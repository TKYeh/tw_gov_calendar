"""
台灣政府行政機關辦公日曆（data.gov.tw）
- 使用 API v2 自動取得所有年度 CSV
- 只移除「例假日」
- 其他日期全部保留
- 輸出 iCalendar (.ics)
"""

import os
import re
import requests
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta, timezone
from ics import Calendar, Event

# =========================
# 基本設定
# =========================

DATASET_API = "https://data.gov.tw/api/v2/rest/dataset/14718"
OUTPUT_FILE = "taiwan_gov_calendar_no_weekend.ics"
OUTPUT_DIR = "output"
CALENDAR_NAME = "台灣政府行事曆（移除例假日）"

# =========================
# 取得所有年度 CSV 下載網址
# =========================

def get_csv_urls() -> list[tuple[str, str]]:
    resp = requests.get(DATASET_API)
    resp.raise_for_status()
    data = resp.json()

    distributions = data["result"]["distribution"]

    csv_urls = []
    for d in distributions:
        if d.get("resourceFormat", "").upper() == "CSV":
            url = d.get("resourceDownloadUrl", "")
            name = d.get("resourceDescription", "")
            # 排除 Google 行事曆版本
            if "Google" not in name:
                csv_urls.append((url, name))

    return csv_urls


# =========================
# 下載並合併所有年度 CSV
# =========================

def load_all_years(csv_urls: list[str]) -> pd.DataFrame:
    frames = []

    for url in csv_urls:
        print(f"下載 CSV：{url}")
        r = requests.get(url)
        r.encoding = "utf-8-sig"  # 使用 utf-8-sig 以處理 BOM
        df = pd.read_csv(StringIO(r.text))
        frames.append(df)

    return pd.concat(frames, ignore_index=True)


def _normalize_column_name(name: str) -> str:
    if not isinstance(name, str):
        return name
    # strip BOM and whitespace, unify full-width/half-width spaces
    name = name.replace('\ufeff', '').strip()
    name = name.replace('\u3000', ' ')
    return name


def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names and map common alternative names to canonical ones.

    Returns a DataFrame where possible alternative column names are copied to
    canonical names: '西元日期', '假日類別', '備註'.
    """
    if df is None or df.empty:
        return df

    # normalize existing column names
    cols = {c: _normalize_column_name(c) for c in df.columns}
    df = df.rename(columns=cols)

    # possible alternatives
    date_alts = ['西元日期', '日期', 'Date', 'date']
    cat_alts = ['假日類別', '假日類型', 'Subject', '類別', '類型']
    remark_alts = ['備註', '備注', '備考', 'Remarks', 'Remark', '說明']

    # helper to copy first matching alt to canonical name
    def map_first(canonical, alts):
        for a in alts:
            if a in df.columns:
                df[canonical] = df[a]
                return True
        return False

    map_first('西元日期', date_alts)
    map_first('假日類別', cat_alts)
    map_first('備註', remark_alts)

    return df


def get_latest_ics(output_dir: str, exclude: tuple[str, ...] = ("basic.ics",), mode: str = "by_year") -> str | None:
    """Return path to the latest .ics in `output_dir`.

    mode:
      - 'by_year': try to parse a year from filename and return the highest year match;
      - 'mtime': return the most recently modified file.
    """
    import re

    latest_fp = None

    try:
        names = os.listdir(output_dir)
    except Exception:
        return None

    if mode == "by_year":
        best_year = -1
        for fname in names:
            if not fname.lower().endswith('.ics') or fname in exclude:
                continue
            m = re.search(r"(\d{3,4})", fname)
            if m:
                try:
                    year = int(m.group(1))
                except Exception:
                    continue
                if year > best_year:
                    best_year = year
                    latest_fp = os.path.join(output_dir, fname)
        if latest_fp:
            return latest_fp

    # fallback or mode == 'mtime'
    latest_mtime = 0.0
    for fname in names:
        if not fname.lower().endswith('.ics') or fname in exclude:
            continue
        fp = os.path.join(output_dir, fname)
        try:
            m = os.path.getmtime(fp)
        except Exception:
            continue
        if m > latest_mtime:
            latest_mtime = m
            latest_fp = fp

    return latest_fp


def get_ics_for_current_year(output_dir: str, exclude: tuple[str, ...] = ("basic.ics",)) -> str | None:
    """Select an .ics file in output_dir that matches the current year.

    Preference order:
      1. Find filename containing ROC year (current_year - 1911), e.g. '115'
      2. Find filename containing Gregorian year, e.g. '2026'
      3. Fallback to get_latest_ics
    """
    now = datetime.now()
    gregorian = now.year
    roc = now.year - 1911

    # search for ROC year in filenames
    for fname in os.listdir(output_dir):
        if not fname.lower().endswith('.ics') or fname in exclude:
            continue
        if str(roc) in fname:
            return os.path.join(output_dir, fname)

    # search for gregorian year
    for fname in os.listdir(output_dir):
        if not fname.lower().endswith('.ics') or fname in exclude:
            continue
        if str(gregorian) in fname:
            return os.path.join(output_dir, fname)

    # fallback
    return get_latest_ics(output_dir, exclude=exclude, mode='mtime')


# =========================
# 移除例假日（只移除這一種）
# =========================

def remove_weekends(df: pd.DataFrame) -> pd.DataFrame:
    """
    欄位說明（政府 CSV 常見）：
    - 日期：西元日期（YYYYMMDD）
    - 假日類型：假日類別
    """
    # 若存在 `假日類別`，維持原行為：只移除值為「例假日」的列
    if "假日類別" in df.columns:
        return df[df["假日類別"] != "例假日"].copy()

    # 備註不是 NaN 的列才視為有意義的節日，其他排除
    if "備註" in df.columns:
        return df[df["備註"].notna()].copy()

    # 若兩個欄位都不存在，回傳空的 DataFrame（並由呼叫端繼續處理其他檔案）
    print("警告：CSV 未包含 '假日類別' 或 '備註' 欄位，該檔案將被跳過")
    return df.iloc[0:0].copy()


# =========================
# 產生 ICS
# =========================

def generate_ics(df: pd.DataFrame, output_path: str):
    def escape_text(s: str) -> str:
        s = s.replace('\\', '\\\\')
        s = s.replace('\n', '\\n')
        s = s.replace('\r', '')
        s = s.replace(',', '\\,')
        s = s.replace(';', '\\;')
        return s

    lines = []
    lines.append('BEGIN:VCALENDAR')
    lines.append('PRODID:-//Generated by tw_gov_calendar//EN')
    lines.append('VERSION:2.0')
    lines.append('CALSCALE:GREGORIAN')
    lines.append('METHOD:PUBLISH')
    lines.append(f'X-WR-CALNAME:{CALENDAR_NAME}')
    lines.append('X-WR-TIMEZONE:Asia/Taipei')
    lines.append('X-WR-CALDESC:Generated by tw_gov_calendar')

    now_stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')

    for _, row in df.iterrows():
        raw_date = row['西元日期']
        if pd.isna(raw_date):
            # 沒有日期的列跳過
            continue

        # 有些 CSV 會把日期讀成浮點數（例如 20240101.0），先安全轉為整數再格式化
        try:
            date_int = int(float(raw_date))
            date_str = f"{date_int:08d}"
        except Exception:
            date_str = str(raw_date).strip()

        try:
            day = datetime.strptime(date_str, '%Y%m%d')
        except Exception:
            # skip invalid date
            continue

        desc = str(row.get('備註', '')).strip()
        if '假日類別' in df.columns:
            category = str(row.get('假日類別', '')).strip()
        else:
            category = desc

        title = category if category else '行事曆'
        if '補班' in desc:
            title = '補班'

        dtstart = date_str
        dtend_day = (datetime.strptime(date_str, '%Y%m%d') + timedelta(days=1)).strftime('%Y%m%d')

        lines.append('BEGIN:VEVENT')
        lines.append(f'DTSTART;VALUE=DATE:{dtstart}')
        lines.append(f'DTEND;VALUE=DATE:{dtend_day}')
        lines.append(f'DTSTAMP:{now_stamp}')
        uid = f'gov-calendar-{date_str}'
        lines.append(f'UID:{uid}')
        lines.append(f'CREATED:{now_stamp}')
        lines.append(f'LAST-MODIFIED:{now_stamp}')
        lines.append('SEQUENCE:0')
        lines.append('STATUS:CONFIRMED')
        lines.append('SUMMARY:' + escape_text(title))
        # TRANSP: match basic default TRANSPARENT unless title contains 補放等需 OPAQUE — keep TRANSPARENT
        lines.append('TRANSP:TRANSPARENT')
        if desc:
            lines.append('DESCRIPTION:' + escape_text(desc))
        lines.append('END:VEVENT')

    lines.append('END:VCALENDAR')

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f'✅ 已輸出：{output_path}')


# =========================
# 主流程
# =========================

def main():
    csv_infos = get_csv_urls()
    print(f"共找到 {len(csv_infos)} 份年度 CSV")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for url, name in csv_infos:
        print(f"下載 CSV：{name}")
        r = requests.get(url)
        content = r.content

        # 嘗試偵測並使用合理的編碼順序解碼（chardet 可能提供建議）
        try:
            import chardet
            det = chardet.detect(content[:4096])
            guessed = det.get('encoding')
        except Exception:
            guessed = None

        tried = []
        df = None
        enc_candidates = [guessed, 'utf-8-sig', 'utf-8', 'big5', 'cp950', 'latin1']
        for enc in enc_candidates:
            if not enc or enc in tried:
                continue
            tried.append(enc)
            try:
                text = content.decode(enc)
                df = pd.read_csv(StringIO(text))
                # print(f"  使用編碼: {enc}")
                break
            except Exception:
                continue

        if df is None:
            # 最後手段：以忽略錯誤方式解碼
            text = content.decode('utf-8', errors='replace')
            df = pd.read_csv(StringIO(text))

        # 正規化並嘗試對應常見替代欄名（處理 BOM/欄名變化）
        df = prepare_df(df)
        df_filtered = remove_weekends(df)

        # 依 resourceDescription 產生安全檔名
        base = re.sub(r"\\.csv$", "", name, flags=re.I)
        base = re.sub(r"[\\/:*?\"<>|]", "_", base)
        base = base.strip()
        if not base:
            base = "calendar"

        out_path = os.path.join(OUTPUT_DIR, f"{base}.ics")
        generate_ics(df_filtered, out_path)

    # 選取 output 下的最新 .ics（優先使用檔名年份，若找不到則用檔案修改時間）
    latest_fp = get_latest_ics(OUTPUT_DIR, exclude=("basic.ics",), mode='by_year')
    if not latest_fp:
        latest_fp = get_latest_ics(OUTPUT_DIR, exclude=("basic.ics",), mode='mtime')

    if latest_fp:
        dest = os.path.abspath('TW_gov_calendar.ics')
        try:
            import shutil
            shutil.copy2(latest_fp, dest)
            print(f'✅ 已複製最新 ICS（來源：{os.path.basename(latest_fp)}）')
        except Exception as e:
            print('複製失敗：', e)
    else:
        print('找不到可用的 ICS 檔案來複製')

    print('✅ 全部處理完成')


if __name__ == "__main__":
    main()