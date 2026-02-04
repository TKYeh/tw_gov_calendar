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
from datetime import datetime, timedelta
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
    cal = Calendar()
    for _, row in df.iterrows():
        raw_date = row["西元日期"]
        if pd.isna(raw_date):
            # 沒有日期的列跳過
            continue

        # 有些 CSV 會把日期讀成浮點數（例如 20240101.0），先安全轉為整數再格式化
        try:
            date_int = int(float(raw_date))
            date_str = f"{date_int:08d}"
        except Exception:
            date_str = str(raw_date).strip()

        day = datetime.strptime(date_str, "%Y%m%d")

        desc = str(row.get("備註", "")).strip()

        # 如果原始資料有 `假日類別` 欄位，就使用它；沒有的話改用 `備註` 作為類別/標題來源
        if "假日類別" in df.columns:
            category = str(row.get("假日類別", "")).strip()
        else:
            category = desc

        title = category if category else "行事曆"
        if "補班" in desc:
            title = "補班"

        event = Event()
        event.uid = f"gov-calendar-{date_str}"
        event.name = title
        event.begin = day
        event.end = day + timedelta(days=1)
        event.make_all_day()

        description_lines = []

        # 只有在原始資料存在 `假日類別` 時，才顯示「類型：...」，避免與備註重複
        if "假日類別" in df.columns and category:
            description_lines.append(f"類型：{category}")
        if desc:
            # 不加上「備註：」前綴，直接放入內容
            description_lines.append(desc)

        if description_lines:
            event.description = "\n".join(description_lines)

        cal.events.add(event)

    # 確保輸出目錄存在並寫檔
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.writelines(cal)
    print(f"✅ 已輸出：{output_path}")


# =========================
# 主流程
# =========================

def main():
    csv_infos = get_csv_urls()
    print(f"共找到 {len(csv_infos)} 份年度 CSV")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for url, name in csv_infos:
        print(f"下載 CSV：{url}")
        r = requests.get(url)
        r.encoding = "utf-8-sig"
        df = pd.read_csv(StringIO(r.text))

        df_filtered = remove_weekends(df)

        # 依 resourceDescription 產生安全檔名
        base = re.sub(r"\\.csv$", "", name, flags=re.I)
        base = re.sub(r"[\\/:*?\"<>|]", "_", base)
        base = base.strip()
        if not base:
            base = "calendar"

        out_path = os.path.join(OUTPUT_DIR, f"{base}.ics")
        generate_ics(df_filtered, out_path)

    print("✅ 全部處理完成")


if __name__ == "__main__":
    main()