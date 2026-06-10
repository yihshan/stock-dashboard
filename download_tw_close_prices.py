#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
下載台灣上市與上櫃個股每日收盤價，輸出為 CSV。

v1.6 更新：
1. 改用 data.gov.tw 推薦的即時 CSV 連結作為資料來源，更新速度較快且穩定。
2. 支援 SSL 憑證忽略，解決 Windows 環境憑證缺失問題。
3. 自動判斷資料日期，僅在有新資料時輸出，或使用 --allow-latest 強制輸出。
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import re
import ssl
import sys
import time
import traceback
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

# 改為直接從環境變數讀取，若沒有則使用預設路徑 (適用於您本地電腦)
DATA_DIR = os.getenv("OUTPUT_DIR", r"C:\Users\yihsh\OneDrive - FBTW\每日收盤")
SAVE_DIR = Path(DATA_DIR)
SAVE_DIR.mkdir(parents=True, exist_ok=True)

# 強制系統標準輸出使用 UTF-8，解決 Windows CMD 亂碼問題
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

TAIWAN_TZ = timezone(timedelta(hours=8))
import os
DEFAULT_OUTPUT_DIR = os.getenv("OUTPUT_DIR", r"C:\Users\yihsh\OneDrive - FBTW\每日收盤")

# 資料來源 (data.gov.tw 推薦之即時 CSV 連結)
TWSE_CSV_URL = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data"
TPEX_CSV_URL = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=data"

# 公司基本資料 (用於篩選個股)
TWSE_COMPANY_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
TPEX_COMPANY_URL = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 tw-close-downloader/1.6"

CSV_COLUMNS = [
    "日期",
    "西元日期",
    "市場",
    "股票代號",
    "股票名稱",
    "收盤價",
    "漲跌",
    "開盤價",
    "最高價",
    "最低價",
    "成交股數",
    "成交金額",
    "成交筆數",
    "資料來源",
    "下載時間",
]

class DownloadError(RuntimeError):
    """資料下載或解析失敗。"""

def now_taiwan() -> datetime:
    return datetime.now(TAIWAN_TZ)

def roc_to_ad_date(roc_date: str) -> Optional[str]:
    """將民國年月日字串轉為 YYYY-MM-DD。支援 1150522 或 115/05/22。"""
    if roc_date is None:
        return None
    text = str(roc_date).strip().replace("/", "").replace("-", "")
    if not re.fullmatch(r"\d{6,7}", text):
        return None
    try:
        year = int(text[:-4]) + 1911
        month = int(text[-4:-2])
        day = int(text[-2:])
        return datetime(year, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return None

def today_roc_yyyymmdd() -> str:
    today = now_taiwan()
    roc_year = today.year - 1911
    return f"{roc_year:03d}{today.month:02d}{today.day:02d}"

def fetch_content(url: str, retries: int = 3, timeout: int = 30) -> str:
    ssl_context = ssl._create_unverified_context()
    last_error: Optional[BaseException] = None
    for attempt in range(1, retries + 1):
        try:
            req = Request(url, headers={"User-Agent": USER_AGENT})
            with urlopen(req, timeout=timeout, context=ssl_context) as response:
                raw = response.read()
            return raw.decode("utf-8-sig")
        except (HTTPError, URLError, TimeoutError) as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(2 * attempt)
            else:
                break
    raise DownloadError(f"無法下載資料：{url}；錯誤：{last_error}")

def fetch_json(url: str) -> Any:
    content = fetch_content(url)
    try:
        return json.loads(content)
    except json.JSONDecodeError as e:
        raise DownloadError(f"解析 JSON 失敗：{url}；錯誤：{e}")

def clean_number(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().replace('"', '')
    if text in {"", "--", "---", "除權", "除息", "除權息"}:
        return text
    return text.replace(",", "")

def is_etf_like(code: str, name: str) -> bool:
    code_text = str(code).strip().upper()
    name_text = str(name).strip().upper()
    if "ETF" in name_text:
        return True
    return bool(re.fullmatch(r"00[0-9A-Z]{2,5}", code_text))

def fetch_stock_code_sets() -> Tuple[set[str], set[str]]:
    try:
        twse_company_json = fetch_json(TWSE_COMPANY_URL)
        tpex_company_json = fetch_json(TPEX_COMPANY_URL)
        twse_codes = {str(row.get("公司代號", "")).strip() for row in twse_company_json if row.get("公司代號")}
        tpex_codes = {str(row.get("SecuritiesCompanyCode", "")).strip() for row in tpex_company_json if row.get("SecuritiesCompanyCode")}
        return twse_codes, tpex_codes
    except Exception:
        # 若公司清單 API 失敗，則不進行精確篩選，改用備援邏輯
        return set(), set()

def download_close_prices(output_dir: Path, include_all_securities: bool, allow_latest: bool) -> Tuple[Optional[Path], int, str]:
    download_time = now_taiwan().strftime("%Y-%m-%d %H:%M:%S %z")
    
    # 取得公司代號清單用於篩選
    twse_stock_codes, tpex_stock_codes = (set(), set()) if include_all_securities else fetch_stock_code_sets()

    all_rows = []

    # 1. 處理上市資料 (TWSE)
    twse_csv_text = fetch_content(TWSE_CSV_URL)
    reader = csv.DictReader(io.StringIO(twse_csv_text))
    for row in reader:
        code = row.get("證券代號", "").strip().replace('"', '')
        name = row.get("證券名稱", "").strip().replace('"', '')
        if not include_all_securities:
            is_stock = code in twse_stock_codes
            is_etf = is_etf_like(code, name)
            if not (is_stock or is_etf):
                continue
        
        roc_date = row.get("日期", "").strip().replace('"', '')
        all_rows.append({
            "日期": roc_date,
            "西元日期": roc_to_ad_date(roc_date) or "",
            "市場": "上市",
            "股票代號": code,
            "股票名稱": name,
            "收盤價": clean_number(row.get("收盤價")),
            "漲跌": clean_number(row.get("漲跌價差")),
            "開盤價": clean_number(row.get("開盤價")),
            "最高價": clean_number(row.get("最高價")),
            "最低價": clean_number(row.get("最低價")),
            "成交股數": clean_number(row.get("成交股數")),
            "成交金額": clean_number(row.get("成交金額")),
            "成交筆數": clean_number(row.get("成交筆數")),
            "資料來源": "TWSE_CSV",
            "下載時間": download_time,
        })

    # 2. 處理上櫃資料 (TPEX)
    tpex_csv_text = fetch_content(TPEX_CSV_URL)
    # 跳過可能的標題列或處理欄位名稱不一致
    tpex_reader = csv.DictReader(io.StringIO(tpex_csv_text))
    for row in tpex_reader:
        code = row.get("代號", "").strip().replace('"', '')
        name = row.get("名稱", "").strip().replace('"', '')
        if not include_all_securities:
            is_stock = code in tpex_stock_codes
            is_etf = is_etf_like(code, name)
            if not (is_stock or is_etf):
                continue
        
        roc_date = row.get("資料日期", "").strip().replace('"', '')
        all_rows.append({
            "日期": roc_date,
            "西元日期": roc_to_ad_date(roc_date) or "",
            "市場": "上櫃",
            "股票代號": code,
            "股票名稱": name,
            "收盤價": clean_number(row.get("收盤")),
            "漲跌": clean_number(row.get("漲跌")),
            "開盤價": clean_number(row.get("開盤")),
            "最高價": clean_number(row.get("最高")),
            "最低價": clean_number(row.get("最低")),
            "成交股數": clean_number(row.get("成交股數")),
            "成交金額": clean_number(row.get("成交金額")),
            "成交筆數": clean_number(row.get("成交筆數")),
            "資料來源": "TPEX_CSV",
            "下載時間": download_time,
        })

    if not all_rows:
        raise DownloadError("抓取不到任何資料。")

    # 日期檢查
    all_rows.sort(key=lambda r: (r["市場"], r["股票代號"]))
    dates = sorted({row["日期"] for row in all_rows if row.get("日期")})
    latest_date = dates[-1] if dates else ""
    today_roc = today_roc_yyyymmdd()

    if latest_date != today_roc and not allow_latest:
        return None, len(all_rows), f"資料日期為 {latest_date}，不是今天 {today_roc}，略過輸出。"

    # 輸出 CSV
    ad_date = roc_to_ad_date(latest_date) or now_taiwan().strftime("%Y-%m-%d")
    output_path = output_dir / f"台股每日收盤價_{ad_date}.csv"
    
    output_dir.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(all_rows)

    return output_path, len(all_rows), f"成功下載 {len(all_rows)} 筆資料（日期：{latest_date}）。"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--include-all-securities", action="store_true")
    parser.add_argument("--allow-latest", action="store_true")
    args = parser.parse_args()
    
    output_dir = Path(args.output_dir)
    log_path = output_dir / "tw_close_downloader.log"
    
    try:
        _path, _count, msg = download_close_prices(output_dir, args.include_all_securities, args.allow_latest)
        print(msg)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"[{now_taiwan()}] {msg}\n")
    except Exception as e:
        err = f"執行失敗：{e}"
        print(err, file=sys.stderr)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"[{now_taiwan()}] {err}\n{traceback.format_exc()}\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
