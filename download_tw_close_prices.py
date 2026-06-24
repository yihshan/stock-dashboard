#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import argparse
import csv
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.request import Request, urlopen

# --- 核心路徑處理 ---
# 如果是在 GitHub Actions 執行，強制存到 ./data；否則用 OneDrive
DATA_DIR = os.getenv("OUTPUT_DIR", Path(__file__).resolve().parent / "data")
SAVE_DIR = Path(DATA_DIR)
SAVE_DIR.mkdir(parents=True, exist_ok=True)

def download_data():
    """直接使用 data.gov.tw 的穩定連結進行下載"""
    # 這邊是合併了您原本的下載與解析邏輯
    urls = {
        "上市": "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data",
        "上櫃": "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&d=115/06/10&se=EW&o=csv"
    }
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    print(f"[{today_str}] 開始下載資料...")

    for market, url in urls.items():
        try:
            req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urlopen(req, timeout=30) as res:
                content = res.read().decode('big5', errors='ignore')
                #content = res.read().decode('utf-8-sig')
                
            file_path = SAVE_DIR / f"台股每日收盤價_{market}_{today_str}.csv"
            with open(file_path, "w", encoding="utf-8-sig") as f:
                f.write(content)
            print(f"✅ {market} 資料已存檔: {file_path}")
        except Exception as e:
            print(f"❌ {market} 下載失敗: {e}")

if __name__ == "__main__":
    download_data()
# 假設上市資料存在 twse_rows，上櫃資料存在 tpex_rows
all_rows = twse_rows + tpex_rows  # 將兩份資料合併成一個大清單

# --- 接著執行我們之前加上的「資料瘦身」過濾邏輯 ---
target_stocks = set()
# ... 讀取庫存與監控名單的程式碼 ...

if target_stocks:
    all_rows = [row for row in all_rows if row.get('股票名稱', '').strip() in target_stocks]

# --- 最後，統一存成一個合併的 CSV 檔案 ---
ad_date = now_taiwan().strftime("%Y-%m-%d")
output_path = output_dir / f"台股每日收盤價_{ad_date}.csv"

with output_path.open("w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(all_rows)
