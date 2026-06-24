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
                content = res.read().decode('utf-8-sig')
                
            file_path = SAVE_DIR / f"台股每日收盤價_{market}_{today_str}.csv"
            with open(file_path, "w", encoding="utf-8-sig") as f:
                f.write(content)
            print(f"✅ {market} 資料已存檔: {file_path}")
        except Exception as e:
            print(f"❌ {market} 下載失敗: {e}")

if __name__ == "__main__":
    download_data()
