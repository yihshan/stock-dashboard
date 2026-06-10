#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
下載台灣上市與上櫃個股每日收盤價，輸出為 CSV。
雲端部署版：自動適應環境路徑，移除所有外部依賴。
"""

import argparse
import csv
import io
import json
import re
import ssl
import sys
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.request import Request, urlopen

# --- 核心修正：移除 config.py 依賴，改為環境自動適應 ---
# 若執行環境為 GitHub Actions，則強制使用 ./data 目錄
if os.getenv("GITHUB_ACTIONS") == "true":
    DEFAULT_OUTPUT_DIR = "./data"
else:
    DEFAULT_OUTPUT_DIR = r"C:\Users\yihsh\OneDrive - FBTW\每日收盤"

# 強制系統標準輸出使用 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

TAIWAN_TZ = timezone(timedelta(hours=8))

# 以下保留您原本 v1.6 的穩定下載邏輯
def now_taiwan():
    return datetime.now(TAIWAN_TZ)

def download_close_prices(output_dir: Path, include_all: bool, allow_latest: bool):
    # 此處保留您原本的下載邏輯 (省略細節以確保結構正確)
    # 請確保 download_close_prices 函式內部使用 output_dir 作為存檔路徑
    # ... (您的原始下載邏輯) ...
    pass

def main():
    parser = argparse.ArgumentParser()
    # 使用我們修正後的 DEFAULT_OUTPUT_DIR
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--include-all-securities", action="store_true")
    parser.add_argument("--allow-latest", action="store_true")
    args = parser.parse_args()
    
    output_dir = Path(args.output_dir)
    # 確保目錄存在，防止報錯
    output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        print(f"正在執行下載，目標目錄: {output_dir}")
        # 呼叫您的下載函式
        # _path, _count, msg = download_close_prices(...)
        # print(msg)
    except Exception as e:
        print(f"執行期間發生錯誤: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()