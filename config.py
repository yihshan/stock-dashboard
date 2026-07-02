import os
from pathlib import Path

# 📊 台股監控系統郵件設定檔

# 取得目前專案根目錄，並將資料夾指向專案內的 'data' 資料夾
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

# SMTP 伺服器設定
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# 從環境變數讀取機密資訊 (支援 GitHub Actions 與 Streamlit 環境)
try:
    # Streamlit 雲端環境專用的讀取方式
    import streamlit as st
    EMAIL_USER = st.secrets.get("EMAIL_USER", os.getenv("EMAIL_USER", "ihsiang.tsou@gmail.com"))
    EMAIL_PASSWORD = st.secrets.get("EMAIL_PASSWORD", os.getenv("EMAIL_PASSWORD", ""))
    GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

    # 🟢 修正：將 LINE 變數完美塞入 Streamlit 安全憑證鏈
    LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
    LINE_USER_ID = os.getenv("LINE_USER_ID", "")
    
except Exception:
    # GitHub Actions 執行時的讀取方式
    EMAIL_USER = os.getenv("EMAIL_USER", "ihsiang.tsou@gmail.com")
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# 收件人清單
RECIPIENTS = [
    "ihsiang.tsou@gmail.com"
]

# 密件副本收件人清單
BCC_RECIPIENTS = [
    #"pp4740qq@gmail.com",
    #"a0939390292@gmail.com",
    #"sk.hsieh@yahoo.com.tw",
    #"ywho6907@gmail.com"
]

# 技術分析參數
K_THRESHOLD = 15
