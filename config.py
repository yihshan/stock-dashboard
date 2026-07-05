import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# 📊 台股監控系統郵件與 LINE 整合設定檔

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# 從環境變數讀取機密資訊
try:
    # 👑 1. 嘗試 Streamlit 雲端憑證鏈
    import streamlit as st
    EMAIL_USER = st.secrets.get("EMAIL_USER", os.getenv("EMAIL_USER", "ihsiang.tsou@gmail.com"))
    EMAIL_PASSWORD = st.secrets.get("EMAIL_PASSWORD", os.getenv("EMAIL_PASSWORD", ""))
    GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

    #LINE_CHANNEL_ACCESS_TOKEN = st.secrets.get("LINE_CHANNEL_ACCESS_TOKEN", os.getenv("LINE_CHANNEL_ACCESS_TOKEN", ""))
    #LINE_USER_ID = st.secrets.get("LINE_USER_ID", os.getenv("LINE_USER_ID", ""))
    
except Exception:
    # 👑 2. 降級回歸 GitHub Actions / 本地環境憑證鏈
    EMAIL_USER = os.getenv("EMAIL_USER", "ihsiang.tsou@gmail.com")
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

    #LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
    #LINE_USER_ID = os.getenv("LINE_USER_ID", "")

# 🛠️ 終極防禦：如果在 Actions 區塊讀取失敗，嘗試直接做全域強抓
#if not LINE_CHANNEL_ACCESS_TOKEN:
    #LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
#if not LINE_USER_ID:
    #LINE_USER_ID = os.environ.get("LINE_USER_ID", "")
    
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
