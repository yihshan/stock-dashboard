# 📊 台股監控系統郵件設定檔

# 數據儲存路徑 (請確保此路徑存在，所有腳本將同步使用此路徑)
DATA_DIR = "C:/Users/yihsh/OneDrive - FBTW/每日收盤"

# SMTP 伺服器設定 (若改用 Gmail 請修改為 smtp.gmail.com)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# 您的郵件帳號
EMAIL_USER = "ihsiang.tsou@gmail.com"

# 您的應用程式密碼
EMAIL_PASSWORD = "gshf qupk rkqq vbpl"

# 收件人清單
RECIPIENTS = [
    "ihsiang.tsou@gmail.com"
]

# 密件副本收件人清單 (可選，用於不公開顯示的收件人)
BCC_RECIPIENTS = [
    "pp4740qq@gmail.com",
    "a0939390292@gmail.com",
    "sk.hsieh@yahoo.com.tw",
    "ywho6907@gmail.com"
    #"yihshan@ihsiangtsou.onmicrosoft.com"
]

# Google Gemini API 設定
GEMINI_API_KEY = "AQ.Ab8RN6L1IPxoM0e0bzuXKaGya7q4ckrEPM2_1yrDHKra-l7utg"

# 技術分析參數
K_THRESHOLD = 15  # K 值低於此數值則發送告警
KD_PERIOD = 9     # K 值計算週期 (K9)

# Line Messaging API 設定 (請前往 Line Developers 取得 Channel Access Token)
LINE_CHANNEL_ACCESS_TOKEN = "raP62DD+supQ+bK5KdZsN3q/Re9IyApL7Nefm+L0OR3Mt8Hm5L9Aesv03hslKx7wzYukXwC4UridkWpy/K4pZ0XLvqsUO8xO3H2PiWp41kbiyl0wqAbXpuI+I2hsvuwXBb3XT6Q4hrvPd/ptneaHUQdB04t89/1O/w1cDnyilFU=" # 請在此填入您的 Line Messaging API Channel Access Token
LINE_USER_ID = "" # 請在此填入您的 Line User ID 或 Group ID，用於廣播訊息
