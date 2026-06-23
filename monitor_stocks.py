import pandas as pd
import numpy as np
import os
import glob
import re
import smtplib
import requests
import urllib.parse
import urllib3
from datetime import datetime, date
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys

# 載入 LINE 發送模組
from line_messaging import send_line_message

# 隱藏 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 強制 UTF-8 輸出
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import config
except ImportError:
    print("❌ 找不到 config.py，請確保設定檔存在。")
    sys.exit(1)

BASE_DIR = Path(config.DATA_DIR)
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"

# --- 日期與字串處理輔助函式 ---
def parse_date(val):
    if pd.isna(val): return None
    s = str(val).strip().split(' ')[0]
    if not s: return None
    for sep in ['/', '-']:
        parts = s.split(sep)
        if len(parts) == 3:
            try: return date(int(parts[0]) if int(parts[0]) > 1911 else int(parts[0]) + 1911, int(parts[1]), int(parts[2]))
            except: continue
    digits = re.sub(r'\D', '', s)
    if len(digits) == 8:
        try: return datetime.strptime(digits, '%Y%m%d').date()
        except: pass
    elif len(digits) == 7:
        try: return date(int(digits[:3]) + 1911, int(digits[3:5]), int(digits[5:]))
        except: pass
    return None

def extract_date_from_filename(filename):
    match = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', filename)
    if match: return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    match = re.search(r'(\d{8})|(\d{7})', filename)
    if match:
        d = match.group()
        return datetime.strptime(d, '%Y%m%d').date() if len(d) == 8 else date(int(d[:3]) + 1911, int(d[3:5]), int(d[5:]))
    return None

def clean_price(val):
    try: return float(str(val).replace(',', '').strip())
    except: return np.nan

# --- 技術指標計算模組 ---
def calculate_k9(df):
    df = df.sort_values('Date').reset_index(drop=True)
    df['L9'] = df['Low'].rolling(window=9).min() if not (df['High'] == df['Close']).all() else df['Close'].rolling(window=9).min()
    df['H9'] = df['High'].rolling(window=9).max() if not (df['High'] == df['Close']).all() else df['Close'].rolling(window=9).max()
    denom = df['H9'] - df['L9']
    df['RSV'] = 100 * (df['Close'] - df['L9']) / denom
    df.loc[denom == 0, 'RSV'] = 50.0
    
    k_vals, d_vals = [], []
    current_k, current_d = 50.0, 50.0
    for rsv in df['RSV']:
        if pd.isna(rsv): 
            k_vals.append(np.nan)
            d_vals.append(np.nan)
        else:
            current_k = (2/3) * current_k + (1/3) * rsv
            current_d = (2/3) * current_d + (1/3) * current_k
            k_vals.append(current_k)
            d_vals.append(current_d)
    df['K'] = k_vals
    df['D'] = d_vals
    return df

def calculate_macd(df):
    df = df.sort_values('Date').copy()
    ema12 = df['Close'].ewm(span=12, adjust=True).mean()
    ema26 = df['Close'].ewm(span=26, adjust=True).mean()
    df['DIF'] = ema12 - ema26
    df['MACD'] = df['DIF'].ewm(span=9, adjust=True).mean()
    df['OSC'] = (df['DIF'] - df['MACD'])
    return df['DIF'], df['MACD'], df['OSC']

# --- CSV 歷史資料讀取模組 ---
def get_stock_history(stock_id, stock_name):
    data = []
    clean_name = re.sub(r'\(.*?\)', '', stock_name).strip()
    s_id = str(stock_id).strip()
    
    for f in glob.glob(str(BASE_DIR / "*.csv")):
        try:
            try: df = pd.read_csv(f, encoding='utf-8-sig')
            except: df = pd.read_csv(f, encoding='cp950')
            
            date_col, name_col, id_col, close_col, high_col, low_col = None, None, None, None, None, None
            
            for col in df.columns:
                c = str(col).strip()
                if any(x in c for x in ['日期', 'Date']): date_col = col
                if any(x in c for x in ['證券代號', 'Code', '股票代號']): id_col = col
                if any(x in c for x in ['名稱', 'Name']): name_col = col
                if any(x in c for x in ['收盤', 'Close', 'ClosingPrice']): close_col = col
                if any(x in c for x in ['最高', 'High', 'HighestPrice']): high_col = col
                if any(x in c for x in ['最低', 'Low', 'LowestPrice']): low_col = col
            
            if id_col and s_id:
                mask = df[id_col].astype(str).str.replace('"', '').str.strip() == s_id
            elif name_col:
                mask = df[name_col].astype(str).str.contains(clean_name, na=False)
            else:
                mask = None
            
            if mask is not None and mask.sum() > 0:
                row = df[mask].iloc[0]
                dt = parse_date(row[date_col]) if date_col else extract_date_from_filename(os.path.basename(f))
                if dt and not pd.isna(clean_price(row[close_col])):
                    c_val = clean_price(row[close_col])
                    h_val = clean_price(row[high_col]) if high_col else c_val
                    l_val = clean_price(row[low_col]) if low_col else c_val
                    data.append({'Date': dt, 'Close': c_val, 'High': h_val, 'Low': l_val})
        except: continue
        
    return pd.DataFrame(data).sort_values('Date').drop_duplicates('Date') if data else pd.DataFrame()

# --- AI 新聞摘要模組 ---
def fetch_stock_news(stock_name):
    try:
        api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
        headers = {'Content-Type': 'application/json', 'x-goog-api-key': str(config.GEMINI_API_KEY).strip()}
        prompt = f"分析台股 {stock_name} 相關新聞，請給 50 字以內的精華摘要。"
        resp = requests.post(api_url, json={"contents": [{"parts": [{"text": prompt}]}]}, headers=headers, timeout=20, verify=False)
        if resp.status_code == 200:
            return "\n【AI 焦點新聞】" + resp.json()['candidates'][0]['content']['parts'][0]['text']
        return "" 
    except: 
        return ""

# --- Email 派送模組 (安全字串拼接版) ---
def send_summary_email(all_stocks):
    if not all_stocks: return
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 每日台股技術指標總覽 - {date.today()}"
    msg['From'] = config.EMAIL_USER
    msg['To'] = ", ".join(config.RECIPIENTS)
    if 'Bcc' in msg: del msg['Bcc']

    k_threshold = getattr(config, 'K_THRESHOLD', 15)
    table_rows = ""
    for s in all_stocks:
        k_style = "color:red; font-weight:bold;" if s['k'] < k_threshold else ""
        row = "<tr>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'>{s['name']} ({s['id']})</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['close']:.2f}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right; {k_style}'>{s['k']:.2f}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['d']:.2f}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['dif']:.2f}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['macd']:.2f}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['osc']:.2f}</td>"
        row += "</tr>"
        table_rows += row
        
    html = (
        "<html>"
        "<body style=\"font-family: 'Microsoft JhengHei', sans-serif;\">"
        "<h2 style=\"color: #1a365d;\">📊 每日台股技術指標自動彙整</h2>"
        f"<p>以下為關注個股最新日線指標 (K值低於 {k_threshold} 會以紅字標註)：</p>"
        "<table style=\"width:100%; border-collapse:collapse; margin-top:20px;\">"
        "<thead><tr style=\"background-color: #004a99; color: white;\">"
        "<th style=\"padding:12px; border:1px solid #ddd;\">股票名稱</th>"
        "<th style=\"padding:12px; border:1px solid #ddd;\">收盤價</th>"
        "<th style=\"padding:12px; border:1px solid #ddd;\">日K值</th>"
        "<th style=\"padding:12px; border:1px solid #ddd;\">日D值</th>"
        "<th style=\"padding:12px; border:1px solid #ddd;\">DIF</th>"
        "<th style=\"padding:12px; border:1px solid #ddd;\">MACD</th>"
        "<th style=\"padding:12px; border:1px solid #ddd;\">OSC(柱狀體)</th>"
        "</tr></thead>"
        f"<tbody>{table_rows}</tbody>"
        "</table>"
        "</body>"
        "</html>"
    )
    msg.attach(MIMEText(html, 'html'))
    
    try:
        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
            server.starttls()
            server.login(config.EMAIL_USER, config.EMAIL_PASSWORD)
            targets = config.RECIPIENTS + getattr(config, 'BCC_RECIPIENTS', [])
            for target in targets:
                server.sendmail(config.EMAIL_USER, target, msg.as_string())
        print("✅ 總覽郵件發送成功")
    except Exception as e: 
        print(f"❌ 郵件發送失敗: {e}")

# --- 主程式 ---
def main():
    print(f"🚀 啟動技術分析監控... (日期: {date.today()})")
    
    if not MONITOR_FILE.exists():
        print(f"❌ 找不到監控檔案: {MONITOR_FILE}")
        return
        
    monitor_df = pd.read_excel(MONITOR_FILE)
    
    name_col = next((col for col in monitor_df.columns if '名稱' in str(col)), monitor_df.columns[0])
    id_col = next((col for col in monitor_df.columns if '代號' in str(col)), None)
    
    all_stocks = []
    for _, row in monitor_df.iterrows():
        name = str(row[name_col]).strip()
        s_id = str(row[id_col]).strip() if id_col and pd.notna(row[id_col]) else ""
        
        if not s_id:
            match = re.search(r'\((\d+)\)', name)
            s_id = match.group(1) if match else ""
            
        if not name or name == 'nan': continue
            
        df = get_stock_history(s_id, name)
        if len(df) < 10: 
            continue
            
        df = calculate_k9(df)
        dif, macd, osc = calculate_macd(df)
        
        latest = df.iloc[-1]
        current_k = latest['K']
        current_d = latest['D']
        current_osc = osc.iloc[-1]
        prev_osc = osc.iloc[-2]
        
        k_threshold = getattr(config, 'K_THRESHOLD', 15)
        if current_k < k_threshold:
# and current_osc < 0 and current_osc > prev_osc:
            news = fetch_stock_news(name)
            msg_text = f"🚨【逢低布局警示】{name} ({s_id})\n最新收盤價：{latest['Close']:.2f}\n狀態：K值 <{k_threshold} 且MACD綠柱收斂！{news}\n\n💡 由Agentic AI系統自動發送"
            send_line_message(getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', ''), getattr(config, 'LINE_USER_ID', ''), msg_text)
            
        all_stocks.append({
            'name': name, 
            'id': s_id, 
            'close': latest['Close'], 
            'k': current_k,
            'd': current_d,
            'dif': dif.iloc[-1],
            'macd': macd.iloc[-1],
            'osc': current_osc
        })
        
    if all_stocks:
        send_summary_email(all_stocks)
    else:
        print("⚠️ 沒有足夠的資料來產生監控報告。")

if __name__ == "__main__":
    main()