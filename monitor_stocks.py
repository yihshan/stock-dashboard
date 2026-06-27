import os
import glob
import re
import csv
import smtplib
import requests
import urllib3
import numpy as np
import pandas as pd
from datetime import datetime, date
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys

# 載入 LINE 發送模組
try:
    from line_messaging import send_line_message
except ModuleNotFoundError:
    print("⚠️ 找不到 line_messaging 模組，將改為日誌輸出。")
    def send_line_message(token, user_id, msg): print(f"[Line 模擬] {msg}")

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
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
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

# --- 讀取最新一日報表日期用 ---
def get_report_date():
    csv_files = glob.glob(str(BASE_DIR / "台股每日收盤價_*.csv"))
    if not csv_files: return str(date.today())
    latest_file = max(csv_files, key=os.path.basename)
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(latest_file))
    return date_match.group(1) if date_match else str(date.today())

# --- 升級版雙料 Email 派送模組 (包含大波段預警 + 技術指標總覽) ---
def send_combined_email(report_date, strategy_alerts, all_stocks):
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 台股策略決策報告與技術指標總覽 - {report_date}"
    msg['From'] = config.EMAIL_USER
    msg['To'] = ", ".join(config.RECIPIENTS)

    # 1. 建立大波段策略警告區塊的 HTML
    if strategy_alerts:
        alert_html = (
            "<div style='background-color: #fffaf0; border-left: 5px solid #dd6b20; padding: 15px; margin-bottom: 25px; border-radius: 4px;'>"
            "<h3 style='color: #dd6b20; margin-top: 0;'>⚠️ 大波段策略買賣觸發提示</h3>"
            "<ul style='padding-left: 20px; line-height: 1.6; color: #2d3748;'>"
        )
        for act in strategy_alerts:
            act_br = act.replace('\n', '<br>')
            alert_html += f"<li style='margin-bottom: 8px;'>{act_br}</li>"
        alert_html += "</ul></div>"
    else:
        alert_html = (
            "<div style='background-color: #f0fff4; border-left: 5px solid #38a169; padding: 15px; margin-bottom: 25px; border-radius: 4px;'>"
            "<h3 style='color: #38a169; margin-top: 0;'>✅ 大波段策略監控正常</h3>"
            "<p style='color: #276749; margin: 0;'>今日現有庫存皆在安全波段中，且觀察名單尚未觸發加倉買進目標價。</p>"
            "</div>"
        )

    # 2. 建立原本的每日技術指標表格 (完美防呆 NaN 數值)
    k_threshold = getattr(config, 'K_THRESHOLD', 15)
    table_rows = ""
    for s in all_stocks:
        is_nan_k = pd.isna(s['k'])
        k_style = "color:red; font-weight:bold;" if not is_nan_k and s['k'] < k_threshold else ""
        
        k_val = f"{s['k']:.2f}" if not is_nan_k else "N/A"
        d_val = f"{s['d']:.2f}" if not pd.isna(s['d']) else "N/A"
        dif_val = f"{s['dif']:.2f}" if not pd.isna(s['dif']) else "N/A"
        macd_val = f"{s['macd']:.2f}" if not pd.isna(s['macd']) else "N/A"
        osc_val = f"{s['osc']:.2f}" if not pd.isna(s['osc']) else "N/A"

        row = "<tr>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'>{s['name']} ({s['id']})</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['close']:.2f}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right; {k_style}'>{k_val}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{d_val}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{dif_val}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{macd_val}</td>"
        row += f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{osc_val}</td>"
        row += "</tr>"
        table_rows += row
        
    html = (
        "<html>"
        "<body style=\"font-family: 'Microsoft JhengHei', sans-serif; padding: 20px;\">"
        "<h2 style=\"color: #1a365d;\">📊 每日台股策略監控與技術指標自動彙整</h2>"
        f"<p style='color: #4a5568;'>數據基準日：{report_date}</p>"
        f"{alert_html}"  
        "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 30px 0;'>"
        "<h3 style=\"color: #004a99;\">📈 關注個股最新日線指標總覽</h3>"
        f"<p>指標狀態提示 (K值低於 {k_threshold} 會以紅字標註)：</p>"
        "<table style=\"width:100%; border-collapse:collapse; margin-top:15px;\">"
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
        "<p style='font-size: 12px; color: #a0aec0; margin-top: 20px;'>💡 本郵件由 Agentic AI 投資決策系統自動彙整發送</p>"
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
        print("📨 [Email 系統] 整合型策略總覽報告郵件已成功寄出！")
    except Exception as e: 
        print(f"❌ [Email 系統] 郵件發送失敗: {e}")

# --- 主程式核心邏輯 ---
def main():
    report_date = get_report_date()
    print(f"🚀 啟動大波段策略與技術分析綜合監控... (基準日: {report_date})")
    
    strategy_alerts = []
    all_stocks_data = {}

    # ==========================================
    # 核心一：技術指標計算與監控名單掃描
    # ==========================================
    if MONITOR_FILE.exists():
        monitor_df = pd.read_excel(MONITOR_FILE)
        name_col = next((col for col in monitor_df.columns if '名稱' in str(col)), monitor_df.columns[0])
        id_col = next((col for col in monitor_df.columns if '代號' in str(col)), None)
        
        for _, row in monitor_df.iterrows():
            name = str(row[name_col]).strip()
            s_id = str(row[id_col]).strip() if id_col and pd.notna(row[id_col]) else ""
            if not s_id:
                match = re.search(r'\((\d+)\)', name)
                s_id = match.group(1) if match else ""
            if not name or name == 'nan': continue
                
            df = get_stock_history(s_id, name)
            if df.empty: continue
            
            # 【重要修正】優先提取今日最新的市價，確保大波段監控不受歷史天數限制！
            latest = df.iloc[-1]
            current_price = latest['Close']
            
            # 🎯 優先進行大波段「買進目標價」比對
            if '買進目標價' in monitor_df.columns and pd.notna(row['買進目標價']):
                target_price = float(row['買進目標價'])
                pe_limit = row['買進本益比上限'] if '買進本益比上限' in monitor_df.columns else None
                if current_price <= target_price:
                    pe_msg = f" (本益比門檻: {pe_limit}x)" if pd.notna(pe_limit) else ""
                    strategy_alerts.append(f"🎯 [監控買進提示] {name}\n  - 今日收盤: {current_price}\n  - 買進目標價: {target_price}{pe_msg}\n  - 股價已修正至大波段安全邊際買點，建議分批佈局！")

            # --- 歷史保護過濾：若天數不滿 10 天，略過技術指標(KD/MACD)計算，但前方的價格監控依然生效！ ---
            if len(df) < 10: 
                print(f"⚠️ {name} 歷史天數不足 10 筆，略過日線技術指標計算。")
                all_stocks_data[name] = {
                    'name': name, 'id': s_id, 'close': current_price, 
                    'k': np.nan, 'd': np.nan, 'dif': np.nan, 'macd': np.nan, 'osc': np.nan
                }
                continue
                
            df = calculate_k9(df)
            dif, macd, osc = calculate_macd(df)
            
            latest = df.iloc[-1]
            current_k = latest['K']
            current_d = latest['D']
            current_osc = osc.iloc[-1]
            
            all_stocks_data[name] = {
                'name': name, 'id': s_id, 'close': current_price, 
                'k': current_k, 'd': current_d, 'dif': dif.iloc[-1], 'macd': macd.iloc[-1], 'osc': current_osc
            }
            
            # 舊版低K值 Line 逢低布局警示保留
            k_threshold = getattr(config, 'K_THRESHOLD', 15)
            if current_k < k_threshold:
                news = fetch_stock_news(name)
                msg_text = f"🚨【逢低布局警示】{name} ({s_id})\n最新收盤價：{current_price:.2f}\n狀態：K值 <{k_threshold} 且MACD綠柱收斂！{news}\n\n💡 由Agentic AI系統自動發送"
                send_line_message(getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', ''), getattr(config, 'LINE_USER_ID', ''), msg_text)

    # ==========================================
    # 核心二：庫存股票大波段「移動停利」追蹤
    # ==========================================
    if INVENTORY_FILE.exists():
        try:
            inv_df = pd.read_excel(INVENTORY_FILE)
            for col in ['移動停利百分比(%)', '波段最高價', '保本停損價']:
                if col not in inv_df.columns: inv_df[col] = None
                
            updated_rows = []
            for name, group in inv_df.groupby('股票名稱'):
                name = str(name).strip()
                total_shares = group['股數'].sum()
                if total_shares <= 0: continue
                
                avg_cost = (group['成本'] * group['股數']).sum() / total_shares
                trail_pct = group['移動停利百分比(%)'].dropna().head(1).values
                trail_pct = float(trail_pct[0]) if len(trail_pct) > 0 and not pd.isna(trail_pct[0]) else 15.0
                base_stop = group['保本停損價'].dropna().head(1).values
                base_stop = float(base_stop[0]) if len(base_stop) > 0 and not pd.isna(base_stop[0]) else avg_cost
                highest_record = group['波段最高價'].dropna().head(1).values
                highest_price = float(highest_record[0]) if len(highest_record) > 0 and not pd.isna(highest_record[0]) else 0.0
                
                current_price = all_stocks_data[name]['close'] if name in all_stocks_data else np.nan
                if pd.isna(current_price):
                    hist_df = get_stock_history("", name)
                    current_price = hist_df.iloc[-1]['Close'] if not hist_df.empty else 0.0

                if current_price > 0:
                    if current_price > highest_price:
                        highest_price = current_price
                        print(f"🚀 {name} 創波段新高！更新最高價為: {highest_price}")
                    
                    sell_trigger_price = highest_price * (1 - (trail_pct / 100))
                    if current_price <= sell_trigger_price:
                        strategy_alerts.append(f"⚠️ [庫存賣出預警] {name}\n  - 今日收盤: {current_price}\n  - 波段高點: {highest_price}\n  - 已回撤超過 {trail_pct}%\n  - 觸發移動停利線 ({sell_trigger_price:.1f})，建議落袋為安！")
                    elif current_price <= base_stop:
                        strategy_alerts.append(f"🛑 [保本停損觸發] {name}\n  - 今日收盤: {current_price}\n  - 綜合平均成本: {avg_cost:.1f}\n  - 觸發保本停損底線 ({base_stop:.1f})，建議全數離場！")
                
                for _, row in group.iterrows():
                    row['移動停利百分比(%)'] = trail_pct
                    row['波段最高價'] = highest_price
                    row['保本停損價'] = base_stop
                    updated_rows.append(row)
                    
            new_inv_df = pd.DataFrame(updated_rows)
            new_inv_df.to_excel(INVENTORY_FILE, index=False)
            print("💾 庫存股票之最新波段最高價已成功同步回存至 Excel。")
        except Exception as e:
            print(f"❌ 處理庫存大波段監控時發生錯誤: {e}")

    # ==========================================
    # 核心三：雙通道分流發送 (Line + 整合 Email)
    # ==========================================
    if strategy_alerts:
        line_token = getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', '')
        line_user = getattr(config, 'LINE_USER_ID', '')
        full_line_msg = f"\n📊 【台股大波段策略決策報告】\n基準日: {report_date}\n" + "\n---------------------\n".join(strategy_alerts)
        send_line_message(line_token, line_user, full_line_msg)
        print("🔔 策略買賣訊號已成功推播至 Line。")

    if all_stocks_data:
        send_combined_email(report_date, strategy_alerts, list(all_stocks_data.values()))

if __name__ == "__main__":
    main()
