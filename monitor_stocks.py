import os
import glob
import re
import csv
import smtplib
import pandas as pd
from datetime import datetime
from pathlib import Path
from email.mime.text import MIMEText
from email.header import Header

# ==========================================
# 1. 外部通知模組整合 (Line & Email)
# ==========================================

# 試圖導入 Line 訊息模組
try:
    from line_messaging import send_line_message
except ModuleNotFoundError:
    print("⚠️ 找不到 line_messaging 模組，將改為僅在日誌中輸出警告。")
    def send_line_message(msg):
        print(f"[Line 模擬器] {msg}")

def send_email_notification(subject, content_text):
    """透過 Gmail SMTP 自動發送電子郵件通知"""
    email_user = os.getenv("EMAIL_USER")
    email_password = os.getenv("EMAIL_PASSWORD")
    
    # 預設收件者為發信者本人，若 config 內有其他設定可在此擴充
    recipients = [email_user] 
    
    if not email_user or not email_password:
        print("⚠️ 未偵測到 EMAIL_USER 或 EMAIL_PASSWORD 環境變數，跳過 Email 發送。")
        return False
        
    try:
        # 建立郵件主體
        msg = MIMEText(content_text, 'plain', 'utf-8')
        msg['From'] = Header(f"台股自動監控機器人 <{email_user}>", 'utf-8')
        msg['To'] = Header(", ".join(recipients), 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        
        # 連線至 Gmail SMTP 伺服器 (使用 TLS 加密)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, email_password)
        server.sendmail(email_user, recipients, msg.as_string())
        server.quit()
        print("📨 [Email 系統] 策略決策報告已成功寄出！")
        return True
    except Exception as e:
        print(f"❌ [Email 系統] 發送郵件失敗: {e}")
        return False

# ==========================================
# 2. 資料路徑與初始化
# ==========================================
try:
    import config
    DATA_DIR = Path(config.DATA_DIR)
except ImportError:
    DATA_DIR = Path("./data")

INVENTORY_FILE = DATA_DIR / "庫存股票.xlsx"
MONITOR_FILE = DATA_DIR / "監控股票.xlsx"

def load_latest_close_prices():
    """讀取最新的台股合併收盤價 CSV 檔案"""
    csv_files = glob.glob(str(DATA_DIR / "台股每日收盤價_*.csv"))
    if not csv_files:
        print("❌ 找不到任何收盤價 CSV 檔案。")
        return {}, None
    
    latest_file = max(csv_files, key=os.path.basename)
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(latest_file))
    report_date = date_match.group(1) if date_match else "未知日期"
    
    price_map = {}
    try:
        try: df = pd.read_csv(latest_file, encoding='cp950')
        except: df = pd.read_csv(latest_file, encoding='utf-8')
        
        name_col, close_col = None, None
        for col in df.columns:
            c = str(col).strip()
            if any(x in c for x in ['名稱', '證券名稱', '股票名稱']): name_col = col
            if any(x in c for x in ['收盤', '收盤價', 'Close']): close_col = col
            
        if name_col and close_col:
            for _, row in df.iterrows():
                name = str(row[name_col]).strip()
                try:
                    p_str = str(row[close_col]).replace(',', '').strip()
                    price_map[name] = float(p_str)
                except: pass
    except Exception as e:
        print(f"❌ 讀取收盤價明細失敗: {e}")
        
    return price_map, report_date

def check_inventory_and_monitor():
    price_map, report_date = load_latest_close_prices()
    if not price_map:
        return
    
    print(f"💡 [基準日: {report_date}] 開始執行大波段策略邏輯監控...")
    alerts = []
    
    # ==========================================
    # 核心邏輯一：庫存股票大波段「移動停利」追蹤
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
                
                current_price = price_map.get(name, 0.0)
                
                if current_price > 0:
                    if current_price > highest_price:
                        highest_price = current_price
                        print(f"🚀 {name} 創波段新高！更新最高價為: {highest_price}")
                    
                    sell_trigger_price = highest_price * (1 - (trail_pct / 100))
                    
                    if current_price <= sell_trigger_price:
                        alerts.append(f"⚠️ [庫存賣出預警] {name}\n  - 今日收盤: {current_price}\n  - 波段高點: {highest_price}\n  - 已回撤超過 {trail_pct}%\n  - 觸發移動停利線 ({sell_trigger_price:.1f})，建議落袋為安！")
                    elif current_price <= base_stop:
                        alerts.append(f"🛑 [保本停損觸發] {name}\n  - 今日收盤: {current_price}\n  - 綜合平均成本: {avg_cost:.1f}\n  - 觸發保本停損底線 ({base_stop:.1f})，建議全數離場！")
                
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
    # 核心邏輯二：監控股票「買進基準線」比對
    # ==========================================
    if MONITOR_FILE.exists():
        try:
            mon_df = pd.read_excel(MONITOR_FILE)
            if '買進目標價' in mon_df.columns:
                for _, row in mon_df.iterrows():
                    name = str(row['股票名稱']).strip()
                    target_price = row['買進目標價']
                    pe_limit = row['買進本益比上限'] if '買進本益比上限' in mon_df.columns else None
                    
                    if pd.isna(target_price): continue
                    
                    current_price = price_map.get(name, 0.0)
                    if current_price > 0 and current_price <= float(target_price):
                        pe_msg = f" (本益比門檻: {pe_limit}x)" if not pd.isna(pe_limit) else ""
                        alerts.append(f"🎯 [監控買進提示] {name}\n  - 今日收盤: {current_price}\n  - 買進目標價: {target_price}{pe_msg}\n  - 股價已修正至大波段安全邊際買點，建議分批佈局！")
        except Exception as e:
            print(f"❌ 處理監控股票買進比對時發生錯誤: {e}")

# ==========================================
    # 3. 雙通道通知發送 (Line + Email)
    # ==========================================
    if alerts:
        email_subject = f"📊 【台股大波段策略決策報告】基準日: {report_date}"
        full_message = f"報告基準日: {report_date}\n\n系統偵測到以下股票已觸發策略基準線：\n\n" + "\n---------------------\n".join(alerts)
        
        # 通道一：發送 Line (適應您 line_messaging.py 的 user_id 與 message 雙參數架構)
        line_msg = "\n" + email_subject + "\n" + "\n---------------------\n".join(alerts)
        try:
            # 從環境變數或 config 讀取 Line User ID，若無則傳入空字串或 None
            line_user_id = os.getenv("LINE_USER_ID") or getattr(config, 'LINE_USER_ID', '')
            send_line_message(line_user_id, line_msg)
        except TypeError:
            # 防呆：如果您的函式參數順序不同，自動嘗試另一種呼叫方式
            try: send_line_message(line_msg)
            except Exception as le: print(f"⚠️ Line 發送失敗: {le}")
        except Exception as e:
            print(f"⚠️ Line 發送時發生未預期錯誤: {e}")
        
        # 通道二：發送 Email 通知
        send_email_notification(email_subject, full_message)        
        print("🔔 雙通道策略通知已執行完畢。")
    else:
        print("✅ 今日全數個股皆處於安全波段中，未觸發任何買賣基準線。")

if __name__ == "__main__":
    check_inventory_and_monitor()
