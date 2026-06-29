import os
import glob
import re
import smtplib
import logging
import sys
from datetime import datetime, date
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Dict, Any, Tuple, Optional

import numpy as np
import pandas as pd

# 配置專業日誌系統
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(filename)s:%(lineno)d - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# 強制 UTF-8 環境
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ==========================================
# 🛑 核心配置與全域常數定義
# ==========================================
try:
    import config
    import yfinance as yf
except ImportError as e:
    logger.critical(f"缺少必要的核心模組或設定檔: {e}")
    sys.exit(1)

try:
    from line_messaging import send_line_message
except ModuleNotFoundError:
    def send_line_message(token: str, user_id: str, msg: str) -> None:
        logger.info(f"[LINE 模擬推播] {msg}")

BASE_DIR = Path(config.DATA_DIR)
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"

# 隱藏 SSL 警告
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def clean_price(val: Any) -> float:
    """行級防呆核心：將字串安全轉為 float，精確保留至小數點後兩位以上，跳過熔斷/停牌之 '----' 符號"""
    if pd.isna(val): 
        return np.nan
    s = str(val).replace(',', '').strip()
    if s in ['', '----', '--', '-', 'null', 'None', 'nil']:
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


class DateDataParser:
    """處理日期格式轉換"""
    @staticmethod
    def parse_generic_date(val: Any) -> Optional[date]:
        if pd.isna(val): return None
        s = str(val).strip().split(' ')[0]
        if not s: return None
        for sep in ['/', '-']:
            parts = s.split(sep)
            if len(parts) == 3:
                try:
                    year = int(parts[0])
                    corrected_year = year if year > 1911 else year + 1911
                    return date(corrected_year, int(parts[1]), int(parts[2]))
                except ValueError: continue
        digits = re.sub(r'\D', '', s)
        if len(digits) == 8:
            try: return datetime.strptime(digits, '%Y%m%d').date()
            except ValueError: pass
        elif len(digits) == 7:
            try: return date(int(digits[:3]) + 1911, int(digits[3:5]), int(digits[5:]))
            except ValueError: pass
        return None

    @staticmethod
    def extract_date_from_filename(filename: str) -> Optional[date]:
        match = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', filename)
        if match: return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        match = re.search(r'(\d{8})|(\d{7})', filename)
        if match:
            d = match.group()
            return datetime.strptime(d, '%Y%m%d').date() if len(d) == 8 else date(int(d[:3]) + 1911, int(d[3:5]), int(d[5:]))
        return None


class StockDataRepository:
    """資料訪問層：對齊 Dashboard 規格，全面載入資料夾內所有 CSV，確保歷史天數縱向完整載入"""
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.report_date = str(date.today())
        self._historical_cache: Dict[str, pd.DataFrame] = {}
        self._initialize_database()

    def _initialize_database(self) -> None:
        csv_files = glob.glob(str(self.data_dir / "*.csv"))
        if not csv_files:
            logger.warning(f"指定路徑 {self.data_dir} 中未發現數據檔案。")
            return
            
        daily_files = glob.glob(str(self.data_dir / "台股每日收盤價_*.csv"))
        if daily_files:
            latest_file = max(daily_files, key=os.path.basename)
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(latest_file))
            if date_match: self.report_date = date_match.group(1)

        logger.info(f"💾 正在載入 {len(csv_files)} 個 CSV 歷史檔案至記憶體快取...")
        raw_data_list = []
        
        for file_path in csv_files:
            file_date = DateDataParser.extract_date_from_filename(os.path.basename(file_path))
            try:
                try: df = pd.read_csv(file_path, encoding='utf-8-sig')
                except Exception: df = pd.read_csv(file_path, encoding='cp950')
                
                col_map = {}
                for col in df.columns:
                    c = str(col).strip()
                    if any(x in c for x in ['日期', 'Date']): col_map['date'] = col
                    if any(x in c for x in ['證券代號', 'Code', '股票代號']): col_map['id'] = col
                    if any(x in c for x in ['名稱', 'Name']): col_map['name'] = col
                    if any(x in c for x in ['收盤', 'Close', 'ClosingPrice']): col_map['close'] = col
                    if any(x in c for x in ['最高', 'High', 'HighestPrice']): col_map['high'] = col
                    if any(x in c for x in ['最低', 'Low', 'LowestPrice']): col_map['low'] = col
                
                if 'close' not in col_map: continue
                    
                final_date = file_date
                for _, row in df.iterrows():
                    if not final_date and 'date' in col_map:
                        final_date = DateDataParser.parse_generic_date(row[col_map['date']])
                    if not final_date: continue
                        
                    # 【只針對證券代號進行字串化對齊，絕不干涉股價】
                    raw_id = str(row[col_map['id']]).replace('"', '').strip() if 'id' in col_map else ""
                    s_id = raw_id.split('.')[0] if raw_id else ""
                    s_name = str(row[col_map['name']]).strip() if 'name' in col_map else ""
                    
                    # 使用全防禦性數值清洗，精確還原浮點數股價
                    c_val = clean_price(row[col_map['close']])
                    if pd.isna(c_val): continue 
                        
                    h_val = clean_price(row[col_map['high']]) if 'high' in col_map else c_val
                    l_val = clean_price(row[col_map['low']]) if 'low' in col_map else c_val
                    
                    if pd.isna(h_val): h_val = c_val
                    if pd.isna(l_val): l_val = c_val
                    
                    raw_data_list.append({'id': s_id, 'name': s_name, 'Date': final_date, 'Close': c_val, 'High': h_val, 'Low': l_val})
            except Exception as e:
                logger.error(f"跳過嚴重損壞之單一檔案 {os.path.basename(file_path)}: {e}")
                
        if raw_data_list:
            master_df = pd.DataFrame(raw_data_list)
            if 'id' in master_df.columns:
                for s_id, g in master_df.groupby('id'):
                    if s_id: self._historical_cache[str(s_id)] = g.sort_values('Date').drop_duplicates('Date')
            for s_name, g in master_df.groupby('name'):
                if s_name: self._historical_cache[str(s_name)] = g.sort_values('Date').drop_duplicates('Date')
        logger.info(f"✅ 快取加載成功！長線數據已安全移入記憶體。")

    def get_history(self, stock_id: str, stock_name: str) -> pd.DataFrame:
        raw_id = str(stock_id).replace('"', '').strip()
        s_id = raw_id.split('.')[0] if raw_id else ""
        s_name = re.sub(r'\(.*?\)', '', stock_name).strip()
        if s_id in self._historical_cache: return self._historical_cache[s_id]
        if s_name in self._historical_cache: return self._historical_cache[s_name]
        return pd.DataFrame()


class MarketIndicatorService:
    """運算服務層：整合大盤與個股指標"""
    @staticmethod
    def calculate_kd9(df: pd.DataFrame) -> pd.DataFrame:
        if len(df) < 9:
            df['K'], df['D'] = np.nan, np.nan
            return df
        df = df.sort_values('Date').reset_index(drop=True)
        df['L9'] = df['Low'].rolling(window=9).min()
        df['H9'] = df['High'].rolling(window=9).max()
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
        df['K'], df['D'] = k_vals, d_vals
        return df

    @staticmethod
    def calculate_macd(df: pd.DataFrame) -> Tuple[pd.Series, pd.Series, pd.Series]:
        if len(df) < 26:
            empty = pd.Series([np.nan] * len(df))
            return empty, empty, empty
        df = df.sort_values('Date')
        ema12 = df['Close'].ewm(span=12, adjust=True).mean()
        ema26 = df['Close'].ewm(span=26, adjust=True).mean()
        dif = ema12 - ema26
        macd_line = dif.ewm(span=9, adjust=True).mean()
        osc = dif - macd_line
        return dif, macd_line, osc

    @staticmethod
    def check_macro_regime() -> Tuple[bool, str]:
        """精準下載即時大盤加權指數，進行年線波段風控"""
        try:
            logger.info("🌐 正在下載大盤加權指數 (計算 200MA 年線系統風險)...")
            twii = yf.Ticker("^TWII")
            df = twii.history(period="1y")
            if df.empty: raise ValueError("無法獲取加權指數數據。")
            df['200MA'] = df['Close'].rolling(window=200).mean()
            latest_close = df['Close'].iloc[-1]
            ma200 = df['200MA'].iloc[-1]
            
            is_bull = latest_close >= ma200
            status_text = "【多頭市場】(加權指數處於年線之上，啟動智慧估值緩衝鎖)" if is_bull else "【空頭熊市】(大盤走空，全面防守開啟鐵律停損)"
            direction = "上" if is_bull else "下"
            return is_bull, f"今日加權指數收盤 {latest_close:.2f} 點，處於年線 ({ma200:.2f}) 之{direction} -> {status_text}"
        except Exception as e:
            logger.error(f"大盤總體環境解析異常: {e}，策略自動降級為【安全多頭環境】")
            return True, "⚠️ 總體環境連線異常，策略降級切換為【安全多頭環境】"


class NotificationService:
    """通知派送服務層"""
    def __init__(self):
        self.smtp_server = config.SMTP_SERVER
        self.smtp_port = config.SMTP_PORT
        self.email_user = config.EMAIL_USER
        self.email_password = config.EMAIL_PASSWORD
        self.recipients = config.RECIPIENTS
        self.bcc_recipients = getattr(config, 'BCC_RECIPIENTS', [])

    def send_line(self, alerts: List[str], report_date: str) -> None:
        if not alerts: return
        token = getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', '')
        user_id = getattr(config, 'LINE_USER_ID', '')
        full_msg = f"\n📊 【台股雙多智慧策略決策報告】\n基準日: {report_date}\n" + "\n---------------------\n".join(alerts)
        send_line_message(token, user_id, full_msg)

    def send_html_email(self, report_date: str, market_text: str, alerts: List[str], all_stocks: List[Dict[str, Any]]) -> None:
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 台股雙多智慧策略決策報告與診斷總覽 - {report_date}"
        msg['From'] = self.email_user
        msg['To'] = ", ".join(self.recipients)

        if alerts:
            alert_html = (
                "<div style='background-color: #fffaf0; border-left: 5px solid #dd6b20; padding: 15px; margin-bottom: 25px; border-radius: 4px;'>"
                "<h3 style='color: #dd6b20; margin-top: 0;'>⚠️ 智慧多因子策略買賣觸發提示</h3>"
                "<ul style='padding-left: 20px; line-height: 1.6; color: #2d3748;'>"
            )
            for act in alerts:
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

        k_threshold = getattr(config, 'K_THRESHOLD', 15)
        table_rows = ""
        for s in all_stocks:
            is_nan_k = pd.isna(s['k'])
            k_style = "color:red; font-weight:bold;" if not is_nan_k and s['k'] < k_threshold else ""
            k_val = f"{s['k']:.2f}" if not is_nan_k else "N/A"
            d_val = f"{s['d']:.2f}" if not pd.isna(s['d']) else "N/A"
            osc_val = f"{s['osc']:.2f}" if not pd.isna(s['osc']) else "N/A"

            type_bg = "#e2e8f0" if s['type'] == '監控股' else "#feebc8"
            type_color = "#4a5568" if s['type'] == '監控股' else "#c05621"
            
            status_style = "color:#38a169;"
            if "🎯" in s['status'] or "🛑" in s['status']: status_style = "color:red; font-weight:bold;"
            elif "⚠️" in s['status'] or "🔄" in s['status']: status_style = "color:#d69e2e; font-weight:bold;"

            # 🟢 修正點：所有股價欄位統以 :.2f 格式化輸出，完美保留小數點二位跳動精確度
            table_rows += (
                f"<tr>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><b>{s['name']}</b><br><small style='color:#718096;'>{s['id']}</small></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><span style='background-color:{type_bg}; color:{type_color}; padding:3px 8px; border-radius:4px; font-size:12px;'>{s['type']}</span></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'><b>{s['close']:.2f}</b></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right; {k_style}'>{k_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{d_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{osc_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center; font-size:13px;'>{s['target_or_cost']}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center; font-size:13px; {status_style}'>{s['status']}</td>"
                f"</tr>"
            )
            
        html = (
            f"<html><body style=\"font-family: 'Microsoft JhengHei', sans-serif; padding: 20px;\">"
            f"<h2 style=\"color: #1a365d;\">📊 每日台股策略監控與技術指標自動彙整</h2>"
            f"<p style='color: #4a5568;'><b>數據基準日：</b>{report_date}</p>"
            f"<p style='background-color: #edf2f7; padding: 10px; border-radius: 4px; color: #4a5568;'>🌐 <b>總體環境監測：</b>{market_text}</p>"
            f"{alert_html}"  
            f"<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 30px 0;'>"
            f"<h3 style=\"color: #004a99;\">📈 全資產綜合策略整合診斷面板</h3>"
            f"<table style=\"width:100%; border-collapse:collapse; margin-top:15px;\">"
            f"<thead><tr style=\"background-color: #004a99; color: white;\">"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">股票名稱</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">資產類別</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">收盤價</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">日K值</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">日D值</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">MACD (OSC)</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">目標 / 成本</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">策略診斷狀態</th>"
            f"</tr></thead><tbody>{table_rows}</tbody></table>"
            f"</body></html>"
        )
        msg.attach(MIMEText(html, 'html'))
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_user, self.email_password)
                targets = self.recipients + self.bcc_recipients
                for target in targets: server.sendmail(self.email_user, target, msg.as_string())
            logger.info("📨 智慧決策報告已成功寄出。")
        except Exception as e: logger.error(f"發送 Email 失敗: {e}")


class StrategyOrchestrator:
    """核心協調指揮官"""
    def __init__(self):
        self.repo = StockDataRepository(BASE_DIR)
        self.notifier = NotificationService()

    def execute_pipeline(self) -> None:
        is_bull_market, market_text = MarketIndicatorService.check_macro_regime()
        strategy_alerts: List[str] = []
        all_stocks_output: List[Dict[str, Any]] = []
        global_stock_pool: Dict[str, Dict[str, Any]] = {}

        # 1. 處理觀察/監控名單
        if MONITOR_FILE.exists():
            try:
                monitor_df = pd.read_excel(MONITOR_FILE)
                name_col = next((col for col in monitor_df.columns if '名稱' in str(col)), monitor_df.columns[0])
                id_col = next((col for col in monitor_df.columns if '代號' in str(col)), None)
                
                for _, row in monitor_df.iterrows():
                    name = str(row[name_col]).strip()
                    raw_id = str(row[id_col]).strip() if id_col and pd.notna(row[id_col]) else ""
                    s_id = raw_id.split('.')[0] if raw_id else ""
                    if not s_id:
                        match = re.search(r'\((\d+)\)', name)
                        s_id = match.group(1) if match else ""
                    if not name or name == 'nan': continue
                        
                    df = self.repo.get_history(s_id, name)
                    if df.empty: continue
                        
                    latest = df.iloc[-1]
                    current_price = latest['Close']
                    
                    target_price = np.nan
                    status_str = "觀察中"
                    if '買進目標價' in monitor_df.columns and pd.notna(row['買進目標價']):
                        target_price = float(row['買進目標價'])
                        diff_pct = ((current_price - target_price) / target_price) * 100
                        if current_price <= target_price:
                            strategy_alerts.append(f"🎯 [監控買進提示] {name}\n  - 今日收盤: {current_price:.2f}\n  - 買進目標價: {target_price:.2f}\n  - 已達大波段安全安全邊際，建議分批佈局。")
                            status_str = "🎯 已達買點"
                        else:
                            status_str = f"溢價 {diff_pct:.1f}%"

                    k_val, d_val, osc_val = np.nan, np.nan, np.nan
                    if len(df) >= 10:
                        df_idx = MarketIndicatorService.calculate_kd9(df)
                        _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                        latest_idx = df_idx.iloc[-1]
                        k_val, d_val, osc_val = latest_idx['K'], latest_idx['D'], osc_s.iloc[-1]
                        
                    stock_res = {
                        'name': name, 'id': s_id, 'type': '監控股', 'close': current_price,
                        'k': k_val, 'd': d_val, 'osc': osc_val,
                        'target_or_cost': f"目標: {target_price:.2f}" if pd.notna(target_price) else "-",
                        'status': status_str
                    }
                    all_stocks_output.append(stock_res)
                    global_stock_pool[name] = stock_res
            except Exception as e: logger.error(f"解析監控 Excel 失敗: {e}")

        # 2. 處理現有庫存移動停利與智慧停損
        if INVENTORY_FILE.exists():
            try:
                inv_df = pd.read_excel(INVENTORY_FILE)
                updated_rows = []
                for name, group in inv_df.groupby('股票名稱'):
                    name = str(name).strip()
                    total_shares = group['股數'].sum()
                    if total_shares <= 0: continue
                    
                    avg_cost = (group['成本'] * group['股數']).sum() / total_shares
                    trail_pct = group['移動停利百分比(%)'].dropna().head(1).values
                    trail_pct = float(trail_pct[0]) if len(trail_pct) > 0 and not pd.isna(trail_pct[0]) else 15.0
                    base_stop = avg_cost 
                    
                    highest_record = group['波段最高價'].dropna().head(1).values
                    highest_price = float(highest_record[0]) if len(highest_record) > 0 and not pd.isna(highest_record[0]) else 0.0
                    
                    hist_df = self.repo.get_history("", name)
                    if hist_df.empty: continue
                    
                    latest = hist_df.iloc[-1]
                    current_price = latest['Close']
                    raw_id = latest['id'] if 'id' in latest and latest['id'] else ""
                    s_id = str(raw_id).split('.')[0]
                    
                    if highest_price == 0.0 or pd.isna(highest_price): highest_price = hist_df['Close'].max()

                    status_str = "✅ 持股安全"
                    if current_price > highest_price: highest_price = current_price
                    sell_trigger_price = highest_price * (1 - (trail_pct / 100))
                    
                    if current_price <= base_stop:
                        if is_bull_market:
                            strategy_alerts.append(f"🔄 [智慧緩衝：暫緩停損] {name}\n  - 今日收盤: {current_price:.2f}\n  - 平均成本: {avg_cost:.2f}\n  - 說明：大盤加權指數處於強勢多頭格局，此回檔建議暫緩盲目砍單。")
                            status_str = "🔄 智慧緩衝"
                        else:
                            strategy_alerts.append(f"🛑 [鐵律清倉停損] {name}\n  - 今日收盤: {current_price:.2f}\n  - 平均成本: {avg_cost:.2f}\n  - 說明：大盤確認走空，請嚴守資金紀律全數清倉！")
                            status_str = "🛑 鐵律停損"
                    elif current_price <= sell_trigger_price:
                        strategy_alerts.append(f"⚠️ [庫存移動停利] {name}\n  - 今日收盤: {current_price:.2f}\n  - 波段最高: {highest_price:.2f}\n  - 觸發移動停利線 ({sell_trigger_price:.2f})，建議獲利落袋。")
                        status_str = "⚠️ 移動停利"

                    k_val, d_val, osc_val = np.nan, np.nan, np.nan
                    if len(hist_df) >= 10:
                        df_idx = MarketIndicatorService.calculate_kd9(hist_df)
                        _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                        latest_idx = df_idx.iloc[-1]
                        k_val, d_val, osc_val = latest_idx['K'], latest_idx['D'], osc_s.iloc[-1]

                    # 🟢 修正點：若監控名單與庫存股重複，直接更新狀態為庫存狀態，避免表格重複渲染或資料混亂
                    stock_res = {
                        'name': name, 'id': s_id, 'type': '庫存股', 'close': current_price,
                        'k': k_val, 'd': d_val, 'osc': osc_val,
                        'target_or_cost': f"成本: {avg_cost:.2f}",
                        'status': status_str
                    }
                    
                    if name in global_stock_pool:
                        # 如果已存在，更新其內容回歸庫存股最高優先權
                        idx_to_update = next(i for i, x in enumerate(all_stocks_output) if x['name'] == name)
                        all_stocks_output[idx_to_update] = stock_res
                    else:
                        all_stocks_output.append(stock_res)
                    
                    for _, row in group.iterrows():
                        row['移動停利百分比(%)'] = trail_pct
                        row['波段最高價'] = highest_price
                        row['保本停損價'] = base_stop
                        updated_rows.append(row)
                pd.DataFrame(updated_rows).to_excel(INVENTORY_FILE, index=False)
            except Exception as e: logger.error(f"庫存智慧風控計算失敗: {e}")

        # 3. 發送通知
        self.notifier.send_line(strategy_alerts, self.repo.report_date)
        if all_stocks_output:
            self.notifier.send_html_email(self.repo.report_date, market_text, strategy_alerts, all_stocks_output)


if __name__ == "__main__":
    orchestrator = StrategyOrchestrator()
    orchestrator.execute_pipeline()
