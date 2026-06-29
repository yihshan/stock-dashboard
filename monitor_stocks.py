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
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

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

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def clean_price(val: Any) -> float:
    if pd.isna(val): return np.nan
    s = str(val).replace(',', '').strip()
    if s in ['', '----', '--', '-', 'null', 'None', 'nil']: return np.nan
    try: return float(s)
    except ValueError: return np.nan


def clean_stock_id(val: Any) -> str:
    if pd.isna(val): return ""
    s = str(val).strip()
    s = re.sub(r'[="\'\s]', '', s)
    s = s.split('.')[0]
    if s.isdigit() and len(s) < 4:
        return s.zfill(4)
    return s


def clean_stock_name(val: Any) -> str:
    if pd.isna(val): return ""
    s = str(val).strip()
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\[.*?\]', '', s)
    return s.replace(' ', '').replace(' ', '')


class DateDataParser:
    @staticmethod
    def parse_generic_date(val: Any) -> Optional[date]:
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
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.report_date = str(date.today())
        self.master_df = pd.DataFrame()
        self._initialize_database()

    def _initialize_database(self) -> None:
        csv_files = glob.glob(str(self.data_dir / "*.csv"))
        if not csv_files: return
            
        daily_files = glob.glob(str(self.data_dir / "台股每日收盤價_*.csv"))
        if daily_files:
            latest_file = max(daily_files, key=os.path.basename)
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(latest_file))
            if date_match: self.report_date = date_match.group(1)

        logger.info(f"💾 正在解析 {len(csv_files)} 個政府網站歷史收盤 CSV...")
        raw_data_list = []
        
        for file_path in csv_files:
            file_date = DateDataParser.extract_date_from_filename(os.path.basename(file_path))
            try:
                try: df = pd.read_csv(file_path, encoding='utf-8-sig')
                except: df = pd.read_csv(file_path, encoding='cp950')
                
                date_col, id_col, name_col, close_col, high_col, low_col = None, None, None, None, None, None
                
                for col in df.columns:
                    c = str(col).strip()
                    if any(x in c for x in ['日期', 'Date', '年月日', '成交日期']): date_col = col
                    if any(x in c for x in ['證券代號', '股票代號', '代號', 'Code', '證券編號']): id_col = col
                    if any(x in c for x in ['證券名稱', '股票名稱', '名稱', 'Name']): name_col = col
                    if any(x in c for x in ['收盤價', '收盤', 'Close', 'ClosingPrice']): close_col = col
                    if any(x in c for x in ['最高價', '最高', 'High', 'HighestPrice']): high_col = col
                    if any(x in c for x in ['最低價', '最低', 'Low', 'LowestPrice']): low_col = col
                
                if not close_col: continue
                    
                for _, row in df.iterrows():
                    row_date = file_date
                    if not row_date and date_col:
                        row_date = DateDataParser.parse_generic_date(row[date_col])
                    if not row_date: continue
                        
                    s_id = clean_stock_id(row[id_col]) if id_col else ""
                    s_name = clean_stock_name(row[name_col]) if name_col else ""
                    if not s_id and not s_name: continue
                        
                    c_val = clean_price(row[close_col])
                    if pd.isna(c_val): continue 
                        
                    h_val = clean_price(row[high_col]) if high_col else c_val
                    l_val = clean_price(row[low_col]) if low_col else c_val
                    
                    raw_data_list.append({
                        'id': s_id, 'name': s_name, 'Date': row_date, 
                        'Close': c_val, 'High': h_val, 'Low': l_val
                    })
            except Exception: continue
                
        if raw_data_list:
            self.master_df = pd.DataFrame(raw_data_list)
        logger.info(f"✅ 政府歷史數據讀取完畢，總紀錄數: {len(self.master_df)}")

    def get_history(self, stock_id: str, stock_name: str) -> pd.DataFrame:
        target_id = clean_stock_id(stock_id)
        target_name = clean_stock_name(stock_name)
        
        if self.master_df.empty: return pd.DataFrame()
        
        if target_id:
            df_res = self.master_df[self.master_df['id'] == target_id]
        else:
            cond_exact = (self.master_df['name'] == target_name)
            cond_remap = (self.master_df['name'] == f"{target_name}光電") | (self.master_df['name'] == f"{target_name}科技")
            df_res = self.master_df[cond_exact | cond_remap]
            
        if df_res.empty: return pd.DataFrame()
        
        return df_res.sort_values('Date').drop_duplicates('Date').reset_index(drop=True)


class MarketIndicatorService:
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
        df = df.sort_values('Date')
        ema12 = df['Close'].ewm(span=12, adjust=True).mean()
        ema26 = df['Close'].ewm(span=26, adjust=True).mean()
        dif = ema12 - ema26
        macd_line = dif.ewm(span=9, adjust=True).mean()
        osc = dif - macd_line
        return dif, macd_line, osc

    @staticmethod
    def check_macro_regime() -> Tuple[bool, str]:
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
            return is_bull, f"今日加權指數收盤 {latest_close:.2f} 點，處於年線 ({ma200:.1f}) 之{'上' if is_bull else '下'} -> {status_text}"
        except Exception as e:
            return True, "⚠️ 總體環境連線異常，策略降級切換為【安全多頭環境】"


class NotificationService:
    def __init__(self):
        self.smtp_server = config.SMTP_SERVER
        self.smtp_port = config.SMTP_PORT
        self.email_user = config.EMAIL_USER
        self.email_password = config.EMAIL_PASSWORD
        self.recipients = config.RECIPIENTS
        self.bcc_recipients = getattr(config, 'BCC_RECIPIENTS', [])

    def send_line(self, alerts: List[Dict[str, str]], report_date: str) -> None:
        if not alerts: return
        token = getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', '')
        user_id = getattr(config, 'LINE_USER_ID', '')
        lines = []
        for a in alerts:
            lines.append(f"{a['icon']} [{a['type']}] {a['name']}\n- 今日收盤: {a['close']}\n- {a['line2']}\n- 說明: {a['desc']}")
        full_msg = f"\n📊 【台股雙多智慧策略決策報告】\n基準日: {report_date}\n" + "\n---------------------\n".join(lines)
        send_line_message(token, user_id, full_msg)

    def send_html_email(self, report_date: str, market_text: str, alerts: List[Dict[str, str]], all_stocks: List[Dict[str, Any]]) -> None:
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 台股雙多智慧策略決策報告與診斷總覽 - {report_date}"
        msg['From'] = self.email_user
        msg['To'] = ", ".join(self.recipients)

        if alerts:
            alert_html = (
                "<div style='border-left: 4px solid #f6ad55; padding-left: 15px; margin-bottom: 25px;'> "
                "<h3 style='color: #dd6b20; font-size: 18px; margin-top: 0; margin-bottom: 15px;'>⚠️ 智慧多因子策略買賣觸發提示</h3>"
                "<ul style='list-style-type: disc; padding-left: 20px; line-height: 1.8; color: #2d3748; font-size: 14px;'>"
            )
            for a in alerts:
                alert_html += (
                    f"<li style='margin-bottom: 12px; list-style-type: disc;'>"
                    f"<b>{a['icon']} [{a['type']}] {a['name']}</b><br>"
                    f"&nbsp;&nbsp;- 今日收盤: {a['close']}<br>"
                    f"&nbsp;&nbsp;- {a['line2']}<br>"
                    f"&nbsp;&nbsp;- 說明：{a['desc']}"
                    f"</li>"
                )
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

            s_name, s_id, s_type, s_close, s_target_or_cost, s_status = s['name'], s['id'], s['type'], f"{s['close']:.2f}", s['target_or_cost'], s['status']

            table_rows += (
                f"<tr>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><b>{s_name}</b><br><small style='color:#718096;'>{s_id}</small></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><span style='background-color:{type_bg}; color:{type_color}; padding:3px 8px; border-radius:4px; font-size:12px;'>{s_type}</span></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'><b>{s_close}</b></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right; {k_style}'>{k_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{d_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{osc_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center; font-size:13px;'>{s_target_or_cost}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center; font-size:13px; {status_style}'>{s_status}</td>"
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
                for t in (self.recipients + self.bcc_recipients): server.sendmail(self.email_user, t, msg.as_string())
            logger.info("📨 智慧決策報告已成功寄出。")
        except Exception as e: logger.error(f"發送 Email 失敗: {e}")


class StrategyOrchestrator:
    def __init__(self):
        self.repo = StockDataRepository(BASE_DIR)
        self.notifier = NotificationService()

    def execute_pipeline(self) -> None:
        is_bull_market, market_text = MarketIndicatorService.check_macro_regime()
        structured_alerts: List[Dict[str, str]] = []
        all_stocks_output: List[Dict[str, Any]] = []
        global_stock_pool: Dict[str, Dict[str, Any]] = {}
        
        core_inventory_symbols = {
            '2330', '3017', '3131', '3443', '6442', '3324', '6510', '3563', 
            '7751', '2308', '7734', '3363', '0056', '00919', '00762', '00990A', '00985A', '00648R', '00724B'
        }

        # 🟢 修正核心：2026-2027 年「今明兩年預估股利與資產類別殖利率估值大腦」
        dividend_valuation_presets = {
            '2330': {'div_2026': 36.0, 'div_2027': 44.0, 'target_yield': 1.8},  # 台積電：成長股溢價，殖利率設 1.8%
            '2308': {'div_2026': 26.0, 'div_2027': 34.0, 'target_yield': 1.8},  # 台達電：調升估值，目標價推算約 1666 元
            '3008': {'div_2026': 95.0, 'div_2027': 115.0, 'target_yield': 2.4}, # 大立光：目標價推算約 4375 元
            '3017': {'div_2026': 32.0, 'div_2027': 42.0, 'target_yield': 1.8},  # 奇鋐
            '3131': {'div_2026': 50.0, 'div_2027': 65.0, 'target_yield': 1.7},  # 弘塑
            '3443': {'div_2026': 45.0, 'div_2027': 60.0, 'target_yield': 1.3},  # 創意
            '6442': {'div_2026': 18.0, 'div_2027': 26.0, 'target_yield': 1.6},  # 光聖
            '3324': {'div_2026': 16.0, 'div_2027': 24.0, 'target_yield': 1.9},  # 雙鴻
            '6510': {'div_2026': 35.0, 'div_2027': 50.0, 'target_yield': 1.6},  # 精測
            '3563': {'div_2026': 12.0, 'div_2027': 18.0, 'target_yield': 2.2},  # 牧德
            '7751': {'div_2026': 20.0, 'div_2027': 30.0, 'target_yield': 1.9},  # 竑騰
            '7734': {'div_2026': 55.0, 'div_2027': 75.0, 'target_yield': 2.1},  # 印能科技
            '8299': {'div_2026': 45.0, 'div_2027': 65.0, 'target_yield': 2.3},  # 群聯
            '8210': {'div_2026': 24.0, 'div_2027': 34.0, 'target_yield': 2.2},  # 勤誠
        }

        # 1. 處理現有庫存移動停利與智慧停損
        if INVENTORY_FILE.exists():
            try:
                inv_df = pd.read_excel(INVENTORY_FILE)
                
                if inv_df.empty or len(inv_df.columns) == 0:
                    logger.warning("⚠️ 庫存股票.xlsx 目前無內容，跳過持股風控運算。")
                else:
                    updated_rows = []
                    inv_name_col = next((c for c in inv_df.columns if '名稱' in str(c) or '股票' in str(c)), inv_df.columns[0])
                    inv_id_col = next((c for c in inv_df.columns if '代號' in str(c) or 'Code' in str(c)), None)
                    inv_shares_col = next((c for c in inv_df.columns if '股數' in str(c) or '張數' in str(c)), None)
                    inv_cost_col = next((c for c in inv_df.columns if '成本' in str(c) or '價位' in str(c)), None)
                    
                    inv_id_map = {}
                    if inv_id_col:
                        inv_id_map = pd.Series(inv_df[inv_id_col].astype(str).values, index=inv_df[inv_name_col].astype(str)).to_dict()

                    for name, group in inv_df.groupby(inv_name_col):
                        name = str(name).strip()
                        total_shares = group[inv_shares_col].sum() if inv_shares_col else 1
                        if total_shares <= 0: continue
                        
                        avg_cost = (group[inv_cost_col] * group[inv_shares_col]).sum() / total_shares if inv_cost_col and inv_shares_col else 0.0
                        trail_pct = group['移動停利百分比(%)'].dropna().head(1).values if '移動停利百分比(%)' in inv_df.columns else []
                        trail_pct = float(trail_pct[0]) if len(trail_pct) > 0 and not pd.isna(trail_pct[0]) else 15.0
                        base_stop = avg_cost 
                        
                        highest_record = group['波段最高價'].dropna().head(1).values if '波段最高價' in inv_df.columns else []
                        highest_price = float(highest_record[0]) if len(highest_record) > 0 and not pd.isna(highest_record[0]) else 0.0
                        
                        s_id_target = clean_stock_id(inv_id_map.get(name, ""))
                        hist_df = self.repo.get_history(s_id_target, name)
                        if hist_df.empty: continue
                        
                        latest = hist_df.iloc[-1]
                        current_price = latest['Close']
                        s_id = latest['id'] if 'id' in latest and latest['id'] else s_id_target
                        
                        if highest_price == 0.0 or pd.isna(highest_price): highest_price = hist_df['Close'].max()

                        status_str = "✅ 持股安全"
                        if current_price > highest_price: highest_price = current_price
                        sell_trigger_price = highest_price * (1 - (trail_pct / 100))
                        
                        c_p_str = f"{current_price:.2f}"
                        a_c_str = f"{avg_cost:.2f}"
                        h_p_str = f"{highest_price:.2f}"
                        s_t_str = f"{sell_trigger_price:.2f}"

                        if current_price <= base_stop:
                            if is_bull_market:
                                structured_alerts.append({
                                    'icon': '🔄', 'type': '智慧緩衝：暫緩停損', 'name': name, 'close': c_p_str,
                                    'line2': f"平均成本: {a_c_str}",
                                    'desc': "大盤加權指數處於強勢多頭格局，此回檔建議暫緩盲目砍單。"
                                })
                                status_str = "🔄 智慧緩衝"
                            else:
                                structured_alerts.append({
                                    'icon': '🛑', 'type': '鐵律清倉停損', 'name': name, 'close': c_p_str,
                                    'line2': f"平均成本: {a_c_str}",
                                    'desc': "大盤確認走空，請嚴守資金紀律全數清倉！"
                                })
                                status_str = "🛑 鐵律停損"
                        elif current_price <= sell_trigger_price:
                            structured_alerts.append({
                                'icon': '⚠️', 'type': '庫存移動停利', 'name': name, 'close': c_p_str,
                                'line2': f"波段最高: {h_p_str}",
                                'desc': f"觸發移動停利線 ({s_t_str})，建議獲利落袋。"
                            })
                            status_str = "⚠️ 移動停利"

                        k_val, d_val, osc_val = np.nan, np.nan, np.nan
                        if len(hist_df) >= 9:
                            df_idx = MarketIndicatorService.calculate_kd9(hist_df)
                            _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                            latest_idx = df_idx.iloc[-1]
                            k_val, d_val, osc_val = latest_idx['K'], latest_idx['D'], osc_s.iloc[-1]

                        stock_res = {
                            'name': name, 'id': s_id, 'type': '庫存股', 'close': current_price,
                            'k': k_val, 'd': d_val, 'osc': osc_val,
                            'target_or_cost': f"成本: {avg_cost:.2f}",
                            'status': status_str
                        }
                        all_stocks_output.append(stock_res)
                        global_stock_pool[name] = stock_res
                        
                        for _, row in group.iterrows():
                            row_copy = row.copy()
                            if '移動停利百分比(%)' in inv_df.columns: row_copy['移動停利百分比(%)'] = trail_pct
                            if '波段最高價' in inv_df.columns: row_copy['波段最高價'] = highest_price
                            if '保本停損價' in inv_df.columns: row_copy['保本停損價'] = base_stop
                            updated_rows.append(row_copy)
                            
                    if updated_rows:
                        pd.DataFrame(updated_rows).to_excel(INVENTORY_FILE, index=False)
            except Exception as e: logger.error(f"庫存智慧風控計算失敗: {e}")

        # 2. 處理觀察/監控名單
        if MONITOR_FILE.exists():
            try:
                monitor_df = pd.read_excel(MONITOR_FILE)
                if monitor_df.empty or len(monitor_df.columns) == 0:
                    logger.warning("⚠️ 監控股票.xlsx 內容為空。")
                else:
                    name_col = next((col for col in monitor_df.columns if '名稱' in str(col)), monitor_df.columns[0])
                    id_col = next((col for col in monitor_df.columns if '代號' in str(col)), None)
                    
                    for _, row in monitor_df.iterrows():
                        name = str(row[name_col]).strip()
                        raw_id = str(row[id_col]).strip() if id_col and pd.notna(row[id_col]) else ""
                        s_id = clean_stock_id(raw_id)
                        if not s_id:
                            match = re.search(r'\((\d+)\)', name)
                            s_id = match.group(1) if match else ""
                        if not name or name == 'nan': continue
                        
                        if s_id in core_inventory_symbols and name not in global_stock_pool:
                            logger.info(f"🛡️ [防禦性攔截] 偵測到庫存核心資產 {name}({s_id}) 被錯誤落入監控流程，已自動攔截隔離。")
                            continue
                            
                        df = self.repo.get_history(s_id, name)
                        if df.empty: continue
                            
                        latest = df.iloc[-1]
                        current_price = latest['Close']
                        
                        # 🟢 根據今明兩年（2026-2027）配息金額與自適應要求殖利率動態計算目標價
                        if s_id in dividend_valuation_presets:
                            cfg = dividend_valuation_presets[s_id]
                            avg_dividend = (cfg['div_2026'] + cfg['div_2027']) / 2
                            target_price = avg_dividend / (cfg['target_yield'] / 100)
                        else:
                            # 備援防禦：若無預設股利資料，則以最新市價打 85 折做為安全邊際買進目標價
                            target_price = current_price * 0.85
                            
                        diff_pct = ((current_price - target_price) / target_price) * 100
                        if current_price <= target_price:
                            structured_alerts.append({
                                'icon': '🎯', 'type': '監控買進提示', 'name': name,
                                'close': f"{current_price:.2f}",
                                'line2': f"買進目標價: {target_price:.2f}",
                                'desc': "已達大波段安全安全邊際，建議分批佈局。"
                            })
                            status_str = "🎯 已達買點"
                        else:
                            status_str = f"溢價 {diff_pct:.1f}%"

                        k_val, d_val, osc_val = np.nan, np.nan, np.nan
                        if len(df) >= 9:
                            df_idx = MarketIndicatorService.calculate_kd9(df)
                            _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                            latest_idx = df_idx.iloc[-1]
                            k_val, d_val, osc_val = latest_idx['K'], latest_idx['D'], osc_s.iloc[-1]
                            
                        stock_res = {
                            'name': name, 'id': s_id, 'type': '監控股', 'close': current_price,
                            'k': k_val, 'd': d_val, 'osc': osc_val,
                            'target_or_cost': f"目標: {target_price:.2f}",
                            'status': status_str
                        }
                        all_stocks_output.append(stock_res)
            except Exception as e: logger.error(f"解析監控 Excel 失敗: {e}")

        # 3. 發送通知
        self.notifier.send_line(structured_alerts, self.repo.report_date)
        if all_stocks_output:
            self.notifier.send_html_email(self.repo.report_date, market_text, structured_alerts, all_stocks_output)


if __name__ == "__main__":
    orchestrator = StrategyOrchestrator()
    orchestrator.execute_pipeline()
