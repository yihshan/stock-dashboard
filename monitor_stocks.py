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

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = Path(config.DATA_DIR)
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"


# ==========================================
# 🟢 獨立通訊核心：LINE 官方一鍵全體廣播發送器
# ==========================================
def send_line_broadcast(token: str, msg: str) -> None:
    """
    使用 LINE Messaging API 廣播機制，不需要知道 User/Group ID，
    一鍵發送一對一私訊給所有加入官方帳號好友的夥伴，徹底根絕群組發言洗版痛點。
    """
    url = "https://api.line.me/v2/bot/message/broadcast"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }
    payload = {
        "messages": [
            {
                "type": "text",
                "text": msg
            }
        ]
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        if response.status_code != 200:
            logger.error(f"LINE 廣播推播失敗: {response.status_code} - {response.text}")
        else:
            logger.info("LINE 決策提示訊息已成功廣播給全體好友。")
    except Exception as e:
        logger.error(f"連線至 LINE 廣播 API 發生異常: {e}")


# ==========================================
# 工具函式區
# ==========================================
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
    if s.isdigit() and len(s) < 4: return s.zfill(4)
    return s

def clean_stock_name(val: Any) -> str:
    if pd.isna(val): return ""
    s = str(val).strip()
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\[.*?\]', '', s)
    return s.replace(' ', '')

def round_to_tw_tick(price: float) -> float:
    if pd.isna(price) or price <= 0: return 0.0
    if price < 10: return np.floor(price * 100) / 100
    elif price < 50: return np.floor(price * 20) / 20
    elif price < 100: return np.floor(price * 10) / 10
    elif price < 500: return np.floor(price * 2) / 2
    elif price < 1000: return np.floor(price)
    else: return np.floor(price / 5) * 5


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
                except Exception: continue
        digits = re.sub(r'\D', '', s)
        if len(digits) == 8:
            try: return datetime.strptime(digits, '%Y%m%d').date()
            except Exception: pass
        elif len(digits) == 7:
            try: return date(int(digits[:3]) + 1911, int(digits[3:5]), int(digits[5:]))
            except Exception: pass
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

        raw_data_list = []
        for file_path in csv_files:
            file_date = DateDataParser.extract_date_from_filename(os.path.basename(file_path))
            try:
                try: df = pd.read_csv(file_path, encoding='utf-8-sig')
                except Exception: df = pd.read_csv(file_path, encoding='cp950')
                
                date_col, id_col, name_col, close_col, high_col, low_col = None, None, None, None, None, None
                for col in df.columns:
                    c = str(col).strip()
                    if any(x in c for x in ['日期', 'Date', '年月日']): date_col = col
                    if any(x in c for x in ['證券代號', '股票代號', '代號', 'Code']): id_col = col
                    if any(x in c for x in ['證券名稱', '股票名稱', '名稱', 'Name']): name_col = col
                    if any(x in c for x in ['收盤價', '收盤', 'Close']): close_col = col
                    if any(x in c for x in ['最高價', '最高', 'High']): high_col = col
                    if any(x in c for x in ['最低價', '最低', 'Low']): low_col = col
                
                if not close_col: continue
                for _, row in df.iterrows():
                    row_date = file_date
                    if not row_date and date_col:
                        row_date = DateDataParser.parse_generic_date(row[date_col])
                    if not row_date: continue
                        
                    s_id = clean_stock_id(row[id_col]) if id_col else ""
                    s_name = clean_stock_name(row[name_col]) if name_col else ""
                    c_val = clean_price(row[close_col])
                    if pd.isna(c_val) or (not s_id and not s_name): continue
                    
                    raw_data_list.append({
                        'id': s_id, 'name': s_name, 'Date': row_date, 
                        'Close': c_val, 
                        'High': clean_price(row[high_col]) if high_col else c_val,
                        'Low': clean_price(row[low_col]) if low_col else c_val
                    })
            except Exception: continue
                
        if raw_data_list:
            self.master_df = pd.DataFrame(raw_data_list)

    def get_history(self, stock_id: str, stock_name: str) -> pd.DataFrame:
        if self.master_df.empty: return pd.DataFrame()
        target_id = clean_stock_id(stock_id)
        target_name = clean_stock_name(stock_name)
        
        if target_id: df_res = self.master_df[self.master_df['id'] == target_id]
        else: df_res = self.master_df[self.master_df['name'] == target_name]
            
        return df_res.sort_values('Date').drop_duplicates('Date').reset_index(drop=True)


class MarketIndicatorService:
    @staticmethod
    def calculate_kd9(df: pd.DataFrame) -> pd.DataFrame:
        if len(df) < 2:
            df = df.copy()
            df['K'], df['D'] = 50.0, 50.0
            return df
            
        df = df.sort_values('Date').reset_index(drop=True)
        df['L9'] = df['Low'].rolling(window=9, min_periods=1).min()
        df['H9'] = df['High'].rolling(window=9, min_periods=1).max()
        denom = df['H9'] - df['L9']
        df['RSV'] = 100 * (df['Close'] - df['L9']) / denom
        df.loc[denom == 0, 'RSV'] = 50.0
        
        k_vals, d_vals = [], []
        current_k, current_d = 50.0, 50.0
        for rsv in df['RSV']:
            if pd.isna(rsv): rsv = 50.0
            current_k = (2/3) * current_k + (1/3) * rsv
            current_d = (2/3) * current_d + (1/3) * current_k
            k_vals.append(current_k)
            d_vals.append(current_d)
        df['K'], df['D'] = k_vals, d_vals
        return df

    @staticmethod
    def calculate_macd(df: pd.DataFrame) -> Tuple[pd.Series, pd.Series, pd.Series]:
        df = df.sort_values('Date')
        ema12 = df['Close'].ewm(span=12, adjust=True, min_periods=1).mean()
        ema26 = df['Close'].ewm(span=26, adjust=True, min_periods=1).mean()
        dif = ema12 - ema26
        macd_line = dif.ewm(span=9, adjust=True, min_periods=1).mean()
        osc = dif - macd_line
        return dif, macd_line, osc

    @staticmethod
    def get_macro_status() -> Tuple[bool, str, float]:
        """
        🟢 功能升級：線性滑動防震引擎 (Linear Smoothing Engine)
        不再使用跳躍式階梯硬切，大盤與年線乖離率改採漸進式乘數，徹底消弭大立光等目標價暴振落差。
        """
        try:
            twii = yf.Ticker("^TWII")
            df = twii.history(period="1y")
            df['200MA'] = df['Close'].rolling(window=200).mean()
            current_close = df['Close'].iloc[-1]
            ma200 = df['200MA'].iloc[-1]
            
            is_bull = current_close >= ma200
            market_text = f"今日加權指數收盤 {current_close:.2f} 點，處於年線 ({ma200:.1f}) 之{'上' if is_bull else '下'}"
            
            # 計算大盤與年線的百分比乖離率
            bias = (current_close - ma200) / ma200
            
            # 線性連續滑動演算法：鎖定極值，其餘区間隨市場水溫自然流暢修正
            if bias > 0:
                multiplier = max(0.85, 1.0 - (bias * 0.5))
            else:
                multiplier = min(1.20, 1.0 - (bias * 1.5))
                
            return is_bull, market_text, round(multiplier, 3)
        except Exception:
            return True, "⚠️ 總體環境連線異常，策略降級切換為【安全多頭環境】", 1.00


class NotificationService:
    def __init__(self):
        self.smtp_server = config.SMTP_SERVER
        self.smtp_port = config.SMTP_PORT
        self.email_user = config.EMAIL_USER
        self.email_password = config.EMAIL_PASSWORD
        self.recipients = config.RECIPIENTS
        self.bcc_recipients = getattr(config, 'BCC_RECIPIENTS', [])

    def send_line(self, alerts: List[Dict[str, Any]], report_date: str) -> None:
        if not alerts: return
        lines = [f"{a['icon']} [{a['type']}] {a['name']}\n- 今日收盤: {a['close']}\n- {a['line2']}\n- 說明: {a['desc']}" for a in alerts]
        msg_body = f"\n📊 【台股智慧策略提示】\n基準日: {report_date}\n" + "\n------------\n".join(lines)
        
        token = getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', '')
        if token:
            send_line_broadcast(token, msg_body)
        else:
            logger.warning("未配置 LINE_CHANNEL_ACCESS_TOKEN，跳過廣播發送")

    def send_html_email(self, report_date: str, market_text: str, alerts: List[Dict[str, Any]], all_stocks: List[Dict[str, Any]], global_stock_pool: Dict[str, Any], triggered_exits: set, macro_multiplier: float, monitor_configs: Dict[str, Dict[str, Any]]) -> None:
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 台股智慧策略決策診斷報告 - {report_date}"
        msg['From'] = self.email_user
        msg['To'] = ", ".join(self.recipients)

        # 🟢 功能增強：頂部置入完全體「採取行動邏輯路徑說明表格」
        logic_explanation_html = """
        <div style='margin-bottom: 25px; background-color: #f8fafc; padding: 15px; border-left: 5px solid #2b6cb0; border-radius: 6px;'>
            <h3 style='margin-top: 0; margin-bottom: 12px; color: #2d3748; font-size: 15px;'>💡 【採取行動邏輯路徑】動態診斷狀態對照表</h3>
            <table style='width:100%; border-collapse:collapse; font-size:13px; border: 1px solid #e2e8f0; background-color: #ffffff;'>
                <thead>
                    <tr style='background-color: #edf2f7; color: #4a5568;'>
                        <th style='padding:8px 12px; border: 1px solid #e2e8f0; text-align:left;'>策略路徑狀態</th>
                        <th style='padding:8px 12px; border: 1px solid #e2e8f0; text-align:left;'>核心量化條件</th>
                        <th style='padding:8px 12px; border: 1px solid #e2e8f0; text-align:left;'>資金戰術與思維</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'><span style='color:#dd6b20; font-weight:bold;'>🎯 路徑 1：現價 ≦ 買點 (無持股)</span></td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>現價已回落至安全邊際區間，且帳戶無庫存部位</td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>已達絕對便宜價 ➜ <b>執行分批建立新倉</b></td>
                    </tr>
                    <tr>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'><span style='color:#e53e3e; font-weight:bold;'>🎯 路徑 2：現價 ≦ 買點 (有持股)</span></td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>現價低於動態估值甜蜜點，且帳戶已有基本部位</td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>擴大獲利安全邊際 ➜ <b>波段逢低分批加碼</b></td>
                    </tr>
                    <tr>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'><span style='color:#3182ce; font-weight:bold;'>⏳ 現價 ＞ 買點 (有持股)</span></td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>股價高於估值點、溢價偏高，但技術面或大盤環境安全</td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>高昂無安全邊際 ➜ <b>現有持股續抱、暫緩加碼</b></td>
                    </tr>
                    <tr>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'><span style='color:#4a5568; font-weight:bold;'>⏳ 現價 ＞ 買點 (無持股)</span></td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>標的並未出現回檔，不具備安全邊際優勢</td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>條件不滿足 ➜ <b>空倉耐心監控，切勿盲目追高</b></td>
                    </tr>
                    <tr>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'><span style='color:#718096; font-weight:bold;'>🛑 風控鎖定：策略執行出場中</span></td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>高檔跌破自訂移動防守線，或空頭環境下觸發停損線</td>
                        <td style='padding:8px 12px; border: 1px solid #e2e8f0;'>進入風控保護程序 ➜ <b>執行落袋或避險，嚴禁逆勢加碼</b></td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

        preset_matrix_rows = ""
        for c_name, cfg in monitor_configs.items():
            calc_target = cfg['dynamic_target']
            match_stock = next((s for s in all_stocks if s['name'] == c_name), None)
            current_price = match_stock['close'] if match_stock else None
            s_id = match_stock['id'] if match_stock else ""
            is_owned = (c_name in global_stock_pool)
            
            if c_name in triggered_exits:
                path_desc = "<span style='color:#718096; font-weight:bold;'>🛑 風控鎖定：策略執行出場中</span><br><small style='color:#718096;'>已觸發移動防守線 ➜ <b>請先執行落袋/避險，禁止加碼</b></small>"
                defense_desc = "⚠️ 技術面破位，進入風控保護程序"
            elif current_price and calc_target > 0 and current_price <= calc_target:
                if is_owned:
                    path_desc = "<span style='color:#e53e3e; font-weight:bold;'>🎯 路徑 2：現價 ≦ 買點 (有持股)</span><br><small style='color:#e53e3e;'>已達估值甜蜜區 ➜ <b>執行大波段逢低加碼</b></small>"
                else:
                    path_desc = "<span style='color:#dd6b20; font-weight:bold;'>🎯 路徑 1：現價 ≦ 買點 (無持股)</span><br><small style='color:#dd6b20;'>已達安全邊際 ➜ <b>執行分批建立新倉</b></small>"
                defense_desc = "🛡️ 策略已觸發，請依資金紀律分批建倉"
            else:
                if is_owned:
                    path_desc = "<span style='color:#3182ce;'>⏳ 現價 ＞ 買點 (有持股)</span><br><small style='color:#718096;'>資產價值高昂 ➜ <b>現有部位續抱、暫緩加碼</b></small>"
                    defense_desc = "🛡️ 多頭環境下享智慧緩衝持股防護網"
                else:
                    path_desc = "<span style='color:#4a5568;'>⏳ 現價 ＞ 買點 (無持股)</span><br><small style='color:#718096;'>未達安全邊際 ➜ <b>建倉條件不滿足，持續監控</b></small>"
                    defense_desc = "🛡️ 耐心等待估值回落，切勿盲目追高"

            eff_yield = cfg['target_yield'] * macro_multiplier if cfg['asset_type'] == '個股' else cfg['target_yield']
            target_str = f"{calc_target:.1f}" if calc_target > 0 else "採用固定目標價"
            div_val_str = f"{cfg['fwd_div']:.2f}元 (外資共識)" if cfg['asset_type'] == '個股' else f"{cfg['fwd_div']:.2f}元 (TTM配息)"
            
            preset_matrix_rows += (
                f"<tr style='border-bottom: 1px solid #e2e8f0;'>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:center; background-color:#f7fafc;'><b>{c_name}</b><br><small style='color:#718096;'>{s_id}</small></td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:center;'><span style='background-color:#e2e8f0; padding:2px 5px; border-radius:3px; font-size:11px;'>{cfg['asset_type']}</span></td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:right;'>{div_val_str}</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:center; font-weight:bold; color:#718096;'>{eff_yield:.2f}%</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:right; font-weight:bold; color:#b7791f; background-color:#fffaf0;'>{target_str}</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; font-size:12px;'>{path_desc}</td>"  
                f"<td style='padding:10px; border:1px solid #e2e8f0; font-size:12px;'>{defense_desc}</td>"
                f"</tr>"
            )
        
        dividend_table_html = (
            f"<div style='margin-bottom: 30px; background-color: #fff; padding: 18px; border: 2px solid #2b6cb0; border-radius: 8px;'>"
            f"<h3 style='color: #2b6cb0; margin-top: 0; margin-bottom: 14px;'>📋 全動態量化智慧估值決策面板 (數據源於 Excel 參數與 YFinance 串接)</h3>"
            f"<table style='width:100%; border-collapse:collapse; font-size:13px; border: 1px solid #e2e8f0;'>"
            f"<thead><tr style='background-color: #2b6cb0; color: white;'>"
            f"<th style='padding:10px;'>標的名稱</th><th style='padding:10px;'>資產類別</th><th style='padding:10px;'>自動推估股利基準</th><th style='padding:10px;'>動態要求殖利率</th><th style='padding:10px;'>演算法目標買點</th><th style='padding:10px;'>採取行動邏輯路徑 (動態)</th><th style='padding:10px;'>大盤聯動與風控機制</th>"
            f"</tr></thead><tbody>{preset_matrix_rows}</tbody></table></div>"
        )

        alert_html = "<div style='border-left: 4px solid #f6ad55; padding-left: 15px; margin-bottom: 25px;'><h3 style='color: #dd6b20;'>⚠️ 智慧多因子策略買賣觸發提示 (LINE 僅放行買進通知)</h3><ul style='line-height: 1.8;'>"
        for a in alerts:
            alert_html += f"<li><b>{a['icon']} [{a['type']}] {a['name']}</b> (今日收盤: {a['close']}) - {a['line2']}<br>說明：{a['desc']}</li>"
        alert_html += "</ul></div>" if alerts else "</div>"

        table_rows = ""
        for s in all_stocks:
            type_bg = "#e2e8f0" if s['type'] == '監控觀察股' else "#feebc8"
            status_style = "color:red; font-weight:bold;" if "🎯" in s['status'] or "⚠️" in s['status'] else "color:#38a169;"
            
            k_threshold = getattr(config, 'K_THRESHOLD', 15)
            is_nan_k = pd.isna(s['k'])
            k_style = "color:red; font-weight:bold;" if not is_nan_k and s['k'] < k_threshold else ""
            
            k_str = f"{s['k']:.2f}" if not is_nan_k else "50.00"
            d_str = f"{s['d']:.2f}" if not pd.isna(s['d']) else "50.00"
            osc_str = f"{s['osc']:.2f}" if not pd.isna(s['osc']) else "0.00"
            
            table_rows += (
                f"<tr>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><b>{s['name']}</b><br><small>{s['id']}</small></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><span style='background-color:{type_bg}; padding:3px 8px; border-radius:4px;'>{s['type']}</span></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'><b>{s['close']:.2f}</b></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right; {k_style}'>{k_str}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{d_str}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{osc_str}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'>{s['target_or_cost']}</td>" 
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center; {status_style}'>{s['status']}</td>"
                f"</tr>"
            )
            
        html = f"<html><body style=\"font-family: 'Microsoft JhengHei'; padding: 20px;\"><h2>📊 台股雙多智慧策略決策總覽</h2>{logic_explanation_html}{dividend_table_html}<p><b>數據基準日：</b>{report_date}</p><p>🌐 {market_text} (當前總經平滑修正係數: {macro_multiplier})</p>{alert_html}<hr><table style=\"width:100%; border-collapse:collapse;\"><thead><tr style=\"background-color: #004a99; color: white;\"><th>標的名稱</th><th>名單分類</th><th>今日收盤</th><th>日K</th><th>日D</th><th>MACD (OSC)</th><th>目標 / 狀態</th><th>策略診斷狀態</th></tr></thead><tbody>{table_rows}</tbody></table></body></html>"
        msg.attach(MIMEText(html, 'html'))
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_user, self.email_password)
                for t in (self.recipients + self.bcc_recipients): server.sendmail(self.email_user, t, msg.as_string())
        except Exception as e: logger.error(f"Mail 寄送失敗: {e}")


class StrategyOrchestrator:
    def __init__(self):
        self.repo = StockDataRepository(BASE_DIR)
        self.notifier = NotificationService()

    def execute_pipeline(self) -> None:
        is_bull_market, market_text, macro_multiplier = MarketIndicatorService.get_macro_status()
        
        structured_alerts, all_stocks_output, global_stock_pool = [], [], {}
        triggered_exit_stocks = set()
        
        inventory_data = {}
        monitor_data = set()
        monitor_configs = {}
        
        # 1. 讀取庫存清單 (精準讀取您自訂之 成本、波段最高、停利百分比 欄位)
        if INVENTORY_FILE.exists():
            try:
                inv_df = pd.read_excel(INVENTORY_FILE)
                if '股票名稱' in inv_df.columns and '成本' in inv_df.columns:
                    for name, group in inv_df.groupby('股票名稱'):
                        c_name = clean_stock_name(name)
                        if not c_name: continue
                        
                        avg_cost = group['成本'].mean()
                        file_highest_p = group['波段最高'].max() if '波段最高' in inv_df.columns else None
                        tp_pct = group['停利百分比'].iloc[0] / 100.0 if '停利百分比' in inv_df.columns else 0.15
                        
                        inventory_data[c_name] = {
                            'avg_cost': avg_cost,
                            'highest_p': file_highest_p,
                            'tp_pct': tp_pct
                        }
            except Exception as e: logger.error(f"庫存股票.xlsx 讀取失敗: {e}")

        # 2. 讀取監控清單與自動建立市場均值估值 (新增資產類別、要求殖利率(%)、目標配息率(%)、固定買進目標價)
        if MONITOR_FILE.exists():
            try:
                mon_df = pd.read_excel(MONITOR_FILE)
                m_name_col = '股票名稱' if '股票名稱' in mon_df.columns else mon_df.columns[0]
                
                for _, row in mon_df.iterrows():
                    c_name = clean_stock_name(row[m_name_col])
                    if not c_name: continue
                    monitor_data.add(c_name)
                    
                    s_id = str(row.get('代號', '')).strip()
                    asset_type = str(row.get('資產類別', '個股')).strip()
                    target_yield = float(row.get('要求殖利率(%)', 2.0))
                    payout_ratio = float(row.get('目標配息率(%)', 50.0)) / 100.0
                    fixed_target = float(row.get('固定買進目標價', 0.0))
                    
                    dynamic_target = 0.0
                    fwd_div = 0.0
                    
                    # 🟢 分流估值演算法 (高股息 ETF 與 個股 雙軌制)
                    try:
                        ticker = yf.Ticker(s_id + ".TW")
                        
                        if asset_type == '高股息ETF':
                            # ETF: 自動回溯加總過去一年 True TTM 配息，不綁定大盤修正係數
                            hist_div = ticker.dividends
                            if not hist_div.empty:
                                fwd_div = hist_div.last("1Y").sum()
                                dynamic_target = fwd_div / (target_yield / 100)
                        else:
                            # 個股: 自動聯網爬取外資/法人的 forwardEps 市場共識均值
                            fwd_eps = ticker.info.get('forwardEps', 0.0)
                            if fwd_eps and fwd_eps > 0:
                                fwd_div = fwd_eps * payout_ratio
                                # 結合平滑化總經乘數算大波段便宜價
                                dynamic_target = fwd_div / ((target_yield * macro_multiplier) / 100)
                    except Exception as e:
                        logger.warning(f"未能獲取 {c_name} 線上共識數據，系統防禦級 fallback 啟用: {e}")
                    
                    # 🛡️ 雙重安全鎖：若爬不到 EPS 或算出低於10元的異常低價，自動Fallback回您設定的固定目標價
                    if dynamic_target <= 10.0 and fixed_target > 0:
                        dynamic_target = fixed_target
                        fwd_div = 0.0
                        
                    monitor_configs[c_name] = {
                        'asset_type': asset_type,
                        'target_yield': target_yield,
                        'fwd_div': fwd_div,
                        'dynamic_target': round_to_tw_tick(dynamic_target)
                    }
            except Exception as e: logger.error(f"監控股票.xlsx 讀取失敗: {e}")

        all_unique_stocks = set(inventory_data.keys()).union(monitor_data)

        # 3. 核心互斥狀態機邏輯處理 (杜絕雙胞胎重複印出與不合理移動停利)
        for c_name in all_unique_stocks:
            try:
                hist_df = self.repo.get_history("", c_name)
                if hist_df.empty: continue
                
                latest = hist_df.iloc[-1]
                current_price = latest['Close']
                
                is_owned = c_name in inventory_data
                is_monitored = c_name in monitor_data
                
                avg_cost = inventory_data[c_name]['avg_cost'] if is_owned else 0
                tp_pct = inventory_data[c_name]['tp_pct'] if is_owned else 0.15
                highest_p = inventory_data[c_name]['highest_p'] if is_owned and pd.notna(inventory_data[c_name]['highest_p']) else hist_df['Close'].max()
                
                target_price = monitor_configs.get(c_name, {}).get('dynamic_target', 0.0) if is_monitored else 0.0

                status_str = ""
                
                # 狀態 1：鐵律停損 (有持股 + 虧損 + 大盤走空)
                if is_owned and current_price <= avg_cost and not is_bull_market:
                    status_str = "🛑 鐵律停損"
                    structured_alerts.append({
                        'icon': '🛑', 'type': '鐵律清倉停損', 'name': c_name, 'close': f"{current_price:.2f}", 
                        'line2': "狀態: 跌破防守線", 'desc': "大盤走空，請嚴守資金紀律全數清倉避險！", 'is_monitor_alert': False
                    })
                    triggered_exit_stocks.add(c_name)
                    
                # 狀態 2：移動停利 (有持股 + 實質获利鎖[現價>成本] + 跌破您設定的停利防禦線)
                elif is_owned and current_price > avg_cost and current_price <= highest_p * (1 - tp_pct):
                    status_str = "⚠️ 移動停利"
                    structured_alerts.append({
                        'icon': '⚠️', 'type': '庫存移動停利', 'name': c_name, 'close': f"{current_price:.2f}", 
                        'line2': f"波段最高: {highest_p:.2f}", 'desc': f"觸發移動停利線 ({round_to_tw_tick(highest_p*(1-tp_pct)):.2f})，建議獲利落袋。", 'is_monitor_alert': False
                    })
                    triggered_exit_stocks.add(c_name)
                    
                # 狀態 3：買進/加碼 (跌破全動態估值便宜買點)
                elif is_monitored and target_price > 0 and current_price <= target_price:
                    if is_owned:
                        status_str = "🔄 建議加碼"
                        structured_alerts.append({
                            'icon': '🔄', 'type': '庫存逢低加碼提示', 'name': c_name, 'close': f"{current_price:.2f}", 
                            'line2': f"加碼目標: {target_price:.2f}", 'desc': "波段趨勢安全，已達估值加倉區間，建議分批加碼。", 'is_monitor_alert': True
                        })
                    else:
                        status_str = "🎯 已達買點"
                        structured_alerts.append({
                            'icon': '🎯', 'type': '監控買進提示', 'name': c_name, 'close': f"{current_price:.2f}", 
                            'line2': f"買進目標: {target_price:.2f}", 'desc': "已達大波段安全邊際，建議分批佈局。", 'is_monitor_alert': True
                        })
                        
                # 狀態 4：智慧緩衝 (有持股 + 虧損 + 大盤多頭 + 未達加碼點)
                elif is_owned and current_price <= avg_cost and is_bull_market:
                    status_str = "🔄 智慧緩衝"
                    
                # 狀態 5：預設常態
                else:
                    if is_owned:
                        status_str = "✅ 持股安全"
                    else:
                        diff_pct = ((current_price - target_price) / target_price) * 100 if target_price > 0 else 0
                        status_str = f"溢價 {diff_pct:.1f}%" if target_price > 0 else "常態監控中"
                        continue

                df_idx = MarketIndicatorService.calculate_kd9(hist_df)
                _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                
                all_stocks_output.append({
                    'name': c_name, 'id': latest['id'], 
                    'type': '現有庫存股' if is_owned else '監控觀察股', 
                    'close': current_price,
                    'k': df_idx.iloc[-1]['K'], 'd': df_idx.iloc[-1]['D'], 'osc': osc_s.iloc[-1],
                    'target_or_cost': f"目標: {target_price:.2f}" if is_monitored and target_price > 0 else "庫存持有中", 
                    'status': status_str
                })
                
                if is_owned:
                    global_stock_pool[c_name] = {'close': current_price}
                    
            except Exception as e: logger.error(f"處理個股 {c_name} 執行失敗: {e}")

        # 4. 派送綜合報告與單向一對一好友廣播
        if all_stocks_output:
            self.notifier.send_html_email(self.repo.report_date, market_text, structured_alerts, all_stocks_output, global_stock_pool, triggered_exit_stocks, macro_multiplier, monitor_configs)
            
            # LINE 廣播僅放行帶有 'is_monitor_alert': True 的落地加碼/建倉提示
            only_monitor_alerts = [a for a in structured_alerts if a.get('is_monitor_alert', False)]
            if only_monitor_alerts:
                self.notifier.send_line(only_monitor_alerts, self.repo.report_date)


if __name__ == "__main__":
    orchestrator = StrategyOrchestrator()
    orchestrator.execute_pipeline()
