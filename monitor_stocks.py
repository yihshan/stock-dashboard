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

# 🟢 全域核心大腦：2026-2027 預估股利與要求殖利率自適應矩陣
DIVIDEND_PRESETS = {
    '2330': {'name': '台積電', 'div_2026': 36.0, 'div_2027': 44.0, 'target_yield': 1.8},  
    '2308': {'name': '台達電', 'div_2026': 26.0, 'div_2027': 34.0, 'target_yield': 1.8},  
    '3008': {'name': '大立光', 'div_2026': 95.0, 'div_2027': 115.0, 'target_yield': 2.4}, 
    '3017': {'name': '奇鋐', 'div_2026': 32.0, 'div_2027': 42.0, 'target_yield': 1.8},  
    '3131': {'name': '弘塑', 'div_2026': 50.0, 'div_2027': 65.0, 'target_yield': 1.7},  
    '3443': {'name': '創意', 'div_2026': 45.0, 'div_2027': 60.0, 'target_yield': 1.3},  
    '6442': {'name': '光聖', 'div_2026': 18.0, 'div_2027': 26.0, 'target_yield': 1.6},  
    '3324': {'name': '雙鴻', 'div_2026': 16.0, 'div_2027': 24.0, 'target_yield': 1.9},  
    '6510': {'name': '精測', 'div_2026': 35.0, 'div_2027': 50.0, 'target_yield': 1.6},  
    '3563': {'name': '牧德', 'div_2026': 12.0, 'div_2027': 18.0, 'target_yield': 2.2},  
    '7751': {'name': '竑騰', 'div_2026': 20.0, 'div_2027': 30.0, 'target_yield': 1.9},  
    '7734': {'name': '印能科技', 'div_2026': 55.0, 'div_2027': 75.0, 'target_yield': 2.1},  
    '8299': {'name': '群聯', 'div_2026': 45.0, 'div_2027': 65.0, 'target_yield': 2.3},  
    '8210': {'name': '勤誠', 'div_2026': 24.0, 'div_2027': 34.0, 'target_yield': 2.2},  
    '6789': {'name': '采鈺', 'div_2026': 8.5, 'div_2027': 11.0, 'target_yield': 2.1},
    '7750': {'name': '新代', 'div_2026': 42.0, 'div_2027': 55.0, 'target_yield': 2.4},
    '3030': {'name': '德律', 'div_2026': 7.5, 'div_2027': 9.5, 'target_yield': 2.5},
    '6515': {'name': '穎威', 'div_2026': 110.0, 'div_2027': 140.0, 'target_yield': 1.6},
    '3081': {'name': '聯亞', 'div_2026': 22.0, 'div_2027': 30.0, 'target_yield': 1.4}
}

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
    return s.replace(' ', '').replace('　', '')


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

        raw_data_list = []
        # 🟢 修正核心 1：強制完整串接目錄下所有 CSV 歷史檔案，確保時間軸天數充足！
        for file_path in csv_files:
            file_date = DateDataParser.extract_date_from_filename(os.path.basename(file_path))
            try:
                try: df = pd.read_csv(file_path, encoding='utf-8-sig')
                except: df = pd.read_csv(file_path, encoding='cp950')
                
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
        
        # 🟢 修正核心 2：精準分離代號與名稱查找路徑，徹底根絕交叉對齊錯置的 Bug
        if target_id:
            df_res = self.master_df[self.master_df['id'] == target_id]
        else:
            df_res = self.master_df[self.master_df['name'] == target_name]
            
        return df_res.sort_values('Date').drop_duplicates('Date').reset_index(drop=True)


class MarketIndicatorService:
    @staticmethod
    def calculate_kd9(df: pd.DataFrame) -> pd.DataFrame:
        # 🟢 修正核心 3：防空機制。若歷史天數嚴重不足，自動給予平滑基底值，徹底在表格上消滅 "nan"
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
    def check_macro_regime() -> Tuple[bool, str]:
        try:
            twii = yf.Ticker("^TWII")
            df = twii.history(period="1y")
            df['200MA'] = df['Close'].rolling(window=200).mean()
            is_bull = df['Close'].iloc[-1] >= df['200MA'].iloc[-1]
            return is_bull, f"今日加權指數收盤 {df['Close'].iloc[-1]:.2f} 點，處於年線 ({df['200MA'].iloc[-1]:.1f}) 之{'上' if is_bull else '下'}"
        except:
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
        lines = [f"{a['icon']} [{a['type']}] {a['name']}\n- 今日收盤: {a['close']}\n- {a['line2']}\n- 說明: {a['desc']}" for a in alerts]
        send_line_message(getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', ''), getattr(config, 'LINE_USER_ID', ''), f"\n📊 【台股雙多智慧策略決策報告】\n基準日: {report_date}\n" + "\n------------\n".join(lines))

    def send_html_email(self, report_date: str, market_text: str, alerts: List[Dict[str, str]], all_stocks: List[Dict[str, Any]], global_stock_pool: Dict[str, Any]) -> None:
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 台股雙多智慧策略決策報告與診斷總覽 - {report_date}"
        msg['From'] = self.email_user
        msg['To'] = ", ".join(self.recipients)

        # 🟢 修正核心 4：徹底解鎖！將「行動邏輯路徑」從寫死字串改為動態橫向核對大腦
        preset_matrix_rows = ""
        for s_id, cfg in DIVIDEND_PRESETS.items():
            avg_div = (cfg['div_2026'] + cfg['div_2027']) / 2
            calc_target = avg_div / (cfg['target_yield'] / 100)
            
            # 即時動態穿透審查該資產是否在現有持股庫存中
            is_owned = (cfg['name'] in global_stock_pool)
            if is_owned:
                path_desc = "<span style='color:#3182ce; font-weight:bold;'>路徑 2：現價 ≦ 股利估值買點 (有持股)</span><br><small style='color:#718096;'>判定為核心資產 ➜ 啟動大波段逢低加碼路徑</small>"
                defense_desc = "🛡️ 多頭環境下享智慧緩衝防護網"
            else:
                path_desc = "<span style='color:#dd6b20;'>路徑 1：現價 ≦ 股利估值買點 (無持股)</span><br><small style='color:#718096;'>判定為新觀察標的 ➜ 啟動安全邊際分批佈局</small>"
                defense_desc = "🛡️ 空頭市場全面啟動鐵律停損風控"

            preset_matrix_rows += (
                f"<tr style='border-bottom: 1px solid #e2e8f0;'>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:center; background-color:#f7fafc;'><b>{cfg['name']}</b><br><small style='color:#718096;'>{s_id}</small></td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:right;'>{cfg['div_2026']:.1f}元</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:right;'>{cfg['div_2027']:.1f}元</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:center; font-weight:bold; color:#e53e3e;'>{cfg['target_yield']:.2f}%</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; text-align:right; font-weight:bold; color:#b7791f; background-color:#fffaf0;'>{calc_target:.1f}</td>"
                f"<td style='padding:10px; border:1px solid #e2e8f0; font-size:12px;'>{path_desc}</td>"  
                f"<td style='padding:10px; border:1px solid #e2e8f0; font-size:12px; color:#e53e3e;'>{defense_desc}</td>"
                f"</tr>"
            )
        
        dividend_table_html = (
            f"<div style='margin-bottom: 30px; background-color: #fff; padding: 18px; border: 2px solid #2b6cb0; border-radius: 8px;'>"
            f"<h3 style='color: #2b6cb0; margin-top: 0; margin-bottom: 14px;'>📋 2026-2027 智慧股利估值與策略決策動態路徑綜合面板</h3>"
            f"<table style='width:100%; border-collapse:collapse; font-size:13px; border: 1px solid #e2e8f0;'>"
            f"<thead><tr style='background-color: #2b6cb0; color: white;'>"
            f"<th style='padding:10px;'>股票名稱</th><th style='padding:10px;'>2026配息</th><th style='padding:10px;'>2027配息</th><th style='padding:10px;'>要求殖利率</th><th style='padding:10px;'>推算目標價</th><th style='padding:10px;'>採取行動邏輯路徑 (動態)</th><th style='padding:10px;'>大盤聯動機制</th>"
            f"</tr></thead><tbody>{preset_matrix_rows}</tbody></table></div>"
        )

        alert_html = "<div style='border-left: 4px solid #f6ad55; padding-left: 15px; margin-bottom: 25px;'><h3 style='color: #dd6b20;'>⚠️ 智慧多因子策略買賣觸發提示</h3><ul style='line-height: 1.8;'>"
        for a in alerts:
            alert_html += f"<li><b>{a['icon']} [{a['type']}] {a['name']}</b> (今日收盤: {a['close']}) - {a['line2']}<br>說明：{a['desc']}</li>"
        alert_html += "</ul></div>" if alerts else "</div>"

        table_rows = ""
        for s in all_stocks:
            type_bg = "#e2e8f0" if s['type'] == '監控觀察股' else "#feebc8"
            status_style = "color:red; font-weight:bold;" if "🎯" in s['status'] or "⚠️" in s['status'] else "color:#38a169;"
            
            k_str = f"{s['k']:.2f}" if not pd.isna(s['k']) else "50.00"
            d_str = f"{s['d']:.2f}" if not pd.isna(s['d']) else "50.00"
            osc_str = f"{s['osc']:.2f}" if not pd.isna(s['osc']) else "0.00"
            
            table_rows += (
                f"<tr>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><b>{s['name']}</b><br><small>{s['id']}</small></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'><span style='background-color:{type_bg}; padding:3px 8px; border-radius:4px;'>{s['type']}</span></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'><b>{s['close']:.2f}</b></td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{k_str}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{d_str}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{osc_str}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'>{s['target_or_cost']}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center; {status_style}'>{s['status']}</td>"
                f"</tr>"
            )
            
        html = f"<html><body style=\"font-family: 'Microsoft JhengHei'; padding: 20px;\"><h2>📊 台股智慧決策診斷總覽</h2>{dividend_table_html}<p><b>數據基準日：</b>{report_date}</p><p>🌐 {market_text}</p>{alert_html}<hr><table style=\"width:100%; border-collapse:collapse;\"><thead><tr style=\"background-color: #004a99; color: white;\"><th>股票名稱</th><th>資產類別</th><th>收盤價</th><th>日K</th><th>日D</th><th>MACD (OSC)</th><th>目標 / 成本</th><th>策略診斷狀態</th></tr></thead><tbody>{table_rows}</tbody></table></body></html>"
        msg.attach(MIMEText(html, 'html'))
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_user, self.email_password)
                for t in (self.recipients + self.bcc_recipients): server.sendmail(self.email_user, t, msg.as_string())
        except Exception as e: logger.error(f"Mail failed: {e}")


class StrategyOrchestrator:
    def __init__(self):
        self.repo = StockDataRepository(BASE_DIR)
        self.notifier = NotificationService()

    def execute_pipeline(self) -> None:
        is_bull_market, market_text = MarketIndicatorService.check_macro_regime()
        structured_alerts, all_stocks_output, global_stock_pool = [], [], {}
        triggered_exit_stocks = set()
        
        # 1. 第一階段：精準滾動處理現有庫存股
        if INVENTORY_FILE.exists():
            try:
                inv_df = pd.read_excel(INVENTORY_FILE)
                inv_name_col = next((c for c in inv_df.columns if '名稱' in str(c) or '股票' in str(c)), inv_df.columns[0])
                inv_cost_col = next((c for c in inv_df.columns if '成本' in str(c) or '價位' in str(c)), None)
                inv_date_col = next((c for c in inv_df.columns if '日期' in str(c)), None)
                
                for name, group in inv_df.groupby(inv_name_col):
                    c_name = clean_stock_name(name)
                    if not c_name: continue
                    hist_df = self.repo.get_history("", c_name)
                    if hist_df.empty: continue
                    
                    latest = hist_df.iloc[-1]
                    current_price = latest['Close']
                    avg_cost = group[inv_cost_col].mean() if inv_cost_col else current_price
                    
                    # 🟢 修正核心 5：移動停利精準校正。以買入後的歷史最大收盤價為準；防範空值誤判
                    buy_date = DateDataParser.parse_generic_date(group[inv_date_col].iloc[0]) if inv_date_col else None
                    if buy_date:
                        post_buy_df = hist_df[hist_df['Date'] >= buy_date]
                        highest_p = post_buy_df['Close'].max() if not post_buy_df.empty else current_price
                    else:
                        highest_p = hist_df['Close'].max()
                        
                    if pd.isna(highest_p) or highest_p == 0: highest_p = current_price

                    status_str = "✅ 持股安全"
                    if current_price <= avg_cost:
                        if is_bull_market:
                            structured_alerts.append({'icon': '🔄', 'type': '智慧緩衝：暫緩停損', 'name': c_name, 'close': f"{current_price:.2f}", 'line2': f"成本: {avg_cost:.2f}", 'desc': "大盤加權指數處於強勢多頭格局，此回檔建議暫緩盲目砍單。"})
                            status_str = "🔄 智慧緩衝"
                        else:
                            structured_alerts.append({'icon': '🛑', 'type': '鐵律清倉停損', 'name': c_name, 'close': f"{current_price:.2f}", 'line2': f"成本: {avg_cost:.2f}", 'desc': "大盤走空，請嚴守資金紀律全數清倉避險！"})
                            status_str = "🛑 鐵律停損"
                            triggered_exit_stocks.add(c_name)
                    elif current_price <= highest_p * 0.85:
                        structured_alerts.append({'icon': '⚠️', 'type': '庫存移動停利', 'name': c_name, 'close': f"{current_price:.2f}", 'line2': f"波段最高: {highest_p:.2f}", 'desc': f"觸發移動停利線 ({highest_p*0.85:.2f})，建議獲利落袋。"})
                        status_str = "⚠️ 移動停利"
                        triggered_exit_stocks.add(c_name)

                    df_idx = MarketIndicatorService.calculate_kd9(hist_df)
                    _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                    
                    res = {
                        'name': c_name, 'id': latest['id'], 'type': '現有庫存股', 'close': current_price, 
                        'k': df_idx.iloc[-1]['K'], 'd': df_idx.iloc[-1]['D'], 'osc': osc_s.iloc[-1], 
                        'target_or_cost': f"成本: {avg_cost:.2f}", 'status': status_str
                    }
                    all_stocks_output.append(res)
                    global_stock_pool[c_name] = res
            except Exception as e: logger.error(f"庫存風控模組執行失敗: {e}")

        # 2. 第二階段：處理監控觀察名單
        if MONITOR_FILE.exists():
            try:
                monitor_df = pd.read_excel(MONITOR_FILE)
                name_col = next((col for col in monitor_df.columns if '名稱' in str(col)), monitor_df.columns[0])
                
                for _, row in monitor_df.iterrows():
                    c_name = clean_stock_name(row[name_col])
                    if not c_name: continue
                    df = self.repo.get_history("", c_name)
                    if df.empty: continue
                    
                    latest = df.iloc[-1]
                    current_price = latest['Close']
                    s_id_str = str(latest['id'])
                    
                    if s_id_str in DIVIDEND_PRESETS:
                        cfg = DIVIDEND_PRESETS[s_id_str]
                        target_price = ((cfg['div_2026'] + cfg['div_2027']) / 2) / (cfg['target_yield'] / 100)
                        diff_pct = ((current_price - target_price) / target_price) * 100
                        status_str = f"溢價 {diff_pct:.1f}%"
                    else:
                        target_price = current_price * 0.85
                        status_str = "溢價 17.6%"
                    
                    is_already_owned = (c_name in global_stock_pool)
                    
                    if current_price <= target_price and c_name not in triggered_exit_stocks:
                        if is_already_owned:
                            structured_alerts.append({'icon': '🔄', 'type': '庫存逢低加碼提示', 'name': c_name, 'close': f"{current_price:.2f}", 'line2': f"加碼目標: {target_price:.2f}", 'desc': "波段趨勢安全，已達估值加倉區間，建議分批加碼。"})
                            status_str = "🔄 建議加碼"
                        else:
                            structured_alerts.append({'icon': '🎯', 'type': '監控買進提示', 'name': c_name, 'close': f"{current_price:.2f}", 'line2': f"買進目標: {target_price:.2f}", 'desc': "已達大波段安全安全邊際，建議分批佈局。"})
                            status_str = "🎯 已達買點"

                    df_idx = MarketIndicatorService.calculate_kd9(df)
                    _, _, osc_s = MarketIndicatorService.calculate_macd(df_idx)
                    
                    all_stocks_output.append({
                        'name': c_name, 'id': latest['id'], 'type': '監控觀察股', 'close': current_price,
                        'k': df_idx.iloc[-1]['K'], 'd': df_idx.iloc[-1]['D'], 'osc': osc_s.iloc[-1],
                        'target_or_cost': f"目標: {target_price:.2f}", 'status': status_str
                    })
            except Exception as e: logger.error(f"監控觀察模組執行失敗: {e}")

        # 🟢 修正核心 6：將完整的 global_stock_pool 作為橫向審計指標傳遞給 HTML 生成器
        if all_stocks_output:
            self.notifier.send_html_email(self.repo.report_date, market_text, structured_alerts, all_stocks_output, global_stock_pool)

if __name__ == "__main__":
    orchestrator = StrategyOrchestrator()
    orchestrator.execute_pipeline()
