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

# 配置生產線等級的日誌系統
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(filename)s:%(lineno)d - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('strategy_robot.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

# 強制 UTF-8 環境
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ==========================================
# 🛑 核心配置與全域常數定義 (最高優先權)
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
    logger.warning("找不到 line_messaging 模組，切換為發送模擬模式。")
    def send_line_message(token: str, user_id: str, msg: str) -> None:
        logger.info(f"[LINE 模擬推播] {msg}")

# 嚴格在最頂層定義全域路徑，確保全檔案 Scope 皆可無縫存取
BASE_DIR = Path(config.DATA_DIR)
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"

# 隱藏 SSL 警告
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class DateDataParser:
    """負責處理台股特有的日期與字串格式轉換（工具類別）"""
    
    @staticmethod
    def parse_generic_date(val: Any) -> Optional[date]:
        if pd.isna(val):
            return None
        s = str(val).strip().split(' ')[0]
        if not s:
            return None
        
        for sep in ['/', '-']:
            parts = s.split(sep)
            if len(parts) == 3:
                try:
                    year = int(parts[0])
                    corrected_year = year if year > 1911 else year + 1911
                    return date(corrected_year, int(parts[1]), int(parts[2]))
                except ValueError:
                    continue
                    
        digits = re.sub(r'\D', '', s)
        if len(digits) == 8:
            try:
                return datetime.strptime(digits, '%Y%m%d').date()
            except ValueError:
                pass
        elif len(digits) == 7:
            try:
                return date(int(digits[:3]) + 1911, int(digits[3:5]), int(digits[5:]))
            except ValueError:
                pass
        return None

    @staticmethod
    def extract_date_from_filename(filename: str) -> Optional[date]:
        match = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', filename)
        if match:
            return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        match = re.search(r'(\d{8})|(\d{7})', filename)
        if match:
            d = match.group()
            return datetime.strptime(d, '%Y%m%d').date() if len(d) == 8 else date(int(d[:3]) + 1911, int(d[3:5]), int(d[5:]))
        return None


class StockDataRepository:
    """資料訪問層 (Data Access Layer)：優化 I/O，一次性全載入記憶體快取"""
    
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.report_date = str(date.today())
        self._historical_cache: Dict[str, pd.DataFrame] = {}
        self._initialize_database()

    def _initialize_database(self) -> None:
        """核心最佳化：單次掃描硬碟，完成所有歷史收盤價 CSV 的快取建立"""
        csv_files = glob.glob(str(self.data_dir / "台股每日收盤價_*.csv"))
        if not csv_files:
            logger.warning(f"資料夾 {self.data_dir} 中未偵測到任何台股收盤價 CSV 檔案。")
            return
            
        latest_file = max(csv_files, key=os.path.basename)
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(latest_file))
        if date_match:
            self.report_date = date_match.group(1)

        logger.info(f"💾 開始建立記憶體快取庫，追蹤歷史 CSV 檔案共 {len(csv_files)} 個...")
        
        raw_data_list = []
        for file_path in csv_files:
            file_date = DateDataParser.extract_date_from_filename(os.path.basename(file_path))
            if not file_date:
                continue
                
            try:
                try:
                    df = pd.read_csv(file_path, encoding='utf-8-sig')
                except Exception:
                    df = pd.read_csv(file_path, encoding='cp950')
                
                col_map = {}
                for col in df.columns:
                    c = str(col).strip()
                    if any(x in c for x in ['日期', 'Date']): col_map['date'] = col
                    if any(x in c for x in ['證券代號', 'Code', '股票代號']): col_map['id'] = col
                    if any(x in c for x in ['名稱', 'Name']): col_map['name'] = col
                    if any(x in c for x in ['收盤', 'Close', 'ClosingPrice']): col_map['close'] = col
                    if any(x in c for x in ['最高', 'High', 'HighestPrice']): col_map['high'] = col
                    if any(x in c for x in ['最低', 'Low', 'LowestPrice']): col_map['low'] = col
                
                if 'close' not in col_map:
                    continue
                    
                for _, row in df.iterrows():
                    s_id = str(row[col_map['id']]).replace('"', '').strip() if 'id' in col_map else ""
                    s_name = str(row[col_map['name']]).strip() if 'name' in col_map else ""
                    
                    c_val = float(str(row[col_map['close']]).replace(',', '').strip()) if pd.notna(row[col_map['close']]) else np.nan
                    h_val = float(str(row[col_map['high']]).replace(',', '').strip()) if 'high' in col_map and pd.notna(row[col_map['high']]) else c_val
                    l_val = float(str(row[col_map['low']]).replace(',', '').strip()) if 'low' in col_map and pd.notna(row[col_map['low']]) else c_val
                    
                    if pd.isna(c_val):
                        continue
                        
                    raw_data_list.append({
                        'id': s_id, 'name': s_name, 'Date': file_date,
                        'Close': c_val, 'High': h_val, 'Low': l_val
                    })
            except Exception as e:
                logger.error(f"解析 CSV 檔案失敗 {os.path.basename(file_path)}: {e}")
                
        if raw_data_list:
            master_df = pd.DataFrame(raw_data_list)
            if 'id' in master_df.columns:
                for s_id, g in master_df.groupby('id'):
                    if s_id: self._historical_cache[str(s_id)] = g.sort_values('Date').drop_duplicates('Date')
            for s_name, g in master_df.groupby('name'):
                if s_name: self._historical_cache[str(s_name)] = g.sort_values('Date').drop_duplicates('Date')
                
        logger.info(f"✅ 快取建立完成，已成功將 {len(self._historical_cache)} 檔個股數據移入記憶體。")

    def get_history(self, stock_id: str, stock_name: str) -> pd.DataFrame:
        """記憶體等級的高速查詢，取代舊版的硬碟迴圈讀取"""
        s_id = str(stock_id).strip()
        s_name = re.sub(r'\(.*?\)', '', stock_name).strip()
        
        if s_id in self._historical_cache:
            return self._historical_cache[s_id]
        if s_name in self._historical_cache:
            return self._historical_cache[s_name]
        return pd.DataFrame()


class MarketIndicatorService:
    """技術指標運算服務層 (Business Logic Layer)"""
    
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
        """抓取大盤年線(200MA)總體環境分析"""
        try:
            logger.info("🌐 正在下載大盤加權指數 (計算 200MA 年線系統風險)...")
            twii = yf.Ticker("^TWII")
            df = twii.history(period="1y")
            if df.empty:
                raise ValueError("無法從 Yahoo Finance 獲取加權指數數據。")
            df['200MA'] = df['Close'].rolling(window=200).mean()
            latest_close = df['Close'].iloc[-1]
            ma200 = df['200MA'].iloc[-1]
            
            if pd.isna(ma200):
                return True, "⚠️ 大盤歷史天數不夠計算年線，預設切換為【友善多頭環境】"
                
            is_bull = latest_close >= ma200
            status_text = "【多頭市場】(年線之上，啟動智慧估值緩衝)" if is_bull else "【空頭熊市】(全面防守，啟動鐵律停損模式)"
            return is_bull, f"今日加權指數收盤 {latest_close:.1f} 位於年線 ({ma200:.1f}) 之{'上' if is_bull else '下'} -> {status_text}"
        except Exception as e:
            logger.error(f"大盤總體環境解析異常: {e}，策略自動降級為【安全多頭環境】防呆。")
            return True, "⚠️ 總體環境伺服器連線異常，策略降級切換為【安全多頭環境】"


class NotificationService:
    """通知派送服務層 (Infrastructure Layer)"""
    
    def __init__(self):
        self.smtp_server = config.SMTP_SERVER
        self.smtp_port = config.SMTP_PORT
        self.email_user = config.EMAIL_USER
        self.email_password = config.EMAIL_PASSWORD
        self.recipients = config.RECIPIENTS
        self.bcc_recipients = getattr(config, 'BCC_RECIPIENTS', [])

    def send_line(self, alerts: List[str], report_date: str) -> None:
        if not alerts:
            return
        token = getattr(config, 'LINE_CHANNEL_ACCESS_TOKEN', '')
        user_id = getattr(config, 'LINE_USER_ID', '')
        full_msg = f"\n📊 【台股雙多策略決策報告】\n基準日: {report_date}\n" + "\n---------------------\n".join(alerts)
        send_line_message(token, user_id, full_msg)
        logger.info("🔔 策略決策訊號已成功推播至 LINE 頻道。")

    def send_html_email(self, report_date: str, market_text: str, alerts: List[str], all_stocks: List[Dict[str, Any]]) -> None:
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 台股雙多智慧策略決策報告 - {report_date}"
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
            dif_val = f"{s['dif']:.2f}" if not pd.isna(s['dif']) else "N/A"
            macd_val = f"{s['macd']:.2f}" if not pd.isna(s['macd']) else "N/A"
            osc_val = f"{s['osc']:.2f}" if not pd.isna(s['osc']) else "N/A"

            table_rows += (
                f"<tr>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:center;'>{s['name']} ({s['id']})</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{s['close']:.2f}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right; {k_style}'>{k_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{d_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{dif_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{macd_val}</td>"
                f"<td style='padding:10px; border:1px solid #ddd; text-align:right;'>{osc_val}</td>"
                f"</tr>"
            )
            
        html = (
            f"<html><body style=\"font-family: 'Microsoft JhengHei', sans-serif; padding: 20px;\">"
            f"<h2 style=\"color: #1a365d;\">📊 每日台股策略監控與技術指標自動彙整</h2>"
            f"<p style='color: #4a5568;'><b>數據基準日：</b>{report_date}</p>"
            f"<p style='background-color: #edf2f7; padding: 10px; border-radius: 4px; color: #4a5568;'>🌐 <b>總體環境監測：</b>{market_text}</p>"
            f"{alert_html}"  
            f"<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 30px 0;'>"
            f"<h3 style=\"color: #004a99;\">📈 關注個股最新日線指標總覽</h3>"
            f"<table style=\"width:100%; border-collapse:collapse; margin-top:15px;\">"
            f"<thead><tr style=\"background-color: #004a99; color: white;\">"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">股票名稱</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">收盤價</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">日K值</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">日D值</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">DIF</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">MACD</th>"
            f"<th style=\"padding:12px; border:1px solid #ddd;\">OSC(柱狀體)</th>"
            f"</tr></thead><tbody>{table_rows}</tbody></table>"
            f"<p style='font-size: 12px; color: #a0aec0; margin-top: 20px;'>💡 本郵件由 Agentic AI 智慧型多因子風控系統自動發送</p>"
            f"</body></html>"
        )
        msg.attach(MIMEText(html, 'html'))
        
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_user, self.email_password)
                targets = self.recipients + self.bcc_recipients
                for target in targets:
                    server.sendmail(self.email_user, target, msg.as_string())
            logger.info("📨 [Email 系統] 整合型智慧決策報告已順利寄達指定信箱。")
        except Exception as e:
            logger.error(f"發送 Email 失敗: {e}", exc_info=True)


class StrategyOrchestrator:
    """架構核心：策略指揮官，負責調配各組服務元件執行完整的監控流程"""
    
    def __init__(self):
        # 這裡會精確存取在最頂端就已經定義好、絕不踩空的全域變數 BASE_DIR
        self.repo = StockDataRepository(BASE_DIR)
        self.notifier = NotificationService()

    def execute_pipeline(self) -> None:
        is_bull_market, market_text = MarketIndicatorService.check_macro_regime()
        strategy_alerts: List[str] = []
        all_stocks_output: Dict[str, Dict[str, Any]] = {}

        # ==========================================
        # 1. 掃描與執行觀察監控名單
        # ==========================================
        if MONITOR_FILE.exists():
            try:
                monitor_df = pd.read_excel(MONITOR_FILE)
                name_col = next((col for col in monitor_df.columns if '名稱' in str(col)), monitor_df.columns[0])
                id_col = next((col for col in monitor_df.columns if '代號' in str(col)), None)
                
                for _, row in monitor_df.iterrows():
                    name = str(row[name_col]).strip()
                    s_id = str(row[id_col]).strip() if id_col and pd.notna(row[id_col]) else ""
                    if not s_id:
                        match = re.search(r'\((\d+)\)', name)
                        s_id = match.group(1) if match else ""
                    if not name or name == 'nan': 
                        continue
                        
                    df = self.repo.get_history(s_id, name)
                    if df.empty:
                        continue
                        
                    latest = df.iloc[-1]
                    current_price = latest['Close']
                    
                    if '買進目標價' in monitor_df.columns and pd.notna(row['買進目標價']):
                        target_price = float(row['買進目標價'])
                        pe_limit = row['買進本益比上限'] if '買進本益比上限' in monitor_df.columns else None
                        if current_price <= target_price:
                            pe_msg = f" (本益比門檻: {pe_limit}x)" if pd.notna(pe_limit) else ""
                            strategy_alerts.append(f"🎯 [監控買進提示] {name}\n  - 今日收盤: {current_price}\n  - 買進目標價: {target_price}{pe_msg}\n  - 股價已修正至大波段安全邊際買點，建議分批佈局！")

                    if len(df) < 10:
                        all_stocks_output[name] = {
                            'name': name, 'id': s_id, 'close': current_price, 
                            'k': np.nan, 'd': np.nan, 'dif': np.nan, 'macd': np.nan, 'osc': np.nan
                        }
                        continue
                        
                    df = MarketIndicatorService.calculate_kd9(df)
                    dif, macd, osc = MarketIndicatorService.calculate_macd(df)
                    df['60MA'] = df['Close'].rolling(window=min(60, len(df))).mean()
                    
                    latest = df.iloc[-1]
                    prev_row = df.iloc[-2] if len(df) > 1 else latest
                    current_k, current_osc, current_ma60 = latest['K'], osc.iloc[-1], latest['60MA']
                    
                    all_stocks_output[name] = {
                        'name': name, 'id': s_id, 'close': current_price, 
                        'k': current_k, 'd': latest['D'], 'dif': dif.iloc[-1], 'macd': macd.iloc[-1], 'osc': current_osc
                    }
                    
                    k_threshold = getattr(config, 'K_THRESHOLD', 15)
                    if current_price >= current_ma60:
                        is_right_side_confirmed = (prev_row['K'] <= 10 and current_k > 15) or (current_osc < 0 and current_osc > osc.iloc[-2] if len(osc)>1 else False)
                        if current_k < k_threshold and is_right_side_confirmed:
                            strategy_alerts.append(f"🚨 [技術右側買點觸發] {name} ({s_id})\n  - 今日收盤: {current_price}\n  - 狀態: 股價在季線之上，且滿足低檔黃金止跌動能！")
            except Exception as e:
                logger.error(f"處理解析監控股票 Excel 失敗: {e}", exc_info=True)

        # ==========================================
        # 2. 庫存股票大波段智慧風控防線
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
                    
                    base_stop = avg_cost 
                    
                    highest_record = group['波段最高價'].dropna().head(1).values
                    highest_price = float(highest_record[0]) if len(highest_record) > 0 and not pd.isna(highest_record[0]) else 0.0
                    
                    current_price = all_stocks_output[name]['close'] if name in all_stocks_output else np.nan
                    hist_df = self.repo.get_history("", name)
                    
                    if pd.isna(current_price) and not hist_df.empty:
                        current_price = hist_df.iloc[-1]['Close']
                    if highest_price == 0.0 or pd.isna(highest_price):
                        highest_price = hist_df['Close'].max() if not hist_df.empty else current_price

                    if current_price > 0:
                        if current_price > highest_price:
                            highest_price = current_price
                            logger.info(f"🚀 {name} 突破歷史新高，天花板上移至: {highest_price}")
                        
                        sell_trigger_price = highest_price * (1 - (trail_pct / 100))
                        
                        if current_price <= base_stop:
                            if is_bull_market:
                                strategy_alerts.append(
                                    f"🔄 [智慧緩衝：暫緩停損] {name}\n  - 今日收盤: {current_price}\n  - 加權平均成本: {avg_cost:.1f}\n"
                                    f"  - 總體風控：大盤處於牛市。此回檔定義為多頭非理性震盪，**系統強烈建議暫緩盲目砍單**，切勿殺在阿呆谷。"
                                )
                            else:
                                strategy_alerts.append(
                                    f"🛑 [鐵律清倉停損觸發] {name}\n  - 今日收盤: {current_price}\n  - 加權平均成本: {avg_cost:.1f}\n"
                                    f"  - 總體風控：大盤已確認走空轉熊！**請嚴守資金紀律，全數清倉停損離場**，保留實力！"
                                )
                        elif current_price <= sell_trigger_price:
                            strategy_alerts.append(
                                f"⚠️ [庫存移動停利提示] {name}\n  - 今日收盤: {current_price}\n  - 波段最高點: {highest_price}\n"
                                f"  - 已自高點回撤超過 {trail_pct}%\n  - 決策建議：已觸發移動利潤線 ({sell_trigger_price:.1f})，建議高檔分批落袋為安。"
                            )
                    
                    for _, row in group.iterrows():
                        row['移動停利百分比(%)'] = trail_pct
                        row['波段最高價'] = highest_price
                        row['保本停損價'] = base_stop
                        updated_rows.append(row)
                        
                pd.DataFrame(updated_rows).to_excel(INVENTORY_FILE, index=False)
                logger.info("💾 庫存 Excel 歷史新高數據已安全同步回寫。")
            except Exception as e:
                logger.error(f"庫存大波段智慧風控核心運算失敗: {e}", exc_info=True)

        # ==========================================
        # 3. 雙通道分流調度通知
        # ==========================================
        self.notifier.send_line(strategy_alerts, self.repo.report_date)
        if all_stocks_output:
            self.notifier.send_html_email(
                self.repo.report_date, market_text, strategy_alerts, list(all_stocks_output.values())
            )


if __name__ == "__main__":
    logger.info("🎬 自動化智慧多因子投資策略監控系統啟動...")
    orchestrator = StrategyOrchestrator()
    orchestrator.execute_pipeline()
    logger.info("🏁 系統執行完畢，安全關閉。")
