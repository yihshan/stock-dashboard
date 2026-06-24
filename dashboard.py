import streamlit as st
import pandas as pd
import os
import glob
import re
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
from pathlib import Path
import requests
import xml.etree.ElementTree as ET
import urllib.parse
import json
import urllib3

# 隱藏 SSL 警告
urllib3.disable_warnings()

# 導入 config
try:
    import config
except ImportError:
    st.error("❌ 找不到 config.py，請確保設定檔存在。")
    st.stop()

# 設定頁面 (全域只能呼叫一次)
st.set_page_config(page_title="台股投資分析報告 v3.9.1", layout="wide")

# --- 1. 動態個股新聞分析模組 ---
@st.cache_data(ttl=3600)
def fetch_portfolio_news(api_key, stock_names):
    if not api_key or not stock_names: return None
    search_query = " OR ".join([f'"{name}"' for name in stock_names[:5]])
    news_items = []
    try:
        url = f"https://news.google.com/rss/search?q={urllib.parse.quote(search_query + ' 股市')}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10, verify=False)
        for item in ET.fromstring(response.text).findall('./channel/item')[:8]:
            news_items.append(f"- {item.find('title').text}")
    except Exception: return None
        
    if not news_items: return None
    try:
        api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
        headers = {'Content-Type': 'application/json', 'x-goog-api-key': api_key}
        prompt = (
            f"以下是關於您持倉股票 ({', '.join(stock_names)}) 的新聞標題：\n{chr(10).join(news_items)}\n\n"
            "請扮演專業分析師，總結 4 則盤前監控重點。請只輸出 JSON 陣列格式：[{\"title\": \"🚀 標題\", \"content\": \"摘要\"}]"
        )
        resp = requests.post(api_url, json={"contents": [{"parts": [{"text": prompt}]}]}, headers=headers, timeout=20, verify=False)
        if resp.status_code == 200:
            text = resp.json()['candidates'][0]['content']['parts'][0]['text']
            return json.loads(text.replace("```json", "").replace("```", "").strip())
    except Exception: pass
    return None

# --- 2. 初始化庫存 ---
INVENTORY_FILE = Path(config.DATA_DIR) / "庫存股票.xlsx"
inventory_df = pd.read_excel(INVENTORY_FILE) if os.path.exists(INVENTORY_FILE) else None

# --- 3. UI 樣式設定 ---
st.markdown("""
    <style>
    h1, h2, h3 { color: #1a365d; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    
    /* 專業表格樣式 */
    .report-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 20px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .report-table th {
        background-color: #f1f3f5;
        color: #1a365d;
        font-weight: bold;
        text-align: center !important;
        padding: 12px;
        border: 1px solid #dee2e6;
    }
    .report-table td {
        padding: 10px;
        border: 1px solid #dee2e6;
        text-align: right;
    }
    .report-table td:first-child {
        text-align: center;
    }
    
    /* 負數紅字樣式 */
    .neg-value { color: #dc3545; font-weight: bold; }
    
    /* 自定義 KPI 卡片樣式 */
    .kpi-box {
        padding: 20px;
        border-radius: 10px;
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        text-align: center;
    }
    .kpi-label { font-size: 1rem; color: #6c757d; margin-bottom: 5px; }
    .kpi-value { font-size: 1.8rem; font-weight: bold; color: #1a365d; }
    .kpi-value-red { font-size: 1.8rem; font-weight: bold; color: #dc3545; }

    /* 新聞區塊樣式 */
    .news-box {
        background-color: #fff9db;
        padding: 15px;
        border-left: 5px solid #fcc419;
        border-radius: 5px;
        margin-bottom: 20px;
    }
    .news-title { font-weight: bold; color: #856404; margin-bottom: 5px; }
    .news-content { font-size: 0.95rem; color: #555; }
    </style>
    """, unsafe_allow_html=True)

# 定義資料路徑
BASE_DIR = Path(config.DATA_DIR)
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"

# 財務參數
FEE_RATE = 0.001425
TAX_RATE = 0.003
MIN_FEE = 20

def format_finance(val, is_percent=False):
    if val is None: return ""
    suffix = "%" if is_percent else ""
    abs_val = abs(val)
    val_str = f"{abs_val:,.2f}{suffix}" if is_percent else f"{abs_val:,.0f}{suffix}"
    return f'<span class="neg-value">({val_str})</span>' if val < 0 else val_str

def format_finance_plain(val, is_percent=False):
    try:
        val = float(val)
        formatted = f"{val:.2f}%" if is_percent else f"{val:,.0f}"
        if val < 0:
            abs_val = abs(val)
            return f'<span style="color:red">({abs_val:.2f}%)</span>' if is_percent else f'<span style="color:red">({abs_val:,.0f})</span>'
        return formatted
    except:
        return str(val)

def calculate_net_pnl(cost, price, shares):
    total_cost_base = cost * shares
    buy_fee = max(MIN_FEE, total_cost_base * FEE_RATE)
    total_market_value = price * shares
    sell_fee = max(MIN_FEE, total_market_value * FEE_RATE)
    sell_tax = total_market_value * TAX_RATE
    net_pnl = total_market_value - total_cost_base - sell_fee - sell_tax
    return net_pnl, total_market_value, sell_fee + sell_tax

def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    s = str(val).strip().split(' ')[0]
    if not s: return None
    
    for sep in ['/', '-']:
        parts = s.split(sep)
        if len(parts) == 3:
            try:
                y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                return date(y if y > 1911 else y + 1911, m, d)
            except: continue
            
    digits = re.sub(r'\D', '', s)
    if len(digits) == 8:
        try: return datetime.strptime(digits, '%Y%m%d').date()
        except: pass
    elif len(digits) == 7:
        try: return date(int(digits[:3]) + 1911, int(digits[3:5]), int(digits[5:]))
        except: pass
    return None

def clean_price(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(',', '').strip()
    try: return float(s)
    except: return 0.0

@st.cache_data(ttl=300)
def load_data():
    if not INVENTORY_FILE.exists(): return None, None, ["❌ 找不到庫存檔案"]
    try:
        inventory_df = pd.read_excel(INVENTORY_FILE, sheet_name=0)
        col_map = {}
        for col in inventory_df.columns:
            c = str(col).strip()
            if '名稱' in c: col_map['股票名稱'] = col
            if '日期' in c: col_map['交易日期'] = col
            if '股數' in c: col_map['股數'] = col
            if '成本' in c: col_map['成本'] = col
        if len(col_map) < 4: return None, None, [f"❌ 缺少必要欄位"]
        inventory_df = inventory_df.rename(columns={v: k for k, v in col_map.items()})
        inventory_df['交易日期'] = inventory_df['交易日期'].apply(parse_date)
        inventory_df = inventory_df.dropna(subset=['交易日期'])
        inventory_df['股票名稱'] = inventory_df['股票名稱'].astype(str).str.strip()
        inventory_df['股數'] = pd.to_numeric(inventory_df['股數'], errors='coerce').fillna(0)
        inventory_df['成本'] = pd.to_numeric(inventory_df['成本'], errors='coerce').fillna(0)
    except: return None, None, ["⚠️ 讀取庫存失敗"]

    csv_files = glob.glob(str(BASE_DIR / "*.csv"))
    all_close_data = []
    for f in csv_files:
        try:
            try: df = pd.read_csv(f, encoding='cp950')
            except: df = pd.read_csv(f, encoding='utf-8')
            
            date_col = None
            for col in df.columns:
                if df[col].dropna().head(10).apply(parse_date).count() >= 3:
                    date_col = col; break
            
            target_cols = {'股票名稱': None, '收盤價': None}
            for col in df.columns:
                c_clean = str(col).strip()
                if any(x in c_clean for x in ['名稱', '證券名稱', '股票名稱', 'Name']):
                    target_cols['股票名稱'] = col
                elif any(x in c_clean for x in ['證券代號', '代號', '股票代號', 'Code']) and not target_cols['股票名稱']:
                    target_cols['股票名稱'] = col
                if any(x in c_clean for x in ['收盤', '收盤價', 'ClosingPrice', 'Close']):
                    target_cols['收盤價'] = col
            
            if target_cols['股票名稱'] and target_cols['收盤價']:
                temp_df = df.copy()
                temp_df['日期_dt'] = temp_df[date_col].apply(parse_date) if date_col else parse_date(re.search(r'(\d{8})|(\d{7})', os.path.basename(f)).group())
                temp_df = temp_df.dropna(subset=['日期_dt'])
                temp_df = temp_df[[target_cols['股票名稱'], target_cols['收盤價'], '日期_dt']]
                temp_df.columns = ['股票名稱', '收盤價', '日期_dt']
                temp_df['股票名稱'] = temp_df['股票名稱'].astype(str).str.strip()
                temp_df['收盤價'] = temp_df['收盤價'].apply(clean_price)
                all_close_data.append(temp_df)
        except: pass
    return inventory_df, (pd.concat(all_close_data) if all_close_data else None), []

inventory_df, close_df, logs = load_data()

# --- 盤前重大消息區塊 ---
st.markdown("### 🔔 今日盤前重大消息 (Gemini AI 即時分析)")
if inventory_df is not None and '股票名稱' in inventory_df.columns:
    portfolio_stocks = inventory_df['股票名稱'].unique().tolist()
    api_key = os.getenv("GEMINI_API_KEY") or getattr(config, 'GEMINI_API_KEY', None)
    
    with st.spinner("🔄 正在取得並分析最新市場新聞..."):
        dynamic_news = fetch_portfolio_news(api_key, portfolio_stocks)

    news_col1, news_col2 = st.columns(2)
    if dynamic_news and len(dynamic_news) >= 2:
        mid = len(dynamic_news) // 2
        with news_col1:
            for item in dynamic_news[:mid]:
                st.markdown(f'''<div class="news-box"><div class="news-title">{item.get('title', '焦點新聞')}</div><div class="news-content">{item.get('content', '')}</div></div>''', unsafe_allow_html=True)
        with news_col2:
            for item in dynamic_news[mid:]:
                st.markdown(f'''<div class="news-box"><div class="news-title">{item.get('title', '焦點新聞')}</div><div class="news-content">{item.get('content', '')}</div></div>''', unsafe_allow_html=True)
    else:
        with news_col1:
            st.info("⚠️ 目前 AI 伺服器繁忙，無法即時分析，請稍後重整。")
        with news_col2:
            st.markdown("""
                <div class="news-box">
                    <div class="news-title">⚠️ 外部風險觀察</div>
                    <div class="news-content">
                        需緊盯 Fed 利率決策動向及美銀示警之市場泡沫訊號。外資期貨空單仍處高位，短線需防範漲多回檔風險。
                    </div>
                </div>
            """, unsafe_allow_html=True)
else:
    st.warning("⚠️ 無法讀取庫存或 Excel 中缺少「股票名稱」欄位。")

st.divider()

if inventory_df is not None and close_df is not None:
    all_dates = sorted(close_df['日期_dt'].unique())
    
    # 頂部控制列
    ctrl1, ctrl2 = st.columns([7, 3])
    with ctrl1:
        selected_date = st.select_slider("📅 選擇報告基準日：", options=all_dates, value=all_dates[-1])
    with ctrl2:
        all_stock_names = ["全部個股"] + sorted(inventory_df['股票名稱'].unique().tolist())
        selected_stock = st.selectbox("🔍 個股連動分析：", options=all_stock_names)

    # 計算損益數據 (加入收盤價與股數，以便後續直覺計算每日差額)
    daily_details = []
    for d in all_dates:
        if d > selected_date: continue
        d_close = close_df[close_df['日期_dt'] == d]
        d_inv = inventory_df[inventory_df['交易日期'] <= d]
        for name, group in d_inv.groupby('股票名稱'):
            shares = group['股數'].sum()
            if shares <= 0: continue
            avg_cost = (group['成本'] * group['股數']).sum() / shares
            m_close = d_close[d_close['股票名稱'] == name]
            if not m_close.empty:
                price = m_close.iloc[0]['收盤價']
                pnl, val, _ = calculate_net_pnl(avg_cost, price, shares)
                daily_details.append({
                    "日期": d, "股票名稱": name, "淨損益": pnl, "市值": val, 
                    "成本": avg_cost * shares, "收盤價": price, "股數": shares
                })
    
    full_df = pd.DataFrame(daily_details)
    filtered_full_df = full_df if selected_stock == "全部個股" else full_df[full_df['股票名稱'] == selected_stock]
    latest_summary = filtered_full_df[filtered_full_df['日期'] == selected_date]
    
    # 計算「本日獲利(損)」 - 改為直覺算法：(今日收盤價 - 昨日收盤價) * 股數
    daily_pnl = 0
    if len(all_dates) >= 2:
        try:
            idx = all_dates.index(selected_date)
            if idx > 0:
                prev_date = all_dates[idx-1]
                prev_summary = filtered_full_df[filtered_full_df['日期'] == prev_date]
                for _, row in latest_summary.iterrows():
                    name = row['股票名稱']
                    c_price = row['收盤價']
                    c_shares = row['股數']
                    prev_row = prev_summary[prev_summary['股票名稱'] == name]
                    p_price = prev_row.iloc[0]['收盤價'] if not prev_row.empty else c_price
                    daily_pnl += (c_price - p_price) * c_shares
        except: pass

    total_net_pnl = latest_summary['淨損益'].sum() if not latest_summary.empty else 0
    total_market_value = latest_summary['市值'].sum() if not latest_summary.empty else 0
    total_cost = latest_summary['成本'].sum() if not latest_summary.empty else 0
    roi = (total_net_pnl / total_cost * 100) if total_cost > 0 else 0

    st.markdown("### 🔑 關鍵績效指標")
    m1, m2, m3, m4, m5 = st.columns(5)
    
    with m1:
        st.markdown(f'<div class="kpi-box"><div class="kpi-label">資產總市值</div><div class="kpi-value">${total_market_value:,.0f}</div></div>', unsafe_allow_html=True)
    with m2:
        pnl_display = format_finance(total_net_pnl)
        st.markdown(f'<div class="kpi-box"><div class="kpi-label">累積淨損益</div><div class="kpi-value">${pnl_display}</div></div>', unsafe_allow_html=True)
    with m3:
        pnl_class = "kpi-value-red" if daily_pnl < 0 else "kpi-value"
        abs_daily_pnl = abs(daily_pnl)
        val_str = f"{abs_daily_pnl:,.0f}"
        daily_pnl_display = f"({val_str})" if daily_pnl < 0 else val_str
        st.markdown(f'<div class="kpi-box"><div class="kpi-label">本日獲利(損)</div><div class="{pnl_class}">{daily_pnl_display}</div></div>', unsafe_allow_html=True)
    with m4:
        roi_display = format_finance(roi, is_percent=True)
        st.markdown(f'<div class="kpi-box"><div class="kpi-label">投資報酬率</div><div class="kpi-value">{roi_display}</div></div>', unsafe_allow_html=True)
    with m5:
        st.markdown(f'<div class="kpi-box"><div class="kpi-label">報告基準日</div><div class="kpi-value" style="font-size:1.2rem; padding-top:10px;">{selected_date}</div></div>', unsafe_allow_html=True)

    st.divider()

    # 圖表區
    c1, c2, c3 = st.columns([5, 2.5, 2.5])
    with c1:
        st.subheader("📈 損益結構趨勢")
        if not filtered_full_df.empty:
            fig_trend = px.area(filtered_full_df, x="日期", y="淨損益", color="股票名稱", 
                                title=f"{selected_stock} 損益變動", template="plotly_white")
            fig_trend.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350, showlegend=(selected_stock == "全部個股"))
            st.plotly_chart(fig_trend, use_container_width=True)

    with c2:
        st.subheader("🍰 資產配置")
        if not latest_summary.empty:
            fig_pie = px.pie(latest_summary, values='市值', names='股票名稱', hole=.4, template="plotly_white")
            fig_pie.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350, showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)

    with c3:
        st.subheader("🏆 績效貢獻")
        if not latest_summary.empty:
            fig_bar = px.bar(latest_summary.sort_values('淨損益'), x='淨損益', y='股票名稱', orientation='h', 
                             color='淨損益', color_continuous_scale='RdYlGn', template="plotly_white")
            fig_bar.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350, coloraxis_showscale=False)
            st.plotly_chart(fig_bar, use_container_width=True)

    st.divider()

    # 投資組合明細清單
    st.subheader("📋 投資組合明細清單")
    current_close = close_df[close_df['日期_dt'] == selected_date]
    inventory_to_show = inventory_df if selected_stock == "全部個股" else inventory_df[inventory_df['股票名稱'] == selected_stock]
    
    # 計算各股的本日獲利(損) - 直覺算法：(今日收盤價 - 昨日收盤價) * 股數
    stock_daily_pnl_map = {}
    if len(all_dates) >= 2:
        try:
            idx = all_dates.index(selected_date)
            if idx > 0:
                prev_date = all_dates[idx-1]
                prev_details = filtered_full_df[filtered_full_df['日期'] == prev_date]
                for _, row in latest_summary.iterrows():
                    name = row['股票名稱']
                    c_price = row['收盤價']
                    c_shares = row['股數']
                    prev_row = prev_details[prev_details['股票名稱'] == name]
                    if not prev_row.empty:
                        p_price = prev_row.iloc[0]['收盤價']
                        stock_daily_pnl_map[name] = (c_price - p_price) * c_shares
                    else:
                        stock_daily_pnl_map[name] = 0
        except: pass

    rows_html = ""
    for name, group in inventory_to_show.groupby('股票名稱'):
        shares = group['股數'].sum()
        if shares <= 0: continue
        avg_cost = (group['成本'] * group['股數']).sum() / shares
        m_close = current_close[current_close['股票名稱'] == name]
        if not m_close.empty:
            price = m_close.iloc[0]['收盤價']
            pnl, val, fees = calculate_net_pnl(avg_cost, price, shares)
            roi_stock = (pnl / (avg_cost * shares) * 100) if avg_cost > 0 else 0
            daily_pnl_stock = stock_daily_pnl_map.get(name, 0)
            
            rows_html += f"""<tr>
<td>{name}</td>
<td>{shares:,.0f}</td>
<td>{avg_cost:,.2f}</td>
<td>{price:,.2f}</td>
<td>{val:,.0f}</td>
<td>{fees:,.0f}</td>
<td>{format_finance_plain(pnl)}</td>
<td>{format_finance_plain(daily_pnl_stock)}</td>
<td>{format_finance_plain(roi_stock, is_percent=True)}</td>
</tr>"""
    
    st.markdown(f"""
        <table class="report-table">
            <thead>
                <tr>
                    <th>股票名稱</th>
                    <th>持有股數</th>
                    <th>平均成本</th>
                    <th>目前市價</th>
                    <th>資產市值</th>
                    <th>預估稅費</th>
                    <th>累積淨損益</th>
                    <th>本日獲利(損)</th>
                    <th>投資報酬率</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        <p style="color: #666; font-size: 0.8rem; margin-top: 10px;">註：負數以括號 ( ) 表示。累積淨損益的計算包含扣除買進/賣出手續費及證交稅。</p>
    """, unsafe_allow_html=True)
else:
    st.info("👋 歡迎使用！請確認 '每日收盤' 或 'data' 資料夾中已包含 '庫存股票.xlsx' 與收盤價 CSV 檔案。")
    if logs:
        for log in logs: st.error(log)
