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

urllib3.disable_warnings()

try:
    import config
except ImportError:
    st.error("❌ 找不到 config.py，請確保設定檔存在。")
    st.stop()

st.set_page_config(page_title="台股投資分析報告 v3.9.1", layout="wide")

# --- 工具函式 ---
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
    except: return None
    if not news_items: return None
    try:
        api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
        headers = {'Content-Type': 'application/json', 'x-goog-api-key': api_key}
        prompt = f"以下是持倉股票 ({', '.join(stock_names)}) 新聞：\n{chr(10).join(news_items)}\n請輸出 JSON：[{{\"title\": \"標題\", \"content\": \"摘要\"}}]"
        resp = requests.post(api_url, json={"contents": [{"parts": [{"text": prompt}]}]}, headers=headers, timeout=20, verify=False)
        return json.loads(resp.json()['candidates'][0]['content']['parts'][0]['text'].replace("```json", "").replace("```", "").strip())
    except: return None

def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (datetime, date)): return val.date() if isinstance(val, datetime) else val
    s = str(val).strip().split(' ')[0]
    digits = re.sub(r'\D', '', s)
    if len(digits) == 8: return datetime.strptime(digits, '%Y%m%d').date()
    return None

def calculate_net_pnl(cost, price, shares):
    val = price * shares
    fee = max(20, cost * shares * 0.001425) + max(20, val * 0.001425)
    tax = val * 0.003
    return val - (cost * shares) - fee - tax, val, fee + tax

@st.cache_data(ttl=60)
def load_data():
    if not os.path.exists(Path(config.DATA_DIR) / "庫存股票.xlsx"): return None, None
    inv = pd.read_excel(Path(config.DATA_DIR) / "庫存股票.xlsx")
    inv.columns = inv.columns.str.strip()
    inv['交易日期'] = inv['交易日期'].apply(parse_date)
    
    csv_files = glob.glob(str(Path(config.DATA_DIR) / "*.csv"))
    all_close = []
    for f in csv_files:
        df = pd.read_csv(f, encoding='cp950', on_bad_lines='skip')
        # 尋找日期與收盤價欄位
        date_col = next((c for c in df.columns if '日期' in c), None)
        price_col = next((c for c in df.columns if '收盤' in c), None)
        name_col = next((c for c in df.columns if '名稱' in c or '代號' in c), None)
        if date_col and price_col and name_col:
            df = df.rename(columns={name_col: '股票名稱', price_col: '收盤價', date_col: '日期'})
            df['日期_dt'] = df['日期'].apply(parse_date)
            all_close.append(df[['股票名稱', '收盤價', '日期_dt']])
    return inv, pd.concat(all_close) if all_close else None

# --- 載入 ---
inventory_df, close_df = load_data()
monitor_configs = pd.read_excel(Path(config.DATA_DIR) / "監控股票.xlsx").set_index('股票名稱').to_dict('index') if os.path.exists(Path(config.DATA_DIR) / "監控股票.xlsx") else {}

# --- UI 渲染 ---
st.markdown("### 🔔 市場與持倉監控")
if inventory_df is not None:
    # 新聞區
    news = fetch_portfolio_news(getattr(config, 'GEMINI_API_KEY', None), inventory_df['股票名稱'].unique().tolist())
    if news:
        cols = st.columns(2)
        for i, item in enumerate(news):
            with cols[i % 2]: st.warning(f"**{item['title']}**\n\n{item['content']}")

    all_dates = sorted(close_df['日期_dt'].unique())
    selected_date = st.select_slider("📅 選擇報告基準日：", options=all_dates, value=all_dates[-1])
    
    # 計算損益
    daily_details = []
    for d in all_dates:
        if d > selected_date: continue
        d_close = close_df[close_df['日期_dt'] == d]
        for name, group in inventory_df[inventory_df['交易日期'] <= d].groupby('股票名稱'):
            shares = group['股數'].sum()
            price = d_close[d_close['股票名稱'] == name]['收盤價'].iloc[0] if name in d_close['股票名稱'].values else 0
            if price > 0:
                pnl, val, _ = calculate_net_pnl((group['成本'] * group['股數']).sum() / shares, price, shares)
                daily_details.append({"日期": d, "股票名稱": name, "淨損益": pnl, "市值": val})
    
    full_df = pd.DataFrame(daily_details)
    latest_summary = full_df[full_df['日期'] == selected_date]

    # --- 儀表板 ---
    st.markdown("### 🛡️ 資產與現金流監控儀表板")
    c1, c2 = st.columns(2)
    with c1:
        total_val = latest_summary['市值'].sum()
        max_stock = latest_summary.nlargest(1, '市值')
        conc = (max_stock['市值'].iloc[0] / total_val) * 100
        st.metric("最大單一個股佔比", f"{conc:.1f}%", max_stock['股票名稱'].iloc[0])
        st.progress(min(conc/30.0, 1.0))
    with c2:
        st.metric("預估年配息", "請確認 Excel 配息欄位")
        st.progress(0.4)

    # --- KPI與明細 ---
    st.markdown("### 🔑 關鍵績效指標")
    st.dataframe(latest_summary)
