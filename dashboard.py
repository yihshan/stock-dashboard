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

st.set_page_config(page_title="台股投資分析報告", layout="wide")

# --- 1. 新聞分析模組 ---
@st.cache_data(ttl=3600)
def fetch_portfolio_news(api_key, stock_names):
    if not api_key or not stock_names: return None
    search_query = " OR ".join([f'"{name}"' for name in stock_names[:5]])
    try:
        url = f"https://news.google.com/rss/search?q={urllib.parse.quote(search_query + ' 股市')}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10, verify=False)
        items = ET.fromstring(response.text).findall('./channel/item')[:4]
        news_list = [f"- {item.find('title').text}" for item in items]
        
        api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
        headers = {'Content-Type': 'application/json', 'x-goog-api-key': api_key}
        prompt = f"分析以下新聞:\n{chr(10).join(news_list)}\n請輸出 JSON：[{{\"title\": \"標題\", \"content\": \"摘要\"}}]"
        resp = requests.post(api_url, json={"contents": [{"parts": [{"text": prompt}]}]}, headers=headers, timeout=20, verify=False)
        return json.loads(resp.json()['candidates'][0]['content']['parts'][0]['text'].replace("```json", "").replace("```", "").strip())
    except: return None

# --- 2. 工具與數據 ---
def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (datetime, date)): return val.date()
    s = str(val).strip().split(' ')[0]
    digits = re.sub(r'\D', '', s)
    return datetime.strptime(digits, '%Y%m%d').date() if len(digits) == 8 else None

@st.cache_data(ttl=60)
def load_data():
    inv = pd.read_excel(Path(config.DATA_DIR) / "庫存股票.xlsx")
    inv.columns = inv.columns.str.strip()
    inv['交易日期'] = inv['交易日期'].apply(parse_date)
    
    csv_files = glob.glob(str(Path(config.DATA_DIR) / "*.csv"))
    all_close = []
    for f in csv_files:
        df = pd.read_csv(f, encoding='cp950', on_bad_lines='skip')
        date_col = next((c for c in df.columns if '日期' in c), None)
        price_col = next((c for c in df.columns if '收盤' in c), None)
        name_col = next((c for c in df.columns if '名稱' in c or '代號' in c), None)
        if date_col and price_col and name_col:
            df = df.rename(columns={name_col: '股票名稱', price_col: '收盤價', date_col: '日期'})
            df['日期_dt'] = df['日期'].apply(parse_date)
            all_close.append(df[['股票名稱', '收盤價', '日期_dt']])
    return inv, pd.concat(all_close) if all_close else None

inventory_df, close_df = load_data()
monitor_configs = pd.read_excel(Path(config.DATA_DIR) / "監控股票.xlsx").set_index('股票名稱').to_dict('index') if os.path.exists(Path(config.DATA_DIR) / "監控股票.xlsx") else {}

# --- 3. UI 主架構 ---
if inventory_df is not None and close_df is not None:
    # 新聞與 Slider
    portfolio_stocks = inventory_df['股票名稱'].unique().tolist()
    news = fetch_portfolio_news(getattr(config, 'GEMINI_API_KEY', None), portfolio_stocks)
    
    all_dates = sorted(close_df['日期_dt'].dropna().unique())
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
                cost = (group['成本'] * group['股數']).sum() / shares
                val = price * shares
                pnl = val - (cost * shares) - max(20, cost * shares * 0.001425) - max(20, val * 0.001425) - (val * 0.003)
                daily_details.append({"日期": d, "股票名稱": name, "淨損益": pnl, "市值": val})
    
    full_df = pd.DataFrame(daily_details)
    latest_summary = full_df[full_df['日期'] == selected_date]

    # --- 新功能：資產監控 ---
    st.markdown("---")
    st.markdown("### 🛡️ 資產與現金流監控儀表板")
    c1, c2 = st.columns(2)
    with c1:
        total_val = latest_summary['市值'].sum()
        max_row = latest_summary.nlargest(1, '市值')
        conc = (max_row['市值'].iloc[0] / total_val) * 100
        st.metric("最大單一個股佔比", f"{conc:.1f}%", max_row['股票名稱'].iloc[0])
        st.progress(min(conc/30.0, 1.0))
    with c2:
        # 計算預估年配息
        est_annual_div = 0
        for name, group in inventory_df.groupby('股票名稱'):
            if name in monitor_configs and '預估每股配息' in monitor_configs[name]:
                est_annual_div += (monitor_configs[name]['預估每股配息'] * group['股數'].sum())
        target = getattr(config, 'TARGET_Q_2026', 280000) * 4
        st.metric("2026 預估年配息", f"${est_annual_div:,.0f}")
        st.progress(min(est_annual_div/target, 1.0))

    # --- KPI 與 圖表 ---
    st.markdown("### 🔑 關鍵績效指標")
    m1, m2, m3, m4 = st.columns(4)
    total_pnl = latest_summary['淨損益'].sum()
    m1.metric("總市值", f"${total_val:,.0f}")
    m2.metric("累積損益", f"${total_pnl:,.0f}")
    
    # 圖表
    cols = st.columns(3)
    cols[0].plotly_chart(px.area(full_df, x="日期", y="淨損益", color="股票名稱"), use_container_width=True)
    cols[1].plotly_chart(px.pie(latest_summary, values='市值', names='股票名稱'), use_container_width=True)
    cols[2].plotly_chart(px.bar(latest_summary, x='股票名稱', y='淨損益'), use_container_width=True)

    # 明細表
    st.markdown("### 📋 投資組合明細")
    prev_date = all_dates[all_dates.index(selected_date) - 1] if all_dates.index(selected_date) > 0 else selected_date
    prev_data = full_df[full_df['日期'] == prev_date]
    
    rows = []
    for name in latest_summary['股票名稱']:
        pnl = latest_summary[latest_summary['股票名稱']==name]['淨損益'].iloc[0]
        prev = prev_data[prev_data['股票名稱']==name]['淨損益'].sum()
        rows.append({"名稱": name, "市值": latest_summary[latest_summary['股票名稱']==name]['市值'].iloc[0], "累計損益": pnl, "本日獲利": pnl - prev})
    st.dataframe(pd.DataFrame(rows))
