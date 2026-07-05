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

# 設定頁面
st.set_page_config(page_title="台股投資分析報告 v3.9.1", layout="wide")

# --- CSS 樣式 ---
st.markdown("""
    <style>
    h1, h2, h3 { color: #1a365d; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .report-table { width: 100%; border-collapse: collapse; margin-top: 20px; font-family: 'Segoe UI', sans-serif; }
    .report-table th { background-color: #f1f3f5; color: #1a365d; font-weight: bold; text-align: center !important; padding: 12px; border: 1px solid #dee2e6; }
    .report-table td { padding: 10px; border: 1px solid #dee2e6; text-align: right; }
    .report-table td:first-child { text-align: center; }
    .neg-value { color: #dc3545; font-weight: bold; }
    .kpi-box { padding: 20px; border-radius: 10px; background-color: #f8f9fa; border: 1px solid #dee2e6; text-align: center; }
    .kpi-label { font-size: 1rem; color: #6c757d; margin-bottom: 5px; }
    .kpi-value { font-size: 1.8rem; font-weight: bold; color: #1a365d; }
    .kpi-value-red { font-size: 1.8rem; font-weight: bold; color: #dc3545; }
    .news-box { background-color: #fff9db; padding: 15px; border-left: 5px solid #fcc419; border-radius: 5px; margin-bottom: 20px; }
    .news-title { font-weight: bold; color: #856404; margin-bottom: 5px; }
    .news-content { font-size: 0.95rem; color: #555; }
    </style>
    """, unsafe_allow_html=True)

# --- 資料路徑 ---
BASE_DIR = Path(config.DATA_DIR)
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"
FEE_RATE, TAX_RATE, MIN_FEE = 0.001425, 0.003, 20

# --- 工具函式 ---
@st.cache_data(ttl=3600)
def fetch_portfolio_news(api_key, stock_names):
    if not api_key or not stock_names: return None
    try:
        url = f"https://news.google.com/rss/search?q={urllib.parse.quote(' OR '.join([f'\"{n}\"' for n in stock_names[:5]]) + ' 股市')}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10, verify=False)
        items = ET.fromstring(response.text).findall('./channel/item')[:8]
        news_list = [f"- {item.find('title').text}" for item in items]
        api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
        resp = requests.post(api_url, json={"contents": [{"parts": [{"text": f"分析:\n{chr(10).join(news_list)}\n輸出 JSON：[{{\"title\": \"標題\", \"content\": \"摘要\"}}]"}]}]}, headers={'Content-Type': 'application/json', 'x-goog-api-key': api_key}, timeout=20, verify=False)
        return json.loads(resp.json()['candidates'][0]['content']['parts'][0]['text'].replace("```json", "").replace("```", "").strip())
    except: return None

def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (datetime, date)): return val.date()
    s = str(val).strip().split(' ')[0]
    digits = re.sub(r'\D', '', s)
    return datetime.strptime(digits, '%Y%m%d').date() if len(digits) == 8 else None

def calculate_net_pnl(cost, price, shares):
    val = price * shares
    fee = max(MIN_FEE, cost * shares * FEE_RATE) + max(MIN_FEE, val * FEE_RATE)
    tax = val * TAX_RATE
    return val - (cost * shares) - fee - tax, val, fee + tax

def format_finance_plain(val, is_percent=False):
    try:
        val = float(val)
        if is_percent: return f"{val:.2f}%"
        if val < 0: return f'<span style="color:red">({abs(val):,.0f})</span>'
        return f"{val:,.0f}"
    except: return str(val)

@st.cache_data(ttl=60)
def load_data():
    if not INVENTORY_FILE.exists(): return None, None, ["❌ 找不到庫存檔案"]
    inv = pd.read_excel(INVENTORY_FILE)
    inv.columns = inv.columns.str.strip()
    inv['交易日期'] = inv['交易日期'].apply(parse_date)
    
    csv_files = glob.glob(str(BASE_DIR / "*.csv"))
    all_close = []
    for f in csv_files:
        try:
            df = pd.read_csv(f, encoding='cp950', on_bad_lines='skip')
            # 自動偵測欄位
            name_col = next((c for c in df.columns if any(x in c for x in ['名稱','代號'])), None)
            price_col = next((c for c in df.columns if '收盤' in c), None)
            date_col = next((c for c in df.columns if '日期' in c), None)
            if name_col and price_col and date_col:
                df = df.rename(columns={name_col: '股票名稱', price_col: '收盤價', date_col: '日期'})
                df['日期_dt'] = df['日期'].apply(parse_date)
                all_close.append(df[['股票名稱', '收盤價', '日期_dt']])
        except: continue
    return inv, pd.concat(all_close, ignore_index=True) if all_close else None, []

def load_monitor_configs():
    if not MONITOR_FILE.exists(): return {}
    df = pd.read_excel(MONITOR_FILE)
    df['股票名稱'] = df['股票名稱'].astype(str).str.strip()
    return df.set_index('股票名稱').to_dict('index')

# --- 載入與初始化 ---
inventory_df, close_df, logs = load_data()
monitor_configs = load_monitor_configs()

# --- UI 呈現 ---
st.markdown("### 🔔 市場與持倉監控")
if inventory_df is not None:
    portfolio_stocks = inventory_df['股票名稱'].unique().tolist()
    news = fetch_portfolio_news(getattr(config, 'GEMINI_API_KEY', None), portfolio_stocks)
    if news:
        cols = st.columns(2)
        for i, item in enumerate(news):
            with cols[i % 2]: st.markdown(f'''<div class="news-box"><div class="news-title">{item.get('title')}</div><div class="news-content">{item.get('content')}</div></div>''', unsafe_allow_html=True)

if inventory_df is not None and close_df is not None:
    all_dates = sorted(close_df['日期_dt'].dropna().unique())
    selected_date = st.select_slider("📅 選擇報告基準日：", options=all_dates, value=all_dates[-1])
    
    # 進行計算
    daily_details = []
    for d in all_dates:
        if d > selected_date: continue
        d_close = close_df[close_df['日期_dt'] == d]
        for name, group in inventory_df[inventory_df['交易日期'] <= d].groupby('股票名稱'):
            shares = group['股數'].sum()
            price = d_close[d_close['股票名稱'] == name]['收盤價'].iloc[0] if name in d_close['股票名稱'].values else 0
            if price > 0:
                cost = (group['成本'] * group['股數']).sum() / shares
                pnl, val, _ = calculate_net_pnl(cost, price, shares)
                daily_details.append({"日期": d, "股票名稱": name, "淨損益": pnl, "市值": val})
    
    full_df = pd.DataFrame(daily_details)
    latest_summary = full_df[full_df['日期'] == selected_date]

    # --- [新增功能] 資產與現金流儀表板 ---
    st.markdown("---")
    st.markdown("### 🛡️ 資產與現金流監控儀表板")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🚩 資產集中度")
        total_val = latest_summary['市值'].sum()
        if total_val > 0:
            max_row = latest_summary.nlargest(1, '市值')
            conc = (max_row['市值'].iloc[0] / total_val) * 100
            st.metric("最大單一個股佔比", f"{conc:.1f}%", max_row['股票名稱'].iloc[0])
            st.progress(min(conc/30.0, 1.0))
    with c2:
        st.markdown("#### 💰 被動收入達標監測")
        est_div = 0
        for name, group in inventory_df.groupby('股票名稱'):
            if name in monitor_configs and '預估每股配息' in monitor_configs[name]:
                est_div += (monitor_configs[name]['預估每股配息'] * group['股數'].sum())
        target = getattr(config, 'TARGET_Q_2026', 280000) * 4
        st.metric("2026 預估年配息", f"${est_div:,.0f}")
        st.progress(min(est_div/target, 1.0))
    st.divider()

    # --- 原有分析與圖表 ---
    cols = st.columns(3)
    cols[0].plotly_chart(px.area(full_df, x="日期", y="淨損益", color="股票名稱"), use_container_width=True)
    cols[1].plotly_chart(px.pie(latest_summary, values='市值', names='股票名稱'), use_container_width=True)
    cols[2].plotly_chart(px.bar(latest_summary, x='股票名稱', y='淨損益'), use_container_width=True)

    # --- 投資組合明細 ---
    st.subheader("📋 投資組合明細清單")
    prev_date = all_dates[all_dates.index(selected_date) - 1] if all_dates.index(selected_date) > 0 else selected_date
    prev_data = full_df[full_df['日期'] == prev_date]
    rows_html = ""
    for name in latest_summary['股票名稱']:
        row = latest_summary[latest_summary['股票名稱']==name]
        prev_pnl = prev_data[prev_data['股票名稱']==name]['淨損益'].sum() if name in prev_data['股票名稱'].values else 0
        rows_html += f"<tr><td>{name}</td><td>{row['市值'].iloc[0]:,.0f}</td><td>{format_finance_plain(row['淨損益'].iloc[0])}</td><td>{format_finance_plain(row['淨損益'].iloc[0] - prev_pnl)}</td></tr>"
    st.markdown(f"<table class='report-table'><thead><tr><th>名稱</th><th>市值</th><th>累計損益</th><th>本日獲利(損)</th></tr></thead><tbody>{rows_html}</tbody></table>", unsafe_allow_html=True)
else:
    st.info("👋 歡迎使用！請確認資料夾中已包含 '庫存股票.xlsx' 與 CSV 檔案。")
