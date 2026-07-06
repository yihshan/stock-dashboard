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

# --- 工具函式 ---
BASE_DIR = Path(config.DATA_DIR)
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"
FEE_RATE, TAX_RATE, MIN_FEE = 0.001425, 0.003, 20

def format_finance_plain(val, is_percent=False):
    try:
        val = float(val)
        fmt = f"{val:.2f}%" if is_percent else f"{val:,.0f}"
        if val < 0: return f'<span style="color:red">({abs(val):.2f}%)</span>' if is_percent else f'<span style="color:red">({abs(val):,.0f})</span>'
        return fmt
    except: return str(val)

def calculate_net_pnl(cost, price, shares):
    total_cost_base = cost * shares
    total_market_value = price * shares
    sell_fee = max(MIN_FEE, total_market_value * FEE_RATE)
    sell_tax = total_market_value * TAX_RATE
    net_pnl = total_market_value - total_cost_base - max(MIN_FEE, total_cost_base * FEE_RATE) - sell_fee - sell_tax
    return net_pnl, total_market_value, sell_fee + sell_tax

def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (datetime, date)): return val.date()
    s = str(val).strip().split(' ')[0]
    digits = re.sub(r'\D', '', s)
    if len(digits) == 8: return datetime.strptime(digits, '%Y%m%d').date()
    return None

def clean_price(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(',', '').strip()
    try: return float(s)
    except: return 0.0

@st.cache_data(ttl=60)
def load_data():
    if not INVENTORY_FILE.exists(): return None, None, ["❌ 找不到庫存檔案"]
    try:
        inv = pd.read_excel(INVENTORY_FILE)
        inv.columns = inv.columns.str.strip()
        inv['交易日期'] = inv['交易日期'].apply(parse_date)
        inv = inv.dropna(subset=['交易日期'])
    except: return None, None, ["⚠️ 讀取庫存失敗"]

    all_close_data = []
    for f in glob.glob(str(BASE_DIR / "*.csv")):
        try:
            df = pd.read_csv(f, encoding='cp950' if 'cp950' in f else 'utf-8', on_bad_lines='skip')
            date_col = next((c for c in df.columns if '日期' in c), None)
            name_col = next((c for c in df.columns if any(x in c for x in ['名稱','代號'])), None)
            price_col = next((c for c in df.columns if '收盤' in c), None)
            if date_col and name_col and price_col:
                df = df.rename(columns={name_col: '股票名稱', price_col: '收盤價', date_col: '日期'})
                df['日期_dt'] = df['日期'].apply(parse_date)
                df['收盤價'] = df['收盤價'].apply(clean_price)
                all_close_data.append(df[['股票名稱', '收盤價', '日期_dt']])
        except: continue
    return inv, pd.concat(all_close_data, ignore_index=True) if all_close_data else None, []

@st.cache_data(ttl=60)
def load_monitor_configs():
    if not MONITOR_FILE.exists(): return {}
    df = pd.read_excel(MONITOR_FILE)
    df['股票名稱'] = df['股票名稱'].astype(str).str.strip()
    return df.set_index('股票名稱').to_dict('index')

inventory_df, close_df, logs = load_data()
monitor_configs = load_monitor_configs()

# --- UI 樣式 ---
st.markdown("""<style>
    .report-table { width: 100%; border-collapse: collapse; margin-top: 20px; font-family: sans-serif; }
    .report-table th { background-color: #f1f3f5; color: #1a365d; text-align: center !important; padding: 12px; border: 1px solid #dee2e6; }
    .report-table td { padding: 10px; border: 1px solid #dee2e6; text-align: right; }
    .kpi-box { padding: 20px; border-radius: 10px; background-color: #f8f9fa; border: 1px solid #dee2e6; text-align: center; }
    .kpi-value { font-size: 1.8rem; font-weight: bold; color: #1a365d; }
    .news-box { background-color: #fff9db; padding: 15px; border-left: 5px solid #fcc419; margin-bottom: 20px; }
    .news-title { font-weight: bold; color: #856404; margin-bottom: 5px; }
</style>""", unsafe_allow_html=True)

# --- 主程式 ---
st.markdown("### 🔔 持倉個股重大消息監控")
if inventory_df is not None:
    news = fetch_portfolio_news(getattr(config, 'GEMINI_API_KEY', None), inventory_df['股票名稱'].unique().tolist())
    if news:
        c1, c2 = st.columns(2)
        for i, item in enumerate(news):
            with [c1, c2][i%2]: st.markdown(f'''<div class="news-box"><div class="news-title">{item.get('title')}</div><div class="news-content">{item.get('content')}</div></div>''', unsafe_allow_html=True)

if inventory_df is not None and close_df is not None:
    all_dates = sorted(close_df['日期_dt'].dropna().unique())
    selected_date = st.select_slider("📅 選擇報告基準日：", options=all_dates, value=all_dates[-1])
    
    # 計算損益數據
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

    # --- 資產與現金流儀表板 ---
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
        st.markdown("#### 💰 被動收入監測")
        est_div = 0
        for name, group in inventory_df.groupby('股票名稱'):
            if name in monitor_configs and '預估每股配息' in monitor_configs[name]:
                est_div += (monitor_configs[name]['預估每股配息'] * group['股數'].sum())
        target = getattr(config, 'TARGET_Q_2026', 280000) * 4
        st.metric("2026 預估年配息", f"${est_div:,.0f}")
        st.progress(min(est_div/target, 1.0))
    st.divider()

    # --- KPI 與圖表 ---
    st.markdown("### 🔑 關鍵績效指標")
    
    # 【本日獲利(損) 直覺演算法】：使用 (今日收盤 - 昨日收盤) * 股數
    current_idx = all_dates.index(selected_date)
    prev_date = all_dates[current_idx - 1] if current_idx > 0 else selected_date
    curr_close_df = close_df[close_df['日期_dt'] == selected_date]
    prev_close_df = close_df[close_df['日期_dt'] == prev_date]
    
    total_daily_pnl = 0
    stock_daily_pnl_map = {}
    
    valid_inventory = inventory_df[inventory_df['交易日期'] <= selected_date]
    for name, group in valid_inventory.groupby('股票名稱'):
        shares = group['股數'].sum()
        curr_price = curr_close_df[curr_close_df['股票名稱'] == name]['收盤價'].iloc[0] if name in curr_close_df['股票名稱'].values else 0
        prev_price = prev_close_df[prev_close_df['股票名稱'] == name]['收盤價'].iloc[0] if name in prev_close_df['股票名稱'].values else curr_price
        
        # 單純的價差相乘
        daily_diff = (curr_price - prev_price) * shares
        stock_daily_pnl_map[name] = daily_diff
        total_daily_pnl += daily_diff

    total_net_pnl = latest_summary['淨損益'].sum() if not latest_summary.empty else 0
    
    m1, m2, m3 = st.columns(3)
    with m1: st.markdown(f'<div class="kpi-box"><div class="kpi-label">資產總市值</div><div class="kpi-value">${total_val:,.0f}</div></div>', unsafe_allow_html=True)
    with m2: st.markdown(f'<div class="kpi-box"><div class="kpi-label">累積淨損益</div><div class="kpi-value">{format_finance_plain(total_net_pnl)}</div></div>', unsafe_allow_html=True)
    with m3: st.markdown(f'<div class="kpi-box"><div class="kpi-label">本日獲利(直覺)</div><div class="kpi-value">{format_finance_plain(total_daily_pnl)}</div></div>', unsafe_allow_html=True)

    cols = st.columns(3)
    cols[0].plotly_chart(px.area(full_df, x="日期", y="淨損益", color="股票名稱"), use_container_width=True)
    cols[1].plotly_chart(px.pie(latest_summary, values='市值', names='股票名稱'), use_container_width=True)
    cols[2].plotly_chart(px.bar(latest_summary, x='股票名稱', y='淨損益'), use_container_width=True)
    
    # --- 投資組合明細清單 ---
    st.subheader("📋 投資組合明細清單")
    rows_html = ""
    for name, group in valid_inventory.groupby('股票名稱'):
        shares = group['股數'].sum()
        if shares <= 0: continue
        avg_cost = (group['成本'] * group['股數']).sum() / shares
        curr_price = curr_close_df[curr_close_df['股票名稱'] == name]['收盤價'].iloc[0] if name in curr_close_df['股票名稱'].values else 0
        
        if curr_price > 0:
            pnl, val, fees = calculate_net_pnl(avg_cost, curr_price, shares)
            roi_stock = (pnl / (avg_cost * shares) * 100) if avg_cost > 0 else 0
            # 取得上方計算好的直覺獲利
            daily_pnl_stock = stock_daily_pnl_map.get(name, 0)
            
            rows_html += f"<tr><td>{name}</td><td>{shares:,.0f}</td><td>{avg_cost:,.2f}</td><td>{curr_price:,.2f}</td><td>{val:,.0f}</td><td>{fees:,.0f}</td><td>{format_finance_plain(pnl)}</td><td>{format_finance_plain(daily_pnl_stock)}</td><td>{format_finance_plain(roi_stock, is_percent=True)}</td></tr>"
            
    st.markdown(f"""
        <table class="report-table">
            <thead>
                <tr>
                    <th>股票名稱</th><th>持有股數</th><th>平均成本</th><th>目前市價</th><th>資產市值</th><th>預估稅費</th><th>累積淨損益</th><th>本日獲利(直覺)</th><th>投資報酬率</th>
                </tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
        <p style="color: #666; font-size: 0.8rem; margin-top: 10px;">註：本日獲利採 (今日收盤 - 昨日收盤) * 持有股數 計算，不含稅費。</p>
    """, unsafe_allow_html=True)
else:
    st.info("👋 歡迎使用！請確認資料夾中已包含 '庫存股票.xlsx' 與當日收盤價 CSV 檔案。")
