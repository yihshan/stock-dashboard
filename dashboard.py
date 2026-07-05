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

# --- 2. 工具函式 ---
BASE_DIR = Path(config.DATA_DIR)
INVENTORY_FILE = BASE_DIR / "庫存股票.xlsx"
MONITOR_FILE = BASE_DIR / "監控股票.xlsx"
FEE_RATE, TAX_RATE, MIN_FEE = 0.001425, 0.003, 20

def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (datetime, date)): return val.date() if isinstance(val, datetime) else val
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
    return None

def clean_price(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(',', '').strip()
    try: return float(s)
    except: return 0.0

def format_finance(val, is_percent=False):
    val_str = f"{abs(val):,.2f}%" if is_percent else f"{abs(val):,.0f}"
    return f'<span style="color:red">({val_str})</span>' if val < 0 else val_str

def format_finance_plain(val, is_percent=False):
    try:
        val = float(val)
        fmt = f"{val:.2f}%" if is_percent else f"{val:,.0f}"
        if val < 0: return f'<span style="color:red">({abs(val):.2f}%)</span>' if is_percent else f'<span style="color:red">({abs(val):,.0f})</span>'
        return fmt
    except: return str(val)

def calculate_net_pnl(cost, price, shares):
    val = price * shares
    return val - (cost * shares) - max(MIN_FEE, cost * shares * FEE_RATE) - max(MIN_FEE, val * FEE_RATE) - (val * TAX_RATE), val, max(MIN_FEE, val * FEE_RATE) + (val * TAX_RATE)

@st.cache_data(ttl=60)
def load_data():
    if not INVENTORY_FILE.exists(): return None, None, ["❌ 找不到庫存檔案"]
    inventory_df = pd.read_excel(INVENTORY_FILE)
    inventory_df.columns = inventory_df.columns.str.strip()
    inventory_df['交易日期'] = inventory_df['交易日期'].apply(parse_date)
    inventory_df = inventory_df.dropna(subset=['交易日期'])
    
    csv_files = glob.glob(str(BASE_DIR / "*.csv"))
    all_close_data = []
    for f in csv_files:
        df = pd.read_csv(f, encoding='cp950' if 'cp950' in f else 'utf-8', on_bad_lines='skip')
        for col in df.columns:
            if '收盤' in col:
                df = df.rename(columns={col: '收盤價'})
                break
        all_close_data.append(df)
    
    return inventory_df, (pd.concat(all_close_data) if all_close_data else None), []

@st.cache_data(ttl=60)
def load_monitor_configs():
    if not MONITOR_FILE.exists(): return {}
    df = pd.read_excel(MONITOR_FILE)
    df['股票名稱'] = df['股票名稱'].astype(str).str.strip()
    return df.set_index('股票名稱').to_dict('index')

# --- 載入數據 ---
inventory_df, close_df, logs = load_data()
monitor_configs = load_monitor_configs()

# --- CSS 設定 ---
st.markdown("""<style>.news-box { padding: 15px; border-radius: 8px; background-color: #f8f9fa; border: 1px solid #dee2e6; margin-bottom: 15px; } .news-title { font-weight: bold; color: #004a99; } .kpi-box { padding: 20px; border-radius: 10px; background-color: #f8f9fa; border: 1px solid #dee2e6; text-align: center; } .kpi-value { font-size: 1.8rem; font-weight: bold; color: #1a365d; }</style>""", unsafe_allow_html=True)

# --- 主程式區塊 ---
st.markdown("### 🔔 持倉個股重大消息監控")
portfolio_stocks = inventory_df['股票名稱'].unique().tolist() if inventory_df is not None else []
api_key = os.getenv("GEMINI_API_KEY") or getattr(config, 'GEMINI_API_KEY', None)

with st.spinner("🔄 正在取得並分析最新市場新聞..."):
    dynamic_news = fetch_portfolio_news(api_key, portfolio_stocks)

news_col1, news_col2 = st.columns(2)
if dynamic_news:
    mid = len(dynamic_news) // 2
    with news_col1:
        for item in dynamic_news[:mid]:
            st.markdown(f'''<div class="news-box"><div class="news-title">{item.get('title', '焦點新聞')}</div><div class="news-content">{item.get('content', '')}</div></div>''', unsafe_allow_html=True)
    with news_col2:
        for item in dynamic_news[mid:]:
            st.markdown(f'''<div class="news-box"><div class="news-title">{item.get('title', '焦點新聞')}</div><div class="news-content">{item.get('content', '')}</div></div>''', unsafe_allow_html=True)

if inventory_df is not None and close_df is not None:
    all_dates = sorted(close_df['日期_dt'].unique())
    selected_date = st.select_slider("📅 選擇報告基準日：", options=all_dates, value=all_dates[-1])
    
    # 計算損益數據 (移至前面以供新儀表板使用)
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
                daily_details.append({"日期": d, "股票名稱": name, "淨損益": pnl, "市值": val, "成本": avg_cost * shares})
    
    full_df = pd.DataFrame(daily_details)
    latest_summary = full_df[full_df['日期'] == selected_date]

    # --- 新功能區塊：資產集中度與配息監測 ---
    st.markdown("### 🛡️ 資產與現金流監控儀表板")
    col_risk, col_div = st.columns(2)
    
    with col_risk:
        st.markdown("#### 🚩 資產集中度")
        if not latest_summary.empty:
            total_val = latest_summary['市值'].sum()
            max_stock = latest_summary.nlargest(1, '市值')
            max_name = max_stock['股票名稱'].values[0]
            concentration = (max_stock['市值'].values[0] / total_val) * 100
            st.metric("最大單一個股佔比", f"{concentration:.1f}%", f"集中標的: {max_name}")
            st.progress(min(concentration/30.0, 1.0))
        else:
            st.info("⚠️ 無庫存資料。")
    
    with col_div:
        st.markdown("#### 💰 2026 被動收入進度")
        target_income = getattr(config, 'TARGET_Q_2026', 280000) * 4
        st.metric("預估年配息", "請確認 Excel 配息欄位", help="數據對接中")
        st.progress(0.45)
        st.caption(f"目標：2026 年預計達成年化配息 ${target_income:,.0f}")

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

    # 計算損益數據
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
                daily_details.append({"日期": d, "股票名稱": name, "淨損益": pnl, "市值": val, "成本": avg_cost * shares})
    
    full_df = pd.DataFrame(daily_details)
    filtered_full_df = full_df if selected_stock == "全部個股" else full_df[full_df['股票名稱'] == selected_stock]
    latest_summary = filtered_full_df[filtered_full_df['日期'] == selected_date]
    
    # 計算「本日獲利(損)」
    daily_pnl = 0
    if len(all_dates) >= 2:
        try:
            idx = all_dates.index(selected_date)
            if idx > 0:
                prev_date = all_dates[idx-1]
                current_net_pnl = latest_summary['淨損益'].sum()
                prev_summary = filtered_full_df[filtered_full_df['日期'] == prev_date]
                prev_net_pnl = prev_summary['淨損益'].sum()
                daily_pnl = current_net_pnl - prev_net_pnl
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
    valid_inventory = inventory_df[inventory_df['交易日期'] <= selected_date]
    inventory_to_show = valid_inventory if selected_stock == "全部個股" else valid_inventory[valid_inventory['股票名稱'] == selected_stock]    
    
    # 準備用於計算個股本日獲利(損)的數據 (修正為單純的日期差異計算)
    stock_daily_pnl_map = {}
    if len(all_dates) >= 2:
        try:
            # 取得當前日期的索引
            current_idx = all_dates.index(selected_date)
            if current_idx > 0:
                prev_date = all_dates[current_idx - 1]
                
                # 過濾出當日與前一日的完整數據
                curr_data = full_df[full_df['日期'] == selected_date]
                prev_data = full_df[full_df['日期'] == prev_date]
                
                # 遍歷當日所有持股進行相減
                for name in curr_data['股票名稱'].unique():
                    curr_pnl = curr_data[curr_data['股票名稱'] == name]['淨損益'].sum()
                    # 若前一日無該股，則視為 0
                    prev_pnl = prev_data[prev_data['股票名稱'] == name]['淨損益'].sum() if name in prev_data['股票名稱'].values else 0
                    
                    stock_daily_pnl_map[name] = curr_pnl - prev_pnl
        except Exception as e:
            logger.error(f"計算本日獲利差異時發生錯誤: {e}")

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
        <p style="color: #666; font-size: 0.8rem; margin-top: 10px;">註：負數以括號 ( ) 表示。稅費包含買入手續費、賣出手續費及證交稅。</p>
    """, unsafe_allow_html=True)
else:
    st.info("👋 歡迎使用！請確認 '每日收盤' 資料夾中已包含 '庫存股票.xlsx' 與當日收盤價 CSV 檔案。")
    if logs:
        for log in logs: st.error(log)
