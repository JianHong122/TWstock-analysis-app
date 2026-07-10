import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import io
import re
import time  # 👈 新增這個，用於防止爬蟲被鎖 IP
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go  
from plotly.subplots import make_subplots 
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ==========================================
# 0. 初始化設定與共用函數
# ==========================================
if 'analyzed_input' not in st.session_state:
    st.session_state.analyzed_input = None
if 'target_date' not in st.session_state:
    st.session_state.target_date = None

st.set_page_config(page_title="牧場小霸王", page_icon="📈", layout="wide") # 調整為 wide 讓圖表更好看

@st.cache_data
def load_stock_list():
    try:
        df = pd.read_excel('TW50100.xlsx', engine='openpyxl', dtype=str)
        return {str(row[df.columns[1]]): str(row[df.columns[0]]).replace('.0', '') for _, row in df.iterrows()}, True
    except: return {}, False


# ==========================================
# 副程式 1：抓取 YFinance 資料
# ==========================================
@st.cache_data(ttl=300, show_spinner=False)
def step1_fetch_yf_data(ticker, raw_ticker, auto_fallback, target_date_str):
    end_dt = pd.to_datetime(target_date_str, format='%Y/%m/%d') + pd.Timedelta(days=1)
    start_dt = end_dt - pd.DateOffset(months=6) 
    
    start_str = start_dt.strftime('%Y-%m-%d')
    end_str = end_dt.strftime('%Y-%m-%d')

    hist = yf.Ticker(ticker).history(start=start_str, end=end_str)
    
    if hist.empty and auto_fallback and raw_ticker:
        ticker_two = f"{raw_ticker}.TWO"
        hist_two = yf.Ticker(ticker_two).history(start=start_str, end=end_str)
        if not hist_two.empty:
            hist = hist_two
            ticker = ticker_two
            
    return hist, ticker


# ==========================================
# 副程式 2：產生 K線、均線、KD、MACD
# ==========================================
def step2_calc_tech_indicators(hist):
    df = hist.copy()
    df['MA5'] = df['Close'].rolling(window=5).mean()
    df['MA10'] = df['Close'].rolling(window=10).mean()
    df['MA20'] = df['Close'].rolling(window=20).mean()

    low_min = df['Low'].rolling(window=9).min()
    high_max = df['High'].rolling(window=9).max()
    rsv = (df['Close'] - low_min) / (high_max - low_min + 1e-9) * 100
    df['K'] = rsv.ewm(com=2, adjust=False).mean()
    df['D'] = df['K'].ewm(com=2, adjust=False).mean()
    
    ema12 = df['Close'].ewm(span=12, adjust=False).mean()
    ema26 = df['Close'].ewm(span=26, adjust=False).mean()
    df['DIF'] = ema12 - ema26
    df['MACD'] = df['DIF'].ewm(span=9, adjust=False).mean()
    df['OSC'] = (df['DIF'] - df['MACD']) * 2 
    df['Volume'] = df['Volume'] / 1000  
    return df.tail(64)


# ==========================================
# 副程式 3：產生分價量統計與圖表
# ==========================================
def step3_process_volume_profile(hist_64):
    current_price_round = round(hist_64['Close'].dropna().iloc[-1], 2)
    max_price, min_price = hist_64['High'].max(), hist_64['Low'].min()
    if max_price == min_price:
        max_price, min_price = min_price * 1.05, min_price * 0.95
    
    bin_size = (max_price - min_price) / 20
    curr_bin_idx = 19 if current_price_round >= max_price else (0 if current_price_round <= min_price else min(19, int((current_price_round - min_price) / bin_size)))
    
    bins_data = [{'idx': i, 'start': min_price + i * bin_size, 'end': min_price + (i + 1) * bin_size, 'mid': (min_price + i * bin_size + min_price + (i + 1) * bin_size) / 2, 'label': f"{min_price + i * bin_size:.2f} ~ {min_price + (i + 1) * bin_size:.2f}", 'disp_label': f"{'** ' if i == curr_bin_idx else ''}{min_price + i * bin_size:.2f} ~ {min_price + (i + 1) * bin_size:.2f}", 'is_current': (i == curr_bin_idx), 'vol': 0} for i in range(20)]
    
    all_price_vols = []
    for _, row in hist_64.iterrows():
        o, h, l, c, v = round(row['Open'], 2), round(row['High'], 2), round(row['Low'], 2), round(row['Close'], 2), row['Volume']
        if l > h: l, h = h, l 
        vol_open, vol_close, vol_dist_total = v * 0.05, v * 0.30, v * 0.65
        
        ticks, curr = [], l
        while curr <= h:
            ticks.append(curr)
            ts = 0.01 if curr < 10 else (0.05 if curr < 50 else (0.1 if curr < 100 else (0.5 if curr < 500 else (1.0 if curr < 1000 else 5.0))))
            curr = round(curr + ts, 2)
        
        n_ticks = len(ticks)
        vol_per_tick = vol_dist_total / n_ticks if n_ticks > 0 else 0
        all_price_vols.extend([{'Price': o, 'Vol': vol_open}, {'Price': c, 'Vol': vol_close}] + [{'Price': t, 'Vol': vol_per_tick} for t in ticks])
            
    df_all_vols = pd.DataFrame(all_price_vols)
    for price, vol in df_all_vols.groupby('Price')['Vol'].sum().items():
        if price >= max_price: bins_data[-1]['vol'] += vol
        elif price <= min_price: bins_data[0]['vol'] += vol
        else: bins_data[min(19, int((price - min_price) / bin_size))]['vol'] += vol
            
    all_intervals_disp = sorted(bins_data, key=lambda x: x['idx'], reverse=True)
    
    df_plot = pd.DataFrame({
        '價格區間': [item['label'] for item in all_intervals_disp],
        '累積成交量 (張)': [int(item['vol']) for item in all_intervals_disp],
        '標記': ['現價所在' if item['is_current'] else '一般區間' for item in all_intervals_disp]
    })
    fig_vol = px.bar(df_plot, x='累積成交量 (張)', y='價格區間', color='標記', color_discrete_map={'現價所在': '#FF4B4B', '一般區間': '#60B4FF'}, orientation='h')
    fig_vol.update_yaxes(categoryorder='array', categoryarray=df_plot['價格區間'])
    fig_vol.update_layout(yaxis=dict(title="價格區間", autorange="reversed"), margin=dict(l=0, r=0, t=30, b=0), height=500)
    
    return bins_data, all_intervals_disp, fig_vol, current_price_round


# ==========================================
# 副程式 4：關鍵分價量支撐
# ==========================================
def step4_find_support_resistance(bins_data, current_price_round):
    top_5_above = sorted(sorted([b for b in bins_data if b['mid'] >= current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
    top_5_below = sorted(sorted([b for b in bins_data if b['mid'] < current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
    return top_5_above, top_5_below


# ==========================================
# 副程式 5 & 6 相關：即時下載外資、投信 CSV 與 融資券 JSON
# ==========================================
def fetch_twse_csv_data(date_str, inst_type, debug=False):
    url = f"https://www.twse.com.tw/rwd/zh/fund/{inst_type}?date={date_str}&response=csv"
    try:
        res = requests.get(url, timeout=5)
        res.encoding = 'big5' 
        df = pd.read_csv(io.StringIO(res.text), names=list(range(20)), on_bad_lines='skip')
        if debug and df.empty: st.warning(f"[{date_str}] {inst_type} CSV 取得為空")
        return df
    except Exception as e: 
        if debug: st.error(f"[{date_str}] {inst_type} CSV 請求錯誤: {e}")
        return pd.DataFrame()


def fetch_margin_json_data(date_str, raw_ticker, debug=False):
    """(新功能) 利用 exchangeReport API 下載融資券 JSON，並正確解析 tables 結構"""
    url = f"https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={date_str}&selectType=ALL"
    try:
        res = requests.get(url, timeout=5).json()
        
        stat = res.get('stat', 'Unknown')
        if debug: st.write(f"👉 請求融資券 API: `{date_str}` | 狀態: `{stat}`")
        
        if stat == 'OK':
            # 【Debug 升級】如實輸出 JSON 最外層的結構，讓你看見真實面貌
            if debug: st.write(f"🔍 成功收到資料！外層擁有的標籤 (Keys): `{list(res.keys())}`")
            
            # 正確對應 MI_MARGN 的 tables 結構
            tables = res.get('tables', [])
            
            # (保險機制) 萬一哪天證交所又改回 data，我們也接得住
            if not tables and 'data' in res:
                tables = [{'data': res['data']}]
                
            found = False
            for table in tables:
                for row in table.get('data', []):
                    # row[0] 是股票代號
                    if str(row[0]).strip() == raw_ticker:
                        found = True
                        if debug: st.success(f"✅ 找到 {raw_ticker} 融資券資料！\n原始陣列內容: `{row}`")
                        
                        # 依據固定欄位位置解析 (5:融資前日, 6:融資今日, 11:融券前日, 12:融券今日)
                        m_prev = int(str(row[5]).replace(',', ''))
                        m_today = int(str(row[6]).replace(',', ''))
                        s_prev = int(str(row[11]).replace(',', ''))
                        s_today = int(str(row[12]).replace(',', ''))
                        
                        return (m_today - m_prev), m_today, (s_today - s_prev), s_today
            
            if not found and debug:
                st.warning(f"⚠️ API 狀態 OK，但遍歷了整個表格，找不到代號 {raw_ticker} 的資料。")
                
        else:
            if debug: st.warning(f"⚠️ {date_str} 證交所回傳狀態為 {stat} (無交易或假日)。")
            
    except Exception as e: 
        if debug: st.error(f"❌ 請求或解析 JSON 發生錯誤: {e}")
        
    return 0, 0, 0, 0

def step6_extract_10day_institutional_data(raw_ticker, hist_64, debug=False):
    last_10_dates = hist_64.index[-10:]
    
    foreign_records = []
    trust_records = []
    margin_records = []
    
    if debug: st.info("開始爬取近 10 日籌碼資料，啟動防擋機制 (每次停頓 1.5 秒)...")
    
    for d in last_10_dates:
        date_api_str = d.strftime('%Y%m%d')
        date_disp_str = d.strftime('%m/%d')
        
        if debug: st.write(f"--- 處理日期: {date_disp_str} ---")
        
        # 1. 外資
        df_foreign = fetch_twse_csv_data(date_api_str, "TWT38U", debug)
        net_f = 0
        if not df_foreign.empty:
            df_foreign[1] = df_foreign[1].astype(str).str.replace(r'[=" ]', '', regex=True)
            target_row = df_foreign[df_foreign[1] == raw_ticker]
            if not target_row.empty:
                try: net_f = round(int(str(target_row.iloc[0, 5]).replace(',', '').strip()) / 1000)
                except: pass
        foreign_records.append({'日期': date_disp_str, '外資買賣超(張)': net_f})
        time.sleep(0.5) # 防擋延遲
        
        # 2. 投信
        df_trust = fetch_twse_csv_data(date_api_str, "TWT44U", debug)
        net_t = 0
        if not df_trust.empty:
            df_trust[1] = df_trust[1].astype(str).str.replace(r'[=" ]', '', regex=True)
            target_row = df_trust[df_trust[1] == raw_ticker]
            if not target_row.empty:
                try: net_t = round(int(str(target_row.iloc[0, 5]).replace(',', '').strip()) / 1000)
                except: pass
        trust_records.append({'日期': date_disp_str, '投信買賣超(張)': net_t})
        time.sleep(0.5) # 防擋延遲
        
        # 3. 融資券
        m_change, m_today, s_change, s_today = fetch_margin_json_data(date_api_str, raw_ticker, debug)
        margin_records.append({
            '日期': date_disp_str,
            '融資變動(張)': m_change,
            '融資餘額(張)': m_today,
            '融券變動(張)': s_change,
            '融券餘額(張)': s_today
        })
        time.sleep(0.5) # 防擋延遲
        
    df_f_res = pd.DataFrame(foreign_records)
    df_t_res = pd.DataFrame(trust_records)
    df_m_res = pd.DataFrame(margin_records)
    
    # 圖 1: 外資
    fig_f = px.bar(df_f_res, x='日期', y='外資買賣超(張)', title='近10日外資買賣超狀況', text_auto=True)
    fig_f.update_traces(marker_color=['#FF4B4B' if val > 0 else '#00B050' for val in df_f_res['外資買賣超(張)']])
    fig_f.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=300)
    
    # 圖 2: 投信
    fig_t = px.bar(df_t_res, x='日期', y='投信買賣超(張)', title='近10日投信買賣超狀況', text_auto=True)
    fig_t.update_traces(marker_color=['#FF4B4B' if val > 0 else '#00B050' for val in df_t_res['投信買賣超(張)']])
    fig_t.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=300)

    # 🛑 移除了融資融券的 Plotly 畫圖邏輯，回傳變數也少掉 fig_ml, fig_ms
    return df_f_res, df_t_res, df_m_res, fig_f, fig_t


# ==========================================
# 介面繪製輔助函數 (Tech Chart)
# ==========================================
def render_tech_chart(hist_64, show_ma5, show_ma10, show_ma20, allow_zoom):
    date_strings = hist_64.index.strftime('%Y-%m-%d')
    fig_k = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.03, row_heights=[0.5, 0.25, 0.25], subplot_titles=("價格與均線", "KD (9,3,3)", "MACD (12,26,9)"))
    
    fig_k.add_trace(go.Candlestick(x=date_strings, open=hist_64['Open'], high=hist_64['High'], low=hist_64['Low'], close=hist_64['Close'], name='K線', increasing_line_color='#FF4B4B', increasing_fillcolor='#FF4B4B', decreasing_line_color='#00B050', decreasing_fillcolor='#00B050'), row=1, col=1)
    if show_ma5: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA5'], name='5MA', line=dict(color='#7A431D', width=1.5)), row=1, col=1)
    if show_ma10: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA10'], name='10MA', line=dict(color='#00E5FF', width=1.5)), row=1, col=1)
    if show_ma20: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA20'], name='20MA', line=dict(color='#0D47A1', width=1.5)), row=1, col=1)
    
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['K'], name='K值', line=dict(color='#FF9900', width=1.2)), row=2, col=1)
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['D'], name='D值', line=dict(color='#0066FF', width=1.2)), row=2, col=1)
    
    macd_colors = ['#FF4B4B' if val > 0 else '#00B050' for val in hist_64['OSC']]
    fig_k.add_trace(go.Bar(x=date_strings, y=hist_64['OSC'], name='OSC', marker_color=macd_colors), row=3, col=1)
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['DIF'], name='DIF', line=dict(color='#FF9900', width=1.2)), row=3, col=1)
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MACD'], name='MACD', line=dict(color='#0066FF', width=1.2)), row=3, col=1)
    
    fig_k.update_layout(xaxis=dict(type='category', visible=False), xaxis2=dict(type='category', visible=False), xaxis3=dict(type='category', visible=True, title="交易日期", nticks=10), yaxis=dict(visible=False), yaxis2=dict(visible=True), yaxis3=dict(visible=True), xaxis_rangeslider_visible=False, margin=dict(l=4, r=4, t=30, b=4), height=700, hovermode='x unified', showlegend=False)
    fig_k.update_xaxes(fixedrange=not allow_zoom)
    fig_k.update_yaxes(fixedrange=not allow_zoom)
    return fig_k


# ==========================================
# 🚀 系統主程式 (Main Program)
# ==========================================
st.title("📊 牧場小霸王")
st.markdown("支援 **技術K線均線**、**KD/MACD**、**分價量防守** 與 **三大法人/融資券籌碼分析**")

name_to_ticker, list_loaded = load_stock_list()
if not list_loaded: st.warning("⚠️ 找不到 'TW50100.xlsx'，請直接輸入股票代號。")

user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")

default_date = datetime.now().strftime("%Y/%m/%d")
target_date_input = st.text_input("📅 請輸入查詢基準日 (西元年/月/日)：", value=default_date, placeholder="例如: 2024/01/01")

if st.button("🚀 開始分析", use_container_width=True):
    input_date_str = target_date_input.strip()
    if not input_date_str: input_date_str = default_date
        
    try:
        datetime.strptime(input_date_str, "%Y/%m/%d")
        st.session_state.analyzed_input = user_input
        st.session_state.target_date = input_date_str
    except ValueError:
        st.error("⚠️ 日期格式錯誤！請輸入正確的「西元年/月/日」格式，例如：2024/01/01")
        st.stop() 

# ------------------------------------
if st.session_state.analyzed_input:
    current_target = st.session_state.analyzed_input
    
    matched_names = [name for name in name_to_ticker.keys() if current_target in name] if list_loaded else []
    if len(matched_names) == 0:
        target_name = f"自訂代號 ({current_target})"
        auto_fallback = False if current_target.lower().endswith(('.tw', '.two')) else True
        raw_ticker = current_target.split('.')[0]
        yf_ticker = current_target.upper() if not auto_fallback else f"{raw_ticker}.TW"
    elif len(matched_names) > 1:
        st.error(f"⚠️ 找到多檔股票，請輸入更明確的名稱：{', '.join(matched_names)}")
        st.stop()
    else:
        target_name = matched_names[0]
        raw_ticker = name_to_ticker[target_name]
        yf_ticker, auto_fallback = f"{raw_ticker}.TW", True

    with st.spinner('📡 正在運算核心技術指標與分價量...'):
        hist, yf_ticker = step1_fetch_yf_data(yf_ticker, raw_ticker, auto_fallback, st.session_state.target_date)
        if hist.empty:
            st.error("❌ 無法取得該日期之前的歷史資料。請確認代號與日期。")
            st.stop()
            
        hist_64 = step2_calc_tech_indicators(hist)
        latest = hist_64.iloc[-1]
        
        bins_data, all_intervals_disp, fig_vol, current_price_round = step3_process_volume_profile(hist_64)
        top_5_above, top_5_below = step4_find_support_resistance(bins_data, current_price_round)

    actual_last_date = hist_64.index[-1].strftime('%Y/%m/%d')
    st.success(f"✅ {target_name} ({yf_ticker}) 分析完成！實際查詢基準日: **{actual_last_date}** / 股價: **{current_price_round:.2f}**")

    # 顯示指標表
    st.subheader("📊 技術指標參考")
    st.table(pd.DataFrame({
        "項目": ["均線狀況", "KD狀況", "MACD狀況"],
        "狀態": ["✅ 多頭" if latest['MA5'] > latest['MA10'] > latest['MA20'] else ("⚠️ 空頭" if latest['MA5'] < latest['MA10'] < latest['MA20'] else "⭕ 盤整"), 
                 "✅ 多" if latest['K'] > latest['D'] else "⚠️ 空", 
                 "✅ 多" if latest['DIF'] > latest['MACD'] else "⚠️ 空"],
        "數值細項": [f"5MA:{latest['MA5']:.1f} / 10MA:{latest['MA10']:.1f}", f"K:{latest['K']:.1f} / D:{latest['D']:.1f}", f"DIF:{latest['DIF']:.1f} / MACD:{latest['MACD']:.1f}"]
    }))

    allow_zoom = st.checkbox("🔍 啟用圖表縮放與拖曳", value=False)
    with st.container(border=True):
        st.subheader("📈 技術分析綜合儀表板")
        c1, c2, c3 = st.columns(3)
        fig_tech = render_tech_chart(hist_64, c1.checkbox("顯示 5MA", value=False), c2.checkbox("顯示 10MA", value=True), c3.checkbox("顯示 20MA", value=False), allow_zoom)
        st.plotly_chart(fig_tech, use_container_width=True)

    st.subheader("📊 64日分價量參考圖")
    fig_vol.update_xaxes(fixedrange=not allow_zoom)
    fig_vol.update_yaxes(fixedrange=not allow_zoom)
    st.plotly_chart(fig_vol, use_container_width=True)

    st.subheader("🎯 關鍵支撐與壓力 (Top 5)")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**⬆️ 向上方壓力區**")
        for item in top_5_above: st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")
    with col2:
        st.write("**⬇️ 向下方支撐區**")
        for item in top_5_below: st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")

   # ----------------------------------------------------
    # 背景獨立執行「法人及融資券籌碼分析」
    # ----------------------------------------------------
    st.divider()
    col_title, col_debug = st.columns([3, 1])
    with col_title:
        st.subheader("📈 近 10 日市場籌碼動向 (外資 / 投信 / 融資券)")
    with col_debug:
        is_debug_mode = st.checkbox("🐛 開啟開發者 Debug 模式")
        
    is_otc = yf_ticker.endswith('.TWO')
    
    df_foreign_export = pd.DataFrame()
    df_trust_export = pd.DataFrame()
    df_margin_export = pd.DataFrame()
    
    if is_otc:
        st.info("⚠️ 該股為上櫃股票，目前僅支援上市股票之三大法人與信用交易籌碼查詢。")
    else:
        with st.spinner("⏳ 正在即時向證交所下載近 10 日 CSV 法人數據與 JSON 融資券數據，請稍候... (為防被鎖 IP 會稍微停頓)"):
            
            if is_debug_mode:
                with st.expander("🛠️ Debug 運作日誌 (展開查看)", expanded=True):
                    # 配合副程式修改，拿掉 fig_ml, fig_ms
                    df_foreign_export, df_trust_export, df_margin_export, fig_f, fig_t = step6_extract_10day_institutional_data(raw_ticker, hist_64, debug=True)
            else:
                df_foreign_export, df_trust_export, df_margin_export, fig_f, fig_t = step6_extract_10day_institutional_data(raw_ticker, hist_64, debug=False)
            
        # UI 佈局：分兩排顯示
        # 第一排：外資與投信 (維持圖表)
        col_f, col_t = st.columns(2)
        with col_f: st.plotly_chart(fig_f, use_container_width=True)
        with col_t: st.plotly_chart(fig_t, use_container_width=True)
        
        # 第二排：融資與融券 (改用乾淨表格呈現)
        if not df_margin_export.empty:
            st.markdown("#### 📊 近 10 日信用交易明細 (張)")
            col_m, col_s = st.columns(2)
            
            # 使用 set_index('日期') 可以讓表格最左邊沒有多餘的數字索引，版面更整齊
            with col_m:
                st.write("**💰 融資狀況 (散戶做多指標)**")
                st.dataframe(df_margin_export[['日期', '融資變動(張)', '融資餘額(張)']].set_index('日期'), use_container_width=True)
            with col_s:
                st.write("**📉 融券狀況 (散戶做空指標)**")
                st.dataframe(df_margin_export[['日期', '融券變動(張)', '融券餘額(張)']].set_index('日期'), use_container_width=True)

    # ----------------------------------------------------
    # 結尾：Excel 報表匯出
    # ----------------------------------------------------
    st.divider()
    st.subheader("💾 匯出完整 Excel 報表")
    try:
        output = io.BytesIO()
        df_sr_excel = pd.DataFrame([{'項次': i+1, '價格級距區間 (TWD)': item['disp_label'], '累積成交量 (張)': int(item['vol'])} for i, item in enumerate(all_intervals_disp)])
        df_top5_excel = pd.DataFrame([{'位置': '⬆️ 向上壓力區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_above] + [{'位置': '🎯 最新股價', '價格級距區間': f"{current_price_round:.2f}", '累積成交量 (張)': 0}] + [{'位置': '⬇️ 向下支撐區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_below])
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_sr_excel.to_excel(writer, sheet_name='區間分價量總表', index=False)
            df_top5_excel.to_excel(writer, sheet_name='關鍵支撐壓力', index=False)
            
            if not df_foreign_export.empty: 
                df_foreign_export.to_excel(writer, sheet_name='外資買賣超(10日)', index=False)
            if not df_trust_export.empty:
                df_trust_export.to_excel(writer, sheet_name='投信買賣超(10日)', index=False)
            if not df_margin_export.empty:
                df_margin_export.to_excel(writer, sheet_name='融資券狀況(10日)', index=False)
            
            workbook = writer.book
            for ws in workbook.worksheets:
                for col in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 25.5
            
            sheet1 = workbook['區間分價量總表']
            chart = BarChart()
            chart.type, chart.style = "bar", 10
            chart.title = f"{target_name} 64日分價量分佈圖"
            chart.x_axis.title, chart.y_axis.title = "價格區間", "成交量"
            chart.height, chart.width = max(10, len(df_sr_excel) * 0.5) * 1.5, 24
            chart.add_data(Reference(sheet1, min_col=3, min_row=1, max_row=len(df_sr_excel) + 1), titles_from_data=True)
            chart.set_categories(Reference(sheet1, min_col=2, min_row=2, max_row=len(df_sr_excel) + 1))
            sheet1.add_chart(chart, "E2")
            
        st.download_button("📥 點我下載 Excel 分析報表", data=output.getvalue(), file_name=f"{re.sub(r'[\\/*?:\"<>|]', '_', target_name)}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    except Exception as e:
        st.error(f"❌ Excel 產生錯誤：{e}")
