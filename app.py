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
import urllib3 # 👈 新增這個
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning) # 👈 關閉煩人的 SSL 警告

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
# 副程式 1：抓取 YFinance 資料 (加入防擋偽裝機制)
# ==========================================
@st.cache_data(ttl=300, show_spinner=False)
def step1_fetch_yf_data(ticker, raw_ticker, auto_fallback, target_date_str):
    end_dt = pd.to_datetime(target_date_str, format='%Y/%m/%d') + pd.Timedelta(days=1)
    start_dt = end_dt - pd.DateOffset(months=6) 
    
    start_str = start_dt.strftime('%Y-%m-%d')
    end_str = end_dt.strftime('%Y-%m-%d')

    # 👇 核心解法：建立自訂 Session，偽裝成一般的 Windows Chrome 瀏覽器
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })

    # 將 session 傳遞給 yf.Ticker
    hist = yf.Ticker(ticker, session=session).history(start=start_str, end=end_str)
    
    if hist.empty and auto_fallback and raw_ticker:
        ticker_two = f"{raw_ticker}.TWO"
        hist_two = yf.Ticker(ticker_two, session=session).history(start=start_str, end=end_str)
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

# 👇 1. TWSE (上市) 專用下載函數
@st.cache_data(ttl=86400, show_spinner=False)
def download_twse_csv_text(date_str, inst_type):
    url = f"https://www.twse.com.tw/rwd/zh/fund/{inst_type}?date={date_str}&response=csv"
    time.sleep(1)
    try:
        res = requests.get(url, timeout=5)
        res.encoding = 'big5'
        if len(res.text) > 100:
            return res.text
        return ""
    except:
        return ""

def fetch_twse_csv_data(date_str, inst_type):
    csv_text = download_twse_csv_text(date_str, inst_type)
    if csv_text:
        try:
            return pd.read_csv(io.StringIO(csv_text), names=list(range(20)), on_bad_lines='skip')
        except: pass
    return pd.DataFrame()

def fetch_margin_json_data(date_str, raw_ticker):
    """利用 exchangeReport API 下載上市融資券 JSON"""
    url = f"https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={date_str}&selectType=ALL"
    try:
        res = requests.get(url, timeout=5).json()
        if res.get('stat') == 'OK':
            tables = res.get('tables', [])
            if not tables and 'data' in res:
                tables = [{'data': res['data']}]
                
            for table in tables:
                for row in table.get('data', []):
                    if str(row[0]).strip() == raw_ticker:
                        m_prev = int(str(row[5]).replace(',', ''))
                        m_today = int(str(row[6]).replace(',', ''))
                        s_prev = int(str(row[11]).replace(',', ''))
                        s_today = int(str(row[12]).replace(',', ''))
                        return (m_today - m_prev), m_today, (s_today - s_prev), s_today
    except: pass
    return 0, 0, 0, 0

# 👇 2. TPEx (上櫃) 專用下載函數
@st.cache_data(ttl=86400, show_spinner=False)
def download_tpex_csv_text(date_str, inst_type):
    """下載上櫃法人 CSV，強制略過 SSL 檢查並使用 big5 解碼"""
    url = f"https://www.tpex.org.tw/www/zh-tw/insti/{inst_type}?type=Daily&date={date_str}&searchType=buy&id=&response=csv"
    time.sleep(1) 
    try:
        res = requests.get(url, timeout=5, verify=False)
        res.encoding = 'big5'
        if len(res.text) > 100: 
            return res.text
    except: pass
    return ""

def fetch_tpex_margin_json_data(roc_date_str, raw_ticker):
    """下載上櫃融資券 JSON，強制略過 SSL 檢查"""
    url = f"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=json&d={roc_date_str}"
    try:
        res = requests.get(url, timeout=5, verify=False).json()
        for row in res.get('aaData', []):
            if str(row[0]).strip() == raw_ticker:
                m_prev = int(str(row[5]).replace(',', ''))
                m_today = int(str(row[6]).replace(',', ''))
                s_prev = int(str(row[11]).replace(',', ''))
                s_today = int(str(row[12]).replace(',', ''))
                return (m_today - m_prev), m_today, (s_today - s_prev), s_today
    except: pass
    return 0, 0, 0, 0

# 👇 3. 核心迴圈：依據上市或上櫃進行資料分流
def step6_extract_institutional_data(raw_ticker, hist_64, is_otc):
    last_20_dates = hist_64.index[-20:]
    last_10_dates = hist_64.index[-10:]
    
    foreign_records = []
    trust_records = []
    margin_records = []
    
    for d in last_20_dates:
        date_disp_str = d.strftime('%m/%d')
        
        if not is_otc:
            # ==========================================
            # 🟢 上市 (TWSE) 邏輯分支
            # ==========================================
            date_api_str = d.strftime('%Y%m%d')
            
            # 1. 外資 (20天)
            df_foreign = fetch_twse_csv_data(date_api_str, "TWT38U")
            net_f = 0
            if not df_foreign.empty:
                df_foreign[1] = df_foreign[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_foreign[df_foreign[1] == raw_ticker]
                if not target_row.empty:
                    try: net_f = round(int(str(target_row.iloc[0, 5]).replace(',', '').strip()) / 1000)
                    except: pass
            foreign_records.append({'日期': date_disp_str, '外資買賣超(張)': net_f})
            
            # 2. 投信 (20天)
            df_trust = fetch_twse_csv_data(date_api_str, "TWT44U")
            net_t = 0
            if not df_trust.empty:
                df_trust[1] = df_trust[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_trust[df_trust[1] == raw_ticker]
                if not target_row.empty:
                    try: net_t = round(int(str(target_row.iloc[0, 5]).replace(',', '').strip()) / 1000)
                    except: pass
            trust_records.append({'日期': date_disp_str, '投信買賣超(張)': net_t})
            
            # 3. 融資券 (10天)
            if d in last_10_dates:
                m_change, m_today, s_change, s_today = fetch_margin_json_data(date_api_str, raw_ticker)
                margin_records.append({'日期': date_disp_str, '融資變動(張)': m_change, '融資餘額(張)': m_today, '融券變動(張)': s_change, '融券餘額(張)': s_today})
                time.sleep(0.5)
                
        else:
            # ==========================================
            # 🟢 上櫃 (TPEx) 邏輯分支
            # ==========================================
            date_tpex_csv_str = d.strftime('%Y%%2F%m%%2F%d')
            
            # 1. 上櫃外資 (20天)
            csv_f_text = download_tpex_csv_text(date_tpex_csv_str, "qfiiStat")
            net_f = 0
            if csv_f_text:
                df_f = pd.read_csv(io.StringIO(csv_f_text), names=list(range(20)), on_bad_lines='skip')
                df_f[0] = df_f[0].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_f[df_f[0] == raw_ticker]
                if not target_row.empty:
                    try: net_f = round(int(str(target_row.iloc[0, 4]).replace(',', '').strip()) / 1000)
                    except: pass
            foreign_records.append({'日期': date_disp_str, '外資買賣超(張)': net_f})
            
            # 2. 上櫃投信 (20天)
            csv_t_text = download_tpex_csv_text(date_tpex_csv_str, "sitcStat")
            net_t = 0
            if csv_t_text:
                df_t = pd.read_csv(io.StringIO(csv_t_text), names=list(range(20)), on_bad_lines='skip')
                df_t[0] = df_t[0].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_t[df_t[0] == raw_ticker]
                if not target_row.empty:
                    try: net_t = round(int(str(target_row.iloc[0, 4]).replace(',', '').strip()) / 1000)
                    except: pass
            trust_records.append({'日期': date_disp_str, '投信買賣超(張)': net_t})
            
            # 3. 上櫃融資券 (10天)
            if d in last_10_dates:
                roc_date_str = f"{d.year - 1911}/{d.strftime('%m/%d')}"
                m_change, m_today, s_change, s_today = fetch_tpex_margin_json_data(roc_date_str, raw_ticker)
                margin_records.append({'日期': date_disp_str, '融資變動(張)': m_change, '融資餘額(張)': m_today, '融券變動(張)': s_change, '融券餘額(張)': s_today})
                time.sleep(0.5)

    # --- 以下統一將資料打包成 DataFrame 並畫圖 ---
    df_f_res = pd.DataFrame(foreign_records)
    df_t_res = pd.DataFrame(trust_records)
    df_m_res = pd.DataFrame(margin_records)
    
    fig_f = px.bar(df_f_res, x='日期', y='外資買賣超(張)', title='近20日外資買賣超狀況', text_auto=True)
    fig_f.update_traces(marker_color=['#FF4B4B' if val > 0 else '#00B050' for val in df_f_res['外資買賣超(張)']])
    fig_f.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=300)
    
    fig_t = px.bar(df_t_res, x='日期', y='投信買賣超(張)', title='近20日投信買賣超狀況', text_auto=True)
    fig_t.update_traces(marker_color=['#FF4B4B' if val > 0 else '#00B050' for val in df_t_res['投信買賣超(張)']])
    fig_t.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=300)

    return df_f_res, df_t_res, df_m_res, fig_f, fig_t


# ==========================================
# 🚀 系統主程式 (Main Program)
# ==========================================
st.title("📊 牧場小霸王")
st.markdown("支援 **技術K線均線**、**KD/MACD**、**分價量防守** 與 **三大法人/融資券籌碼分析**")

name_to_ticker, list_loaded = load_stock_list()
if not list_loaded: st.warning("⚠️ 找不到 'TW50100.xlsx'，請直接輸入股票代號。")

user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")

default_date = get_latest_trading_date()
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
    
    # 模糊搜尋：找出所有包含輸入字串的股票名稱
    matched_names = [name for name in name_to_ticker.keys() if current_target in name] if list_loaded else []
    
    # 👇👇👇 核心修正：完全命中攔截 👇👇👇
    if len(matched_names) > 1 and current_target in matched_names:
        matched_names = [current_target]  # 如果有完全一模一樣的，就只留下它
    # 👆👆👆 核心修正結束 👆👆👆

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

    # 👇👇👇 從這裡開始插入新的 MA 落點分析 👇👇👇
    
    st.divider()
    st.subheader("📏 均線落點分價區間")
    col_ma1, col_ma2, col_ma3 = st.columns(3)
    
    # 將三個 MA 的數值與對應的 UI 欄位綁定
    ma_settings = [
        (col_ma1, "5MA", latest['MA5']), 
        (col_ma2, "10MA", latest['MA10']), 
        (col_ma3, "20MA", latest['MA20'])
    ]
    
    for col, ma_name, ma_val in ma_settings:
        with col:
            if pd.isna(ma_val):  # 防呆：如果上市天數不足，均線算不出數值
                st.write(f"**{ma_name}**：無資料")
                continue
                
            # 尋找該 MA 坐落的籌碼區間
            target_bin = None
            for b in all_intervals_disp:
                if b['start'] <= ma_val <= b['end']:
                    target_bin = b
                    break
                    
            # 極端防呆：如果均線價格噴太高或跌太深，超出了目前 K 線畫出的 20 個區間
            if not target_bin:
                if ma_val > all_intervals_disp[0]['end']: 
                    target_bin = all_intervals_disp[0]  # 代入最高區間
                else: 
                    target_bin = all_intervals_disp[-1] # 代入最低區間
            
            # 顯示結果
            st.markdown(f"**{ma_name} ({ma_val:.2f})**")
            st.write(f"落於區間：`{target_bin['label']}`")
            st.write(f"區間籌碼：**{int(target_bin['vol']):,}** 張")

    # 👆👆👆 插入結束 👆👆👆
    
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
    
    # 🐛 加入 Debug 模式開關
    show_debug = st.checkbox("🐛 開啟爬蟲 Debug 模式 (用來檢視上櫃原始資料的欄位位置)")
    
    c_title, c_btn = st.columns([4, 1])
    with c_title:
        st.subheader("📈 近期市場籌碼動向 (外資投信 20日 / 融資券 10日)")
    with c_btn:
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 重新下載快取", use_container_width=True):
            download_twse_csv_text.clear()
            download_tpex_csv_text.clear()
            
    is_otc = yf_ticker.endswith('.TWO')
    
    # ==========================================
    # 🐛 Debug 模式專用顯示區塊
    # ==========================================
    if show_debug and is_otc:
        st.warning("🔍 【Debug 模式】目前為上櫃股票，以下印出最新一天的原始資料，請從 0 開始數，確認目標在第幾個欄位 (Index)：")
        test_d = hist_64.index[-1]
        test_tpex_csv = test_d.strftime('%Y%%2F%m%%2F%d')
        test_roc = f"{test_d.year - 1911}/{test_d.strftime('%m/%d')}"
        
        c_d1, c_d2 = st.columns(2)
        with c_d1:
            st.write(f"👉 **外資 CSV 原始檔** ({test_tpex_csv})")
            f_txt = download_tpex_csv_text(test_tpex_csv, "qfiiStat")
            if f_txt:
                df_debug = pd.read_csv(io.StringIO(f_txt), names=list(range(12)), on_bad_lines='skip')
                st.dataframe(df_debug.head(10))
            else:
                st.error("無資料或連線失敗")
        with c_d2:
            st.write(f"👉 **融資券 JSON 第一筆資料** ({test_roc})")
            url_m = f"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=json&d={test_roc}"
            try:
                res_m = requests.get(url_m, timeout=5).json()
                if 'aaData' in res_m and len(res_m['aaData']) > 0:
                    st.write(res_m['aaData'][0]) # 展開第一筆陣列讓您數欄位
                else:
                    st.write("該日無 JSON 資料回傳")
            except Exception as e:
                st.error(f"解析失敗: {e}")
        
        st.info("💡 檢查完畢後，請取消勾選 Debug 模式，即可恢復正常繪圖。")
        st.stop() # 暫停主程式，不再往下畫圖，保持畫面乾淨
    # ==========================================
    
    with st.spinner("⏳ 正在即時向證交所/櫃買中心調閱籌碼數據，請稍候..."):
        df_foreign_export, df_trust_export, df_margin_export, fig_f, fig_t = step6_extract_institutional_data(raw_ticker, hist_64, is_otc)
        
    # UI 佈局：分兩排顯示
    col_f, col_t = st.columns(2)
    with col_f: st.plotly_chart(fig_f, use_container_width=True)
    with col_t: st.plotly_chart(fig_t, use_container_width=True)
    
    # 第二排：融資與融券
    if not df_margin_export.empty:
        st.markdown("#### 📊 近 10 日信用交易明細 (張)")
        col_m, col_s = st.columns(2)
        
        df_margin_reversed = df_margin_export.iloc[::-1].set_index('日期')
        
        with col_m:
            st.write("**💰 融資狀況 (散戶做多指標)**")
            st.dataframe(df_margin_reversed[['融資變動(張)', '融資餘額(張)']], use_container_width=True)
        with col_s:
            st.write("**📉 融券狀況 (散戶做空指標)**")
            st.dataframe(df_margin_reversed[['融券變動(張)', '融券餘額(張)']], use_container_width=True)

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
                df_foreign_export.to_excel(writer, sheet_name='外資買賣超(20日)', index=False)
            if not df_trust_export.empty:
                df_trust_export.to_excel(writer, sheet_name='投信買賣超(20日)', index=False)
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
