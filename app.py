import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import math
import re
import io
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

st.set_page_config(page_title="台股籌碼分析工具", page_icon="📈", layout="centered")

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
def step1_fetch_yf_data(ticker, raw_ticker, auto_fallback):
    """負責從 Yahoo Finance 抓取歷史股價資料"""
    hist = yf.Ticker(ticker).history(period="6mo")
    if hist.empty and auto_fallback and raw_ticker:
        ticker_two = f"{raw_ticker}.TWO"
        hist_two = yf.Ticker(ticker_two).history(period="6mo")
        if not hist_two.empty:
            hist = hist_two
            ticker = ticker_two
    return hist, ticker


# ==========================================
# 副程式 2：產生 K線、均線、KD、MACD
# ==========================================
def step2_calc_tech_indicators(hist):
    """負責所有技術指標的數學運算"""
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
    """負責 20 級距切分、籌碼攤平演算法，並產出 Plotly 圖表"""
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
    """負責過濾出上下 Top 5 的支撐與壓力區"""
    top_5_above = sorted(sorted([b for b in bins_data if b['mid'] >= current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
    top_5_below = sorted(sorted([b for b in bins_data if b['mid'] < current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
    return top_5_above, top_5_below


# ==========================================
# 副程式 5：從證交所抓取法人資料
# ==========================================
def step5_fetch_twse_data(raw_ticker, hist_64, is_otc):
    """負責爬取近 9 日的三大法人買賣超資料"""
    last_9_dates = hist_64.index[-9:]
    daily_records = []
    
    if is_otc:
        for d in last_9_dates:
            daily_records.append({'date_disp': d.strftime('%Y-%m-%d'), 'net_buy': 0, 'volume': int(hist_64.loc[d, 'Volume'])})
        return daily_records
        
    for d in last_9_dates:
        month_str, date_str = d.strftime('%Y%m01'), d.strftime('%Y%m%d')
        tw_date_str = f"{d.year - 1911}/{d.strftime('%m/%d')}"
        daily_vol, net_buy = 0, 0
        
        try:
            res_vol = requests.get(f"https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date={month_str}&stockNo={raw_ticker}", timeout=2).json()
            if res_vol.get('stat') == 'OK':
                df_vol = pd.DataFrame(res_vol['data'], columns=res_vol['fields'])
                vol_val = dict(zip(df_vol['日期'], df_vol['成交股數'])).get(tw_date_str, "0")
                daily_vol = round((int(vol_val.replace(',', '')) if isinstance(vol_val, str) else int(vol_val)) / 1000)
        except: pass
        
        try:
            res_t86 = requests.get(f"https://www.twse.com.tw/rwd/zh/fund/T86?date={date_str}&selectType=ALL", timeout=2).json()
            if res_t86.get('stat') == 'OK':
                df_t86 = pd.DataFrame(res_t86['data'], columns=res_t86['fields'])
                target_row = df_t86[df_t86['證券代號'] == raw_ticker]
                if not target_row.empty:
                    buy_col = '三大法人買賣超股數' if '三大法人買賣超股數' in df_t86.columns else df_t86.columns[-1]
                    net_buy = round(int(target_row[buy_col].values[0].replace(',', '')) / 1000)
        except: pass
        
        if daily_vol == 0: daily_vol = int(hist_64.loc[d, 'Volume']) 
        daily_records.append({'date_disp': d.strftime('%Y-%m-%d'), 'net_buy': net_buy, 'volume': daily_vol})
        
    return daily_records


# ==========================================
# 副程式 6：分析法人強度
# ==========================================
def step6_analyze_inst_strength(daily_records):
    """負責計算 5 日滾動買賣強度並產出資料表與 9 日圖表"""
    chip_results = []
    for i in range(4, 9):
        window = daily_records[i-4 : i+1] 
        t_net_buy = sum(w['net_buy'] for w in window)
        t_vol = sum(w['volume'] for w in window)
        ratio = (t_net_buy / t_vol * 100) if t_vol > 0 else 0
        chip_results.append({
            '結算日期': daily_records[i]['date_disp'],
            '當日買賣超 (張)': f"{daily_records[i]['net_buy']:,}",
            '5日累計買賣超 (張)': f"{t_net_buy:,}",
            '5日累計成交量 (張)': f"{t_vol:,}",
            '買賣強度 (%)': f"{ratio:.2f}%"
        })
    df_chip = pd.DataFrame(list(reversed(chip_results)))
    
    # 建立 9 日直條圖物件
    df_9_days = pd.DataFrame(daily_records)
    inst_colors_9 = ['#FF4B4B' if val > 0 else '#00B050' for val in df_9_days['net_buy']]
    fig_inst = go.Figure(data=[go.Bar(x=df_9_days['date_disp'], y=df_9_days['net_buy'], marker_color=inst_colors_9, text=df_9_days['net_buy'], textposition='auto')])
    fig_inst.update_layout(xaxis=dict(type='category', title="交易日期"), yaxis=dict(title="買賣超 (張)"), margin=dict(l=0, r=0, t=30, b=0), height=300)
    
    return df_chip, fig_inst


# ==========================================
# 【新增】副程式 7：從證交所抓取月的融資融券資料並分析
# ==========================================
@st.cache_data(ttl=300, show_spinner=False)
def step7_fetch_and_analyze_margin(raw_ticker, hist_64, is_otc):
    """負責從證交所抓取月信用交易統計，並精算近 64 日融資券數據與每日變動量"""
    if is_otc:
        return pd.DataFrame()
        
    # 提取近 64 日橫跨的所有不重複月份 (格式: YYYYMM01)
    unique_months = list(set([d.strftime('%Y%m01') for d in hist_64.index]))
    raw_margin_data = {}
    
    for m in unique_months:
        try:
            url = f"https://www.twse.com.tw/rwd/zh/marginTrading/CREDIT_ALI?date={m}&stockNo={raw_ticker}"
            res = requests.get(url, timeout=3).json()
            if res.get('stat') == 'OK':
                for row in res['data']:
                    # row[0]:日期(民國), row[4]:融資前日餘額, row[5]:融資今日餘額, row[10]:融券前日餘額, row[11]:融券今日餘額
                    tw_date = row[0].split('/')
                    ad_date_str = f"{int(tw_date[0])+1911}-{tw_date[1]}-{tw_date[2]}"
                    
                    margin_today = int(str(row[5]).replace(',', ''))
                    margin_yesterday = int(str(row[4]).replace(',', ''))
                    short_today = int(str(row[11]).replace(',', ''))
                    short_yesterday = int(str(row[10]).replace(',', ''))
                    
                    raw_margin_data[ad_date_str] = {
                        '融資餘額(張)': margin_today,
                        '融資變動(張)': margin_today - margin_yesterday,
                        '融券餘額(張)': short_today,
                        '融券變動(張)': short_today - short_yesterday
                    }
        except: pass

    # 將計算好的每日信用交易數據，精準對齊回近 64 個交易日的日期結構中
    margin_records = []
    for d in hist_64.index:
        d_str = d.strftime('%Y-%m-%d')
        day_data = raw_margin_data.get(d_str, {'融資餘額(張)': 0, '融資變動(張)': 0, '融券餘額(張)': 0, '融券變動(張)': 0})
        day_data['日期'] = d_str
        margin_records.append(day_data)
        
    df_margin_64 = pd.DataFrame(margin_records).set_index('日期')
    return df_margin_64


# ==========================================
# 介面繪製輔助函數 (Tech Chart)
# ==========================================
def render_tech_chart(hist_64, show_ma5, show_ma10, show_ma20, allow_zoom):
    """輔助副程式：負責產生包含 K線、KD、MACD 的 3層子圖表"""
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
# 🚀 系統主架構 (指揮中心中樞)
# ==========================================
name_to_ticker, list_loaded = load_stock_list()
if not list_loaded: st.warning("⚠️ 找不到 'TW50100.xlsx'，請直接輸入股票代號。")

user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")
if st.button("🚀 開始分析", use_container_width=True):
    st.session_state.analyzed_input = user_input

if st.session_state.analyzed_input:
    current_target = st.session_state.analyzed_input
    
    # [準備工作] 解析輸入代號
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

    # ----------------------------------------------------
    # 前半段：極速運算區 (Step 1 ~ Step 4，秒速渲染)
    # ----------------------------------------------------
    with st.spinner('📡 正在運算核心技術指標與分價量...'):
        # ▶ 副程式 1：抓取 YFinance
        hist, yf_ticker = step1_fetch_yf_data(yf_ticker, raw_ticker, auto_fallback)
        if hist.empty:
            st.error("❌ 無法取得歷史資料。請確認代號。")
            st.stop()
            
        # ▶ 副程式 2：計算技術指標
        hist_64 = step2_calc_tech_indicators(hist)
        latest = hist_64.iloc[-1]
        
        # ▶ 副程式 3：計算分價量與圖表
        bins_data, all_intervals_disp, fig_vol, current_price_round = step3_process_volume_profile(hist_64)
        
        # ▶ 副程式 4：尋找支撐與壓力
        top_5_above, top_5_below = step4_find_support_resistance(bins_data, current_price_round)

    # ---------------- UI 顯示區 (前半部技術面) ----------------
    st.success(f"✅ {target_name} ({yf_ticker}) 價格與技術面指標載入完成！最新股價: {current_price_round:.2f}")

    # 顯示 2 的結果 (技術指標表)
    st.subheader("📊 技術指標參考")
    st.table(pd.DataFrame({
        "項目": ["均線狀況", "KD狀況", "MACD狀況"],
        "狀態": ["✅ 多頭" if latest['MA5'] > latest['MA10'] > latest['MA20'] else ("⚠️ 空頭" if latest['MA5'] < latest['MA10'] < latest['MA20'] else "⭕ 盤整"), 
                 "✅ 多" if latest['K'] > latest['D'] else "⚠️ 空", 
                 "✅ 多" if latest['DIF'] > latest['MACD'] else "⚠️ 空"],
        "數值細項": [f"5MA:{latest['MA5']:.1f} / 10MA:{latest['MA10']:.1f} / 20MA:{latest['MA20']:.1f}", f"K:{latest['K']:.1f} / D:{latest['D']:.1f}", f"DIF:{latest['DIF']:.1f} / MACD:{latest['MACD']:.1f}"]
    }))

    # 顯示 2 的結果 (K線綜合儀表板)
    allow_zoom = st.checkbox("🔍 啟用圖表縮放與拖曳", value=False)
    with st.container(border=True):
        st.subheader("📈 技術分析綜合儀表板")
        c1, c2, c3 = st.columns(3)
        fig_tech = render_tech_chart(hist_64, c1.checkbox("顯示 5MA", value=False), c2.checkbox("顯示 10MA", value=True), c3.checkbox("顯示 20MA", value=False), allow_zoom)
        st.plotly_chart(fig_tech, use_container_width=True)

    # 顯示 3 的結果 (分價量圖)
    st.subheader("📊 64日分價量參考圖")
    fig_vol.update_xaxes(fixedrange=not allow_zoom)
    fig_vol.update_yaxes(fixedrange=not allow_zoom)
    st.plotly_chart(fig_vol, use_container_width=True)

    # 顯示 4 的結果 (支撐壓力)
    st.subheader("🎯 關鍵支撐與壓力 (Top 5)")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**⬆️ 向上方壓力區**")
        for item in top_5_above: st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")
    with col2:
        st.write("**⬇️ 向下方支撐區**")
        for item in top_5_below: st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")


    # ---------------- 核心籌碼區 (Step 5 ~ Step 7，背景緩慢執行) ----------------
    is_otc = yf_ticker.endswith('.TWO')
    df_chip = pd.DataFrame()
    df_margin = pd.DataFrame()
    
    if is_otc:
        st.subheader("📈 近 5 日法人買賣強度")
        st.info("⚠️ 該股為上櫃股票，目前僅支援上市股票之法人與融資券籌碼統計。")
    else:
        # ▶ 接續在主程式背景繼續執行 5 ~ 6
        st.subheader("📈 近 5 日法人買賣強度")
        with st.spinner("⏳ 正在向證交所 API 抓取近 5 日法人籌碼數據..."):
            # 副程式 5：抓取法人
            daily_records = step5_fetch_twse_data(raw_ticker, hist_64, is_otc)
            # 副程式 6：分析強度與產出 9 日圖
            df_chip, fig_inst = step6_analyze_inst_strength(daily_records)
            
        if not df_chip.empty:
            st.dataframe(df_chip, use_container_width=True)
            st.subheader("📊 近 9 日法人買賣超直條圖")
            fig_inst.update_xaxes(fixedrange=not allow_zoom)
            fig_inst.update_yaxes(fixedrange=not allow_zoom)
            st.plotly_chart(fig_inst, use_container_width=True)

        # ▶ 【全新接續】執行副程式 7
        st.subheader("👥 信用交易籌碼分析 (近 64 日融資融券全覽)")
        with st.spinner("⏳ 正在向證交所快取月度信用交易統計，並計算餘額變動量..."):
            # 副程式 7：抓取整月並過濾分析 64 天的變動
            df_margin = step7_fetch_and_analyze_margin(raw_ticker, hist_64, is_otc)
            
        if not df_margin.empty:
            # 呈現最新一天的融資券結果與變動
            latest_margin = df_margin.iloc[-1]
            st.markdown(
                f"最新交易日信用交易結算：\n"
                f"* 融資餘額：**{latest_margin['融資餘額(張)']:,}** 張（當日變動：**{latest_margin['融資變動(張)']_:+:,}** 張）\n"
                f"* 融券餘額：**{latest_margin['融券餘額(張)']:,}** 張（當日變動：**{latest_margin['融券變動(張)']_:+:,}** 張）"
            )
            # 顯示近 10 天的詳細明細表供對照
            with st.expander("查看近 10 日融資融券異動明細表"):
                st.dataframe(df_margin.tail(10).sort_index(ascending=False), use_container_width=True)


    # ----------------------------------------------------
    # 結尾：Excel 報表匯出 (整合 1~7 所有運算結果)
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
            if not df_chip.empty: df_chip.to_excel(writer, sheet_name='法人買賣強度', index=False)
            if not df_margin.empty: df_margin.reset_index().to_excel(writer, sheet_name='融資融券追蹤', index=False)
            
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
