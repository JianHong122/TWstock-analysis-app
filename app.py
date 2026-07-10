import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import io
import re
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

st.set_page_config(page_title="台股籌碼分析工具", page_icon="📈", layout="centered")

@st.cache_data
def load_stock_list():
    try:
        df = pd.read_excel('TW50100.xlsx', engine='openpyxl', dtype=str)
        return {str(row[df.columns[1]]): str(row[df.columns[0]]).replace('.0', '') for _, row in df.iterrows()}, True
    except: return {}, False


# ==========================================
# 副程式 1：抓取 YFinance 資料 (加入日期區間邏輯)
# ==========================================
@st.cache_data(ttl=300, show_spinner=False)
def step1_fetch_yf_data(ticker, raw_ticker, auto_fallback, target_date_str):
    """副程式 1：負責從 Yahoo Finance 抓取歷史股價資料，並支援指定結束日期"""
    
    # 處理日期：利用 pandas 處理日期加減
    # YFinance 的 end 是 exclusive (不包含)，所以我們要把目標日 + 1天
    end_dt = pd.to_datetime(target_date_str, format='%Y/%m/%d') + pd.Timedelta(days=1)
    # 往前抓 6 個月，確保資料足以運算 64天分價量與 MA/MACD 等技術指標
    start_dt = end_dt - pd.DateOffset(months=6) 
    
    start_str = start_dt.strftime('%Y-%m-%d')
    end_str = end_dt.strftime('%Y-%m-%d')

    hist = yf.Ticker(ticker).history(start=start_str, end=end_str)
    
    # 處理上櫃股票 (.TWO) 的自動退退切換
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
    """副程式 2：負責所有技術指標的數學運算"""
    df = hist.copy()
    
    # MA 均線
    df['MA5'] = df['Close'].rolling(window=5).mean()
    df['MA10'] = df['Close'].rolling(window=10).mean()
    df['MA20'] = df['Close'].rolling(window=20).mean()

    # KD 指標 (9, 3, 3)
    low_min = df['Low'].rolling(window=9).min()
    high_max = df['High'].rolling(window=9).max()
    rsv = (df['Close'] - low_min) / (high_max - low_min + 1e-9) * 100
    df['K'] = rsv.ewm(com=2, adjust=False).mean()
    df['D'] = df['K'].ewm(com=2, adjust=False).mean()
    
    # MACD 指標 (12, 26, 9)
    ema12 = df['Close'].ewm(span=12, adjust=False).mean()
    ema26 = df['Close'].ewm(span=26, adjust=False).mean()
    df['DIF'] = ema12 - ema26
    df['MACD'] = df['DIF'].ewm(span=9, adjust=False).mean()
    df['OSC'] = (df['DIF'] - df['MACD']) * 2 
    
    # 轉換 Volume 單位為「張」
    df['Volume'] = df['Volume'] / 1000  
    return df.tail(64)


# ==========================================
# 副程式 3：產生分價量統計與圖表
# ==========================================
def step3_process_volume_profile(hist_64):
    """副程式 3：負責 20 級距切分、籌碼攤平演算法，並產出 Plotly 圖表"""
    current_price_round = round(hist_64['Close'].dropna().iloc[-1], 2)
    max_price, min_price = hist_64['High'].max(), hist_64['Low'].min()
    if max_price == min_price:
        max_price, min_price = min_price * 1.05, min_price * 0.95
    
    bin_size = (max_price - min_price) / 20
    curr_bin_idx = 19 if current_price_round >= max_price else (0 if current_price_round <= min_price else min(19, int((current_price_round - min_price) / bin_size)))
    
    # 初始化 20 個級距
    bins_data = [{'idx': i, 'start': min_price + i * bin_size, 'end': min_price + (i + 1) * bin_size, 'mid': (min_price + i * bin_size + min_price + (i + 1) * bin_size) / 2, 'label': f"{min_price + i * bin_size:.2f} ~ {min_price + (i + 1) * bin_size:.2f}", 'disp_label': f"{'** ' if i == curr_bin_idx else ''}{min_price + i * bin_size:.2f} ~ {min_price + (i + 1) * bin_size:.2f}", 'is_current': (i == curr_bin_idx), 'vol': 0} for i in range(20)]
    
    # 5/30/65 籌碼攤平邏輯
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
            
    # 統計加總
    df_all_vols = pd.DataFrame(all_price_vols)
    for price, vol in df_all_vols.groupby('Price')['Vol'].sum().items():
        if price >= max_price: bins_data[-1]['vol'] += vol
        elif price <= min_price: bins_data[0]['vol'] += vol
        else: bins_data[min(19, int((price - min_price) / bin_size))]['vol'] += vol
            
    all_intervals_disp = sorted(bins_data, key=lambda x: x['idx'], reverse=True)
    
    # 產生 Plotly 圖表物件
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
    """副程式 4：負責過濾出上下 Top 5 的支撐與壓力區"""
    top_5_above = sorted(sorted([b for b in bins_data if b['mid'] >= current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
    top_5_below = sorted(sorted([b for b in bins_data if b['mid'] < current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
    return top_5_above, top_5_below


# ==========================================
# 副程式 5：從證交所抓取法人資料
# ==========================================
def step5_fetch_twse_data(raw_ticker, hist_64, is_otc):
    """副程式 5：負責爬取近 9 日的三大法人買賣超資料"""
    last_9_dates = hist_64.index[-9:]
    daily_records = []
    
    if is_otc:
        # 上櫃股票不支援
        for d in last_9_dates:
            daily_records.append({'date_disp': d.strftime('%Y-%m-%d'), 'net_buy': 0, 'volume': int(hist_64.loc[d, 'Volume'])})
        return daily_records
        
    for d in last_9_dates:
        month_str, date_str = d.strftime('%Y%m01'), d.strftime('%Y%m%d')
        tw_date_str = f"{d.year - 1911}/{d.strftime('%m/%d')}"
        daily_vol, net_buy = 0, 0
        
        try: # 抓成交量
            res_vol = requests.get(f"https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date={month_str}&stockNo={raw_ticker}", timeout=2).json()
            if res_vol.get('stat') == 'OK':
                df_vol = pd.DataFrame(res_vol['data'], columns=res_vol['fields'])
                vol_val = dict(zip(df_vol['日期'], df_vol['成交股數'])).get(tw_date_str, "0")
                daily_vol = round((int(vol_val.replace(',', '')) if isinstance(vol_val, str) else int(vol_val)) / 1000)
        except: pass
        
        try: # 抓三大法人
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
    """副程式 6：負責計算 5 日滾動買賣強度並產出資料表"""
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
    return df_chip


# ==========================================
# 介面繪製輔助函數 (Tech Chart)
# ==========================================
def render_tech_chart(hist_64, show_ma5, show_ma10, show_ma20, allow_zoom):
    """輔助副程式：繪製 K線與指標 3層子圖表"""
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
st.markdown("支援 **技術K線均線**、**KD/MACD**、**分價量防守** 與 **五日法人買賣強度**")

name_to_ticker, list_loaded = load_stock_list()
if not list_loaded: st.warning("⚠️ 找不到 'TW50100.xlsx'，請直接輸入股票代號。")

user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")

# --- 新增的日期輸入與防呆處理區塊 ---
default_date = datetime.now().strftime("%Y/%m/%d")
target_date_input = st.text_input("📅 請輸入查詢基準日 (西元年/月/日)：", value=default_date, placeholder="例如: 2024/01/01")

if st.button("🚀 開始分析", use_container_width=True):
    input_date_str = target_date_input.strip()
    
    # 空白處理：如果使用者清空輸入框，自動代入預設日期(今天)
    if not input_date_str:
        input_date_str = default_date
        
    # 格式防呆：檢查是否符合 YYYY/MM/DD
    try:
        # 嘗試轉換日期，確認格式正確
        datetime.strptime(input_date_str, "%Y/%m/%d")
        
        # 驗證成功，存入 session_state
        st.session_state.analyzed_input = user_input
        st.session_state.target_date = input_date_str
    except ValueError:
        st.error("⚠️ 日期格式錯誤！請輸入正確的「西元年/月/日」格式，例如：2024/01/01")
        st.stop() # 停止後續運算

# ------------------------------------

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
    # 以下為嚴格遵守「模組化 1~6 順序」的執行流程
    # ----------------------------------------------------
    
    with st.spinner('📡 正在運算核心技術指標與分價量...'):
        
        # ▶ 執行 1：抓取 YFinance (將基準日期傳入)
        hist, yf_ticker = step1_fetch_yf_data(yf_ticker, raw_ticker, auto_fallback, st.session_state.target_date)
        if hist.empty:
            st.error("❌ 無法取得該日期之前的歷史資料。請確認代號與日期。")
            st.stop()
            
        # ▶ 執行 2：計算技術指標
        hist_64 = step2_calc_tech_indicators(hist)
        latest = hist_64.iloc[-1]
        
        # ▶ 執行 3：計算分價量與圖表
        bins_data, all_intervals_disp, fig_vol, current_price_round = step3_process_volume_profile(hist_64)
        
        # ▶ 執行 4：尋找支撐與壓力
        top_5_above, top_5_below = step4_find_support_resistance(bins_data, current_price_round)

    # ---------------- UI 顯示區 (前半部) ----------------
    # 抓出 YFinance 實際取得的「最新一日」日期，展現"自動尋找最近交易日"的效果
    actual_last_date = hist_64.index[-1].strftime('%Y/%m/%d')
    st.success(f"✅ {target_name} ({yf_ticker}) 分析完成！實際查詢基準日: **{actual_last_date}** / 股價: **{current_price_round:.2f}**")

    # 顯示 2 的結果 (技術指標表)
    st.subheader("📊 技術指標參考")
    st.table(pd.DataFrame({
        "項目": ["均線狀況", "KD狀況", "MACD狀況"],
        "狀態": ["✅ 多頭" if latest['MA5'] > latest['MA10'] > latest['MA20'] else ("⚠️ 空頭" if latest['MA5'] < latest['MA10'] < latest['MA20'] else "⭕ 盤整"), 
                 "✅ 多" if latest['K'] > latest['D'] else "⚠️ 空", 
                 "✅ 多" if latest['DIF'] > latest['MACD'] else "⚠️ 空"],
        "數值細項": [f"5MA:{latest['MA5']:.1f} / 10MA:{latest['MA10']:.1f}", f"K:{latest['K']:.1f} / D:{latest['D']:.1f}", f"DIF:{latest['DIF']:.1f} / MACD:{latest['MACD']:.1f}"]
    }))

    # 顯示 2 的結果 (K線圖)
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

    # ----------------------------------------------------
    # 背景繼續執行 5~6
    # ----------------------------------------------------
    st.subheader("📈 近 5 日法人買賣強度")
    is_otc = yf_ticker.endswith('.TWO')
    
    if is_otc:
        st.info("⚠️ 該股為上櫃股票，目前僅支援上市股票之法人籌碼查詢。")
        df_chip = pd.DataFrame()
    else:
        with st.spinner("⏳ 正在向證交所 API 抓取近 5 日法人籌碼數據，請稍候..."):
            # ▶ 執行 5：抓取證交所資料
            daily_records = step5_fetch_twse_data(raw_ticker, hist_64, is_otc)
            
            # ▶ 執行 6：分析法人強度
            df_chip = step6_analyze_inst_strength(daily_records)
            
        # 顯示 6 的結果
        if not df_chip.empty:
            st.dataframe(df_chip, use_container_width=True)

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
            if not df_chip.empty: df_chip.to_excel(writer, sheet_name='法人買賣強度', index=False)
            
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
