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
# 0. 初始化 Session State (記憶體)
# ==========================================
if 'analyzed_input' not in st.session_state:
    st.session_state.analyzed_input = None

# ==========================================
# 1. 網頁基本設定 
# ==========================================
st.set_page_config(page_title="台股籌碼分析工具", page_icon="📈", layout="centered")
st.title("📊 台股籌碼與技術指標綜合分析")
st.markdown("支援 **技術K線均線**、**KD/MACD**、**融資券追蹤**、**分價量防守** 與 **五日法人買賣強度**")

# ==========================================
# 2. 爬蟲與資料快取函數
# ==========================================
@st.cache_data
def load_stock_list():
    file_path = 'TW50100.xlsx'
    name_to_ticker = {}
    try:
        df_excel = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_ticker, col_name = df_excel.columns[0], df_excel.columns[1]
        for _, row in df_excel.iterrows():
            if pd.notna(row[col_name]) and pd.notna(row[col_ticker]):
                name = str(row[col_name]).strip()
                ticker = str(row[col_ticker]).strip()
                if ticker.endswith('.0'): ticker = ticker[:-2]
                name_to_ticker[name] = ticker
        return name_to_ticker, True
    except Exception:
        return {}, False

name_to_ticker, list_loaded = load_stock_list()

if not list_loaded:
    st.warning("⚠️ 找不到 'TW50100.xlsx'，請直接輸入股票代號 (例如: 2330)。")

@st.cache_data(ttl=300, show_spinner=False)
def fetch_stock_history(ticker):
    try:
        stock_data = yf.Ticker(ticker)
        return stock_data.history(period="6mo")
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def get_margin_data(date_str, ticker):
    """抓取整月融資券資料"""
    try:
        url = f"https://www.twse.com.tw/rwd/zh/marginTrading/CREDIT_ALI?date={date_str}&stockNo={ticker}"
        res = requests.get(url, timeout=3).json()
        data = {}
        if res.get('stat') == 'OK':
            for row in res['data']:
                # row[0] 是日期(113/05/15), row[5] 是融資餘額, row[11] 是融券餘額
                tw_date = row[0].split('/')
                ad_date = f"{int(tw_date[0])+1911}{tw_date[1]}{tw_date[2]}"
                data[ad_date] = {
                    'Margin': int(str(row[5]).replace(',', '')), 
                    'Short': int(str(row[11]).replace(',', ''))
                }
        return data
    except Exception: 
        return {}

# ==========================================
# 3. 網頁 UI：使用者輸入區
# ==========================================
user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")
analyze_button = st.button("🚀 開始分析", use_container_width=True)

if analyze_button and user_input:
    st.session_state.analyzed_input = user_input

if st.session_state.analyzed_input:
    current_target_input = st.session_state.analyzed_input
    
    with st.spinner('📡 正在向證交所與雲端伺服器抓取大數據，請稍候...'):
        
        # --- 解析股票代號 ---
        target_name, yf_ticker, raw_ticker, auto_fallback = "", "", "", True
        matched_names = [name for name in name_to_ticker.keys() if current_target_input in name] if list_loaded else []
        
        if len(matched_names) == 0:
            target_name = f"自訂代號 ({current_target_input})"
            if current_target_input.lower().endswith('.tw') or current_target_input.lower().endswith('.two'):
                yf_ticker = current_target_input.upper()
                auto_fallback = False 
                raw_ticker = current_target_input.split('.')[0]
            else:
                raw_ticker = current_target_input
                yf_ticker = f"{raw_ticker}.TW"
        elif len(matched_names) > 1:
            st.error(f"⚠️ 找到多檔包含 '{current_target_input}' 的股票，請輸入更明確的名稱：{', '.join(matched_names)}")
            st.stop()
        else:
            target_name = matched_names[0]
            raw_ticker = name_to_ticker[target_name]
            yf_ticker = f"{raw_ticker}.TW"

        # --- 抓取股價與計算技術指標 ---
        try:
            hist = fetch_stock_history(yf_ticker)
            if hist.empty and auto_fallback and raw_ticker:
                yf_ticker_two = f"{raw_ticker}.TWO"
                hist_two = fetch_stock_history(yf_ticker_two)
                if not hist_two.empty:
                    yf_ticker = yf_ticker_two
                    hist = hist_two
            
            if hist.empty:
                st.error("❌ 無法取得歷史資料。可能是股票下市、代號錯誤，或 Yahoo API 暫時阻擋，請 5 分鐘後再試。")
                st.stop()
            
            # MA 均線
            hist['MA5'] = hist['Close'].rolling(window=5).mean()
            hist['MA10'] = hist['Close'].rolling(window=10).mean()
            hist['MA20'] = hist['Close'].rolling(window=20).mean()

            # KD 指標
            low_min = hist['Low'].rolling(window=9).min()
            high_max = hist['High'].rolling(window=9).max()
            rsv = (hist['Close'] - low_min) / (high_max - low_min + 1e-9) * 100
            hist['K'] = rsv.ewm(com=2, adjust=False).mean()
            hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
            
            # MACD 指標
            ema12 = hist['Close'].ewm(span=12, adjust=False).mean()
            ema26 = hist['Close'].ewm(span=26, adjust=False).mean()
            hist['DIF'] = ema12 - ema26
            hist['MACD'] = hist['DIF'].ewm(span=9, adjust=False).mean()
            hist['OSC'] = (hist['DIF'] - hist['MACD']) * 2 
            
            hist_64 = hist.tail(64).copy()
            hist_64['Volume'] = hist_64['Volume'] / 1000  
            latest = hist_64.iloc[-1]
            
            # --- 抓取整月融資券 (規劃A) ---
            margin_map = {}
            if not yf_ticker.endswith('.TWO'): # 上市才抓信用交易
                unique_months = list(set([d.strftime('%Y%m01') for d in hist_64.index]))
                for m in unique_months:
                    margin_map.update(get_margin_data(m, raw_ticker))
            
            hist_64['Margin'] = [margin_map.get(d.strftime('%Y%m%d'), {}).get('Margin', 0) for d in hist_64.index]
            hist_64['Short'] = [margin_map.get(d.strftime('%Y%m%d'), {}).get('Short', 0) for d in hist_64.index]

            # --- 籌碼區間與分價量計算 ---
            current_price_round = round(latest['Close'], 2)
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
            top_5_above_disp = sorted(sorted([b for b in bins_data if b['mid'] >= current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
            top_5_below_disp = sorted(sorted([b for b in bins_data if b['mid'] < current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)

            # --- 抓取法人買賣超 (精簡版 9 天迴圈) ---
            last_9_dates = hist_64.index[-9:]
            daily_records = []
            
            if not yf_ticker.endswith('.TWO'):
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
            else:
                for d in last_9_dates:
                    daily_records.append({'date_disp': d.strftime('%Y-%m-%d'), 'net_buy': 0, 'volume': int(hist_64.loc[d, 'Volume'])})

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            st.stop()

    # ==========================================
    # 4. 網頁 UI：呈現結果
    # ==========================================
    st.success(f"✅ {target_name} ({yf_ticker}) 分析完成！最新股價: {current_price_round:.2f}")

    # --- 1. 技術指標總表 ---
    ma_trend = "✅ 多頭" if latest['MA5'] > latest['MA10'] > latest['MA20'] else ("⚠️ 空頭" if latest['MA5'] < latest['MA10'] < latest['MA20'] else "⭕ 盤整")
    kd_trend = "✅ 多" if latest['K'] > latest['D'] else "⚠️ 空"
    macd_trend = "✅ 多" if latest['DIF'] > latest['MACD'] else "⚠️ 空"
    
    st.subheader("📊 技術指標參考")
    st.table(pd.DataFrame({
        "項目": ["均線狀況", "KD狀況", "MACD狀況", "今日融資餘額", "今日融券餘額"],
        "狀態": [ma_trend, kd_trend, macd_trend, f"{latest['Margin']:,}", f"{latest['Short']:,}"],
        "數值細項": [f"5MA:{latest['MA5']:.1f} / 10MA:{latest['MA10']:.1f}", f"K:{latest['K']:.1f} / D:{latest['D']:.1f}", f"DIF:{latest['DIF']:.1f} / MACD:{latest['MACD']:.1f}", "張", "張"]
    }))

    # --- 2. 綜合四層圖表 ---
    allow_zoom = st.checkbox("🔍 啟用圖表縮放與拖曳功能 (防手機誤觸)", value=False)
    with st.container(border=True):
        st.subheader("📈 技術分析與籌碼指標綜合儀表板")
        col_ma1, col_ma2, col_ma3 = st.columns(3)
        show_ma5 = col_ma1.checkbox("顯示 5MA", value=False)
        show_ma10 = col_ma2.checkbox("顯示 10MA", value=True)
        show_ma20 = col_ma3.checkbox("顯示 20MA", value=False)
        
        date_strings = hist_64.index.strftime('%Y-%m-%d')
        fig_k = make_subplots(
            rows=4, cols=1, shared_xaxes=True, vertical_spacing=0.03,
            row_heights=[0.45, 0.15, 0.20, 0.20],
            subplot_titles=("價格與均線", "KD (9,3,3)", "MACD (12,26,9)", "融資餘額(紫) / 融券餘額(綠)")
        )
        
        # R1: K線
        fig_k.add_trace(go.Candlestick(x=date_strings, open=hist_64['Open'], high=hist_64['High'], low=hist_64['Low'], close=hist_64['Close'], name='K線', increasing_line_color='#FF4B4B', increasing_fillcolor='#FF4B4B', decreasing_line_color='#00B050', decreasing_fillcolor='#00B050'), row=1, col=1)
        if show_ma5: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA5'], name='5MA', line=dict(color='#7A431D', width=1.5)), row=1, col=1)
        if show_ma10: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA10'], name='10MA', line=dict(color='#00E5FF', width=1.5)), row=1, col=1)
        if show_ma20: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA20'], name='20MA', line=dict(color='#0D47A1', width=1.5)), row=1, col=1)
        
        # R2: KD
        fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['K'], name='K值', line=dict(color='#FF9900', width=1.2)), row=2, col=1)
        fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['D'], name='D值', line=dict(color='#0066FF', width=1.2)), row=2, col=1)
        
        # R3: MACD
        macd_colors = ['#FF4B4B' if val > 0 else '#00B050' for val in hist_64['OSC']]
        fig_k.add_trace(go.Bar(x=date_strings, y=hist_64['OSC'], name='OSC', marker_color=macd_colors), row=3, col=1)
        fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['DIF'], name='DIF', line=dict(color='#FF9900', width=1.2)), row=3, col=1)
        fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MACD'], name='MACD', line=dict(color='#0066FF', width=1.2)), row=3, col=1)
        
        # R4: 融資券餘額 (面積圖與折線圖)
        fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['Margin'], name='融資餘額', line=dict(color='#8B5CF6', width=2), fill='tozeroy', fillcolor='rgba(139, 92, 246, 0.2)'), row=4, col=1)
        fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['Short'], name='融券餘額', line=dict(color='#059669', width=2)), row=4, col=1)
        
        fig_k.update_layout(
            xaxis=dict(type='category', visible=False), xaxis2=dict(type='category', visible=False), xaxis3=dict(type='category', visible=False), xaxis4=dict(type='category', visible=True, title="交易日期", nticks=10),
            yaxis=dict(visible=False), yaxis2=dict(visible=True), yaxis3=dict(visible=True), yaxis4=dict(visible=True),
            xaxis_rangeslider_visible=False, margin=dict(l=4, r=4, t=30, b=4), height=800, hovermode='x unified', showlegend=False
        )
        fig_k.update_xaxes(fixedrange=not allow_zoom)
        fig_k.update_yaxes(fixedrange=not allow_zoom)
        st.plotly_chart(fig_k, use_container_width=True)

    # --- 3. 64 日分價量圖 ---
    st.subheader("📊 64日分價量參考圖")
    df_plot = pd.DataFrame({
        '價格區間': [item['label'] for item in all_intervals_disp],
        '累積成交量 (張)': [int(item['vol']) for item in all_intervals_disp],
        '標記': ['現價所在' if item['is_current'] else '一般區間' for item in all_intervals_disp]
    })
    fig = px.bar(df_plot, x='累積成交量 (張)', y='價格區間', color='標記', color_discrete_map={'現價所在': '#FF4B4B', '一般區間': '#60B4FF'}, orientation='h')
    fig.update_yaxes(categoryorder='array', categoryarray=df_plot['價格區間'])
    fig.update_layout(yaxis=dict(title="價格區間", autorange="reversed"), margin=dict(l=0, r=0, t=30, b=0), height=500)
    fig.update_xaxes(fixedrange=not allow_zoom)
    fig.update_yaxes(fixedrange=not allow_zoom)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("🎯 關鍵支撐與壓力 (Top 5)")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**⬆️ 向上方壓力區**")
        for item in top_5_above_disp: st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")
    with col2:
        st.write("**⬇️ 向下方支撐區**")
        for item in top_5_below_disp: st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")

    # --- 4. 近 5 日法人買賣強度 ---
    st.subheader("📈 近 5 日法人買賣強度")
    if yf_ticker.endswith('.TWO'):
        st.info("⚠️ 該股為上櫃股票，目前僅支援上市股票之籌碼查詢。")
        df_chip = pd.DataFrame()
    else:
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
        st.dataframe(df_chip, use_container_width=True)

    # --- 5. Excel 匯出 ---
    st.divider()
    st.subheader("💾 匯出完整 Excel 報表")
    
    df_sr_excel = pd.DataFrame([{'項次': i+1, '價格級距區間 (TWD)': item['disp_label'], '累積成交量 (張)': int(item['vol'])} for i, item in enumerate(all_intervals_disp)])
    df_top5_excel = pd.DataFrame(
        [{'位置': '⬆️ 向上壓力區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_above_disp] +
        [{'位置': '🎯 最新股價 (中軸)', '價格級距區間': f"{current_price_round:.2f}", '累積成交量 (張)': 0}] +
        [{'位置': '⬇️ 向下支撐區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_below_disp]
    )

    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_sr_excel.to_excel(writer, sheet_name='區間分價量總表', index=False)
            df_top5_excel.to_excel(writer, sheet_name='關鍵支撐壓力', index=False)
            if not df_chip.empty: df_chip.to_excel(writer, sheet_name='法人買賣強度', index=False)
            
            workbook = writer.book
            for sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                for col in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 25.5
            
            sheet1 = workbook['區間分價量總表']
            chart = BarChart()
            chart.type, chart.style = "bar", 10
            chart.title = f"{target_name} 64日分價量分佈圖"
            chart.x_axis.title, chart.y_axis.title = "價格級距區間 (TWD)", "累積成交量 (張)" 
            chart.x_axis.scaling.orientation, chart.y_axis.crosses = "maxMin", "max"
            chart.height, chart.width = max(10, len(df_sr_excel) * 0.5) * 1.5, 16 * 1.5
            
            max_r = len(df_sr_excel) + 1
            chart.add_data(Reference(sheet1, min_col=3, min_row=1, max_row=max_r), titles_from_data=True)
            chart.set_categories(Reference(sheet1, min_col=2, min_row=2, max_row=max_r))
            sheet1.add_chart(chart, "E2")
        
        safe_target_name = re.sub(r'[\\/*?:"<>|]', '_', target_name)
        today_str = datetime.now().strftime("%Y%m%d")
        
        st.download_button(
            label="📥 點我下載 Excel 分析報表",
            data=output.getvalue(),
            file_name=f"{safe_target_name}_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    except Exception as e:
        st.error(f"❌ Excel 產生發生錯誤：{e}")
