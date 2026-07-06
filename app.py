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
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ==========================================
# 0. 初始化 Session State (記憶體)，解決打勾重新整理的問題
# ==========================================
if 'analyzed_input' not in st.session_state:
    st.session_state.analyzed_input = None

# ==========================================
# 1. 網頁基本設定 (支援手機版 RWD)
# ==========================================
st.set_page_config(page_title="台股籌碼分析工具", page_icon="📈", layout="centered")
st.title("📊 台股區間支撐壓力與法人籌碼分析")
st.markdown("支援 **技術K線均線**、**20級距全覽**、**Top 5 關鍵防守** 與 **五日法人買賣強度**")

# ==========================================
# 2. 讀取股票清單與抓取股價 (加入 Cache 防止被 Yahoo 封鎖)
# ==========================================
@st.cache_data
def load_stock_list():
    file_path = 'TW50100.xlsx'
    name_to_ticker = {}
    try:
        df_excel = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_ticker = df_excel.columns[0]
        col_name = df_excel.columns[1]
        for _, row in df_excel.iterrows():
            if pd.notna(row[col_name]) and pd.notna(row[col_ticker]):
                name = str(row[col_name]).strip()
                ticker = str(row[col_ticker]).strip()
                if ticker.endswith('.0'):
                    ticker = ticker[:-2]
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
        hist = stock_data.history(period="6mo")
        return hist
    except Exception:
        return pd.DataFrame()

# ==========================================
# 3. 網頁 UI：使用者輸入區
# ==========================================
user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")
analyze_button = st.button("🚀 開始分析", use_container_width=True)

# 當按鈕被按下時，將輸入值存入記憶體中
if analyze_button and user_input:
    st.session_state.analyzed_input = user_input

# 只要記憶體中有值，就顯示分析結果 (這樣打勾時就不會消失了)
if st.session_state.analyzed_input:
    current_target_input = st.session_state.analyzed_input
    
    with st.spinner('正在根據證交所 Tick 檔位精算盤中籌碼分佈，請稍候...'):
        
        # 判斷輸入邏輯
        target_name = ""
        yf_ticker = ""
        raw_ticker = ""
        auto_fallback = True
        
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

        # ==========================================
        # 4. 抓取股價與運算技術指標、20 級距
        # ==========================================
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
            
            # 在切分 64 天前，先利用半年資料精算 5MA、10MA、20MA
            hist['MA5'] = hist['Close'].rolling(window=5).mean()
            hist['MA10'] = hist['Close'].rolling(window=10).mean()
            hist['MA20'] = hist['Close'].rolling(window=20).mean()
            
            # 裁切出近 64 日資料
            hist_64 = hist.tail(64).copy()
            hist_64['Volume'] = hist_64['Volume'] / 1000  # 轉換為張數
            
            current_price = hist_64['Close'].dropna().iloc[-1]
            current_price_round = round(current_price, 2)
            
            max_price = hist_64['High'].max()
            min_price = hist_64['Low'].min()
            
            if max_price == min_price:
                max_price = min_price * 1.05
                min_price = min_price * 0.95
                
            bin_size = (max_price - min_price) / 20
            
            if current_price_round >= max_price:
                curr_bin_idx = 19
            elif current_price_round <= min_price:
                curr_bin_idx = 0
            else:
                curr_bin_idx = int((current_price_round - min_price) / bin_size)
                if curr_bin_idx > 19: curr_bin_idx = 19
            
            bins_data = []
            for i in range(20):
                start = min_price + i * bin_size
                end = min_price + (i + 1) * bin_size
                is_current = (i == curr_bin_idx)
                mark = "** " if is_current else ""
                bins_data.append({
                    'idx': i,
                    'start': start,
                    'end': end,
                    'mid': (start + end) / 2,
                    'label': f"{start:.2f} ~ {end:.2f}",
                    'disp_label': f"{mark}{start:.2f} ~ {end:.2f}",
                    'is_current': is_current,
                    'vol': 0
                })
            
            # 籌碼攤平邏輯
            all_price_vols = []
            for index, row in hist_64.iterrows():
                o = round(row['Open'], 2)
                h = round(row['High'], 2)
                l = round(row['Low'], 2)
                c = round(row['Close'], 2)
                v = row['Volume']
                
                if l > h: l, h = h, l 
                
                vol_open = v * 0.05
                vol_close = v * 0.30
                vol_dist_total = v * 0.65
                
                ticks = []
                curr = l
                while curr <= h:
                    ticks.append(curr)
                    if curr < 10: ts = 0.01
                    elif curr < 50: ts = 0.05
                    elif curr < 100: ts = 0.1
                    elif curr < 500: ts = 0.5
                    elif curr < 1000: ts = 1.0
                    else: ts = 5.0
                    
                    curr = round(curr + ts, 2)
                
                n_ticks = len(ticks)
                vol_per_tick = vol_dist_total / n_ticks if n_ticks > 0 else 0
                
                all_price_vols.append({'Price': o, 'Vol': vol_open})
                all_price_vols.append({'Price': c, 'Vol': vol_close})
                for t in ticks:
                    all_price_vols.append({'Price': t, 'Vol': vol_per_tick})
                    
            df_all_vols = pd.DataFrame(all_price_vols)
            price_vol = df_all_vols.groupby('Price')['Vol'].sum()
            
            for price, vol in price_vol.items():
                if price >= max_price: bins_data[-1]['vol'] += vol
                elif price <= min_price: bins_data[0]['vol'] += vol
                else:
                    idx = int((price - min_price) / bin_size)
                    if idx > 19: idx = 19
                    bins_data[idx]['vol'] += vol
                    
            all_intervals_disp = sorted(bins_data, key=lambda x: x['idx'], reverse=True)
            
            top_5_above_disp = sorted(sorted([b for b in bins_data if b['mid'] >= current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
            top_5_below_disp = sorted(sorted([b for b in bins_data if b['mid'] < current_price_round and b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)

            # ==========================================
            # 5. 網頁 UI：顯示結果與圖表
            # ==========================================
            st.success(f"✅ {target_name} ({yf_ticker}) 分析完成！最新股價: {current_price_round:.2f}")
            
            # --- 繪製 K 線與均線圖 ---
            st.subheader("📈 64日技術K線與移動平均線")
            col_ma1, col_ma2, col_ma3 = st.columns(3)
            # 【修改點】：只保留 10MA 為 True，其他為 False
            show_ma5 = col_ma1.checkbox("顯示 5MA (咖啡色)", value=False)
            show_ma10 = col_ma2.checkbox("顯示 10MA (亮藍綠)", value=True)
            show_ma20 = col_ma3.checkbox("顯示 20MA (深藍色)", value=False)
            
            date_strings = hist_64.index.strftime('%Y-%m-%d')
            fig_k = go.Figure()
            
            # 實心純色的 K 線設計
            fig_k.add_trace(go.Candlestick(
                x=date_strings,
                open=hist_64['Open'],
                high=hist_64['High'],
                low=hist_64['Low'],
                close=hist_64['Close'],
                name='K線',
                increasing_line_color='#FF4B4B',  # 上漲外框紅
                increasing_fillcolor='#FF4B4B',   # 上漲實體紅
                decreasing_line_color='#00B050',  # 下跌外框綠
                decreasing_fillcolor='#00B050'    # 下跌實體綠
            ))
            
            if show_ma5:
                fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA5'], name='5MA', line=dict(color='#7A431D', width=1.5))) 
            if show_ma10:
                fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA10'], name='10MA', line=dict(color='#00E5FF', width=1.5))) 
            if show_ma20:
                fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA20'], name='20MA', line=dict(color='#0D47A1', width=1.5))) 
            
            # 設定 xaxis type 為 'category' 來自動略過沒有交易的假日與週末
            fig_k.update_layout(
                xaxis=dict(
                    title="交易日期", 
                    type='category', 
                    tickmode='auto', 
                    nticks=10  
                ),
                yaxis=dict(title="價格 (TWD)"),
                xaxis_rangeslider_visible=False,  
                margin=dict(l=0, r=0, t=20, b=0),
                height=420,
                hovermode='x unified',             
                legend=dict(orientation="h", y=1.05, x=0, yanchor="bottom") 
            )
            st.plotly_chart(fig_k, use_container_width=True)

            # --- 繪製實體分價量分佈圖 ---
            st.subheader("📊 64日實體分價量分佈圖")
            df_plot = pd.DataFrame({
                '價格區間': [item['label'] for item in all_intervals_disp],
                '累積成交量 (張)': [int(item['vol']) for item in all_intervals_disp],
                '標記': ['現價所在' if item['is_current'] else '一般區間' for item in all_intervals_disp]
            })
            
            fig = px.bar(df_plot, x='累積成交量 (張)', y='價格區間', color='標記', 
                         color_discrete_map={'現價所在': '#FF4B4B', '一般區間': '#60B4FF'},
                         orientation='h')
            fig.update_yaxes(categoryorder='array', categoryarray=df_plot['價格區間'])
            
            # 補回 autorange="reversed"，讓高價位在上方，低價位在下方
            fig.update_layout(
                yaxis=dict(title="價格區間", autorange="reversed"),
                margin=dict(l=0, r=0, t=30, b=0), 
                height=500
            )
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("🎯 關鍵支撐與壓力 (Top 5)")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**⬆️ 向上方壓力區**")
                for item in top_5_above_disp:
                    st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")
            with col2:
                st.write("**⬇️ 向下方支撐區**")
                for item in top_5_below_disp:
                    st.write(f"`{item['disp_label']:<20}` | **{int(item['vol']):,}** 張")

            # ==========================================
            # 6. 法人買賣強度
            # ==========================================
            st.subheader("📈 近 5 日法人買賣強度")
            if yf_ticker.endswith('.TWO'):
                st.info("⚠️ 該股為上櫃股票，證交所 API 僅支援上市股票之法人籌碼查詢。")
                df_chip = pd.DataFrame()
            else:
                last_9_dates = hist_64.index[-9:]
                daily_records = []
                for d in last_9_dates:
                    month_str, date_str = d.strftime('%Y%m01'), d.strftime('%Y%m%d')
                    tw_date_str = f"{d.year - 1911}/{d.strftime('%m/%d')}"
                    
                    daily_vol, net_buy = 0, 0
                    
                    try:
                        res_vol = requests.get(f"https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date={month_str}&stockNo={raw_ticker}", timeout=3).json()
                        if res_vol.get('stat') == 'OK':
                            df_vol = pd.DataFrame(res_vol['data'], columns=res_vol['fields'])
                            vol_val = dict(zip(df_vol['日期'], df_vol['成交股數'])).get(tw_date_str, "0")
                            daily_vol_shares = int(vol_val.replace(',', '')) if isinstance(vol_val, str) else int(vol_val)
                            daily_vol = round(daily_vol_shares / 1000)
                    except: pass
                    
                    try:
                        res_t86 = requests.get(f"https://www.twse.com.tw/rwd/zh/fund/T86?date={date_str}&selectType=ALL", timeout=3).json()
                        if res_t86.get('stat') == 'OK':
                            df_t86 = pd.DataFrame(res_t86['data'], columns=res_t86['fields'])
                            target_row = df_t86[df_t86['證券代號'] == raw_ticker]
                            if not target_row.empty:
                                buy_col = '三大法人買賣超股數' if '三大法人買賣超股數' in df_t86.columns else df_t86.columns[-1]
                                net_buy_shares = int(target_row[buy_col].values[0].replace(',', ''))
                                net_buy = round(net_buy_shares / 1000)
                    except: pass
                    
                    if daily_vol == 0: 
                        daily_vol = int(hist_64.loc[d, 'Volume']) 
                        
                    daily_records.append({'date_disp': d.strftime('%Y-%m-%d'), 'net_buy': net_buy, 'volume': daily_vol})

                chip_results = []
                for i in range(4, 9):
                    window = daily_records[i-4 : i+1] 
                    t_net_buy, t_vol = sum(w['net_buy'] for w in window), sum(w['volume'] for w in window)
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

            # ==========================================
            # 7. Excel 記憶體內產生與下載按鈕
            # ==========================================
            st.divider()
            st.subheader("💾 匯出完整 Excel 報表")
            
            df_sr_excel = pd.DataFrame([{
                '項次': i+1, '價格級距區間 (TWD)': item['disp_label'], '累積成交量 (張)': int(item['vol'])
            } for i, item in enumerate(all_intervals_disp)])
            
            df_top5_excel = pd.DataFrame(
                [{'位置': '⬆️ 向上壓力區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_above_disp] +
                [{'位置': '🎯 最新股價 (中軸)', '價格級距區間': f"{current_price_round:.2f}", '累積成交量 (張)': 0}] +
                [{'位置': '⬇️ 向下支撐區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_below_disp]
            )

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_sr_excel.to_excel(writer, sheet_name='區間分價量總表', index=False)
                df_top5_excel.to_excel(writer, sheet_name='關鍵支撐壓力', index=False)
                if not df_chip.empty:
                    df_chip.to_excel(writer, sheet_name='法人買賣強度', index=False)
                
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
            st.error(f"❌ 發生錯誤：{e}")
