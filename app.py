import streamlit as st
import pandas as pd
import yfinance as yf
import requests
import math
import re
import io
from datetime import datetime
import plotly.express as px
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 網頁基本設定 (支援手機版 RWD)
# ==========================================
st.set_page_config(page_title="台股籌碼分析工具", page_icon="📈", layout="centered")
st.title("📊 台股區間支撐壓力與法人籌碼分析")
st.markdown("支援 **OHLC加權 20級距全覽**、**Top 5 關鍵防守** 與 **五日法人買賣強度**")

# ==========================================
# 2. 讀取股票清單 (使用快取加速載入)
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

# ==========================================
# 3. 網頁 UI：使用者輸入區
# ==========================================
user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")
analyze_button = st.button("🚀 開始分析", use_container_width=True)

# 點擊分析按鈕後執行
if analyze_button and user_input:
    with st.spinner('正在抓取大數據並進行運算，請稍候...'):
        
        # 判斷輸入邏輯
        target_name = ""
        yf_ticker = ""
        raw_ticker = ""
        auto_fallback = True
        
        matched_names = [name for name in name_to_ticker.keys() if user_input in name] if list_loaded else []
        
        if len(matched_names) == 0:
            target_name = f"自訂代號 ({user_input})"
            if user_input.lower().endswith('.tw') or user_input.lower().endswith('.two'):
                yf_ticker = user_input.upper()
                auto_fallback = False 
                raw_ticker = user_input.split('.')[0]
            else:
                raw_ticker = user_input
                yf_ticker = f"{raw_ticker}.TW"
        elif len(matched_names) > 1:
            st.error(f"⚠️ 找到多檔包含 '{user_input}' 的股票，請輸入更明確的名稱：{', '.join(matched_names)}")
            st.stop()
        else:
            target_name = matched_names[0]
            raw_ticker = name_to_ticker[target_name]
            yf_ticker = f"{raw_ticker}.TW"

        # ==========================================
        # 4. 抓取股價與運算 20 級距 (OHLC加權)
        # ==========================================
        try:
            stock_data = yf.Ticker(yf_ticker)
            hist = stock_data.history(period="6mo")
            
            if hist.empty and auto_fallback and raw_ticker:
                yf_ticker_two = f"{raw_ticker}.TWO"
                stock_data = yf.Ticker(yf_ticker_two)
                hist = stock_data.history(period="6mo")
                if not hist.empty:
                    yf_ticker = yf_ticker_two
            
            if hist.empty:
                st.error("❌ 無法取得歷史資料。請確認該股票是否已下市或代號錯誤。")
                st.stop()
            
            hist_64 = hist.tail(64).copy()
            
            # 【核心修改】：將成交量由「股數」轉換為「張數」(除以 1000)
            hist_64['Volume'] = hist_64['Volume'] / 1000
            
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
            
            # OHLC 加權分配
            hist_64['Vol_OpenClose'] = hist_64['Volume'] * 0.35
            hist_64['Vol_HighLow'] = hist_64['Volume'] * 0.15
            
            df_open = pd.DataFrame({'Price': hist_64['Open'].round(2), 'Vol': hist_64['Vol_OpenClose']})
            df_high = pd.DataFrame({'Price': hist_64['High'].round(2), 'Vol': hist_64['Vol_HighLow']})
            df_low = pd.DataFrame({'Price': hist_64['Low'].round(2), 'Vol': hist_64['Vol_HighLow']})
            df_close = pd.DataFrame({'Price': hist_64['Close'].round(2), 'Vol': hist_64['Vol_OpenClose']})
            
            price_vol = pd.concat([df_open, df_high, df_low, df_close]).groupby('Price')['Vol'].sum()
            
            for price, vol in price_vol.items():
                if price >= max_price: bins_data[-1]['vol'] += vol
                elif price <= min_price: bins_data[0]['vol'] += vol
                else:
                    idx = int((price - min_price) / bin_size)
                    if idx > 19: idx = 19
                    bins_data[idx]['vol'] += vol
                    
            all_intervals_disp = sorted(bins_data, key=lambda x: x['idx'], reverse=True)
            
            # Top 5 計算
            above_bins = [b for b in bins_data if b['mid'] >= current_price_round and b['vol'] > 0]
            below_bins = [b for b in bins_data if b['mid'] < current_price_round and b['vol'] > 0]
            
            top_5_above_disp = sorted(sorted(above_bins, key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)
            top_5_below_disp = sorted(sorted(below_bins, key=lambda x: x['vol'], reverse=True)[:5], key=lambda x: x['start'], reverse=True)

            # ==========================================
            # 5. 網頁 UI：顯示結果與圖表
            # ==========================================
            st.success(f"✅ {target_name} ({yf_ticker}) 分析完成！最新股價: {current_price_round:.2f}")
            
            # --- 繪製互動式手機圖表 (Plotly) ---
            st.subheader("📊 64日實體分價量分佈圖")
            df_plot = pd.DataFrame({
                '價格區間': [item['label'] for item in all_intervals_disp],
                '累積成交量 (張)': [int(item['vol']) for item in all_intervals_disp], # 改為張數
                '標記': ['現價所在' if item['is_current'] else '一般區間' for item in all_intervals_disp]
            })
            
            # 使用 Plotly 畫水平長條圖
            fig = px.bar(df_plot, x='累積成交量 (張)', y='價格區間', color='標記', 
                         color_discrete_map={'現價所在': '#FF4B4B', '一般區間': '#60B4FF'},
                         orientation='h')
                         
            # 強制鎖定價格排序，避免被 Plotly 自動按字首排序
            fig.update_yaxes(categoryorder='array', categoryarray=df_plot['價格區間'])
            
            # 反轉 Y 軸讓高價在上
            fig.update_layout(yaxis=dict(autorange="reversed"), margin=dict(l=0, r=0, t=30, b=0), height=500)
            st.plotly_chart(fig, use_container_width=True)

            # --- 顯示 Top 5 支撐壓力 ---
            st.subheader("🎯 關鍵支撐與壓力 (Top 5)")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**⬆️ 向上方壓力區**")
                for item in top_5_above_disp:
                    # 單位替換為「張」
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
                            # 轉換為張數
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
                                # 轉換為張數
                                net_buy = round(net_buy_shares / 1000)
                    except: pass
                    
                    if daily_vol == 0: 
                        # 取用前段已被除以1000的 Volume 資料
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
            
            # 整理要存檔的 DataFrame，更新所有「股」為「張」
            df_sr_excel = pd.DataFrame([{
                '項次': i+1, '價格級距區間 (TWD)': item['disp_label'], '累積成交量 (張)': int(item['vol'])
            } for i, item in enumerate(all_intervals_disp)])
            
            df_top5_excel = pd.DataFrame(
                [{'位置': '⬆️ 向上壓力區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_above_disp] +
                [{'位置': '🎯 最新股價 (中軸)', '價格級距區間': f"{current_price_round:.2f}", '累積成交量 (張)': 0}] +
                [{'位置': '⬇️ 向下支撐區', '價格級距區間': b['disp_label'], '累積成交量 (張)': int(b['vol'])} for b in top_5_below_disp]
            )

            # 在記憶體中建立 Excel
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
                chart.x_axis.title, chart.y_axis.title = "價格級距區間 (TWD)", "累積成交量 (張)" # 更新軸標題
                chart.x_axis.scaling.orientation, chart.y_axis.crosses = "maxMin", "max"
                chart.height, chart.width = max(10, len(df_sr_excel) * 0.5) * 1.5, 16 * 1.5
                
                max_r = len(df_sr_excel) + 1
                chart.add_data(Reference(sheet1, min_col=3, min_row=1, max_row=max_r), titles_from_data=True)
                chart.set_categories(Reference(sheet1, min_col=2, min_row=2, max_row=max_r))
                sheet1.add_chart(chart, "E2")
            
            # 產生下載按鈕
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
