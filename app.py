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

# 👇 1. 請把取得日期的函數放在這裡 (讓 Python 先認識它)
@st.cache_data(ttl=3600, show_spinner=False)
def get_latest_trading_date():
    """抓取台股大盤指數，自動避開假日與颱風天，取得最後真實交易日"""
    try:
        idx_data = yf.Ticker("^TWII").history(period="5d")
        return idx_data.index[-1].strftime("%Y/%m/%d")
    except:
        return datetime.now().strftime("%Y/%m/%d")
        
# 👇 2. 這是原本讀取 Excel 的函數 (維持原樣)        
@st.cache_data
def load_stock_list():
    try:
        df = pd.read_excel('TW50100.xlsx', engine='openpyxl', dtype=str)
        return {str(row[df.columns[1]]): str(row[df.columns[0]]).replace('.0', '') for _, row in df.iterrows()}, True
    except: return {}, False


# ==========================================
# 副程式 1：抓取 YFinance 資料 (加入防擋偽裝與錯誤攔截機制)
# ==========================================
def step1_fetch_yf_data(ticker, raw_ticker, auto_fallback, target_date_str):
    end_dt = pd.to_datetime(target_date_str, format='%Y/%m/%d') + pd.Timedelta(days=1)
    start_dt = end_dt - pd.DateOffset(months=6) 
    
    start_str = start_dt.strftime('%Y-%m-%d')
    end_str = end_dt.strftime('%Y-%m-%d')

    # 建立自訂 Session，偽裝成一般的 Windows Chrome 瀏覽器
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })

    hist = pd.DataFrame()
    
    # 🟢 第一層抓取 (預設抓上市 .TW)
    try:
        hist = yf.Ticker(ticker, session=session).history(start=start_str, end=end_str)
    except Exception:
        pass # 如果發生 404 找不到或其他錯誤，直接略過，讓 hist 保持空值

    # 🟢 如果第一層沒抓到，且允許備用方案，啟動第二層抓取 (改抓上櫃 .TWO)
    if hist.empty and auto_fallback and raw_ticker:
        ticker_two = f"{raw_ticker}.TWO"
        try:
            hist_two = yf.Ticker(ticker_two, session=session).history(start=start_str, end=end_str)
            if not hist_two.empty:
                hist = hist_two
                ticker = ticker_two
        except Exception:
            pass

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

# ==========================================
# 🟢 升級版：上市 (TWSE) 融資券下載與解析 (支援快取)
# ==========================================
@st.cache_data(ttl=86400, show_spinner=False)
def download_twse_margin_json(date_str):
    """下載並快取上市融資券 JSON (全市場總表)"""
    url = f"https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={date_str}&selectType=ALL"
    time.sleep(1) # 只有真正連網時才會觸發延遲，保護 IP
    try:
        res = requests.get(url, timeout=5).json()
        if res.get('stat') == 'OK':
            return res
    except: pass
    return {}

def fetch_margin_json_data(date_str, raw_ticker):
    """從快取的全市場資料中，撈出特定上市股票"""
    res = download_twse_margin_json(date_str)
    if not res: return 0, 0, 0, 0
    
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
    return 0, 0, 0, 0

# ==========================================
# TPEx (上櫃) 專用下載與解析函數 (終極防禦版 + 支援買賣雙向)
# ==========================================
@st.cache_data(ttl=86400, show_spinner=False)
def download_tpex_csv_text(date_str, inst_type, search_type="buy"):
    """下載上櫃法人 CSV，強制略過 SSL 檢查並安全解碼，支援 searchType 分流"""
    # 🟢 網址加入 {search_type} 變數，動態抓取 buy 或 sell
    url = f"https://www.tpex.org.tw/www/zh-tw/insti/{inst_type}?type=Daily&date={date_str}&searchType={search_type}&id=&response=csv"
    time.sleep(1) 
    try:
        res = requests.get(url, timeout=5, verify=False)
        if len(res.content) > 100: 
            return res.content.decode('cp950', errors='ignore')
    except: pass
    return ""

# ==========================================
# 🟢 升級版：上櫃 (TPEx) 融資券下載與解析 (支援快取)
# ==========================================
@st.cache_data(ttl=86400, show_spinner=False)
def download_tpex_margin_json(roc_date_str):
    """下載並快取上櫃融資券 JSON (全市場總表)"""
    url = f"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=json&d={roc_date_str}"
    time.sleep(1) # 只有真正連網時才會觸發延遲
    try:
        # 強制略過 SSL 檢查
        return requests.get(url, timeout=5, verify=False).json()
    except: pass
    return {}

def fetch_tpex_margin_json_data(roc_date_str, raw_ticker):
    """從快取的全市場資料中，撈出特定上櫃股票"""
    res = download_tpex_margin_json(roc_date_str)
    if not res: return 0, 0, 0, 0
    
    tables = res.get('tables', [])
    for table in tables:
        for row in table.get('data', []):
            if str(row[0]).strip() == raw_ticker:
                # 依據 fields 定位：前資(2)、資餘(6)、前券(10)、券餘(14)
                m_prev = int(str(row[2]).replace(',', ''))
                m_today = int(str(row[6]).replace(',', ''))
                s_prev = int(str(row[10]).replace(',', ''))
                s_today = int(str(row[14]).replace(',', ''))
                return (m_today - m_prev), m_today, (s_today - s_prev), s_today
    return 0, 0, 0, 0


# ==========================================
# 🟢 修正 2：籌碼主迴圈 (加入終極負數清洗器)
# ==========================================
def step6_extract_institutional_data(raw_ticker, hist_64, is_otc):
    last_20_dates = hist_64.index[-20:]
    last_10_dates = hist_64.index[-10:]
    
    foreign_records = []
    trust_records = []
    margin_records = []
    
    # 🛡️ 終極過濾器：精準捕捉全形負號、括號、三角形等詭異負數格式
    def safe_parse_int(val_str):
        s = str(val_str).strip()
        if not s: return 0
        
        # 1. 判斷是否為負數 (涵蓋全形/半形減號、括號、財報三角形)
        is_negative = False
        if s.startswith('-') or s.startswith('－') or s.startswith('−') or \
           (s.startswith('(') and s.endswith(')')) or '△' in s or '▲' in s:
            is_negative = True
            
        # 2. 暴力拔除所有非數字字元 (只留下純數字)
        cleaned = re.sub(r'\D', '', s)
        if not cleaned: return 0
        
        # 3. 轉為整數，若為負數則掛上負號
        val = int(cleaned)
        return -val if is_negative else val
    
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
                    net_f = round(safe_parse_int(target_row.iloc[0, 5]) / 1000)
            foreign_records.append({'日期': date_disp_str, '外資買賣超(張)': net_f})
            
            # 2. 投信 (20天)
            df_trust = fetch_twse_csv_data(date_api_str, "TWT44U")
            net_t = 0
            if not df_trust.empty:
                df_trust[1] = df_trust[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_trust[df_trust[1] == raw_ticker]
                if not target_row.empty:
                    net_t = round(safe_parse_int(target_row.iloc[0, 5]) / 1000)
            trust_records.append({'日期': date_disp_str, '投信買賣超(張)': net_t})
            
            # 3. 融資券 (10天)
            if d in last_10_dates:
                m_change, m_today, s_change, s_today = fetch_margin_json_data(date_api_str, raw_ticker)
                margin_records.append({'日期': date_disp_str, '融資變動(張)': m_change, '融資餘額(張)': m_today, '融券變動(張)': s_change, '融券餘額(張)': s_today})
                
        else:
            # ==========================================
            # 🟢 上櫃 (TPEx) 邏輯分支 (支援買賣超雙向查詢)
            # ==========================================
            date_tpex_csv_str = d.strftime('%Y/%m/%d')
            
            # ----------------------------------
            # 1. 上櫃外資 (先查買超，沒有再查賣超)
            # ----------------------------------
            net_f = 0
            found_f = False
            
            # 先查 Buy 檔
            csv_f_buy = download_tpex_csv_text(date_tpex_csv_str, "qfiiStat", "buy")
            if csv_f_buy:
                df_f_buy = pd.read_csv(io.StringIO(csv_f_buy), names=list(range(20)), on_bad_lines='skip')
                df_f_buy[1] = df_f_buy[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_f_buy[df_f_buy[1] == raw_ticker]
                if not target_row.empty:
                    net_f = safe_parse_int(target_row.iloc[0, 5])
                    found_f = True
            
            # 買超找不到，改查 Sell 檔
            if not found_f:
                csv_f_sell = download_tpex_csv_text(date_tpex_csv_str, "qfiiStat", "sell")
                if csv_f_sell:
                    df_f_sell = pd.read_csv(io.StringIO(csv_f_sell), names=list(range(20)), on_bad_lines='skip')
                    df_f_sell[1] = df_f_sell[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                    target_row = df_f_sell[df_f_sell[1] == raw_ticker]
                    if not target_row.empty:
                        val = safe_parse_int(target_row.iloc[0, 5])
                        # 🟢 確保賣超資料轉為負數 (無論原始表有無負號，強制取絕對值後掛負號)
                        net_f = -abs(val)
                        
            foreign_records.append({'日期': date_disp_str, '外資買賣超(張)': net_f})
            
            # ----------------------------------
            # 2. 上櫃投信 (先查買超，沒有再查賣超)
            # ----------------------------------
            net_t = 0
            found_t = False
            
            # 先查 Buy 檔
            csv_t_buy = download_tpex_csv_text(date_tpex_csv_str, "sitcStat", "buy")
            if csv_t_buy:
                df_t_buy = pd.read_csv(io.StringIO(csv_t_buy), names=list(range(20)), on_bad_lines='skip')
                df_t_buy[1] = df_t_buy[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                target_row = df_t_buy[df_t_buy[1] == raw_ticker]
                if not target_row.empty:
                    net_t = safe_parse_int(target_row.iloc[0, 5])
                    found_t = True
                    
            # 買超找不到，改查 Sell 檔
            if not found_t:
                csv_t_sell = download_tpex_csv_text(date_tpex_csv_str, "sitcStat", "sell")
                if csv_t_sell:
                    df_t_sell = pd.read_csv(io.StringIO(csv_t_sell), names=list(range(20)), on_bad_lines='skip')
                    df_t_sell[1] = df_t_sell[1].astype(str).str.replace(r'[=" ]', '', regex=True)
                    target_row = df_t_sell[df_t_sell[1] == raw_ticker]
                    if not target_row.empty:
                        val = safe_parse_int(target_row.iloc[0, 5])
                        # 🟢 確保賣超資料轉為負數
                        net_t = -abs(val)
                        
            trust_records.append({'日期': date_disp_str, '投信買賣超(張)': net_t})
            
            # ----------------------------------
            # 3. 上櫃融資券 (10天) 維持不動
            # ----------------------------------
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
# 介面繪製輔助函數 (Tech Chart)
# ==========================================
def render_tech_chart(hist_64, show_ma5, show_ma10, show_ma20, allow_zoom):
    date_strings = hist_64.index.strftime('%Y-%m-%d')
    
    # 🟢 包含 4 個子圖與成交量均量的版本
    fig_k = make_subplots(rows=4, cols=1, shared_xaxes=True, vertical_spacing=0.03, 
                          row_heights=[0.4, 0.2, 0.2, 0.2], 
                          subplot_titles=("價格與均線", "KD (9,3,3)", "MACD (12,26,9)", "成交量與64日均量"))
    
    # --- Row 1: 價格與均線 ---
    fig_k.add_trace(go.Candlestick(x=date_strings, open=hist_64['Open'], high=hist_64['High'], low=hist_64['Low'], close=hist_64['Close'], name='K線', increasing_line_color='#FF4B4B', increasing_fillcolor='#FF4B4B', decreasing_line_color='#00B050', decreasing_fillcolor='#00B050'), row=1, col=1)
    if show_ma5: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA5'], name='5MA', line=dict(color='#7A431D', width=1.5)), row=1, col=1)
    if show_ma10: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA10'], name='10MA', line=dict(color='#00E5FF', width=1.5)), row=1, col=1)
    if show_ma20: fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MA20'], name='20MA', line=dict(color='#0D47A1', width=1.5)), row=1, col=1)
    
    # --- Row 2: KD ---
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['K'], name='K值', line=dict(color='#FF9900', width=1.2)), row=2, col=1)
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['D'], name='D值', line=dict(color='#0066FF', width=1.2)), row=2, col=1)
    
    # --- Row 3: MACD ---
    macd_colors = ['#FF4B4B' if val > 0 else '#00B050' for val in hist_64['OSC']]
    fig_k.add_trace(go.Bar(x=date_strings, y=hist_64['OSC'], name='OSC', marker_color=macd_colors), row=3, col=1)
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['DIF'], name='DIF', line=dict(color='#FF9900', width=1.2)), row=3, col=1)
    fig_k.add_trace(go.Scatter(x=date_strings, y=hist_64['MACD'], name='MACD', line=dict(color='#0066FF', width=1.2)), row=3, col=1)
    
    # --- Row 4: 成交量與 64 日均量 ---
    vol_colors = ['#FF4B4B' if row['Close'] >= row['Open'] else '#00B050' for idx, row in hist_64.iterrows()]
    fig_k.add_trace(go.Bar(x=date_strings, y=hist_64['Volume'], name='成交量(張)', marker_color=vol_colors), row=4, col=1)
    
    avg_vol = hist_64['Volume'].mean()
    fig_k.add_trace(go.Scatter(x=date_strings, y=[avg_vol]*len(hist_64), name=f'64日均量({int(avg_vol)}張)', mode='lines', line=dict(color='#FFD700', width=2, dash='dash')), row=4, col=1)
    
    fig_k.update_layout(
        xaxis=dict(type='category', visible=False), 
        xaxis2=dict(type='category', visible=False), 
        xaxis3=dict(type='category', visible=False), 
        xaxis4=dict(type='category', visible=True, title="交易日期", nticks=10),
        yaxis=dict(visible=False), 
        yaxis2=dict(visible=True), 
        yaxis3=dict(visible=True),
        yaxis4=dict(visible=True), 
        xaxis_rangeslider_visible=False, 
        margin=dict(l=4, r=4, t=30, b=4), 
        height=850, 
        hovermode='x unified', 
        showlegend=False
    )
    
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
    
    # 🐛 Debug 模式開關 (已從 UI 隱藏，未來若需使用，將下方兩行註解反轉即可)
    # show_debug = st.checkbox("🐛 開啟爬蟲 Debug 模式 (用來檢視原始資料的欄位位置)")
    show_debug = False  # 👈 強制設為 False，直接跳過 Debug 顯示區塊
    
    c_title, c_btn = st.columns([4, 1])
    with c_title:
        st.subheader("📈 近期市場籌碼動向 (外資投信 20日 / 融資券 10日)")
    with c_btn:
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 重新下載快取", use_container_width=True):
            download_twse_csv_text.clear()
            download_tpex_csv_text.clear()
            # 🟢 補上清除融資券快取的指令
            download_twse_margin_json.clear()
            download_tpex_margin_json.clear()
            
    is_otc = yf_ticker.endswith('.TWO')
    
   # ==========================================
    # 🐛 Debug 模式專用顯示區塊 (安全版)
    # ==========================================
    if show_debug and is_otc:
        st.warning("🔍 【Debug 模式】目前為上櫃股票，以下印出最新一天的原始純文字 (已關閉表格渲染以防止當機)：")
        test_d = hist_64.index[-1]
        # 🟢 這裡也要改回正常的斜線
        test_tpex_csv = test_d.strftime('%Y/%m/%d') 
        test_roc = f"{test_d.year - 1911}/{test_d.strftime('%m/%d')}"
        
        c_d1, c_d2 = st.columns(2)
        with c_d1:
            st.write(f"👉 **外資 CSV 原始文字檔前 500 字** ({test_tpex_csv})")
            f_txt = download_tpex_csv_text(test_tpex_csv, "qfiiStat")
            if f_txt:
                # 🟢 徹底避開 PyArrow，直接印出純文字，保證絕對不當機！
                st.text(f_txt[:500])
            else:
                st.error("無資料或連線失敗")
        with c_d2:
            st.write(f"👉 **融資券 JSON 原始回傳** ({test_roc})")
            url_m = f"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=json&d={test_roc}"
            try:
                # 🟢 加上 verify=False，並直接讀取純文字 text
                res_raw = requests.get(url_m, timeout=5, verify=False).text
                st.text(res_raw[:500])
            except Exception as e:
                st.error(f"解析失敗: {e}")
        
        st.info("💡 檢查完畢後，請取消勾選 Debug 模式，即可恢復正常繪圖。")
        st.stop()
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
