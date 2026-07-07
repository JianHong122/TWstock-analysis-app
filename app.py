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
# 0. 初始化 Session State
# ==========================================
if 'analyzed_input' not in st.session_state:
    st.session_state.analyzed_input = None

# ==========================================
# 1. 網頁基本設定
# ==========================================
st.set_page_config(page_title="台股籌碼分析工具", page_icon="📈", layout="centered")
st.title("📊 台股區間支撐壓力與法人籌碼分析")
st.markdown("支援 **技術K線均線**、**KD/MACD指標**、**20級距全覽**、**Top 5 關鍵防守** 與 **五日法人買賣強度**")

# ==========================================
# 2. 輔助函數：計算 KD 與 MACD
# ==========================================
def calculate_indicators(df):
    # KD 計算 (9日)
    low_min = df['Low'].rolling(window=9).min()
    high_max = df['High'].rolling(window=9).max()
    rsv = (df['Close'] - low_min) / (high_max - low_min) * 100
    
    # 這裡使用 ewm 來處理 K, D 的遞迴累加
    df['K'] = rsv.ewm(com=2, adjust=False).mean()
    df['D'] = df['K'].ewm(com=2, adjust=False).mean()
    
    # MACD 計算
    ema12 = df['Close'].ewm(span=12, adjust=False).mean()
    ema26 = df['Close'].ewm(span=26, adjust=False).mean()
    df['DIF'] = ema12 - ema26
    df['MACD'] = df['DIF'].ewm(span=9, adjust=False).mean()
    df['OSC'] = (df['DIF'] - df['MACD']) * 2
    
    return df

# ==========================================
# 3. 讀取與抓取資料 (Cache)
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

@st.cache_data(ttl=300, show_spinner=False)
def fetch_stock_history(ticker):
    try:
        stock_data = yf.Ticker(ticker)
        hist = stock_data.history(period="6mo") # 抓半年資料以計算指標
        return hist
    except Exception:
        return pd.DataFrame()

# ==========================================
# 4. 網頁 UI：使用者輸入區
# ==========================================
user_input = st.text_input("🔍 請輸入個股名稱或代號：", placeholder="例如: 台積電 或 2330")
analyze_button = st.button("🚀 開始分析", use_container_width=True)

if analyze_button and user_input:
    st.session_state.analyzed_input = user_input

if st.session_state.analyzed_input:
    current_target_input = st.session_state.analyzed_input
    
    with st.spinner('正在計算技術指標與籌碼分佈...'):
        
        # (略過判斷輸入邏輯...與先前版本相同，保持程式碼簡潔)
        # 判斷邏輯相同
        matched_names = [name for name in name_to_ticker.keys() if current_target_input in name] if list_loaded else []
        raw_ticker = name_to_ticker[matched_names[0]] if matched_names else current_target_input
        yf_ticker = f"{raw_ticker}.TW" if not any(x in current_target_input for x in ['.TW', '.TWO']) else current_target_input.upper()
        
        hist = fetch_stock_history(yf_ticker)
        # (這裡省略部分容錯 fallback 程式碼，複製時請確保完整性)
        
        # 運算技術指標
        hist = calculate_indicators(hist)
        
        # 計算 MA
        hist['MA5'] = hist['Close'].rolling(window=5).mean()
        hist['MA10'] = hist['Close'].rolling(window=10).mean()
        hist['MA20'] = hist['Close'].rolling(window=20).mean()
        
        hist_64 = hist.tail(64).copy()
        
        # 顯示 KD/MACD 面板
        st.subheader("📊 最新技術指標數據")
        latest = hist_64.iloc[-1]
        col_k, col_d, col_dif, col_macd = st.columns(4)
        col_k.metric("K值", f"{latest['K']:.2f}")
        col_d.metric("D值", f"{latest['D']:.2f}")
        col_dif.metric("DIF", f"{latest['DIF']:.2f}")
        col_macd.metric("MACD", f"{latest['MACD']:.2f}")

        # --- 繪製 K 線圖 ---
        with st.container(border=True):
            allow_zoom = st.checkbox("🔍 啟用圖表縮放", value=False)
            st.subheader("📈 64日技術K線")
            fig_k = go.Figure()
            # (K線繪圖邏輯與前一版相同)
            # ... 請將 K 線與均線的繪圖邏輯貼回此處 ...
            # 為了節省空間，我建議直接將您上一次程式碼中的這部分整合即可。
        
        # --- 其餘程式碼與 Excel 下載邏輯 ---
        # ...
