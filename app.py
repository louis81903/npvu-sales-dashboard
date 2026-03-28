import streamlit as st
import pandas as pd
import plotly.express as px

# 設定網頁標題與寬度
st.set_page_config(page_title="nPVU 銷售業績互動報表", layout="wide")
st.title("📊 nPVU 銷售業績互動報表 (2021-2025)")
st.markdown("透過左側選單篩選資料，圖表與數據將即時連動更新。")

# 讀取資料 (這裡讀取您的原始銷售明細)
@st.cache_data
def load_data():
    # 讀取 Sales data，跳過前四行標題空白列
    df = pd.read_csv("nPVU sales performance.xlsx", skiprows=4)
    
    # 資料清理與型態轉換
    df['Month'] = df['Month'].astype(str)
    df['Year'] = df['Month'].str[:4] # 擷取年份
    
    # 確保銷售額與數量是數字型態
    df['Net Sales Amount'] = pd.to_numeric(df['Net Sales Amount'], errors='coerce').fillna(0)
    df['Net Sales Quantity'] = pd.to_numeric(df['Net Sales Quantity'], errors='coerce').fillna(0)
    
    # 填補空值
    df['Product'] = df['Product'].fillna('Unknown')
    df['Salesman Name'] = df['Salesman Name'].fillna('Unknown')
    return df

# 載入資料
try:
    df = load_data()
except Exception as e:
    st.error(f"資料讀取失敗，請確認檔案是否存在。錯誤訊息：{e}")
    st.stop()

# --- 側邊欄：互動篩選器 ---
st.sidebar.header("🔍 資料篩選器")

# 年份篩選
years = st.sidebar.multiselect("選擇年份", options=sorted(df['Year'].unique()), default=sorted(df['Year'].unique()))

# 產品篩選
products = st.sidebar.multiselect("選擇產品 (Product)", options=df['Product'].unique(), default=df['Product'].unique())

# 業務員篩選
salesmen = st.sidebar.multiselect("選擇業務員 (Salesman Name)", options=df['Salesman Name'].unique(), default=df['Salesman Name'].unique())

# 根據篩選條件過濾資料
filtered_df = df[
    (df['Year'].isin(years)) &
    (df['Product'].isin(products)) &
    (df['Salesman Name'].isin(salesmen))
]

# --- 主要指標 (KPI) ---
st.markdown("### 💰 銷售總覽")
col1, col2, col3 = st.columns(3)
col1.metric("總銷售額 (Net Sales Amount)", f"NT$ {filtered_df['Net Sales Amount'].sum():,.0f}")
col2.metric("總銷售量 (Net Sales Quantity)", f"{filtered_df['Net Sales Quantity'].sum():,.0f}")
col3.metric("出貨客戶數量 (Customer)", f"{filtered_df['Customer Name'].nunique()}")

st.divider()

# --- 互動圖表區 ---
col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("#### 📈 每月銷售額趨勢")
    if not filtered_df.empty:
        trend_data = filtered_df.groupby('Month')['Net Sales Amount'].sum().reset_index()
        fig_trend = px.line(trend_data, x='Month', y='Net Sales Amount', markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

with col_chart2:
    st.markdown("#### 🏆 各業務員銷售表現")
    if not filtered_df.empty:
        salesrep_data = filtered_df.groupby('Salesman Name')['Net Sales Amount'].sum().reset_index().sort_values('Net Sales Amount', ascending=False)
        fig_bar = px.bar(salesrep_data, x='Salesman Name', y='Net Sales Amount', color='Salesman Name')
        st.plotly_chart(fig_bar, use_container_width=True)

st.divider()

# --- 原始資料表 ---
st.markdown("#### 📄 篩選後的明細資料")
st.dataframe(filtered_df[['Month', 'Transaction Date', 'Product', 'Item Description', 'Salesman Name', 'Customer Name', 'Net Sales Amount', 'Net Sales Quantity']], height=300)