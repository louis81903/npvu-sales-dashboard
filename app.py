import streamlit as st
import pandas as pd
import plotly.express as px

# 設定網頁標題與寬度
st.set_page_config(page_title="nPVU 銷售業績互動報表", layout="wide")
st.title("📊 nPVU 銷售業績互動報表 (2021-2025)")
st.markdown("透過左側選單篩選資料，圖表與數據將即時連動更新。")

@st.cache_data
def load_data():
    file_path = "nPVU sales performance.xlsx" # 確認你的檔名是這個
    
    try:
        # 嘗試讀取 Excel 檔案，尋找包含 'Sales data' 的工作表，並跳過前 4 行標題空白列
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_names = xls.sheet_names
        
        # 自動尋找名稱裡有 "Sales data" 的工作表，找不到就讀取第一個
        target_sheet = next((s for s in sheet_names if 'Sales data' in s), sheet_names[3] if len(sheet_names)>3 else sheet_names[0])
        
        df = pd.read_excel(file_path, sheet_name=target_sheet, skiprows=4, engine='openpyxl')
        
        # 資料清理與型態轉換
        df['Month'] = df['Month'].astype(str).str.replace(r'\.0$', '', regex=True) # 確保月份格式乾淨
        df['Year'] = df['Month'].str[:4] # 擷取年份
        
        # 確保銷售額與數量是數字型態
        df['Net Sales Amount'] = pd.to_numeric(df['Net Sales Amount'], errors='coerce').fillna(0)
        df['Net Sales Quantity'] = pd.to_numeric(df['Net Sales Quantity'], errors='coerce').fillna(0)
        
        # 填補空值
        df['Product'] = df['Product'].fillna('Unknown')
        df['Salesman Name'] = df['Salesman Name'].fillna('Unknown')
        
        return df
        
    except FileNotFoundError:
        st.error(f"找不到檔案 '{file_path}'，請確認檔案已經上傳到 GitHub，且檔名完全一致 (包含大小寫與空格)。")
        st.stop()
    except Exception as e:
        st.error(f"讀取 Excel 發生錯誤，請確認工作表格式：{e}")
        st.stop()

# 載入資料
df = load_data()

# --- 側邊欄：互動篩選器 ---
st.sidebar.header("🔍 資料篩選器")

# 年份篩選
available_years = sorted([y for y in df['Year'].unique() if y.isdigit() and len(y)==4])
years = st.sidebar.multiselect("選擇年份", options=available_years, default=available_years)

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
# 挑選重要欄位顯示
display_columns = ['Month', 'Transaction Date', 'Product', 'Item Description', 'Salesman Name', 'Customer Name', 'Net Sales Amount', 'Net Sales Quantity']
# 確保欄位存在於資料中才顯示
existing_columns = [col for col in display_columns if col in filtered_df.columns]

st.dataframe(filtered_df[existing_columns], height=300)
