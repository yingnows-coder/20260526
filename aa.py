import streamlit as st
# 導入服務帳戶連接核心套件
from streamlit_gsheets import GSheetsConnection

st.set_page_config(layout="wide")
st.title("階段二：雲端資料庫讀取與原始表格分析")
st.caption("授權標註：edit by 闕河正")

# 1. 建立雲端連接器
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. 從 Google 試算表讀取 "Tasks" 工作表
# 核心細節：ttl="0" 代表快取時間為 0 秒，強迫它每次重整都即時去雲端抓最新，不准用舊記憶
df = conn.read(worksheet="Tasks", ttl="0")

st.write("---")
st.write("### 這是從 Google 雲端硬碟抓回來的原始黑白表格（Bare Data）：")

# 3. 直接用 st.dataframe() 把整張表格原汁原味印在網頁上
st.dataframe(df)

# 4. 拆解底層資訊給學生看
st.write("經過 Python 分析，這張表格擁有的『直欄欄位名稱（Columns）』有：", list(df.columns))
