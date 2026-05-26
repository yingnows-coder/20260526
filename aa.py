import streamlit as st

# 1. 網頁初始化設定（必須放在程式碼第一行）
# layout="wide" 會把網頁兩邊的留白填滿，變成寬螢幕，最適合看多欄位的看板
st.set_page_config(layout="wide")

st.title("階段一：Trello 畫布空間規劃測試")
st.caption("授權標註：edit by 闕河正 | 專屬資淺初學者講義")

st.write("---")

# 2. 呼叫 st.columns(3)，在網頁橫向切出三個一模一樣寬度的大直欄變數
col1, col2, col3 = st.columns(3)

# 3. 運用 with 語法，像填空一樣把文字塞進對應的直欄空間裡
with col1:
    st.markdown("### To Do (待辦)")
    st.write("這裡未來要放『待辦事項』的卡片")

with col2:
    st.markdown("### In Progress (執行中)")
    st.write("這裡未來要放『執行中』的卡片")

with col3:
    st.markdown("### Done (已完成)")
    st.write("這裡未來要放『已完成』的卡片")
