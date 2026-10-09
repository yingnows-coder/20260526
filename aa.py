import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
import uuid
import io
import os
import plotly.express as px
import plotly.graph_objects as go

# Excel 匯出為選用功能；即使環境尚未安裝 openpyxl，主系統仍可正常執行
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.drawing.image import Image as XLImage
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# =========================================================
# 1. 基本設定
# =========================================================

st.set_page_config(
    page_title="企業版 RFQ 詢價追蹤系統",
    page_icon="📌",
    layout="wide"
)

# =========================================================
# 系統標題 + JENN-WEI LOGO
# =========================================================

logo_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "jenn_wei_logo.png"
)

header_logo, header_title = st.columns(
    [1, 9],
    vertical_alignment="center"
)

with header_logo:
    if os.path.exists(logo_path):
        st.image(logo_path, width=85)
    else:
        st.warning("Logo 找不到")

with header_title:
    st.markdown(
        """
        <h1 style="
            margin-top:0;
            margin-bottom:2px;
            padding-top:0;
        ">
            企業版：RFQ 詢價追蹤管理系統
        </h1>
        """,
        unsafe_allow_html=True
    )
    st.caption("RFQ Workflow + 報價產生器 | edit by 林溫城")


# =========================================================
# 2. BB 頁面
# =========================================================

# cc.py 獨立部署於 Streamlit Community Cloud 後，將網址填入此處。
# 例如：https://your-service-app.streamlit.app/
SERVICE_APP_URL = "https://jwcncmaintancesheet.streamlit.app/"

nav_spacer, nav_trello, nav_service = st.columns([5, 2.5, 2.5])

with nav_trello:
    st.link_button(
        "📋 雲端 Trello 管理系統",
        "https://aazzyybb.streamlit.app/",
        use_container_width=True,
    )

with nav_service:
    if SERVICE_APP_URL.startswith(("https://", "http://")):
        st.link_button(
            "🛠️ 售服維修管理系統",
            SERVICE_APP_URL,
            use_container_width=True,
        )
    else:
        st.caption("🛠️ 售服維修管理系統：請先設定獨立網址")


# =========================================================
