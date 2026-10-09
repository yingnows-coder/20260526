"""共用 Employees 員工主檔讀取與選擇元件；僅 bb.py 可編輯主檔。"""
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

EMPLOYEE_COLUMNS = ['employee_id','name','department','job_title','phone','email','hire_date','supervisor','status','notes','created_at','updated_at']

def employee_names():
    try:
        direct = st.secrets.get('SERVICE_SPREADSHEET_URL', '')
        config = st.secrets.get('connections', {}).get('gsheets', {})
        url = str(direct or config.get('spreadsheet', '') or config.get('spreadsheet_url', '')).strip()
        conn = st.connection('gsheets', type=GSheetsConnection)
        options = {'worksheet': 'Employees', 'ttl': 30}
        if url:
            options['spreadsheet'] = url
        df = conn.read(**options)
        if df is None or df.empty or 'name' not in df.columns:
            return []
        statuses = df['status'].fillna('').astype(str).str.strip() if 'status' in df.columns else pd.Series([''] * len(df), index=df.index)
        active = df.loc[~statuses.isin(['離職','留職停薪','停用']), 'name']
        return sorted({str(v).strip() for v in active if pd.notna(v) and str(v).strip()})
    except Exception as exc:
        st.warning(f'無法讀取共用員工主檔 Employees：{exc}。仍可手動輸入。')
        return []

def employee_selector(label, key, current='', names=None):
    current = str(current or '').strip()
    names = employee_names() if names is None else names
    choices = ['✍️ 手動輸入'] + list(names)
    if current and current not in choices:
        choices.append(current)
    index = choices.index(current) if current in choices else 0
    selected = st.selectbox(label, choices, index=index, key=f'{key}_employee_choice')
    if selected == '✍️ 手動輸入':
        return st.text_input(f'{label}（手動輸入）', value=current, key=f'{key}_employee_manual')
    return selected
