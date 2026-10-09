"""JENN-WEI shared customer/supplier master data for aa.py, bb.py, cc.py."""
import uuid
from datetime import datetime
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

CUSTOMER_COLUMNS = ['customer_id','company_name','tax_id','country','address','contact','phone','email','industry','status','notes','created_at','updated_at']
SUPPLIER_COLUMNS = ['supplier_id','company_name','tax_id','category','country','address','contact','phone','email','payment_terms','lead_time','status','notes','created_at','updated_at']


def spreadsheet_url():
    try:
        return str(st.secrets.get('SERVICE_SPREADSHEET_URL','') or st.secrets.get('connections',{}).get('gsheets',{}).get('spreadsheet','')).strip()
    except Exception:
        return ''


def connection():
    return st.connection('gsheets', type=GSheetsConnection)


def load_master(kind):
    sheet, columns = config(kind)
    url = spreadsheet_url()
    if not url:
        st.error('請在本 App Secrets 的 [connections.gsheets] 設定 spreadsheet（與 aa.py 相同的網址）。')
        return None
    try:
        data = connection().read(spreadsheet=url, worksheet=sheet, ttl=60)
    except Exception as exc:
        st.error(f'無法讀取 {sheet}。請先在 Google Sheets 建立該分頁及第一列欄位。')
        st.caption(f'{type(exc).__name__}: {exc}')
        return None
    if data is None:
        data = pd.DataFrame()
    for col in columns:
        if col not in data.columns:
            data[col] = ''
    return data[columns].fillna('').astype(object).copy()


def config(kind):
    if kind == 'customer':
        return 'Customers', CUSTOMER_COLUMNS
    if kind == 'supplier':
        return 'Suppliers', SUPPLIER_COLUMNS
    raise ValueError('Unknown master type')


def save_master(kind, data):
    sheet, cols = config(kind)
    try:
        connection().update(spreadsheet=spreadsheet_url(), worksheet=sheet, data=data[cols].fillna(''))
        st.cache_data.clear()
        st.success('主檔已儲存，其他 App 最多可能需要 60 秒刷新快取。')
        st.rerun()
    except Exception as exc:
        st.error(f'無法儲存 {sheet}：{type(exc).__name__}: {exc}')


def render_master(kind, prefix='shared'):
    sheet, columns = config(kind)
    data = load_master(kind)
    if data is None:
        st.code('\t'.join(columns), language='text')
        return
    id_col = columns[0]
    label = '客戶' if kind == 'customer' else '廠商'
    st.subheader(f'{label}主檔管理')
    st.caption(f'共用工作表：{sheet}｜使用編號關聯，停用資料不刪除歷史紀錄')
    search = st.text_input('搜尋公司／編號／聯絡人／電話', key=f'{prefix}_{kind}_search')
    if search:
        filtered = data[data.astype(str).apply(lambda col: col.str.contains(search, case=False, regex=False, na=False)).any(axis=1)]
    else:
        filtered = data
    st.dataframe(filtered[[id_col,'company_name','contact','phone','status','updated_at']], use_container_width=True, hide_index=True)
    choices = ['➕ 新增'] + [str(x) for x in data[id_col] if str(x).strip()]
    choice = st.selectbox('新增或編輯', choices, key=f'{prefix}_{kind}_choice')
    editing = choice != '➕ 新增'
    record = data[data[id_col].astype(str) == choice].iloc[0].to_dict() if editing else {}
    def old(key):
        value = record.get(key,'')
        return '' if pd.isna(value) else str(value)
    with st.form(f'{prefix}_{kind}_form'):
        inputs = {}
        for col in columns:
            if col in (id_col,'created_at','updated_at'):
                continue
            display = {'company_name':'公司名稱 *','tax_id':'統一編號','country':'國家／地區','address':'地址','contact':'聯絡人','phone':'電話','email':'Email','industry':'產業別','category':'供應類別','payment_terms':'付款條件','lead_time':'交期','notes':'備註','status':'狀態'}
            if col == 'status':
                opts = ['啟用','停用']
                inputs[col] = st.selectbox(display[col], opts, index=1 if old(col)=='停用' else 0)
            elif col == 'notes':
                inputs[col] = st.text_area(display[col], value=old(col))
            else:
                inputs[col] = st.text_input(display.get(col,col), value=old(col))
        submitted = st.form_submit_button('💾 儲存主檔', type='primary')
    if submitted:
        if not inputs['company_name'].strip():
            st.error('公司名稱為必填。')
            return
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if editing:
            idx = data.index[data[id_col].astype(str)==choice][0]
            for key, value in inputs.items():
                data.at[idx,key] = value.strip() if isinstance(value,str) else value
            data.at[idx,'updated_at'] = now
        else:
            company = inputs['company_name'].strip().casefold()
            if (data['company_name'].astype(str).str.strip().str.casefold()==company).any():
                st.error('已有相同公司名稱，請編輯既有主檔，避免重複建檔。')
                return
            new = {c:'' for c in columns}
            new.update(inputs)
            new[id_col] = ('CUS' if kind=='customer' else 'SUP') + '-' + datetime.now().strftime('%Y%m%d') + '-' + uuid.uuid4().hex[:6].upper()
            new['created_at'] = now
            new['updated_at'] = now
            data = pd.concat([data,pd.DataFrame([new])],ignore_index=True)
        save_master(kind,data)


def customer_options():
    data = load_master('customer')
    if data is None:
        return [], {}
    active = data[data['status'].astype(str)!='停用']
    labels = [f"{row['company_name']}｜{row['customer_id']}" for _,row in active.iterrows() if str(row['company_name']).strip()]
    mapping = {f"{row['company_name']}｜{row['customer_id']}":row.to_dict() for _,row in active.iterrows() if str(row['company_name']).strip()}
    return labels, mapping
