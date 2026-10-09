"""JENN-WEI Trello 任務管理：共用客戶與廠商主檔版本。
注意：本檔是依既有 Tasks/StatusLog 欄位重新建立的完整可執行版，並非原 bb.py 逐行升級。
"""
import uuid
from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection
from shared_contacts import render_master, load_master

st.set_page_config(page_title='JENN-WEI Trello 任務管理', layout='wide', page_icon='📊')
st.title('📊 JENN-WEI 雲端 Trello 任務管理')
st.caption('共用客戶與廠商主檔｜Google Sheets')
a,b,_=st.columns([2,2,5])
with a: st.link_button('📋 RFQ 詢價管理','https://aazzyyii.streamlit.app/',use_container_width=True)
with b: st.link_button('🛠️ 售服維修管理','https://jwcncmaintancesheet.streamlit.app/',use_container_width=True)

def spreadsheet_url():
    try:
        return str(st.secrets.get('SERVICE_SPREADSHEET_URL','') or st.secrets.get('connections',{}).get('gsheets',{}).get('spreadsheet','')).strip()
    except Exception: return ''
URL=spreadsheet_url()
if not URL:
    st.error('請在 bb.py App 的 Secrets 設定 SERVICE_SPREADSHEET_URL 或 [connections.gsheets] spreadsheet。')
    st.stop()
conn=st.connection('gsheets',type=GSheetsConnection)
COLUMNS=['id','department','customer','title','status','owner','created_time','due_time','updated_time','next_step','evidence','exception','RFQ_ID','version','approved_quote','quote_sent','first_followup_due']
STATES=['新詢價','待補件','工程評估','待核價','已報價','追蹤中','結案']

def read(sheet,cols):
    try:
        df=conn.read(spreadsheet=URL,worksheet=sheet,ttl=30)
        if df is None: df=pd.DataFrame()
        for col in cols:
            if col not in df.columns: df[col]=''
        return df.copy().astype(object)
    except Exception as e:
        st.error(f'讀取 {sheet} 失敗：{e}')
        st.stop()

def save(sheet,df):
    try:
        conn.update(spreadsheet=URL,worksheet=sheet,data=df.fillna(''))
        st.cache_data.clear()
        st.rerun()
    except Exception as e: st.error(f'儲存 {sheet} 失敗：{e}')

def s(value): return '' if pd.isna(value) else str(value)
def today(): return date.today().isoformat()

tasks=read('Tasks',COLUMNS)
for col in COLUMNS:
    if col not in tasks: tasks[col]=''
work,customer_tab,supplier_tab=st.tabs(['📋 任務看板','👥 客戶管理','🏭 廠商管理'])
with customer_tab: render_master('customer','bb')
with supplier_tab: render_master('supplier','bb')
with work:
    st.subheader('新增任務')
    customers=load_master('customer')
    names=[] if customers is None else [s(x) for x in customers.loc[customers['status'].astype(str)!='停用','company_name'].tolist() if s(x)]
    with st.form('new_task'):
        c1,c2=st.columns(2)
        with c1:
            title=st.text_input('任務名稱 *')
            customer=st.selectbox('客戶（共用主檔）',['']+sorted(set(names)))
            department=st.text_input('部門')
            owner=st.text_input('負責人')
        with c2:
            status=st.selectbox('狀態',STATES)
            due=st.date_input('到期日',value=date.today()+timedelta(days=3))
            next_step=st.text_input('下一步')
            evidence=st.text_input('證據／連結')
        if st.form_submit_button('➕ 建立任務',type='primary'):
            if not title.strip(): st.warning('請輸入任務名稱')
            else:
                record={c:'' for c in tasks.columns}
                record.update({'id':uuid.uuid4().hex[:12], 'title':title.strip(),'customer':customer,'department':department,'owner':owner,'status':status,'created_time':datetime.now().isoformat(timespec='seconds'),'updated_time':datetime.now().isoformat(timespec='seconds'),'due_time':due.isoformat(),'next_step':next_step,'evidence':evidence,'version':'V1'})
                save('Tasks',pd.concat([tasks,pd.DataFrame([record])],ignore_index=True))
    st.divider()
    search=st.text_input('🔎 搜尋任務、客戶、負責人、RFQ 編號')
    view=tasks.copy()
    if search:
        mask=view.astype(str).apply(lambda c:c.str.contains(search,case=False,regex=False,na=False)).any(axis=1)
        view=view.loc[mask]
    counts=tasks['status'].astype(str).value_counts()
    cols=st.columns(len(STATES))
    for col,state in zip(cols,STATES):
        with col:
            st.metric(state,int(counts.get(state,0)))
            subset=view.loc[view['status'].astype(str)==state]
            for idx,item in subset.iterrows():
                identifier=s(item.get('id')) or str(idx)
                with st.container(border=True):
                    st.markdown(f"**{s(item.get('title')) or '未命名任務'}**")
                    st.caption(f"客戶：{s(item.get('customer')) or '—'}｜負責：{s(item.get('owner')) or '—'}")
                    st.caption(f"到期：{s(item.get('due_time')) or '—'}")
                    with st.expander('編輯任務'):
                        with st.form(f'edit_{idx}_{identifier}'):
                            new_status=st.selectbox('狀態',STATES,index=STATES.index(state),key=f'st_{idx}')
                            new_owner=st.text_input('負責人',value=s(item.get('owner')))
                            new_step=st.text_input('下一步',value=s(item.get('next_step')))
                            new_evidence=st.text_input('證據',value=s(item.get('evidence')))
                            new_exception=st.text_input('異常',value=s(item.get('exception')))
                            approved=st.checkbox('核準報價',value=s(item.get('approved_quote')).lower() in ('true','1','yes'))
                            sent=st.checkbox('實際寄送',value=s(item.get('quote_sent')).lower() in ('true','1','yes'))
                            if st.form_submit_button('儲存修改'):
                                if new_status=='已報價' and not (approved and sent):
                                    st.error('未同時具備核準報價與實際寄送紀錄，不可標記已報價。')
                                else:
                                    for k,v in {'status':new_status,'owner':new_owner,'next_step':new_step,'evidence':new_evidence,'exception':new_exception,'approved_quote':approved,'quote_sent':sent,'updated_time':datetime.now().isoformat(timespec='seconds')}.items():
                                        tasks.at[idx,k]=v
                                    save('Tasks',tasks)
    st.divider()
    st.subheader('任務總覽')
    st.dataframe(view,use_container_width=True,hide_index=True)
