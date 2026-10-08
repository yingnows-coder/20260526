import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, date
import uuid
import io
import base64
from PIL import Image
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as PDFImage, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xml.sax.saxutils import escape

try:
    from streamlit_drawable_canvas import st_canvas
except ImportError:
    st_canvas = None

SIGNATURE_COLUMNS = ['customer_signature', 'signed_by', 'signed_at']

st.set_page_config(page_title='JENN-WEI 售服維修管理系統', page_icon='🛠️', layout='wide')
st.title('🛠️ JENN-WEI 售服維修管理系統')
st.caption('Service Management | 報修、派工、維修、驗收與售服績效')
# aa.py 獨立部署網址
AA_APP_URL = "https://aazzyyii.streamlit.app/"
BB_APP_URL = "https://aazzyybb.streamlit.app/"
nav_aa, nav_bb, nav_space = st.columns([1.5, 1.5, 4])
with nav_aa:
    st.link_button("📋 返回 RFQ 詢價管理", AA_APP_URL or "https://share.streamlit.io/", disabled=not AA_APP_URL, use_container_width=True)
with nav_bb:
    st.link_button("📊 返回 Trello 任務管理", BB_APP_URL, use_container_width=True)

SHEET = 'ServiceTickets'
COLUMNS = ['ticket_id','created_at','customer','contact','phone','machine_model','serial_number','warranty','issue','priority','status','engineer','scheduled_date','diagnosis','repair_action','parts','parts_cost','labor_hours','labor_rate','total_cost','resolution','closed_at','updated_at'] + SIGNATURE_COLUMNS
STATUSES = ['新報修','待確認','待派工','待零件','維修中','待驗收','已結案']
PRIORITIES = ['一般','急件','緊急']
# cc.py 是獨立部署的 App，必須在此 App 的 Secrets 指定試算表。
# 建議在 Streamlit Cloud > 售服 App > Settings > Secrets 設定：
# SERVICE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/你的ID/edit"
def get_spreadsheet_url():
    try:
        direct = st.secrets.get("SERVICE_SPREADSHEET_URL", "")
        if direct:
            return str(direct).strip()
        settings = st.secrets.get("connections", {})
        gsheets = settings.get("gsheets", {})
        return str(gsheets.get("spreadsheet", "") or gsheets.get("spreadsheet_url", "")).strip()
    except Exception:
        return ""

SPREADSHEET_URL = get_spreadsheet_url()
if not SPREADSHEET_URL:
    st.error("尚未設定售服系統的 Google Spreadsheet 網址。")
    st.info('請到售服 App 的 Streamlit Cloud → Settings → Secrets，新增 SERVICE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/你的試算表ID/edit"。並保留 gsheets 連線的服務帳號憑證。')
    st.stop()

conn = st.connection('gsheets', type=GSheetsConnection)

def read_tickets():
    try:
        data = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=SHEET, ttl=300)
        if data is None:
            data = pd.DataFrame()
        for c in COLUMNS:
            if c not in data.columns:
                data[c] = ''
        return data[COLUMNS].copy()
    except Exception as exc:
        st.error(f'無法讀取 {SHEET} 工作表。請檢查試算表網址、ServiceTickets 分頁及服務帳號的共用權限。')
        st.exception(exc)
        st.stop()

def save_tickets(data):
    try:
        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=SHEET, data=data[COLUMNS].fillna(''))
        st.cache_data.clear()
        st.success('資料已儲存')
        st.rerun()
    except Exception as exc:
        st.error('Google Sheets 寫入失敗')
        st.exception(exc)

def number(value):
    try:
        result = pd.to_numeric(value, errors='coerce')
        return 0.0 if pd.isna(result) else float(result)
    except (TypeError, ValueError):
        return 0.0

def safe_str(value):
    return '' if pd.isna(value) else str(value)

def signature_to_base64(canvas_data):
    """縮小並轉為 PNG；避免超過 Google Sheets 單格 50,000 字元上限。"""
    if canvas_data is None:
        return ''
    image = Image.fromarray(canvas_data.astype('uint8'), 'RGBA')
    image.thumbnail((550, 190))
    background = Image.new('RGB', image.size, 'white')
    background.paste(image, mask=image.getchannel('A'))
    background = background.convert('L').point(lambda x: 255 if x > 195 else 0, '1')
    output = io.BytesIO()
    background.save(output, format='PNG', optimize=True)
    encoded = base64.b64encode(output.getvalue()).decode('ascii')
    if len(encoded) > 48000:
        raise ValueError('簽名圖片超過 Google Sheets 單格限制，請清除後簡化簽名。')
    return encoded


def make_ticket_pdf(ticket):
    """維修服務單 PDF，含簽名及中文欄位。"""
    font_path = next((path for path in [
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf',
    ] if __import__('os').path.exists(path)), None)
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont('JWChinese', font_path, subfontIndex=0))
            font_name = 'JWChinese'
        except Exception:
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            pdfmetrics.registerFont(UnicodeCIDFont('MSung-Light'))
            font_name = 'MSung-Light'
    else:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont('MSung-Light'))
        font_name = 'MSung-Light'
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=(210*mm, 297*mm),
                            rightMargin=17*mm, leftMargin=17*mm,
                            topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='JWTitle', fontName=font_name, fontSize=17, leading=25, alignment=TA_CENTER, spaceAfter=12))
    styles.add(ParagraphStyle(name='JWCell', fontName=font_name, fontSize=9, leading=15, wordWrap='CJK'))
    styles.add(ParagraphStyle(name='JWSmall', fontName=font_name, fontSize=8, leading=12))
    def para(value):
        return Paragraph(escape(safe_str(value)).replace('\\n','<br/>'), styles['JWCell'])
    story = [Paragraph('JENN-WEI 震唯機械｜售服維修服務單', styles['JWTitle'])]
    fields = [
        ('維修單號', 'ticket_id'), ('建立日期', 'created_at'),
        ('客戶名稱', 'customer'), ('聯絡人', 'contact'),
        ('聯絡電話', 'phone'), ('機台型號', 'machine_model'),
        ('機台序號', 'serial_number'), ('保固狀態', 'warranty'),
        ('優先程度', 'priority'), ('維修狀態', 'status'),
        ('負責工程師', 'engineer'), ('預定處理日期', 'scheduled_date'),
        ('故障現象', 'issue'), ('故障診斷', 'diagnosis'),
        ('維修處置', 'repair_action'), ('更換零件', 'parts'),
        ('測試／客戶驗收', 'resolution'),
        ('零件費用', 'parts_cost'), ('維修工時', 'labor_hours'),
        ('每小時工資', 'labor_rate'), ('維修費用合計', 'total_cost'),
        ('結案日期', 'closed_at')
    ]
    rows = [[para(label), para(ticket.get(key, ''))] for label, key in fields]
    table = Table(rows, colWidths=[38*mm, 136*mm], hAlign='LEFT')
    table.setStyle(TableStyle([
        ('FONTNAME', (0,0),(-1,-1),font_name),
        ('VALIGN', (0,0),(-1,-1),'TOP'),
        ('BACKGROUND',(0,0),(0,-1),colors.HexColor('#EDF2F8')),
        ('GRID',(0,0),(-1,-1),0.35,colors.HexColor('#D6DEE8')),
        ('LEFTPADDING',(0,0),(-1,-1),8),
        ('RIGHTPADDING',(0,0),(-1,-1),8),
        ('TOPPADDING',(0,0),(-1,-1),5),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
    ]))
    story += [table, Spacer(1, 10*mm)]
    signed_by = safe_str(ticket.get('signed_by',''))
    signed_at = safe_str(ticket.get('signed_at',''))
    signature = safe_str(ticket.get('customer_signature',''))
    sign_items = [Paragraph('客戶簽收／驗收確認', styles['JWCell'])]
    if signature:
        try:
            sign_image = PDFImage(io.BytesIO(base64.b64decode(signature)))
            sign_image.drawWidth = 65*mm
            sign_image.drawHeight = 22*mm
            sign_items.append(sign_image)
        except Exception:
            sign_items.append(Paragraph('簽名圖片無法載入', styles['JWSmall']))
    else:
        sign_items.append(Spacer(1, 22*mm))
    sign_items.append(Paragraph(f'簽收人：{escape(signed_by)}　簽收時間：{escape(signed_at)}', styles['JWSmall']))
    sign_items.append(Paragraph('本單據記錄現場維修及簽收資訊，簽收不代表另行承諾保固或費用條件。', styles['JWSmall']))
    story.append(KeepTogether(sign_items))
    doc.build(story)
    return output.getvalue()


df = read_tickets()

new_tab, manage_tab, dashboard_tab, history_tab, signature_tab = st.tabs(['📝 建立報修單','🔧 維修工單管理','📊 售服 KPI','📚 維修紀錄','✍️ 客戶簽名／PDF'])
with new_tab:
    st.subheader('建立客戶報修工單')
    with st.form('new_service_ticket', clear_on_submit=True):
        a,b,c = st.columns(3)
        with a:
            customer = st.text_input('客戶名稱 *')
            contact = st.text_input('聯絡人')
            phone = st.text_input('電話 / 聯絡方式')
        with b:
            model = st.text_input('機台型號 *')
            serial = st.text_input('機台序號')
            warranty = st.selectbox('保固狀態', ['待確認','保固內','保固外'])
        with c:
            priority = st.selectbox('緊急程度', PRIORITIES)
            engineer = st.text_input('預定負責工程師')
            schedule = st.date_input('預計處理日期', value=date.today())
        issue = st.text_area('故障現象 / 客戶反映 *')
        submitted = st.form_submit_button('➕ 建立維修工單', type='primary')
    if submitted:
        if not customer.strip() or not model.strip() or not issue.strip():
            st.error('請填寫客戶名稱、機台型號及故障現象。')
        else:
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ticket = {c:'' for c in COLUMNS}
            ticket.update(ticket_id=f'SRV-{datetime.now():%Y%m%d}-{uuid.uuid4().hex[:4].upper()}',created_at=now,customer=customer.strip(),contact=contact.strip(),phone=phone.strip(),machine_model=model.strip(),serial_number=serial.strip(),warranty=warranty,issue=issue.strip(),priority=priority,status='新報修',engineer=engineer.strip(),scheduled_date=str(schedule),parts_cost=0.0,labor_hours=0.0,labor_rate=0.0,total_cost=0.0,updated_at=now)
            save_tickets(pd.concat([df,pd.DataFrame([ticket])],ignore_index=True))

with manage_tab:
    st.subheader('維修工單管理')
    f1,f2,f3 = st.columns(3)
    with f1: status_filter = st.selectbox('維修狀態', ['全部']+STATUSES)
    with f2: engineer_filter = st.selectbox('負責工程師', ['全部']+sorted(x for x in df['engineer'].dropna().astype(str).unique() if x.strip()))
    with f3: keyword = st.text_input('搜尋工單 / 客戶 / 型號 / 序號')
    view = df.copy()
    if status_filter != '全部': view = view[view['status']==status_filter]
    if engineer_filter != '全部': view = view[view['engineer']==engineer_filter]
    if keyword: view = view[view.astype(str).apply(lambda s:s.str.contains(keyword,case=False,regex=False,na=False)).any(axis=1)]
    st.caption(f'符合條件工單：{len(view)} 筆')
    if view.empty:
        st.info('目前沒有符合條件的維修工單。')
    else:
        st.dataframe(view[['ticket_id','customer','machine_model','serial_number','priority','status','engineer','scheduled_date','updated_at']],hide_index=True,use_container_width=True)
        ticket_id = st.selectbox('選擇要更新的工單',view['ticket_id'].astype(str).tolist())
        matches = df.index[df['ticket_id'].astype(str)==ticket_id]
        if len(matches):
            idx=matches[0]; item=df.loc[idx]
            with st.form('edit_service_ticket'):
                st.markdown(f'**{ticket_id}｜{safe_str(item["customer"])}｜{safe_str(item["machine_model"])}**')
                c1,c2,c3=st.columns(3)
                with c1:
                    current=safe_str(item['status']); status=st.selectbox('維修狀態',STATUSES,index=STATUSES.index(current) if current in STATUSES else 0)
                    priority=safe_str(item['priority']); priority_new=st.selectbox('優先程度',PRIORITIES,index=PRIORITIES.index(priority) if priority in PRIORITIES else 0)
                with c2:
                    engineer_new=st.text_input('負責工程師',value=safe_str(item['engineer']))
                    warranty=safe_str(item['warranty']); warranty_new=st.selectbox('保固', ['待確認','保固內','保固外'],index=['待確認','保固內','保固外'].index(warranty) if warranty in ['待確認','保固內','保固外'] else 0)
                with c3:
                    parts_cost=st.number_input('零件費用',min_value=0.0,value=number(item['parts_cost']))
                    labor_hours=st.number_input('維修工時（小時）',min_value=0.0,value=number(item['labor_hours']))
                    labor_rate=st.number_input('每小時工資',min_value=0.0,value=number(item['labor_rate']))
                diagnosis=st.text_area('故障診斷',value=safe_str(item['diagnosis']))
                action=st.text_area('維修處置',value=safe_str(item['repair_action']))
                parts=st.text_area('更換零件',value=safe_str(item['parts']))
                resolution=st.text_area('測試 / 客戶驗收紀錄',value=safe_str(item['resolution']))
                st.metric('本次維修費用合計',f'{parts_cost+labor_hours*labor_rate:,.0f}')
                saved=st.form_submit_button('💾 儲存維修紀錄',type='primary')
            if saved:
                updates=dict(status=status,priority=priority_new,engineer=engineer_new.strip(),warranty=warranty_new,parts_cost=parts_cost,labor_hours=labor_hours,labor_rate=labor_rate,total_cost=parts_cost+labor_hours*labor_rate,diagnosis=diagnosis,repair_action=action,parts=parts,resolution=resolution,updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                if status=='已結案' and not safe_str(item['closed_at']): updates['closed_at']=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                if status!='已結案': updates['closed_at']=''
                for k,v in updates.items(): df.at[idx,k]=v
                save_tickets(df)

with dashboard_tab:
    st.subheader('售服管理 KPI')
    total=len(df); closed=int((df['status']=='已結案').sum()); active=total-closed; urgent=int(df['priority'].isin(['急件','緊急']).sum())
    a,b,c,d=st.columns(4)
    a.metric('累計報修',total); b.metric('處理中',active); c.metric('已結案',closed); d.metric('急件 / 緊急',urgent)
    st.markdown('#### 維修狀態分布')
    st.bar_chart(df['status'].value_counts().reindex(STATUSES,fill_value=0))
    st.markdown('#### 工程師工單數')
    if not df.empty:
        st.bar_chart(df['engineer'].replace('', '未指派').fillna('未指派').value_counts())
    st.metric('累計維修費用',f'{pd.to_numeric(df["total_cost"],errors="coerce").fillna(0).sum():,.0f}')

with history_tab:
    st.subheader('維修紀錄查詢 / 匯出')
    st.dataframe(df.sort_values('created_at',ascending=False),hide_index=True,use_container_width=True)
    st.download_button('⬇️ 匯出維修工單 CSV',data=df.to_csv(index=False).encode('utf-8-sig'),file_name='service_tickets.csv',mime='text/csv')

with signature_tab:
    st.subheader('✍️ 客戶手機簽名與維修單 PDF')
    st.caption('手機可用手指直接在簽名區書寫。簽名會儲存至同一份 Google Sheets 的 ServiceTickets 分頁。')
    if df.empty:
        st.info('請先建立維修工單。')
    else:
        selected_id = st.selectbox('選擇維修單號', df['ticket_id'].astype(str).tolist(), key='sign_ticket')
        selected_idx = df.index[df['ticket_id'].astype(str) == selected_id][0]
        ticket = df.loc[selected_idx]
        st.write(f'**客戶：** {safe_str(ticket["customer"])}　**機台：** {safe_str(ticket["machine_model"])}')
        signed = bool(safe_str(ticket.get('customer_signature', '')))
        if signed:
            st.success(f'已簽收：{safe_str(ticket.get("signed_by", ""))}　{safe_str(ticket.get("signed_at", ""))}')
            try:
                st.image(base64.b64decode(safe_str(ticket['customer_signature'])), width=440)
            except Exception:
                st.warning('已存簽名資料，但無法顯示預覽。')
        else:
            st.info('本工單尚未簽名。')
        signer = st.text_input('客戶簽收人姓名', value=safe_str(ticket.get('signed_by', '')), key=f'signer_{selected_id}')
        if st_canvas is None:
            st.error('尚未安裝手機簽名元件。請在 requirements.txt 加入 streamlit-drawable-canvas。')
        else:
            st.caption('請在下方白色區域簽名。需要重簽時請按畫布的清除鍵。')
            canvas = st_canvas(
                fill_color='rgba(255,255,255,0)',
                stroke_width=3,
                stroke_color='#17263C',
                background_color='#FFFFFF',
                update_streamlit=True,
                height=190,
                width=550,
                drawing_mode='freedraw',
                key=f'canvas_{selected_id}',
            )
            if st.button('💾 儲存客戶簽名', type='primary', key=f'save_signature_{selected_id}'):
                if not signer.strip():
                    st.error('請先輸入簽收人姓名。')
                elif canvas.image_data is None or not canvas.json_data or not canvas.json_data.get('objects'):
                    st.error('請先在簽名區書寫。')
                else:
                    try:
                        encoded = signature_to_base64(canvas.image_data)
                        df.at[selected_idx, 'customer_signature'] = encoded
                        df.at[selected_idx, 'signed_by'] = signer.strip()
                        df.at[selected_idx, 'signed_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        df.at[selected_idx, 'updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        save_tickets(df)
                    except Exception as exc:
                        st.error(f'簽名儲存失敗：{exc}')
        st.divider()
        st.download_button(
            '📄 下載維修服務單 PDF（含已儲存簽名）',
            data=make_ticket_pdf(ticket),
            file_name=f'{selected_id}_service_report.pdf',
            mime='application/pdf',
            use_container_width=True,
        )
        st.caption('先儲存簽名，頁面重新載入後再下載 PDF，才會包含最新簽名。')

with st.expander('⚙️ Google Sheets 工作表設定'):
    st.write(f'請在與 aa.py 相同的 Google Spreadsheet 中新增工作表 **{SHEET}**，第一列依序建立以下欄位：')
    st.code(','.join(COLUMNS))
    st.info('新增的 3 欄為 customer_signature、signed_by、signed_at，請加在 ServiceTickets 工作表第一列原有欄位後方。')
    st.caption('讀取快取 TTL 為 300 秒；寫入後清除 Streamlit 資料快取。')
