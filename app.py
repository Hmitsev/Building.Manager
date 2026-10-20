import io
import os
from datetime import date
from decimal import Decimal
import pandas as pd
from PIL import Image
import streamlit as st
from auth import hash_password, verify_password, login_user, logout_user, current_user, has_role
from config import APP_NAME, CURRENCY, IBAN, PAYEE
from database import init_db, fetch_all, fetch_one, execute
from services import MONTHS, create_year_obligations, obligation_rows, create_payment, audit

st.set_page_config(page_title=APP_NAME, page_icon='🏢', layout='wide', initial_sidebar_state='expanded')
st.markdown('''<style>.block-container{padding-top:1.2rem;max-width:1250px}.bm-card{background:white;border:1px solid #e2e8f0;border-radius:16px;padding:18px;box-shadow:0 4px 14px rgba(15,23,42,.05)}.status-paid{color:#15803d;font-weight:700}.status-due{color:#b91c1c;font-weight:700}.small{color:#64748b;font-size:.88rem}div.stButton>button{border-radius:10px;font-weight:650}</style>''', unsafe_allow_html=True)

try:
    init_db()
except Exception as exc:
    st.error(str(exc)); st.info('В Render добавете DATABASE_URL, ADMIN_EMAIL и ADMIN_PASSWORD в Environment.'); st.stop()


def fmt(value): return f"{Decimal(value or 0):,.2f} {CURRENCY}".replace(',', ' ')
def rerun(): st.rerun()

def login_page():
    st.title('🏢 Building.Manager')
    tab1,tab2=st.tabs(['Вход','Регистрация'])
    with tab1:
        with st.form('login'):
            email=st.text_input('Имейл').strip().lower(); password=st.text_input('Парола',type='password'); submit=st.form_submit_button('Вход',use_container_width=True)
        if submit:
            u=fetch_one("SELECT u.*,a.apartment_number FROM users u LEFT JOIN apartments a ON a.id=u.apartment_id WHERE lower(u.email)=lower(%s)",(email,))
            if u and verify_password(password,u['password_hash']):
                if u['account_status']!='approved': st.warning('Регистрацията още не е одобрена.')
                else: login_user(dict(u)); rerun()
            else: st.error('Невалиден имейл или парола.')
    with tab2:
        apartments=fetch_all("SELECT id,apartment_number FROM apartments WHERE active ORDER BY apartment_number")
        if not apartments: st.info('Администраторът първо трябва да добави апартаменти.')
        with st.form('register'):
            apt_map={a['apartment_number']:a['id'] for a in apartments}; apt=st.selectbox('Апартамент',list(apt_map) or ['Няма добавени'])
            c1,c2=st.columns(2); first=c1.text_input('Име'); last=c2.text_input('Фамилия')
            email=st.text_input('Имейл').strip().lower(); phone=st.text_input('Телефон')
            password=st.text_input('Парола',type='password'); password2=st.text_input('Повтори паролата',type='password')
            reg=st.form_submit_button('Изпрати регистрация',use_container_width=True)
        if reg:
            if not apartments or not all([first,last,email,password]): st.error('Попълнете задължителните полета.')
            elif password!=password2 or len(password)<8: st.error('Паролите трябва да съвпадат и да са поне 8 знака.')
            else:
                try:
                    execute("INSERT INTO users(email,password_hash,first_name,last_name,phone,role,apartment_id,account_status) VALUES(%s,%s,%s,%s,%s,'neighbor',%s,'pending')",(email,hash_password(password),first,last,phone,apt_map[apt])); st.success('Регистрацията е изпратена за одобрение.')
                except Exception: st.error('Имейлът вече е използван или данните са невалидни.')

def banner():
    a=fetch_one("SELECT * FROM announcements WHERE active AND (starts_on IS NULL OR starts_on<=CURRENT_DATE) AND (ends_on IS NULL OR ends_on>=CURRENT_DATE) ORDER BY created_at DESC LIMIT 1")
    if a:
        if a['image_data']: st.image(bytes(a['image_data']),use_container_width=True)
        st.subheader(a['title']); st.write(a['body'] or '')
    else:
        st.markdown("<div class='bm-card'><h2>Добре дошли в Building.Manager</h2><p>Такси, плащания, разходи и информация за входа на едно място.</p></div>",unsafe_allow_html=True)

def home(user):
    banner(); st.write('')
    if user.get('apartment_id'):
        rows=fetch_all("""SELECT o.amount,COALESCE((SELECT SUM(pa.amount) FROM payment_allocations pa JOIN payments p ON p.id=pa.payment_id WHERE pa.obligation_id=o.id AND p.status='confirmed'),0) paid FROM obligations o WHERE o.apartment_id=%s""",(user['apartment_id'],))
        due=sum(Decimal(r['amount'])-Decimal(r['paid']) for r in rows); paid=sum(Decimal(r['paid']) for r in rows)
        c1,c2,c3=st.columns(3); c1.metric('Апартамент',user.get('apartment_number') or '-'); c2.metric('Платено',fmt(paid)); c3.metric('Остатък',fmt(due))
    st.subheader('Последни разходи')
    st.dataframe(fetch_all("SELECT expense_date AS дата,category AS категория,amount AS сума,source AS каса,description AS описание FROM expenses ORDER BY expense_date DESC,id DESC LIMIT 10"),use_container_width=True,hide_index=True)

def fees_page(user):
    if not user.get('apartment_id'): st.info('Този профил не е свързан с апартамент.'); return
    year=st.selectbox('Година',list(range(date.today().year-2,date.today().year+2)),index=2)
    create_year_obligations(user['apartment_id'],year); rows=obligation_rows(user['apartment_id'],year)
    selected=[]
    for fee_type,label in [('entrance','Такса вход'),('fro','Фонд ремонт и обновяване')]:
        st.subheader(label); cols=st.columns(4)
        for i,r in enumerate([x for x in rows if x['fee_type']==fee_type]):
            remaining=Decimal(r['amount'])-Decimal(r['paid'])
            with cols[i%4]:
                st.markdown(f"**{MONTHS[r['month']]}**  \\nНачислено: {fmt(r['amount'])}  \\nПлатено: {fmt(r['paid'])}")
                if remaining>0 and st.checkbox(f"Избери {fmt(remaining)}",key=f"ob_{r['id']}"): selected.append({'id':r['id'],'amount':remaining,'label':f"{label} - {MONTHS[r['month']]} {year}"})
                elif remaining<=0: st.success('Платено')
    if selected:
        st.divider(); st.subheader('Кошница')
        for x in selected: st.write(f"• {x['label']}: **{fmt(x['amount'])}**")
        st.metric('Общо',fmt(sum(x['amount'] for x in selected)))
        if st.button('Създай банково плащане',type='primary',use_container_width=True):
            num=create_payment(user['apartment_id'],selected,'bank',user['id'],'pending','Създадено от клиента')
            st.success(f'Заявката е създадена. Основание: {num}'); st.write(f'Получател: **{PAYEE}**'); st.write(f'IBAN: **{IBAN or "ще бъде предоставен от домоуправителя"}**')

def archive_page(user):
    if not user.get('apartment_id'): return
    rows=fetch_all("SELECT payment_number AS номер,created_at AS създадено,payment_date AS дата,amount AS сума,payment_method AS метод,status AS статус,note AS бележка FROM payments WHERE apartment_id=%s ORDER BY created_at DESC",(user['apartment_id'],))
    st.subheader('Архив на плащанията'); st.dataframe(rows,use_container_width=True,hide_index=True)

def expenses_page():
    st.subheader('Разходи на входа'); rows=fetch_all("SELECT expense_date AS дата,category AS категория,amount AS сума,source AS каса,description AS описание FROM expenses ORDER BY expense_date DESC,id DESC")
    st.dataframe(rows,use_container_width=True,hide_index=True)

def profile_page(user):
    st.subheader('Моят профил'); st.write(f"**{user['first_name']} {user['last_name']}**"); st.write(user['email']); st.write(f"Апартамент: **{user.get('apartment_number') or '-'}**")
    if user.get('apartment_id'):
        apt=fetch_one('SELECT * FROM apartments WHERE id=%s',(user['apartment_id'],)); st.write(f"Текущ брой пълнолетни: **{apt['adult_residents']}**")
        with st.form('change_request'):
            adults=st.number_input('Нов брой лица над 18 г.',0,20,int(apt['adult_residents'])); effective=st.date_input('В сила от',date.today().replace(day=1)); reason=st.text_area('Причина'); submit=st.form_submit_button('Изпрати искане')
        if submit:
            execute("INSERT INTO change_requests(apartment_id,requested_adults,effective_from,reason,requested_by) VALUES(%s,%s,%s,%s,%s)",(user['apartment_id'],adults,effective,reason,user['id'])); st.success('Искането е изпратено.')

def admin_page(user):
    st.title('Административен панел')
    tabs=st.tabs(['Табло','Апартаменти','Регистрации','Плащания','Разходи','Банер','Промени','Експорт'])
    with tabs[0]:
        stats=fetch_one("SELECT (SELECT COUNT(*) FROM apartments WHERE active) apartments,(SELECT COUNT(*) FROM users WHERE account_status='pending') pending,(SELECT COALESCE(SUM(amount),0) FROM payments WHERE status='confirmed') income,(SELECT COALESCE(SUM(amount),0) FROM expenses) expenses")
        c1,c2,c3,c4=st.columns(4); c1.metric('Апартаменти',stats['apartments']); c2.metric('Чакащи регистрации',stats['pending']); c3.metric('Постъпления',fmt(stats['income'])); c4.metric('Разходи',fmt(stats['expenses']))
    with tabs[1]:
        with st.form('add_apt'):
            c1,c2,c3=st.columns(3); num=c1.text_input('Номер/обект'); typ=c2.selectbox('Тип',['Апартамент','Офис','Магазин','Гараж','Паркомясто','Друг']); floor=c3.text_input('Етаж')
            c4,c5=st.columns(2); adults=c4.number_input('Лица над 18 г.',0,20,1); fro=c5.number_input('Годишен ФРО (€)',0.0,10000.0,0.0,step=0.01)
            notes=st.text_area('Бележка'); add=st.form_submit_button('Добави обект')
        if add and num:
            try:
                row=execute("INSERT INTO apartments(apartment_number,property_type,floor,adult_residents,annual_fro,notes) VALUES(%s,%s,%s,%s,%s,%s) RETURNING id",(num,typ,floor,adults,fro,notes),True)
                execute("INSERT INTO resident_history(apartment_id,adults,effective_from,reason,created_by) VALUES(%s,%s,%s,'Начална стойност',%s)",(row['id'],adults,date(date.today().year,1,1),user['id'])); st.success('Обектът е добавен.'); rerun()
            except Exception: st.error('Този номер/обект вече съществува.')
        st.dataframe(fetch_all("SELECT apartment_number AS обект,property_type AS тип,floor AS етаж,adult_residents AS пълнолетни,annual_fro AS годишен_ФРО,status AS статус FROM apartments ORDER BY apartment_number"),use_container_width=True,hide_index=True)
    with tabs[2]:
        regs=fetch_all("SELECT u.id,u.first_name,u.last_name,u.email,a.apartment_number FROM users u LEFT JOIN apartments a ON a.id=u.apartment_id WHERE u.account_status='pending' ORDER BY u.created_at")
        for r in regs:
            c1,c2,c3=st.columns([4,1,1]); c1.write(f"**{r['first_name']} {r['last_name']}** | Ап. {r['apartment_number']} | {r['email']}")
            if c2.button('Одобри',key=f"approve{r['id']}"): execute("UPDATE users SET account_status='approved' WHERE id=%s",(r['id'],)); audit(user['id'],'approve_user','user',r['id']); rerun()
            if c3.button('Откажи',key=f"reject{r['id']}"): execute("UPDATE users SET account_status='rejected' WHERE id=%s",(r['id'],)); rerun()
    with tabs[3]:
        pending=fetch_all("SELECT p.id,p.payment_number,p.amount,p.created_at,a.apartment_number FROM payments p JOIN apartments a ON a.id=p.apartment_id WHERE p.status='pending' ORDER BY p.created_at")
        for p in pending:
            c1,c2=st.columns([5,1]); c1.write(f"**{p['payment_number']}** | Ап. {p['apartment_number']} | {fmt(p['amount'])}")
            if c2.button('Потвърди',key=f"pay{p['id']}"): execute("UPDATE payments SET status='confirmed',payment_date=CURRENT_DATE,confirmed_by=%s,confirmed_at=NOW() WHERE id=%s",(user['id'],p['id'])); audit(user['id'],'confirm_payment','payment',p['id']); rerun()
        st.divider(); st.caption('За ръчно плащане първо генерирайте годишните задължения от клиентския профил или чрез избор на апартамент.')
        apts=fetch_all("SELECT id,apartment_number FROM apartments WHERE active ORDER BY apartment_number"); amap={a['apartment_number']:a['id'] for a in apts}
        if apts:
            apt_num=st.selectbox('Апартамент за ръчно плащане',list(amap)); yr=st.number_input('Година',2022,2100,date.today().year); create_year_obligations(amap[apt_num],int(yr)); dues=[r for r in obligation_rows(amap[apt_num],int(yr)) if Decimal(r['amount'])>Decimal(r['paid'])]
            labels={f"{('Такса вход' if r['fee_type']=='entrance' else 'ФРО')} {MONTHS[r['month']]} - {fmt(Decimal(r['amount'])-Decimal(r['paid']))}":r for r in dues}
            picked=st.multiselect('Покрити месеци',list(labels)); method=st.selectbox('Метод',['cash','bank']); note=st.text_input('Бележка')
            if st.button('Запиши ръчно плащане') and picked:
                items=[{'id':labels[x]['id'],'amount':Decimal(labels[x]['amount'])-Decimal(labels[x]['paid'])} for x in picked]; create_payment(amap[apt_num],items,method,user['id'],'confirmed',note); st.success('Плащането е записано.'); rerun()
    with tabs[4]:
        with st.form('expense'):
            d=st.date_input('Дата',date.today()); category=st.selectbox('Категория',['Електроенергия','Асансьор','Почистване','Интернет','Банкови такси','Домоуправител','Касиер','Ремонт','Консумативи','Гараж','Други']); amount=st.number_input('Сума (€)',0.0,100000.0,0.0,step=0.01); source=st.selectbox('Каса',['cash','bank']); description=st.text_area('Описание'); save=st.form_submit_button('Добави разход')
        if save and amount>0: execute("INSERT INTO expenses(expense_date,category,amount,source,description,created_by) VALUES(%s,%s,%s,%s,%s,%s)",(d,category,amount,source,description,user['id'])); st.success('Разходът е добавен.'); rerun()
        st.dataframe(fetch_all("SELECT expense_date,category,amount,source,description FROM expenses ORDER BY expense_date DESC,id DESC"),use_container_width=True,hide_index=True)
    with tabs[5]:
        with st.form('banner'):
            title=st.text_input('Заглавие'); body=st.text_area('Текст'); image=st.file_uploader('Снимка',type=['png','jpg','jpeg','webp']); starts=st.date_input('Начална дата',date.today()); ends=st.date_input('Крайна дата',value=None); publish=st.form_submit_button('Публикувай')
        if publish and title:
            data=mime=None
            if image:
                im=Image.open(image).convert('RGB'); im.thumbnail((1600,700)); buf=io.BytesIO(); im.save(buf,format='WEBP',quality=82,method=6); data=buf.getvalue(); mime='image/webp'
            execute("UPDATE announcements SET active=FALSE"); execute("INSERT INTO announcements(title,body,image_data,image_mime,starts_on,ends_on,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s)",(title,body,data,mime,starts,ends,user['id'])); st.success('Банерът е публикуван.'); rerun()
    with tabs[6]:
        reqs=fetch_all("SELECT cr.*,a.apartment_number,u.first_name,u.last_name FROM change_requests cr JOIN apartments a ON a.id=cr.apartment_id LEFT JOIN users u ON u.id=cr.requested_by WHERE cr.status='pending' ORDER BY cr.created_at")
        for r in reqs:
            st.write(f"Ап. **{r['apartment_number']}**: {r['requested_adults']} лица от {r['effective_from']} | {r['reason'] or ''}")
            c1,c2=st.columns(2)
            if c1.button('Одобри промяната',key=f"chga{r['id']}"):
                execute("UPDATE change_requests SET status='approved',reviewed_by=%s,reviewed_at=NOW() WHERE id=%s",(user['id'],r['id'])); execute("UPDATE apartments SET adult_residents=%s WHERE id=%s",(r['requested_adults'],r['apartment_id'])); execute("INSERT INTO resident_history(apartment_id,adults,effective_from,reason,created_by) VALUES(%s,%s,%s,%s,%s) ON CONFLICT(apartment_id,effective_from) DO UPDATE SET adults=EXCLUDED.adults,reason=EXCLUDED.reason",(r['apartment_id'],r['requested_adults'],r['effective_from'],r['reason'],user['id'])); rerun()
            if c2.button('Откажи промяната',key=f"chgr{r['id']}"): execute("UPDATE change_requests SET status='rejected',reviewed_by=%s,reviewed_at=NOW() WHERE id=%s",(user['id'],r['id'])); rerun()
    with tabs[7]:
        dataset=fetch_all("SELECT a.apartment_number,o.fee_type,o.year,o.month,o.amount,COALESCE(SUM(CASE WHEN p.status='confirmed' THEN pa.amount ELSE 0 END),0) paid FROM obligations o JOIN apartments a ON a.id=o.apartment_id LEFT JOIN payment_allocations pa ON pa.obligation_id=o.id LEFT JOIN payments p ON p.id=pa.payment_id GROUP BY a.apartment_number,o.id ORDER BY a.apartment_number,o.year,o.month")
        df=pd.DataFrame(dataset); buf=io.BytesIO(); df.to_excel(buf,index=False,engine='openpyxl'); st.download_button('Изтегли Excel отчет',buf.getvalue(),'Building_Manager_Report.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

user=current_user()
if not user: login_page(); st.stop()
with st.sidebar:
    st.title('🏢 Building.Manager'); st.caption(f"{user['first_name']} {user['last_name']} | {user['role']}")
    choices=['Начало','Моите такси','Архив','Разходи','Профил']
    if has_role('manager','cashier'): choices.append('Админ панел')
    page=st.radio('Навигация',choices)
    if st.button('Изход',use_container_width=True): logout_user(); rerun()
if page=='Начало': home(user)
elif page=='Моите такси': fees_page(user)
elif page=='Архив': archive_page(user)
elif page=='Разходи': expenses_page()
elif page=='Профил': profile_page(user)
elif page=='Админ панел': admin_page(user)
