import io
import os
from datetime import date
from decimal import Decimal
from urllib.parse import quote

import pandas as pd
from PIL import Image
import streamlit as st

from auth import hash_password, verify_password, login_user, logout_user, current_user
from config import APP_NAME, CURRENCY
from database import init_db, fetch_all, fetch_one, execute
from services import MONTHS, create_year_obligations, obligation_rows, create_payment, audit

st.set_page_config(page_title=APP_NAME, page_icon="🏢", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
:root{--line:rgba(148,163,184,.16);--text:#f8fafc}.stApp{color:var(--text);background:radial-gradient(circle at 85% -10%,rgba(37,99,235,.18),transparent 30%),linear-gradient(180deg,#070b14,#09111f)}
.block-container{padding-top:1.3rem;padding-bottom:3rem;max-width:1320px}h1,h2,h3{color:#f8fafc!important;letter-spacing:-.035em;font-weight:800!important}p,label,.stMarkdown{color:#e2e8f0}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#111827,#0b1220);border-right:1px solid var(--line)}section[data-testid="stSidebar"] div[role="radiogroup"] label{border:1px solid transparent;border-radius:12px;padding:8px 10px;margin:2px 0}section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked){background:linear-gradient(135deg,rgba(37,99,235,.30),rgba(8,145,178,.16));border-color:rgba(96,165,250,.38)}
div[data-testid="stMetric"]{position:relative;overflow:hidden;background:linear-gradient(145deg,rgba(25,38,61,.98),rgba(13,22,38,.98))!important;border:1px solid rgba(96,165,250,.18);border-radius:20px;padding:20px;min-height:120px;box-shadow:0 16px 38px rgba(0,0,0,.23)}div[data-testid="stMetric"]:before{content:"";position:absolute;left:0;top:0;right:0;height:3px;background:linear-gradient(90deg,#2563eb,#22d3ee)}[data-testid="stMetricLabel"] p{color:#9fb0c7!important;font-weight:650!important}[data-testid="stMetricValue"]{color:#f8fafc!important;font-size:1.9rem!important;font-weight:800!important}
.stTabs [data-baseweb="tab-list"]{gap:7px;background:rgba(15,23,38,.75);border:1px solid var(--line);border-radius:15px;padding:6px;overflow-x:auto}.stTabs [data-baseweb="tab"]{height:42px;border-radius:10px;padding:0 14px;color:#aebdd0}.stTabs [aria-selected="true"]{color:white!important;background:linear-gradient(135deg,#2563eb,#0891b2)!important}.stTabs [data-baseweb="tab-highlight"]{display:none}
div.stButton>button,div.stDownloadButton>button,a[data-testid="stLinkButton"]{color:#f8fafc!important;background:linear-gradient(135deg,#2563eb,#0891b2)!important;border:1px solid rgba(125,211,252,.24)!important;border-radius:12px!important;min-height:44px;font-weight:750!important;box-shadow:0 10px 24px rgba(37,99,235,.22)}div.stButton>button:hover{filter:brightness(1.1);transform:translateY(-1px)}section[data-testid="stSidebar"] div.stButton>button{background:rgba(255,255,255,.04)!important;box-shadow:none}
[data-baseweb="input"]>div,[data-baseweb="select"]>div,[data-baseweb="textarea"]>div,[data-testid="stNumberInput"]>div>div,[data-testid="stDateInput"]>div>div{background:#111b2c!important;border-color:rgba(148,163,184,.2)!important;border-radius:12px!important;color:#f8fafc!important}input,textarea{color:#f8fafc!important}[role="listbox"]{background:#111b2c!important;color:white!important}
[data-testid="stForm"]{background:linear-gradient(145deg,rgba(18,28,46,.94),rgba(11,18,31,.94));border:1px solid var(--line);border-radius:18px;padding:20px}[data-testid="stDataFrame"],[data-testid="stTable"]{border:1px solid var(--line);border-radius:16px;overflow:hidden}
@media(max-width:768px){.block-container{padding:1rem .8rem 2rem}h1{font-size:2rem!important}[data-testid="stMetricValue"]{font-size:1.45rem!important}div[data-testid="stMetric"]{min-height:100px;padding:15px}}
</style>
""", unsafe_allow_html=True)

try:
    init_db()
    execute("ALTER TABLE payments ADD COLUMN IF NOT EXISTS provider VARCHAR(50)")
    execute("ALTER TABLE payments ADD COLUMN IF NOT EXISTS provider_transaction_id VARCHAR(150)")
    execute("ALTER TABLE payments ADD COLUMN IF NOT EXISTS external_status VARCHAR(40)")
except Exception as exc:
    st.error(str(exc)); st.stop()


def fmt(v): return f"{Decimal(v or 0):,.2f} {CURRENCY}".replace(","," ")
def rerun(): st.rerun()
def method_bg(v): return {"cash":"В брой","bank":"По банка","card_softpos":"Карта / SoftPOS"}.get(v,v)
def kind_bg(v): return {"entrance":"Такса вход","fro":"ФРО","old_debt":"Старо задължение"}.get(v,v)

def login_page():
    st.title("🏢 Building.Manager")
    a,b=st.tabs(["Вход","Регистрация за справка"])
    with a:
        with st.form("login"):
            email=st.text_input("Имейл").strip().lower(); pw=st.text_input("Парола",type="password")
            go=st.form_submit_button("Вход",use_container_width=True)
        if go:
            u=fetch_one("SELECT u.*,a.apartment_number FROM users u LEFT JOIN apartments a ON a.id=u.apartment_id WHERE lower(u.email)=lower(%s)",(email,))
            if u and verify_password(pw,u["password_hash"]):
                if u["account_status"]!="approved": st.warning("Профилът очаква одобрение.")
                else: login_user(dict(u)); rerun()
            else: st.error("Невалиден имейл или парола.")
    with b:
        apts=fetch_all("SELECT id,apartment_number FROM apartments WHERE active ORDER BY apartment_number")
        amap={str(x["apartment_number"]):x["id"] for x in apts}
        with st.form("register"):
            apt=st.selectbox("Апартамент / обект",list(amap) or ["Няма обекти"])
            c1,c2=st.columns(2); first=c1.text_input("Име"); last=c2.text_input("Фамилия")
            email=st.text_input("Имейл").strip().lower(); phone=st.text_input("Телефон")
            pw=st.text_input("Парола",type="password"); pw2=st.text_input("Повтори паролата",type="password")
            reg=st.form_submit_button("Изпрати регистрация",use_container_width=True)
        if reg:
            if not amap or not all([first,last,email,pw]) or pw!=pw2 or len(pw)<8: st.error("Провери данните и паролата.")
            else:
                try:
                    execute("INSERT INTO users(email,password_hash,first_name,last_name,phone,role,apartment_id,account_status) VALUES(%s,%s,%s,%s,%s,'neighbor',%s,'pending')",(email,hash_password(pw),first,last,phone,amap[apt])); st.success("Регистрацията е изпратена.")
                except Exception: st.error("Имейлът вече е използван.")

def payment_rows(apartment_id,year):
    create_year_obligations(apartment_id,int(year)); return obligation_rows(apartment_id,int(year))

def apartment_ledger(apartment_id,year,editable=False,key="ledger"):
    rows=payment_rows(apartment_id,year); data=[]
    for r in rows:
        paid=Decimal(r["paid"]); left=max(Decimal("0"),Decimal(r["amount"])-paid)
        data.append({"Плати":False,"Вид":kind_bg(r["fee_type"]),"Месец":MONTHS.get(r["month"],"Еднократно"),"Година":r["year"],"Начислено":float(r["amount"]),"Платено":float(paid),"Остатък":float(left),"Статус":"Платено" if left<=0 else "Частично" if paid>0 else "Неплатено","_id":r["id"]})
    df=pd.DataFrame(data)
    if df.empty: st.info("Няма задължения."); return [],Decimal("0")
    base=df.drop(columns=["_id"])
    if editable:
        view=st.data_editor(base,use_container_width=True,hide_index=True,disabled=["Вид","Месец","Година","Начислено","Платено","Остатък","Статус"],column_config={"Плати":st.column_config.CheckboxColumn("Плати"),"Начислено":st.column_config.NumberColumn(format="%.2f €"),"Платено":st.column_config.NumberColumn(format="%.2f €"),"Остатък":st.column_config.NumberColumn(format="%.2f €")},key=key)
    else:
        view=base.drop(columns=["Плати"]); st.dataframe(view,use_container_width=True,hide_index=True); return [],sum(Decimal(str(x)) for x in df["Остатък"])
    selected=[]
    for i,row in view.iterrows():
        if bool(row["Плати"]) and Decimal(str(row["Остатък"]))>0: selected.append({"id":int(df.iloc[i]["_id"]),"amount":Decimal(str(row["Остатък"])),"label":f"{row['Вид']} {row['Месец']} {row['Година']}"})
    return selected,sum((x["amount"] for x in selected),Decimal("0"))

def cashier_page(user):
    st.title("💳 Каса")
    apts=fetch_all("SELECT id,apartment_number,property_type,adult_residents,annual_fro FROM apartments WHERE active ORDER BY apartment_number")
    amap={f"{a['apartment_number']} · {a['property_type']}":a for a in apts}
    q=st.text_input("Търси апартамент или обект",placeholder="8, Офис 5, ПМ 26")
    choices=[k for k in amap if q.lower().strip() in k.lower()] if q.strip() else list(amap)
    if not choices: st.warning("Няма намерен обект."); return
    selected_name=st.selectbox("Апартамент / обект",choices); apt=amap[selected_name]
    years=[x["year"] for x in fetch_all("SELECT DISTINCT year FROM obligations WHERE apartment_id=%s ORDER BY year DESC",(apt["id"],))]
    if date.today().year not in years: years.insert(0,date.today().year)
    year=st.selectbox("Година",years or [date.today().year])
    rows=payment_rows(apt["id"],year); paid=sum(Decimal(x["paid"]) for x in rows); due=sum(max(Decimal("0"),Decimal(x["amount"])-Decimal(x["paid"])) for x in rows)
    c1,c2,c3,c4=st.columns(4); c1.metric("Обект",apt["apartment_number"]); c2.metric("Пълнолетни",apt["adult_residents"]); c3.metric("Платено",fmt(paid)); c4.metric("Остава",fmt(due))
    st.subheader("Финансов картон")
    selected,total=apartment_ledger(apt["id"],year,True,f"ledger_{apt['id']}_{year}")
    st.divider(); c1,c2=st.columns([2,1]); c1.write(f"**Избрани позиции: {len(selected)}**"); c2.metric("Обща сума",fmt(total))
    note=st.text_input("Бележка / номер на платежно")
    b1,b2,b3=st.columns(3)
    if b1.button("💵 В брой",use_container_width=True,disabled=not selected): create_payment(apt["id"],selected,"cash",user["id"],"confirmed",note); st.success("Плащането е записано."); rerun()
    if b2.button("🏦 По банка",use_container_width=True,disabled=not selected): create_payment(apt["id"],selected,"bank",user["id"],"confirmed",note); st.success("Плащането е записано."); rerun()
    if b3.button("💳 Карта / NFC",use_container_width=True,disabled=not selected):
        num=create_payment(apt["id"],selected,"card_softpos",user["id"],"pending",note or "SoftPOS")
        p=fetch_one("SELECT id FROM payments WHERE payment_number=%s",(num,)); execute("UPDATE payments SET provider='mypos_glass',external_status='initiated' WHERE id=%s",(p["id"],))
        st.session_state.softpos={"id":p["id"],"number":num,"amount":str(total)}; rerun()
    pending=st.session_state.get("softpos")
    if pending:
        st.warning(f"SoftPOS операция **{pending['number']}**, сума **{fmt(pending['amount'])}**. Потвърди само след успешно плащане в myPOS Glass.")
        template=os.getenv("MYPOS_GLASS_DEEP_LINK_TEMPLATE","")
        if template:
            url=template.replace("{amount}",quote(pending["amount"])).replace("{reference}",quote(pending["number"])).replace("{currency}","EUR")
            st.link_button("Отвори myPOS Glass",url,use_container_width=True)
        else: st.info("Стартирай сумата ръчно в myPOS Glass. Автоматичният бутон се активира след добавяне на одобрения deep-link шаблон в Render.")
        ref=st.text_input("Референция от успешната картова транзакция")
        x,y=st.columns(2)
        if x.button("✅ Потвърди успешно плащане",disabled=not ref.strip(),use_container_width=True):
            execute("UPDATE payments SET status='confirmed',payment_date=CURRENT_DATE,confirmed_by=%s,confirmed_at=NOW(),provider_transaction_id=%s,external_status='approved' WHERE id=%s AND status='pending'",(user["id"],ref.strip(),pending["id"])); audit(user["id"],"confirm_softpos","payment",pending["id"],ref); st.session_state.pop("softpos",None); rerun()
        if y.button("Откажи операцията",use_container_width=True): execute("UPDATE payments SET status='cancelled',external_status='cancelled' WHERE id=%s",(pending["id"],)); st.session_state.pop("softpos",None); rerun()
    st.subheader("Последни плащания")
    hist=fetch_all("SELECT payment_date AS дата,payment_number AS номер,amount AS сума,payment_method AS метод,status AS статус,note AS бележка FROM payments WHERE apartment_id=%s ORDER BY created_at DESC LIMIT 20",(apt["id"],))
    st.dataframe([{**dict(x),"метод":method_bg(x["метод"])} for x in hist],use_container_width=True,hide_index=True)

def daily_report(user):
    st.title("📊 Дневен отчет")
    r=fetch_one("SELECT COALESCE(SUM(amount) FILTER(WHERE payment_method='cash' AND status='confirmed'),0) cash,COALESCE(SUM(amount) FILTER(WHERE payment_method='bank' AND status='confirmed'),0) bank,COALESCE(SUM(amount) FILTER(WHERE payment_method='card_softpos' AND status='confirmed'),0) card FROM payments WHERE payment_date=CURRENT_DATE AND confirmed_by=%s",(user["id"],))
    c1,c2,c3,c4=st.columns(4); c1.metric("В брой",fmt(r["cash"])); c2.metric("По банка",fmt(r["bank"])); c3.metric("С карта",fmt(r["card"])); c4.metric("Общо",fmt(Decimal(r["cash"])+Decimal(r["bank"])+Decimal(r["card"])))
    rows=fetch_all("SELECT p.payment_date AS дата,a.apartment_number AS обект,p.payment_number AS номер,p.amount AS сума,p.payment_method AS метод,p.status AS статус FROM payments p JOIN apartments a ON a.id=p.apartment_id WHERE p.payment_date=CURRENT_DATE AND p.confirmed_by=%s ORDER BY p.confirmed_at DESC",(user["id"],))
    display=[{**dict(x),"метод":method_bg(x["метод"])} for x in rows]; st.dataframe(display,use_container_width=True,hide_index=True)
    if display:
        out=io.BytesIO(); pd.DataFrame(display).to_excel(out,index=False,engine="openpyxl"); st.download_button("Изтегли дневен отчет",out.getvalue(),f"cashier_{date.today()}.xlsx")

def expenses_page(user):
    st.title("Разходи")
    if user["role"] in ("manager","cashier"):
        cats=[x["category"] for x in fetch_all("SELECT DISTINCT category FROM expenses ORDER BY category")]
        defaults=["Електроенергия","Асансьор","Почистване","Интернет","Банкови такси","Домоуправител","Касиер","Ремонт","Консумативи","Гараж","Други"]
        with st.form("expense"):
            d=st.date_input("Дата",date.today()); cat=st.selectbox("Категория",list(dict.fromkeys(defaults+cats))); amount=st.number_input("Сума (€)",0.0,100000.0,0.0,.01); source=st.radio("Начин",["В брой","По банка"],horizontal=True); desc=st.text_area("Описание"); save=st.form_submit_button("Добави разход")
        if save and amount>0: execute("INSERT INTO expenses(expense_date,category,amount,source,description,created_by) VALUES(%s,%s,%s,%s,%s,%s)",(d,cat,amount,"cash" if source=="В брой" else "bank",desc,user["id"])); rerun()
    rows=fetch_all("SELECT expense_date AS дата,category AS категория,amount AS сума,CASE source WHEN 'cash' THEN 'В брой' ELSE 'По банка' END AS начин,description AS описание FROM expenses ORDER BY expense_date DESC,id DESC")
    st.dataframe(rows,use_container_width=True,hide_index=True)

def neighbor_page(user):
    st.title("Моите задължения")
    if not user.get("apartment_id"): st.info("Профилът не е свързан с обект."); return
    years=[x["year"] for x in fetch_all("SELECT DISTINCT year FROM obligations WHERE apartment_id=%s ORDER BY year DESC",(user["apartment_id"],))]
    year=st.selectbox("Година",years or [date.today().year]); apartment_ledger(user["apartment_id"],year,False)

def neighbor_archive(user):
    st.title("Архив")
    rows=fetch_all("SELECT payment_date AS дата,payment_number AS номер,amount AS сума,payment_method AS метод,status AS статус,note AS бележка FROM payments WHERE apartment_id=%s AND status='confirmed' ORDER BY payment_date DESC",(user["apartment_id"],)) if user.get("apartment_id") else []
    st.dataframe([{**dict(x),"метод":method_bg(x["метод"])} for x in rows],use_container_width=True,hide_index=True)

def admin_panel(user):
    st.title("Административен панел")
    tabs=st.tabs(["Табло","Апартаменти","Регистрации","Роли","Плащания","Банер","Експорт"])
    with tabs[0]:
        s=fetch_one("SELECT (SELECT COUNT(*) FROM apartments WHERE active) a,(SELECT COUNT(*) FROM users WHERE account_status='pending') p,(SELECT COALESCE(SUM(amount),0) FROM payments WHERE status='confirmed') i,(SELECT COALESCE(SUM(amount),0) FROM expenses) e")
        c1,c2,c3,c4=st.columns(4); c1.metric("Обекти",s["a"]); c2.metric("Чакащи",s["p"]); c3.metric("Постъпления",fmt(s["i"])); c4.metric("Разходи",fmt(s["e"]))
    with tabs[1]:
        with st.form("apt"):
            c1,c2,c3=st.columns(3); num=c1.text_input("Номер/обект"); typ=c2.selectbox("Тип",["Апартамент","Офис","Магазин","Гараж","Паркомясто"]); floor=c3.text_input("Етаж"); adults=st.number_input("Лица над 18 г.",0,20,1); fro=st.number_input("Годишен ФРО",0.0,10000.0,0.0,.01); add=st.form_submit_button("Добави")
        if add and num:
            try: execute("INSERT INTO apartments(apartment_number,property_type,floor,adult_residents,annual_fro) VALUES(%s,%s,%s,%s,%s)",(num,typ,floor,adults,fro)); rerun()
            except Exception: st.error("Обектът вече съществува.")
        st.dataframe(fetch_all("SELECT apartment_number AS обект,property_type AS тип,adult_residents AS лица,annual_fro AS ФРО,status AS статус FROM apartments ORDER BY apartment_number"),use_container_width=True,hide_index=True)
    with tabs[2]:
        regs=fetch_all("SELECT u.id,u.first_name,u.last_name,u.email,a.apartment_number FROM users u LEFT JOIN apartments a ON a.id=u.apartment_id WHERE u.account_status='pending'")
        for x in regs:
            c1,c2,c3=st.columns([4,1,1]); c1.write(f"**{x['first_name']} {x['last_name']}** · ап. {x['apartment_number']} · {x['email']}")
            if c2.button("Одобри",key=f"ok{x['id']}"): execute("UPDATE users SET account_status='approved' WHERE id=%s",(x["id"],)); rerun()
            if c3.button("Откажи",key=f"no{x['id']}"): execute("UPDATE users SET account_status='rejected' WHERE id=%s",(x["id"],)); rerun()
    with tabs[3]:
        users=fetch_all("SELECT id,first_name,last_name,email,role FROM users ORDER BY first_name,last_name")
        for x in users:
            c1,c2,c3=st.columns([4,2,1]); c1.write(f"**{x['first_name']} {x['last_name']}** · {x['email']}"); roles=["neighbor","cashier","manager"]; role=c2.selectbox("Роля",roles,index=roles.index(x["role"]) if x["role"] in roles else 0,key=f"role{x['id']}",label_visibility="collapsed")
            if c3.button("Запази",key=f"rs{x['id']}"): execute("UPDATE users SET role=%s,account_status='approved' WHERE id=%s",(role,x["id"])); audit(user["id"],"change_role","user",x["id"],role); rerun()
    with tabs[4]:
        rows=fetch_all("SELECT p.created_at AS дата,a.apartment_number AS обект,p.payment_number AS номер,p.amount AS сума,p.payment_method AS метод,p.status AS статус,p.provider_transaction_id AS референция FROM payments p JOIN apartments a ON a.id=p.apartment_id ORDER BY p.created_at DESC LIMIT 200")
        st.dataframe([{**dict(x),"метод":method_bg(x["метод"])} for x in rows],use_container_width=True,hide_index=True)
    with tabs[5]:
        with st.form("banner"):
            title=st.text_input("Заглавие"); body=st.text_area("Текст"); image=st.file_uploader("Снимка",type=["png","jpg","jpeg","webp"]); publish=st.form_submit_button("Публикувай")
        if publish and title:
            data=mime=None
            if image: im=Image.open(image).convert("RGB"); im.thumbnail((1600,700)); b=io.BytesIO(); im.save(b,"WEBP",quality=82); data=b.getvalue(); mime="image/webp"
            execute("UPDATE announcements SET active=FALSE"); execute("INSERT INTO announcements(title,body,image_data,image_mime,active,created_by) VALUES(%s,%s,%s,%s,TRUE,%s)",(title,body,data,mime,user["id"])); rerun()
    with tabs[6]:
        rows=fetch_all("SELECT a.apartment_number,o.fee_type,o.year,o.month,o.amount,COALESCE(SUM(CASE WHEN p.status='confirmed' THEN pa.amount ELSE 0 END),0) paid FROM obligations o JOIN apartments a ON a.id=o.apartment_id LEFT JOIN payment_allocations pa ON pa.obligation_id=o.id LEFT JOIN payments p ON p.id=pa.payment_id GROUP BY a.apartment_number,o.id ORDER BY a.apartment_number,o.year,o.month")
        out=io.BytesIO(); pd.DataFrame(rows).to_excel(out,index=False,engine="openpyxl"); st.download_button("Изтегли общ отчет",out.getvalue(),"Building_Manager_Report.xlsx")

user=current_user()
if not user: login_page(); st.stop()
with st.sidebar:
    st.title("🏢 Building.Manager"); st.caption(f"{user['first_name']} {user['last_name']} | {user['role']}")
    if user["role"]=="manager": pages=["Каса","Дневен отчет","Разходи","Админ панел"]
    elif user["role"]=="cashier": pages=["Каса","Дневен отчет","Разходи"]
    else: pages=["Моите задължения","Архив","Разходи"]
    page=st.radio("Навигация",pages)
    if st.button("Изход",use_container_width=True): logout_user(); rerun()
if page=="Каса": cashier_page(user)
elif page=="Дневен отчет": daily_report(user)
elif page=="Разходи": expenses_page(user)
elif page=="Моите задължения": neighbor_page(user)
elif page=="Архив": neighbor_archive(user)
elif page=="Админ панел": admin_panel(user)
