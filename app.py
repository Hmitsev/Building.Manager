import io
import re
import hashlib
from collections import Counter
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
from PIL import Image
from openpyxl import load_workbook
import streamlit as st

from auth import hash_password, verify_password, login_user, logout_user, current_user, has_role
from config import APP_NAME, CURRENCY, IBAN, PAYEE
from database import init_db, fetch_all, fetch_one, execute, connection
from services import MONTHS, create_year_obligations, obligation_rows, create_payment, audit

st.set_page_config(page_title=APP_NAME, page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
:root{--bg:#070b14;--panel:#111b2c;--panel2:#17243a;--line:rgba(148,163,184,.16);--text:#f8fafc;--muted:#94a3b8;--blue:#2563eb;--cyan:#0891b2}
.stApp{color:var(--text);background:radial-gradient(circle at 85% -10%,rgba(37,99,235,.18),transparent 30%),radial-gradient(circle at 5% 105%,rgba(8,145,178,.12),transparent 34%),linear-gradient(180deg,#070b14,#09111f)}
.block-container{padding-top:1.3rem;padding-bottom:3rem;max-width:1320px}h1,h2,h3{color:#f8fafc!important;letter-spacing:-.035em;font-weight:800!important}p,label,.stMarkdown{color:#e2e8f0}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#111827,#0b1220);border-right:1px solid var(--line)}section[data-testid="stSidebar"] h1{font-size:1.55rem!important}section[data-testid="stSidebar"] div[role="radiogroup"] label{border:1px solid transparent;border-radius:12px;padding:8px 10px;margin:2px 0;transition:.18s}section[data-testid="stSidebar"] div[role="radiogroup"] label:hover{background:rgba(59,130,246,.10);border-color:rgba(59,130,246,.22)}section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked){background:linear-gradient(135deg,rgba(37,99,235,.30),rgba(8,145,178,.16));border-color:rgba(96,165,250,.38);box-shadow:0 8px 24px rgba(2,6,23,.25)}
.bm-card{background:linear-gradient(145deg,rgba(23,36,58,.97),rgba(12,20,34,.97));border:1px solid var(--line);border-radius:20px;padding:22px;box-shadow:0 18px 44px rgba(0,0,0,.24),inset 0 1px 0 rgba(255,255,255,.04)}
div[data-testid="stMetric"]{position:relative;overflow:hidden;background:linear-gradient(145deg,rgba(25,38,61,.98),rgba(13,22,38,.98))!important;border:1px solid rgba(96,165,250,.18);border-radius:20px;padding:20px;min-height:124px;box-shadow:0 16px 38px rgba(0,0,0,.23),inset 0 1px 0 rgba(255,255,255,.04)}div[data-testid="stMetric"]:before{content:"";position:absolute;left:0;top:0;right:0;height:3px;background:linear-gradient(90deg,#2563eb,#22d3ee)}[data-testid="stMetricLabel"] p{color:#9fb0c7!important;font-weight:650!important}[data-testid="stMetricValue"]{color:#f8fafc!important;font-size:2rem!important;font-weight:800!important;letter-spacing:-.04em}
.stTabs [data-baseweb="tab-list"]{gap:7px;background:rgba(15,23,38,.75);border:1px solid var(--line);border-radius:15px;padding:6px;overflow-x:auto}.stTabs [data-baseweb="tab"]{height:42px;border-radius:10px;padding:0 14px;color:#aebdd0;background:transparent}.stTabs [data-baseweb="tab"]:hover{background:rgba(59,130,246,.1);color:white}.stTabs [aria-selected="true"]{color:white!important;background:linear-gradient(135deg,#2563eb,#0891b2)!important;box-shadow:0 8px 20px rgba(37,99,235,.26)}.stTabs [data-baseweb="tab-highlight"]{display:none}
div.stButton>button,div.stDownloadButton>button,button[kind="primaryFormSubmit"],button[kind="secondaryFormSubmit"]{color:#f8fafc!important;background:linear-gradient(135deg,#2563eb,#0891b2)!important;border:1px solid rgba(125,211,252,.24)!important;border-radius:12px!important;min-height:44px;font-weight:750!important;box-shadow:0 10px 24px rgba(37,99,235,.22),inset 0 1px 0 rgba(255,255,255,.14);transition:.16s}div.stButton>button:hover,div.stDownloadButton>button:hover,button[kind="primaryFormSubmit"]:hover,button[kind="secondaryFormSubmit"]:hover{transform:translateY(-1px);filter:brightness(1.1);box-shadow:0 14px 30px rgba(37,99,235,.31)}div.stButton>button:disabled{opacity:.45;transform:none}section[data-testid="stSidebar"] div.stButton>button{background:rgba(255,255,255,.04)!important;border-color:rgba(148,163,184,.24)!important;box-shadow:none}
[data-baseweb="input"]>div,[data-baseweb="select"]>div,[data-baseweb="textarea"]>div,[data-testid="stNumberInput"]>div>div,[data-testid="stDateInput"]>div>div{background:#111b2c!important;border-color:rgba(148,163,184,.2)!important;border-radius:12px!important;color:#f8fafc!important}input,textarea{color:#f8fafc!important;caret-color:#67e8f9!important}input::placeholder,textarea::placeholder{color:#64748b!important}[data-baseweb="popover"],[role="listbox"]{background:#111b2c!important;color:white!important}[role="option"]:hover{background:rgba(59,130,246,.18)!important}
[data-testid="stForm"]{background:linear-gradient(145deg,rgba(18,28,46,.94),rgba(11,18,31,.94));border:1px solid var(--line);border-radius:18px;padding:20px;box-shadow:0 14px 34px rgba(0,0,0,.18)}[data-testid="stFileUploaderDropzone"]{background:rgba(15,23,42,.85)!important;border:1px dashed rgba(96,165,250,.38)!important;border-radius:16px!important}[data-testid="stAlert"]{border-radius:14px;border:1px solid var(--line)}hr{border-color:var(--line)!important}
[data-testid="stDataFrame"],[data-testid="stTable"]{border:1px solid var(--line);border-radius:16px;overflow:hidden;box-shadow:0 14px 32px rgba(0,0,0,.17)}
.month-due,.month-partial,.month-paid{position:relative;border-radius:17px;padding:16px;min-height:122px;box-shadow:0 12px 28px rgba(0,0,0,.22),inset 0 1px 0 rgba(255,255,255,.04)}.month-due{background:linear-gradient(145deg,rgba(69,10,10,.92),rgba(32,12,18,.96));border:1px solid rgba(248,113,113,.34)}.month-partial{background:linear-gradient(145deg,rgba(67,20,7,.94),rgba(35,18,8,.96));border:1px solid rgba(251,146,60,.36)}.month-paid{background:linear-gradient(145deg,rgba(5,46,22,.92),rgba(8,30,20,.96));border:1px solid rgba(74,222,128,.3)}.month-title{font-size:1.02rem;font-weight:800;letter-spacing:-.02em;margin-bottom:10px}.month-sum{font-size:1.45rem;font-weight:850;letter-spacing:-.04em}.month-note{font-size:.78rem;color:#cbd5e1;margin-top:6px}.legend{display:flex;gap:18px;flex-wrap:wrap;color:#a9bad0;margin:9px 0 20px;font-size:.88rem}.dot{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:7px;box-shadow:0 0 12px currentColor}
@media(max-width:768px){.block-container{padding:1rem .85rem 2rem}h1{font-size:2rem!important}[data-testid="stMetricValue"]{font-size:1.55rem!important}div[data-testid="stMetric"]{min-height:104px;padding:16px}.stTabs [data-baseweb="tab"]{padding:0 10px;font-size:.84rem}.month-due,.month-partial,.month-paid{min-height:108px}}
</style>
""", unsafe_allow_html=True)

try:
    init_db()
except Exception as exc:
    st.error(str(exc))
    st.info("В Render добавете DATABASE_URL, ADMIN_EMAIL и ADMIN_PASSWORD в Environment.")
    st.stop()

# Допълнителните таблици са миграция без изтриване на съществуващи данни.
execute("""CREATE TABLE IF NOT EXISTS import_registry (
    id BIGSERIAL PRIMARY KEY, file_hash VARCHAR(64) UNIQUE NOT NULL,
    file_name TEXT, imported_by BIGINT REFERENCES users(id),
    imported_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), summary TEXT)""")


def fmt(value):
    return f"{Decimal(value or 0):,.2f} {CURRENCY}".replace(",", " ")


def rerun():
    st.rerun()


def classify_property(value):
    text = str(value).strip().lower()
    if "офис" in text: return "Офис"
    if "магазин" in text: return "Магазин"
    if "гараж" in text: return "Гараж"
    if "пм" in text or "парко" in text: return "Паркомясто"
    return "Апартамент"


def parse_number(value):
    if value is None or value == "": return None
    if isinstance(value, (int, float, Decimal)): return Decimal(str(value))
    match = re.search(r"-?\d+(?:[.,]\d+)?", str(value).replace(" ", ""))
    if not match: return None
    return Decimal(match.group(0).replace(",", "."))


def parse_excel_date(value):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, (int, float)):
        try: return datetime(1899, 12, 30).date() + pd.to_timedelta(int(value), unit="D")
        except Exception: return None
    text = str(value or "").strip().replace(",", ".")
    for pattern in ("%d.%m.%Y", "%m.%d.%Y", "%d.%m.%y"):
        try: return datetime.strptime(text, pattern).date()
        except ValueError: pass
    return None


def login_page():
    st.title("🏢 Building.Manager")
    tab1, tab2 = st.tabs(["Вход", "Регистрация"])
    with tab1:
        with st.form("login"):
            email = st.text_input("Имейл").strip().lower()
            password = st.text_input("Парола", type="password")
            submit = st.form_submit_button("Вход", use_container_width=True)
        if submit:
            user = fetch_one("""SELECT u.*,a.apartment_number FROM users u
                LEFT JOIN apartments a ON a.id=u.apartment_id WHERE lower(u.email)=lower(%s)""", (email,))
            if user and verify_password(password, user["password_hash"]):
                if user["account_status"] != "approved": st.warning("Регистрацията още не е одобрена.")
                else: login_user(dict(user)); rerun()
            else: st.error("Невалиден имейл или парола.")
    with tab2:
        apartments = fetch_all("SELECT id,apartment_number FROM apartments WHERE active ORDER BY apartment_number")
        if not apartments: st.info("Администраторът първо трябва да добави или импортира апартаментите.")
        with st.form("register"):
            apt_map = {str(a["apartment_number"]): a["id"] for a in apartments}
            apt = st.selectbox("Апартамент/обект", list(apt_map) or ["Няма добавени"])
            c1, c2 = st.columns(2)
            first = c1.text_input("Име")
            last = c2.text_input("Фамилия")
            email = st.text_input("Имейл").strip().lower()
            phone = st.text_input("Телефон")
            password = st.text_input("Парола", type="password")
            password2 = st.text_input("Повтори паролата", type="password")
            reg = st.form_submit_button("Изпрати регистрация", use_container_width=True)
        if reg:
            if not apartments or not all([first, last, email, password]): st.error("Попълнете задължителните полета.")
            elif password != password2 or len(password) < 8: st.error("Паролите трябва да съвпадат и да са поне 8 знака.")
            else:
                try:
                    execute("""INSERT INTO users(email,password_hash,first_name,last_name,phone,role,apartment_id,account_status)
                        VALUES(%s,%s,%s,%s,%s,'neighbor',%s,'pending')""",
                        (email, hash_password(password), first, last, phone, apt_map[apt]))
                    st.success("Регистрацията е изпратена за одобрение.")
                except Exception: st.error("Имейлът вече е използван или данните са невалидни.")


def banner():
    item = fetch_one("""SELECT * FROM announcements WHERE active
        AND (starts_on IS NULL OR starts_on<=CURRENT_DATE)
        AND (ends_on IS NULL OR ends_on>=CURRENT_DATE)
        ORDER BY created_at DESC LIMIT 1""")
    if item:
        if item["image_data"]: st.image(bytes(item["image_data"]), use_container_width=True)
        st.subheader(item["title"]); st.write(item["body"] or "")
    else:
        st.markdown("<div class='bm-card'><h2>Добре дошли в Building.Manager</h2><p>Такси, плащания, разходи и информация за входа на едно място.</p></div>", unsafe_allow_html=True)


def home(user):
    banner(); st.write("")
    if user.get("apartment_id"):
        rows = fetch_all("""SELECT o.amount,COALESCE((SELECT SUM(pa.amount) FROM payment_allocations pa
            JOIN payments p ON p.id=pa.payment_id WHERE pa.obligation_id=o.id AND p.status='confirmed'),0) paid
            FROM obligations o WHERE o.apartment_id=%s""", (user["apartment_id"],))
        due = sum(max(Decimal("0"), Decimal(r["amount"]) - Decimal(r["paid"])) for r in rows)
        paid = sum(Decimal(r["paid"]) for r in rows)
        c1, c2, c3 = st.columns(3)
        c1.metric("Апартамент", user.get("apartment_number") or "-")
        c2.metric("Платено", fmt(paid)); c3.metric("За плащане", fmt(due))
    st.subheader("Последни разходи")
    st.dataframe(fetch_all("""SELECT expense_date AS дата,category AS категория,amount AS сума,
        CASE source WHEN 'cash' THEN 'В брой' ELSE 'По банка' END AS начин,description AS описание
        FROM expenses ORDER BY expense_date DESC,id DESC LIMIT 10"""), use_container_width=True, hide_index=True)


def month_card(row, label, year, selected):
    remaining = max(Decimal("0"), Decimal(row["amount"]) - Decimal(row["paid"]))
    if remaining <= 0: return
    partial = Decimal(row["paid"]) > 0
    css = "month-partial" if partial else "month-due"
    note = f"Платено частично: {fmt(row['paid'])}" if partial else "Неплатено"
    st.markdown(f"""<div class='{css}'><div class='month-title'>{MONTHS.get(row['month'],'Старо задължение')}</div>
        <div class='month-sum'>{fmt(remaining)}</div><div class='month-note'>{note}</div></div>""", unsafe_allow_html=True)
    if st.checkbox("Добави в кошницата", key=f"ob_{row['id']}"):
        selected.append({"id": row["id"], "amount": remaining,
                         "label": f"{label} - {MONTHS.get(row['month'],'старо задължение')} {year}"})


def fees_page(user):
    if not user.get("apartment_id"): st.info("Този профил не е свързан с апартамент."); return
    years = [r["year"] for r in fetch_all("SELECT DISTINCT year FROM obligations WHERE apartment_id=%s ORDER BY year DESC", (user["apartment_id"],))]
    if date.today().year not in years: years.insert(0, date.today().year)
    year = st.selectbox("Година", years or [date.today().year])
    create_year_obligations(user["apartment_id"], int(year))
    rows = obligation_rows(user["apartment_id"], int(year))
    selected = []
    st.markdown("<div class='legend'><span><i class='dot' style='background:#dc2626'></i>Неплатено</span><span><i class='dot' style='background:#f97316'></i>Частично платено</span><span>Платените месеци са в Архив</span></div>", unsafe_allow_html=True)
    sections = [("old_debt", "Стари задължения"), ("entrance", "Такса вход"), ("fro", "Фонд ремонт и обновяване")]
    for fee_type, label in sections:
        due_rows = [r for r in rows if r["fee_type"] == fee_type and Decimal(r["amount"]) > Decimal(r["paid"])]
        if not due_rows: continue
        st.subheader(label)
        cols = st.columns(4)
        for i, row in enumerate(due_rows):
            with cols[i % 4]: month_card(row, label, year, selected)
    if not any(Decimal(r["amount"]) > Decimal(r["paid"]) for r in rows): st.success("Няма неплатени задължения за избраната година.")
    if selected:
        st.divider(); st.subheader("Кошница")
        for item in selected: st.write(f"• {item['label']}: **{fmt(item['amount'])}**")
        st.metric("Общо", fmt(sum(x["amount"] for x in selected)))
        if st.button("Създай плащане по банков път", type="primary", use_container_width=True):
            number = create_payment(user["apartment_id"], selected, "bank", user["id"], "pending", "Създадено от клиента")
            st.success(f"Заявката е създадена. Основание: {number}")
            st.write(f"Получател: **{PAYEE}**"); st.write(f"IBAN: **{IBAN or 'ще бъде предоставен от домоуправителя'}**")


def archive_page(user):
    if not user.get("apartment_id"): return
    st.subheader("Архив на платените задължения")
    rows = fetch_all("""SELECT p.payment_number AS номер,p.payment_date AS дата,
        CASE o.fee_type WHEN 'entrance' THEN 'Такса вход' WHEN 'fro' THEN 'ФРО' ELSE 'Старо задължение' END AS вид,
        o.year AS година,o.month AS месец,pa.amount AS сума,
        CASE p.payment_method WHEN 'cash' THEN 'В брой' WHEN 'bank' THEN 'По банка' ELSE p.payment_method END AS начин
        FROM payment_allocations pa JOIN payments p ON p.id=pa.payment_id
        JOIN obligations o ON o.id=pa.obligation_id
        WHERE p.apartment_id=%s AND p.status='confirmed' ORDER BY p.payment_date DESC,o.year DESC,o.month DESC""", (user["apartment_id"],))
    display = []
    for row in rows:
        row = dict(row); row["месец"] = MONTHS.get(row["месец"], "Еднократно")
        display.append(row)
    st.dataframe(display, use_container_width=True, hide_index=True)


def expenses_page():
    st.subheader("Разходи на входа")
    source = st.selectbox("Покажи", ["Всички", "В брой", "По банка"])
    where, params = "", ()
    if source != "Всички": where, params = "WHERE source=%s", (("cash" if source == "В брой" else "bank"),)
    rows = fetch_all(f"""SELECT expense_date AS дата,category AS категория,amount AS сума,
        CASE source WHEN 'cash' THEN 'В брой' ELSE 'По банка' END AS начин,description AS описание
        FROM expenses {where} ORDER BY expense_date DESC,id DESC""", params)
    st.dataframe(rows, use_container_width=True, hide_index=True)


def profile_page(user):
    st.subheader("Моят профил"); st.write(f"**{user['first_name']} {user['last_name']}**")
    st.write(user["email"]); st.write(f"Апартамент: **{user.get('apartment_number') or '-'}**")
    if user.get("apartment_id"):
        apt = fetch_one("SELECT * FROM apartments WHERE id=%s", (user["apartment_id"],))
        st.write(f"Текущ брой пълнолетни: **{apt['adult_residents']}**")
        with st.form("change_request"):
            adults = st.number_input("Нов брой лица над 18 г.", 0, 20, int(apt["adult_residents"]))
            effective = st.date_input("В сила от", date.today().replace(day=1)); reason = st.text_area("Причина")
            submit = st.form_submit_button("Изпрати искане")
        if submit:
            execute("INSERT INTO change_requests(apartment_id,requested_adults,effective_from,reason,requested_by) VALUES(%s,%s,%s,%s,%s)",
                    (user["apartment_id"], adults, effective, reason, user["id"]))
            st.success("Искането е изпратено.")


def preview_import(data):
    wb = load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.worksheets[0]
    start = next((r for r in range(1, ws.max_row + 1) if any(str(ws.cell(r, c).value).strip() == "2026" for c in range(1, 4))), None)
    objects = []
    if start:
        for r in range(start + 1, ws.max_row + 1):
            obj = ws.cell(r, 3).value
            if obj is None: continue
            if isinstance(obj, (int, float)) and 1 <= int(obj) <= 59 or isinstance(obj, str) and any(x in obj.lower() for x in ["офис", "пм"]):
                months = [ws.cell(r, c).value for c in range(4, 16)]
                objects.append((r, str(int(obj)) if isinstance(obj, float) and obj.is_integer() else str(obj), months, ws.cell(r,16).value, ws.cell(r,18).value))
    ws2 = wb.worksheets[1]
    cash_exp = sum(1 for r in range(2, ws2.max_row + 1) if parse_excel_date(ws2.cell(r,1).value) and parse_number(ws2.cell(r,3).value) is not None)
    bank_exp = sum(1 for r in range(2, ws2.max_row + 1) if parse_excel_date(ws2.cell(r,9).value) and parse_number(ws2.cell(r,11).value) is not None)
    ws3 = wb.worksheets[2]
    debts = 0
    for r in range(2, ws3.max_row + 1):
        for ac, vc in ((1,2),(4,5)):
            if ws3.cell(r,ac).value is not None and parse_number(ws3.cell(r,vc).value) not in (None, Decimal("0")): debts += 1
    return wb, objects, cash_exp, bank_exp, debts


def import_excel(data, filename, user_id):
    digest = hashlib.sha256(data).hexdigest()
    if fetch_one("SELECT id FROM import_registry WHERE file_hash=%s", (digest,)):
        raise ValueError("Този Excel файл вече е импортиран.")
    wb, objects, cash_count, bank_count, debt_count = preview_import(data)
    with connection() as conn:
        cur = conn.cursor()
        imported_objects = imported_payments = imported_expenses = imported_debts = 0
        # Лист 1: блокът 2026, C=обект, D:O=платени месеци, P=платен ФРО, R=годишен ФРО EUR.
        for row_no, obj, months, fro_paid_raw, annual_fro_raw in objects:
            annual_fro = parse_number(annual_fro_raw) or Decimal("0")
            numeric_months = [parse_number(v) for v in months if parse_number(v) is not None and parse_number(v) >= 0]
            monthly_fee = Counter(numeric_months).most_common(1)[0][0] if numeric_months else Decimal("6")
            adults = max(0, int((monthly_fee / Decimal("6")).quantize(Decimal("1"), rounding=ROUND_HALF_UP)))
            cur.execute("""INSERT INTO apartments(apartment_number,property_type,adult_residents,annual_fro,notes)
                VALUES(%s,%s,%s,%s,%s) ON CONFLICT(apartment_number) DO UPDATE SET
                property_type=EXCLUDED.property_type,adult_residents=EXCLUDED.adult_residents,
                annual_fro=EXCLUDED.annual_fro RETURNING id""",
                (obj, classify_property(obj), adults, annual_fro, "Импорт от Excel 2026"))
            apt_id = cur.fetchone()["id"]; imported_objects += 1
            cur.execute("""INSERT INTO resident_history(apartment_id,adults,effective_from,reason,created_by)
                VALUES(%s,%s,%s,'Импорт от Excel 2026',%s)
                ON CONFLICT(apartment_id,effective_from) DO UPDATE SET adults=EXCLUDED.adults""", (apt_id, adults, date(2026,1,1), user_id))
            fro_base = (annual_fro / Decimal("12")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if annual_fro else Decimal("0")
            allocated = Decimal("0")
            obligation_ids = []
            for month in range(1,13):
                cell_amount = parse_number(months[month-1])
                amount = cell_amount if cell_amount is not None and cell_amount >= 0 else monthly_fee
                cur.execute("""INSERT INTO obligations(apartment_id,fee_type,year,month,amount,note)
                    VALUES(%s,'entrance',2026,%s,%s,'Импорт от Excel')
                    ON CONFLICT(apartment_id,fee_type,year,month) DO UPDATE SET amount=EXCLUDED.amount RETURNING id""", (apt_id, month, amount))
                oid = cur.fetchone()["id"]
                if months[month-1] is not None and cell_amount is not None and cell_amount > 0:
                    number = f"IMP-ENT-2026-{apt_id}-{month}"
                    cur.execute("""INSERT INTO payments(payment_number,apartment_id,amount,payment_method,status,payment_date,note,created_by,confirmed_by,confirmed_at)
                        VALUES(%s,%s,%s,'cash','confirmed',%s,'Импорт: платен месец от Excel',%s,%s,NOW())
                        ON CONFLICT(payment_number) DO NOTHING RETURNING id""", (number, apt_id, amount, date(2026,month,1), user_id, user_id))
                    payment = cur.fetchone()
                    if payment:
                        cur.execute("INSERT INTO payment_allocations(payment_id,obligation_id,amount) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING", (payment["id"], oid, amount)); imported_payments += 1
                fro_amount = annual_fro - allocated if month == 12 else fro_base; allocated += fro_amount
                cur.execute("""INSERT INTO obligations(apartment_id,fee_type,year,month,amount,note)
                    VALUES(%s,'fro',2026,%s,%s,'Импорт от Excel')
                    ON CONFLICT(apartment_id,fee_type,year,month) DO UPDATE SET amount=EXCLUDED.amount RETURNING id""", (apt_id, month, max(Decimal("0"), fro_amount)))
                obligation_ids.append((cur.fetchone()["id"], max(Decimal("0"), fro_amount)))
            fro_paid = parse_number(fro_paid_raw) or Decimal("0")
            remaining_paid = max(Decimal("0"), min(fro_paid, annual_fro))
            if remaining_paid > 0:
                number = f"IMP-FRO-2026-{apt_id}"
                cur.execute("""INSERT INTO payments(payment_number,apartment_id,amount,payment_method,status,payment_date,note,created_by,confirmed_by,confirmed_at)
                    VALUES(%s,%s,%s,'cash','confirmed',%s,'Импорт: платен ФРО от Excel',%s,%s,NOW())
                    ON CONFLICT(payment_number) DO NOTHING RETURNING id""", (number, apt_id, remaining_paid, date(2026,1,1), user_id, user_id))
                payment = cur.fetchone()
                if payment:
                    for oid, amount in obligation_ids:
                        part = min(remaining_paid, amount)
                        if part > 0: cur.execute("INSERT INTO payment_allocations(payment_id,obligation_id,amount) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING", (payment["id"], oid, part)); remaining_paid -= part
                    imported_payments += 1
        # Лист 2: A:D разходи в брой, I:L разходи по банка.
        ws2 = wb.worksheets[1]
        for row in range(2, ws2.max_row + 1):
            for source, dc, cc, ac, nc in (("cash",1,2,3,4),("bank",9,10,11,12)):
                d = parse_excel_date(ws2.cell(row,dc).value); amount = parse_number(ws2.cell(row,ac).value)
                if not d or amount is None or amount <= 0: continue
                category = str(ws2.cell(row,cc).value or "Други").strip(); note = str(ws2.cell(row,nc).value or "").strip()
                cur.execute("INSERT INTO expenses(expense_date,category,amount,source,description,created_by) VALUES(%s,%s,%s,%s,%s,%s)", (d,category,amount,source,note,user_id)); imported_expenses += 1
        # Лист 3: A:B и D:E са старите задължения за 2024.
        ws3 = wb.worksheets[2]
        for row in range(2, ws3.max_row + 1):
            for ac, vc in ((1,2),(4,5)):
                obj = ws3.cell(row,ac).value; amount = parse_number(ws3.cell(row,vc).value)
                if obj is None or amount is None or amount <= 0: continue
                obj = str(int(obj)) if isinstance(obj,float) and obj.is_integer() else str(obj).strip()
                cur.execute("""INSERT INTO apartments(apartment_number,property_type,adult_residents,annual_fro,notes)
                    VALUES(%s,%s,1,0,'Създаден от импорт на стари задължения') ON CONFLICT(apartment_number) DO UPDATE SET apartment_number=EXCLUDED.apartment_number RETURNING id""", (obj,classify_property(obj)))
                apt_id = cur.fetchone()["id"]
                cur.execute("""INSERT INTO obligations(apartment_id,fee_type,year,month,amount,note)
                    VALUES(%s,'old_debt',2024,0,%s,'Старо задължение от Excel')
                    ON CONFLICT(apartment_id,fee_type,year,month) DO UPDATE SET amount=EXCLUDED.amount""", (apt_id,amount)); imported_debts += 1
        summary = f"Обекти: {imported_objects}; плащания: {imported_payments}; разходи: {imported_expenses}; стари задължения: {imported_debts}"
        cur.execute("INSERT INTO import_registry(file_hash,file_name,imported_by,summary) VALUES(%s,%s,%s,%s)", (digest,filename,user_id,summary))
        conn.commit()
    return summary


def admin_page(user):
    st.title("Административен панел")
    tabs = st.tabs(["Табло","Апартаменти","Регистрации","Плащания","Разходи","Банер","Промени","Импорт Excel","Експорт"])
    with tabs[0]:
        stats = fetch_one("""SELECT (SELECT COUNT(*) FROM apartments WHERE active) apartments,
            (SELECT COUNT(*) FROM users WHERE account_status='pending') pending,
            (SELECT COALESCE(SUM(amount),0) FROM payments WHERE status='confirmed') income,
            (SELECT COALESCE(SUM(amount),0) FROM expenses) expenses""")
        c1,c2,c3,c4 = st.columns(4); c1.metric("Обекти",stats["apartments"]); c2.metric("Чакащи регистрации",stats["pending"]); c3.metric("Постъпления",fmt(stats["income"])); c4.metric("Разходи",fmt(stats["expenses"]))
    with tabs[1]:
        with st.form("add_apt"):
            c1,c2,c3=st.columns(3); num=c1.text_input("Номер/обект"); typ=c2.selectbox("Тип",["Апартамент","Офис","Магазин","Гараж","Паркомясто","Друг"]); floor=c3.text_input("Етаж")
            c4,c5=st.columns(2); adults=c4.number_input("Лица над 18 г.",0,20,1); fro=c5.number_input("Годишен ФРО (€)",0.0,10000.0,0.0,step=0.01)
            notes=st.text_area("Бележка"); add=st.form_submit_button("Добави обект")
        if add and num:
            try:
                row=execute("INSERT INTO apartments(apartment_number,property_type,floor,adult_residents,annual_fro,notes) VALUES(%s,%s,%s,%s,%s,%s) RETURNING id",(num,typ,floor,adults,fro,notes),True)
                execute("INSERT INTO resident_history(apartment_id,adults,effective_from,reason,created_by) VALUES(%s,%s,%s,'Начална стойност',%s)",(row["id"],adults,date(date.today().year,1,1),user["id"])); st.success("Обектът е добавен."); rerun()
            except Exception as exc: st.error(f"Обектът не беше добавен: {exc}")
        st.dataframe(fetch_all("SELECT apartment_number AS обект,property_type AS тип,floor AS етаж,adult_residents AS пълнолетни,annual_fro AS годишен_ФРО,status AS статус FROM apartments ORDER BY apartment_number"),use_container_width=True,hide_index=True)
    with tabs[2]:
        regs=fetch_all("SELECT u.id,u.first_name,u.last_name,u.email,a.apartment_number FROM users u LEFT JOIN apartments a ON a.id=u.apartment_id WHERE u.account_status='pending' ORDER BY u.created_at")
        for r in regs:
            c1,c2,c3=st.columns([4,1,1]); c1.write(f"**{r['first_name']} {r['last_name']}** | Ап. {r['apartment_number']} | {r['email']}")
            if c2.button("Одобри",key=f"approve{r['id']}"): execute("UPDATE users SET account_status='approved' WHERE id=%s",(r["id"],)); audit(user["id"],"approve_user","user",r["id"]); rerun()
            if c3.button("Откажи",key=f"reject{r['id']}"): execute("UPDATE users SET account_status='rejected' WHERE id=%s",(r["id"],)); rerun()
    with tabs[3]:
        pending=fetch_all("SELECT p.id,p.payment_number,p.amount,p.created_at,a.apartment_number FROM payments p JOIN apartments a ON a.id=p.apartment_id WHERE p.status='pending' ORDER BY p.created_at")
        for p in pending:
            c1,c2=st.columns([5,1]); c1.write(f"**{p['payment_number']}** | Ап. {p['apartment_number']} | {fmt(p['amount'])}")
            if c2.button("Потвърди",key=f"pay{p['id']}"): execute("UPDATE payments SET status='confirmed',payment_date=CURRENT_DATE,confirmed_by=%s,confirmed_at=NOW() WHERE id=%s",(user["id"],p["id"])); rerun()
        st.divider(); st.subheader("Отразяване на плащане в брой или по банка")
        apts=fetch_all("SELECT id,apartment_number FROM apartments WHERE active ORDER BY apartment_number"); amap={str(a["apartment_number"]):a["id"] for a in apts}
        if apts:
            apt_num=st.selectbox("Апартамент/обект",list(amap)); yr=st.number_input("Година",2022,2100,date.today().year)
            create_year_obligations(amap[apt_num],int(yr)); dues=[r for r in obligation_rows(amap[apt_num],int(yr)) if Decimal(r["amount"])>Decimal(r["paid"])]
            labels={f"{('Такса вход' if r['fee_type']=='entrance' else 'ФРО' if r['fee_type']=='fro' else 'Старо задължение')} {MONTHS.get(r['month'],'')} - {fmt(Decimal(r['amount'])-Decimal(r['paid']))}":r for r in dues}
            picked=st.multiselect("Какво е платено",list(labels)); method=st.radio("Начин на плащане",["В брой","По банка"],horizontal=True); note=st.text_input("Бележка/номер на платежно")
            if st.button("Запиши плащането",type="primary") and picked:
                items=[{"id":labels[x]["id"],"amount":Decimal(labels[x]["amount"])-Decimal(labels[x]["paid"])} for x in picked]
                create_payment(amap[apt_num],items,"cash" if method=="В брой" else "bank",user["id"],"confirmed",note); st.success("Плащането е записано и месеците са преместени в Архив."); rerun()
    with tabs[4]:
        existing=[r["category"] for r in fetch_all("SELECT DISTINCT category FROM expenses ORDER BY category")]
        defaults=["Електроенергия","Асансьор","Почистване","Интернет","Банкови такси","Домоуправител","Касиер","Ремонт","Консумативи","Гараж","Други"]
        categories=list(dict.fromkeys(defaults+existing))
        with st.form("expense"):
            d=st.date_input("Дата",date.today()); category=st.selectbox("Категория",categories); amount=st.number_input("Сума (€)",0.0,100000.0,0.0,step=0.01)
            source=st.radio("Разходът е платен",["В брой","По банка"],horizontal=True); description=st.text_area("Описание"); save=st.form_submit_button("Добави разход")
        if save and amount>0: execute("INSERT INTO expenses(expense_date,category,amount,source,description,created_by) VALUES(%s,%s,%s,%s,%s,%s)",(d,category,amount,"cash" if source=="В брой" else "bank",description,user["id"])); st.success("Разходът е добавен."); rerun()
        st.dataframe(fetch_all("SELECT expense_date AS дата,category AS категория,amount AS сума,CASE source WHEN 'cash' THEN 'В брой' ELSE 'По банка' END AS начин,description AS описание FROM expenses ORDER BY expense_date DESC,id DESC"),use_container_width=True,hide_index=True)
    with tabs[5]:
        with st.form("banner_form"):
            title=st.text_input("Заглавие"); body=st.text_area("Текст"); image=st.file_uploader("Снимка",type=["png","jpg","jpeg","webp"]); starts=st.date_input("Начална дата",date.today()); ends=st.date_input("Крайна дата",value=None); publish=st.form_submit_button("Публикувай")
        if publish and title:
            data=mime=None
            if image:
                im=Image.open(image).convert("RGB"); im.thumbnail((1600,700)); buf=io.BytesIO(); im.save(buf,format="WEBP",quality=82,method=6); data=buf.getvalue(); mime="image/webp"
            execute("UPDATE announcements SET active=FALSE"); execute("INSERT INTO announcements(title,body,image_data,image_mime,starts_on,ends_on,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s)",(title,body,data,mime,starts,ends,user["id"])); st.success("Банерът е публикуван."); rerun()
    with tabs[6]:
        reqs=fetch_all("SELECT cr.*,a.apartment_number FROM change_requests cr JOIN apartments a ON a.id=cr.apartment_id WHERE cr.status='pending' ORDER BY cr.created_at")
        for r in reqs:
            st.write(f"Ап. **{r['apartment_number']}**: {r['requested_adults']} лица от {r['effective_from']} | {r['reason'] or ''}")
            c1,c2=st.columns(2)
            if c1.button("Одобри",key=f"chga{r['id']}"):
                execute("UPDATE change_requests SET status='approved',reviewed_by=%s,reviewed_at=NOW() WHERE id=%s",(user["id"],r["id"])); execute("UPDATE apartments SET adult_residents=%s WHERE id=%s",(r["requested_adults"],r["apartment_id"])); execute("INSERT INTO resident_history(apartment_id,adults,effective_from,reason,created_by) VALUES(%s,%s,%s,%s,%s) ON CONFLICT(apartment_id,effective_from) DO UPDATE SET adults=EXCLUDED.adults,reason=EXCLUDED.reason",(r["apartment_id"],r["requested_adults"],r["effective_from"],r["reason"],user["id"])); rerun()
            if c2.button("Откажи",key=f"chgr{r['id']}"): execute("UPDATE change_requests SET status='rejected',reviewed_by=%s,reviewed_at=NOW() WHERE id=%s",(user["id"],r["id"])); rerun()
    with tabs[7]:
        st.subheader("Импорт от първите 3 листа")
        st.info("Лист 1: плащания и ФРО за 2026; Лист 2: разходи в брой и по банка; Лист 3: стари задължения. Другите листове не се четат.")
        upload=st.file_uploader("Качи Excel файла",type=["xlsx"],key="excel_import")
        if upload:
            data=upload.getvalue()
            try:
                _, objects, cash_count, bank_count, debt_count=preview_import(data)
                c1,c2,c3,c4=st.columns(4); c1.metric("Обекти",len(objects)); c2.metric("Разходи в брой",cash_count); c3.metric("Разходи по банка",bank_count); c4.metric("Стари задължения",debt_count)
                confirm=st.checkbox("Потвърждавам импорта в базата")
                if st.button("Импортирай данните",type="primary",disabled=not confirm):
                    summary=import_excel(data,upload.name,user["id"]); st.success("Импортът приключи: "+summary); rerun()
            except Exception as exc: st.error(f"Файлът не може да бъде импортиран: {exc}")
        history=fetch_all("SELECT file_name AS файл,imported_at AS дата,summary AS резултат FROM import_registry ORDER BY imported_at DESC")
        if history: st.dataframe(history,use_container_width=True,hide_index=True)
    with tabs[8]:
        dataset=fetch_all("SELECT a.apartment_number,o.fee_type,o.year,o.month,o.amount,COALESCE(SUM(CASE WHEN p.status='confirmed' THEN pa.amount ELSE 0 END),0) paid FROM obligations o JOIN apartments a ON a.id=o.apartment_id LEFT JOIN payment_allocations pa ON pa.obligation_id=o.id LEFT JOIN payments p ON p.id=pa.payment_id GROUP BY a.apartment_number,o.id ORDER BY a.apartment_number,o.year,o.month")
        df=pd.DataFrame(dataset); buf=io.BytesIO(); df.to_excel(buf,index=False,engine="openpyxl"); st.download_button("Изтегли Excel отчет",buf.getvalue(),"Building_Manager_Report.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


user=current_user()
if not user: login_page(); st.stop()
with st.sidebar:
    st.title("🏢 Building.Manager"); st.caption(f"{user['first_name']} {user['last_name']} | {user['role']}")
    choices=["Начало","Моите такси","Архив","Разходи","Профил"]
    if has_role("manager","cashier"): choices.append("Админ панел")
    page=st.radio("Навигация",choices)
    if st.button("Изход",use_container_width=True): logout_user(); rerun()
if page=="Начало": home(user)
elif page=="Моите такси": fees_page(user)
elif page=="Архив": archive_page(user)
elif page=="Разходи": expenses_page()
elif page=="Профил": profile_page(user)
elif page=="Админ панел": admin_page(user)
