from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
import secrets
from database import fetch_all, fetch_one, execute, money
from config import ENTRY_FEE_PER_ADULT

MONTHS = {1:'Януари',2:'Февруари',3:'Март',4:'Април',5:'Май',6:'Юни',7:'Юли',8:'Август',9:'Септември',10:'Октомври',11:'Ноември',12:'Декември'}

def audit(user_id, action, entity_type=None, entity_id=None, details=None):
    execute("INSERT INTO audit_log(user_id,action,entity_type,entity_id,details) VALUES(%s,%s,%s,%s,%s)", (user_id,action,entity_type,entity_id,details))

def create_year_obligations(apartment_id:int, year:int):
    apt=fetch_one("SELECT * FROM apartments WHERE id=%s",(apartment_id,))
    if not apt: return
    fro_month=(Decimal(apt['annual_fro'])/Decimal('12')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    allocated=Decimal('0.00')
    for month in range(1,13):
        adults_row=fetch_one("SELECT adults FROM resident_history WHERE apartment_id=%s AND effective_from<=%s ORDER BY effective_from DESC LIMIT 1",(apartment_id,date(year,month,1)))
        adults=adults_row['adults'] if adults_row else apt['adult_residents']
        entry=money(Decimal(adults)*Decimal(str(ENTRY_FEE_PER_ADULT)))
        execute("INSERT INTO obligations(apartment_id,fee_type,year,month,amount) VALUES(%s,'entrance',%s,%s,%s) ON CONFLICT(apartment_id,fee_type,year,month) DO NOTHING",(apartment_id,year,month,entry))
        fro = money(Decimal(apt['annual_fro'])-allocated) if month==12 else fro_month
        allocated += fro
        execute("INSERT INTO obligations(apartment_id,fee_type,year,month,amount) VALUES(%s,'fro',%s,%s,%s) ON CONFLICT(apartment_id,fee_type,year,month) DO NOTHING",(apartment_id,year,month,fro))

def obligation_rows(apartment_id, year):
    return fetch_all("""SELECT o.*, COALESCE((SELECT SUM(pa.amount) FROM payment_allocations pa JOIN payments p ON p.id=pa.payment_id WHERE pa.obligation_id=o.id AND p.status='confirmed'),0) paid FROM obligations o WHERE o.apartment_id=%s AND o.year=%s ORDER BY o.fee_type,o.month""",(apartment_id,year))

def create_payment(apartment_id, items, method, user_id, status='pending', note=''):
    total=sum(Decimal(str(i['amount'])) for i in items)
    number=f"BM-{apartment_id}-{datetime.now():%Y%m%d}-{secrets.token_hex(3).upper()}"
    row=execute("INSERT INTO payments(payment_number,apartment_id,amount,payment_method,status,payment_date,note,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id",(number,apartment_id,total,method,status,date.today() if status=='confirmed' else None,note,user_id),True)
    for item in items:
        execute("INSERT INTO payment_allocations(payment_id,obligation_id,amount) VALUES(%s,%s,%s)",(row['id'],item['id'],money(item['amount'])))
    audit(user_id,'create_payment','payment',row['id'],number)
    return number
