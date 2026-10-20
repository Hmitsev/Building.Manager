from contextlib import contextmanager
from decimal import Decimal
import os
import psycopg
from psycopg.rows import dict_row
from auth import hash_password
from config import DATABASE_URL, ADMIN_EMAIL, ADMIN_PASSWORD

SCHEMA = r"""
CREATE TABLE IF NOT EXISTS apartments (
    id BIGSERIAL PRIMARY KEY,
    apartment_number VARCHAR(30) UNIQUE NOT NULL,
    property_type VARCHAR(30) NOT NULL DEFAULT 'Апартамент',
    floor VARCHAR(20),
    status VARCHAR(30) NOT NULL DEFAULT 'Обитава се',
    adult_residents INTEGER NOT NULL DEFAULT 1 CHECK (adult_residents >= 0),
    annual_fro NUMERIC(12,2) NOT NULL DEFAULT 0,
    notes TEXT,
    active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS users (
    id BIGSERIAL PRIMARY KEY,
    email VARCHAR(255) UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    first_name VARCHAR(100) NOT NULL,
    last_name VARCHAR(100) NOT NULL,
    phone VARCHAR(40),
    role VARCHAR(30) NOT NULL DEFAULT 'neighbor',
    apartment_id BIGINT REFERENCES apartments(id),
    account_status VARCHAR(30) NOT NULL DEFAULT 'pending',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS resident_history (
    id BIGSERIAL PRIMARY KEY,
    apartment_id BIGINT NOT NULL REFERENCES apartments(id),
    adults INTEGER NOT NULL CHECK (adults >= 0),
    effective_from DATE NOT NULL,
    reason TEXT,
    created_by BIGINT REFERENCES users(id),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(apartment_id, effective_from)
);
CREATE TABLE IF NOT EXISTS obligations (
    id BIGSERIAL PRIMARY KEY,
    apartment_id BIGINT NOT NULL REFERENCES apartments(id),
    fee_type VARCHAR(30) NOT NULL CHECK (fee_type IN ('entrance','fro','old_debt')),
    year INTEGER NOT NULL,
    month INTEGER NOT NULL CHECK (month BETWEEN 0 AND 12),
    amount NUMERIC(12,2) NOT NULL CHECK (amount >= 0),
    note TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(apartment_id, fee_type, year, month)
);
CREATE TABLE IF NOT EXISTS payments (
    id BIGSERIAL PRIMARY KEY,
    payment_number VARCHAR(60) UNIQUE NOT NULL,
    apartment_id BIGINT NOT NULL REFERENCES apartments(id),
    amount NUMERIC(12,2) NOT NULL CHECK (amount > 0),
    payment_method VARCHAR(30) NOT NULL,
    status VARCHAR(30) NOT NULL DEFAULT 'pending',
    payment_date DATE,
    note TEXT,
    created_by BIGINT REFERENCES users(id),
    confirmed_by BIGINT REFERENCES users(id),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    confirmed_at TIMESTAMPTZ
);
CREATE TABLE IF NOT EXISTS payment_allocations (
    id BIGSERIAL PRIMARY KEY,
    payment_id BIGINT NOT NULL REFERENCES payments(id) ON DELETE CASCADE,
    obligation_id BIGINT NOT NULL REFERENCES obligations(id),
    amount NUMERIC(12,2) NOT NULL CHECK (amount > 0),
    UNIQUE(payment_id, obligation_id)
);
CREATE TABLE IF NOT EXISTS expenses (
    id BIGSERIAL PRIMARY KEY,
    expense_date DATE NOT NULL,
    category VARCHAR(100) NOT NULL,
    amount NUMERIC(12,2) NOT NULL CHECK (amount >= 0),
    source VARCHAR(30) NOT NULL,
    description TEXT,
    document_name TEXT,
    created_by BIGINT REFERENCES users(id),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS announcements (
    id BIGSERIAL PRIMARY KEY,
    title VARCHAR(200) NOT NULL,
    body TEXT,
    image_data BYTEA,
    image_mime VARCHAR(80),
    active BOOLEAN NOT NULL DEFAULT TRUE,
    starts_on DATE,
    ends_on DATE,
    created_by BIGINT REFERENCES users(id),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS change_requests (
    id BIGSERIAL PRIMARY KEY,
    apartment_id BIGINT NOT NULL REFERENCES apartments(id),
    requested_adults INTEGER NOT NULL CHECK (requested_adults >= 0),
    effective_from DATE NOT NULL,
    reason TEXT,
    status VARCHAR(30) NOT NULL DEFAULT 'pending',
    requested_by BIGINT REFERENCES users(id),
    reviewed_by BIGINT REFERENCES users(id),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    reviewed_at TIMESTAMPTZ
);
CREATE TABLE IF NOT EXISTS audit_log (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT REFERENCES users(id),
    action VARCHAR(100) NOT NULL,
    entity_type VARCHAR(60),
    entity_id BIGINT,
    details TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_obligations_apartment ON obligations(apartment_id, year, month);
CREATE INDEX IF NOT EXISTS idx_payments_apartment ON payments(apartment_id, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_expenses_date ON expenses(expense_date DESC);
"""

@contextmanager
def connection():
    url = os.getenv("DATABASE_URL") or DATABASE_URL
    if not url:
        raise RuntimeError("Липсва DATABASE_URL. Добавете го в Render Environment.")
    with psycopg.connect(url, row_factory=dict_row, connect_timeout=10) as conn:
        yield conn


def init_db():
    with connection() as conn:
        with conn.cursor() as cur:
            cur.execute(SCHEMA)
        conn.commit()
    seed_admin()


def seed_admin():
    email = os.getenv("ADMIN_EMAIL") or ADMIN_EMAIL
    password = os.getenv("ADMIN_PASSWORD") or ADMIN_PASSWORD
    if not email or not password:
        return
    with connection() as conn:
        row = conn.execute("SELECT id FROM users WHERE lower(email)=lower(%s)", (email,)).fetchone()
        if not row:
            conn.execute("INSERT INTO users(email,password_hash,first_name,last_name,role,account_status) VALUES(%s,%s,'Главен','Администратор','manager','approved')", (email, hash_password(password)))
            conn.commit()


def fetch_all(sql, params=()):
    with connection() as conn:
        return conn.execute(sql, params).fetchall()


def fetch_one(sql, params=()):
    with connection() as conn:
        return conn.execute(sql, params).fetchone()


def execute(sql, params=(), returning=False):
    with connection() as conn:
        row = conn.execute(sql, params).fetchone() if returning else conn.execute(sql, params)
        conn.commit()
        return row if returning else None


def money(v):
    return Decimal(str(v)).quantize(Decimal("0.01"))
