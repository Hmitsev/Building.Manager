import os
from pathlib import Path

APP_NAME = "Building.Manager"
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
DATABASE_URL = os.getenv("DATABASE_URL", "")
APP_SECRET = os.getenv("APP_SECRET", "change-this-in-render")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "")
ENTRY_FEE_PER_ADULT = 6.00
CURRENCY = "EUR"
IBAN = os.getenv("BUILDING_IBAN", "")
PAYEE = os.getenv("BUILDING_PAYEE", "Етажна собственост")
