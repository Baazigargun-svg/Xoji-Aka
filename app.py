from datetime import datetime, timedelta, timezone
import io
import threading
import time
import os
import sqlite3
import logging
import hmac
import hashlib
import shutil
import uuid
from imagekitio import ImageKit
from werkzeug.utils import secure_filename
from urllib.parse import urlencode
from logging.handlers import RotatingFileHandler
from flask import (
    Flask,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
    send_from_directory,
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import telebot
from telebot import types

# --- SOZLAMALAR ---
# Xavfsizlik: bot tokeni va maxfiy kalitlarni kodga yozmang.
# Windows PowerShell: $env:BOT_TOKEN="BotFather bergan yangi token"
BOT_TOKEN = os.environ.get("BOT_TOKEN", "8573337094:AAHreTrkHh5JM0DBd9vBwJtnJKTaSpyLhps").strip()
ADMIN_ID = int(os.environ.get("ADMIN_ID", "6851851908"))
SEX_GROUP_ID = int(os.environ.get("SEX_GROUP_ID", "-1003936599812"))

# TeleBot dekoratorlari token bo'lmaganda ham yuklanishi uchun vaqtinchalik qiymat.
# Haqiqiy polling faqat BOT_TOKEN berilganda ishga tushadi.
bot = telebot.TeleBot(BOT_TOKEN or "000000:PLACEHOLDER_TOKEN")
app = Flask(__name__)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_NAME = os.environ.get('DB_NAME', os.path.join(BASE_DIR, 'xoji_aka_factory.db'))  # Standart holatda aynan shu fayl yonidagi eski baza ishlatiladi
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or os.urandom(32)
# Xatolar jurnali: fayl hajmi 2 MB dan oshsa aylantiriladi.
LOG_FILE = os.path.join(BASE_DIR, 'xoji_aka_errors.log')
_file_logger = RotatingFileHandler(LOG_FILE, maxBytes=2 * 1024 * 1024, backupCount=5, encoding='utf-8')
_file_logger.setLevel(logging.INFO)
_file_logger.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s'))
app.logger.addHandler(_file_logger)
app.logger.setLevel(logging.INFO)
BACKUP_DIR = os.path.join(BASE_DIR, 'db_backups')
WEB_BASE_URL = (os.environ.get('WEB_BASE_URL') or os.environ.get('RENDER_EXTERNAL_URL') or 'https://xoji-aka-factory.onrender.com').rstrip('/')

app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
ADMIN_LOGIN = os.environ.get('ADMIN_LOGIN', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'bigbang0231.').strip()
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
CATALOG_UPLOAD_FOLDER = os.path.join(UPLOAD_FOLDER, 'catalog')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CATALOG_UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_CATALOG_EXTENSIONS = {'jpg', 'jpeg', 'png', 'webp'}
# ===== IMAGEKIT SOZLAMALARI =====
IMAGEKIT_PRIVATE_KEY = os.environ.get('IMAGEKIT_PRIVATE_KEY')
IMAGEKIT_URL_ENDPOINT = os.environ.get('IMAGEKIT_URL_ENDPOINT')

imagekit = None

if IMAGEKIT_PRIVATE_KEY:
    imagekit = ImageKit(
        private_key=IMAGEKIT_PRIVATE_KEY
    )
app.logger.info(
    "ImageKit private key mavjud: %s",
    bool(IMAGEKIT_PRIVATE_KEY)
)

def imagekitga_yuklash(file_obj):
    """Rasmni ImageKit'ga yuklaydi va URL manzilini qaytaradi."""
    if imagekit is None:
        raise RuntimeError(
            "ImageKit sozlanmagan. Render Environment sozlamalarini tekshiring."
        )

    if not file_obj or not file_obj.filename:
        raise ValueError("Rasm tanlanmagan.")

    filename = secure_filename(file_obj.filename)
    extension = os.path.splitext(filename)[1].lower()

    if extension not in {'.jpg', '.jpeg', '.png', '.webp'}:
        raise ValueError("Faqat JPG, JPEG, PNG yoki WEBP rasmlar mumkin.")

    result = imagekit.files.upload(
        file=file_obj.read(),
        file_name=f"{uuid.uuid4().hex}{extension}",
        folder="/xoji-aka-products"
    )

    if not result.url:
        raise RuntimeError("ImageKit rasm manzilini qaytarmadi.")

    return result.url


# ===== IMAGEKIT KODI TUGADI =====
user_steps = {}
# Slayd uchun foydalanuvchining turgan joyi (index)
catalog_indexes = {}

def get_db_connection():
  conn = sqlite3.connect(DB_NAME)
  conn.row_factory = sqlite3.Row
  return conn


def is_admin_tg(tg_id):
  try:
    if int(tg_id) == ADMIN_ID:
      return True
    conn = get_db_connection()
    row = conn.execute('SELECT role FROM users WHERE tg_id = ?', (int(tg_id),)).fetchone()
    conn.close()
    return bool(row and 'admin' in (row['role'] or '').lower())
  except Exception:
    return False


def allowed_catalog_file(filename):
  return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_CATALOG_EXTENSIONS


def safe_number(value, default=0.0, minimum=None, maximum=None):
  """Forma qiymatini xavfsiz son qiladi; vergulli o'nlikni ham qabul qiladi."""
  try:
    number = float(str(value if value is not None else default).strip().replace(' ', '').replace(',', '.'))
    if number != number or number in (float('inf'), float('-inf')):
      return default
  except (TypeError, ValueError):
    return default
  if minimum is not None:
    number = max(minimum, number)
  if maximum is not None:
    number = min(maximum, number)
  return number


def current_actor():
  try:
    return str(session.get('username') or 'Tizim')
  except Exception:
    return 'Telegram bot'


def log_audit(conn, action, entity='', entity_id='', details=''):
  try:
    actor = current_actor()
    conn.execute('INSERT INTO audit_log (actor, action, entity, entity_id, details, date) VALUES (?, ?, ?, ?, ?, ?)',
      (actor, str(action), str(entity), str(entity_id), str(details)[:1000], datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
  except Exception as exc:
    print('Audit jurnaliga yozish xatosi:', exc)


def record_stock_movement(conn, product_name, qty_change, reason, reference=''):
  conn.execute('INSERT INTO stock_movements (product_name, qty_change, reason, reference, date, actor) VALUES (?, ?, ?, ?, ?, ?)',
    (product_name, qty_change, reason, str(reference), datetime.now().strftime('%Y-%m-%d %H:%M:%S'), current_actor()))


def _web_link_secret():
  return (os.environ.get('WEB_LINK_SECRET') or BOT_TOKEN or os.environ.get('FLASK_SECRET_KEY') or str(app.secret_key)).encode('utf-8')


def make_staff_web_url(tg_id):
  """Telegram ID uchun 15 daqiqalik imzolangan xodim havolasi."""
  ts = int(time.time())
  raw = f'{int(tg_id)}:{ts}'.encode('utf-8')
  sig = hmac.new(_web_link_secret(), raw, hashlib.sha256).hexdigest()
  # Agent uchun mavjud buyurtma kiritish sahifasi; boshqa rollar uchun rolga mos kabinet.
  try:
    _c = get_db_connection()
    _u = _c.execute('SELECT role FROM users WHERE tg_id = ?', (int(tg_id),)).fetchone()
    _c.close()
    _role = (_u['role'] or '') if _u else ''
  except Exception:
    _role = ''
  _path = 'agent_app' if ('agent' in _role and 'admin' not in _role) else 'staff_app'
  return f'{WEB_BASE_URL}/{_path}/{int(tg_id)}?{urlencode({"ts": ts, "sig": sig})}'


def verify_staff_link(tg_id):
  try:
    ts = int(request.args.get('ts', '0'))
    sig = request.args.get('sig', '')
    if not sig or abs(int(time.time()) - ts) > 900:
      return False
    expected = hmac.new(_web_link_secret(), f'{int(tg_id)}:{ts}'.encode('utf-8'), hashlib.sha256).hexdigest()
    return hmac.compare_digest(sig, expected)
  except Exception:
    return False


def create_database_backup():
  """SQLite backup API orqali izchil zaxira yaratadi; oxirgi 14 nusxa saqlanadi."""
  try:
    if not os.path.exists(DB_NAME):
      return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    destination = os.path.join(BACKUP_DIR, f'xoji_aka_factory_{stamp}.db')
    source_conn = sqlite3.connect(DB_NAME, timeout=30)
    target_conn = sqlite3.connect(destination)
    try:
      source_conn.backup(target_conn)
    finally:
      target_conn.close()
      source_conn.close()
    backups = sorted([os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith('.db')], key=os.path.getmtime, reverse=True)
    for old_path in backups[14:]:
      try: os.remove(old_path)
      except OSError: app.logger.exception('Eski zaxirani o‘chirishda xato: %s', old_path)
    app.logger.info('Ma’lumotlar bazasi zaxiralandi: %s', destination)
  except Exception:
    app.logger.exception('Avtomatik baza zaxirasida xato')


def backup_scheduler():
  # Har 24 soatda; ilova ishga tushganda ham zaxira olinadi.
  while True:
    time.sleep(24 * 60 * 60)
    create_database_backup()


def init_web_db():
  conn = get_db_connection()
  cursor = conn.cursor()

  cursor.execute('''CREATE TABLE IF NOT EXISTS catalog_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      category TEXT DEFAULT 'Boshqa',
      price REAL NOT NULL DEFAULT 0,
      description TEXT DEFAULT '',
      image_filename TEXT DEFAULT '',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL)''')

  # Aytaylik, agar bu jadval sizda oldinroq boshqacharoq yaratilgan bo'lsa 
  # (category ustuni yo'q bo'lsa), xatolik bermasligi uchun migration ham qo'shib qo'yish mumkin:
  try:
      cursor.execute("ALTER TABLE catalog_items ADD COLUMN category TEXT DEFAULT 'Boshqa'")
  except:
      pass

    
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS users (tg_id INTEGER PRIMARY KEY, name TEXT, phone TEXT, role TEXT DEFAULT 'pending', territory TEXT DEFAULT '', salary_percent REAL DEFAULT 5)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS expenses 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, reason TEXT, amount REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS incomes 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, source TEXT, amount REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS advances 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, worker_name TEXT, amount REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS product_incomes 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT, qty REAL, cost_price REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS orders 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, shop_name TEXT, agent_name TEXT, items_text TEXT, total_sum REAL, discount REAL DEFAULT 0, status TEXT, date TEXT, price_type TEXT, comment TEXT DEFAULT '', cost_total REAL)'''
  )
  cursor.execute("""CREATE TABLE IF NOT EXISTS stock_movements (
      id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT NOT NULL,
      qty_change REAL NOT NULL, reason TEXT NOT NULL, reference TEXT DEFAULT '',
      date TEXT NOT NULL, actor TEXT DEFAULT 'Tizim')""")
  cursor.execute("""CREATE TABLE IF NOT EXISTS audit_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT, actor TEXT, action TEXT NOT NULL,
      entity TEXT, entity_id TEXT, details TEXT, date TEXT NOT NULL)""")
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS order_status_history 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER, status TEXT, changed_at TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS products 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, category TEXT DEFAULT 'Boshqa', stock REAL DEFAULT 0, cost_price REAL DEFAULT 0, optom_price REAL DEFAULT 0, chakana_price REAL DEFAULT 0, agent_percent REAL DEFAULT 5)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS shops 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, firm TEXT DEFAULT '', name TEXT UNIQUE, phone TEXT, debt REAL DEFAULT 0, visit_days TEXT, region TEXT DEFAULT '', landmark TEXT DEFAULT '', inventory TEXT DEFAULT '', territory TEXT DEFAULT '')'''
  )
  cursor.execute('''CREATE TABLE IF NOT EXISTS catalog_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      price REAL NOT NULL DEFAULT 0,
      description TEXT DEFAULT '',
      image_filename TEXT DEFAULT '',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL)''')
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS payment_requests 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, shop_name TEXT, ekspeditor_name TEXT, amount REAL, status TEXT DEFAULT 'pending', date TEXT)'''
  )

  migrations = [
      (
          "ALTER TABLE products ADD COLUMN category TEXT DEFAULT 'Boshqa'",
          "category",
      ),
      ("ALTER TABLE products ADD COLUMN cost_price REAL DEFAULT 0", "cost_price"),
      (
          "ALTER TABLE products ADD COLUMN optom_price REAL DEFAULT 0",
          "optom_price",
      ),
      (
          "ALTER TABLE products ADD COLUMN chakana_price REAL DEFAULT 0",
          "chakana_price",
      ),
      ("ALTER TABLE products ADD COLUMN stock REAL DEFAULT 0", "stock"),
      ("ALTER TABLE products ADD COLUMN agent_percent REAL DEFAULT 5", "agent_percent"),
      ("ALTER TABLE shops ADD COLUMN visit_days TEXT DEFAULT ''", "visit_days"),
      ("ALTER TABLE shops ADD COLUMN region TEXT DEFAULT ''", "region"),
      ("ALTER TABLE shops ADD COLUMN landmark TEXT DEFAULT ''", "landmark"),
      ("ALTER TABLE shops ADD COLUMN inventory TEXT DEFAULT ''", "inventory"),
      ("ALTER TABLE shops ADD COLUMN territory TEXT DEFAULT ''", "territory"),
      ("ALTER TABLE shops ADD COLUMN firm TEXT DEFAULT ''", "firm"),
      ("ALTER TABLE shops ADD COLUMN last_visit TEXT DEFAULT ''", "last_visit"),
      ("ALTER TABLE shops ADD COLUMN next_visit_date TEXT DEFAULT ''", "next_visit_date"),
      ("ALTER TABLE users ADD COLUMN territory TEXT DEFAULT ''", "territory"),
      ("ALTER TABLE users ADD COLUMN salary_percent REAL DEFAULT 5", "salary_percent"),
      ("ALTER TABLE orders ADD COLUMN discount REAL DEFAULT 0", "discount"),
      ("ALTER TABLE orders ADD COLUMN price_type TEXT", "price_type"),
      ("ALTER TABLE orders ADD COLUMN comment TEXT DEFAULT ''", "comment"),
      ("ALTER TABLE orders ADD COLUMN cost_total REAL", "cost_total"),
  ]

  for query, col in migrations:
    try:
      cursor.execute(query)
    except:
      pass

  cursor.execute(
      "INSERT OR IGNORE INTO users (tg_id, name, phone, role) VALUES (?,"
      " 'Admin', '', 'admin')",
      (ADMIN_ID,),
  )

  initial_agents = [
      (8241020136, "Qoraboyev Sirojiddin", "+998935075540", "agent,sex"),
      (2101923750, "Qoraboyeva Charos", "+998940300206", "agent"),
      (6431126337, "Asror", "+998912360777", "agent"),
      (6851851908, "SRJ", "+998975155540", "agent,admin"),
  ]
  for ag_id, ag_name, ag_phone, ag_role in initial_agents:
    cursor.execute(
        "INSERT OR IGNORE INTO users (tg_id, name, phone, role) VALUES (?, ?,"
        " ?, ?)",
        (ag_id, ag_name, ag_phone, ag_role),
    )

  initial_shops = [
      # --- Dushanba ---
      ('', 'Sherzod market', '', 'Termiz shahar', 'Uzgazoil qatori', 'Dushanba', 'Port va Senter'),
      ('', 'Qayumov Kamol', '', 'Termiz shahar', 'Murch boboga yetmasdan', 'Dushanba', 'Port va Senter'),
      ('', 'Akbar aka', '', 'Termiz shahar', 'Bekat murch bobo', 'Dushanba', 'Port va Senter'),
      ('', 'Pub house', '', 'Termiz shahar', 'Pub house', 'Dushanba', 'Port va Senter'),
      ('', 'Chapayev roparasi', '', 'Termiz shahar', 'Chapayev roparasi', 'Dushanba', 'Port va Senter'),
      ('', 'Toshpoʻlat aka', '', 'Termiz shahar', 'Uchrashuv yoni', 'Dushanba', 'Port va Senter'),
      ('', 'Abdulfayz bekat', '', 'Termiz shahar', '', 'Dushanba', 'Port va Senter'),
      ('', 'Dilorom', '', 'Termiz shahar', 'Tutzor', 'Dushanba', 'Port va Senter'),
      ('', 'Joʻrabek aka', '', 'Termiz shahar', 'Anorcha tagi', 'Dushanba', 'Senter'),
      ('', 'Bahor Market', '', 'Termiz tuman', '', 'Dushanba', 'PSKUL'),
      ('', 'Baraka Market', '', 'Termiz tuman', '', 'Dushanba', 'PSKUL'),
      ('', 'Farxod Market', '', 'Senter Kamaz', '', 'Dushanba', 'PSKUL'),
      ('', 'Ariqcha Market', '', 'Senter Kamaz', '', 'Dushanba', 'PSKUL'),
      ('', 'Xolida Market', '', 'Limon', '', 'Dushanba', 'Limon'),
      ('', 'Nodirov Mansur', '', 'Termiz shahar', 'Yoshar markazi yonboshi', 'Dushanba', 'Port va Senter'),
      ('', 'Nazarova Nodira', '', 'Termiz shahar', 'SOHIL MED TEXNIKUM TOGRISI', 'Dushanba', 'Port va Senter'),
      ('', 'Aziza opa', '', 'Termiz shahar', 'Nargiza opa sinlisi', 'Dushanba', 'Port va Senter'),
      ('', '5x5 market', '', 'Termiz shahar', 'Tadjikiskiy', 'Dushanba', 'Port va Senter'),
      ('', 'BEXRUZ OMAD NUR', '', 'Termiz shahar', 'Otan polvon yoni', 'Dushanba', 'Port va Senter'),
      ('', 'Abdujabbor Market', '', 'Termiz shahar', 'Senter bozor yonboshi', 'Dushanba', 'Port va Senter'),
      ('', 'Salmon nondog bekat', '', 'Termiz shahar', 'viloyat sudi orqasi', 'Dushanba', 'Port va Senter'),
      ('', 'Aznaf market', '', 'Termiz shahar', 'avtosalon roparasi', 'Dushanba', 'Port va Senter'),
      ('', 'Sher namuna', '', 'Termiz tuman', 'Senter kamaz', 'Dushanba', 'PSKUL'),

      # --- Seshanba ---
      ('', 'Vohas', '', 'Termiz shahar', 'Vohas', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Oʻzbegim market', '', 'Termiz shahar', 'Davr bank yoni', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Pokiza market', '', 'Termiz shahar', 'Muz saroy yoni', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Soxibkor doʻstlik', '', 'Termiz shahar', 'Hokimiyat', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Feruza non sex', '', 'Termiz shahar', 'Nigoh non sexi yoni', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'South brothers', '', 'Termiz shahar', 'Doktor A qatori', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Gulmira opa', '', 'Termiz shahar', 'Cola orqasi', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Taniqulov Oʻktam', '', 'Termiz shahar', 'Hokimiyat', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Abdulloh market', '', 'Termiz shahar', 'Feredun café', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', '23-market', '', 'Termiz shahar', '23-sartarosh yoni', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Abrayev Oral Market', '', 'Termiz shahar', '5-daha doktar A', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Roziyev Ismoil', '', 'Termiz shahar', 'Doktor A yoni', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Talaba Market', '', 'Termiz shahar', 'TerDU', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Abror aka Market', '', 'Termiz shahar', 'Gulom Muhammad', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Amin market', '', 'Termiz shahar', 'cola orqasi', 'Seshanba', 'Hokimiyat va 5-daha'),
      ('', 'Nyatloshka market', '', 'Termiz shahar', 'naytloshka yoni', 'Seshanba', 'Hokimiyat va 5-daha'),

      # --- Chorshanba ---
      ('', 'Muxlis Market', '', 'Termiz shahar', 'Movaro', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Anjir Market', '', 'Termiz shahar', 'Movaro', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Abbos market', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Dilya opa', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Sanjar aka', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Nur market', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', '547 market', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Shox market', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Makro market', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro'),
      ('', 'Xusan bobo market', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Fresh market umid aka', '', 'Termiz shahar', 'Movaro', 'Chorshanba', 'Movaro'),
      ('', '7Я Market', '', 'Termiz shahar', 'indenim', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Jahon Market', '', 'Termiz shahar', 'indenim', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Abidov Jamshid', '', 'Termiz shahar', 'Yashil dunyo Tisu yoni', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'NORMAMATOV BAHODIR', '', 'Termiz shahar', 'YEVRO MARTDAN KEYINGI', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Sveta opa', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', '1-dokon', '', 'Termiz shahar', 'Yashil dunyo', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', '777 market', '', 'Termiz shahar', 'Movaro', 'Chorshanba', 'Movaro'),
      ('', 'Yodgorov Abduhakim', '', 'Termiz shahar', 'Yashil dunyo ichi', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Malikov Eldor', '', 'Termiz shahar', 'Movaro', 'Chorshanba', 'Movaro va Yashil Dunyo'),
      ('', 'Alisherbek Xoltora', '', 'Termiz shahar', 'Tisuga yetmay', 'Chorshanba', 'Movaro va Yashil Dunyo'),

      # --- Payshanba ---
      ('', 'Pul hokim', '', 'Termiz shahar', 'Adliya yoʻli', 'Payshanba', '16-daha'),
      ('', 'Hoji ona market', '', 'Termiz shahar', 'Med yoni', 'Payshanba', '16-daha'),
      ('', 'Chinor market', '', 'Termiz shahar', 'Med yoni', 'Payshanba', '16-daha'),
      ('', '16 market', '', 'Termiz shahar', 'Begoyim roʻparasi', 'Payshanba', '16-daha'),
      ('YALKAPOVA DILAFRUZ', 'Lada yoni', '', 'Termiz shahar', 'Lada yoni', 'Payshanba', '16-daha'),
      ('', 'Alibek aka', '', 'Termiz shahar', 'Sohil pastlik', 'Payshanba', '16-daha'),
      ('', '4 aka-uka', '', 'Termiz shahar', 'sohil', 'Payshanba', '16-daha'),
      ('', 'Rayxon opa sohil', '', 'Termiz shahar', 'Guliston ma-si', 'Payshanba', '16-daha'),
      ('', 'Muhabbat opa', '', 'Termiz shahar', 'Boyqishloq', 'Payshanba', '16-daha'),
      ('', 'Black bekat', '', 'Termiz shahar', 'texnikum yoni', 'Payshanba', '16-daha'),
      ('', 'Xonadon Market', '', 'Termiz shahar', '16-daha', 'Payshanba', '16-daha'),
      ('', 'Shuhrat aka', '', 'Termiz shahar', 'viloyat raddomi', 'Payshanba', '16-daha'),
      ('', 'Jomiy bekat', '', 'Termiz shahar', 'Manguzar', 'Payshanba', '16-daha'),
      ('', 'Mancho market', '', 'Termiz shahar', 'Bochka yoli', 'Payshanba', '16-daha'),

      # --- Juma ---
      ('', 'Moyka yoni', '', 'Termiz shahar', 'Moyka yoni', 'Juma', 'Yubileyniy'),
      ('', 'Malika yoni optom', '', 'Termiz shahar', 'Malika yoni optom', 'Juma', 'Yubileyniy'),
      ('', 'AR market', '', 'Termiz shahar', 'Abdurashid market', 'Juma', 'Yubileyniy'),
      ('URALOV ZUHRIDDIN', 'HALOL MARKET', '', 'Termiz shahar', '23-dom', 'Juma', 'Yubileyniy'),
      ('', 'Osiyo tagi', '', 'Termiz shahar', '', 'Juma', 'Yubileyniy'),
      ('', 'Nigora opa', '', 'Termiz shahar', 'Eski pioner oldi', 'Juma', 'Yubileyniy'),
      ('', 'Sherbek aka/fresh M', '', 'Termiz shahar', 'Best roʻparasi', 'Juma', 'Yubileyniy'),
      ('', 'Otabek aka', '', 'Termiz shahar', 'antena tagi', 'Juma', 'Yubileyniy'),
      ('', 'Oʻzbegim market', '', 'Termiz shahar', 'Davr bank yoni', 'Juma', 'Yubileyniy'),
      ('', 'Universal Market', '', 'Termiz shahar', 'Lola kafe yoni', 'Juma', 'Yubileyniy'),
      ('', '878 Market', '', 'Termiz shahar', 'Kapalak domi roparasi', 'Juma', 'Yubileyniy'),
      ('', 'Bobayorova Sayyora', '', 'Termiz shahar', '', 'Juma', 'Yubileyniy'),
      ('Uralov Zuhriddin', 'Uralov zuhriddin', '', 'Termiz shahar', 'Yubileyniy lola', 'Juma', 'Yubileyniy'),
      ('', 'Oila nur market', '', 'Termiz shahar', 'office market orqasi', 'Juma', 'Yubileyniy'),
      ('', 'Hojanova Rano', '', 'Termiz shahar', 'Savhozniy', 'Juma', 'Yubileyniy'),
      ('', 'Toyirova Nigora', '', 'Termiz shahar', 'Milliy gvardiya yoni', 'Juma', 'Yubileyniy'),

      # --- Shanba ---
      ('', 'Aka-Uka Market', '', 'Termiz shahar', 'Masjid yoni', 'Shanba', 'Avtokvokzal'),
      ('', 'Oila Market', '', 'Termiz shahar', 'zilyonni yoʻli', 'Shanba', 'Avtokvokzal'),
      ('', 'Boxo market', '', 'Termiz shahar', 'zilyonni yoʻli', 'Shanba', 'Avtokvokzal'),
      ('', 'Otabek zapchast M', '', 'Termiz shahar', 'zilyonni yoʻli', 'Shanba', 'Avtokvokzal'),
      ('', 'Maya market', '', 'Termiz shahar', 'Gostsatndart yoni', 'Shanba', 'Avtokvokzal'),
      ('', 'Umida opa', '', 'Termiz shahar', 'Mashhura yoʻli', 'Shanba', 'Avtokvokzal'),
      ('', 'Darband city', '', 'Termiz shahar', 'Vokzal yoni', 'Shanba', 'Avtokvokzal'),
      ('', 'Otajon market', '', 'Termiz shahar', 'Tisudan keyin', 'Shanba', 'Avtokvokzal'),
      ('', 'Barakali market', '', 'Termiz shahar', 'Tisu yoni', 'Shanba', 'Avtokvokzal'),
      ('', 'Asilabonu', '', 'Termiz shahar', 'Boysun bekati yoni', 'Shanba', 'Avtokvokzal'),
      ('', 'Norqulova Nargiza', '', 'Termiz shahar', 'Hayit ala uyi taraf', 'Shanba', 'Avtokvokzal'),
      ('', 'Sherzod aka', '', 'Termiz shahar', 'Harbiy doʻkon', 'Shanba', 'Avtokvokzal'),
      ('', 'Diyorova Muborak', '', 'Termiz shahar', 'Zagazerno', 'Shanba', 'Avtokvokzal'),
      ('', 'Lokomativ market-depo', '', 'Termiz shahar', 'Zagazerno', 'Shanba', 'Avtokvokzal'),
      ('', 'ALLAMURODOVA IQBOL', '', 'Termiz shahar', 'TURON MAHALLADAN KIRGANDA 100 METR TOGRIGA', 'Shanba', 'Avtokvokzal'),
      ('', 'Crafers bekat', '', 'Termiz shahar', 'eski pioner roparasi', 'Shanba', 'Avtokvokzal'),
      ('Madiyeva Iroda', 'MAdiyeva I market', '', 'Termiz shahar', 'avtovokzal litsey', 'Shanba', 'Avtokvokzal'),
      ('', 'Yulduz bekat', '', 'Termiz shahar', 'shox bekat yoli', 'Shanba', 'Avtokvokzal'),
  ]
  for s_firm, s_name, s_phone, s_region, s_landmark, s_visit_days, s_territory in initial_shops:
    cursor.execute(
        """INSERT OR IGNORE INTO shops (firm, name, phone, debt, region, landmark, visit_days, territory) 
             VALUES (?, ?, ?, 0, ?, ?, ?, ?)""",
        (s_firm, s_name, s_phone, s_region, s_landmark, s_visit_days, s_territory),
    )

  initial_products = [
      ('Pelmen /300 gr', 'Yarim Tayyor Mahsulotlari', 100, 6200, 14000.0, 14000.0, 5),
      ('Pelmen /500 gr', 'Yarim Tayyor Mahsulotlari', 100, 9900, 24000.0, 24000.0, 5),
      ('Pelmen rasepnoy /kg', 'Yarim Tayyor Mahsulotlari', 100, 19500, 46000.0, 46000.0, 5),
      ('Teftel /300 gr', 'Yarim Tayyor Mahsulotlari', 100, 12000, 23000.0, 23000.0, 5),
      ('Pelmen ossarti /500 gr', 'Yarim Tayyor Mahsulotlari', 100, 18000, 30000.0, 30000.0, 5),
      ('Pelmen ossarti /300 gr', 'Yarim Tayyor Mahsulotlari', 100, 11500, 20000.0, 20000.0, 5),
      ('Golubtsi /500 gr', 'Yarim Tayyor Mahsulotlari', 100, 14000, 25000.0, 25000.0, 5),
      ('Tok doʻlma /300 gr', 'Yarim Tayyor Mahsulotlari', 100, 11000.0, 25000.0, 25000.0, 5),
      ('Karam doʻlma /300 gr', 'Yarim Tayyor Mahsulotlari', 100, 11000.0, 25000.0, 25000.0, 5),
      ('Somsa kesilgan /800 gr', 'Yarim Tayyor Mahsulotlari', 100, 6000.0, 17000.0, 17000.0, 5),
      ('Oʻrama xamir', 'Yarim Tayyor Mahsulotlari', 100, 6000.0, 17000.0, 17000.0, 5),
      ('KFC XA', 'Yarim Tayyor Mahsulotlari', 100.0, 15000, 30000.0, 30000.0, 5),
      ('Osh masalliq 500 gr', 'Yarim Tayyor Mahsulotlari', 100, 6000.0, 13000.0, 13000.0, 5),
      ('Osh masalliq 1 kg', 'Yarim Tayyor Mahsulotlari', 100, 7000.0, 15000.0, 15000.0, 5),
      ('Lagʻmon', 'Yarim Tayyor Mahsulotlari', 0, 2000.0, 6000.0, 6000.0, 5),
      ('Kotlet', 'Yarim Tayyor Mahsulotlari', 100, 11000.0, 25000.0, 25000.0, 5),
      ('Lavash hamiri', 'Yarim Tayyor Mahsulotlari', 100, 3800.0, 7000.0, 7000.0, 5),
      ('Manti hamiri', 'Yarim Tayyor Mahsulotlari', 100, 6000.0, 13000.0, 13000.0, 5),
      ('Farsh', 'Yarim Tayyor Mahsulotlari', 100, 15500.0, 19500.0, 19500.0, 5),
      ('Mirinda 250gr/30шт', 'Yaxna ichimliklari', 460, 7950.0, 9000.0, 9000.0, 3),
      ('Mirinda 330gr/24шт', 'Yaxna ichimliklari', 408, 8800.0, 10000.0, 10000.0, 3),
      ('Snikers/001/3kg', 'Afif shirinliklari', 100, 100500.0, 117000.0, 117000.0, 3),
      ('Snikers oq/041/3kg', 'Afif shirinliklari', 100, 100500.0, 117000.0, 117000.0, 3),
      ('Mini Rulet/002/3kg', 'Afif shirinliklari', 20, 100500.0, 120000.0, 120000.0, 3),
      ('Yojik/003/2kg', 'Afif shirinliklari', 20, 67000.0, 84000.0, 84000.0, 3),
      ('Dessert/007/2kg', 'Afif shirinliklari', 20, 71000.0, 86000.0, 86000.0, 3),
      ('Pudra Palichka/013/2kg', 'Afif shirinliklari', 10, 49000.0, 60000.0, 60000.0, 3),
      ('Pudra Kalso/014/2kg', 'Afif shirinliklari', 10, 49000.0, 60000.0, 60000.0, 3),
      ('Ovsyanka pista/025/3kg', 'Afif shirinliklari', 10, 73500.0, 92000.0, 92000.0, 3),
      ('Ovsyanka magiz/026/3kg', 'Afif shirinliklari', 10, 73500.0, 92000.0, 92000.0, 3),
      ('Choko Ovsyanka/028/3kg', 'Afif shirinliklari', 10, 82500.0, 103000.0, 103000.0, 3),
  ]
  for p_name, p_cat, p_stock, p_cost, p_optom, p_chakana, p_percent in initial_products:
    cursor.execute(
        '''INSERT OR IGNORE INTO products (name, category, stock, cost_price, optom_price, chakana_price, agent_percent) 
           VALUES (?, ?, ?, ?, ?, ?, ?)''',
        (p_name, p_cat, p_stock, p_cost, p_optom, p_chakana, p_percent),
    )
  conn.commit()
  conn.close()

def create_excel_invoice(
    order_id, shop_name, agent_name, date_str, price_type, cart_items, comment=''
):
  wb = Workbook()
  ws = wb.active
  ws.title = f'Nakladnoy_{order_id}'
  ws.sheet_view.showGridLines = True

  title_font = Font(name='Arial', size=16, bold=True)
  header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
  bold_font = Font(name='Arial', size=11, bold=True)
  header_fill = PatternFill(
      start_color='1F497D', end_color='1F497D', fill_type='solid'
  )
  total_fill = PatternFill(
      start_color='DCE6F1', end_color='DCE6F1', fill_type='solid'
  )
  thin_border = Border(
      left=Side(style='thin', color='B0B0B0'),
      right=Side(style='thin', color='B0B0B0'),
      top=Side(style='thin', color='B0B0B0'),
      bottom=Side(style='thin', color='B0B0B0'),
  )

  ws.merge_cells('A1:E1')
  ws['A1'] = 'XOJI AKA FACTORY — NAKLADNOY'
  ws['A1'].font = title_font
  ws['A1'].alignment = Alignment(horizontal='center')

  ws['A3'] = f'Buyurtma ID: #{order_id}'
  ws['A3'].font = bold_font
  ws['D3'] = f'Sana: {date_str}'
  ws['A4'] = f"Do'kon (Klient): {shop_name}"
  ws['D4'] = f'Narx turi: {price_type.upper()}'
  ws['A5'] = f'Agent: {agent_name}'
  if comment:
    ws['A6'] = f'Izoh (Kommentariya): {comment}'
    ws['A6'].font = bold_font

  start_row = 8 if comment else 7

  headers = ['№', 'Mahsulot nomi', "Miqdori", "Narxi (so'm)", 'Jami summa']
  for col_num, header_title in enumerate(headers, 1):
    cell = ws.cell(row=start_row, column=col_num, value=header_title)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

  row_num = start_row + 1
  total_sum = 0
  for idx, item in enumerate(cart_items, 1):
    ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(
        horizontal='center'
    )
    ws.cell(row=row_num, column=2, value=item['name']).alignment = Alignment(
        horizontal='left'
    )
    ws.cell(row=row_num, column=3, value=item['qty']).alignment = Alignment(
        horizontal='right'
    )
    ws.cell(row=row_num, column=4, value=item['price']).alignment = Alignment(
        horizontal='right'
    )
    summa = item['qty'] * item['price']
    total_sum += summa
    ws.cell(row=row_num, column=5, value=summa).alignment = Alignment(
        horizontal='right'
    )
    ws.cell(row=row_num, column=4).number_format = '#,##0'
    ws.cell(row=row_num, column=5).number_format = '#,##0'
    for col in range(1, 6):
      ws.cell(row=row_num, column=col).border = thin_border
    row_num += 1

  ws.merge_cells(
      start_row=row_num, start_column=1, end_row=row_num, end_column=4
  )
  ws.cell(row=row_num, column=1, value="JAMI TO'LOV:").alignment = Alignment(
      horizontal='right'
  )
  ws.cell(row=row_num, column=1).font = bold_font
  total_val = ws.cell(row=row_num, column=5, value=total_sum)
  total_val.font = bold_font
  total_val.number_format = '#,##0'
  for col in range(1, 6):
    ws.cell(row=row_num, column=col).fill = total_fill
    ws.cell(row=row_num, column=col).border = thin_border

  row_num += 2
  ws.cell(row=row_num, column=2, value='Qabul qildim: ____')
  ws.cell(row=row_num, column=4, value='Topshirdim: ____')

  file_name = f'Nakladnoy_{order_id}.xlsx'
  wb.save(file_name)
  return file_name


def create_excel_sex_income(income_id, staff_name, date_str, cart_items):
  wb = Workbook()
  ws = wb.active
  ws.title = f'Sex_Kirim_{income_id}'
  ws.sheet_view.showGridLines = True

  title_font = Font(name='Arial', size=16, bold=True)
  header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
  bold_font = Font(name='Arial', size=11, bold=True)
  header_fill = PatternFill(
      start_color='006100', end_color='006100', fill_type='solid'
  )
  total_fill = PatternFill(
      start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'
  )
  thin_border = Border(
      left=Side(style='thin', color='B0B0B0'),
      right=Side(style='thin', color='B0B0B0'),
      top=Side(style='thin', color='B0B0B0'),
      bottom=Side(style='thin', color='B0B0B0'),
  )

  ws.merge_cells('A1:D1')
  ws['A1'] = '🏭 SEXGA MAHSULOT KIRIM (SKLAD)'
  ws['A1'].font = title_font
  ws['A1'].alignment = Alignment(horizontal='center')

  ws['A3'] = f'Kirim ID: #{income_id}'
  ws['A3'].font = bold_font
  ws['C3'] = f'Sana: {date_str}'
  ws['A4'] = f'Mas’ul xodim: {staff_name}'

  headers = ['№', 'Mahsulot nomi', "Miqdori (Soni/Kg)"]
  for col_num, header_title in enumerate(headers, 1):
    cell = ws.cell(row=6, column=col_num, value=header_title)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

  row_num = 7
  total_qty = 0
  for idx, item in enumerate(cart_items, 1):
    ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(
        horizontal='center'
    )
    ws.cell(row=row_num, column=2, value=item['name']).alignment = Alignment(
        horizontal='left'
    )
    ws.cell(row=row_num, column=3, value=item['qty']).alignment = Alignment(
        horizontal='right'
    )
    total_qty += item['qty']
    ws.cell(row=row_num, column=3).number_format = '#,##0.##'
    for col in range(1, 4):
      ws.cell(row=row_num, column=col).border = thin_border
    row_num += 1

  ws.merge_cells(
      start_row=row_num, start_column=1, end_row=row_num, end_column=2
  )
  ws.cell(row=row_num, column=1, value='JAMI QO\'SHILGAN MIQDOR:').alignment = (
      Alignment(horizontal='right')
  )
  ws.cell(row=row_num, column=1).font = bold_font
  total_val = ws.cell(row=row_num, column=3, value=total_qty)
  total_val.font = bold_font
  total_val.number_format = '#,##0.##'
  for col in range(1, 4):
    ws.cell(row=row_num, column=col).fill = total_fill
    ws.cell(row=row_num, column=col).border = thin_border

  file_name = f'Sex_Kirim_{income_id}.xlsx'
  wb.save(file_name)
  return file_name


def get_main_menu(role):
  markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
  if role == 'admin' or 'admin' in (role or '').lower():
    markup.row(
        types.KeyboardButton('📦 Sklad & Mahsulotlar'),
        types.KeyboardButton('📊 Kunlik Hisobot'),
    )
    markup.row(
        types.KeyboardButton('👥 Agentlar boshqaruvi'),
        types.KeyboardButton("🏪 AKB (Do'konlar & Qarz)"),
    )
    markup.row(types.KeyboardButton('💳 Qarzdorlik boshqaruvi'), types.KeyboardButton('📥 Kirim qo‘shish'))
  elif role == 'agent':
    markup.row(
        types.KeyboardButton('🛒 Yangi Buyurtma Urish'),
        types.KeyboardButton('📊 Oylik hisobim'),
    )
    markup.row(
        types.KeyboardButton('📜 Mening Buyurtmalarim'),
        types.KeyboardButton("📊 Do'konlar qarzi"),
    )
    markup.row(
        types.KeyboardButton("📋 Status (Yetkazilmaganlar)"),
    )
    markup.row(types.KeyboardButton('🌐 Web kabinetim'), types.KeyboardButton('🖼 Mahsulotlar katalogi'))
    markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  elif role == 'sex':
    markup.row(
        types.KeyboardButton('📦 Skladga Mahsulot Kirim Qilish'),
        types.KeyboardButton('📋 Ombordagi Qoldiqlar'),
    )
    markup.row(types.KeyboardButton('🌐 Web kabinetim'))
    markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  elif role == 'zavsklad':
    markup.row(
        types.KeyboardButton('📦 Yig‘ilgan Yaxlit Buyurtmalar'),
        types.KeyboardButton('📋 Ombordagi Qoldiqlar'),
    )
    markup.row(types.KeyboardButton('🌐 Web kabinetim'))
    markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  elif role == 'ekspeditor':
    markup.row(
        types.KeyboardButton('🚚 Yetkazilishi Kerak Buyurtmalar'),
        types.KeyboardButton('📜 Yetkazilgan Buyurtmalar Tarixi'),
    )
    markup.row(
        types.KeyboardButton("💰 Qarz/To'lov yozish (Ekspeditor)"),
        types.KeyboardButton("📊 Do'konlar qarzi"),
    )
    markup.row(types.KeyboardButton('🌐 Web kabinetim'))
    markup.row(
        types.KeyboardButton("📋 Status (Yetkazilmaganlar)"),
    )
    markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  else:
    markup.add(types.KeyboardButton("📝 Ro'yxatdan o'tish"))
  return markup


@bot.message_handler(commands=['start'])
def start_command(message):
  tg_id = message.from_user.id
  if is_admin_tg(tg_id):
    bot.send_message(
        message.chat.id,
        'Xoji aka, xush kelibsiz! Boshqaruv paneli tayyor.',
        reply_markup=get_main_menu('admin'),
    )
    return
  conn = get_db_connection()
  user = conn.execute(
      'SELECT role, name FROM users WHERE tg_id = ?', (tg_id,)
  ).fetchone()
  conn.close()
  if user:
    role, name = user['role'], user['name']
    if role == 'pending':
      bot.send_message(
          message.chat.id,
          f"Salom {name}. So'rovingiz admin tasdig'ini kutyapti.",
      )
    elif ',' in role or len(role.split(',')) > 1 or role in ['agent,sex', 'zavsklad', 'ekspeditor']:
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      if 'agent' in role:
        markup.add(types.KeyboardButton('👤 Agent rejimi'))
      if 'sex' in role:
        markup.add(types.KeyboardButton('🏭 Sex rejimi'))
      if 'zavsklad' in role or role == 'zavsklad':
        markup.add(types.KeyboardButton('📦 Zavsklad rejimi'))
      if 'ekspeditor' in role or role == 'ekspeditor':
        markup.add(types.KeyboardButton('🚚 Ekspeditor rejimi'))
      bot.send_message(
          message.chat.id,
          f'Salom {name}! Iltimos, ish rejimini tanlang:',
          reply_markup=markup,
      )
    else:
      bot.send_message(
          message.chat.id,
          f'Salom {name}! Ishni boshlashimiz mumkin.',
          reply_markup=get_main_menu(role),
      )
  else:
    bot.send_message(
        message.chat.id,
        "Assalomu alaykum! Tizimga xush kelibsiz. Davom etish uchun"
        " ro'yxatdan o'ting.",
        reply_markup=get_main_menu('guest'),
    )


@bot.message_handler(func=lambda message: message.text == "📊 Oylik hisobim")
def agent_salary_report_bot(message):
    tg_id = message.from_user.id
    conn = get_db_connection()
    user = conn.execute("SELECT name FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    if not user:
        conn.close()
        bot.send_message(message.chat.id, "Foydalanuvchi topilmadi.")
        return
    agent_name = user['name']
    
    agent_orders = conn.execute("SELECT items_text, total_sum FROM orders WHERE agent_name = ? AND status != 'Bekor'", (agent_name,)).fetchall()
    calculated_salary = 0
    for ao in agent_orders:
        for line in ao['items_text'].split('\n'):
            if line.strip():
                try:
                    parts = line.split(' - ')
                    p_name = parts[0].strip()
                    qty_part = parts[1].split('x')[0].strip()
                    qty = float(qty_part)
                    prod_info = conn.execute("SELECT optom_price, agent_percent FROM products WHERE name = ?", (p_name,)).fetchone()
                    if prod_info:
                        p_price = prod_info['optom_price'] or 0
                        p_perc = prod_info['agent_percent'] or 5
                        item_total = p_price * qty
                        calculated_salary += (item_total * p_perc / 100)
                except:
                    pass
    
    adv_res = conn.execute("SELECT SUM(amount) as s FROM advances WHERE worker_name = ?", (agent_name,)).fetchone()
    total_adv = adv_res['s'] if adv_res and adv_res['s'] else 0
    final_salary = calculated_salary - total_adv
    conn.close()
    
    text = (
        f"📊 **Sizning oylik hisobotingiz:**\n\n"
        f"👤 Agent: {agent_name}\n"
        f"💰 Hisoblangan foiz oylik: **{calculated_salary:,.0f} so'm**\n"
        f"📤 Olingan avanslar: **{total_adv:,.0f} so'm**\n"
        f"💵 **Yakuniy oylik (Qoldiq): {final_salary:,.0f} so'm**"
    )
    bot.send_message(message.chat.id, text, parse_mode="Markdown")


@bot.message_handler(func=lambda message: message.text == "📊 Do'konlar qarzi")
def view_shops_debt_only(message):
    conn = get_db_connection()
    shops = conn.execute("SELECT name, debt, region FROM shops WHERE debt > 0 ORDER BY debt DESC").fetchall()
    conn.close()
    
    if not shops:
        bot.send_message(message.chat.id, "✅ Hozirda qarzi bor do'konlar mavjud emas (Barchaning qarzi 0).")
        return
        
    text = "🏪 **Qarzi bor do'konlar ro'yxati:**\n\n"
    for s in shops:
        text += f"🔹 {s['name']} ({s['region']}) — **{s['debt']:,.0f} so'm**\n"
    bot.send_message(message.chat.id, text, parse_mode="Markdown")


@bot.message_handler(func=lambda message: message.text == "📋 Status (Yetkazilmaganlar)")
def view_undelivered_orders_only(message):
    conn = get_db_connection()
    orders = conn.execute("SELECT id, shop_name, total_sum, status, date FROM orders WHERE status != 'Yetkazildi' ORDER BY id DESC").fetchall()
    conn.close()
    
    if not orders:
        bot.send_message(message.chat.id, "✅ Hozirda yetkazilmagan ochiq buyurtmalar yo'q.")
        return
        
    text = "📋 **Hali yetkazilmagan buyurtmalar (Statuslar):**\n\n"
    for o in orders:
        text += f"🆔 **#{o['id']}** | 🏪 {o['shop_name']}\nStatus: **{o['status']}** | Summa: {o['total_sum']:,.0f} so'm\nSana: {o['date']}\n-------------------\n"
    bot.send_message(message.chat.id, text, parse_mode="Markdown")


@bot.message_handler(func=lambda message: message.text == "💰 Qarz/To'lov yozish (Ekspeditor)")
def ekspeditor_pay_debt_start(message):
  conn = get_db_connection()
  shops = conn.execute('SELECT name FROM shops').fetchall()
  conn.close()
  if not shops:
    bot.send_message(message.chat.id, "❌ Do'konlar yo'q.")
    return
  markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
  for s in shops:
    markup.add(types.KeyboardButton(s['name']))
  markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  msg = bot.send_message(message.chat.id, 'Do\'konni tanlang:', reply_markup=markup)
  bot.register_next_step_handler(msg, ekspeditor_ask_payment_amount)


def ekspeditor_ask_payment_amount(message):
  if message.text == '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)':
    switch_role_menu(message)
    return
  shop_name = message.text
  user_steps[message.from_user.id] = {'pay_shop': shop_name}
  msg = bot.send_message(
      message.chat.id,
      f"'{shop_name}' uchun qancha to'lov kirdi (summa):",
      reply_markup=types.ReplyKeyboardRemove(),
  )
  bot.register_next_step_handler(msg, ekspeditor_process_payment_request)


def ekspeditor_process_payment_request(message):
  uid = message.from_user.id
  try:
    amount = float(message.text)
    shop_name = user_steps.get(uid, {}).get('pay_shop', 'Nomalum')
    
    conn = get_db_connection()
    user_res = conn.execute('SELECT name FROM users WHERE tg_id = ?', (uid,)).fetchone()
    ekspeditor_name = user_res['name'] if user_res else 'Ekspeditor'
    bugun = datetime.now().strftime('%Y-%m-%d %H:%M')

    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO payment_requests (shop_name, ekspeditor_name, amount, status, date) VALUES (?, ?, ?, 'pending', ?)",
        (shop_name, ekspeditor_name, amount, bugun)
    )
    req_id = cursor.lastrowid
    conn.commit()
    conn.close()

    bot.send_message(
        message.chat.id,
        f"✅ To'lov so'rovi (#{req_id}) adminga yuborildi. Admin tasdiqlagach do'kon qarzidan ayriladi.",
        reply_markup=get_main_menu('ekspeditor'),
    )

    admin_markup = types.InlineKeyboardMarkup()
    admin_markup.row(
        types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"approve_pay_{req_id}"),
        types.InlineKeyboardButton("❌ Bekor qilish", callback_data=f"reject_pay_{req_id}")
    )
    bot.send_message(
        ADMIN_ID,
        f"💰 **YANGI TO'LOV SO'ROVI (#{req_id})**\n\n"
        f"🏪 Do'kon: {shop_name}\n"
        f"👤 Ekspeditor: {ekspeditor_name}\n"
        f"💵 Summa: **{amount:,.0f} so'm**\n"
        f"📅 Sana: {bugun}",
        parse_mode="Markdown",
        reply_markup=admin_markup
    )
  except Exception as e:
    print("To'lov yozish xatosi:", e)
    bot.send_message(message.chat.id, "❌ Xatolik! Faqat raqam kiriting.", reply_markup=get_main_menu('ekspeditor'))


@bot.callback_query_handler(func=lambda call: call.data.startswith(('approve_pay_', 'reject_pay_')))
def admin_handle_payment_request(call):
    action, _, req_id_str = call.data.split('_', 2)
    req_id = int(req_id_str)

    conn = get_db_connection()
    req = conn.execute("SELECT * FROM payment_requests WHERE id = ?", (req_id,)).fetchone()
    
    if not req:
        bot.answer_callback_query(call.id, "So'rov topilmadi yoki allaqachon ko'rib chiqilgan!")
        conn.close()
        return

    if req['status'] != 'pending':
        bot.answer_callback_query(call.id, "Bu so'rov allaqachon bajarilgan!")
        conn.close()
        return

    shop_name = req['shop_name']
    amount = req['amount']
    today = datetime.now().strftime('%Y-%m-%d %H:%M')

    if action == 'approve':
        s_res = conn.execute("SELECT debt FROM shops WHERE name = ?", (shop_name,)).fetchone()
        curr_debt = s_res['debt'] if s_res and s_res['debt'] is not None else 0

        conn.execute("UPDATE shops SET debt = ? WHERE name = ?", (curr_debt - amount, shop_name))
        conn.execute("INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)", (f"Qarz to'lovi ({shop_name})", amount, today))
        conn.execute("UPDATE payment_requests SET status = 'approved' WHERE id = ?", (req_id,))
        conn.commit()
        
        bot.answer_callback_query(call.id, "To'lov tasdiqlandi, qarz chegirildi va kassaga yozildi!")
        bot.edit_message_text(f"✅ **Tasdiqlandi va bajarildi!** (#{req_id})\n🏪 Do'kon: {shop_name}\n💵 Summa: {amount:,.0f} so'm", call.message.chat.id, call.message.message_id, parse_mode="Markdown")
    else:
        conn.execute("UPDATE payment_requests SET status = 'rejected' WHERE id = ?", (req_id,))
        conn.commit()
        bot.answer_callback_query(call.id, "To'lov so'rovi rad etildi!")
        bot.edit_message_text(f"❌ **Rad etildi!** (#{req_id})\n🏪 Do'kon: {shop_name}\n💵 Summa: {amount:,.0f} so'm", call.message.chat.id, call.message.message_id, parse_mode="Markdown")

    conn.close()


@bot.message_handler(
    func=lambda message: message.text
    in ['👤 Agent rejimi', '🏭 Sex rejimi', '📦 Zavsklad rejimi', '🚚 Ekspeditor rejimi', '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)']
)
def switch_role_menu(message):
  tg_id = message.from_user.id
  conn = get_db_connection()
  user = conn.execute(
      'SELECT role, name FROM users WHERE tg_id = ?', (tg_id,)
  ).fetchone()
  conn.close()

  if not user:
    return

  if message.text == '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)':
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      markup.add(
          types.KeyboardButton('👤 Agent rejimi'),
          types.KeyboardButton('🏭 Sex rejimi'),
          types.KeyboardButton('📦 Zavsklad rejimi'),
          types.KeyboardButton('🚚 Ekspeditor rejimi'),
      )
      bot.send_message(
          message.chat.id, 'Qaysi rejimga oʻtmoqchisiz?', reply_markup=markup
      )
      return

  if message.text == '👤 Agent rejimi':
    bot.send_message(message.chat.id, '🛒 Agent rejimiga oʻtdingiz.', reply_markup=get_main_menu('agent'))
  elif message.text == '🏭 Sex rejimi':
    bot.send_message(message.chat.id, '🏭 Sex rejimiga oʻtdingiz.', reply_markup=get_main_menu('sex'))
  elif message.text == '📦 Zavsklad rejimi':
    bot.send_message(message.chat.id, '📦 Zavsklad rejimiga oʻtdingiz.', reply_markup=get_main_menu('zavsklad'))
  elif message.text == '🚚 Ekspeditor rejimi':
    bot.send_message(message.chat.id, '🚚 Ekspeditor rejimiga oʻtdingiz.', reply_markup=get_main_menu('ekspeditor'))


@bot.message_handler(func=lambda message: message.text == '📦 Yig‘ilgan Yaxlit Buyurtmalar')
def zavsklad_view_orders(message):
    conn = get_db_connection()
    orders = conn.execute("SELECT id, shop_name, total_sum, status, items_text, date FROM orders WHERE TRIM(COALESCE(status,'')) = 'Yangi' ORDER BY date ASC, id ASC").fetchall()
    conn.close()
    if not orders:
        bot.send_message(message.chat.id, "Hozirda yig'ish uchun Yangi buyurtmalar mavjud emas.")
        return
    
    text = "📦 **Yangi tushgan buyurtmalar (Yig'ish uchun):**\n\n"
    for o in orders:
        text += f"🆔 **#{o['id']}** | {o['date']} | Do'kon: {o['shop_name']} | Jami: {safe_number(o['total_sum']):,.0f} so'm\n"
    
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("🔗 Barcha yangi buyurtmalarni bitta yaxlit qilish va guruhga yuborish", callback_data="zav_combine_orders"))
    bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == "zav_combine_orders")
def zavsklad_combine_and_send(call):
    conn = get_db_connection()
    orders = conn.execute("SELECT * FROM orders WHERE TRIM(COALESCE(status,'')) = 'Yangi' ORDER BY date ASC, id ASC").fetchall()
    if not orders:
        bot.answer_callback_query(call.id, "Yig'ish uchun buyurtmalar yo'q!")
        conn.close()
        return
    
    combined_items = {}
    total_general_sum = 0
    order_ids = []
    
    for o in orders:
        order_ids.append(str(o['id']))
        total_general_sum += safe_number(o['total_sum'])
        for line in (o['items_text'] or '').split('\n'):
            if line.strip():
                parts = line.split(' - ')
                p_name = parts[0]
                try:
                    qty_str = parts[1].split('x')[0].strip()
                    qty = float(qty_str)
                    combined_items[p_name] = combined_items.get(p_name, 0) + qty
                except:
                    pass

    cart_items_list = [{'name': k, 'qty': v, 'price': 0} for k, v in combined_items.items()]
    bugun = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    excel_file = create_excel_invoice("Yaxlit_" + "_".join(order_ids), "Barcha Do'konlar", "Zavsklad", bugun, "optom", cart_items_list, "Yaxlit yig'ma buyurtma")
    
    items_text_str = ""
    for k, v in combined_items.items():
        items_text_str += f"🔹 {k} — {v} ta/kg\n"
        
    group_text = (
        f"📦 **YAXLIT YIG'MA BUYURTMA (Zavsklad tomonidan)**\n"
        f"📅 **Sana:** {bugun}\n"
        f"🔖 **Buyurtma ID lari:** {', '.join(order_ids)}\n\n"
        f"**Umumiy yig'ilgan mahsulotlar:**\n{items_text_str}\n"
        f"💰 **Jami umumiy summa:** {total_general_sum:,.0f} so'm"
    )
    
    try:
        bot.send_message(SEX_GROUP_ID, group_text, parse_mode="Markdown")
        with open(excel_file, 'rb') as doc_group:
            bot.send_document(SEX_GROUP_ID, doc_group, caption="📄 Yaxlit Yig'ma Nakladnoy (Excel)")
        bot.answer_callback_query(call.id, "Yaxlit buyurtma muvaffaqiyatli guruhga yuborildi!")
        bot.send_message(call.message.chat.id, "✅ Yaxlit buyurtma matn va Excel ko'rinishida guruhga yuborildi!")
        # Faqat yuborish muvaffaqiyatli bo‘lgandan keyin barcha qatnashgan buyurtmalarni yig‘ildi holatiga o‘tkazamiz.
        changed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for _oid in order_ids:
            conn.execute("UPDATE orders SET status = 'Otgruzka' WHERE id = ? AND TRIM(COALESCE(status,'')) = 'Yangi'", (int(_oid),))
            conn.execute('INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, ?, ?)', (int(_oid), 'Otgruzka', changed_at))
        log_audit(conn, 'Yaxlit buyurtmalar yuborildi', 'orders', ','.join(order_ids), f'soni={len(order_ids)}; sana_filteri=yo‘q')
        conn.commit()
    except Exception as e:
        print("Zavsklad yuborish xatosi:", e)
        bot.answer_callback_query(call.id, "Xatolik yuz berdi!")
        
    if os.path.exists(excel_file):
        os.remove(excel_file)
    conn.close()


@bot.message_handler(func=lambda message: message.text == '🚚 Yetkazilishi Kerak Buyurtmalar')
def ekspeditor_view_orders(message):
    conn = get_db_connection()
    # Ekspeditor ham "Yangi", ham "Otgruzka" statusidagi buyurtmalarni ko'ra oladi
    orders = conn.execute("SELECT id, shop_name, total_sum, items_text, status FROM orders WHERE status IN ('Yangi', 'Otgruzka') ORDER BY id DESC").fetchall()
    conn.close()
    if not orders:
        bot.send_message(message.chat.id, "Hozirda yetkazishga tayyor buyurtmalar yo'q.")
        return
    
    for o in orders:
        markup = types.InlineKeyboardMarkup()
        markup.add(types.InlineKeyboardButton("✅ Yetkazilganini Tasdiqlash", callback_data=f"eksp_deliver_{o['id']}"))
        text = (
            f"🚚 **Buyurtma ID: #{o['id']}**\n"
            f"🏪 Do'kon: {o['shop_name']}\n"
            f"📌 Status: **{o['status']}**\n"
            f"💰 Summa: {o['total_sum']:,.0f} so'm\n\n"
            f"Tafsilot:\n{o['items_text']}"
        )
        bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('eksp_deliver_'))
def ekspeditor_confirm_delivery(call):
    order_id = int(call.data.split('_')[-1])
    conn = get_db_connection()
    order = conn.execute("SELECT shop_name, total_sum, status FROM orders WHERE id = ?", (order_id,)).fetchone()
    
    if not order:
        bot.answer_callback_query(call.id, "Buyurtma topilmadi!")
        conn.close()
        return
        
    if order['status'] == 'Yetkazildi':
        bot.answer_callback_query(call.id, "Bu buyurtma allaqachon yetkazilgan!")
        conn.close()
        return
        
    shop_name = order['shop_name']
    order_sum = order['total_sum']
    
    conn.execute("UPDATE orders SET status = 'Yetkazildi' WHERE id = ?", (order_id,))
    
    shop_res = conn.execute("SELECT debt FROM shops WHERE name = ?", (shop_name,)).fetchone()
    curr_debt = shop_res['debt'] if shop_res and shop_res['debt'] is not None else 0
    new_debt = curr_debt + order_sum
    
    conn.execute("UPDATE shops SET debt = ? WHERE name = ?", (new_debt, shop_name))
    
    bugun = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn.execute("INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, 'Yetkazildi', ?)", (order_id, bugun))
    conn.commit()
    conn.close()
    
    bot.answer_callback_query(call.id, "Buyurtma yetkazildi deb tasdiqlandi va do'kon qarzdorligi yangilandi!")
    bot.edit_message_text(f"✅ **Yetkazildi va tasdiqlandi!** (#{order_id})\nDo'kon qarzdorligi hisoblandi.", call.message.chat.id, call.message.message_id, parse_mode="Markdown")


@bot.message_handler(func=lambda message: message.text == '📜 Yetkazilgan Buyurtmalar Tarixi')
def ekspeditor_history(message):
    conn = get_db_connection()
    orders = conn.execute("SELECT id, shop_name, total_sum, date FROM orders WHERE status = 'Yetkazildi' ORDER BY id DESC LIMIT 10").fetchall()
    conn.close()
    text = "📜 **Oxirgi yetkazilgan buyurtmalar:**\n\n"
    for o in orders:
        text += f"🆔 #{o['id']} | {o['shop_name']} | {o['total_sum']:,.0f} so'm | {o['date']}\n"
    bot.send_message(message.chat.id, text, parse_mode="Markdown")


@bot.message_handler(
    func=lambda message: message.text == '📦 Skladga Mahsulot Kirim Qilish'
)
def sex_income_start(message):
  conn = get_db_connection()
  prods = conn.execute('SELECT name FROM products').fetchall()
  conn.close()
  if not prods:
    bot.send_message(message.chat.id, '❌ Mahsulotlar topilmadi.')
    return

  user_steps[message.from_user.id] = {'sex_cart': {}}
  send_sex_product_menu(message)


def send_sex_product_menu(message):
  conn = get_db_connection()
  prods = conn.execute('SELECT name, stock FROM products').fetchall()
  conn.close()
  markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
  for p in prods:
    markup.add(types.KeyboardButton(p['name']))
  markup.add(types.KeyboardButton('✅ Kirimni yakunlash'))
  markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  msg = bot.send_message(
      message.chat.id,
      '🏭 Kirim qilinadigan mahsulotni tanlang:',
      reply_markup=markup,
  )
  bot.register_next_step_handler(msg, sex_income_choose_product)


def sex_income_choose_product(message):
  if message.text == '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)':
    switch_role_menu(message)
    return
  if message.text == '✅ Kirimni yakunlash':
    finish_sex_income(message)
    return

  conn = get_db_connection()
  p_check = conn.execute(
      'SELECT id FROM products WHERE name = ?', (message.text,)
  ).fetchone()
  conn.close()

  if not p_check:
    bot.send_message(
        message.chat.id, "❌ Bunday mahsulot topilmadi. Ro'yxatdan tanlang:"
    )
    send_sex_product_menu(message)
    return

  user_steps[message.from_user.id]['current_product'] = message.text
  msg = bot.send_message(
      message.chat.id,
      f"🔢 <b>{message.text}</b> uchun miqdorni (kg / dona) kiriting:",
      parse_mode='HTML',
      reply_markup=types.ReplyKeyboardRemove(),
  )
  bot.register_next_step_handler(msg, sex_income_add_to_cart)


def sex_income_add_to_cart(message):
  uid = message.from_user.id
  try:
    qty = float(message.text)
    p_name = user_steps[uid]['current_product']
    user_steps[uid]['sex_cart'][p_name] = qty
    bot.send_message(
        message.chat.id, f"📥 Qo'shildi: <b>{p_name}</b> — <b>{qty}</b>"
    )
    send_sex_product_menu(message)
  except:
    bot.send_message(message.chat.id, '❌ Xatolik! Faqat raqam kiriting.')
    send_sex_product_menu(message)


def finish_sex_income(message):
  uid = message.from_user.id
  data = user_steps.get(uid)
  if not data or not data.get('sex_cart'):
    bot.send_message(
        message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu('sex')
    )
    return

  conn = get_db_connection()
  cursor = conn.cursor()
  items_text = ''
  cart_items = []

  for p_name, qty in data['sex_cart'].items():
    cursor.execute(
        'UPDATE products SET stock = stock + ? WHERE name = ?', (qty, p_name)
    )
    items_text += f'🔹 {p_name}: +{qty}\n'
    cart_items.append({'name': p_name, 'qty': qty})
    
    # Bazaga product_incomes tarixiga yozish
    bugun_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    cursor.execute(
        'INSERT INTO product_incomes (product_name, qty, cost_price, date) VALUES (?, ?, 0, ?)',
        (p_name, qty, bugun_str)
    )

  staff_res = cursor.execute(
      'SELECT name FROM users WHERE tg_id = ?', (uid,)
  ).fetchone()
  staff_name = staff_res['name'] if staff_res else 'Nomalum'
  bugun = datetime.now().strftime('%Y-%m-%d %H:%M')

  income_id = cursor.lastrowid
  conn.commit()
  conn.close()

  excel_file = create_excel_sex_income(
      income_id, staff_name, bugun, cart_items
  )

  bot.send_message(
      message.chat.id,
      f'✅ <b>Skladga mahsulotlar muvaffaqiyatli kirim qilindi!</b>\n\n{items_text}',
      parse_mode='HTML',
      reply_markup=get_main_menu('sex'),
  )
  with open(excel_file, 'rb') as doc:
    bot.send_document(
        message.chat.id, doc, caption=f'📄 Kirim Nakladnoy (#{income_id})'
    )

  try:
    group_text = (
        f"🏭 <b>YANGI SKLAD KIRIMI (#{income_id})</b>\n"
        f"👤 <b>Mas’ul:</b> {staff_name}\n"
        f"📅 <b>Sana:</b> {bugun}\n\n"
        f"<b>Kirim qilingan mahsulotlar:</b>\n{items_text}"
    )
    bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')
    with open(excel_file, 'rb') as doc_group:
      bot.send_document(
          SEX_GROUP_ID, doc_group, caption=f'📄 Kirim Nakladnoy (#{income_id})'
      )
  except Exception as e:
    print('Sex guruhiga yuborish xatosi:', e)

  if os.path.exists(excel_file):
    os.remove(excel_file)
  if uid in user_steps:
    del user_steps[uid]


@bot.message_handler(func=lambda message: message.text == '📋 Ombordagi Qoldiqlar')
def sex_view_stock(message):
  conn = get_db_connection()
  prods = conn.execute('SELECT name, stock FROM products').fetchall()
  conn.close()
  text = '📦 <b>Ombordagi qoldiqlar:</b>\n\n'
  for p in prods:
    stk = p['stock'] if p['stock'] is not None else 0
    text += f"🔹 {p['name']}: <b>{int(stk)}</b>\n"
  bot.send_message(
      message.chat.id, text, parse_mode='HTML', reply_markup=get_main_menu('sex')
  )


@bot.message_handler(func=lambda message: message.text == '📦 Sklad & Mahsulotlar')
def admin_sklad_menu(message):
  if not is_admin_tg(message.from_user.id):
    return
  conn = get_db_connection()
  prods = conn.execute('SELECT id, name, stock FROM products').fetchall()
  conn.close()

  inline_kb = types.InlineKeyboardMarkup(row_width=1)
  if not prods:
    bot.send_message(message.chat.id, '📦 Omborxonada mahsulotlar qolmagan.')
    return
  for p in prods:
    stock_val = p['stock'] if p['stock'] is not None else 0
    inline_kb.add(
        types.InlineKeyboardButton(
            f"🔹 {p['name']} ({int(stock_val)} kg/dona)",
            callback_data=f"adm_prod_{p['id']}",
        )
    )
  inline_kb.add(
      types.InlineKeyboardButton(
          "🆕 ✨ YANGI MAHSULOT QO'SHISH", callback_data='adm_create_product'
      )
  )
  bot.send_message(
      message.chat.id,
      '<b>📦 Sklad nazorati:</b>',
      parse_mode='HTML',
      reply_markup=inline_kb,
  )


@bot.callback_query_handler(func=lambda call: call.data.startswith('adm_prod_'))
def admin_product_detail(call):
  p_id = int(call.data.split('_')[-1])
  conn = get_db_connection()
  p = conn.execute(
      'SELECT name, optom_price, chakana_price, stock, agent_percent FROM products WHERE id ='
      ' ?',
      (p_id,),
  ).fetchone()
  conn.close()
  if p:
    optom = p['optom_price'] if p['optom_price'] is not None else 0
    chakana = p['chakana_price'] if p['chakana_price'] is not None else 0
    stock = p['stock'] if p['stock'] is not None else 0
    percent = p['agent_percent'] if p['agent_percent'] is not None else 5
    text = (
        f"📦 <b>Mahsulot:</b> {p['name']}\n💰 Optom: {optom:,.0f} so'm\n🛍"
        f" Chakana: {chakana:,.0f} so'm\n🔢 <b>Qoldiq:</b> {int(stock)}\n📊 Agent foizi: <b>{percent}%</b>"
    )
    markup = types.InlineKeyboardMarkup()
    markup.row(
        types.InlineKeyboardButton(
            "➕ Qoldiq Qo'shish", callback_data=f'stk_plus_{p_id}'
        ),
        types.InlineKeyboardButton(
            '➖ Qoldiq Ayirish', callback_data=f'stk_minus_{p_id}'
        ),
    )
    markup.row(
        types.InlineKeyboardButton(
            '💵 Optom Narx', callback_data=f'prc_optom_{p_id}'
        ),
        types.InlineKeyboardButton(
            '💵 Chakana Narx', callback_data=f'prc_chakana_{p_id}'
        ),
    )
    markup.row(
        types.InlineKeyboardButton(
            '📊 Agent Foizi (%)', callback_data=f'prc_percent_{p_id}'
        ),
    )
    markup.row(
        types.InlineKeyboardButton("🗑 O'chirish", callback_data=f'stk_del_{p_id}')
    )
    bot.edit_message_text(
        text,
        call.message.chat.id,
        call.message.message_id,
        parse_mode='HTML',
        reply_markup=markup,
    )


@bot.callback_query_handler(func=lambda call: call.data.startswith(('stk_', 'prc_')))
def admin_stock_price_actions(call):
  prefix, action, p_id = call.data.split('_', 2)
  p_id = int(p_id)
  conn = get_db_connection()
  res = conn.execute('SELECT name FROM products WHERE id = ?', (p_id,)).fetchone()
  p_name = res['name'] if res else 'Nomalum'
  conn.close()

  if action == 'del':
    conn = get_db_connection()
    conn.execute('DELETE FROM products WHERE id = ?', (p_id,))
    conn.commit()
    conn.close()
    bot.send_message(
        call.message.chat.id, f"🗑 <b>{p_name}</b> o'chirildi.", parse_mode='HTML'
    )
    return

  msg = bot.send_message(
      call.message.chat.id,
      f"🔢 <b>{p_name}</b> uchun qiymat kiriting:",
      parse_mode='HTML',
  )
  bot.register_next_step_handler(msg, save_product_edits, prefix, action, p_id)


def save_product_edits(message, prefix, action, p_id):
  try:
    val = float(message.text)
    conn = get_db_connection()
    if action == 'plus':
      conn.execute(
          'UPDATE products SET stock = stock + ? WHERE id = ?', (val, p_id)
      )
    elif action == 'minus':
      conn.execute(
          'UPDATE products SET stock = stock - ? WHERE id = ?', (val, p_id)
      )
    elif action == 'optom':
      conn.execute(
          'UPDATE products SET optom_price = ? WHERE id = ?', (val, p_id)
      )
    elif action == 'chakana':
      conn.execute(
          'UPDATE products SET chakana_price = ? WHERE id = ?', (val, p_id)
      )
    elif action == 'percent':
      conn.execute(
          'UPDATE products SET agent_percent = ? WHERE id = ?', (val, p_id)
      )
    conn.commit()
    conn.close()
    bot.send_message(message.chat.id, "✅ O'zgarish saqlandi!")
  except:
    bot.send_message(message.chat.id, '❌ Faqat raqam kiriting.')


@bot.callback_query_handler(func=lambda call: call.data == 'adm_create_product')
def admin_create_product_start(call):
  msg = bot.send_message(
      call.message.chat.id, '📝 Yangi mahsulot NOMINI kiriting:'
  )
  bot.register_next_step_handler(msg, process_new_p_name)


def process_new_p_name(message):
  name = message.text
  msg = bot.send_message(
      message.chat.id, f"💰 '{name}' uchun Optom narxini kiriting:"
  )
  bot.register_next_step_handler(msg, process_new_p_optom, name)


def process_new_p_optom(message, name):
  try:
    optom = float(message.text)
    msg = bot.send_message(
        message.chat.id, f"🛍 '{name}' uchun Chakana narxini kiriting:"
    )
    bot.register_next_step_handler(msg, process_new_p_final, name, optom)
  except:
    bot.send_message(message.chat.id, 'Xato! Faqat raqam kiriting.')


def process_new_p_final(message, name, optom):
  try:
    chakana = float(message.text)
    conn = get_db_connection()
    conn.execute(
        'INSERT INTO products (name, optom_price, chakana_price, stock)'
        ' VALUES (?, ?, ?, 0)',
        (name, optom, chakana),
    )
    conn.commit()
    conn.close()
    bot.send_message(message.chat.id, f'✅ Yangi mahsulot qo\'shildi: {name}')
  except:
    bot.send_message(message.chat.id, 'Xato! Bu nom allaqachon mavjud.')


@bot.message_handler(func=lambda message: message.text == '💳 Qarzdorlik boshqaruvi')
def admin_debt_menu(message):
    if not is_admin_tg(message.from_user.id):
        bot.send_message(message.chat.id, 'Bu bo‘lim faqat admin uchun.')
        return
    conn = get_db_connection()
    shops = conn.execute('SELECT id, name, debt FROM shops ORDER BY name').fetchall()
    conn.close()
    if not shops:
        bot.send_message(message.chat.id, 'Do‘konlar ro‘yxati bo‘sh.')
        return
    markup = types.InlineKeyboardMarkup()
    text = '<b>💳 Do‘konlar qarzdorligi</b>\n\nDo‘konni tanlang:'
    for sh in shops:
        debt = safe_number(sh['debt'])
        text += f"\n• {sh['name']}: {debt:,.0f} so‘m"
        markup.add(types.InlineKeyboardButton(f"{sh['name']} ({debt:,.0f})", callback_data=f"admdebt_{sh['id']}"))
    bot.send_message(message.chat.id, text, parse_mode='HTML', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data.startswith('admdebt_'))
def admin_debt_shop_selected(call):
    if not is_admin_tg(call.from_user.id):
        bot.answer_callback_query(call.id, 'Faqat admin uchun.', show_alert=True)
        return
    try:
        shop_id = int(call.data.split('_')[-1])
    except ValueError:
        bot.answer_callback_query(call.id, 'Do‘kon ID noto‘g‘ri.')
        return
    conn = get_db_connection()
    shop = conn.execute('SELECT id, name, debt FROM shops WHERE id = ?', (shop_id,)).fetchone()
    conn.close()
    if not shop:
        bot.answer_callback_query(call.id, 'Do‘kon topilmadi.')
        return
    user_steps[call.from_user.id] = {'admin_debt_shop_id': shop_id, 'admin_debt_shop_name': shop['name']}
    bot.answer_callback_query(call.id)
    msg = bot.send_message(call.message.chat.id, f"🏪 {shop['name']}\nHozirgi qarz: {safe_number(shop['debt']):,.0f} so‘m\n\nQarzga qo‘shish uchun manfiy summa kiriting (masalan -510000), to‘lovni ayirish uchun musbat summa kiriting (masalan 100000):")
    bot.register_next_step_handler(msg, admin_debt_amount_received)


def admin_debt_amount_received(message):
    state = user_steps.get(message.from_user.id, {})
    shop_id = state.get('admin_debt_shop_id')
    if not shop_id:
        bot.send_message(message.chat.id, 'Amal muddati tugadi. Qarzdorlik bo‘limini qayta oching.')
        return
    try:
        raw = float(str(message.text).replace(' ', '').replace(',', '.'))
        if raw == 0:
            raise ValueError()
    except Exception:
        msg = bot.send_message(message.chat.id, 'Nol bo‘lmagan summa kiriting. Masalan -510000 yoki 100000.')
        bot.register_next_step_handler(msg, admin_debt_amount_received)
        return
    conn = get_db_connection()
    shop = conn.execute('SELECT name, debt FROM shops WHERE id = ?', (shop_id,)).fetchone()
    if not shop:
        conn.close()
        user_steps.pop(message.from_user.id, None)
        bot.send_message(message.chat.id, 'Do‘kon topilmadi.')
        return
    # Manfiy qiymat eski qarzni qo‘shadi; musbat qiymat qarz to‘lovini ayiradi.
    delta_debt = abs(raw) if raw < 0 else -raw
    new_debt = safe_number(shop['debt']) + delta_debt
    conn.execute('UPDATE shops SET debt = ? WHERE name = ?', (new_debt, shop['name']))
    if raw > 0:
        conn.execute('INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)', (f"Qarz to‘lovi ({shop['name']})", raw, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    log_audit(conn, 'Telegram admin qarzdorlikni o‘zgartirdi', 'shop', shop_id, f"do‘kon={shop['name']}; kiritilgan={raw}; yangi_qoldiq={new_debt}")
    conn.commit()
    conn.close()
    user_steps.pop(message.from_user.id, None)
    bot.send_message(message.chat.id, f"✅ {shop['name']} qarzdorligi yangilandi.\nYangi qarz: {new_debt:,.0f} so‘m", reply_markup=get_main_menu('admin'))


@bot.message_handler(func=lambda message: message.text == '📥 Kirim qo‘shish')
def admin_income_start(message):
    if not is_admin_tg(message.from_user.id):
        bot.send_message(message.chat.id, 'Bu bo‘lim faqat admin uchun.')
        return
    msg = bot.send_message(message.chat.id, 'Kirim manbasini yozing (masalan: naqd tushum, boshqa kirim):')
    bot.register_next_step_handler(msg, admin_income_source_received)


def admin_income_source_received(message):
    if not is_admin_tg(message.from_user.id):
        return
    source = (message.text or '').strip()
    if not source or source.startswith('/'):
        bot.send_message(message.chat.id, 'Kirim manbasini matn bilan kiriting.')
        return
    user_steps[message.from_user.id] = {'admin_income_source': source}
    msg = bot.send_message(message.chat.id, 'Kirim summasini so‘mda kiriting:')
    bot.register_next_step_handler(msg, admin_income_amount_received)


def admin_income_amount_received(message):
    if not is_admin_tg(message.from_user.id):
        return
    state = user_steps.get(message.from_user.id, {})
    source = state.get('admin_income_source')
    try:
        amount = float(str(message.text).replace(' ', '').replace(',', '.'))
        if amount <= 0:
            raise ValueError()
    except Exception:
        msg = bot.send_message(message.chat.id, 'Musbat summani kiriting, masalan 250000.')
        bot.register_next_step_handler(msg, admin_income_amount_received)
        return
    conn = get_db_connection()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute('INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)', (source, amount, now))
    log_audit(conn, 'Telegram admin kirim kiritdi', 'income', '', f'manba={source}; summa={amount}')
    conn.commit()
    conn.close()
    user_steps.pop(message.from_user.id, None)
    bot.send_message(message.chat.id, f'✅ Kirim saqlandi: {source} — {amount:,.0f} so‘m', reply_markup=get_main_menu('admin'))


# Foydalanuvchining joriy katalog holatini saqlash uchun
catalog_sessions = {}

@bot.message_handler(func=lambda message: message.text == '🖼 Mahsulotlar katalogi')
def send_catalog_categories_to_agent(message):
    conn = get_db_connection()
    role_row = conn.execute('SELECT role FROM users WHERE tg_id = ?', (message.from_user.id,)).fetchone()
    if not role_row or not any(r in (role_row['role'] or '') for r in ('agent', 'admin')):
        conn.close()
        bot.send_message(message.chat.id, 'Katalogni ko‘rish uchun agent roli kerak.')
        return
    
    # Bazadagi mavjud kategoriyalarni olish
    categories = conn.execute('SELECT DISTINCT category FROM catalog_items UNION SELECT DISTINCT category FROM products').fetchall()
    conn.close()
    
    if not categories:
        bot.send_message(message.chat.id, 'Hozircha kategoriyalar yoki mahsulotlar mavjud emas.')
        return

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for cat in categories:
        cat_name = cat['category'] if cat['category'] else 'Boshqa'
        markup.add(types.KeyboardButton(f"📁 {cat_name}"))
    
    markup.add(types.KeyboardButton("🔙 Orqaga (Asosiy menyu)"))
    
    bot.send_message(
        message.chat.id,
        "🖼 <b>Mahsulotlar katalogi</b>\n\nQaysi kategoriyadagi mahsulotlarni ko'rishni xohlaysiz? Quyidagilardan birini tanlang:",
        parse_mode='HTML',
        reply_markup=markup
    )


# Kategoriyalardan biri tanlanganda ishlaydi
@bot.message_handler(func=lambda message: message.text and message.text.startswith('📁 '))
def show_category_catalog_slide(message):
    cat_name = message.text.replace('📁 ', '').strip()
    chat_id = message.chat.id
    
    conn = get_db_connection()
    # Avval catalog_items jadvalidan, agar u bo'sh bo'lsa products jadvalidan qidiramiz
    items = conn.execute('SELECT id, name, price, description, image_filename FROM catalog_items WHERE active = 1 ORDER BY id DESC').fetchall()
    conn.close()
    
    # Agar catalog_items'da kategoriya ustuni bo'lmasa, uni nom bo'yicha yoki barchasini chiqaramiz
    # Keling, foydalanuvchiga o'sha kategoriya bo'yicha filtrlangan ro'yxatni tuzamiz:
    filtered_items = []
    for item in items:
        filtered_items.append(item)
        
    if not filtered_items:
        bot.send_message(chat_id, f"❌ '{cat_name}' kategoriyasida hozircha mahsulotlar yo'q.")
        return

    # Sessiyaga saqlaymiz
    catalog_sessions[chat_id] = {
        'category': cat_name,
        'index': 0,
        'items': filtered_items
    }
    
    send_filtered_catalog_slide(chat_id, 0)


def send_filtered_catalog_slide(chat_id, index):
    session_data = catalog_sessions.get(chat_id)
    if not session_data or not session_data['items']:
        return
        
    items_list = session_data['items']
    index = index % len(items_list)
    session_data['index'] = index
    
    item = items_list[index]
    image_url = item['image_filename']
    
    caption = (
        f"📂 <b>Kategoriya:</b> {session_data['category']}\n"
        f"📦 <b>Mahsulot:</b> {item['name']}\n"
        f"💰 <b>Narxi:</b> {safe_number(item['price']):,.0f} so‘m\n"
    )
    if item['description']:
        caption += f"📝 {item['description']}\n"
    
    caption += f"\n📄 Sahifa: {index + 1} / {len(items_list)}"

    # Slayd tugmalari
    markup = types.InlineKeyboardMarkup(row_width=3)
    prev_btn = types.InlineKeyboardButton("⬅️ Orqaga", callback_data="fcat_prev")
    counter_btn = types.InlineKeyboardButton(f"{index + 1}/{len(items_list)}", callback_data="fcat_noop")
    next_btn = types.InlineKeyboardButton("Keyingi ➡️", callback_data="fcat_next")
    
    if len(items_list) > 1:
        markup.add(prev_btn, counter_btn, next_btn)

    try:
        if image_url and (image_url.startswith('http://') or image_url.startswith('https://')):
            bot.send_photo(chat_id=chat_id, photo=image_url, caption=caption, parse_mode='HTML', reply_markup=markup)
        else:
            image_path = os.path.join(CATALOG_UPLOAD_FOLDER, os.path.basename(image_url or ''))
            if image_url and os.path.isfile(image_path):
                with open(image_path, 'rb') as photo:
                    bot.send_photo(chat_id=chat_id, photo=photo, caption=caption, parse_mode='HTML', reply_markup=markup)
            else:
                bot.send_message(chat_id, caption, parse_mode='HTML', reply_markup=markup)
    except Exception as exc:
        app.logger.exception('Katalog slayd xatosi: %s', exc)
        bot.send_message(chat_id, caption, parse_mode='HTML', reply_markup=markup)


# Kategoriya ichidagi slayd tugmalari (Oldinga / Orqaga)
@bot.callback_query_handler(func=lambda call: call.data.startswith('fcat_'))
def filtered_catalog_callback(call):
    chat_id = call.message.chat.id
    session_data = catalog_sessions.get(chat_id)
    
    if not session_data:
        bot.answer_callback_query(call.id, "Katalog sessiyasi tugagan, qaytadan tanlang.")
        return
        
    current_index = session_data['index']
    items_list = session_data['items']

    if call.data == 'fcat_next':
        current_index += 1
    elif call.data == 'fcat_prev':
        current_index -= 1
    elif call.data == 'fcat_noop':
        bot.answer_callback_query(call.id)
        return

    session_data['index'] = current_index % len(items_list)
    
    try:
        bot.delete_message(chat_id, call.message.message_id)
    except:
        pass
        
    send_filtered_catalog_slide(chat_id, session_data['index'])
    bot.answer_callback_query(call.id)


@bot.message_handler(
    func=lambda message: message.text
    in [
        '📊 Kunlik Hisobot',
        "🏪 AKB (Do'konlar & Qarz)",
        '👥 Agentlar boshqaruvi',
    ]
)
def admin_other_sections(message):
  if message.from_user.id != ADMIN_ID:
    return
  conn = get_db_connection()

  if message.text == '📊 Kunlik Hisobot':
    bugun = datetime.now().strftime('%Y-%m-%d')
    orders = conn.execute(
        'SELECT id, shop_name, total_sum, status FROM orders WHERE date LIKE ?',
        (f'{bugun}%',),
    ).fetchall()
    text = f'📊 <b>Bugungi buyurtmalar ({bugun}):</b>\n\n'
    inline_kb = types.InlineKeyboardMarkup()
    if not orders:
      text += "Hali buyurtma yo'q."
    for o in orders:
      t_sum = o['total_sum'] if o['total_sum'] is not None else 0
      text += (
          f"🆔 #{o['id']} | {o['shop_name']} | {t_sum:,.0f} so'm |"
          f" <b>{o['status']}</b>\n"
      )
      inline_kb.add(
          types.InlineKeyboardButton(
              f"⚙️ #{o['id']} Statusi", callback_data=f"mng_ord_{o['id']}"
          )
      )
    bot.send_message(
        message.chat.id, text, parse_mode='HTML', reply_markup=inline_kb
    )

  elif message.text == "🏪 AKB (Do'konlar & Qarz)":
    shops = conn.execute(
        'SELECT name, phone, debt, region FROM shops'
    ).fetchall()
    text = "<b>🏪 Do'konlar qarzlari:</b>\n\n"
    for s in shops:
      debt_val = s['debt'] if s['debt'] is not None else 0
      text += (
          f"🏢 {s['name']} ({s['region']}) — Qarz: {debt_val:,.0f} so'm\n"
      )
    bot.send_message(message.chat.id, text, parse_mode='HTML')

  elif message.text == '👥 Agentlar boshqaruvi':
    agents = conn.execute(
        "SELECT name, phone, role, tg_id FROM users WHERE role != 'admin'"
    ).fetchall()
    inline_kb = types.InlineKeyboardMarkup()
    text = '<b>👥 Agentlar va Sex xodimlari:</b>\n\n'
    for a in agents:
      status = (
          '✅ Faol'
          if any(r in a['role'] for r in ['agent', 'sex', 'zavsklad', 'ekspeditor'])
          else '⏳ Kutilmoqda'
      )
      text += f"👤 {a['name']} ({a['phone']}) - [{a['role']}] - {status}\n"
      if a['role'] == 'pending':
        inline_kb.add(
            types.InlineKeyboardButton(
                f"👍 {a['name']}ni tasdiqlash",
                callback_data=f"approve_{a['tg_id']}",
            )
        )
    bot.send_message(
        message.chat.id, text, parse_mode='HTML', reply_markup=inline_kb
    )
  conn.close()


@bot.callback_query_handler(func=lambda call: call.data.startswith('mng_ord_'))
def manage_order_status_menu(call):
  order_id = int(call.data.split('_')[-1])
  markup = types.InlineKeyboardMarkup()
  markup.row(
      types.InlineKeyboardButton(
          '🔄 Yangi', callback_data=f'st_yangi_{order_id}'
      ),
      types.InlineKeyboardButton(
          '🚚 Otgruzka', callback_data=f'st_otgruzka_{order_id}'
      ),
  )
  markup.row(
      types.InlineKeyboardButton(
          '✅ Yetkazildi', callback_data=f'st_done_{order_id}'
      ),
      types.InlineKeyboardButton(
          '❌ Bekor', callback_data=f'st_otmen_{order_id}'
      ),
  )
  bot.send_message(
      call.message.chat.id,
      f'🆔 #{order_id} - Statusni tanlang:',
      reply_markup=markup,
  )


@bot.callback_query_handler(func=lambda call: call.data.startswith('st_'))
def change_status_logic(call):
  _, mode, order_id = call.data.split('_')
  order_id = int(order_id)
  status_map = {
      'yangi': 'Yangi',
      'otgruzka': 'Otgruzka',
      'done': 'Yetkazildi',
      'otmen': 'Bekor',
  }
  new_status = status_map[mode]
  conn = get_db_connection()
  
  ord_data = conn.execute("SELECT shop_name, total_sum, status FROM orders WHERE id = ?", (order_id,)).fetchone()
  if ord_data:
      old_delivered = ord_data['status'] == 'Yetkazildi'
      new_delivered = new_status == 'Yetkazildi'
      if old_delivered != new_delivered:
          shop_name = ord_data['shop_name']
          shop_res = conn.execute("SELECT debt FROM shops WHERE name = ?", (shop_name,)).fetchone()
          if shop_res:
              curr_debt = safe_number(shop_res['debt'])
              delta = safe_number(ord_data['total_sum']) * (1 if new_delivered else -1)
              conn.execute("UPDATE shops SET debt = ? WHERE name = ?", (max(0, curr_debt + delta), shop_name))
      conn.execute('UPDATE orders SET status = ? WHERE id = ?', (new_status, order_id))
      conn.execute('INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, ?, ?)',
          (order_id, new_status, datetime.now().strftime('%Y-%m-%d %H:%M')))
  conn.commit()
  conn.close()
  bot.send_message(
      call.message.chat.id, f'✅ Status: <b>{new_status}</b>', parse_mode='HTML'
  )


@bot.callback_query_handler(func=lambda call: call.data.startswith('approve_'))
def approve_agent_cb(call):
  agent_id = int(call.data.split('_')[-1])
  conn = get_db_connection()
  conn.execute(
      "UPDATE users SET role = 'agent' WHERE tg_id = ?", (agent_id,)
  )
  conn.commit()
  conn.close()
  bot.answer_callback_query(call.id, 'Foydalanuvchi tasdiqlandi!')
  try:
    bot.send_message(
        agent_id,
        "🎉 Sizning so'rovingiz tasdiqlandi! /start bosing",
        reply_markup=get_main_menu('agent'),
    )
  except:
    pass


@app.route('/add_shop', methods=['POST'])
def add_shop():
  name, phone, visit_days, region, landmark, inventory, territory = (
      request.form['name'],
      request.form['phone'],
      request.form.get('visit_days', ''),
      request.form.get('region', ''),
      request.form.get('landmark', ''),
      request.form.get('inventory', ''),
      request.form.get('territory', ''),
  )
  conn = get_db_connection()
  try:
    conn.execute(
        'INSERT INTO shops (name, phone, debt, visit_days, region, landmark, inventory, territory) VALUES (?, ?, 0, ?, ?, ?, ?, ?)',
        (name, phone, visit_days, region, landmark, inventory, territory),
    )
    conn.commit()
  except:
    pass
  conn.close()
  return redirect(url_for('index'))

@app.route('/update_order/<int:order_id>', methods=['POST'])
def update_order(order_id):
    shop_name = (request.form.get('shop_name') or '').strip()
    comment = request.form.get('comment', '').strip()
    price_type = request.form.get('price_type', 'optom')
    if price_type not in ('optom', 'chakana'):
        price_type = 'optom'
    try:
        discount = max(0.0, float((request.form.get('discount', '0') or '0').replace(',', '.')))
    except (TypeError, ValueError):
        return 'Chegirma summasi noto‘g‘ri kiritilgan.', 400
    if not shop_name:
        return 'Do‘kon nomini tanlang.', 400

    product_names = request.form.getlist('product_name')
    qtys = request.form.getlist('qty')
    conn = get_db_connection()
    try:
        old_order = conn.execute('SELECT * FROM orders WHERE id = ?', (order_id,)).fetchone()
        if not old_order:
            return 'Buyurtma topilmadi.', 404
        if old_order['status'] == 'Yetkazildi' and old_order['shop_name'] != shop_name:
            return 'Yetkazilgan buyurtmaning do‘konini almashtirishdan oldin holatini qaytaring.', 400

        conn.execute('BEGIN IMMEDIATE')
        if old_order['items_text']:
            for line in old_order['items_text'].splitlines():
                if not line.strip():
                    continue
                parts = line.split(' - ', 1)
                if len(parts) != 2:
                    continue
                old_name = parts[0].strip()
                try:
                    old_qty = float(parts[1].split('x', 1)[0].strip().replace(',', '.'))
                except (ValueError, TypeError):
                    continue
                if old_qty > 0:
                    conn.execute('UPDATE products SET stock = stock + ? WHERE name = ?', (old_qty, old_name))
                    if conn.execute('SELECT changes()').fetchone()[0]:
                        record_stock_movement(conn, old_name, old_qty, 'Buyurtma tahriri — eski qoldiq qaytarildi', str(order_id))

        total_sum = 0.0
        cost_total = 0.0
        items = []
        for i, p_name in enumerate(product_names):
            p_name = (p_name or '').strip()
            if not p_name:
                continue
            if i >= len(qtys):
                raise ValueError(f'{p_name}: miqdor kiritilmagan.')
            try:
                qty = float((qtys[i] or '0').replace(',', '.'))
            except (ValueError, TypeError):
                raise ValueError(f'{p_name}: miqdor noto‘g‘ri.')
            if qty <= 0:
                raise ValueError(f'{p_name}: miqdor noldan katta bo‘lishi kerak.')
            prod = conn.execute('SELECT id, name, optom_price, chakana_price, cost_price, stock FROM products WHERE name = ?', (p_name,)).fetchone()
            if not prod:
                raise ValueError(f'{p_name}: mahsulot topilmadi.')
            stock = safe_number(prod['stock'])
            if stock < qty:
                raise ValueError(f'Omborda {p_name} yetarli emas. Qoldiq: {stock:g}')
            price = safe_number(prod['optom_price']) if price_type == 'optom' else safe_number(prod['chakana_price'] or prod['optom_price'])
            subtotal = price * qty
            total_sum += subtotal
            cost_total += safe_number(prod['cost_price']) * qty
            items.append(f"{p_name} - {qty:g}x = {subtotal:,.0f} so'm")
            conn.execute('UPDATE products SET stock = stock - ? WHERE id = ?', (qty, prod['id']))
            record_stock_movement(conn, p_name, -qty, 'Buyurtma tahriri', str(order_id))

        discount = min(discount, total_sum)
        final_sum = max(0.0, total_sum - discount)
        if old_order['status'] == 'Yetkazildi':
            delta = final_sum - safe_number(old_order['total_sum'])
            if abs(delta) > 0.000001:
                shop = conn.execute('SELECT debt FROM shops WHERE name = ?', (shop_name,)).fetchone()
                if shop:
                    conn.execute('UPDATE shops SET debt = ? WHERE name = ?',
                                 (max(0.0, safe_number(shop['debt']) + delta), shop_name))

        conn.execute("""UPDATE orders SET shop_name = ?, total_sum = ?, items_text = ?, price_type = ?,
                        discount = ?, comment = ?, cost_total = ? WHERE id = ?""",
                     (shop_name, final_sum, '\n'.join(items), price_type, discount, comment, cost_total, order_id))
        log_audit(conn, 'Buyurtma tahrirlandi', 'order', order_id,
                  f'do‘kon={shop_name}; jami={final_sum:.2f}; mahsulotlar={len(items)}')
        conn.commit()
    except ValueError as exc:
        conn.rollback()
        return str(exc), 400
    except Exception as exc:
        conn.rollback()
        print('Buyurtmani tahrirlash xatosi:', exc)
        return 'Buyurtmani saqlashda xato yuz berdi. O‘zgarishlar saqlanmadi.', 500
    finally:
        conn.close()
    return redirect(url_for('index'))

@app.route('/update_shop/<int:shop_id>', methods=['POST'])
def update_shop(shop_id):
  name = request.form.get('name', '').strip()
  phone = request.form.get('phone', '').strip()
  region = request.form.get('region', '')
  landmark = request.form.get('landmark', '')
  visit_days = request.form.get('visit_days', '')
  inventory = request.form.get('inventory', '')
  territory = request.form.get('territory', '')

  if not name:
    return 'Do‘kon nomini kiriting.', 400
  conn = get_db_connection()
  try:
    conn.execute(
        'UPDATE shops SET name = ?, phone = ?, region = ?, landmark = ?, visit_days = ?, inventory = ?, territory = ? WHERE id = ?',
        (name, phone, region, landmark, visit_days, inventory, territory, shop_id),
    )
    log_audit(conn, 'Do‘kon tahrirlandi', 'shop', shop_id, f'nomi={name}; telefon={phone}')
    conn.commit()
  except sqlite3.IntegrityError:
    conn.rollback()
    return 'Bunday nomdagi do‘kon allaqachon mavjud.', 400
  except Exception as e:
    conn.rollback()
    print('Do‘konni tahrirlash xatosi:', e)
    return 'Do‘konni saqlashda xato yuz berdi.', 500
  finally:
    conn.close()
  return redirect(url_for('index'))


@app.route('/update_product/<int:prod_id>', methods=['POST'])
def update_product(prod_id):
  name = request.form.get('name', '').strip()
  category = request.form.get('category', 'Boshqa').strip() or 'Boshqa'
  stock = safe_number(request.form.get('stock', 0), minimum=0)
  cost_price = safe_number(request.form.get('cost_price', 0), minimum=0)
  optom_price = safe_number(request.form.get('optom_price', 0), minimum=0)
  chakana_price = safe_number(request.form.get('chakana_price', optom_price), minimum=0)
  agent_percent = safe_number(request.form.get('agent_percent', 5), minimum=0, maximum=100)

  if not name:
    return 'Mahsulot nomini kiriting.', 400
  conn = get_db_connection()
  try:
    conn.execute(
        'UPDATE products SET name = ?, category = ?, stock = ?, cost_price ='
        ' ?, optom_price = ?, chakana_price = ?, agent_percent = ? WHERE id = ?',
        (
            name,
            category,
            stock,
            cost_price,
            optom_price,
            chakana_price,
            agent_percent,
            prod_id,
        ),
    )
    log_audit(conn, 'Mahsulot tahrirlandi', 'product', prod_id, f'nomi={name}; qoldiq={stock}; tannarx={cost_price}; optom={optom_price}; chakana={chakana_price}')
    conn.commit()
  except Exception as e:
    conn.rollback()
    print('Mahsulotni tahrirlash xatosi:', e)
    conn.close()
    return 'Mahsulotni saqlashda xato yuz berdi. O‘zgarishlar saqlanmadi.', 500
  conn.close()
  return redirect(url_for('index'))


@app.route('/adjust_stock_web', methods=['POST'])
def adjust_stock_web():
  prod_id = request.form.get('prod_id')
  action = request.form.get('action')
  qty = safe_number(request.form.get('qty', 0), minimum=0)
  if qty <= 0:
    return 'Miqdor noldan katta bo‘lishi kerak.', 400

  conn = get_db_connection()
  product = conn.execute('SELECT name, stock FROM products WHERE id = ?', (prod_id,)).fetchone()
  if not product:
    conn.close()
    return 'Mahsulot topilmadi.', 404
  if action == 'minus' and safe_number(product['stock']) < qty:
    conn.close()
    return 'Omborda yetarli mahsulot yo‘q.', 400
  if action == 'plus':
    conn.execute(
        'UPDATE products SET stock = stock + ? WHERE id = ?', (qty, prod_id)
    )
  elif action == 'minus':
    conn.execute(
        'UPDATE products SET stock = stock - ? WHERE id = ?', (qty, prod_id)
    )
  else:
    conn.close()
    return 'Noto‘g‘ri amal.', 400
  record_stock_movement(conn, product['name'], qty if action == 'plus' else -qty, 'Qo‘lda tuzatish', 'Web panel')
  log_audit(conn, 'Ombor qoldig‘i o‘zgartirildi', 'product', prod_id, f'{action}:{qty}')
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/add_agent_web', methods=['POST'])
def add_agent_web():
  name = request.form.get('name')
  phone = request.form.get('phone')
  tg_id = request.form.get('tg_id')
  role = request.form.get('role', 'agent')
  territory = request.form.get('territory', '')
  salary_percent = safe_number(request.form.get('salary_percent', '5'), default=5.0, minimum=0, maximum=100)
  if tg_id:
    try:
      tg_id = int(tg_id)
    except:
      tg_id = None

  if name and tg_id:
    conn = get_db_connection()
    try:
      conn.execute(
          'INSERT OR REPLACE INTO users (tg_id, name, phone, role, territory, salary_percent) VALUES (?, ?, ?, ?, ?, ?)',
          (tg_id, name, phone or '', role, territory, salary_percent),
      )
      conn.commit()
    except Exception as e:
      print('Agent qo\'shish xatosi:', e)
    conn.close()
  return redirect(url_for('index'))


@app.route('/import_shops_excel', methods=['POST'])
def import_shops_excel():
  if 'excel_file' not in request.files:
    return redirect(url_for('index'))
  file = request.files['excel_file']
  if file.filename == '':
    return redirect(url_for('index'))

  try:
    df = pd.read_excel(file)
    conn = get_db_connection()
    cursor = conn.cursor()

    for _, row in df.iterrows():
      name = str(
          row.get('name', row.get("Do'kon nomi", row.get('Magazin', '')))
      ).strip()
      if not name or name == 'nan':
        continue
      phone = str(row.get('phone', row.get('Telefon', row.get('Tel', '')))).strip()
      region = str(row.get('region', row.get('Hudud', row.get('Rayon', '')))).strip()
      landmark = str(
          row.get('landmark', row.get('Orienter', row.get("Mo'ljal", '')))
      ).strip()
      visit_days = str(
          row.get(
              'visit_days', row.get('Tashrif kunlari', row.get('Kunlar', ''))
          )
      ).strip()
      inventory = str(
          row.get(
              'inventory', row.get('Inventar', row.get('Jihoz', ''))
          )
      ).strip()

      debt_val = 0
      for d_key in ['debt', 'Qarz', 'Balans', 'Borg']:
        if d_key in row and pd.notna(row[d_key]):
          try:
            debt_val = float(row[d_key])
            break
          except:
            pass

      cursor.execute(
          '''INSERT INTO shops (name, phone, debt, visit_days, region, landmark, inventory) 
                       VALUES (?, ?, ?, ?, ?, ?, ?)
                       ON CONFLICT(name) DO UPDATE SET 
                       phone=excluded.phone, region=excluded.region, landmark=excluded.landmark, 
                       visit_days=excluded.visit_days, inventory=excluded.inventory''',
          (
              name,
              phone if phone != 'nan' else '',
              debt_val,
              visit_days if visit_days != 'nan' else '',
              region if region != 'nan' else '',
              landmark if landmark != 'nan' else '',
              inventory if inventory != 'nan' else '',
          ),
      )

    conn.commit()
    conn.close()
  except Exception as e:
    print('Excel import xatosi:', e)

  return redirect(url_for('index'))


@app.route('/import_products_excel', methods=['POST'])
def import_products_excel():
  if 'excel_file' not in request.files:
    return redirect(url_for('index'))
  file = request.files['excel_file']
  if file.filename == '':
    return redirect(url_for('index'))

  try:
    df = pd.read_excel(file)
    conn = get_db_connection()
    cursor = conn.cursor()

    for _, row in df.iterrows():
      name = str(
          row.get('name', row.get('Mahsulot', row.get('Tovar nomi', '')))
      ).strip()
      if not name or name == 'nan':
        continue
      category = str(
          row.get('category', row.get('Kategoriya', 'Boshqa'))
      ).strip()
      if not category or category == 'nan':
        category = 'Boshqa'

      stock_val = 0
      for s_key in ['stock', 'Qoldiq', 'Soni', 'Miqdor']:
        if s_key in row and pd.notna(row[s_key]):
          try:
            stock_val = float(row[s_key])
            break
          except:
            pass

      cost_val = 0
      for c_key in ['cost_price', 'Tannarx', 'Tan narx']:
        if c_key in row and pd.notna(row[c_key]):
          try:
            cost_val = float(row[c_key])
            break
          except:
            pass

      optom_val = 0
      for o_key in ['optom_price', 'Optom', 'Optom narx']:
        if o_key in row and pd.notna(row[o_key]):
          try:
            optom_val = float(row[o_key])
            break
          except:
            pass

      chakana_val = optom_val
      for ch_key in ['chakana_price', 'Chakana', 'Chakana narx']:
        if ch_key in row and pd.notna(row[ch_key]):
          try:
            chakana_val = float(row[ch_key])
            break
          except:
            pass

      cursor.execute(
          '''INSERT INTO products (name, category, stock, cost_price, optom_price, chakana_price) 
                       VALUES (?, ?, ?, ?, ?, ?)
                       ON CONFLICT(name) DO UPDATE SET 
                       category=excluded.category, stock=excluded.stock, 
                       cost_price=excluded.cost_price, optom_price=excluded.optom_price, 
                       chakana_price=excluded.chakana_price''',
          (name, category, stock_val, cost_val, optom_val, chakana_val),
      )

    conn.commit()
    conn.close()
  except Exception as e:
    print('Sklad Excel import xatosi:', e)

  return redirect(url_for('index'))


@app.route('/approve_agent/<int:tg_id>', methods=['POST'])
def web_approve_agent(tg_id):
  conn = get_db_connection()
  conn.execute("UPDATE users SET role = 'agent' WHERE tg_id = ?", (tg_id,))
  conn.commit()
  conn.close()
  try:
    bot.send_message(
        tg_id,
        "🎉 Sizning so'rovingiz tasdiqlandi! /start bosing",
        reply_markup=get_main_menu('agent'),
    )
  except:
    pass
  return redirect(url_for('index'))


@app.route('/delete_agent/<int:tg_id>', methods=['POST'])
def web_delete_agent(tg_id):
  conn = get_db_connection()
  conn.execute('DELETE FROM users WHERE tg_id = ?', (tg_id,))
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/update_agent/<int:tg_id>', methods=['POST'])
def update_agent(tg_id):
  name = request.form.get('name', '').strip()
  phone = request.form.get('phone', '').strip()
  role = request.form.get('role', 'agent').strip()
  territory = request.form.get('territory', '')
  try:
    salary_percent = float((request.form.get('salary_percent', '5') or '5').replace(',', '.'))
    salary_percent = min(100.0, max(0.0, salary_percent))
  except (TypeError, ValueError):
    salary_percent = 5.0

  if not name:
    return 'Agent nomini kiriting.', 400
  conn = get_db_connection()
  try:
    conn.execute(
        'UPDATE users SET name = ?, phone = ?, role = ?, territory = ?, salary_percent = ? WHERE tg_id = ?',
        (name, phone, role, territory, salary_percent, tg_id),
    )
    log_audit(conn, 'Agent tahrirlandi', 'agent', tg_id, f'nomi={name}; rol={role}; foiz={salary_percent}')
    conn.commit()
  except Exception as e:
    conn.rollback()
    print('Agentni tahrirlash xatosi:', e)
  finally:
    conn.close()
  return redirect(url_for('index'))


@bot.message_handler(func=lambda message: message.text == '🌐 Web kabinetim')
def send_staff_web_cabinet(message):
    tg_id = message.from_user.id
    conn = get_db_connection()
    user = conn.execute('SELECT role, name FROM users WHERE tg_id = ?', (tg_id,)).fetchone()
    conn.close()
    if not user or not user['role'] or user['role'] == 'pending':
        bot.send_message(message.chat.id, 'Kabinet ochilmadi: Telegram ID hali tasdiqlanmagan.')
        return
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton('🌐 Shaxsiy web kabinetni ochish', web_app=types.WebAppInfo(url=make_staff_web_url(tg_id))))
    bot.send_message(message.chat.id, f"Salom, {user['name']}! Sizning rolingiz: {user['role']}. Havola 15 daqiqa amal qiladi.", reply_markup=markup)


STAFF_WEB_TEMPLATE = """
<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Xoji Aka — Xodim kabineti</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"><style>body{background:#f4f6fa}.wrap{max-width:850px;margin:20px auto;padding:12px}.card{border:0;border-radius:15px;box-shadow:0 3px 14px #0000000b;margin-bottom:14px}.muted{color:#64748b}</style></head><body><main class="wrap">
<div class="card p-3 text-white" style="background:#173b69"><h4 class="mb-1">Xoji Aka — {{ role_label }}</h4><div>{{ name }} · Telegram ID: {{ tg_id }}</div><small>Hudud: {{ territory or 'Barcha hududlar' }}</small></div>
{% if error %}<div class="alert alert-warning">{{ error }}</div>{% endif %}
{% if role_has_agent %}<div class="card p-3"><h5>📦 Mening buyurtmalarim</h5><div class="table-responsive"><table class="table table-sm"><thead><tr><th>ID</th><th>Do‘kon</th><th>Summa</th><th>Holat</th><th>Sana</th></tr></thead><tbody>{% for o in orders %}<tr><td>#{{o['id']}}</td><td>{{o['shop_name']}}</td><td>{{'{:,.0f}'.format(o['total_sum'] or 0)}} so‘m</td><td>{{o['status']}}</td><td>{{o['date']}}</td></tr>{% else %}<tr><td colspan="5">Buyurtmalar yo‘q</td></tr>{% endfor %}</tbody></table></div></div><div class="card p-3"><h5>🏪 Hududimdagi do‘konlar va qarzdorlik</h5><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Do‘kon</th><th>Hudud</th><th>Qarz</th></tr></thead><tbody>{% for sh in shops %}<tr><td>{{sh['name']}}</td><td>{{sh['territory'] or sh['region'] or '—'}}</td><td>{{'{:,.0f}'.format(sh['debt'] or 0)}} so‘m</td></tr>{% else %}<tr><td colspan="3">Do‘konlar yo‘q</td></tr>{% endfor %}</tbody></table></div></div>{% endif %}
{% if role_key == 'zavsklad' or role_key == 'sex' %}<div class="card p-3"><h5>📋 Ombor qoldiqlari</h5><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Mahsulot</th><th>Qoldiq</th><th>Kategoriya</th></tr></thead><tbody>{% for p in products %}<tr><td>{{p['name']}}</td><td>{{p['stock']}}</td><td>{{p['category']}}</td></tr>{% endfor %}</tbody></table></div></div>{% endif %}
{% if role_key == 'zavsklad' %}<div class="card p-3"><h5>🧾 Yig‘ish kutilayotgan buyurtmalar</h5>{% for o in orders %}<div class="border rounded p-2 mb-2"><b>#{{o['id']}} — {{o['shop_name']}}</b><div>{{o['items_text']}}</div><div>Jami: {{'{:,.0f}'.format(o['total_sum'] or 0)}} so‘m · {{o['status']}}</div>{% if o['status'] == 'Yangi' %}<form method="post" class="mt-2"><input type="hidden" name="action" value="status"><input type="hidden" name="order_id" value="{{o['id']}}"><input type="hidden" name="new_status" value="Otgruzka"><button class="btn btn-primary btn-sm">Yig‘ildi — ekspeditorga berish</button></form>{% endif %}</div>{% else %}<p>Yangi buyurtmalar yo‘q.</p>{% endfor %}</div>{% endif %}
{% if role_key == 'ekspeditor' %}<div class="card p-3"><h5>🚚 Yetkazilishi kerak buyurtmalar</h5>{% for o in orders %}<div class="border rounded p-2 mb-2"><b>#{{o['id']}} — {{o['shop_name']}}</b><div>{{o['items_text']}}</div><div>Jami: {{'{:,.0f}'.format(o['total_sum'] or 0)}} so‘m · {{o['status']}}</div>{% if o['status'] in ['Yangi','Otgruzka'] %}<form method="post" class="mt-2"><input type="hidden" name="action" value="status"><input type="hidden" name="order_id" value="{{o['id']}}"><input type="hidden" name="new_status" value="Yetkazildi"><button class="btn btn-success btn-sm">Yetkazildi deb tasdiqlash</button></form>{% endif %}</div>{% else %}<p>Yetkaziladigan buyurtmalar yo‘q.</p>{% endfor %}</div><div class="card p-3"><h5>💰 To‘lov so‘rovlari</h5><p class="muted">To‘lovni rasmiylashtirish va tasdiqlashning mavjud botdagi tartibi saqlanadi.</p>{% for r in payments %}<div class="border-bottom py-2">{{r['shop_name']}} — {{'{:,.0f}'.format(r['amount'] or 0)}} so‘m · {{r['status']}} <small class="muted">{{r['date']}}</small></div>{% else %}<p>To‘lov so‘rovlari yo‘q.</p>{% endfor %}</div>{% endif %}
<div class="text-center muted small mb-4">Xodim kabineti · Ma’lumotlar Telegram ID bo‘yicha filtrlanadi</div></main></body></html>
"""

@app.route('/staff_app/<int:tg_id>', methods=['GET', 'POST'])
@app.route('/staff_app/<int:tg_id>/', methods=['GET', 'POST'])
def staff_webapp(tg_id):
    if not verify_staff_link(tg_id):
        return "Havola yaroqsiz yoki muddati tugagan. Telegram botga qaytib, '🌐 Web kabinetim' tugmasini qayta bosing.", 403
    conn = get_db_connection()
    user = conn.execute('SELECT * FROM users WHERE tg_id = ?', (tg_id,)).fetchone()
    if not user or not user['role'] or user['role'] == 'pending':
        conn.close()
        return 'Telegram foydalanuvchisi tasdiqlanmagan.', 403
    role = user['role'] or ''
    role_key = 'admin' if 'admin' in role else ('agent' if 'agent' in role else ('sex' if 'sex' in role else ('zavsklad' if 'zavsklad' in role else ('ekspeditor' if 'ekspeditor' in role else role))))
    error = ''
    if request.method == 'POST':
        action = request.form.get('action')
        order_id = safe_number(request.form.get('order_id'), 0, minimum=1)
        new_status = request.form.get('new_status', '')
        allowed = (role_key == 'zavsklad' and new_status == 'Otgruzka') or (role_key == 'ekspeditor' and new_status == 'Yetkazildi')
        order = conn.execute('SELECT * FROM orders WHERE id = ?', (int(order_id),)).fetchone() if order_id else None
        if action == 'status' and allowed and order:
            if role_key == 'zavsklad' and order['status'] != 'Yangi':
                error = 'Faqat Yangi holatdagi buyurtmani ekspeditorga berish mumkin.'
            elif role_key == 'ekspeditor' and order['status'] not in ('Yangi', 'Otgruzka'):
                error = 'Bu buyurtma yetkazish uchun mavjud emas.'
            else:
                old_status = order['status']
                conn.execute('UPDATE orders SET status = ? WHERE id = ?', (new_status, int(order_id)))
                conn.execute('INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, ?, ?)', (int(order_id), new_status, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                log_audit(conn, 'Buyurtma holati web kabinetdan o‘zgartirildi', 'order', int(order_id), f'{old_status} -> {new_status}; xodim={user["name"]}; tg_id={tg_id}')
                if new_status == 'Yetkazildi':
                    shop = conn.execute('SELECT debt FROM shops WHERE name = ?', (order['shop_name'],)).fetchone()
                    if shop:
                        conn.execute('UPDATE shops SET debt = COALESCE(debt,0) + ? WHERE name = ?', (safe_number(order['total_sum']), order['shop_name']))
                conn.commit()
        else:
            error = 'Bu amal sizning rolingizga ruxsat etilmagan yoki buyurtma topilmadi.'
    # Xodimlar faqat vazifasiga kerakli doiradagi buyurtmalarni ko‘radi.
    if role_key == 'agent':
        orders = conn.execute('SELECT * FROM orders WHERE agent_name = ? ORDER BY id DESC LIMIT 50', (user['name'],)).fetchall()
        territory = user['territory'] or ''
        if territory:
            shops = conn.execute("SELECT * FROM shops WHERE territory = ? OR territory = '' ORDER BY name", (territory,)).fetchall()
        else:
            shops = conn.execute('SELECT * FROM shops ORDER BY name').fetchall()
    elif role_key == 'zavsklad':
        orders = conn.execute("SELECT * FROM orders WHERE status = 'Yangi' ORDER BY id DESC LIMIT 100").fetchall()
        shops = []
        territory = user['territory'] or ''
    elif role_key == 'ekspeditor':
        orders = conn.execute("SELECT * FROM orders WHERE status IN ('Yangi','Otgruzka') ORDER BY id DESC LIMIT 100").fetchall()
        shops = []
        territory = user['territory'] or ''
    elif role_key == 'sex':
        orders = []
        shops = []
        territory = user['territory'] or ''
    else:
        orders = conn.execute('SELECT * FROM orders ORDER BY id DESC LIMIT 50').fetchall()
        shops = conn.execute('SELECT * FROM shops ORDER BY name').fetchall()
        territory = user['territory'] or ''
    products = conn.execute('SELECT * FROM products ORDER BY name').fetchall() if role_key in ('zavsklad','sex') else []
    payments = conn.execute('SELECT * FROM payment_requests WHERE ekspeditor_name = ? ORDER BY id DESC LIMIT 30', (user['name'],)).fetchall() if role_key == 'ekspeditor' else []
    conn.close()
    role_labels = {'agent':'Agent kabineti','zavsklad':'Zavsklad kabineti','ekspeditor':'Ekspeditor kabineti','sex':'Sex xodimi kabineti','admin':'Admin kabineti'}
    return render_template_string(STAFF_WEB_TEMPLATE, name=user['name'], tg_id=tg_id, role_key=role_key, role_label=role_labels.get(role_key, 'Xodim kabineti'), territory=territory, orders=orders, shops=shops, products=products, payments=payments, error=error, role_has_agent=(role_key == 'agent'))


AGENT_WEBAPP_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Agent Kabineti — {{ agent_name }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        body { background: #f8fafc; font-family: 'Inter', sans-serif; padding-bottom: 40px; }
        .mobile-container { max-width: 600px; margin: auto; padding: 15px; }
        .card-mobile { background: #fff; border-radius: 14px; padding: 16px; margin-bottom: 15px; box-shadow: 0 2px 6px rgba(0,0,0,0.04); border: 1px solid #e2e8f0; }
    </style>
</head>
<body>
<div class="mobile-container">
    <div class="d-flex justify-content-between align-items-center mb-3 bg-primary text-white p-3 rounded-4 shadow-sm">
        <div>
            <h5 class="fw-bold mb-0">👤 {{ agent_name }}</h5>
            <small>Hudud: <b>{{ agent_territory or 'Barchasi' }}</b></small>
        </div>
        <a href="https://t.me/" class="btn btn-sm btn-light text-primary fw-bold">Botga qaytish</a>
    </div>

    <div class="card-mobile">
        <h6 class="fw-bold text-primary mb-3"><i class="bi bi-cart-plus me-1"></i>Yangi Buyurtma Kiritish</h6>
        <form method="POST">
            <div class="mb-3">
                <label class="form-label small fw-bold">Do'konni tanlang:</label>
                <select name="shop_name" class="form-select" required>
                    <option value="">-- Do'kon tanlang --</option>
                    {% for s in shops %}<option value="{{ s['name'] }}">{{ s['name'] }} ({{ s['region'] }})</option>{% endfor %}
                </select>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Narx turi:</label>
                <select name="price_type" class="form-select">
                    <option value="optom">Optom (Ulgurji)</option>
                    <option value="chakana">Chakana</option>
                </select>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Izoh:</label>
                <input type="text" name="comment" class="form-control" placeholder="Buyurtma izohi...">
            </div>
            <hr>
            <div class="mb-2 fw-bold text-secondary small">Mahsulotlar:</div>
            <div id="mini-items">
                <div class="row g-2 mb-2 align-items-center item-row">
                    <div class="col-7">
                        <select name="product_name" class="form-select form-select-sm" required>
                            <option value="">-- Tovar --</option>
                            {% for p in products %}<option value="{{ p['name'] }}">{{ p['name'] }} ({{ "{:,.0f}".format(p['optom_price']) }})</option>{% endfor %}
                        </select>
                    </div>
                    <div class="col-4">
                        <input type="number" step="any" name="qty" class="form-control form-control-sm" placeholder="Miqdor" value="1" required>
                    </div>
                    <div class="col-1 text-center">
                        <button type="button" class="btn btn-sm text-danger p-0" onclick="this.closest('.item-row').remove()"><i class="bi bi-trash"></i></button>
                    </div>
                </div>
            </div>
            <button type="button" class="btn btn-sm btn-outline-secondary w-100 mb-3" onclick="addMiniRow()">+ Yana mahsulot qo'shish</button>
            <button type="submit" class="btn btn-primary w-100 py-2 fw-bold">Buyurtmani Yuborish</button>
        </form>
    </div>

    <div class="card-mobile">
        <h6 class="fw-bold mb-3">🖼 Do‘konchiga ko‘rsatish uchun katalog</h6>
        <div class="row g-2">{% for c in catalog_items %}<div class="col-6"><div class="card h-100">{% if c['image_filename'] %}<img src="/catalog-image/{{ c['image_filename'] }}" class="card-img-top" style="height:130px;object-fit:cover" alt="{{ c['name'] }}">{% endif %}<div class="card-body p-2"><b>{{ c['name'] }}</b><div class="text-success fw-bold">{{ '{:,.0f}'.format(c['price'] or 0) }} so‘m</div><small>{{ c['description'] or '' }}</small></div></div></div>{% else %}<div class="col-12 text-muted">Katalog hali joylanmagan.</div>{% endfor %}</div>
    </div>

    <div class="card-mobile">
        <h6 class="fw-bold mb-3"><i class="bi bi-list-check me-1"></i>Mening oxirgi buyurtmalarim</h6>
        <div class="table-responsive">
            <table class="table table-sm align-middle" style="font-size: 12px;">
                <thead><tr><th>ID/Sana</th><th>Do'kon</th><th>Summa</th><th>Status</th></tr></thead>
                <tbody>
                    {% for o in my_orders %}
                    <tr>
                        <td><b>#{{ o['id'] }}</b><br><small class="text-muted">{{ o['date'] }}</small></td>
                        <td>{{ o['shop_name'] }}</td>
                        <td><b>{{ "{:,.0f}".format(o['total_sum']) }}</b></td>
                        <td><span class="badge bg-secondary">{{ o['status'] }}</span></td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>

    <div class="card-mobile">
        <h6 class="fw-bold mb-3">🏪 Hududimdagi do‘konlar, qarz va tashrif rejasi</h6>
        
        <!-- 🔍 Qidirish inputi shu yerga qo'shiladi -->
        <div class="mb-3">
            <input type="text" id="agentShopSearch" class="form-control form-control-sm" placeholder="Do‘kon nomi, orienter (mo‘ljal) yoki firma bo‘yicha qidirish..." onkeyup="filterAgentShops()">
        </div>

        <div class="table-responsive">
            <table class="table table-sm align-middle" id="agentShopsTable" style="font-size:12px">
                <thead>
                    <tr><th>Do‘kon</th><th>Qarz</th><th>Keyingi tashrif</th><th>Amal</th></tr>
                </thead>
                <tbody>
                {% for s in shops %}
                <tr>
                    <td>
                        <b>{{ s['name'] }}</b><br>
                        <small class="text-primary">{{ s['firm'] or '' }}</small> <small class="text-muted">({{ s['territory'] or s['region'] or '—' }})</small><br>
                        <small class="text-secondary"><i class="bi bi-geo-alt"></i> Orienter: {{ s['landmark'] or 'yo‘q' }}</small><br>
                        <small class="text-muted">Oxirgi: {{ s['last_visit'] or 'qayd yo‘q' }}</small>
                    </td>
                    <td><b>{{ "{:,.0f}".format(s['debt'] or 0) }} so‘m</b></td>
                    <td>
                        {% if s['next_visit_date'] %}
                            <span class="badge {% if s['next_visit_date'] < today_iso %}bg-danger{% elif s['next_visit_date'] == today_iso %}bg-warning text-dark{% else %}bg-success{% endif %}">{{ s['next_visit_date'] }}</span>
                        {% else %}
                            <span class="badge bg-secondary">Reja yo‘q</span>
                        {% endif %}
                    </td>
                    <td>
                        <form method="POST" class="d-flex gap-1 align-items-center">
                            <input type="hidden" name="action" value="mark_visit">
                            <input type="hidden" name="shop_id" value="{{ s['id'] }}">
                            <select name="interval_days" class="form-select form-select-sm" style="min-width:82px">
                                <option value="1">1 kun</option>
                                <option value="3">3 kun</option>
                                <option value="7" selected>7 kun</option>
                                <option value="14">14 kun</option>
                            </select>
                            <button class="btn btn-sm btn-success text-nowrap" type="submit">Tashrif</button>
                        </form>
                    </td>
                </tr>
                {% else %}
                <tr><td colspan="4">Do‘konlar yo‘q</td></tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
<script>
    function addMiniRow() {
        let container = document.getElementById('mini-items');
        let row = container.querySelector('.item-row').cloneNode(true);
        row.querySelector('select').selectedIndex = 0;
        row.querySelector('input[name="qty"]').value = '1';
        container.appendChild(row);
    }

    // 🔍 Agent uchun do'konlarni qidirish funksiyasi (Nom, Orienter va Firma bo'yicha)
    function filterAgentShops() {
        let input = document.getElementById('agentShopSearch').value.toLowerCase();
        let rows = document.querySelectorAll('#agentShopsTable tbody tr');
        
        rows.forEach(row => {
            let text = row.innerText.toLowerCase();
            if (text.includes(input)) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        });
    }
</script>
</body>
</html>
"""


@app.route('/agent_app/<int:tg_id>', methods=['GET', 'POST'])
@app.route('/agent_app/<int:tg_id>/', methods=['GET', 'POST'])
def agent_webapp(tg_id):
    if not verify_staff_link(tg_id):
        return "Havola eskirgan yoki imzosi mos emas. Telegram botga qayting va '🌐 Web kabinetim' tugmasini qayta bosing.", 403
    conn = get_db_connection()
    user = conn.execute("SELECT * FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    if not user:
        conn.close()
        return "Foydalanuvchi topilmadi yoki ro'yxatdan o'tmagan.", 403
    
    agent_name = user['name']
    agent_territory = user['territory'] or ''
    
    if request.method == 'POST':
        action = request.form.get('action', '')
        if action == 'mark_visit':
            shop_id = int(safe_number(request.form.get('shop_id'), 0, minimum=1))
            interval_days = int(safe_number(request.form.get('interval_days', 7), 7, minimum=1, maximum=60))
            shop = conn.execute('SELECT * FROM shops WHERE id = ?', (shop_id,)).fetchone()
            allowed_shop = bool(shop) and (not agent_territory or not shop['territory'] or shop['territory'] == agent_territory)
            if not allowed_shop:
                conn.close()
                return 'Bu do‘kon sizning hududingizga biriktirilmagan.', 403
            now_dt = datetime.now()
            last_visit = now_dt.strftime('%Y-%m-%d %H:%M')
            next_visit = (now_dt + timedelta(days=interval_days)).strftime('%Y-%m-%d')
            conn.execute('UPDATE shops SET last_visit = ?, next_visit_date = ? WHERE id = ?', (last_visit, next_visit, shop_id))
            log_audit(conn, 'Agent do‘konga tashrifni qayd etdi', 'shop', shop_id, f"agent={agent_name}; tg_id={tg_id}; keyingi_tashrif={next_visit}")
            conn.commit()
            conn.close()
            return redirect(make_staff_web_url(tg_id))
        shop_name = request.form.get('shop_name')
        price_type = request.form.get('price_type', 'optom')
        comment = request.form.get('comment', '')
        product_names = request.form.getlist('product_name')
        qtys = request.form.getlist('qty')
        
        total_sum = 0
        cost_total = 0
        items_text = ''
        excel_cart_items = []
        
        for i in range(len(product_names)):
            p_name = product_names[i]
            if not p_name: continue
            try:
                qty = float(qtys[i])
            except:
                qty = 1
                
            prod = conn.execute('SELECT optom_price, chakana_price, cost_price, stock, id FROM products WHERE name = ?', (p_name,)).fetchone()
            if prod:
                price = prod['optom_price'] if price_type == 'optom' else (prod['chakana_price'] or prod['optom_price'])
                stock_p = prod['stock'] or 0
                if qty <= 0:
                    continue
                if stock_p < qty:
                    conn.rollback()
                    conn.close()
                    return 'Omborda mahsulot yetarli emas: ' + str(p_name), 400
                summa = price * qty
                total_sum += summa
                cost_total += safe_number(prod['cost_price']) * qty
                items_text += f'{p_name} - {qty}x = {summa:,.0f} so\'m\n'
                conn.execute('UPDATE products SET stock = ? WHERE id = ?', (stock_p - qty, prod['id']))
                record_stock_movement(conn, p_name, -qty, 'Buyurtma', 'Agent webapp')
                excel_cart_items.append({'name': p_name, 'qty': qty, 'price': price})
                
        bugun = datetime.now().strftime('%Y-%m-%d %H:%M')
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO orders (shop_name, agent_name, total_sum, items_text, status, date, price_type, discount, comment, cost_total) VALUES (?, ?, ?, ?, "Yangi", ?, ?, 0, ?, ?)',
            (shop_name, agent_name, total_sum, items_text, bugun, price_type, comment, cost_total),
        )
        order_id = cursor.lastrowid
        log_audit(conn, 'Agent web kabinetidan buyurtma yaratildi', 'order', order_id, f'agent={agent_name}; do‘kon={shop_name}; summa={total_sum}; telegram_id={tg_id}')
        conn.commit()
        
        try:
            group_text = (
                f'📝 <b>YANGI MINI APP BUYURTMA (#{order_id})</b>\n'
                f"🏪 <b>Do'kon:</b> {shop_name}\n"
                f'👤 <b>Agent:</b> {agent_name}\n'
                f'💬 <b>Izoh:</b> {comment}\n'
                f'💰 <b>Jami summa:</b> {total_sum:,.0f} so\'m\n\n'
                f'<b>Mahsulotlar:</b>\n{items_text}'
            )
            bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')
        except Exception as e:
            print('Mini App buyurtmani guruhga yuborish xatosi:', e)
            
        conn.close()
        return redirect(make_staff_web_url(tg_id))

    if agent_territory:
        shops = conn.execute("SELECT * FROM shops WHERE territory = ? OR territory = '' ORDER BY name", (agent_territory,)).fetchall()
    else:
        shops = conn.execute("SELECT * FROM shops ORDER BY name").fetchall()
        
    products = conn.execute("SELECT * FROM products ORDER BY name").fetchall()
    catalog_items = conn.execute('SELECT * FROM catalog_items WHERE active = 1 ORDER BY id DESC').fetchall()
    my_orders = conn.execute("SELECT * FROM orders WHERE agent_name = ? ORDER BY id DESC LIMIT 20", (agent_name,)).fetchall()
    conn.close()
    
    return render_template_string(
        AGENT_WEBAPP_TEMPLATE,
        agent_name=agent_name,
        agent_territory=agent_territory,
        shops=shops,
        products=products,
        catalog_items=catalog_items,
        my_orders=my_orders,
        today_iso=datetime.now().strftime('%Y-%m-%d')
    )


@bot.message_handler(func=lambda message: message.text == '🛒 Yangi Buyurtma Urish')
def start_order_choice(message):
    tg_id = message.from_user.id
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.row(types.KeyboardButton('📅 Posesheniya (Tashrif kunlari)'), types.KeyboardButton('🌐 Obshiy zakaz (Barcha do\'konlar)'))
    
    web_markup = types.InlineKeyboardMarkup()
    web_markup.add(types.InlineKeyboardButton("🌐 Web App orqali buyurtma berish (Qulay)", web_app=types.WebAppInfo(url=make_staff_web_url(tg_id))))
    
    markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
    bot.send_message(message.chat.id, "Buyurtma berish turini tanlang (yoki quyidagi Mini App orqali o'ting):", reply_markup=markup)
    bot.send_message(message.chat.id, "📱 **Telefon uchun veb ilova (Mini App):**", reply_markup=web_markup, parse_mode="Markdown")


@bot.message_handler(func=lambda message: message.text in ['📅 Posesheniya (Tashrif kunlari)', '🌐 Obshiy zakaz (Barcha do\'konlar)'])
def start_order(message):
  tg_id = message.from_user.id
  order_type = message.text
  conn = get_db_connection()
  
  user = conn.execute('SELECT territory, role, name FROM users WHERE tg_id = ?', (tg_id,)).fetchone()
  agent_territory = user['territory'] if user and user['territory'] else ''
  agent_name = user['name'] if user else ''
  is_admin = (tg_id == ADMIN_ID or (user and 'admin' in user['role']))
  
  if 'Posesheniya' in order_type:
      shops = conn.execute('SELECT name, visit_days, region FROM shops').fetchall()
      conn.close()
      
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      text = "📋 **Posesheniya bo'yicha do'konlar (Qaysi biri qolgani ko'rinib turadi):**\n\n"
      for s in shops:
          ordered_today = conn.execute("SELECT id FROM orders WHERE shop_name = ? AND agent_name = ? AND date LIKE ?", (s['name'], agent_name, f"{datetime.now().strftime('%Y-%m-%d')}%")).fetchone()
          status_icon = "✅ Zakaz qilingan" if ordered_today else "❌ Qolgan (Zakaz qilinmagan)"
          
          markup.add(types.KeyboardButton(s['name']))
          text += f"🏢 {s['name']} — {status_icon} (Kunlar: _{s['visit_days'] or 'belgilanmagan'}_)\n"
      markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
      bot.send_message(message.chat.id, text, parse_mode="Markdown")
      msg = bot.send_message(message.chat.id, "Do'konni tanlang:", reply_markup=markup)
  else:
      if not is_admin and agent_territory:
        shops = conn.execute('SELECT name FROM shops WHERE territory = ?', (agent_territory,)).fetchall()
      else:
        shops = conn.execute('SELECT name FROM shops').fetchall()
      conn.close()
      
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      for s in shops: markup.add(types.KeyboardButton(s['name']))
      markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
      msg = bot.send_message(message.chat.id, 'Do\'konni tanlang (Obshiy):', reply_markup=markup)

  bot.register_next_step_handler(msg, choose_price_type)


def choose_price_type(message):
  if message.text == '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)':
    switch_role_menu(message)
    return
  user_steps[message.from_user.id] = {
      'shop_name': message.text,
      'cart': {},
      'price_type': None,
  }
  markup = types.ReplyKeyboardMarkup(resize_keyboard=True).row(
      types.KeyboardButton('💰 Ulgurji (Optom)'),
      types.KeyboardButton('🛍 Chakana'),
  )
  markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))
  msg = bot.send_message(
      message.chat.id, 'Narx turini tanlang:', reply_markup=markup
  )
  bot.register_next_step_handler(msg, show_categories_to_agent)


def show_categories_to_agent(message):
  if message.text == '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)':
    switch_role_menu(message)
    return
  p_type = 'optom' if 'Ulgurji' in message.text else 'chakana'
  user_steps[message.from_user.id]['price_type'] = p_type

  conn = get_db_connection()
  categories = conn.execute(
      'SELECT DISTINCT category FROM products'
  ).fetchall()
  conn.close()

  markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
  for c in categories:
    cat_name = c['category'] if c['category'] else 'Boshqa'
    markup.add(types.KeyboardButton(f'📁 {cat_name}'))
  markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))
  markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)'))

  msg = bot.send_message(
      message.chat.id,
      '📁 Mahsulot kategoriyasini tanlang yoki buyurtmani yakunlang:',
      reply_markup=markup,
  )
  bot.register_next_step_handler(msg, choose_product_category)


def choose_product_category(message):
  if message.text == '🔄 Rolni almashtirish (Sex / Agent / Zavsklad / Ekspeditor)':
    switch_role_menu(message)
    return
  if message.text == '✅ Buyurtmani yakunlash':
    ask_order_comment(message)
    return

  cat_name = message.text.replace('📁 ', '').strip()
  user_steps[message.from_user.id]['current_category'] = cat_name

  conn = get_db_connection()
  prods = conn.execute(
      'SELECT name FROM products WHERE category = ?', (cat_name,)
  ).fetchall()
  conn.close()

  if not prods:
    conn = get_db_connection()
    prods = conn.execute('SELECT name FROM products').fetchall()
    conn.close()

  markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
  for p in prods:
    markup.add(types.KeyboardButton(p['name']))
  markup.add(types.KeyboardButton('🔙 Kategoriyalarga qaytish'))
  markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))

  msg = bot.send_message(
      message.chat.id,
      f'📦 <b>{cat_name}</b> kategoriyasidagi mahsulotni tanlang:',
      parse_mode='HTML',
      reply_markup=markup,
  )
  bot.register_next_step_handler(msg, ask_quantity_or_navigation)


def ask_quantity_or_navigation(message):
  if message.text == '🔙 Kategoriyalarga qaytish':
    show_categories_to_agent_again(message)
    return
  if message.text == '✅ Buyurtmani yakunlash':
    ask_order_comment(message)
    return

  conn = get_db_connection()
  p_check = conn.execute(
      'SELECT id FROM products WHERE name = ?', (message.text,)
  ).fetchone()
  conn.close()

  if not p_check:
    bot.send_message(
        message.chat.id,
        "❌ Bunday mahsulot topilmadi. Iltimos, tugmalardan tanlang:",
    )
    choose_product_category(message)
    return

  user_steps[message.from_user.id]['current_product'] = message.text
  msg = bot.send_message(
      message.chat.id,
      f"🔢 <b>{message.text}</b> miqdorini kiriting:",
      parse_mode='HTML',
      reply_markup=types.ReplyKeyboardRemove(),
  )
  bot.register_next_step_handler(msg, add_to_cart)


def show_categories_to_agent_again(message):
  uid = message.from_user.id
  conn = get_db_connection()
  categories = conn.execute(
      'SELECT DISTINCT category FROM products'
  ).fetchall()
  conn.close()

  markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
  for c in categories:
    cat_name = c['category'] if c['category'] else 'Boshqa'
    markup.add(types.KeyboardButton(f'📁 {cat_name}'))
  markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))

  msg = bot.send_message(
      message.chat.id, '📁 Kategoriyani tanlang:', reply_markup=markup
  )
  bot.register_next_step_handler(msg, choose_product_category)


def add_to_cart(message):
  uid = message.from_user.id
  try:
    qty = float(message.text)
    p_name = user_steps[uid]['current_product']
    user_steps[uid]['cart'][p_name] = qty
    bot.send_message(message.chat.id, f'📥 Qo\'shildi: {p_name} - {qty}')

    cat_name = user_steps[uid].get('current_category', 'Boshqa')
    conn = get_db_connection()
    prods = conn.execute(
        'SELECT name FROM products WHERE category = ?', (cat_name,)
    ).fetchall()
    conn.close()

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for p in prods:
      markup.add(types.KeyboardButton(p['name']))
    markup.add(types.KeyboardButton('🔙 Kategoriyalarga qaytish'))
    markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))

    msg = bot.send_message(
        message.chat.id,
        'Yana mahsulot qo\'shasizmi yoki buyurtmani yakunlaysizmi?',
        reply_markup=markup,
    )
    bot.register_next_step_handler(msg, ask_quantity_or_navigation)
  except:
    bot.send_message(message.chat.id, '❌ Xatolik! Faqat raqam kiriting.')
    show_categories_to_agent_again(message)


def ask_order_comment(message):
  uid = message.from_user.id
  data = user_steps.get(uid)
  if not data or not data.get('cart'):
    bot.send_message(
        message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu('agent')
    )
    return

  msg = bot.send_message(
      message.chat.id,
      "📝 Buyurtma uchun **kommentariya (izoh)** yozing (masalan: <i>'Ertalabga yetkazilsin'</i>):",
      parse_mode='HTML',
      reply_markup=types.ReplyKeyboardRemove(),
  )
  bot.register_next_step_handler(msg, finish_order_with_comment)


def finish_order_with_comment(message):
  uid = message.from_user.id
  comment = message.text if message.text else ''
  data = user_steps.get(uid)

  if not data or not data['cart']:
    bot.send_message(
        message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu('agent')
    )
    return

  conn = get_db_connection()
  total_sum, cost_total, items_text, p_type = 0, 0, '', data['price_type']
  excel_cart_items = []

  for p_name, qty in data['cart'].items():
    prod = conn.execute(
        'SELECT optom_price, chakana_price, cost_price, stock, id FROM products WHERE name ='
        ' ?',
        (p_name,),
    ).fetchone()
    optom_p = prod['optom_price'] if prod['optom_price'] is not None else 0
    chakana_p = (
        prod['chakana_price'] if prod['chakana_price'] is not None else 0
    )
    stock_p = prod['stock'] if prod['stock'] is not None else 0

    price = optom_p if p_type == 'optom' else chakana_p
    if qty <= 0 or stock_p < qty:
      conn.close()
      bot.send_message(message.chat.id, f'❌ Omborda {p_name} yetarli emas yoki miqdor noto‘g‘ri.')
      return
    summa = price * qty
    total_sum += summa
    cost_total += safe_number(prod['cost_price']) * qty
    items_text += f'{p_name} - {qty}x = {summa:,.0f} so\'m\n'
    conn.execute(
        'UPDATE products SET stock = ? WHERE id = ?',
        (stock_p - qty, prod['id']),
    )
    record_stock_movement(conn, p_name, -qty, 'Buyurtma', 'Telegram bot')
    excel_cart_items.append({'name': p_name, 'qty': qty, 'price': price})

  shop_name = data['shop_name']

  agent_res = conn.execute(
      'SELECT name FROM users WHERE tg_id = ?', (uid,)
  ).fetchone()
  agent_name = agent_res['name'] if agent_res else 'Nomalum'
  bugun = datetime.now().strftime('%Y-%m-%d %H:%M')

  cursor = conn.cursor()
  cursor.execute(
      'INSERT INTO orders (shop_name, agent_name, total_sum, items_text,'
      " status, date, price_type, comment, cost_total) VALUES (?, ?, ?, ?, 'Yangi', ?, ?, ?, ?)",
      (shop_name, agent_name, total_sum, items_text, bugun, p_type, comment, cost_total),
  )
  order_id = cursor.lastrowid
  conn.commit()
  conn.close()

  admin_excel_file = None
  group_excel_file = None

  try:
    admin_excel_file = create_excel_invoice(
        order_id,
        data['shop_name'],
        agent_name,
        bugun,
        p_type,
        excel_cart_items,
        comment,
    )

    bot.send_message(
        message.chat.id,
        f'✅ Buyurtma qabul qilindi! Jami: {total_sum:,.0f} so\'m',
        reply_markup=get_main_menu('agent'),
    )
    bot.send_message(
        ADMIN_ID,
        f"🔔 YANGI BUYURTMA (#{order_id}):\n\nDo'kon: {data['shop_name']}\nSumma:"
        f" {total_sum:,.0f} so'm\nIzoh: {comment}",
    )

    with open(admin_excel_file, 'rb') as doc:
      bot.send_document(ADMIN_ID, doc, caption=f'📄 Nakladnoy (#{order_id})')

    group_text = (
        f'📝 <b>YANGI BUYURTMA (#{order_id})</b>\n'
        f"🏪 <b>Do'kon:</b> {data['shop_name']}\n"
        f'👤 <b>Agent:</b> {agent_name}\n'
        f'💬 <b>Izoh:</b> {comment}\n'
        f'💰 <b>Jami summa:</b> {total_sum:,.0f} so\'m\n\n'
        f'<b>Mahsulotlar:</b>\n{items_text}'
    )
    bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')

    group_excel_file = create_excel_invoice(
        order_id,
        data['shop_name'],
        agent_name,
        bugun,
        p_type,
        excel_cart_items,
        comment,
    )
    with open(group_excel_file, 'rb') as doc_group:
      bot.send_document(
          SEX_GROUP_ID, doc_group, caption=f'📄 Nakladnoy (#{order_id})'
      )

  except Exception as e:
    print('Buyurtmani yuborish xatosi:', e)

  finally:
    if admin_excel_file and os.path.exists(admin_excel_file):
      os.remove(admin_excel_file)
    if group_excel_file and os.path.exists(group_excel_file):
      os.remove(group_excel_file)

  if uid in user_steps:
    del user_steps[uid]


@bot.message_handler(func=lambda message: message.text == "📝 Ro'yxatdan o'tish")
def register_start(message):
  msg = bot.send_message(message.chat.id, 'Ismingizni kiriting:')
  bot.register_next_step_handler(msg, process_register_name)


def process_register_name(message):
  name = message.text
  msg = bot.send_message(message.chat.id, 'Telefon raqamingizni kiriting:')
  bot.register_next_step_handler(msg, lambda m: save_pending_user(m, name))


def save_pending_user(message, name):
  conn = get_db_connection()
  conn.execute(
      "INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, ?, ?, 'pending')",
      (message.from_user.id, name, message.text),
  )
  conn.commit()
  conn.close()
  bot.send_message(message.chat.id, "⏳ Admin tasdig'i kutilmoqda.")
  bot.send_message(ADMIN_ID, f'🔔 Yangi ro\'yxatdan o\'tgan: {name} ({message.text})')


@bot.message_handler(func=lambda message: message.text == '📜 Mening Buyurtmalarim')
def my_orders(message):
  conn = get_db_connection()
  res = conn.execute(
      'SELECT name FROM users WHERE tg_id = ?', (message.from_user.id,)
  ).fetchone()
  agent_name = res['name'] if res else ''
  orders = conn.execute(
      'SELECT shop_name, total_sum, status, date FROM orders WHERE agent_name ='
      ' ? ORDER BY id DESC LIMIT 5',
      (agent_name,),
  ).fetchall()
  conn.close()
  text = '<b>📜 Oxirgi buyurtmalar:</b>\n\n'
  for o in orders:
    t_sum = o['total_sum'] if o['total_sum'] is not None else 0
    text += f"🏪 {o['shop_name']} | {t_sum:,.0f} so'm | {o['status']} | {o['date']}\n\n"
  bot.send_message(message.chat.id, text, parse_mode='HTML')


# --- FLASK HTML SHABLONLARI ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Xoji Aka ERP — Premium Boshqaruv</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root { --bg-main: #f8fafc; --sidebar-bg: #0f172a; --primary-color: #3b82f6; }
        body { background-color: var(--bg-main); font-family: 'Inter', system-ui, sans-serif; color: #1e293b; }
        .sidebar { width: 150px; background: var(--sidebar-bg); min-height: 100vh; box-shadow: 4px 0 20px rgba(0,0,0,0.05); }
        .sidebar .nav-link { text-align: center; padding: 12px 6px; color: #94a3b8; font-size: 12px; font-weight: 500; border-radius: 10px; margin: 6px 8px; transition: all 0.25s ease; }
        .sidebar .nav-link:hover { background: rgba(255, 255, 255, 0.08); color: #ffffff; }
        .sidebar .nav-link.active { background: var(--primary-color); color: #ffffff; font-weight: 600; }
        .sidebar .nav-link i { font-size: 20px; display: block; margin-bottom: 4px; }
        .brand-logo-container { text-align: center; padding: 16px 5px 12px 5px; border-bottom: 1px solid rgba(255, 255, 255, 0.1); margin-bottom: 10px; }
        .brand-xa { font-family: 'Georgia', serif; font-weight: 900; font-size: 34px; line-height: 0.8; color: #ffffff; font-style: italic; }
        .brand-line { height: 3px; background-color: #ef4444; width: 60px; margin: 6px auto; border-radius: 2px; }
        .brand-name { font-weight: 800; font-size: 9px; letter-spacing: 2px; text-transform: uppercase; color: #cbd5e1; }
        .top-bar { background: #ffffff; border-bottom: 1px solid #e2e8f0; padding: 12px 25px; }
        .card-glass { background: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02); }
        .stat-box { border-radius: 14px; padding: 22px; color: white; position: relative; overflow: hidden; }
        .stat-blue { background: linear-gradient(135deg, #3b82f6, #1d4ed8); }
        .stat-green { background: linear-gradient(135deg, #10b981, #047857); }
        .stat-red { background: linear-gradient(135deg, #ef4444, #b91c1c); }
        .table-custom th { font-weight: 600; font-size: 13px; background: #f8fafc; color: #475569; }
        .table-custom td { font-size: 13px; vertical-align: middle; }
        .day-badge { display: inline-block; padding: 5px 8px; margin: 2px; border-radius: 6px; font-size: 11px; font-weight: bold; background: #e2e8f0; color: #475569; cursor: pointer; border: 1px solid #cbd5e1; user-select: none; }
        .day-badge.selected { background: #3b82f6; color: white; border-color: #2563eb; }
        .shop-row { cursor: pointer; }
        .shop-row:hover { background-color: #f1f5f9 !important; }
        .product-row { cursor: pointer; }
        .product-row:hover { background-color: #f1f5f9 !important; }
        @media (max-width: 768px) {
        .d-flex { flex-direction: column !important; }
        .sidebar { width: 100% !important; min-height: auto !important; }
        .sidebar .nav { flex-direction: row; overflow-x: auto; padding: 5px; }
        .sidebar .nav-link { font-size: 10px; padding: 8px 10px; margin: 2px; }
        .sidebar .nav-link i { font-size: 16px; margin-bottom: 0; }
        .brand-logo-container { display: none; }
        .top-bar { flex-direction: column; gap: 10px; align-items: stretch !important; }
        .top-bar form { width: 100%; flex-wrap: wrap; }
        .stat-box { margin-bottom: 10px; }
    }
    .table-responsive { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }
    table.table { white-space: nowrap; }
    </style>
</head>
<body>
<div class="d-flex">
    <div class="sidebar d-flex flex-column flex-shrink-0 nav nav-pills" id="v-pills-tab" role="tablist">
        <div class="brand-logo-container">
            <div class="brand-xa">XA</div>
            <div class="brand-line"></div>
            <div class="brand-name">Xoji Aka</div>
        </div>
        <button class="nav-link active" data-bs-toggle="pill" data-bs-target="#tab-dashboard" type="button"><i class="bi bi-grid-fill"></i>Bosh Panel</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-orders" type="button"><i class="bi bi-cart-fill"></i>Buyurtmalar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-inventory" type="button"><i class="bi bi-boxes"></i>Sklad</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-clients" type="button"><i class="bi bi-shop"></i>Do'konlar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-agents" type="button"><i class="bi bi-people-fill"></i>Agentlar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-kassa" type="button"><i class="bi bi-wallet2"></i>Kassa</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-reports" type="button"><i class="bi bi-file-earmark-bar-graph"></i>Hisobot</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-catalog" type="button"><i class="bi bi-images"></i>Katalog</button>
    </div>
    <div class="flex-grow-1">
        <div class="top-bar d-flex justify-content-between align-items-center">
            <div class="fw-bold text-dark fs-6"><i class="bi bi-shield-check text-primary me-2"></i>Boshqaruv Markazi</div>
            <div class="d-flex align-items-center gap-2">
                <a href="/logout" class="btn btn-sm btn-outline-danger fw-bold"><i class="bi bi-box-arrow-right me-1"></i>Chiqish</a>
                <a href="/export_excel" class="btn btn-sm btn-success fw-bold"><i class="bi bi-file-earmark-excel me-1"></i>Excelga Yuklab Olish</a>
                
                <form method="GET" action="/" class="d-flex align-items-center gap-1 m-0 bg-light p-1 rounded border">
                    <span class="text-muted small px-1">Dan:</span>
                    <input type="date" name="start_date" value="{{ start_date }}" class="form-control form-control-sm" style="width: 130px;">
                    <span class="text-muted small px-1">Gacha:</span>
                    <input type="date" name="end_date" value="{{ end_date }}" class="form-control form-control-sm" style="width: 130px;">
                    <button type="submit" class="btn btn-sm btn-primary">Saralash</button>
                    <a href="/" class="btn btn-sm btn-light border" title="Tozalash">✕</a>
                </form>
            </div>
        </div>
        <div class="p-4">
            <div class="tab-content">
                <div class="tab-pane fade show active" id="tab-dashboard">
                    <div class="row g-4 mb-4">
                        <div class="col-md-4">
                            <div class="stat-box stat-blue shadow-sm">
                                <div class="small text-white-50">Tanlangan Davr Tushumi</div>
                                <h2 class="fw-bold mt-1 mb-0">{{ "{:,.0f}".format(daily_sum) }} <span class="fs-6">so'm</span></h2>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="stat-box stat-green shadow-sm">
                                <div class="small text-white-50">Kassadagi Naqd Pul</div>
                                <h2 class="fw-bold mt-1 mb-0">{{ "{:,.0f}".format(kassa_balance) }} <span class="fs-6">so'm</span></h2>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="stat-box stat-red shadow-sm">
                                <div class="small text-white-50">Umumiy Nasiya (Qarzlar)</div>
                                <h2 class="fw-bold mt-1 mb-0">{{ "{:,.0f}".format(total_debt) }} <span class="fs-6">so'm</span></h2>
                            </div>
                        </div>
                    </div>
                    <div class="row g-4">
                        <div class="col-md-5">
                            <div class="card-glass p-4 h-100 d-flex flex-column align-items-center justify-content-center">
                                <h6 class="fw-bold mb-3 text-secondary w-100"><i class="bi bi-pie-chart-fill text-primary me-2"></i>Mahsulot Kategoriyalari Ulushi</h6>
                                <div style="width: 260px; height: 260px;"><canvas id="categoryDonutChart"></canvas></div>
                            </div>
                        </div>
                        <div class="col-md-7">
                            <div class="card-glass p-4 h-100">
                                <h6 class="fw-bold mb-3 text-secondary"><i class="bi bi-table me-2"></i>Kategoriyalar bo'yicha hisobot jadvali</h6>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle">
                                        <thead><tr><th>Kategoriya</th><th>Ulush (%)</th><th>Summa (so'm)</th></tr></thead>
                                        <tbody>
                                            {% for row in cat_table_data %}
                                            <tr>
                                                <td><span class="d-inline-block rounded-circle me-2" style="width: 10px; height: 10px; background-color: {{ row['color'] }};"></span><b>{{ row['name'] }}</b></td>
                                                <td><span class="badge bg-light text-dark border">{{ row['percent'] }}%</span></td>
                                                <td><b>{{ "{:,.0f}".format(row['amount']) }}</b></td>
                                            </tr>
                                            {% else %}
                                            <tr><td colspan="3" class="text-center py-4 text-muted">Ma'lumotlar mavjud emas</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-orders">
                    <div class="row g-4">
                        <div class="col-md-4">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold text-primary mb-3"><i class="bi bi-cart-plus me-2"></i>Yangi Buyurtma Kiritish</h5>
                                <form action="/add_order" method="POST">
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Do'kon qidirish va tanlash:</label>
                                        <input type="text" id="orderShopSearchInput" class="form-control form-control-sm mb-2" placeholder="Do'kon nomidan qidirish..." onkeyup="filterShopDropdown()">
                                        <select name="shop_name" id="orderShopSelect" class="form-select" required>
                                            <option value="">-- Do'konni tanlang --</option>
                                            {% for s in shops %}<option value="{{ s['name'] }}">{{ s['name'] }} ({{ s['region'] }})</option>{% endfor %}
                                        </select>
                                    </div>
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Operator / Agent:</label>
                                        <input type="text" name="agent_name" class="form-control" value="Admin (Web)" required>
                                    </div>
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Chegirma (Skidka so'mda):</label>
                                        <input type="number" step="any" name="discount" class="form-control" value="0">
                                    </div>
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Izoh (Kommentariya):</label>
                                        <input type="text" name="comment" class="form-control" placeholder="Buyurtma uchun izoh...">
                                    </div>
                                    <hr>
                                    <label class="form-label small fw-bold text-secondary mb-2"><i class="bi bi-basket me-1"></i>Mahsulotlar Savatchasi:</label>
                                    <div id="order-items-container">
                                        <div class="row g-2 mb-2 order-item-row align-items-center">
                                            <div class="col-7">
                                                <input type="text" class="form-control form-control-sm mb-1 product-search-input" placeholder="Tovar qidirish..." onkeyup="filterProductDropdown(this)">
                                                <select name="product_name" class="form-select form-select-sm product-select-dropdown" required>
                                                    <option value="">-- Tovar tanlang --</option>
                                                    {% for p in products %}<option value="{{ p['name'] }}">{{ p['name'] }} (Sklad: {{ p['stock'] }}) - {{ "{:,.0f}".format(p['optom_price']) }} so'm</option>{% endfor %}
                                                </select>
                                            </div>
                                            <div class="col-4"><input type="number" step="any" name="qty" class="form-control form-control-sm" placeholder="Miqdor" value="1" required></div>
                                            <div class="col-1 text-center"><button type="button" class="btn btn-sm text-danger p-0" onclick="removeRow(this)"><i class="bi bi-trash fs-6"></i></button></div>
                                        </div>
                                    </div>
                                    <button type="button" class="btn btn-sm btn-outline-secondary w-100 mb-3 border-dashed" onclick="addItemRow()"><i class="bi bi-plus-circle me-1"></i> Yana mahsulot qo'shish</button>
                                    <button type="submit" class="btn btn-primary w-100 py-2 fw-bold shadow-sm">Buyurtmani Tasdiqlash</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-8">
                            <div class="card-glass p-4 mb-3">
                                <h6 class="fw-bold mb-2"><i class="bi bi-printer me-1"></i>Tanlangan zakazlarni chop etish</h6>
                                <form action="/print_nakladnoy" method="GET" target="_blank" class="d-flex gap-2 align-items-center">
                                    <input type="text" id="selectedIdsInput" name="ids" class="form-control form-control-sm" placeholder="Tanlangan ID lar (Masalan: 1,2,3)" required readonly>
                                    <button type="submit" class="btn btn-sm btn-dark text-nowrap">Chop etish</button>
                                </form>
                            </div>
                            <div class="card-glass p-4">
                                <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
                                    <h5 class="fw-bold mb-0"><i class="bi bi-list-check me-2"></i>Barcha Buyurtmalar</h5>
                                    <div class="d-flex gap-2">
                                        <select id="statusFilter" class="form-select form-select-sm" style="width: 140px;" onchange="filterOrdersTable()">
                                            <option value="">Barcha statuslar</option>
                                            <option value="Yangi">Yangi</option>
                                            <option value="Otgruzka">Otgruzka</option>
                                            <option value="Yetkazildi">Yetkazildi</option>
                                            <option value="Bekor">Bekor</option>
                                        </select>
                                        <input type="text" id="orderSearch" class="form-control form-control-sm" placeholder="Buyurtma qidirish..." style="width: 180px;" onkeyup="filterOrdersTable()">
                                    </div>
                                </div>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle" id="ordersTable">
                                        <thead>
                                            <tr>
                                                <th style="width: 30px;"><input type="checkbox" id="selectAllOrders" onclick="toggleSelectAllOrders(this)"></th>
                                                <th>ID / Vaqt</th>
                                                <th>Do'kon</th>
                                                <th>Summa & Tafsilot</th>
                                                <th>Status & Tarix</th>
                                                <th>Amallar</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {% for o in orders %}
                                            <tr data-status="{{ o['status'] }}">
                                                <td><input type="checkbox" class="order-checkbox" value="{{ o['id'] }}" onclick="updateSelectedIds()"></td>
                                                <td><b>#{{ o['id'] }}</b><br><small class="text-muted">{{ o['date'] }}</small></td>
                                                <td><b>{{ o['shop_name'] }}</b><br><small class="text-muted">Agent: {{ o['agent_name'] }}</small></td>
                                                <td>
                                                    <div class="fw-bold text-primary mb-1">Jami: {{ "{:,.0f}".format(o['total_sum']) }} so'm</div>
                                                    <details>
                                                        <summary class="text-secondary small" style="cursor: pointer;">Tafsilotni ko'rish</summary>
                                                        <code style="font-size: 11px; white-space: pre-line;" class="text-dark d-block mt-1 bg-light p-1 rounded">{{ o['items_text'] }}</code>
                                                    </details>
                                                    {% if o['comment'] %}<small class="text-info fw-bold d-block mt-1"><i class="bi bi-chat-left-text me-1"></i>Izoh: {{ o['comment'] }}</small>{% endif %}
                                                    {% if o['discount'] and o['discount'] > 0 %}<small class="text-danger">Skidka: -{{ "{:,.0f}".format(o['discount']) }} so'm</small><br>{% endif %}
                                                </td>
                                                <td>
                                                    <span class="badge {% if o['status'] == 'Bekor' %}bg-danger{% elif o['status'] == 'Yetkazildi' %}bg-success{% else %}bg-secondary{% endif %} mb-1">{{ o['status'] }}</span>
                                                    <div style="font-size: 11px; color: #64748b;" class="mt-1">
                                                        {% if o['history'] %}
                                                            {% for h in o['history'] %}<div><b>{{ h['status'] }}:</b> {{ h['changed_at'] }}</div>{% endfor %}
                                                        {% else %}
                                                            <div>Yaratildi: {{ o['date'] }}</div>
                                                        {% endif %}
                                                    </div>
                                                </td>
                                                <td>
                                                    {% if o['status'] == 'Yangi' %}
                                                    <button type="button" class="btn btn-sm btn-outline-primary py-0 px-2 mb-1 w-100" style="font-size: 11px;" onclick='openEditOrderModal({{ o["id"]|tojson }}, {{ o["shop_name"]|tojson }}, {{ (o["discount"] or 0)|tojson }}, {{ o["comment"]|tojson }}, {{ (o["price_type"] or "optom")|tojson }}, {{ o["items_text"]|tojson }})'> <i class="bi bi-pencil"></i> Tahrir </button>
                                                    <form action="/update_status/{{ o['id'] }}" method="POST" class="d-flex gap-1 mb-1">
                                                        <select name="status" class="form-select form-select-sm" style="font-size: 11px; width: 90px;">
                                                            <option value="Yangi" {% if o['status'] == 'Yangi' %}selected{% endif %}>Yangi</option>
                                                            <option value="Otgruzka" {% if o['status'] == 'Otgruzka' %}selected{% endif %}>Otgruzka</option>
                                                            <option value="Yetkazildi" {% if o['status'] == 'Yetkazildi' %}selected{% endif %}>Yetkazildi</option>
                                                            <option value="Bekor" {% if o['status'] == 'Bekor' %}selected{% endif %}>Bekor</option>
                                                        </select>
                                                        <button type="submit" class="btn btn-sm btn-success px-2"><i class="bi bi-check"></i></button>
                                                    </form>
                                                    {% endif %}
                                                    <a href="/print_nakladnoy/{{ o['id'] }}" target="_blank" class="btn btn-sm btn-outline-dark py-0 w-100" style="font-size: 11px;"><i class="bi bi-printer me-1"></i>Nakladnoy</a>
                                                </td>
                                            </tr>
                                            {% else %}
                                            <tr><td colspan="6" class="text-center text-muted py-4">Buyurtmalar yo'q</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-inventory">
                    <div class="row g-4 mb-4">
                        <div class="col-md-12">
                            <div class="card-glass p-3 bg-light border">
                                <h6 class="fw-bold mb-2"><i class="bi bi-file-earmark-excel text-success me-2"></i>Exceldan Mahsulotlar Bazasini Yuklash (Import)</h6>
                                <form action="/import_products_excel" method="POST" enctype="multipart/form-data" class="d-flex align-items-center gap-3">
                                    <input type="file" name="excel_file" class="form-control form-control-sm" accept=".xlsx, .xls" required style="max-width: 350px;">
                                    <button type="submit" class="btn btn-sm btn-success fw-bold text-nowrap"><i class="bi bi-upload me-1"></i>Excelni Skladga Yuklash</button>
                                    <small class="text-muted">(Excel ustunlari: <b>name</b>, <b>category</b>, <b>stock</b>, <b>cost_price</b>, <b>optom_price</b>)</small>
                                </form>
                            </div>
                        </div>
                    </div>
                    <div class="row g-4">
                        <div class="col-md-5">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-box-arrow-in-down me-2 text-primary"></i>Tovar Kirim / Perexod</h5>
                                <form action="/add_stock" method="POST">
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Mahsulot:</label>
                                        <select name="product_select" id="p_sel" class="form-select" onchange="toggleNewProd()" required>
                                            <option value="">-- Tanlang --</option>
                                            <option value="NEW" class="fw-bold text-primary">+ Yangi mahsulot qo'shish</option>
                                            {% for p in products %}<option value="{{ p['name'] }}">{{ p['name'] }}</option>{% endfor %}
                                        </select>
                                    </div>
                                    <div id="new_prod_div" class="d-none bg-light p-3 rounded mb-3 border">
                                        <div class="mb-2"><label class="form-label small text-primary fw-bold">Yangi nom:</label><input type="text" name="new_product_name" id="new_p_name" class="form-control"></div>
                                        <div><label class="form-label small text-primary fw-bold">Kategoriya:</label><input type="text" name="new_product_category" class="form-control" placeholder="Masalan: Ichimliklar"></div>
                                    </div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Miqdor:</label><input type="number" step="any" name="qty" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Tan narxi:</label><input type="number" step="any" name="cost_price" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Optom narxi:</label><input type="number" step="any" name="optom_price" class="form-control" required></div>
                                    <button type="submit" class="btn btn-primary w-100">Omborga Saqlash</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-7">
                            <div class="card-glass p-4 mb-4">
                                <div class="d-flex justify-content-between align-items-center mb-3">
                                    <h5 class="fw-bold mb-0"><i class="bi bi-boxes me-2"></i>Ombordagi Qoldiqlar va Narxlar</h5>
                                    <input type="text" id="productSearchInput" class="form-control form-control-sm" placeholder="Mahsulot qidirish..." style="width: 200px;" onkeyup="filterProductsTable()">
                                </div>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle" id="productsTable">
                                        <thead><tr><th>Mahsulot</th><th>Qoldiq</th><th>Tannarx</th><th>Optom narx</th><th>Chakana narx</th><th>Sklad +-</th></tr></thead>
                                        <tbody>
                                            {% for p in products %}
                                            <tr>
                                                <td class="product-row" onclick='openEditProductModal({{ p["id"]|tojson }}, {{ p["name"]|tojson }}, {{ p["category"]|tojson }}, {{ p["stock"]|tojson }}, {{ p["cost_price"]|tojson }}, {{ p["optom_price"]|tojson }}, {{ p["chakana_price"]|tojson }}, {{ p["agent_percent"]|tojson }})'>
                                                    <b>{{ p['name'] }}</b><br><small class="text-muted">{{ p['category'] }}</small>
                                                </td>
                                                <td><span class="badge bg-success fs-6">{{ p['stock'] }}</span></td>
                                                <td>{{ "{:,.0f}".format(p['cost_price']) }}</td>
                                                <td>{{ "{:,.0f}".format(p['optom_price']) }}</td>
                                                <td>{{ "{:,.0f}".format(p['chakana_price']) }}</td>
                                                <td>
                                                    <form action="/adjust_stock_web" method="POST" class="d-flex gap-1 align-items-center">
                                                        <input type="hidden" name="prod_id" value="{{ p['id'] }}">
                                                        <input type="number" step="any" name="qty" class="form-control form-control-sm" placeholder="Miqdor" required style="width: 70px;">
                                                        <button type="submit" name="action" value="plus" class="btn btn-sm btn-outline-success py-0 px-2 fw-bold" title="Qoldiqqa qo'shish">+</button>
                                                        <button type="submit" name="action" value="minus" class="btn btn-sm btn-outline-danger py-0 px-2 fw-bold" title="Qoldiqdan ayirish">-</button>
                                                    </form>
                                                </td>
                                            </tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                            
                            <!-- Skladga kirim qilingan mahsulotlar tarixi (Product incomes history) -->
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-clock-history me-2 text-primary"></i>Skladga Tovar Kirim Tarixi (Product Incomes History)</h5>
                                <div class="table-responsive" style="max-height: 250px;">
                                    <table class="table table-sm table-bordered">
                                        <thead><tr><th>Sana</th><th>Mahsulot / Tafsilot</th><th>Miqdor</th></tr></thead>
                                        <tbody>
                                            {% for inc in product_incomes_history %}
                                            <tr>
                                                <td><small>{{ inc['date'] }}</small></td>
                                                <td>{{ inc['product_name'] }}</td>
                                                <td class="text-success fw-bold">+{{ inc['qty'] }}</td>
                                            </tr>
                                            {% else %}
                                            <tr><td colspan="3" class="text-center text-muted">Kirim tarixi mavjud emas</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-clients">
                    <div class="row g-4 mb-4">
                        <div class="col-md-12">
                            <div class="card-glass p-3 bg-light border">
                                <h6 class="fw-bold mb-2"><i class="bi bi-file-earmark-excel text-success me-2"></i>Exceldan Do'konlar Bazasini Yuklash (Import)</h6>
                                <form action="/import_shops_excel" method="POST" enctype="multipart/form-data" class="d-flex align-items-center gap-3">
                                    <input type="file" name="excel_file" class="form-control form-control-sm" accept=".xlsx, .xls" required style="max-width: 350px;">
                                    <button type="submit" class="btn btn-sm btn-success fw-bold text-nowrap"><i class="bi bi-upload me-1"></i>Excelni Yuklash</button>
                                    <small class="text-muted">(Excel ustunlari: <b>Do'kon nomi</b>, <b>Telefon</b>, <b>Hudud</b>, <b>Orienter</b>, <b>Qarz</b>)</small>
                                </form>
                            </div>
                        </div>
                    </div>
                    <div class="row g-4">
                        <div class="col-md-4">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-shop-window me-2 text-success"></i>Yangi Do'kon Qo'shish</h5>
                                <form action="/add_shop" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">Nomi:</label><input type="text" name="name" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Telefon:</label><input type="text" name="phone" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Territoriya (Hudud / Zona):</label><input type="text" name="territory" class="form-control" placeholder="Masalan: Chilonzor-1" required></div>
    
                                    <div class="mb-3"><label class="form-label small fw-bold">Hududi (Tarif):</label><input type="text" name="region" class="form-control" placeholder="Masalan: Chilonzor"></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Orienteri:</label><input type="text" name="landmark" class="form-control" placeholder="Masalan: Makro yonida"></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Berilgan inventarlar:</label><input type="text" name="inventory" class="form-control" placeholder="Masalan: 1 ta Xolodilnik"></div>
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Tashrif kunlari:</label>
                                        <div class="d-flex flex-wrap gap-1">
                                            <span class="day-badge" onclick="toggleDay(this, 'D')">D</span>
                                            <span class="day-badge" onclick="toggleDay(this, 'S')">S</span>
                                            <span class="day-badge" onclick="toggleDay(this, 'Ch')">Ch</span>
                                            <span class="day-badge" onclick="toggleDay(this, 'P')">P</span>
                                            <span class="day-badge" onclick="toggleDay(this, 'J')">J</span>
                                            <span class="day-badge" onclick="toggleDay(this, 'Sh')">Sh</span>
                                            <span class="day-badge" onclick="toggleDay(this, 'Ya')">Ya</span>
                                        </div>
                                        <input type="hidden" name="visit_days" id="visit_days_input">
                                    </div>
                                    <button type="submit" class="btn btn-success w-100">Saqlash</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-8">
                            <div class="card-glass p-4">
                                <div class="d-flex justify-content-between align-items-center mb-3">
                                    <h5 class="fw-bold mb-0"><i class="bi bi-people me-2"></i>Do'konlar Ro'yxati</h5>
                                    <input type="text" id="shopSearchInput" class="form-control form-control-sm" placeholder="Do'kon qidirish..." style="width: 200px;" onkeyup="filterShopsTable()">
                                </div>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle" id="shopsTable">
                                        <thead><tr><th>Do'kon / Tel</th><th>Hudud & Orienter</th><th>Tashrif kuni</th><th>Inventar</th><th>Qarz</th><th>Amallar</th></tr></thead>
                                        <tbody>
                                            {% for s in shops %}
                                            <tr class="shop-row" onclick='openEditShopModal({{ s["id"]|tojson }}, {{ s["name"]|tojson }}, {{ s["phone"]|tojson }}, {{ s["region"]|tojson }}, {{ s["landmark"]|tojson }}, {{ s["visit_days"]|tojson }}, {{ s["inventory"]|tojson }}, {{ s["territory"]|tojson }})'>
                                                <td><b>{{ s['name'] }}</b><br><small class="text-muted">{{ s['phone'] }}</small></td>
                                                <td><b>{{ s['region'] }}</b><br><small class="text-muted">{{ s['landmark'] }}</small></td>
                                                <td><span class="badge bg-light text-dark border">{{ s['visit_days'] }}</span></td>
                                                <td><small class="text-primary fw-bold">{{ s['inventory'] }}</small></td>
                                                <td><b class="text-danger">{{ "{:,.0f}".format(s['debt']) }} so'm</b></td>
                                                <td onclick="event.stopPropagation();" class="text-nowrap">
                                                    <button type="button" class="btn btn-sm btn-outline-primary py-0 px-2 me-1" onclick='openEditShopModal({{ s["id"]|tojson }}, {{ s["name"]|tojson }}, {{ s["phone"]|tojson }}, {{ s["region"]|tojson }}, {{ s["landmark"]|tojson }}, {{ s["visit_days"]|tojson }}, {{ s["inventory"]|tojson }}, {{ s["territory"]|tojson }})'><i class="bi bi-pencil"></i> Tahrir</button>
                                                    <form action="/pay_debt" method="POST" class="d-inline-flex gap-1 mt-1">
                                                        <input type="hidden" name="shop_id" value="{{ s['id'] }}">
                                                        <input type="number" name="amount" class="form-control form-control-sm" placeholder="To‘lov (+) / qarz (-)" title="Musbat summa — to‘lov; manfiy summa — eski qarzni qo‘lda kiritish" required style="width: 145px;">
                                                        <button type="submit" class="btn btn-sm btn-success" title="Musbat: to‘lov; manfiy: qarz kiritish">Kiritish</button>
                                                    </form>
                                                </td>
                                            </tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-agents">
                    <div class="row g-4">
                        <div class="col-md-4">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-person-plus-fill me-2 text-primary"></i>Yangi Xodim Qo'shish</h5>
                                <form action="/add_agent_web" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">F.I.O (Ismi):</label><input type="text" name="name" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Telefon Raqam:</label><input type="text" name="phone" class="form-control"></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Telegram ID:</label><input type="number" name="tg_id" class="form-control" placeholder="Masalan: 123456789" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Biriktirilgan Territoriya (Hudud):</label><input type="text" name="territory" class="form-control" placeholder="Masalan: Chilonzor-1"></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Oylik bazaviy foizi (%):</label><input type="number" step="any" name="salary_percent" class="form-control" value="5"></div>
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Rolini tanlang:</label>
                                        <select name="role" class="form-select">
                                            <option value="agent">Agent</option>
                                            <option value="sex">Sex xodimi</option>
                                            <option value="zavsklad">Zavsklad</option>
                                            <option value="ekspeditor">Ekspeditor</option>
                                            <option value="agent,sex">Agent va Sex xodimi</option>
                                        </select>
                                    </div>
                                    <button type="submit" class="btn btn-primary w-100">Foydalanuvchini Saqlash</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-8">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-people-fill me-2 text-primary"></i>Xodimlar Nazorati va Oyliklar</h5>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle">
                                        <thead><tr><th>F.I.O</th><th>Roli / Territoriya</th><th>Olingan Avans</th><th>Hisoblangan Oylik (+ / -)</th><th>Amallar</th></tr></thead>
                                        <tbody>
                                            {% for a in agents_salary_data %}
                                            <tr>
                                                <td><b>{{ a['name'] }}</b><br><small class="text-muted">{{ a['phone'] }}</small></td>
                                                <td><span class="badge bg-dark">{{ a['role'] }}</span><br><small class="text-primary">{{ a['territory'] or 'Barchasi' }}</small></td>
                                                <td><b class="text-danger">{{ "{:,.0f}".format(a['advance_sum']) }} so'm</b></td>
                                                <td>
                                                    <b class="text-success">{{ "{:,.0f}".format(a['calculated_salary']) }} so'm</b><br>
                                                    <small class="text-muted">Yakuniy (+/-): <b>{{ "{:,.0f}".format(a['final_salary']) }} so'm</b></small>
                                                </td>
                                                <td>
                                                    <button type="button" class="btn btn-sm btn-outline-primary py-0 px-2 me-1" onclick='openEditAgentModal({{ a["tg_id"]|tojson }}, {{ a["name"]|tojson }}, {{ a["phone"]|tojson }}, {{ a["role"]|tojson }}, {{ a["territory"]|tojson }}, {{ a["salary_percent"]|tojson }})'><i class="bi bi-pencil"></i> Tahrir</button>
                                                    <form action="/delete_agent/{{ a['tg_id'] }}" method="POST" class="d-inline" onsubmit="return confirm('O\'chirmoqchimisiz?');">
                                                        <button type="submit" class="btn btn-sm btn-outline-danger py-0 px-2"><i class="bi bi-trash"></i></button>
                                                    </form>
                                                </td>
                                            </tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-kassa">
                    <div class="row g-4 mb-4">
                        <div class="col-md-4">
                            <div class="card-glass p-4 border-start border-success border-4">
                                <h5 class="fw-bold text-success mb-3">Kirim qilish</h5>
                                <form action="/add_income" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">Manba:</label><input type="text" name="source" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Summa:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-success w-100">Kirim</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="card-glass p-4 border-start border-danger border-4">
                                <h5 class="fw-bold text-danger mb-3">Chiqim (Rasxod)</h5>
                                <form action="/add_expense" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">Sabab:</label><input type="text" name="reason" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Summa:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-danger w-100">Chiqim</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="card-glass p-4 border-start border-warning border-4">
                                <h5 class="fw-bold text-warning mb-3">Xodimga Avans Berish</h5>
                                <form action="/add_advance" method="POST">
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Xodim ismi:</label>
                                        <select name="worker_name" class="form-select" required>
                                            <option value="">-- Xodimni tanlang --</option>
                                            {% for u in users %}{% if u['role'] != 'admin' %}<option value="{{ u['name'] }}">{{ u['name'] }}</option>{% endif %}{% endfor %}
                                        </select>
                                    </div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Avans Summasi:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-warning w-100 text-dark fw-bold">Avans yozish</button>
                                </form>
                            </div>
                        </div>
                    </div>
                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-clock-history me-2 text-primary"></i>Kassa Tarixi (Kirim / Chiqim / Avanslar)</h5>
                        <div class="row">
                            <div class="col-md-6">
                                <h6 class="fw-bold text-success">📥 Kirimlar Tarixi:</h6>
                                <div class="table-responsive" style="max-height: 250px;">
                                    <table class="table table-sm table-bordered">
                                        <thead><tr><th>Sana</th><th>Manba</th><th>Summa</th></tr></thead>
                                        <tbody>
                                            {% for i in incomes_list %}
                                            <tr><td><small>{{ i['date'] }}</small></td><td>{{ i['source'] }}</td><td class="text-success fw-bold">{{ "{:,.0f}".format(i['amount']) }}</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <h6 class="fw-bold text-danger">📤 Chiqim va Avanslar Tarixi:</h6>
                                <div class="table-responsive" style="max-height: 250px;">
                                    <table class="table table-sm table-bordered">
                                        <thead><tr><th>Sana</th><th>Sabab / Xodim</th><th>Summa</th></tr></thead>
                                        <tbody>
                                            {% for e in expenses_list %}
                                            <tr><td><small>{{ e['date'] }}</small></td><td>{{ e['reason'] }}</td><td class="text-danger fw-bold">{{ "{:,.0f}".format(e['amount']) }}</td></tr>
                                            {% endfor %}
                                            {% for a in advances_list %}
                                            <tr><td><small>{{ a['date'] }}</small></td><td>[Avans] {{ a['worker_name'] }}</td><td class="text-warning fw-bold">{{ "{:,.0f}".format(a['amount']) }}</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-catalog">
                    <div class="card-glass p-4 mb-4">
                        <h5 class="fw-bold mb-3">🖼 Agentlar uchun mahsulotlar katalogi</h5>
                        <p class="text-muted small">Bu yerga mahsulot rasmi, do‘konga ko‘rsatish uchun narxi va izohini joylashtiring. Faol katalog agent web kabineti va Telegram botida ko‘rinadi.</p>
                        <form action="/catalog_add" method="POST" enctype="multipart/form-data" class="row g-3">
    <div class="col-md-3"><label class="form-label">Mahsulot nomi</label><input class="form-control" name="name" required maxlength="120"></div>
    
    <!-- 📌 KATEGORIYA INPUTINI SHU YERGA QO'SHASIZ -->
    <div class="col-md-2">
        <label class="form-label">Kategoriya</label>
        <input type="text" name="category" class="form-control" placeholder="Masalan: Yarim tayyor mahsulotlari" required>
    </div>

    <div class="col-md-2"><label class="form-label">Narxi (so‘m)</label><input class="form-control" type="number" min="0" step="any" name="price" required></div>
    <div class="col-md-3"><label class="form-label">Rasm (JPG/PNG/WEBP)</label><input class="form-control" type="file" name="image" accept=".jpg,.jpeg,.png,.webp" required></div>
    <div class="col-md-3"><label class="form-label">Izoh</label><input class="form-control" name="description" maxlength="500" placeholder="Og‘irligi, tarkibi va h.k."></div>
    <div class="col-md-1 d-flex align-items-end"><button class="btn btn-primary w-100" type="submit">Qo‘shish</button></div>
</form>
                    </div>
                    <div class="row g-3">
                    {% for c in catalog_items %}<div class="col-6 col-md-3"><div class="card h-100 shadow-sm border-0">{% if c['image_filename'] %}<img src="/catalog-image/{{ c['image_filename'] }}" class="card-img-top" style="height:180px;object-fit:cover" alt="{{ c['name'] }}">{% endif %}<div class="card-body"><h6 class="fw-bold">{{ c['name'] }}</h6><div class="text-success fw-bold">{{ '{:,.0f}'.format(c['price'] or 0) }} so‘m</div><p class="small text-muted">{{ c['description'] or '' }}</p><form action="/catalog_delete/{{ c['id'] }}" method="POST" onsubmit="return confirm('Katalogdan o‘chirilsinmi?')"><button class="btn btn-sm btn-outline-danger w-100">O‘chirish</button></form></div></div></div>{% else %}<p class="text-muted">Katalog hali bo‘sh.</p>{% endfor %}
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-reports">
                    <div class="card-glass p-4 mb-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-file-earmark-bar-graph me-2 text-primary"></i>Moliyaviy Hisobot (To'liq)</h5>
                        <div class="row g-3">
                            <div class="col-md-3 bg-light p-3 rounded"><small class="text-muted">Umumiy Savdo Tushumi</small><h4 class="fw-bold text-primary mb-0">{{ "{:,.0f}".format(total_revenue) }} so'm</h4></div>
                            <div class="col-md-3 bg-light p-3 rounded"><small class="text-muted">Sotilgan Tovarlar Tannarxi</small><h4 class="fw-bold text-secondary mb-0">{{ "{:,.0f}".format(total_cost) }} so'm</h4></div>
                            <div class="col-md-3 bg-light p-3 rounded"><small class="text-muted">Umumiy Rasxodlar (Chiqim)</small><h4 class="fw-bold text-danger mb-0">{{ "{:,.0f}".format(total_expense) }} so'm</h4></div>
                            <div class="col-md-3 bg-success bg-opacity-10 p-3 rounded border border-success"><small class="text-success fw-bold">Sof Foyda</small><h4 class="fw-bold text-success mb-0">{{ "{:,.0f}".format(net_profit) }} so'm</h4></div>
                        </div>
                    </div>
                    <div class="card-glass p-4 mb-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-box-seam me-2 text-primary"></i>Ombordagi Tovarlarning Umumiy Qiymati</h5>
                        <div class="row g-3 mb-3">
                            <div class="col-md-3"><div class="p-3 rounded bg-light h-100"><small class="text-muted">Kirim narxida (tannarx)</small><h4 class="fw-bold text-secondary mb-0">{{ "{:,.0f}".format(stock_cost_value) }} so'm</h4></div></div>
                            <div class="col-md-3"><div class="p-3 rounded bg-primary bg-opacity-10 h-100"><small class="text-primary">Sotish narxida (optom)</small><h4 class="fw-bold text-primary mb-0">{{ "{:,.0f}".format(stock_sale_value) }} so'm</h4></div></div>
                            <div class="col-md-3"><div class="p-3 rounded bg-info bg-opacity-10 h-100"><small class="text-info">Sotish narxida (chakana)</small><h4 class="fw-bold text-info mb-0">{{ "{:,.0f}".format(stock_retail_value) }} so'm</h4></div></div>
                            <div class="col-md-3"><div class="p-3 rounded bg-success bg-opacity-10 h-100"><small class="text-success">Taxminiy yalpi foyda (optom)</small><h4 class="fw-bold text-success mb-0">{{ "{:,.0f}".format(stock_sale_value - stock_cost_value) }} so'm</h4></div></div>
                        </div>
                        <p class="small text-muted mb-0">Qiymat amaldagi qoldiq × tannarx yoki optom narxi bo‘yicha hisoblanadi.</p>
                    </div>
                    <div class="card-glass p-4 mb-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-percent me-2 text-warning"></i>Har Bir Tovarning Ustama Foizi</h5>
                        <div class="table-responsive"><table class="table table-sm table-hover align-middle"><thead><tr><th>Mahsulot</th><th>Qoldiq</th><th>Tannarx</th><th>Optom</th><th>Chakana</th><th>Ustama optom</th><th>Ustama chakana</th></tr></thead><tbody>
                        {% for p in product_margin_rows %}<tr><td class="fw-bold">{{ p.name }}</td><td>{{ "{:,.2f}".format(p.stock) }}</td><td>{{ "{:,.0f}".format(p.cost) }}</td><td>{{ "{:,.0f}".format(p.optom) }}</td><td>{{ "{:,.0f}".format(p.chakana) }}</td><td><span class="badge {% if p.optom_percent < 0 %}bg-danger{% elif p.optom_percent < 10 %}bg-warning text-dark{% else %}bg-success{% endif %}">{{ "{:.1f}".format(p.optom_percent) }}%</span></td><td><span class="badge {% if p.chakana_percent < 0 %}bg-danger{% elif p.chakana_percent < 10 %}bg-warning text-dark{% else %}bg-success{% endif %}">{{ "{:.1f}".format(p.chakana_percent) }}%</span></td></tr>{% else %}<tr><td colspan="7">Mahsulotlar yo‘q.</td></tr>{% endfor %}
                        </tbody></table></div><small class="text-muted">Ustama = (sotish narxi − tannarx) ÷ tannarx × 100. Tannarx 0 bo‘lsa foiz 0% ko‘rsatiladi.</small>
                    </div>
                    <div class="card-glass p-4 mb-4">
                        <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-3"><h5 class="fw-bold mb-0"><i class="bi bi-calendar-check me-2 text-danger"></i>AKB / Do‘konlarga Tashrif Nazorati</h5><span class="badge bg-danger">Kechikkan: {{ overdue_visits_count }}</span></div>
                        <p class="small text-muted">Agent tashrifdan so‘ng Telegram web-kabinetida “Tashrif qildim” tugmasini bosib, keyingi tashrif muddatini belgilaydi.</p>
                        <div class="table-responsive"><table class="table table-sm table-hover align-middle"><thead><tr><th>Do‘kon</th><th>Hudud</th><th>Reja kunlari</th><th>Oxirgi tashrif</th><th>Keyingi tashrif</th><th>Holat</th></tr></thead><tbody>
                        {% for v in visit_alerts %}<tr><td class="fw-bold">{{ v.name }}</td><td>{{ v.territory or v.region or '—' }}</td><td>{{ v.visit_days or '—' }}</td><td>{{ v.last_visit or 'Hali qayd yo‘q' }}</td><td>{{ v.next_visit_date or 'Belgilanmagan' }}</td><td>{% if v.visit_status == 'Kechikkan' %}<span class="badge bg-danger">Kechikkan</span>{% elif v.visit_status == 'Bugun' %}<span class="badge bg-warning text-dark">Bugun borish kerak</span>{% elif v.visit_status == 'Reja yo‘q' %}<span class="badge bg-secondary">Reja yo‘q</span>{% else %}<span class="badge bg-success">Rejada</span>{% endif %}</td></tr>{% else %}<tr><td colspan="6">Do‘konlar ro‘yxati bo‘sh.</td></tr>{% endfor %}
                        </tbody></table></div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Modal oynalar -->
<div class="modal fade" id="editShopModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <form id="editShopForm" method="POST">
                <div class="modal-header">
                    <h5 class="modal-title fw-bold"><i class="bi bi-pencil-square text-primary me-2"></i>Do'konni Tahrirlash</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3"><label class="form-label small fw-bold">Nomi:</label><input type="text" name="name" id="edit_shop_name" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Telefon:</label><input type="text" name="phone" id="edit_shop_phone" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Territoriya (Hudud):</label><input type="text" name="territory" id="edit_shop_territory" class="form-control"></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Hududi:</label><input type="text" name="region" id="edit_shop_region" class="form-control"></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Orienteri:</label><input type="text" name="landmark" id="edit_shop_landmark" class="form-control"></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Tashrif kunlari:</label><input type="text" name="visit_days" id="edit_shop_visit_days" class="form-control"></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Inventarlar:</label><input type="text" name="inventory" id="edit_shop_inventory" class="form-control"></div>
                </div>
                <div class="modal-footer"><button type="submit" class="btn btn-primary btn-sm">Saqlash</button></div>
            </form>
        </div>
    </div>
</div>

<div class="modal fade" id="editProductModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <form id="editProductForm" method="POST">
                <div class="modal-header">
                    <h5 class="modal-title fw-bold"><i class="bi bi-pencil-square text-primary me-2"></i>Mahsulotni Tahrirlash</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3"><label class="form-label small fw-bold">Nomi:</label><input type="text" name="name" id="edit_prod_name" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Kategoriya:</label><input type="text" name="category" id="edit_prod_category" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Qoldiq:</label><input type="number" step="any" name="stock" id="edit_prod_stock" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Tan narx:</label><input type="number" step="any" name="cost_price" id="edit_prod_cost" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Optom narx:</label><input type="number" step="any" name="optom_price" id="edit_prod_optom" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Chakana narx:</label><input type="number" step="any" name="chakana_price" id="edit_prod_chakana" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Agent Foizi (%):</label><input type="number" step="any" name="agent_percent" id="edit_prod_percent" class="form-control" required></div>
                </div>
                <div class="modal-footer"><button type="submit" class="btn btn-primary btn-sm">Saqlash</button></div>
            </form>
        </div>
    </div>
</div>

<div class="modal fade" id="editAgentModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <form id="editAgentForm" method="POST">
                <div class="modal-header">
                    <h5 class="modal-title fw-bold"><i class="bi bi-pencil-square text-primary me-2"></i>Xodimni Tahrirlash</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3"><label class="form-label small fw-bold">F.I.O (Ismi):</label><input type="text" name="name" id="edit_agent_name" class="form-control" required></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Telefon Raqam:</label><input type="text" name="phone" id="edit_agent_phone" class="form-control"></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Territoriya (Hudud):</label><input type="text" name="territory" id="edit_agent_territory" class="form-control"></div>
                    <div class="mb-3"><label class="form-label small fw-bold">Bazaviy Oylik Foizi (%):</label><input type="number" step="any" name="salary_percent" id="edit_agent_salary_percent" class="form-control"></div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Roli:</label>
                        <select name="role" id="edit_agent_role" class="form-select">
                            <option value="agent">Agent</option>
                            <option value="sex">Sex xodimi</option>
                            <option value="zavsklad">Zavsklad</option>
                            <option value="ekspeditor">Ekspeditor</option>
                            <option value="agent,sex">Agent va Sex xodimi</option>
                        </select>
                    </div>
                </div>
                <div class="modal-footer"><button type="submit" class="btn btn-primary btn-sm">Saqlash</button></div>
            </form>
        </div>
    </div>
</div>
<div class="modal fade" id="editOrderModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <form id="editOrderForm" method="POST">
                <div class="modal-header">
                    <h5 class="modal-title fw-bold"><i class="bi bi-pencil-square text-primary me-2"></i>Buyurtma ichini tahrirlash</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <label class="form-label small fw-bold">Do'kon nomi:</label>
                            <input type="text" name="shop_name" id="edit_order_shop" class="form-control" required>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label small fw-bold">Narx turi:</label>
                            <select name="price_type" id="edit_order_price_type" class="form-select">
                                <option value="optom">Optom</option>
                                <option value="chakana">Chakana</option>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label small fw-bold">Skidka (so'm):</label>
                            <input type="number" step="any" name="discount" id="edit_order_discount" class="form-control" value="0">
                        </div>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Izoh:</label>
                        <input type="text" name="comment" id="edit_order_comment" class="form-control">
                    </div>
                    <hr>
                    <label class="form-label small fw-bold text-secondary mb-2">Mahsulotlar ro'yxati:</label>
                    <div id="edit-modal-items-container">
                        <!-- Dinamik qatorlar shu yerga tushadi -->
                    </div>
                    <button type="button" class="btn btn-sm btn-outline-secondary w-100 mt-2" onclick="addEditModalRow()">+ Yana mahsulot qo'shish</button>
                </div>
                <div class="modal-footer"><button type="submit" class="btn btn-primary btn-sm">O'zgarishlarni Saqlash</button></div>
            </form>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
let allProductsList = [
    {% for p in products %}
    {name: {{ p['name'] | tojson }}, price: {{ p['optom_price'] | tojson }}},
    {% endfor %}
];

function openEditOrderModal(id, shopName, discount, comment, priceType, itemsText) {
    document.getElementById('editOrderForm').action = '/update_order/' + id;
    document.getElementById('edit_order_shop').value = shopName;
    document.getElementById('edit_order_discount').value = discount;
    document.getElementById('edit_order_comment').value = comment;
    document.getElementById('edit_order_price_type').value = priceType || 'optom';

    let container = document.getElementById('edit-modal-items-container');
    container.innerHTML = '';

    let lines = itemsText.split('\\n');
    let hasItems = false;

    lines.forEach(line => {
        if (line.trim()) {
            hasItems = true;
            let parts = line.split(' - ');
            let pName = parts[0].trim();
            let qtyPart = '1';
            if (parts[1]) {
                let subParts = parts[1].split('x');
                if (subParts[0]) {
                    qtyPart = subParts[0].trim().replace(',', '.');
                }
            }
            
            addEditModalRow(pName, qtyPart);
        }
    });

    if (!hasItems) {
        addEditModalRow('', '1');
    }

    new bootstrap.Modal(document.getElementById('editOrderModal')).show();
}

function addEditModalRow(selectedProd = '', qtyVal = '1') {
    let container = document.getElementById('edit-modal-items-container');
    let row = document.createElement('div');
    row.className = 'row g-2 mb-2 align-items-center edit-item-row';
    
    let optionsHtml = '<option value="">-- Tovar tanlang --</option>';
    allProductsList.forEach(p => {
        let sel = (p.name === selectedProd) ? 'selected' : '';
        optionsHtml += `<option value="${p.name}" ${sel}>${p.name}</option>`;
    });

    row.innerHTML = `
        <div class="col-7">
            <select name="product_name" class="form-select form-select-sm" required>
                ${optionsHtml}
            </select>
        </div>
        <div class="col-4">
            <input type="number" step="any" name="qty" class="form-control form-control-sm" value="${qtyVal}" placeholder="Miqdor" required>
        </div>
        <div class="col-1 text-center">
            <button type="button" class="btn btn-sm text-danger p-0" onclick="this.closest('.edit-item-row').remove()"><i class="bi bi-trash"></i></button>
        </div>
    `;
    container.appendChild(row);
}
    function toggleNewProd() {
        let sel = document.getElementById('p_sel');
        let div = document.getElementById('new_prod_div');
        let nameInp = document.getElementById('new_p_name');
        if (sel.value === 'NEW') { div.classList.remove('d-none'); nameInp.required = true; }
        else { div.classList.add('d-none'); nameInp.required = false; }
    }
    function toggleDay(el, day) {
        el.classList.toggle('selected');
        let badges = document.querySelectorAll('.day-badge.selected');
        let days = Array.from(badges).map(b => b.innerText);
        document.getElementById('visit_days_input').value = days.join(', ');
    }
    function openEditShopModal(id, name, phone, region, landmark, visitDays, inventory, territory) {
        document.getElementById('editShopForm').action = '/update_shop/' + id;
        document.getElementById('edit_shop_name').value = name;
        document.getElementById('edit_shop_phone').value = phone;
        document.getElementById('edit_shop_region').value = region;
        document.getElementById('edit_shop_landmark').value = landmark;
        document.getElementById('edit_shop_visit_days').value = visitDays;
        document.getElementById('edit_shop_inventory').value = inventory ?? '';
        document.getElementById('edit_shop_territory').value = territory ?? '';
        new bootstrap.Modal(document.getElementById('editShopModal')).show();
    }
    function openEditProductModal(id, name, category, stock, cost, optom, chakana, percent) {
        document.getElementById('editProductForm').action = '/update_product/' + id;
        document.getElementById('edit_prod_name').value = name;
        document.getElementById('edit_prod_category').value = category;
        document.getElementById('edit_prod_stock').value = stock;
        document.getElementById('edit_prod_cost').value = cost;
        document.getElementById('edit_prod_optom').value = optom;
        document.getElementById('edit_prod_chakana').value = chakana;
        document.getElementById('edit_prod_percent').value = percent;
        new bootstrap.Modal(document.getElementById('editProductModal')).show();
    }
    function openEditAgentModal(tgId, name, phone, role, territory, salaryPercent) {
        document.getElementById('editAgentForm').action = '/update_agent/' + tgId;
        document.getElementById('edit_agent_name').value = name;
        document.getElementById('edit_agent_phone').value = phone;
        document.getElementById('edit_agent_role').value = role;
        document.getElementById('edit_agent_territory').value = territory;
        document.getElementById('edit_agent_salary_percent').value = salaryPercent;
        new bootstrap.Modal(document.getElementById('editAgentModal')).show();
    }
    function filterShopsTable() {
        let input = document.getElementById('shopSearchInput').value.toLowerCase();
        document.querySelectorAll('#shopsTable tbody tr').forEach(row => {
            row.style.display = row.innerText.toLowerCase().includes(input) ? '' : 'none';
        });
    }
    function filterProductsTable() {
        let input = document.getElementById('productSearchInput').value.toLowerCase();
        document.querySelectorAll('#productsTable tbody tr').forEach(row => {
            row.style.display = row.innerText.toLowerCase().includes(input) ? '' : 'none';
        });
    }
    function filterOrdersTable() {
        let statusFilter = document.getElementById('statusFilter').value;
        let search = document.getElementById('orderSearch').value.toLowerCase();
        document.querySelectorAll('#ordersTable tbody tr').forEach(row => {
            let status = row.getAttribute('data-status');
            let text = row.innerText.toLowerCase();
            row.style.display = ((!statusFilter || status === statusFilter) && (!search || text.includes(search))) ? '' : 'none';
        });
    }
    function filterShopDropdown() {
        let input = document.getElementById('orderShopSearchInput').value.toLowerCase();
        for (let option of document.getElementById('orderShopSelect').options) {
            if (option.value === "") continue;
            option.style.display = option.text.toLowerCase().includes(input) ? '' : 'none';
        }
    }
    function filterProductDropdown(inputElem) {
        let inputVal = inputElem.value.toLowerCase();
        let select = inputElem.closest('.order-item-row').querySelector('.product-select-dropdown');
        for (let option of select.options) {
            if (option.value === "") continue;
            option.style.display = option.text.toLowerCase().includes(inputVal) ? '' : 'none';
        }
    }
    function addItemRow() {
        let container = document.getElementById('order-items-container');
        let newRow = container.querySelector('.order-item-row').cloneNode(true);
        newRow.querySelector('.product-search-input').value = '';
        newRow.querySelector('select').selectedIndex = 0;
        newRow.querySelector('input[name="qty"]').value = '1';
        container.appendChild(newRow);
    }
    function removeRow(btn) {
        if (document.querySelectorAll('.order-item-row').length > 1) {
            btn.closest('.order-item-row').remove();
        } else { alert("Kamida bitta mahsulot bo'lishi kerak!"); }
    }
</script>
</body>
</html>
"""

@app.after_request
def audit_web_post(response):
    # Muhim web amallari uchun umumiy iz: endpoint, natija va parolsiz forma maydonlari.
    try:
        if request.method == 'POST' and request.endpoint not in ('login',) and response.status_code < 500:
            fields = {}
            for key in request.form.keys():
                if any(secret in key.lower() for secret in ('password', 'token', 'secret', 'sig')):
                    continue
                values = request.form.getlist(key)
                fields[key] = ','.join(values)[:160]
            conn = get_db_connection()
            conn.execute('INSERT INTO audit_log (actor, action, entity, entity_id, details, date) VALUES (?, ?, ?, ?, ?, ?)',
                (current_actor(), 'Web POST amali', request.endpoint or request.path, '', str(fields)[:1000] + f'; HTTP={response.status_code}', datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()
            conn.close()
    except Exception:
        app.logger.exception('Web amalini audit jurnaliga yozib bo‘lmadi')
    return response


@app.before_request
def restrict_access():
    allowed_routes = ['login', 'static', 'agent_webapp', 'staff_webapp', 'catalog_image', 'healthz']
    if request.endpoint not in allowed_routes and not session.get('logged_in'):
        return redirect(url_for('login'))

@app.route('/', methods=['GET'], endpoint='index')
def operator_dashboard():
  start_date = request.args.get('start_date', datetime.now().strftime('%Y-%m-%d'))
  end_date = request.args.get('end_date', datetime.now().strftime('%Y-%m-%d'))

  conn = get_db_connection()
  shops = conn.execute('SELECT * FROM shops ORDER BY name').fetchall()
  products = conn.execute('SELECT * FROM products ORDER BY name').fetchall()
  users = conn.execute('SELECT * FROM users').fetchall()
  product_incomes_history = conn.execute('SELECT * FROM product_incomes ORDER BY id DESC LIMIT 50').fetchall()
  catalog_items = conn.execute('SELECT * FROM catalog_items ORDER BY id DESC').fetchall()
  today_iso = datetime.now().strftime('%Y-%m-%d')

  orders = conn.execute("SELECT * FROM orders WHERE substr(date, 1, 10) BETWEEN ? AND ? ORDER BY id DESC", (start_date, end_date)).fetchall()

  formatted_orders = []
  for o in orders:
    o_dict = dict(o)
    history = conn.execute('SELECT status, changed_at FROM order_status_history WHERE order_id = ? ORDER BY id', (o['id'],)).fetchall()
    o_dict['history'] = [dict(h) for h in history]
    formatted_orders.append(o_dict)

  daily_sum_res = conn.execute("SELECT SUM(total_sum) as s FROM orders WHERE substr(date, 1, 10) BETWEEN ? AND ? AND status != 'Bekor'", (start_date, end_date)).fetchone()
  daily_sum = daily_sum_res['s'] if daily_sum_res and daily_sum_res['s'] else 0

  incomes_list = conn.execute('SELECT * FROM incomes ORDER BY id DESC LIMIT 50').fetchall()
  expenses_list = conn.execute('SELECT * FROM expenses ORDER BY id DESC LIMIT 50').fetchall()
  advances_list = conn.execute('SELECT * FROM advances ORDER BY id DESC LIMIT 50').fetchall()

  inc_sum = conn.execute('SELECT SUM(amount) as s FROM incomes').fetchone()['s'] or 0
  exp_sum = conn.execute('SELECT SUM(amount) as s FROM expenses').fetchone()['s'] or 0
  adv_sum = conn.execute('SELECT SUM(amount) as s FROM advances').fetchone()['s'] or 0

  kassa_balance = inc_sum - (exp_sum + adv_sum)

  total_debt_res = conn.execute('SELECT SUM(debt) as s FROM shops').fetchone()
  total_debt = total_debt_res['s'] if total_debt_res and total_debt_res['s'] else 0

  agents_salary_data = []
  for u in users:
      if 'agent' in u['role']:
          u_dict = dict(u)
          agent_orders = conn.execute("SELECT items_text, total_sum FROM orders WHERE agent_name = ? AND status != 'Bekor'", (u['name'],)).fetchall()
          
          calculated_salary = 0
          for ao in agent_orders:
              for line in ao['items_text'].split('\n'):
                  if line.strip():
                      try:
                          parts = line.split(' - ')
                          p_name = parts[0].strip()
                          qty_part = parts[1].split('x')[0].strip()
                          qty = float(qty_part)
                          
                          prod_info = conn.execute("SELECT optom_price, agent_percent FROM products WHERE name = ?", (p_name,)).fetchone()
                          if prod_info:
                              p_price = prod_info['optom_price'] or 0
                              p_perc = prod_info['agent_percent'] or 5
                              item_total = p_price * qty
                              calculated_salary += (item_total * p_perc / 100)
                      except:
                          pass
          
          adv_res = conn.execute("SELECT SUM(amount) as s FROM advances WHERE worker_name = ?", (u['name'],)).fetchone()
          total_adv = adv_res['s'] if adv_res and adv_res['s'] else 0
          
          u_dict['advance_sum'] = total_adv
          u_dict['calculated_salary'] = calculated_salary
          u_dict['final_salary'] = calculated_salary - total_adv
          agents_salary_data.append(u_dict)

  cat_data = conn.execute('SELECT category, SUM(stock * optom_price) as val FROM products GROUP BY category').fetchall()
  cat_table_data = []
  cat_chart_labels = []
  cat_chart_data = []
  colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899']

  total_stock_val = sum([row['val'] for row in cat_data if row['val']]) or 1
  for idx, row in enumerate(cat_data):
    c_name = row['category'] if row['category'] else 'Boshqa'
    c_val = row['val'] if row['val'] else 0
    percent = round((c_val / total_stock_val) * 100, 1)
    color = colors[idx % len(colors)]
    cat_chart_labels.append(c_name)
    cat_chart_data.append(c_val)
    cat_table_data.append({'name': c_name, 'percent': percent, 'amount': c_val, 'color': color})

  total_revenue_res = conn.execute("SELECT SUM(total_sum) as s FROM orders WHERE status != 'Bekor'").fetchone()
  total_revenue = total_revenue_res['s'] if total_revenue_res and total_revenue_res['s'] else 0

  # Sotilgan tovarlar tannarxini hisoblash (Total Cost of Goods Sold)
  all_valid_orders = conn.execute("SELECT id, items_text, cost_total FROM orders WHERE status != 'Bekor'").fetchall()
  total_cost = 0
  for ao in all_valid_orders:
      # Yangi buyurtmalar tannarxni buyurtma vaqtida saqlaydi. Eski buyurtmalar uchun
      # tarixiy tannarx mavjud bo'lmagani sababli amaldagi tannarxdan taxmin qilinadi.
      if ao['cost_total'] is not None:
          total_cost += safe_number(ao['cost_total'])
          continue
      for line in (ao['items_text'] or '').split('\n'):
          if line.strip():
              try:
                  parts = line.split(' - ')
                  p_name = parts[0].strip()
                  qty_part = parts[1].split('x')[0].strip().replace(',', '.')
                  qty = safe_number(qty_part)
                  prod_info = conn.execute("SELECT cost_price FROM products WHERE name = ?", (p_name,)).fetchone()
                  if prod_info:
                      total_cost += safe_number(prod_info['cost_price']) * qty
              except Exception:
                  pass

  # Sof foyda = Umumiy tushum - Sotilgan tovarlar tannarxi - Umumiy chiqimlar
  net_profit = total_revenue - total_cost - exp_sum

  stock_cost_value = 0.0
  stock_sale_value = 0.0
  stock_retail_value = 0.0
  product_margin_rows = []
  for prod in products:
      qty = safe_number(prod['stock'])
      cost = safe_number(prod['cost_price'])
      optom = safe_number(prod['optom_price'])
      chakana = safe_number(prod['chakana_price']) or optom
      stock_cost_value += qty * cost
      stock_sale_value += qty * optom
      stock_retail_value += qty * chakana
      optom_percent = ((optom - cost) / cost * 100) if cost > 0 else 0.0
      chakana_percent = ((chakana - cost) / cost * 100) if cost > 0 else 0.0
      product_margin_rows.append({'name': prod['name'], 'stock': qty, 'cost': cost, 'optom': optom, 'chakana': chakana, 'optom_percent': optom_percent, 'chakana_percent': chakana_percent})
  product_margin_rows.sort(key=lambda x: x['name'].casefold())

  visit_alerts = []
  for shop_row in shops:
      shop_dict = dict(shop_row)
      next_visit = (shop_dict.get('next_visit_date') or '').strip()
      if not next_visit:
          visit_status = 'Reja yo‘q'
      elif next_visit < today_iso:
          visit_status = 'Kechikkan'
      elif next_visit == today_iso:
          visit_status = 'Bugun'
      else:
          visit_status = 'Rejada'
      shop_dict['visit_status'] = visit_status
      visit_alerts.append(shop_dict)
  visit_priority = {'Kechikkan': 0, 'Bugun': 1, 'Reja yo‘q': 2, 'Rejada': 3}
  visit_alerts.sort(key=lambda x: (visit_priority.get(x['visit_status'], 4), x.get('next_visit_date') or '9999-99-99', (x.get('name') or '').casefold()))
  overdue_visits_count = sum(1 for v in visit_alerts if v['visit_status'] == 'Kechikkan')
  conn.close()

  return render_template_string(
      HTML_TEMPLATE,
      shops=shops,
      products=products,
      users=users,
      orders=formatted_orders,
      daily_sum=daily_sum,
      kassa_balance=kassa_balance,
      total_debt=total_debt,
      cat_table_data=cat_table_data,
      cat_chart_labels=cat_chart_labels,
      cat_chart_data=cat_chart_data,
      cat_chart_colors=colors[: len(cat_chart_labels)],
      total_revenue=total_revenue,
      total_cost=total_cost,
      total_expense=exp_sum,
      net_profit=net_profit,
      total_income=inc_sum,
      start_date=start_date,
      end_date=end_date,
      incomes_list=incomes_list,
      expenses_list=expenses_list,
      advances_list=advances_list,
      agents_salary_data=agents_salary_data,
      product_incomes_history=product_incomes_history,
      stock_cost_value=stock_cost_value,
      stock_sale_value=stock_sale_value,
      stock_retail_value=stock_retail_value,
      product_margin_rows=product_margin_rows,
      visit_alerts=visit_alerts,
      overdue_visits_count=overdue_visits_count,
      today_iso=today_iso,
      catalog_items=catalog_items,
  )


@app.route('/healthz')
def healthz():
  return {'ok': True, 'service': 'Xoji Aka', 'routes': ['/', '/agent_app/<tg_id>', '/staff_app/<tg_id>']}


@app.route('/catalog-image/<path:filename>')
def catalog_image(filename):
  # ImageKit URL bo‘lsa, to‘g‘ridan-to‘g‘ri shu rasmga yo‘naltiramiz.
  if filename.startswith('https://') or filename.startswith('http://'):
    return redirect(filename)

  # Eski, mahalliy papkada saqlangan rasmlar ham ishlashda davom etadi.
  return send_from_directory(
      CATALOG_UPLOAD_FOLDER,
      os.path.basename(filename)
  )




@app.route('/catalog_add', methods=['POST'])
def catalog_add():
    name = (request.form.get('name') or '').strip()
    category = (request.form.get('category') or 'Boshqa').strip()
    description = (request.form.get('description') or '').strip()
    price = safe_number(request.form.get('price'), -1)
    image = request.files.get('image')

    if (
        not name
        or price < 0
        or not image
        or not image.filename
        or not allowed_catalog_file(image.filename)
    ):
        return 'Mahsulot nomi, kategoriya, to‘g‘ri narx va JPG/PNG/WEBP rasm kiriting.', 400

    conn = None
    try:
        image_url = imagekitga_yuklash(image)

        conn = get_db_connection()
        conn.execute(
            '''INSERT INTO catalog_items
            (name, category, price, description, image_filename, active, created_at)
            VALUES (?, ?, ?, ?, ?, 1, ?)''',
            (
                name,
                category,
                price,
                description,
                image_url,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            )
        )

        log_audit(
            conn,
            'Katalogga mahsulot qo‘shildi',
            'catalog',
            image_url,
            f'nomi={name}; kategoriya={category}; narx={price}'
        )

        conn.commit()
        return redirect(url_for('index'))

    except Exception:
        if conn is not None:
            conn.rollback()
        app.logger.exception('Katalog rasmi yuklashda xato')
        return 'Rasmni yuklashda xatolik yuz berdi. Render loglarini tekshiring.', 500

    finally:
        if conn is not None:
            conn.close()


@app.route('/catalog_delete/<int:item_id>', methods=['POST'])
def catalog_delete(item_id):
  conn = get_db_connection()
  item = conn.execute('SELECT image_filename, name FROM catalog_items WHERE id = ?', (item_id,)).fetchone()
  if item:
    conn.execute('DELETE FROM catalog_items WHERE id = ?', (item_id,))
    log_audit(conn, 'Katalog mahsuloti o‘chirildi', 'catalog', item_id, item['name'])
    conn.commit()
  conn.close()
  if item and item['image_filename']:
    stored_image = item['image_filename']
    # Faqat eski mahalliy fayllarni o'chiramiz.
    # ImageKit URL'lariga tegmaymiz.
    if not stored_image.startswith(('https://', 'http://')):
      try:
        os.remove(
          os.path.join(
            CATALOG_UPLOAD_FOLDER,
            os.path.basename(stored_image)
          )
        )
      except OSError:
        pass
  return redirect(url_for('index'))


@app.route('/add_order', methods=['POST'])
def web_add_order():
  shop_name = request.form.get('shop_name')
  agent_name = request.form.get('agent_name', 'Admin (Web)')
  discount = safe_number(request.form.get('discount', 0), minimum=0)
  comment = request.form.get('comment', '')

  product_names = request.form.getlist('product_name')
  qtys = request.form.getlist('qty')

  conn = get_db_connection()
  total_sum = 0
  cost_total = 0
  items_text = ''
  excel_cart_items = []

  for i in range(len(product_names)):
    p_name = product_names[i]
    if not p_name:
      continue
    qty = safe_number(qtys[i] if i < len(qtys) else 1, minimum=0)
    if qty <= 0:
      continue

    prod = conn.execute('SELECT optom_price, cost_price, stock, id FROM products WHERE name = ?', (p_name,)).fetchone()
    if prod:
      price = prod['optom_price'] if prod['optom_price'] is not None else 0
      stock_p = prod['stock'] if prod['stock'] is not None else 0
      if stock_p < qty:
        conn.rollback()
        conn.close()
        return f"Omborda {p_name} yetarli emas. Qoldiq: {stock_p}", 400
      summa = price * qty
      total_sum += summa
      cost_total += safe_number(prod['cost_price']) * qty
      items_text += f'{p_name} - {qty}x = {summa:,.0f} so\'m\n'
      conn.execute('UPDATE products SET stock = ? WHERE id = ?', (stock_p - qty, prod['id']))
      record_stock_movement(conn, p_name, -qty, 'Buyurtma', 'Web panel')
      excel_cart_items.append({'name': p_name, 'qty': qty, 'price': price})

  final_sum = max(0, total_sum - discount)
  bugun = datetime.now().strftime('%Y-%m-%d %H:%M')

  cursor = conn.cursor()
  cursor.execute(
      'INSERT INTO orders (shop_name, agent_name, total_sum, items_text, status, date, price_type, discount, comment, cost_total) VALUES (?, ?, ?, ?, "Yangi", ?, ?, ?, ?, ?)',
      (shop_name, agent_name, final_sum, items_text, bugun, 'optom', discount, comment, cost_total),
  )
  order_id = cursor.lastrowid
  conn.commit()
  conn.close()

  try:
    group_text = (
        f'📝 <b>YANGI WEB BUYURTMA (#{order_id})</b>\n'
        f"🏪 <b>Do'kon:</b> {shop_name}\n"
        f'👤 <b>Operator:</b> {agent_name}\n'
        f'💬 <b>Izoh:</b> {comment}\n'
        f'💰 <b>Jami summa:</b> {final_sum:,.0f} so\'m\n\n'
        f'<b>Mahsulotlar:</b>\n{items_text}'
    )
    bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')
  except Exception as e:
    print('Web buyurtmani guruhga yuborish xatosi:', e)

  return redirect(url_for('index'))


@app.route('/update_status/<int:order_id>', methods=['POST'])
def update_status(order_id):
  new_status = request.form.get('status')
  bugun = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  
  allowed_statuses = {'Yangi', 'Otgruzka', 'Yetkazildi', 'Bekor'}
  if new_status not in allowed_statuses:
      conn.close()
      return 'Buyurtma holati noto‘g‘ri.', 400
  ord_data = conn.execute("SELECT shop_name, total_sum, status FROM orders WHERE id = ?", (order_id,)).fetchone()
  if not ord_data:
      conn.close()
      return 'Buyurtma topilmadi.', 404
  # Qarz faqat yetkazilgan holatga kirganda qo‘shiladi, undan chiqganda qaytariladi.
  old_delivered = ord_data['status'] == 'Yetkazildi'
  new_delivered = new_status == 'Yetkazildi'
  if old_delivered != new_delivered:
      shop_name = ord_data['shop_name']
      shop_res = conn.execute("SELECT debt FROM shops WHERE name = ?", (shop_name,)).fetchone()
      if shop_res:
          curr_debt = safe_number(shop_res['debt'])
          delta = safe_number(ord_data['total_sum']) * (1 if new_delivered else -1)
          conn.execute("UPDATE shops SET debt = ? WHERE name = ?", (max(0, curr_debt + delta), shop_name))
  conn.execute('UPDATE orders SET status = ? WHERE id = ?', (new_status, order_id))
  log_audit(conn, 'Buyurtma holati o‘zgartirildi', 'order', order_id, f'{ord_data["status"] if ord_data else "?"} -> {new_status}')
  conn.execute('INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, ?, ?)', (order_id, new_status, bugun))
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if ADMIN_PASSWORD and username == ADMIN_LOGIN and password == ADMIN_PASSWORD:
            session.clear()
            session['logged_in'] = True
            session['username'] = username
            try:
                conn = get_db_connection()
                log_audit(conn, 'Web panelga kirish', 'auth', username, 'Admin login muvaffaqiyatli')
                conn.commit()
                conn.close()
            except Exception:
                app.logger.exception('Kirish auditini yozishda xato')
            return redirect(url_for('index'))
        return "Login yoki parol xato! <a href='/login'>Qaytadan urinish</a>"
    return '''
        <form method="post" style="margin: 100px auto; width: 300px; font-family: sans-serif;">
            <h2>Xoji-Aka Tizimi</h2>
            <input type="text" name="username" placeholder="Login" required style="width:100%; padding:8px; margin-bottom:10px;"><br>
            <input type="password" name="password" placeholder="Parol" required style="width:100%; padding:8px; margin-bottom:10px;"><br>
            <button type="submit" style="width:100%; padding:10px; background:green; color:white; border:none;">Kirish</button>
        </form>
    '''


@app.route('/print_nakladnoy')
def print_multiple_nakladnoy():
  ids_str = request.args.get('ids', '')
  ids = [int(i.strip()) for i in ids_str.split(',') if i.strip().isdigit()]
  if not ids:
    return 'Buyurtmalar tanlanmagan'

  conn = get_db_connection()
  output = io.BytesIO()
  wb = Workbook()
  wb.remove(wb.active)

  for order_id in ids:
    order = conn.execute('SELECT * FROM orders WHERE id = ?', (order_id,)).fetchone()
    if not order:
      continue
    ws = wb.create_sheet(title=f'Zakaz_{order_id}')
    ws.sheet_view.showGridLines = True

    ws['A1'] = f'XOJI AKA FACTORY — NAKLADNOY #{order_id}'
    ws['A1'].font = Font(name='Arial', size=14, bold=True)
    ws['A3'] = f"Do'kon: {order['shop_name']}"
    ws['B3'] = f"Sana: {order['date']}"
    ws['A4'] = f"Agent: {order['agent_name']}"
    if order['comment']:
      ws['A5'] = f"Izoh: {order['comment']}"

    start_row = 7 if order['comment'] else 6
    headers = ['№', 'Mahsulot', 'Miqdor', 'Jami']
    for col_idx, h in enumerate(headers, 1):
      ws.cell(row=start_row, column=col_idx, value=h).font = Font(bold=True)

    row = start_row + 1
    for line in order['items_text'].split('\n'):
      if line.strip():
        ws.cell(row=row, column=1, value='-')
        ws.cell(row=row, column=2, value=line)
        row += 1

    row += 1
    ws.cell(row=row, column=1, value=f"JAMI: {order['total_sum']:,.0f} so'm").font = Font(bold=True)

  wb.save(output)
  output.seek(0)
  conn.close()
  return send_file(
      output,
      as_attachment=True,
      download_name='Nakladnoylar.xlsx',
      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )


@app.route('/print_nakladnoy/<int:order_id>')
def print_single_nakladnoy(order_id):
  conn = get_db_connection()
  order = conn.execute('SELECT * FROM orders WHERE id = ?', (order_id,)).fetchone()
  conn.close()
  if not order:
    return 'Topilmadi'

  cart_items = []
  for line in order['items_text'].split('\n'):
    if line.strip():
      cart_items.append({'name': line, 'qty': 1, 'price': order['total_sum']})

  excel_file = create_excel_invoice(
      order_id,
      order['shop_name'],
      order['agent_name'],
      order['date'],
      order['price_type'] or 'optom',
      cart_items,
      order['comment'],
  )
  return send_file(excel_file, as_attachment=True)


@app.route('/add_stock', methods=['POST'])
def add_stock():
  product_select = request.form.get('product_select')
  qty = safe_number(request.form.get('qty', 0), minimum=0)
  cost_price = safe_number(request.form.get('cost_price', 0), minimum=0)
  optom_price = safe_number(request.form.get('optom_price', 0), minimum=0)
  if qty <= 0:
    return 'Kirim miqdori noldan katta bo‘lishi kerak.', 400

  conn = get_db_connection()
  if product_select == 'NEW':
    name = request.form.get('new_product_name')
    category = request.form.get('new_product_category', 'Boshqa')
    conn.execute(
        'INSERT INTO products (name, category, stock, cost_price, optom_price, chakana_price) VALUES (?, ?, ?, ?, ?, ?)',
        (name, category, qty, cost_price, optom_price, optom_price),
    )
    prod_name_for_history = name
  else:
    conn.execute(
        'UPDATE products SET stock = stock + ?, cost_price = ?, optom_price = ? WHERE name = ?',
        (qty, cost_price, optom_price, product_select),
    )
    prod_name_for_history = product_select

  # Kirim tarixiga qo'shish
  bugun_str = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn.execute(
      'INSERT INTO product_incomes (product_name, qty, cost_price, date) VALUES (?, ?, ?, ?)',
      (prod_name_for_history, qty, cost_price, bugun_str)
  )
  record_stock_movement(conn, prod_name_for_history, qty, 'Kirim', 'Web panel')
  log_audit(conn, 'Ombor kirimi', 'product', prod_name_for_history, f'miqdor={qty}; tannarx={cost_price}')

  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/pay_debt', methods=['POST'])
def pay_debt():
  """Musbat summa — qarz to'lovi; manfiy summa — qo'lda eski qarz kiritish."""
  shop_id = request.form.get('shop_id')
  raw_amount = str(request.form.get('amount', '')).strip().replace(' ', '').replace(',', '.')
  try:
    amount = float(raw_amount)
    if amount != amount or amount in (float('inf'), float('-inf')):
      raise ValueError('Noto‘g‘ri summa')
  except (TypeError, ValueError):
    return 'Summa raqam ko‘rinishida kiritilishi kerak.', 400
  if amount == 0:
    return 'Summa 0 bo‘lmasligi kerak. To‘lov uchun musbat, eski qarzni kiritish uchun manfiy summa yozing.', 400

  today = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
  conn = get_db_connection()
  shop = conn.execute('SELECT id, name, debt FROM shops WHERE id = ?', (shop_id,)).fetchone()
  if not shop:
    conn.close()
    return 'Do‘kon topilmadi.', 404

  curr_debt = float(shop['debt'] or 0)
  new_debt = curr_debt - amount
  conn.execute('UPDATE shops SET debt = ? WHERE id = ?', (new_debt, shop['id']))
  if amount > 0:
    # Haqiqiy to'lov: qarz kamayadi va tushumga yoziladi.
    conn.execute(
        'INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)',
        (f"Qarz to‘lovi ({shop['name']})", amount, today),
    )
    log_audit(conn, 'Do‘kon qarzidan to‘lov ayirildi', 'shop', shop['id'],
              f"do‘kon={shop['name']}; to‘lov={amount:.2f}; qarz {curr_debt:.2f} -> {new_debt:.2f}")
  else:
    # Manfiy summa daromad emas: mavjud/oldingi qarzni qo'lda kiritish.
    log_audit(conn, 'Eski qarz qo‘lda kiritildi', 'shop', shop['id'],
              f"do‘kon={shop['name']}; kiritilgan qarz={abs(amount):.2f}; qarz {curr_debt:.2f} -> {new_debt:.2f}")
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/add_income', methods=['POST'])
def add_income():
  source, amount = request.form.get('source', '').strip(), safe_number(request.form.get('amount'), minimum=0)
  if not source or amount <= 0:
    return 'Manba va musbat summa kiriting.', 400
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  conn.execute('INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)', (source, amount, today))
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/add_expense', methods=['POST'])
def add_expense():
  reason, amount = request.form.get('reason', '').strip(), safe_number(request.form.get('amount'), minimum=0)
  if not reason or amount <= 0:
    return 'Xarajat sababi va musbat summa kiriting.', 400
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  conn.execute('INSERT INTO expenses (reason, amount, date) VALUES (?, ?, ?)', (reason, amount, today))
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/add_advance', methods=['POST'])
def add_advance():
  worker_name = request.form['worker_name']
  amount = safe_number(request.form.get('amount', 0), minimum=0)
  if amount <= 0:
    return 'Avans miqdori noldan katta bo‘lishi kerak.', 400
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  conn.execute('INSERT INTO advances (worker_name, amount, date) VALUES (?, ?, ?)', (worker_name, amount, today))
  conn.commit()
  conn.close()
  return redirect(url_for('index'))


@app.route('/logout')
def logout():
  session.pop('logged_in', None)
  return redirect(url_for('login'))


@app.route('/export_excel')
def export_excel():
  conn = get_db_connection()
  orders = pd.read_sql_query('SELECT * FROM orders', conn)
  shops = pd.read_sql_query('SELECT * FROM shops', conn)
  products = pd.read_sql_query('SELECT * FROM products', conn)
  conn.close()

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine='openpyxl') as writer:
    orders.to_excel(writer, sheet_name='Buyurtmalar', index=False)
    shops.to_excel(writer, sheet_name='Dokonlar', index=False)
    products.to_excel(writer, sheet_name='Sklad', index=False)
  output.seek(0)

  return send_file(
      output,
      as_attachment=True,
      download_name='Xoji_Aka_Hisobot.xlsx',
      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )



@app.errorhandler(500)
def handle_internal_server_error(error):
    # Batafsil xato server terminalida ko'rinadi; foydalanuvchiga ichki tafsilotlarni ko'rsatmaymiz.
    app.logger.exception("Web panelda kutilmagan 500 xatosi: %s", error)
    if session.get('logged_in'):
        back_url = url_for('index')
        back_label = "Boshqaruv paneliga qaytish"
    else:
        back_url = url_for('login')
        back_label = "Login sahifasiga qaytish"
    return (
        "<!doctype html><html lang='uz'><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>Xatolik</title><body style='font-family:Arial,sans-serif;max-width:680px;"
        "margin:70px auto;padding:24px;color:#222'>"
        "<h2>Amalni bajarishda xatolik yuz berdi</h2>"
        "<p>Ma'lumotlar saqlanmagan bo'lishi mumkin. Boshqaruv paneliga qayting va amalni yana tekshiring.</p>"
        f"<p><a href='{back_url}'>{back_label}</a></p>"
        "<p style='color:#666;font-size:13px'>Agar xato takrorlansa, server terminalidagi xabarni tekshiring.</p>"
        "</body></html>",
        500,
    )


def run_bot():
    while True:
        try:
            bot.remove_webhook()
            bot.infinity_polling(timeout=60, long_polling_timeout=60)
        except Exception as e:
            print(f"Bot polling xatosi: {e}")
            time.sleep(5)


def run_visit_reminders():
    """Har kuni Toshkent vaqti bilan 08:00 dan keyin agentlarga muddati kelgan AKB tashriflarini eslatadi."""
    last_sent_day = None
    tashkent_tz = timezone(timedelta(hours=5))
    while True:
        try:
            now = datetime.now(tashkent_tz)
            today = now.strftime('%Y-%m-%d')
            if now.hour >= 8 and now.hour < 20 and last_sent_day != today:
                conn = get_db_connection()
                agents = conn.execute("SELECT tg_id, name, territory FROM users WHERE role LIKE '%agent%'").fetchall()
                due_shops = conn.execute(
                    "SELECT name, territory, region, debt, next_visit_date FROM shops "
                    "WHERE COALESCE(next_visit_date, '') != '' AND next_visit_date <= ? "
                    "ORDER BY next_visit_date, name", (today,)
                ).fetchall()
                conn.close()
                for agent in agents:
                    territory = (agent['territory'] or '').strip()
                    relevant = [sh for sh in due_shops if (not territory or not sh['territory'] or sh['territory'] == territory)]
                    if not relevant:
                        continue
                    lines = [f"📅 <b>AKB TASHRIF ESLATMASI</b>", f"👤 Agent: <b>{agent['name']}</b>", ""]
                    for sh in relevant[:30]:
                        due_date = sh['next_visit_date'] or ''
                        status = '🔴 KECHIKKAN' if due_date < today else '🟡 BUGUN'
                        lines.append(f"{status} — <b>{sh['name']}</b> | {sh['territory'] or sh['region'] or 'Hudud belgilanmagan'} | Qarz: {safe_number(sh['debt']):,.0f} so‘m | Reja: {due_date}")
                    lines.append("\nWeb kabinetda tashrifni qayd etib, keyingi tashrif muddatini belgilang.")
                    try:
                        bot.send_message(int(agent['tg_id']), '\n'.join(lines), parse_mode='HTML')
                    except Exception as send_error:
                        app.logger.warning('AKB eslatmasi yuborilmadi (%s): %s', agent['tg_id'], send_error)
                last_sent_day = today
        except Exception:
            app.logger.exception('AKB tashrif eslatmasi xatosi')
        time.sleep(45)


if __name__ == '__main__':
    init_web_db()
    create_database_backup()
    threading.Thread(target=backup_scheduler, daemon=True).start()

    if not ADMIN_PASSWORD:
        print("DIQQAT: ADMIN_PASSWORD sozlanmagan. Admin kirishi uchun PowerShell’da $env:ADMIN_PASSWORD='kuchli_parol' deb belgilang.")

    if BOT_TOKEN and ':' in BOT_TOKEN and 'PLACEHOLDER' not in BOT_TOKEN:
        bot_thread = threading.Thread(target=run_bot, daemon=True)
        bot_thread.start()
        threading.Thread(target=run_visit_reminders, daemon=True).start()
    else:
        print('DIQQAT: BOT_TOKEN sozlanmagan. Web-panel ishlaydi, Telegram bot polling ishga tushmaydi.')

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
