from datetime import datetime
import io
import os
import sqlite3
from flask import (
    Flask,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import telebot
from telebot import types

# --- SOZLAMALAR ---
BOT_TOKEN = "8573337094:AAGhQXE6IheONVsJgxyLfpeyjAqY_xbtYJk"
ADMIN_ID = 6851851908
SEX_GROUP_ID = -1003936599812  # Foydalanuvchi ko'rsatgan guruh ID si

bot = telebot.TeleBot(BOT_TOKEN)
app = Flask(__name__)
DB_NAME = 'xoji_aka_factory.db'
app.secret_key = 'xoji_aka_maxfiy_kalit_2026'

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
  os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
user_steps = {}


def get_db_connection():
  conn = sqlite3.connect(DB_NAME)
  conn.row_factory = sqlite3.Row
  return conn


def init_web_db():
  conn = get_db_connection()
  cursor = conn.cursor()

  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS users (tg_id INTEGER PRIMARY KEY, name TEXT, phone TEXT, role TEXT DEFAULT 'pending', password TEXT DEFAULT '1234')'''
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
      '''CREATE TABLE IF NOT EXISTS pending_incomes 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, staff_name TEXT, amount REAL, reason TEXT, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS product_incomes 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT, qty REAL, cost_price REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS orders 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, shop_name TEXT, agent_name TEXT, items_text TEXT, total_sum REAL, discount REAL DEFAULT 0, status TEXT, date TEXT, price_type TEXT, comment TEXT DEFAULT '')'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS order_status_history 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER, status TEXT, changed_at TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS products 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, category TEXT DEFAULT 'Boshqa', stock REAL DEFAULT 0, cost_price REAL DEFAULT 0, optom_price REAL DEFAULT 0, chakana_price REAL DEFAULT 0)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS shops 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, phone TEXT, debt REAL DEFAULT 0, visit_days TEXT, region TEXT DEFAULT '', landmark TEXT DEFAULT '', inventory TEXT DEFAULT '')'''
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
      ("ALTER TABLE shops ADD COLUMN visit_days TEXT DEFAULT ''", "visit_days"),
      ("ALTER TABLE shops ADD COLUMN region TEXT DEFAULT ''", "region"),
      ("ALTER TABLE shops ADD COLUMN landmark TEXT DEFAULT ''", "landmark"),
      ("ALTER TABLE shops ADD COLUMN inventory TEXT DEFAULT ''", "inventory"),
      ("ALTER TABLE orders ADD COLUMN discount REAL DEFAULT 0", "discount"),
      ("ALTER TABLE orders ADD COLUMN price_type TEXT", "price_type"),
      ("ALTER TABLE orders ADD COLUMN comment TEXT DEFAULT ''", "comment"),
      ("ALTER TABLE users ADD COLUMN password TEXT DEFAULT '1234'", "password"),
  ]

  for query, col in migrations:
    try:
      cursor.execute(query)
    except:
      pass

  cursor.execute(
      "INSERT OR REPLACE INTO users (tg_id, name, phone, role, password) VALUES"
      " (?, 'Admin', '', 'admin', 'admin2026')",
      (ADMIN_ID,),
  )

  initial_agents = [
      (8241020136, "Qoraboyev Sirojiddin", "+998935075540", "agent,sex"),
      (2101923750, "Qoraboyeva Charos", "+998940300206", "agent"),
      (6851851908, "SRJ", "+998975155540", "agent,admin"),
  ]
  for ag_id, ag_name, ag_phone, ag_role in initial_agents:
    cursor.execute(
        "INSERT OR IGNORE INTO users (tg_id, name, phone, role, password)"
        " VALUES (?, ?, ?, ?, '1234')",
        (ag_id, ag_name, ag_phone, ag_role),
    )

  initial_shops = [
      ('Vanselling', '', 'SEX Gulim'),
      ('Sherzod market', '', 'Uzgazoil qatori'),
      ('Qayumov Kamol', '', 'Murch boboga yetmasdan'),
      ('Akbar aka', '', 'Bekat murch bobo'),
      ('Pub house', '', 'Pub house'),
      ('Chapayev roparasi', '', 'Chapayev roparasi'),
      ('Toshpoʻlat aka', '', 'Uchrashuv yoni'),
      ('Abdulfayz bekat', '', ''),
      ('Dilorom', '', 'Tutzor'),
      ('Joʻrabek aka', '', 'Anorcha tagi'),
      ('Vohas', '', 'Vohas'),
      ('Oʻzbegim market', '', 'Davr bank yoni'),
      ('Pokiza market', '', 'Muz saroy yoni'),
      ('Soxibkor doʻstlik', '', ''),
      ('Feruza non sex', '', 'Nigoh non sexi yoni'),
      ('South brothers', '', 'Doktor A qatori'),
      ('Gulmira opa', '', 'Cola orqasi'),
      ('Taniqulov Oʻktam', '', ''),
      ('Abdulloh market', '', 'Feredun café'),
      ('23-market', '', '23-sartarosh yoni'),
      ('Cola market', '', 'South brothersga yetmay'),
      ('Muxlis Market', '', 'Movaro'),
      ('Anjir Market', '', 'Movaro'),
      ('Sevimli Market', '', 'Yashil dunyo'),
      ('Abbos market', '', 'Yashil dunyo'),
      ('Dilya opa', '', 'Yashil dunyo'),
      ('Sanjar aka', '', 'Yashil dunyo'),
      ('Nur market', '', 'Yashil dunyo'),
      ('547 market', '', 'Yashil dunyo'),
      ('Shox market', '', 'Yashil dunyo'),
      ('Makro market', '', 'Yashil dunyo'),
      ('Xusan bobo market', '', 'Yashil dunyo'),
      ('Fresh market umid aka', '', 'Movaro'),
      ('Pul hokim', '', 'Adliya yoʻli'),
      ('Hoji ona market', '', 'Med yoni'),
      ('Chinor market', '', 'Med yoni'),
      ('16 market', '', 'Begoyim roʻparasi'),
      ('Lada yoni', '', 'Lada yoni'),
      ('Alibek aka', '', 'Sohil pastlik'),
      ('4 aka-uka', '', 'sohil'),
      ('Rayxon opa sohil', '', 'Guliston ma-si'),
      ('Muhabbat opa', '', 'Boyqishloq'),
      ('Moyka yoni', '', 'Moyka yoni'),
      ('Malika yoni optom', '', 'Malika yoni optom'),
      ('AR market', '', 'Abdurashid market'),
      ('Kam-kam Market', '', '23-dom yoni'),
      ('Osiyo tagi', '', ''),
      ('Nigora opa', '', 'Eski pioner oldi'),
      ('Sherbek aka/fresh M', '', 'Best roʻparasi'),
      ('Otabek aka', '', 'antena tagi'),
      ('Universal Market', '', 'Lola kafe yoni'),
      ('Aka-Uka Market', '', 'Masjid yoni'),
      ('Oila Market', '', 'zilyonni yoʻli'),
      ('Boxo market', '', 'zilyonni yoʻli'),
      ('Otabek zapchast M', '', 'zilyonni yoʻli'),
      ('Maya market', '', 'Gostsatndart yoni'),
      ('Umida opa', '', 'Mashhura yoʻli'),
      ('Darband city', '', 'Vokzal yoni'),
      ('Otajon market', '', 'Tisudan keyin'),
      ('Barakali market', '', 'Tisu yoni'),
      ('Asilabonu', '', 'Boysun bekati yoni'),
      ('Norqulova Nargiza', '', 'Hayit ala uyi taraf'),
      ('Sherzod aka', '', 'Harbiy doʻkon'),
      ('Bahor Market', '', 'Termiz tuman'),
      ('Baraka Market', '', 'Termiz tuman'),
      ('Farxod Market', '', 'Senter Kamaz'),
      ('Ariqcha Market', '', 'Senter Kamaz'),
      ('Xolida Market', '', 'Limon'),
      ('7Я Market', '', 'indenim'),
  ]
  for s_name, s_phone, s_region in initial_shops:
    cursor.execute(
        '''INSERT OR IGNORE INTO shops (name, phone, debt, region) VALUES (?, ?, 0, ?)''',
        (s_name, s_phone, s_region),
    )

  initial_products = [
      ('Pelmen /300 gr', 'Yarim Tayyor Mahsulotlari', 1000, 6200, 14000.0, 14000.0),
      ('Pelmen /500 gr', 'Yarim Tayyor Mahsulotlari', 1000, 9900, 24000.0, 24000.0),
      ('Pelmen rasepnoy /kg', 'Yarim Tayyor Mahsulotlari', 1000, 19500, 46000.0, 46000.0),
      ('Teftel /300 gr', 'Yarim Tayyor Mahsulotlari', 1000, 12000, 23000.0, 23000.0),
      ('Pelmen ossarti /500 gr', 'Yarim Tayyor Mahsulotlari', 1000, 18000, 30000.0, 30000.0),
      ('Pelmen ossarti /300 gr', 'Yarim Tayyor Mahsulotlari', 1000, 11500, 20000.0, 20000.0),
      ('Golubtsi /500 gr', 'Yarim Tayyor Mahsulotlari', 1000, 14000, 25000.0, 25000.0),
      ('Tok doʻlma /300 gr', 'Yarim Tayyor Mahsulotlari', 1000, 11000.0, 25000.0, 25000.0),
      ('Karam doʻlma /300 gr', 'Yarim Tayyor Mahsulotlari', 1000, 11000.0, 25000.0, 25000.0),
      ('Somsa kesilgan /800 gr', 'Yarim Tayyor Mahsulotlari', 1000, 6000.0, 17000.0, 17000.0),
      ('Oʻrama xamir', 'Yarim Tayyor Mahsulotlari', 1000, 6000.0, 17000.0, 17000.0),
      ('KFC', 'Yarim Tayyor Mahsulotlari', 1000, 27500, 30000.0, 30000.0),
      ('KFC Gulim', 'Yarim Tayyor Mahsulotlari', 1000.0, 15000, 30000.0, 30000.0),
      ('Osh masalliq 500 gr', 'Yarim Tayyor Mahsulotlari', 1000, 6000.0, 13000.0, 13000.0),
      ('Osh masalliq 1 kg', 'Yarim Tayyor Mahsulotlari', 1000, 7000.0, 15000.0, 15000.0),
      ('Lagʻmon', 'Yarim Tayyor Mahsulotlari', 10000, 2000.0, 6000.0, 6000.0),
      ('Kotlet', 'Yarim Tayyor Mahsulotlari', 1000, 11000.0, 25000.0, 25000.0),
      ('Lavash hamiri', 'Yarim Tayyor Mahsulotlari', 1000, 3800.0, 7000.0, 7000.0),
      ('Manti hamiri', 'Yarim Tayyor Mahsulotlari', 1000, 6000.0, 13000.0, 13000.0),
      ('Mirinda 250gr/30шт', 'Yaxna ichimliklari', 1000, 7950.0, 9000.0, 9000.0),
      ('Mirinda 330gr/24шт', 'Yaxna ichimliklari', 1000, 8800.0, 9000.0, 9000.0),
      ('Snikers/001/3kg', 'Afif shirinliklari', 100, 100500.0, 117000.0, 117000.0),
      ('Mini Rulet/002/3kg', 'Afif shirinliklari', 100, 100500.0, 117000.0, 117000.0),
      ('Yojik/003/2kg', 'Afif shirinliklari', 20, 67000.0, 840000.0, 84000.0),
      ('Dessert/007/2kg', 'Afif shirinliklari', 20, 71000.0, 860000.0, 86000.0),
      ('Pudra Palichka/013/2kg', 'Afif shirinliklari', 10, 49000.0, 60000.0, 60000.0),
      ('Pudra Kalso/014/2kg', 'Afif shirinliklari', 10, 49000.0, 60000.0, 60000.0),
      ('Ovsyanka pista/025/3kg', 'Afif shirinliklari', 10, 73500.0, 92000.0, 92000.0),
      ('Ovsyanka magiz/026/3kg', 'Afif shirinliklari', 10, 73500.0, 92000.0, 92000.0),
      ('Choko Ovsyanka/028/3kg', 'Afif shirinliklari', 10, 82500.0, 103000.0, 103000.0),
  ]
  for p_name, p_cat, p_stock, p_cost, p_optom, p_chakana in initial_products:
    cursor.execute(
        '''INSERT OR IGNORE INTO products (name, category, stock, cost_price, optom_price, chakana_price) 
           VALUES (?, ?, ?, ?, ?, ?)''',
        (p_name, p_cat, p_stock, p_cost, p_optom, p_chakana),
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

  headers = ['№', 'Mahsulot nomi', 'Miqdori', "Narxi (so'm)", 'Jami summa']
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


# --- ZAVSKLAD UCHUN YIG'MA EXCEL (BUNCHA BUYURTMALAR YIG'indisi) ---
def create_excel_consolidated_invoice(aggregated_items, order_ids_str):
  wb = Workbook()
  ws = wb.active
  ws.title = 'Zavsklad_Yigma'
  ws.sheet_view.showGridLines = True

  title_font = Font(name='Arial', size=16, bold=True)
  header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
  bold_font = Font(name='Arial', size=11, bold=True)
  header_fill = PatternFill(
      start_color='595959', end_color='595959', fill_type='solid'
  )
  total_fill = PatternFill(
      start_color='F2F2F2', end_color='F2F2F2', fill_type='solid'
  )
  thin_border = Border(
      left=Side(style='thin', color='B0B0B0'),
      right=Side(style='thin', color='B0B0B0'),
      top=Side(style='thin', color='B0B0B0'),
      bottom=Side(style='thin', color='B0B0B0'),
  )

  ws.merge_cells('A1:C1')
  ws['A1'] = '🏭 ZAVSKLAD UCHUN YIG\'MA MAHSULOTLAR RO\'YXATI'
  ws['A1'].font = title_font
  ws['A1'].alignment = Alignment(horizontal='center')

  ws['A3'] = f'Buyurtmalar ID lari: #{order_ids_str}'
  ws['A3'].font = bold_font
  ws['C3'] = f'Sana: {datetime.now().strftime("%Y-%m-%d %H:%M")}'

  headers = ['№', 'Mahsulot nomi', 'Jami Miqdori (Soni/Kg)']
  for col_num, header_title in enumerate(headers, 1):
    cell = ws.cell(row=5, column=col_num, value=header_title)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

  row_num = 6
  total_qty = 0
  for idx, (p_name, qty) in enumerate(aggregated_items.items(), 1):
    ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(
        horizontal='center'
    )
    ws.cell(row=row_num, column=2, value=p_name).alignment = Alignment(
        horizontal='left'
    )
    ws.cell(row=row_num, column=3, value=qty).alignment = Alignment(
        horizontal='right'
    )
    ws.cell(row=row_num, column=3).number_format = '#,##0.##'
    total_qty += qty
    for col in range(1, 4):
      ws.cell(row=row_num, column=col).border = thin_border
    row_num += 1

  ws.merge_cells(
      start_row=row_num, start_column=1, end_row=row_num, end_column=2
  )
  ws.cell(row=row_num, column=1, value='JAMI YIG\'INDI MIQDOR:').alignment = (
      Alignment(horizontal='right')
  )
  ws.cell(row=row_num, column=1).font = bold_font
  tot_cell = ws.cell(row=row_num, column=3, value=total_qty)
  tot_cell.font = bold_font
  tot_cell.number_format = '#,##0.##'
  for col in range(1, 4):
    ws.cell(row=row_num, column=col).fill = total_fill
    ws.cell(row=row_num, column=col).border = thin_border

  file_name = f'Zavsklad_Yigma_{order_ids_str.replace(",", "_")}.xlsx'
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

  headers = ['№', 'Mahsulot nomi', 'Miqdori (Soni/Kg)']
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
  if role == 'admin':
    markup.row(
        types.KeyboardButton('📦 Sklad & Mahsulotlar'),
        types.KeyboardButton('📊 Kunlik Hisobot'),
    )
    markup.row(
        types.KeyboardButton('👥 Agentlar boshqaruvi'),
        types.KeyboardButton("🏪 AKB (Do'konlar & Qarz)"),
    )
  elif role == 'agent':
    markup.row(
        types.KeyboardButton('🛒 Yangi Buyurtma Urish'),
        types.KeyboardButton("🏪 Do'kon qo'shish (AKB)"),
    )
    markup.row(
        types.KeyboardButton("💰 Qarz/To'lov yozish"),
        types.KeyboardButton('📜 Mening Buyurtmalarim'),
    )
    markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
  elif role == 'sex':
    markup.row(
        types.KeyboardButton('📦 Skladga Mahsulot Kirim Qilish'),
        types.KeyboardButton('💰 Kassaga Kirim Qilish'),
    )
    markup.row(
        types.KeyboardButton('📋 Ombordagi Qoldiqlar'),
        types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'),
    )
  else:
    markup.add(types.KeyboardButton("📝 Ro'yxatdan o'tish"))
  return markup


@bot.message_handler(commands=['start'])
def start_command(message):
  tg_id = message.from_user.id
  if tg_id == ADMIN_ID:
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
    elif ',' in role:
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      if 'agent' in role:
        markup.add(types.KeyboardButton('👤 Agent rejimi'))
      if 'sex' in role:
        markup.add(types.KeyboardButton('🏭 Sex rejimi'))
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


@bot.message_handler(
    func=lambda message: message.text
    in ['👤 Agent rejimi', '🏭 Sex rejimi', '🔄 Rolni almashtirish (Sex / Agent)']
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

  role_str = user['role']
  if message.text == '👤 Agent rejimi' or (
      'agent' in role_str and 'sex' in role_str and message.text != '🏭 Sex rejimi'
  ):
    if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      markup.add(
          types.KeyboardButton('👤 Agent rejimi'),
          types.KeyboardButton('🏭 Sex rejimi'),
      )
      bot.send_message(
          message.chat.id, 'Qaysi rejimga oʻtmoqchisiz?', reply_markup=markup
      )
      return

    bot.send_message(
        message.chat.id,
        '🛒 Agent rejimiga oʻtdingiz.',
        reply_markup=get_main_menu('agent'),
    )
  elif message.text == '🏭 Sex rejimi':
    bot.send_message(
        message.chat.id,
        '🏭 Sex rejimiga oʻtdingiz.',
        reply_markup=get_main_menu('sex'),
    )


# --- ZAVSKLAD BOTDAN KASSAGA KIRIM QILISH (ADMIN TASDIQI BILAN) ---
@bot.message_handler(func=lambda message: message.text == '💰 Kassaga Kirim Qilish')
def zavsklad_cash_income_start(message):
  msg = bot.send_message(
      message.chat.id,
      "💵 Kassaga kirim qilinadigan **summani** kiriting (faqat raqam):",
      parse_mode='HTML',
      reply_markup=types.ReplyKeyboardRemove(),
  )
  bot.register_next_step_handler(msg, zavsklad_cash_income_get_amount)


def zavsklad_cash_income_get_amount(message):
  try:
    amount = float(message.text)
    user_steps[message.from_user.id] = {'cash_income_amount': amount}
    msg = bot.send_message(
        message.chat.id,
        "📝 Kirim bo'yicha **izoh yoki sababni** yozing (masalan: <i>'Plastik karta orqali toʻlov'</i>):",
        parse_mode='HTML',
    )
    bot.register_next_step_handler(msg, zavsklad_cash_income_finish)
  except:
    bot.send_message(
        message.chat.id,
        '❌ Xatolik! Faqat raqam kiriting.',
        reply_markup=get_main_menu('sex'),
    )


def zavsklad_cash_income_finish(message):
  uid = message.from_user.id
  reason = message.text if message.text else 'Izohsiz'
  data = user_steps.get(uid)
  if not data or 'cash_income_amount' not in data:
    bot.send_message(
        message.chat.id, '❌ Xatolik yuz berdi.', reply_markup=get_main_menu('sex')
    )
    return

  amount = data['cash_income_amount']
  conn = get_db_connection()
  u_res = conn.execute(
      'SELECT name FROM users WHERE tg_id = ?', (uid,)
  ).fetchone()
  staff_name = u_res['name'] if u_res else 'Zavsklad xodimi'
  date_str = datetime.now().strftime('%Y-%m-%d %H:%M')

  cursor = conn.cursor()
  cursor.execute(
      'INSERT INTO pending_incomes (staff_name, amount, reason, date) VALUES'
      ' (?, ?, ?, ?)',
      (staff_name, amount, reason, date_str),
  )
  pending_id = cursor.lastrowid
  conn.commit()
  conn.close()

  bot.send_message(
      message.chat.id,
      f"✅ Kirim so'rovi adminga yuborildi!\n💰 Summa: <b>{amount:,.0f} so'm</b>\n📝 Sabab: {reason}\n\n<i>Admin tasdiqlagach kassaga qo'shiladi.</i>",
      parse_mode='HTML',
      reply_markup=get_main_menu('sex'),
  )

  # Adminga tasdiqlash uchun tugma yuborish
  admin_markup = types.InlineKeyboardMarkup()
  admin_markup.row(
      types.InlineKeyboardButton(
          '✅ Tasdiqlash', callback_data=f'approve_income_{pending_id}'
      ),
      types.InlineKeyboardButton(
          '❌ Rad etish', callback_data=f'reject_income_{pending_id}'
      ),
  )
  try:
    bot.send_message(
        ADMIN_ID,
        f"🔔 <b>ZAVSKLADDAN KASSA KIRIMI SO'ROVI</b>\n\n👤 Xodim: {staff_name}\n💰 Summa: <b>{amount:,.0f} so'm</b>\n📝 Sabab: {reason}\n📅 Sana: {date_str}",
        parse_mode='HTML',
        reply_markup=admin_markup,
    )
  except Exception as e:
    print('Adminga xabar yuborish xatosi:', e)

  if uid in user_steps:
    del user_steps[uid]


@bot.callback_query_handler(
    func=lambda call: call.data.startswith(('approve_income_', 'reject_income_'))
)
def admin_income_approval_callback(call):
  parts = call.data.split('_')
  action = parts[0]
  pending_id = int(parts[2])

  conn = get_db_connection()
  cursor = conn.cursor()
  p_inc = cursor.execute(
      'SELECT * FROM pending_incomes WHERE id = ?', (pending_id,)
  ).fetchone()

  if not p_inc:
    bot.answer_callback_query(
        call.id, "Bu so'rov allaqachon ko'rib chiqilgan!"
    )
    conn.close()
    return

  if action == 'approve':
    # Kassaga (incomes table) qo'shamiz
    cursor.execute(
        'INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)',
        (f"Zavsklad kirimi ({p_inc['staff_name']} - {p_inc['reason']})", p_inc['amount'], p_inc['date']),
    )
    cursor.execute('DELETE FROM pending_incomes WHERE id = ?', (pending_id,))
    conn.commit()
    conn.close()

    bot.edit_message_text(
        f"✅ <b>Kirim tasdiqlandi!</b>\n💰 {p_inc['amount']:,.0f} so'm kassaga qo'shildi.",
        call.message.chat.id,
        call.message.message_id,
        parse_mode='HTML',
    )
    bot.answer_callback_query(call.id, 'Tasdiqlandi va kassaga qo\'shildi!')
  else:
    cursor.execute('DELETE FROM pending_incomes WHERE id = ?', (pending_id,))
    conn.commit()
    conn.close()

    bot.edit_message_text(
        '❌ <b>Kirim so\'rovi rad etildi.</b>',
        call.message.chat.id,
        call.message.message_id,
        parse_mode='HTML',
    )
    bot.answer_callback_query(call.id, 'So\'rov rad etildi.')


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
  markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
  msg = bot.send_message(
      message.chat.id,
      '🏭 Kirim qilinadigan mahsulotni tanlang:',
      reply_markup=markup,
  )
  bot.register_next_step_handler(msg, sex_income_choose_product)


def sex_income_choose_product(message):
  if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
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

  staff_res = cursor.execute(
      'SELECT name FROM users WHERE tg_id = ?', (uid,)
  ).fetchone()
  staff_name = staff_res['name'] if staff_res else 'Nomalum'
  bugun = datetime.now().strftime('%Y-%m-%d %H:%M')

  cursor.execute(
      'INSERT INTO product_incomes (product_name, qty, cost_price, date) VALUES'
      " (?, ?, 0, ?)",
      (items_text, sum(data['sex_cart'].values()), bugun),
  )
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
  if message.from_user.id != ADMIN_ID:
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
      'SELECT name, optom_price, chakana_price, stock FROM products WHERE id ='
      ' ?',
      (p_id,),
  ).fetchone()
  conn.close()
  if p:
    optom = p['optom_price'] if p['optom_price'] is not None else 0
    chakana = p['chakana_price'] if p['chakana_price'] is not None else 0
    stock = p['stock'] if p['stock'] is not None else 0
    text = (
        f"📦 <b>Mahsulot:</b> {p['name']}\n💰 Optom: {optom:,.0f} so'm\n🛍"
        f' Chakana: {chakana:,.0f} so\'m\n🔢 <b>Qoldiq:</b> {int(stock)}'
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
  prefix, action, p_id = call.data.split('_')
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
          if any(r in a['role'] for r in ['agent', 'sex'])
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
  conn.execute(
      'UPDATE orders SET status = ? WHERE id = ?', (new_status, order_id)
  )
  conn.commit()
  conn.close()
  bot.send_message(
      call.message.chat.id, f'✅ Status: <b>{new_status}</b>', parse_mode='HTML'
  )


@bot.callback_query_handler(func=lambda call: call.data.startswith('approve_'))
def approve_agent_cb(call):
  if call.data.startswith('approve_income_'):
    return
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


# --- WEB LOGIN & AUTHENTICATION (TO'G'rilandi) ---
@app.route('/login', methods=['GET', 'POST'])
def login():
  error = None
  if request.method == 'POST':
    username = request.form.get('username')
    password = request.form.get('password')

    conn = get_db_connection()
    user = conn.execute(
        "SELECT * FROM users WHERE name = ? OR role = 'admin'", (username,)
    ).fetchone()
    conn.close()

    if username == 'Admin' and password == 'admin2026':
      session['logged_in'] = True
      session['user_name'] = 'Admin'
      return redirect(url_for('operator_dashboard'))
    elif user and user['password'] == password:
      session['logged_in'] = True
      session['user_name'] = user['name']
      return redirect(url_for('operator_dashboard'))
    else:
      error = 'Login yoki parol noto\'g\'ri!'

  return render_template_string(LOGIN_TEMPLATE, error=error)


@app.route('/logout')
def logout():
  session.clear()
  return redirect(url_for('login'))


# --- WEB DASHBOARD & ROUTES ---
@app.route('/')
def operator_dashboard():
  if not session.get('logged_in'):
    return redirect(url_for('login'))

  start_date = request.args.get('start_date', '')
  end_date = request.args.get('end_date', '')

  conn = get_db_connection()
  cursor = conn.cursor()

  # Filterli yoki barcha buyurtmalar
  if start_date and end_date:
    orders_raw = cursor.execute(
        'SELECT * FROM orders WHERE date BETWEEN ? AND ? ORDER BY id DESC',
        (start_date + ' 00:00', end_date + ' 23:59'),
    ).fetchall()
  else:
    orders_raw = cursor.execute(
        'SELECT * FROM orders ORDER BY id DESC'
    ).fetchall()

  orders = []
  for o in orders_raw:
    ord_dict = dict(o)
    history = cursor.execute(
        'SELECT status, changed_at FROM order_status_history WHERE order_id = ?',
        (o['id'],),
    ).fetchall()
    ord_dict['history'] = [dict(h) for h in history]
    orders.append(ord_dict)

  products = cursor.execute(
      'SELECT * FROM products ORDER BY name ASC'
  ).fetchall()
  shops = cursor.execute('SELECT * FROM shops ORDER BY name ASC').fetchall()
  agents = cursor.execute(
      "SELECT * FROM users WHERE role LIKE '%agent%'"
  ).fetchall()

  # Kassa va Qarz hisobi
  total_incomes = (
      cursor.execute('SELECT SUM(amount) FROM incomes').fetchone()[0] or 0
  )
  total_expenses = (
      cursor.execute('SELECT SUM(amount) FROM expenses').fetchone()[0] or 0
  )
  kassa_balance = total_incomes - total_expenses
  total_debt = cursor.execute('SELECT SUM(debt) FROM shops').fetchone()[0] or 0

  daily_sum = sum(o['total_sum'] for o in orders if o['status'] != 'Bekor')

  # Kategoriyalar bo'yicha ma'lumot
  cat_data = cursor.execute(
      'SELECT category, SUM(stock * optom_price) as val FROM products GROUP BY'
      ' category'
  ).fetchall()
  cat_table_data = []
  colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899']
  for idx, c in enumerate(cat_data):
    cat_name = c['category'] or 'Boshqa'
    val = c['val'] or 0
    cat_table_data.append({
        'name': cat_name,
        'amount': val,
        'percent': 15,
        'color': colors[idx % len(colors)],
    })

  conn.close()

  return render_template_string(
      HTML_TEMPLATE,
      orders=orders,
      products=products,
      shops=shops,
      agents=agents,
      kassa_balance=kassa_balance,
      total_debt=total_debt,
      daily_sum=daily_sum,
      cat_table_data=cat_table_data,
      start_date=start_date,
      end_date=end_date,
  )


# --- WEB ORQALI ZAVSKLAD BOTIGA (SEX_GROUP_ID) MATN JO'NATISH VA BUYURTMA QO'SHISH ---
@app.route('/add_order', methods=['POST'])
def add_order():
  if not session.get('logged_in'):
    return redirect(url_for('login'))

  shop_name = request.form.get('shop_name')
  agent_name = request.form.get('agent_name', 'Admin (Web)')
  discount = float(request.form.get('discount', 0) or 0)
  comment = request.form.get('comment', '')

  product_names = request.form.getlist('product_name')
  quantities = request.form.getlist('qty')

  conn = get_db_connection()
  cursor = conn.cursor()

  total_sum = 0
  items_text = ''
  excel_cart_items = []

  for p_name, q_str in zip(product_names, quantities):
    if not p_name:
      continue
    try:
      qty = float(q_str)
    except:
      qty = 1

    prod = cursor.execute(
        'SELECT optom_price, stock, id FROM products WHERE name = ?',
        (p_name,),
    ).fetchone()
    if prod:
      price = prod['optom_price'] if prod['optom_price'] is not None else 0
      summa = price * qty
      total_sum += summa
      items_text += f'{p_name} - {qty}x = {summa:,.0f} so\'m\n'
      excel_cart_items.append({'name': p_name, 'qty': qty, 'price': price})
      cursor.execute(
          'UPDATE products SET stock = stock - ? WHERE id = ?',
          (qty, prod['id']),
      )

  final_sum = total_sum - discount
  bugun = datetime.now().strftime('%Y-%m-%d %H:%M')

  cursor.execute(
      'INSERT INTO orders (shop_name, agent_name, total_sum, items_text,'
      " status, date, price_type, discount, comment) VALUES (?, ?, ?, ?, 'Yangi',"
      ' ?, \'optom\', ?, ?)',
      (
          shop_name,
          agent_name,
          final_sum,
          items_text,
          bugun,
          discount,
          comment,
      ),
  )
  order_id = cursor.lastrowid

  # Do'kon qarzini oshiramiz
  shop_res = cursor.execute(
      'SELECT debt FROM shops WHERE name = ?', (shop_name,)
  ).fetchone()
  curr_debt = (
      shop_res['debt'] if shop_res and shop_res['debt'] is not None else 0
  )
  cursor.execute(
      'UPDATE shops SET debt = ? WHERE name = ?',
      (curr_debt + final_sum, shop_name),
  )

  conn.commit()
  conn.close()

  # Excel nakladnoy yaratish va zavsklad botiga (SEX_GROUP_ID) matn ko'rinishida jo'natish
  excel_file = create_excel_invoice(
      order_id, shop_name, agent_name, bugun, 'optom', excel_cart_items, comment
  )
  try:
    group_text = (
        f"📝 <b>WEB ORQALI YANGI BUYURTMA (#{order_id})</b>\n"
        f"🏪 <b>Do'kon:</b> {shop_name}\n"
        f"👤 <b>Agent:</b> {agent_name}\n"
        f"💬 <b>Izoh:</b> {comment}\n"
        f"💰 <b>Jami summa:</b> {final_sum:,.0f} so'm\n\n"
        f"<b>Mahsulotlar:</b>\n{items_text}"
    )
    bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')
    with open(excel_file, 'rb') as doc_group:
      bot.send_document(
          SEX_GROUP_ID, doc_group, caption=f'📄 Nakladnoy (#{order_id})'
      )
  except Exception as e:
    print('Zavsklad botiga yuborish xatosi:', e)

  if os.path.exists(excel_file):
    os.remove(excel_file)

  return redirect(url_for('operator_dashboard'))


# --- ZAVSKLAD UCHUN YIG'MA EXCEL CHOP ETISH ROUTE ---
@app.route('/print_nakladnoy')
def print_nakladnoy_web():
  if not session.get('logged_in'):
    return redirect(url_for('login'))

  ids_str = request.args.get('ids', '')
  if not ids_str:
    return 'Buyurtma ID lari tanlanmagan!', 400

  try:
    order_ids = [int(i.strip()) for i in ids_str.split(',') if i.strip()]
  except:
    return 'ID lar formati xato!', 400

  conn = get_db_connection()
  cursor = conn.cursor()

  aggregated = {}
  for o_id in order_ids:
    ord_row = cursor.execute(
        'SELECT items_text FROM orders WHERE id = ?', (o_id,)
    ).fetchone()
    if ord_row and ord_row['items_text']:
      lines = ord_row['items_text'].strip().split('\n')
      for line in lines:
        if ' - ' in line and 'x =' in line:
          try:
            parts = line.split(' - ')
            p_name = parts[0].strip()
            rest = parts[1].split('x =')[0].strip()
            qty = float(rest)
            aggregated[p_name] = aggregated.get(p_name, 0) + qty
          except:
            pass

  conn.close()

  if not aggregated:
    return 'Tanlangan buyurtmalardan mahsulotlar topilmadi!', 400

  excel_path = create_excel_consolidated_invoice(aggregated, ids_str)
  return send_file(excel_path, as_attachment=True)


@app.route('/print_nakladnoy/<int:order_id>')
def print_single_nakladnoy(order_id):
  if not session.get('logged_in'):
    return redirect(url_for('login'))

  conn = get_db_connection()
  o = cursor = conn.execute(
      'SELECT * FROM orders WHERE id = ?', (order_id,)
  ).fetchone()
  conn.close()
  if not o:
    return 'Buyurtma topilmadi', 404

  # Items parse
  cart_items = []
  if o['items_text']:
    for line in o['items_text'].strip().split('\n'):
      if ' - ' in line:
        try:
          parts = line.split(' - ')
          p_name = parts[0].strip()
          rest = parts[1].split('x =')[0].strip()
          qty = float(rest)
          # taxminiy narx
          cart_items.append({'name': p_name, 'qty': qty, 'price': 0})
        except:
          pass

  excel_file = create_excel_invoice(
      o['id'],
      o['shop_name'],
      o['agent_name'],
      o['date'],
      o['price_type'] or 'optom',
      cart_items,
      o['comment'],
  )
  return send_file(excel_file, as_attachment=True)


@app.route('/add_shop', methods=['POST'])
def add_shop():
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  name, phone, visit_days, region, landmark, inventory = (
      request.form['name'],
      request.form['phone'],
      request.form.get('visit_days', ''),
      request.form.get('region', ''),
      request.form.get('landmark', ''),
      request.form.get('inventory', ''),
  )
  conn = get_db_connection()
  try:
    conn.execute(
        'INSERT INTO shops (name, phone, debt, visit_days, region, landmark,'
        ' inventory) VALUES (?, ?, 0, ?, ?, ?, ?)',
        (name, phone, visit_days, region, landmark, inventory),
    )
    conn.commit()
  except:
    pass
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/update_shop/<int:shop_id>', methods=['POST'])
def update_shop(shop_id):
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  name = request.form['name']
  phone = request.form['phone']
  region = request.form.get('region', '')
  landmark = request.form.get('landmark', '')
  visit_days = request.form.get('visit_days', '')
  inventory = request.form.get('inventory', '')

  conn = get_db_connection()
  try:
    conn.execute(
        'UPDATE shops SET name = ?, phone = ?, region = ?, landmark = ?,'
        ' visit_days = ?, inventory = ? WHERE id = ?',
        (name, phone, region, landmark, visit_days, inventory, shop_id),
    )
    conn.commit()
  except Exception as e:
    print(e)
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/update_product/<int:prod_id>', methods=['POST'])
def update_product(prod_id):
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  name = request.form['name']
  category = request.form.get('category', 'Boshqa')
  stock = float(request.form.get('stock', 0))
  cost_price = float(request.form.get('cost_price', 0))
  optom_price = float(request.form.get('optom_price', 0))
  chakana_price = float(request.form.get('chakana_price', optom_price))

  conn = get_db_connection()
  try:
    conn.execute(
        'UPDATE products SET name = ?, category = ?, stock = ?, cost_price ='
        ' ?, optom_price = ?, chakana_price = ? WHERE id = ?',
        (
            name,
            category,
            stock,
            cost_price,
            optom_price,
            chakana_price,
            prod_id,
        ),
    )
    conn.commit()
  except Exception as e:
    print('Mahsulotni tahrirlash xatosi:', e)
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/adjust_stock_web', methods=['POST'])
def adjust_stock_web():
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  prod_id = request.form.get('prod_id')
  action = request.form.get('action')
  qty = float(request.form.get('qty', 0))

  conn = get_db_connection()
  if action == 'plus':
    conn.execute(
        'UPDATE products SET stock = stock + ? WHERE id = ?', (qty, prod_id)
    )
  elif action == 'minus':
    conn.execute(
        'UPDATE products SET stock = stock - ? WHERE id = ?', (qty, prod_id)
    )
  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/add_agent_web', methods=['POST'])
def add_agent_web():
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  name = request.form.get('name')
  phone = request.form.get('phone')
  tg_id = request.form.get('tg_id')
  role = request.form.get('role', 'agent')
  if tg_id:
    try:
      tg_id = int(tg_id)
    except:
      tg_id = None

  if name and tg_id:
    conn = get_db_connection()
    try:
      conn.execute(
          'INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, ?,'
          ' ?, ?)',
          (tg_id, name, phone or '', role),
      )
      conn.commit()
    except Exception as e:
      print('Agent qo\'shish xatosi:', e)
    conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/import_shops_excel', methods=['POST'])
def import_shops_excel():
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  if 'excel_file' not in request.files:
    return redirect(url_for('operator_dashboard'))
  file = request.files['excel_file']
  if file.filename == '':
    return redirect(url_for('operator_dashboard'))

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

  return redirect(url_for('operator_dashboard'))


@app.route('/import_products_excel', methods=['POST'])
def import_products_excel():
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  if 'excel_file' not in request.files:
    return redirect(url_for('operator_dashboard'))
  file = request.files['excel_file']
  if file.filename == '':
    return redirect(url_for('operator_dashboard'))

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

  return redirect(url_for('operator_dashboard'))


@app.route('/approve_agent/<int:tg_id>', methods=['POST'])
def web_approve_agent(tg_id):
  if not session.get('logged_in'):
    return redirect(url_for('login'))
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
  return redirect(url_for('operator_dashboard'))


@app.route('/delete_agent/<int:tg_id>', methods=['POST'])
def web_delete_agent(tg_id):
  if not session.get('logged_in'):
    return redirect(url_for('login'))
  conn = get_db_connection()
  conn.execute('DELETE FROM users WHERE tg_id = ?', (tg_id,))
  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


# --- LOGIN HTML SHABLONI ---
LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <title>Xoji Aka ERP — Kirish</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #0f172a; height: 100vh; display: flex; align-items: center; justify-content: center; font-family: 'Inter', sans-serif; }
        .login-card { background: #ffffff; padding: 40px; border-radius: 16px; width: 100%; max-width: 400px; box-shadow: 0 10px 25px rgba(0,0,0,0.2); }
        .brand-xa { font-family: 'Georgia', serif; font-weight: 900; font-size: 40px; color: #0f172a; font-style: italic; text-align: center; }
        .brand-line { height: 3px; background-color: #ef4444; width: 60px; margin: 8px auto 20px auto; border-radius: 2px; }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="brand-xa">XA</div>
        <div class="brand-line"></div>
        <h5 class="text-center mb-4 fw-bold text-secondary">Tizimga Kirish</h5>
        {% if error %}<div class="alert alert-danger py-2 small">{{ error }}</div>{% endif %}
        <form method="POST">
            <div class="mb-3">
                <label class="form-label small fw-bold">Login (Ism):</label>
                <input type="text" name="username" class="form-control" required placeholder="Admin">
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Parol:</label>
                <input type="password" name="password" class="form-control" required placeholder="Parol">
            </div>
            <button type="submit" class="btn btn-primary w-100 py-2 fw-bold">Kirish</button>
        </form>
    </div>
</body>
</html>
"""

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
        @media (max-width: 768px) {
        .d-flex { flex-direction: column !important; }
        .sidebar { width: 100% !important; min-height: auto !important; }
        .sidebar .nav { flex-direction: row; overflow-x: auto; padding: 5px; }
        .sidebar .nav-link { font-size: 10px; padding: 8px 10px; margin: 2px; }
        .sidebar .nav-link i { font-size: 16px; margin-bottom: 0; }
        .brand-logo-container { display: none; }
        .top-bar { flex-direction: column; gap: 10px; align-items: stretch !important; }
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
    </div>
    <div class="flex-grow-1">
        <div class="top-bar d-flex justify-content-between align-items-center">
            <div class="fw-bold text-dark fs-6"><i class="bi bi-shield-check text-primary me-2"></i>Boshqaruv Markazi ({{ session.get('user_name') }})</div>
            <div class="d-flex align-items-center gap-2">
                <a href="/logout" class="btn btn-sm btn-outline-danger fw-bold"><i class="bi bi-box-arrow-right me-1"></i>Chiqish</a>
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
                </div>
                <div class="tab-pane fade" id="tab-orders">
                    <div class="row g-4">
                        <div class="col-md-4">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold text-primary mb-3"><i class="bi bi-cart-plus me-2"></i>Yangi Buyurtma Kiritish</h5>
                                <form action="/add_order" method="POST">
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Do'konni tanlang:</label>
                                        <select name="shop_name" class="form-select" required>
                                            <option value="">-- Do'konni tanlang --</option>
                                            {% for s in shops %}<option value="{{ s['name'] }}">{{ s['name'] }} ({{ s['region'] }})</option>{% endfor %}
                                        </select>
                                    </div>
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Agent / Operator:</label>
                                        <input type="text" name="agent_name" class="form-control" value="{{ session.get('user_name', 'Admin') }}" required>
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
                                    <label class="form-label small fw-bold text-secondary mb-2"><i class="bi bi-basket me-1"></i>Mahsulot:</label>
                                    <div id="order-items-container">
                                        <div class="row g-2 mb-2 align-items-center">
                                            <div class="col-7">
                                                <select name="product_name" class="form-select form-select-sm" required>
                                                    <option value="">-- Tovar tanlang --</option>
                                                    {% for p in products %}<option value="{{ p['name'] }}">{{ p['name'] }} (Sklad: {{ p['stock'] }}) - {{ "{:,.0f}".format(p['optom_price']) }} so'm</option>{% endfor %}
                                                </select>
                                            </div>
                                            <div class="col-4"><input type="number" step="any" name="qty" class="form-control form-control-sm" placeholder="Miqdor" value="1" required></div>
                                        </div>
                                    </div>
                                    <button type="submit" class="btn btn-primary w-100 py-2 fw-bold shadow-sm mt-2">Buyurtmani Tasdiqlash</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-8">
                            <!-- ZAVSKLAD UCHIN YIG'MA EXCEL CHOP ETISH BLoki -->
                            <div class="card-glass p-4 mb-3">
                                <h6 class="fw-bold mb-2"><i class="bi bi-printer me-1"></i>Tanlangan zakazlarni zavsklad uchun yig'ma qilib chop etish</h6>
                                <form action="/print_nakladnoy" method="GET" target="_blank" class="d-flex gap-2 align-items-center">
                                    <input type="text" id="selectedIdsInput" name="ids" class="form-control form-control-sm" placeholder="Tanlangan ID lar (Masalan: 1,2,3)" required readonly>
                                    <button type="submit" class="btn btn-sm btn-dark text-nowrap">Yig'ma Excel Chop Etish</button>
                                </form>
                            </div>
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-list-check me-2"></i>Barcha Buyurtmalar</h5>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle" id="ordersTable">
                                        <thead>
                                            <tr>
                                                <th style="width: 30px;"><input type="checkbox" id="selectAllOrders" onclick="toggleSelectAllOrders(this)"></th>
                                                <th>ID / Vaqt</th>
                                                <th>Do'kon</th>
                                                <th>Tafsilot & Izoh</th>
                                                <th>Status</th>
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
                                                    <code style="font-size: 11px; white-space: pre-line;" class="text-dark d-block">{{ o['items_text'] }}</code>
                                                    {% if o['comment'] %}<small class="text-primary fw-bold d-block mt-1"><i class="bi bi-chat-left-text me-1"></i>Izoh: {{ o['comment'] }}</small>{% endif %}
                                                    <div class="fw-bold text-primary mt-1">Jami: {{ "{:,.0f}".format(o['total_sum']) }} so'm</div>
                                                </td>
                                                <td><span class="badge bg-secondary">{{ o['status'] }}</span></td>
                                                <td>
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
                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-boxes me-2"></i>Ombordagi Qoldiqlar</h5>
                        <div class="table-responsive">
                            <table class="table table-custom table-hover align-middle">
                                <thead><tr><th>Mahsulot</th><th>Qoldiq</th><th>Optom narx</th></tr></thead>
                                <tbody>
                                    {% for p in products %}
                                    <tr>
                                        <td><b>{{ p['name'] }}</b><br><small class="text-muted">{{ p['category'] }}</small></td>
                                        <td><span class="badge bg-success fs-6">{{ p['stock'] }}</span></td>
                                        <td>{{ "{:,.0f}".format(p['optom_price']) }}</td>
                                    </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
                <div class="tab-pane fade" id="tab-clients"><div class="card-glass p-4"><h5>Do'konlar ro'yxati</h5></div></div>
                <div class="tab-pane fade" id="tab-agents"><div class="card-glass p-4"><h5>Agentlar ro'yxati</h5></div></div>
                <div class="tab-pane fade" id="tab-kassa"><div class="card-glass p-4"><h5>Kassa ma'lumotlari</h5></div></div>
                <div class="tab-pane fade" id="tab-reports"><div class="card-glass p-4"><h5>Hisobotlar</h5></div></div>
            </div>
        </div>
    </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    function toggleSelectAllOrders(source) {
        checkboxes = document.querySelectorAll('.order-checkbox');
        checkboxes.forEach(cb => cb.checked = source.checked);
        updateSelectedIds();
    }
    function updateSelectedIds() {
        const checked = document.querySelectorAll('.order-checkbox:checked');
        const ids = Array.from(checked).map(cb => cb.value);
        document.getElementById('selectedIdsInput').value = ids.join(',');
    }
</script>
</body>
</html>
"""

def run_bot():
    while True:
        try:
            bot.remove_webhook()
            bot.infinity_polling(timeout=60, long_polling_timeout=60)
        except Exception as e:
            print(f"Bot polling xatosi: {e}")


if __name__ == '__main__':
    init_web_db()

    bot_thread = threading.Thread(target=run_bot)
    bot_thread.daemon = True
    bot_thread.start()

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
