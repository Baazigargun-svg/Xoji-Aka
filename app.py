import sqlite3
import telebot
import os
import threading
from datetime import datetime
from flask import Flask, redirect, render_template_string, request, url_for, send_file, session
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BOT_TOKEN = "8573337094:AAHH1kNQnNrJyMfG6d5z6IO8lgkS1UdurR8"
ADMIN_ID = 6851851908  
SEX_GROUP_ID = -6851851908 

bot = telebot.TeleBot(BOT_TOKEN)
app = Flask(__name__)
app.secret_key = 'xoji_aka_maxfiy_kalit_2026'
DB_NAME = 'xoji_aka_factory.db'
user_steps = {}

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (tg_id INTEGER PRIMARY KEY, name TEXT, phone TEXT, role TEXT DEFAULT 'pending')''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS expenses 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, reason TEXT, amount REAL, date TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS incomes 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, source TEXT, amount REAL, date TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS product_incomes 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT, qty REAL, cost_price REAL, date TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, shop_name TEXT, agent_name TEXT, items_text TEXT, total_sum REAL, discount REAL DEFAULT 0, price_type TEXT, status TEXT DEFAULT 'Yangi', date TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS order_status_history 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER, status TEXT, changed_at TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS products 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, category TEXT DEFAULT 'Boshqa', stock REAL DEFAULT 0, cost_price REAL DEFAULT 0, optom_price REAL DEFAULT 0, chakana_price REAL DEFAULT 0)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS shops 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, phone TEXT, debt REAL DEFAULT 0, visit_days TEXT DEFAULT '')''')

    try:
        cursor.execute("ALTER TABLE products ADD COLUMN chakana_price REAL DEFAULT 0")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE products ADD COLUMN category TEXT DEFAULT 'Boshqa'")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE products ADD COLUMN cost_price REAL DEFAULT 0")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE products ADD COLUMN optom_price REAL DEFAULT 0")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE products ADD COLUMN stock REAL DEFAULT 0")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE shops ADD COLUMN visit_days TEXT DEFAULT ''")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE orders ADD COLUMN discount REAL DEFAULT 0")
    except:
        pass

    conn.commit()
    conn.close()

def auto_insert_products():
    products_list = [
        ("Pelmen 300 gr", 14000), ("Pelmen 500 gr", 24000), ("Pelmen frikadel 500 gr", 28000),
        ("Teftel 300 gr", 23000), ("Golubtsi 500 gr", 25000), ("Tok do'lma 300 gr", 25000),
        ("Karam do'lma 300 gr", 25000), ("Somsa kesilgani (20 dona)", 17000), ("O'rama hamir", 17000),
        ("Osh masalliq 0.5 kg", 13000), ("Osh masalliq 1 kg", 15000), ("Farsh mol go'shtidan", 19000),
        ("Frikadel 300 gr", 20000), ("Frikadel 500 gr", 30000), ("KFC 500 gr", 30000),
        ("Mantini hamiri 25 dona", 13000), ("Kotlet 5 dona", 25000)
    ]
    conn = get_db_connection()
    cursor = conn.cursor()
    for name, price in products_list:
        try:
            cursor.execute("INSERT OR IGNORE INTO products (name, optom_price, chakana_price, stock, category) VALUES (?, ?, ?, 0, 'Yarim tayyor')", (name, price, price))
        except:
            continue
    conn.commit()
    conn.close()

# --- TELEGRAM BOT QISMI ---
def create_excel_invoice(order_id, shop_name, agent_name, date_str, price_type, cart_items):
    wb = Workbook()
    ws = wb.active
    ws.title = f"Nakladnoy_{order_id}"
    ws.sheet_view.showGridLines = True
    
    title_font = Font(name='Arial', size=16, bold=True)
    header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
    bold_font = Font(name='Arial', size=11, bold=True)
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    total_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'), top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0'))
    
    ws.merge_cells('A1:E1')
    ws['A1'] = "XOJI AKA FACTORY — NAKLADNOY"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center')
    
    ws['A3'] = f"Buyurtma ID: #{order_id}"
    ws['A3'].font = bold_font
    ws['D3'] = f"Sana: {date_str}"
    ws['A4'] = f"Do'kon (Klient): {shop_name}"
    ws['D4'] = f"Narx turi: {str(price_type).upper()}"
    ws['A5'] = f"Agent: {agent_name}"
    
    headers = ["№", "Mahsulot nomi", "Miqdori", "Narxi (so'm)", "Jami summa"]
    for col_num, header_title in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col_num, value=header_title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        
    row_num = 8
    total_sum = 0
    for idx, item in enumerate(cart_items, 1):
        ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(horizontal='center')
        ws.cell(row=row_num, column=2, value=item['name']).alignment = Alignment(horizontal='left')
        ws.cell(row=row_num, column=3, value=item['qty']).alignment = Alignment(horizontal='right')
        ws.cell(row=row_num, column=4, value=item['price']).alignment = Alignment(horizontal='right')
        summa = item['qty'] * item['price']
        total_sum += summa
        ws.cell(row=row_num, column=5, value=summa).alignment = Alignment(horizontal='right')
        ws.cell(row=row_num, column=4).number_format = '#,##0'
        ws.cell(row=row_num, column=5).number_format = '#,##0'
        for col in range(1, 6): ws.cell(row=row_num, column=col).border = thin_border
        row_num += 1
        
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
    ws.cell(row=row_num, column=1, value="JAMI TO'LOV:").alignment = Alignment(horizontal='right')
    ws.cell(row=row_num, column=1).font = bold_font
    total_val = ws.cell(row=row_num, column=5, value=total_sum)
    total_val.font = bold_font
    total_val.number_format = '#,##0'
    for col in range(1, 6):
        ws.cell(row=row_num, column=col).fill = total_fill
        ws.cell(row=row_num, column=col).border = thin_border
        
    row_num += 2
    ws.cell(row=row_num, column=2, value="Qabul qildim: ____")
    ws.cell(row=row_num, column=4, value="Topshirdim: ____")

    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(row=row, column=col).value or '')) for row in range(1, ws.max_row + 1))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    file_name = f"Nakladnoy_{order_id}.xlsx"
    wb.save(file_name)
    return file_name

def get_main_menu(role):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    if role == "admin":
        markup.row(types.KeyboardButton("📦 Sklad & Mahsulotlar"), types.KeyboardButton("📊 Kunlik Hisobot"))
        markup.row(types.KeyboardButton("👥 Agentlar boshqaruvi"), types.KeyboardButton("🏪 AKB (Do'konlar & Qarz)"))
    elif role == "agent":
        markup.row(types.KeyboardButton("🛒 Yangi Buyurtma Urish"), types.KeyboardButton("🏪 Do'kon qo'shish (AKB)"))
        markup.row(types.KeyboardButton("💰 Qarz/To'lov yozish"), types.KeyboardButton("📜 Mening Buyurtmalarim"))
    else:
        markup.add(types.KeyboardButton("📝 Ro'yxatdan o'tish"))
    return markup

@bot.message_handler(commands=['start'])
def start_command(message):
    tg_id = message.from_user.id
    if tg_id == ADMIN_ID:
        conn = get_db_connection()
        conn.execute("INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, 'Admin', '', 'admin')", (tg_id,))
        conn.commit()
        conn.close()
        bot.send_message(message.chat.id, "Xoji aka, xush kelibsiz! Boshqaruv paneli tayyor.", reply_markup=get_main_menu("admin"))
        return
    conn = get_db_connection()
    user = conn.execute("SELECT role, name FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    conn.close()
    if user:
        role, name = user['role'], user['name']
        if role == "pending":
            bot.send_message(message.chat.id, f"Salom {name}. So'rovingiz admin tasdig'ini kutyapti.")
        else:
            bot.send_message(message.chat.id, f"Salom {name}! Ishni boshlashimiz mumkin.", reply_markup=get_main_menu(role))
    else:
        bot.send_message(message.chat.id, "Assalomu alaykum! Tizimga xush kelibsiz. Davom etish uchun ro'yxatdan o'ting.", reply_markup=get_main_menu("guest"))

@bot.message_handler(func=lambda message: message.text == "📦 Sklad & Mahsulotlar")
def admin_sklad_menu(message):
    if message.from_user.id != ADMIN_ID: return
    conn = get_db_connection()
    prods = conn.execute("SELECT id, name, stock FROM products").fetchall()
    conn.close()

    inline_kb = types.InlineKeyboardMarkup(row_width=1)
    if not prods:
        bot.send_message(message.chat.id, "📦 Omborxonada mahsulotlar qolmagan.")
        return
    for p in prods:
        stock_val = p['stock'] if p['stock'] is not None else 0
        inline_kb.add(types.InlineKeyboardButton(f"🔹 {p['name']} ({int(stock_val)} kg/dona)", callback_data=f"adm_prod_{p['id']}"))
    inline_kb.add(types.InlineKeyboardButton("🆕 ✨ YANGI MAHSULOT QO'SHISH", callback_data="adm_create_product"))
    bot.send_message(message.chat.id, "<b>📦 Sklad nazorati:</b>", parse_mode="HTML", reply_markup=inline_kb)

@bot.callback_query_handler(func=lambda call: call.data.startswith("adm_prod_"))
def admin_product_detail(call):
    p_id = int(call.data.split("_")[-1])
    conn = get_db_connection()
    p = conn.execute("SELECT name, optom_price, chakana_price, stock FROM products WHERE id = ?", (p_id,)).fetchone()
    conn.close()
    if p:
        optom = p['optom_price'] if p['optom_price'] is not None else 0
        chakana = p['chakana_price'] if p['chakana_price'] is not None else 0
        stock = p['stock'] if p['stock'] is not None else 0
        text = f"📦 <b>Mahsulot:</b> {p['name']}\n💰 Optom: {optom:,.0f} so'm\n🛍 Chakana: {chakana:,.0f} so'm\n🔢 <b>Qoldiq:</b> {int(stock)}"
        markup = types.InlineKeyboardMarkup()
        markup.row(types.InlineKeyboardButton("➕ Qoldiq Qo'shish", callback_data=f"stk_plus_{p_id}"), types.InlineKeyboardButton("➖ Qoldiq Ayirish", callback_data=f"stk_minus_{p_id}"))
        markup.row(types.InlineKeyboardButton("💵 Optom Narx", callback_data=f"prc_optom_{p_id}"), types.InlineKeyboardButton("💵 Chakana Narx", callback_data=f"prc_chakana_{p_id}"))
        markup.row(types.InlineKeyboardButton("🗑 O'chirish", callback_data=f"stk_del_{p_id}"))
        bot.edit_message_text(text, call.message.chat.id, call.message.message_id, parse_mode="HTML", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith(("stk_", "prc_")))
def admin_stock_price_actions(call):
    prefix, action, p_id = call.data.split("_")
    p_id = int(p_id)
    conn = get_db_connection()
    res = conn.execute("SELECT name FROM products WHERE id = ?", (p_id,)).fetchone()
    p_name = res['name'] if res else "Nomalum"
    conn.close()

    if action == "del":
        conn = get_db_connection()
        conn.execute("DELETE FROM products WHERE id = ?", (p_id,))
        conn.commit()
        conn.close()
        bot.send_message(call.message.chat.id, f"🗑 <b>{p_name}</b> o'chirildi.", parse_mode="HTML")
        return

    msg = bot.send_message(call.message.chat.id, f"🔢 <b>{p_name}</b> uchun qiymat kiriting:", parse_mode="HTML")
    bot.register_next_step_handler(msg, save_product_edits, prefix, action, p_id)

def save_product_edits(message, prefix, action, p_id):
    try:
        val = float(message.text)
        conn = get_db_connection()
        cursor = conn.cursor()
        if action == "plus": cursor.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (val, p_id))
        elif action == "minus": cursor.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (val, p_id))
        elif action == "optom": cursor.execute("UPDATE products SET optom_price = ? WHERE id = ?", (val, p_id))
        elif action == "chakana": cursor.execute("UPDATE products SET chakana_price = ? WHERE id = ?", (val, p_id))
        conn.commit()
        conn.close()
        bot.send_message(message.chat.id, "✅ O'zgarish saqlandi!")
    except: 
        bot.send_message(message.chat.id, "❌ Faqat raqam kiriting.")

@bot.callback_query_handler(func=lambda call: call.data == "adm_create_product")
def admin_create_product_start(call):
    msg = bot.send_message(call.message.chat.id, "📝 Yangi mahsulot NOMINI kiriting:")
    bot.register_next_step_handler(msg, process_new_p_name)

def process_new_p_name(message):
    name = message.text
    msg = bot.send_message(message.chat.id, f"💰 '{name}' uchun Optom narxini kiriting:")
    bot.register_next_step_handler(msg, process_new_p_optom, name)

def process_new_p_optom(message, name):
    try:
        optom = float(message.text)
        msg = bot.send_message(message.chat.id, f"🛍 '{name}' uchun Chakana narxini kiriting:")
        bot.register_next_step_handler(msg, process_new_p_final, name, optom)
    except: 
        bot.send_message(message.chat.id, "Xato! Faqat raqam kiriting.")

def process_new_p_final(message, name, optom):
    try:
        chakana = float(message.text)
        conn = get_db_connection()
        conn.execute("INSERT INTO products (name, optom_price, chakana_price, stock) VALUES (?, ?, ?, 0)", (name, optom, chakana))
        conn.commit()
        conn.close()
        bot.send_message(message.chat.id, f"✅ Yangi mahsulot qo'shildi: {name}")
    except: 
        bot.send_message(message.chat.id, "Xato! Bu nom allaqachon mavjud.")

@bot.message_handler(func=lambda message: message.text in ["📊 Kunlik Hisobot", "🏪 AKB (Do'konlar & Qarz)", "👥 Agentlar boshqaruvi"])
def admin_other_sections(message):
    if message.from_user.id != ADMIN_ID: return
    conn = get_db_connection()
    
    if message.text == "📊 Kunlik Hisobot":
        bugun = datetime.now().strftime("%Y-%m-%d")
        orders = conn.execute("SELECT id, shop_name, total_sum, status FROM orders WHERE date LIKE ?", (f"{bugun}%",)).fetchall()
        text = f"📊 <b>Bugungi buyurtmalar ({bugun}):</b>\n\n"
        inline_kb = types.InlineKeyboardMarkup()
        if not orders: text += "Hali buyurtma yo'q."
        for o in orders:
            t_sum = o['total_sum'] if o['total_sum'] is not None else 0
            text += f"🆔 #{o['id']} | {o['shop_name']} | {t_sum:,.0f} so'm | <b>{o['status']}</b>\n"
            inline_kb.add(types.InlineKeyboardButton(f"⚙️ #{o['id']} Statusi", callback_data=f"mng_ord_{o['id']}"))
        bot.send_message(message.chat.id, text, parse_mode="HTML", reply_markup=inline_kb)
        
    elif message.text == "🏪 AKB (Do'konlar & Qarz)":
        shops = conn.execute("SELECT name, phone, debt FROM shops").fetchall()
        text = "<b>🏪 Do'konlar qarzlari:</b>\n\n"
        for s in shops: 
            debt_val = s['debt'] if s['debt'] is not None else 0
            text += f"🏢 {s['name']} ({s['phone']}) — Qarz: {debt_val:,.0f} so'm\n"
        inline_kb = types.InlineKeyboardMarkup().add(types.InlineKeyboardButton("➕ Do'kon Qo'shish", callback_data="admin_add_shop"))
        bot.send_message(message.chat.id, text, parse_mode="HTML", reply_markup=inline_kb)
        
    elif message.text == "👥 Agentlar boshqaruvi":
        agents = conn.execute("SELECT name, phone, role, tg_id FROM users WHERE role != 'admin'").fetchall()
        inline_kb = types.InlineKeyboardMarkup()
        text = "<b>👥 Agentlar:</b>\n\n"
        for a in agents:
            status = "✅ Faol" if a['role'] == 'agent' else "⏳ Kutilmoqda"
            text += f"👤 {a['name']} ({a['phone']}) - {status}\n"
            if a['role'] == 'pending': inline_kb.add(types.InlineKeyboardButton(f"👍 {a['name']}ni tasdiqlash", callback_data=f"approve_{a['tg_id']}"))
        bot.send_message(message.chat.id, text, parse_mode="HTML", reply_markup=inline_kb)
    conn.close()

@bot.callback_query_handler(func=lambda call: call.data.startswith("mng_ord_"))
def manage_order_status_menu(call):
    order_id = int(call.data.split("_")[-1])
    markup = types.InlineKeyboardMarkup()
    markup.row(types.InlineKeyboardButton("🔄 Yangi", callback_data=f"st_yangi_{order_id}"), types.InlineKeyboardButton("🚚 Otgruzka", callback_data=f"st_otgruzka_{order_id}"))
    markup.row(types.InlineKeyboardButton("✅ Yetkazildi", callback_data=f"st_done_{order_id}"), types.InlineKeyboardButton("❌ Bekor", callback_data=f"st_otmen_{order_id}"))
    bot.send_message(call.message.chat.id, f"🆔 #{order_id} - Statusni tanlang:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith("st_"))
def change_status_logic(call):
    _, mode, order_id = call.data.split("_")
    order_id = int(order_id)
    status_map = {"yangi": "Yangi", "otgruzka": "Otgruzka", "done": "Yetkazildi", "otmen": "Bekor"}
    new_status = status_map[mode]
    conn = get_db_connection()
    conn.execute("UPDATE orders SET status = ? WHERE id = ?", (new_status, order_id))
    conn.commit()
    conn.close()
    bot.send_message(call.message.chat.id, f"✅ Status: <b>{new_status}</b>", parse_mode="HTML")

@bot.callback_query_handler(func=lambda call: call.data == "admin_add_shop")
def admin_inline_clicks(call):
    msg = bot.send_message(call.message.chat.id, "🏢 Do'kon NOMINI kiriting:")
    bot.register_next_step_handler(msg, process_shop_name_step)

@bot.message_handler(func=lambda message: message.text == "🏪 Do'kon qo'shish (AKB)")
def agent_add_shop_menu(message):
    msg = bot.send_message(message.chat.id, "🏢 Do'kon NOMINI kiriting:")
    bot.register_next_step_handler(msg, process_shop_name_step)

def process_shop_name_step(message):
    shop_name = message.text
    msg = bot.send_message(message.chat.id, f"📞 '{shop_name}' uchun TELEFON RAQAM:")
    bot.register_next_step_handler(msg, process_shop_phone_final, shop_name)

def process_shop_phone_final(message, shop_name):
    phone = message.text
    conn = get_db_connection()
    try:
        conn.execute("INSERT INTO shops (name, phone, debt) VALUES (?, ?, 0)", (shop_name, phone))
        conn.commit()
        bot.send_message(message.chat.id, f"✅ Do'kon saqlandi: {shop_name}")
    except: 
        bot.send_message(message.chat.id, "❌ Bu do'kon allaqachon mavjud.")
    finally: 
        conn.close()

@bot.message_handler(func=lambda message: message.text == "🛒 Yangi Buyurtma Urish")
def start_order(message):
    conn = get_db_connection()
    shops = conn.execute("SELECT name FROM shops").fetchall()
    conn.close()
    if not shops:
        bot.send_message(message.chat.id, "❌ Do'konlar yo'q.")
        return
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for s in shops: markup.add(types.KeyboardButton(s['name']))
    msg = bot.send_message(message.chat.id, "Do'konni tanlang:", reply_markup=markup)
    bot.register_next_step_handler(msg, choose_price_type)

def choose_price_type(message):
    user_steps[message.from_user.id] = {'shop_name': message.text, 'cart': {}, 'price_type': None}
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True).row(types.KeyboardButton("💰 Ulgurji (Optom)"), types.KeyboardButton("🛍 Chakana"))
    msg = bot.send_message(message.chat.id, "Narx turini tanlang:", reply_markup=markup)
    bot.register_next_step_handler(msg, show_products_to_agent)

def show_products_to_agent(message):
    p_type = "optom" if "Ulgurji" in message.text else "chakana"
    user_steps[message.from_user.id]['price_type'] = p_type
    send_product_list_menu(message)

def send_product_list_menu(message):
    conn = get_db_connection()
    prods = conn.execute("SELECT name FROM products").fetchall()
    conn.close()
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for p in prods: markup.add(types.KeyboardButton(p['name']))
    markup.add(types.KeyboardButton("✅ Buyurtmani yakunlash"))
    msg = bot.send_message(message.chat.id, "Mahsulotni tanlang:", reply_markup=markup)
    bot.register_next_step_handler(msg, ask_quantity)

def ask_quantity(message):
    if message.text == "✅ Buyurtmani yakunlash":
        finish_order(message)
        return
    user_steps[message.from_user.id]['current_product'] = message.text
    msg = bot.send_message(message.chat.id, f"🔢 Miqdorini kiriting:", reply_markup=types.ReplyKeyboardRemove())
    bot.register_next_step_handler(msg, add_to_cart)

def add_to_cart(message):
    uid = message.from_user.id
    try:
        qty = float(message.text)
        p_name = user_steps[uid]['current_product']
        user_steps[uid]['cart'][p_name] = qty
        bot.send_message(message.chat.id, f"📥 Qo'shildi: {p_name} - {qty}")
        send_product_list_menu(message)
    except:
        bot.send_message(message.chat.id, "❌ Faqat raqam kiriting.")
        send_product_list_menu(message)

def finish_order(message):
    uid = message.from_user.id
    data = user_steps.get(uid)
    if not data or not data['cart']:
        bot.send_message(message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu("agent"))
        return
        
    conn = get_db_connection()
    cursor = conn.cursor()
    total_sum, items_text, p_type = 0, "", data['price_type']
    excel_cart_items = []
    
    for p_name, qty in data['cart'].items():
        prod = cursor.execute("SELECT optom_price, chakana_price, stock, id FROM products WHERE name = ?", (p_name,)).fetchone()
        optom_p = prod['optom_price'] if prod['optom_price'] is not None else 0
        chakana_p = prod['chakana_price'] if prod['chakana_price'] is not None else 0
        stock_p = prod['stock'] if prod['stock'] is not None else 0
        
        price = optom_p if p_type == "optom" else chakana_p
        summa = price * qty
        total_sum += summa
        items_text += f"{p_name} - {qty}x = {summa:,.0f} so'm\n"
        cursor.execute("UPDATE products SET stock = ? WHERE id = ?", (stock_p - qty, prod['id']))
        excel_cart_items.append({'name': p_name, 'qty': qty, 'price': price})
        
    shop_res = cursor.execute("SELECT debt FROM shops WHERE name = ?", (data['shop_name'],)).fetchone()
    current_debt = shop_res['debt'] if shop_res and shop_res['debt'] is not None else 0
    
    cursor.execute("UPDATE shops SET debt = ? WHERE name = ?", (current_debt + total_sum, data['shop_name']))
    agent_res = cursor.execute("SELECT name FROM users WHERE tg_id = ?", (uid,)).fetchone()
    agent_name = agent_res['name'] if agent_res else "Nomalum"
    bugun = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    cursor.execute("INSERT INTO orders (shop_name, agent_name, total_sum, items_text, price_type, status, date) VALUES (?, ?, ?, ?, ?, 'Yangi', ?)", (data['shop_name'], agent_name, total_sum, items_text, p_type, bugun))
    order_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    excel_file = create_excel_invoice(order_id, data['shop_name'], agent_name, bugun, p_type, excel_cart_items)
    
    bot.send_message(message.chat.id, f"✅ Buyurtma qabul qilindi! Jami: {total_sum:,.0f} so'm", reply_markup=get_main_menu("agent"))
    bot.send_message(ADMIN_ID, f"🔔 YANGI BUYURTMA:\n\nDo'kon: {data['shop_name']}\nSumma: {total_sum:,.0f} so'm")
    
    with open(excel_file, 'rb') as doc:
        bot.send_document(ADMIN_ID, doc, caption=f"📄 Nakladnoy (#{order_id})")
    if os.path.exists(excel_file): os.remove(excel_file)
    
    group_text = f"📝 <b>YANGI BUYURTMA</b>\n🏪 Do'kon: {data['shop_name']}\n\n{items_text}"
    try: bot.send_message(SEX_GROUP_ID, group_text, parse_mode="HTML")
    except: pass
    
    if uid in user_steps: del user_steps[uid]

@bot.callback_query_handler(func=lambda call: call.data.startswith("approve_"))
def approve_agent_cb(call):
    agent_id = int(call.data.split("_")[-1])
    conn = get_db_connection()
    conn.execute("UPDATE users SET role = 'agent' WHERE tg_id = ?", (agent_id,))
    conn.commit()
    conn.close()
    bot.answer_callback_query(call.id, "Agent tasdiqlandi!")
    try: bot.send_message(agent_id, "🎉 Sizning so'rovingiz tasdiqlandi! /start bosing", reply_markup=get_main_menu("agent"))
    except: pass

@bot.message_handler(func=lambda message: message.text == "📝 Ro'yxatdan o'tish")
def register_start(message):
    msg = bot.send_message(message.chat.id, "Ismingizni kiriting:")
    bot.register_next_step_handler(msg, process_register_name)

def process_register_name(message):
    name = message.text
    msg = bot.send_message(message.chat.id, "Telefon raqamingizni kiriting:")
    bot.register_next_step_handler(msg, lambda m: save_pending_user(m, name))

def save_pending_user(message, name):
    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, ?, ?, 'pending')", (message.from_user.id, name, message.text))
    conn.commit()
    conn.close()
    bot.send_message(message.chat.id, "⏳ Admin tasdig'i kutilmoqda.")
    bot.send_message(ADMIN_ID, f"🔔 Yangi ro'yxatdan o'tgan: {name} ({message.text})")

@bot.message_handler(func=lambda message: message.text == "📜 Mening Buyurtmalarim")
def my_orders(message):
    conn = get_db_connection()
    res = conn.execute("SELECT name FROM users WHERE tg_id = ?", (message.from_user.id,)).fetchone()
    agent_name = res['name'] if res else ""
    orders = conn.execute("SELECT shop_name, total_sum, status, date FROM orders WHERE agent_name = ? ORDER BY id DESC LIMIT 5", (agent_name,)).fetchall()
    conn.close()
    text = "<b>📜 Oxirgi buyurtmalar:</b>\n\n"
    for o in orders: 
        t_sum = o['total_sum'] if o['total_sum'] is not None else 0
        text += f"🏪 {o['shop_name']} | {t_sum:,.0f} so'm | {o['status']} | {o['date']}\n\n"
    bot.send_message(message.chat.id, text, parse_mode="HTML")

@bot.message_handler(func=lambda message: message.text == "💰 Qarz/To'lov yozish")
def pay_debt_start(message):
    conn = get_db_connection()
    shops = conn.execute("SELECT name FROM shops").fetchall()
    conn.close()
    if not shops:
        bot.send_message(message.chat.id, "❌ Do'konlar yo'q.")
        return
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for s in shops: markup.add(types.KeyboardButton(s['name']))
    msg = bot.send_message(message.chat.id, "Do'konni tanlang:", reply_markup=markup)
    bot.register_next_step_handler(msg, ask_payment_amount)

def ask_payment_amount(message):
    shop_name = message.text
    msg = bot.send_message(message.chat.id, f"'{shop_name}' qancha to'lov qildi (summa):", reply_markup=types.ReplyKeyboardRemove())
    bot.register_next_step_handler(msg, process_payment, shop_name)

def process_payment(message, shop_name):
    try:
        amount = float(message.text)
        today = datetime.now().strftime("%Y-%m-%d %H:%M")
        conn = get_db_connection()
        cursor = conn.cursor()
        
        s_res = cursor.execute("SELECT debt FROM shops WHERE name = ?", (shop_name,)).fetchone()
        curr_debt = s_res['debt'] if s_res and s_res['debt'] is not None else 0
        
        cursor.execute("UPDATE shops SET debt = ? WHERE name = ?", (curr_debt - amount, shop_name))
        cursor.execute("INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)", (f"Qarz to'lovi ({shop_name})", amount, today))
        conn.commit()
        conn.close()
        bot.send_message(message.chat.id, f"✅ To'lov yozildi: {amount:,.0f} so'm chegirildi.", reply_markup=get_main_menu("agent"))
    except: 
        bot.send_message(message.chat.id, "❌ Faqat raqam kiriting.", reply_markup=get_main_menu("agent"))


# --- FLASK VEB INTERFEYS QISMI ---
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
        .sidebar .nav-link.active { background: var(--primary-color); color: #ffffff; box-shadow: 0 4px 15px rgba(59, 130, 246, 0.4); font-weight: 600; }
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
    </style>
</head>
<body>
<div class="d-flex">
    <div class="sidebar d-flex flex-column flex-shrink-0 nav nav-pills" role="tablist">
        <div class="brand-logo-container">
            <div class="brand-xa">XA</div>
            <div class="brand-line"></div>
            <div class="brand-name">Xoji Aka</div>
        </div>
        <button class="nav-link active" data-bs-toggle="pill" data-bs-target="#tab-dashboard" type="button"><i class="bi bi-grid-fill"></i>Bosh Panel</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-orders" type="button"><i class="bi bi-cart-fill"></i>Buyurtmalar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-inventory" type="button"><i class="bi bi-boxes"></i>Sklad</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-clients" type="button"><i class="bi bi-shop"></i>Do'konlar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-kassa" type="button"><i class="bi bi-wallet2"></i>Kassa</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-reports" type="button"><i class="bi bi-file-earmark-bar-graph"></i>Hisobot</button>
    </div>

    <div class="flex-grow-1">
        <div class="top-bar d-flex justify-content-between align-items-center">
            <div class="fw-bold text-dark fs-6"><i class="bi bi-shield-check text-primary me-2"></i>Boshqaruv Markazi</div>
            <div class="d-flex align-items-center gap-3">
                <a href="/logout" class="btn btn-sm btn-outline-danger fw-bold"><i class="bi bi-box-arrow-right me-1"></i>Chiqish</a>
                <a href="/export_excel" class="btn btn-sm btn-success fw-bold"><i class="bi bi-file-earmark-excel me-1"></i>Excelga Yuklab Olish</a>
                <form method="GET" action="/" class="d-flex align-items-center gap-2 m-0">
                    <span class="text-muted small">Sana filtri:</span>
                    <input type="date" name="filter_date" value="{{ filter_date }}" class="form-control form-control-sm" style="width: 140px;" onchange="this.form.submit()">
                    <a href="/" class="btn btn-sm btn-light border">Barchasi</a>
                </form>
            </div>
        </div>

        <div class="p-4">
            <div class="tab-content">
                <div class="tab-pane fade show active" id="tab-dashboard">
                    <div class="row g-4 mb-4">
                        <div class="col-md-4">
                            <div class="stat-box stat-blue shadow-sm">
                                <div class="small text-white-50">Tanlangan Sana Tushumi</div>
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
                    <!-- Buyurtmalar paneli -->
                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3">Buyurtmalar Ro'yxati</h5>
                        <table class="table table-custom">
                            <thead><tr><th>ID</th><th>Do'kon</th><th>Summa</th><th>Status</th><th>Sana</th></tr></thead>
                            <tbody>
                                {% for o in orders %}
                                <tr>
                                    <td>#{{ o['id'] }}</td>
                                    <td>{{ o['shop_name'] }}</td>
                                    <td>{{ "{:,.0f}".format(o['total_sum']) }} so'm</td>
                                    <td><span class="badge bg-secondary">{{ o['status'] }}</span></td>
                                    <td>{{ o['date'] }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="tab-pane fade" id="tab-inventory">
                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3">Ombor Qoldiqlari</h5>
                        <table class="table table-custom">
                            <thead><tr><th>Mahsulot</th><th>Kategoriya</th><th>Qoldiq</th><th>Optom narx</th></tr></thead>
                            <tbody>
                                {% for p in products %}
                                <tr>
                                    <td>{{ p['name'] }}</td>
                                    <td>{{ p['category'] }}</td>
                                    <td>{{ p['stock'] }}</td>
                                    <td>{{ "{:,.0f}".format(p['optom_price'] if 'optom_price' in p.keys() else 0) }} so'm</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="tab-pane fade" id="tab-clients">
                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3">Do'konlar va Qarzlar</h5>
                        <table class="table table-custom">
                            <thead><tr><th>Do'kon</th><th>Telefon</th><th>Qarzdorlik</th></tr></thead>
                            <tbody>
                                {% for s in shops %}
                                <tr>
                                    <td>{{ s['name'] }}</td>
                                    <td>{{ s['phone'] }}</td>
                                    <td class="text-danger fw-bold">{{ "{:,.0f}".format(s['debt']) }} so'm</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="tab-pane fade" id="tab-kassa">
                    <div class="row g-4">
                        <div class="col-md-6">
                            <div class="card-glass p-4">
                                <h5 class="text-success mb-3">Kirim qilish</h5>
                                <form action="/add_income" method="POST">
                                    <div class="mb-3"><label class="small fw-bold">Manba:</label><input type="text" name="source" class="form-control" required></div>
                                    <div class="mb-3"><label class="small fw-bold">Summa:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-success w-100">Kirim</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="card-glass p-4">
                                <h5 class="text-danger mb-3">Chiqim (Rasxod)</h5>
                                <form action="/add_expense" method="POST">
                                    <div class="mb-3"><label class="small fw-bold">Sabab:</label><input type="text" name="reason" class="form-control" required></div>
                                    <div class="mb-3"><label class="small fw-bold">Summa:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-danger w-100">Chiqim</button>
                                </form>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="tab-pane fade" id="tab-reports">
                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3">Moliyaviy Hisobot</h5>
                        <p>Umumiy daromad va tahlillar shu yerda ko'rsatiladi.</p>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <title>Xoji Aka ERP — Kirish</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #0f172a; height: 100vh; display: flex; align-items: center; justify-content: center; font-family: sans-serif; }
        .login-card { background: #ffffff; padding: 30px; border-radius: 16px; width: 100%; max-width: 400px; box-shadow: 0 10px 25px rgba(0,0,0,0.3); }
        .brand-xa { font-family: 'Georgia', serif; font-weight: 900; font-size: 38px; color: #0f172a; font-style: italic; text-align: center; }
        .brand-line { height: 3px; background-color: #ef4444; width: 50px; margin: 5px auto 20px auto; border-radius: 2px; }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="brand-xa">XA</div>
        <div class="brand-line"></div>
        <h5 class="text-center mb-4 text-secondary fw-bold">Xoji Aka ERP Tizimi</h5>
        {% if error %}<div class="alert alert-danger py-2 small text-center">{{ error }}</div>{% endif %}
        <form method="POST">
            <div class="mb-3"><label class="form-label small fw-bold">Login:</label><input type="text" name="username" class="form-control" required autocomplete="off"></div>
            <div class="mb-4"><label class="form-label small fw-bold">Parol:</label><input type="password" name="password" class="form-control" required></div>
            <button type="submit" class="btn btn-primary w-100 py-2 fw-bold">Tizimga Kirish</button>
        </form>
    </div>
</body>
</html>
"""

@app.before_request
def require_login():
    allowed_routes = ['login', 'static']
    if request.endpoint not in allowed_routes and not session.get('logged_in'):
        return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if username == 'xoji aka' and password == '023123.+':
            session['logged_in'] = True
            return redirect(url_for('operator_dashboard'))
        else:
            error = 'Login yoki parol xato kiritildi!'
    return render_template_string(LOGIN_TEMPLATE, error=error)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/')
def operator_dashboard():
    selected_date = request.args.get('filter_date', '')
    conn = get_db_connection()
    orders = conn.execute('SELECT * FROM orders ORDER BY id DESC').fetchall()
    products = conn.execute('SELECT * FROM products').fetchall()
    shops = conn.execute('SELECT * FROM shops').fetchall()
    total_debt = conn.execute('SELECT SUM(debt) FROM shops').fetchone()[0] or 0
    total_income = conn.execute('SELECT SUM(amount) FROM incomes').fetchone()[0] or 0
    total_expense = conn.execute('SELECT SUM(amount) FROM expenses').fetchone()[0] or 0
    kassa_balance = total_income - total_expense
    daily_sum = 0
    conn.close()
    return render_template_string(HTML_TEMPLATE, orders=orders, products=products, shops=shops, daily_sum=daily_sum, total_debt=total_debt, kassa_balance=kassa_balance, filter_date=selected_date)

@app.route('/add_income', methods=['POST'])
def add_income():
    source, amount = request.form['source'], float(request.form['amount'])
    today = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn = get_db_connection()
    conn.execute('INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)', (source, amount, today))
    conn.commit()
    conn.close()
    return redirect(url_for('operator_dashboard'))

@app.route('/add_expense', methods=['POST'])
def add_expense():
    reason, amount = request.form['reason'], float(request.form['amount'])
    today = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn = get_db_connection()
    conn.execute('INSERT INTO expenses (reason, amount, date) VALUES (?, ?, ?)', (reason, amount, today))
    conn.commit()
    conn.close()
    return redirect(url_for('operator_dashboard'))

@app.route('/export_excel')
def export_excel():
    conn = get_db_connection()
    orders_df = pd.read_sql_query("SELECT id, shop_name, agent_name, items_text, total_sum, discount, status, date FROM orders", conn)
    products_df = pd.read_sql_query("SELECT name, category, stock, cost_price, optom_price FROM products", conn)
    shops_df = pd.read_sql_query("SELECT name, phone, debt, visit_days FROM shops", conn)
    conn.close()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        orders_df.to_excel(writer, sheet_name='Buyurtmalar', index=False)
        products_df.to_excel(writer, sheet_name='Ombor', index=False)
        shops_df.to_excel(writer, sheet_name='Do\'konlar', index=False)
    output.seek(0)
    return send_file(output, download_name="xoji_aka_erp_report.xlsx", as_attachment=True)


if __name__ == '__main__':
    init_db()
    auto_insert_products()
    
    # Telegram botni alohida potokda (thread) ishga tushiramiz
    bot_thread = threading.Thread(target=lambda: bot.infinity_polling(skip_pending=True), daemon=True)
    bot_thread.start()
    
    # Flask veb ilovasini ishga tushiramiz
    app.run(host='0.0.0.0', port=5000, debug=False)
