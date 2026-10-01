import telebot
from telebot import types
from datetime import datetime
import os
from database import get_db_connection, ADMIN_ID
from excel_utils import create_excel_invoice, create_excel_sex_income

BOT_TOKEN = "8573337094:AAGhQXE6IheONVsJgxyLfpeyjAqY_xbtYJk"
SEX_GROUP_ID = -1003936599812  

bot = telebot.TeleBot(BOT_TOKEN)
user_steps = {}

def get_main_menu(role):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    if role == 'admin':
        markup.row(types.KeyboardButton('📦 Sklad & Mahsulotlar'), types.KeyboardButton('📊 Kunlik Hisobot'))
        markup.row(types.KeyboardButton('👥 Agentlar boshqaruvi'), types.KeyboardButton("🏪 AKB (Do'konlar & Qarz)"))
    elif role == 'agent':
        markup.row(types.KeyboardButton('🛒 Yangi Buyurtma Urish'), types.KeyboardButton("🏪 Do'kon qo'shish (AKB)"))
        markup.row(types.KeyboardButton("💰 Qarz/To'lov yozish"), types.KeyboardButton('📜 Mening Buyurtmalarim'))
        markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
    elif role == 'sex':
        markup.row(types.KeyboardButton('📦 Skladga Mahsulot Kirim Qilish'), types.KeyboardButton('📋 Ombordagi Qoldiqlar'))
        markup.row(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
    else:
        markup.add(types.KeyboardButton("📝 Ro'yxatdan o'tish"))
    return markup

@bot.message_handler(commands=['start'])
def start_command(message):
    tg_id = message.from_user.id
    if tg_id == ADMIN_ID:
        bot.send_message(message.chat.id, 'Xoji aka, xush kelibsiz! Boshqaruv paneli tayyor.', reply_markup=get_main_menu('admin'))
        return
    conn = get_db_connection()
    user = conn.execute('SELECT role, name FROM users WHERE tg_id = ?', (tg_id,)).fetchone()
    conn.close()
    if user:
        role, name = user['role'], user['name']
        if role == 'pending':
            bot.send_message(message.chat.id, f"Salom {name}. So'rovingiz admin tasdig'ini kutyapti.")
        elif ',' in role:
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            if 'agent' in role: markup.add(types.KeyboardButton('👤 Agent rejimi'))
            if 'sex' in role: markup.add(types.KeyboardButton('🏭 Sex rejimi'))
            bot.send_message(message.chat.id, f"Salom {name}! Iltimos, ish rejimini tanlang:", reply_markup=markup)
        else:
            bot.send_message(message.chat.id, f"Salom {name}! Ishni boshlashimiz mumkin.", reply_markup=get_main_menu(role))
    else:
        bot.send_message(message.chat.id, "Assalomu alaykum! Tizimga xush kelibsiz. Davom etish uchun ro'yxatdan o'ting.", reply_markup=get_main_menu('guest'))

@bot.message_handler(func=lambda message: message.text in ['👤 Agent rejimi', '🏭 Sex rejimi', '🔄 Rolni almashtirish (Sex / Agent)'])
def switch_role_menu(message):
    tg_id = message.from_user.id
    conn = get_db_connection()
    user = conn.execute('SELECT role, name FROM users WHERE tg_id = ?', (tg_id,)).fetchone()
    conn.close()
    if not user: return
    role_str = user['role']
    if message.text == '👤 Agent rejimi' or ('agent' in role_str and 'sex' in role_str and message.text != '🏭 Sex rejimi'):
        if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            markup.add(types.KeyboardButton('👤 Agent rejimi'), types.KeyboardButton('🏭 Sex rejimi'))
            bot.send_message(message.chat.id, 'Qaysi rejimga oʻtmoqchisiz?', reply_markup=markup)
            return
        bot.send_message(message.chat.id, '🛒 Agent rejimiga oʻtdingiz.', reply_markup=get_main_menu('agent'))
    elif message.text == '🏭 Sex rejimi':
        bot.send_message(message.chat.id, '🏭 Sex rejimiga oʻtdingiz.', reply_markup=get_main_menu('sex'))

# --- SEX KIRIM ---
@bot.message_handler(func=lambda message: message.text == '📦 Skladga Mahsulot Kirim Qilish')
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
    msg = bot.send_message(message.chat.id, '🏭 Kirim qilinadigan mahsulotni tanlang:', reply_markup=markup)
    bot.register_next_step_handler(msg, sex_income_choose_product)

def sex_income_choose_product(message):
    if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
        switch_role_menu(message)
        return
    if message.text == '✅ Kirimni yakunlash':
        finish_sex_income(message)
        return
    conn = get_db_connection()
    p_check = conn.execute('SELECT id FROM products WHERE name = ?', (message.text,)).fetchone()
    conn.close()
    if not p_check:
        bot.send_message(message.chat.id, "❌ Bunday mahsulot topilmadi. Ro'yxatdan tanlang:")
        send_sex_product_menu(message)
        return
    user_steps[message.from_user.id]['current_product'] = message.text
    msg = bot.send_message(message.chat.id, f"🔢 <b>{message.text}</b> uchun miqdorni (kg / dona) kiriting:", parse_mode='HTML', reply_markup=types.ReplyKeyboardRemove())
    bot.register_next_step_handler(msg, sex_income_add_to_cart)

def sex_income_add_to_cart(message):
    uid = message.from_user.id
    try:
        qty = float(message.text)
        p_name = user_steps[uid]['current_product']
        user_steps[uid]['sex_cart'][p_name] = qty
        bot.send_message(message.chat.id, f"📥 Qo'shildi: <b>{p_name}</b> — <b>{qty}</b>", parse_mode='HTML')
        send_sex_product_menu(message)
    except:
        bot.send_message(message.chat.id, '❌ Xatolik! Faqat raqam kiriting.')
        send_sex_product_menu(message)

def finish_sex_income(message):
    uid = message.from_user.id
    data = user_steps.get(uid)
    if not data or not data.get('sex_cart'):
        bot.send_message(message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu('sex'))
        return
    conn = get_db_connection()
    cursor = conn.cursor()
    items_text = ''
    cart_items = []
    for p_name, qty in data['sex_cart'].items():
        cursor.execute('UPDATE products SET stock = stock + ? WHERE name = ?', (qty, p_name))
        items_text += f"🔹 {p_name}: +{qty}\n"
        cart_items.append({'name': p_name, 'qty': qty})
    
    staff_res = cursor.execute('SELECT name FROM users WHERE tg_id = ?', (uid,)).fetchone()
    staff_name = staff_res['name'] if staff_res else 'Nomalum'
    bugun = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    cursor.execute('INSERT INTO product_incomes (product_name, qty, cost_price, date) VALUES (?, ?, 0, ?)', (items_text, sum(data['sex_cart'].values()), bugun))
    income_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    excel_file = create_excel_sex_income(income_id, staff_name, bugun, cart_items)
    bot.send_message(message.chat.id, f'✅ <b>Skladga mahsulotlar muvaffaqiyatli kirim qilindi!</b>\n\n{items_text}', parse_mode='HTML', reply_markup=get_main_menu('sex'))
    with open(excel_file, 'rb') as doc:
        bot.send_document(message.chat.id, doc, caption=f'📄 Kirim Nakladnoy (#{income_id})')
    try:
        group_text = f"🏭 <b>YANGI SKLAD KIRIMI (#{income_id})</b>\n👤 <b>Mas’ul:</b> {staff_name}\n📅 <b>Sana:</b> {bugun}\n\n<b>Kirim qilingan mahsulotlar:</b>\n{items_text}"
        bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')
        with open(excel_file, 'rb') as doc_group:
            bot.send_document(SEX_GROUP_ID, doc_group, caption=f'📄 Kirim Nakladnoy (#{income_id})')
    except Exception as e:
        print('Sex guruhiga yuborish xatosi:', e)
    if os.path.exists(excel_file): os.remove(excel_file)
    if uid in user_steps: del user_steps[uid]

@bot.message_handler(func=lambda message: message.text == '📋 Ombordagi Qoldiqlar')
def sex_view_stock(message):
    conn = get_db_connection()
    prods = conn.execute('SELECT name, stock FROM products').fetchall()
    conn.close()
    text = "📦 <b>Ombordagi qoldiqlar:</b>\n\n"
    for p in prods:
        stk = p['stock'] if p['stock'] is not None else 0
        text += f"🔹 {p['name']}: <b>{int(stk)}</b>\n"
    bot.send_message(message.chat.id, text, parse_mode='HTML', reply_markup=get_main_menu('sex'))

# --- BUYURTMA BERISH (KATEGORIYA VA KOMMENTARIYA BILAN) ---
@bot.message_handler(func=lambda message: message.text == '🛒 Yangi Buyurtma Urish')
def start_order(message):
    conn = get_db_connection()
    shops = conn.execute('SELECT name FROM shops').fetchall()
    conn.close()
    if not shops:
        bot.send_message(message.chat.id, "❌ Do'konlar yo'q.")
        return
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for s in shops:
        markup.add(types.KeyboardButton(s['name']))
    markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
    msg = bot.send_message(message.chat.id, 'Do\'konni tanlang:', reply_markup=markup)
    bot.register_next_step_handler(msg, choose_price_type)

def choose_price_type(message):
    if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
        switch_role_menu(message)
        return
    user_steps[message.from_user.id] = {'shop_name': message.text, 'cart': {}, 'price_type': None}
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True).row(
        types.KeyboardButton('💰 Ulgurji (Optom)'), types.KeyboardButton('🛍 Chakana')
    )
    markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
    msg = bot.send_message(message.chat.id, 'Narx turini tanlang:', reply_markup=markup)
    bot.register_next_step_handler(msg, show_categories_to_agent)

def show_categories_to_agent(message):
    if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
        switch_role_menu(message)
        return
    p_type = 'optom' if 'Ulgurji' in message.text else 'chakana'
    user_steps[message.from_user.id]['price_type'] = p_type
    conn = get_db_connection()
    categories = conn.execute('SELECT DISTINCT category FROM products').fetchall()
    conn.close()
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for c in categories:
        cat_name = c['category'] if c['category'] else 'Boshqa'
        markup.add(types.KeyboardButton(f'📁 {cat_name}'))
    markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))
    markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
    msg = bot.send_message(message.chat.id, '📁 Mahsulot kategoriyasini tanlang yoki buyurtmani yakunlang:', reply_markup=markup)
    bot.register_next_step_handler(msg, choose_product_category)

def choose_product_category(message):
    if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
        switch_role_menu(message)
        return
    if message.text == '✅ Buyurtmani yakunlash':
        ask_order_comment(message)
        return
    cat_name = message.text.replace('📁 ', '').strip()
    user_steps[message.from_user.id]['current_category'] = cat_name
    conn = get_db_connection()
    prods = conn.execute('SELECT name FROM products WHERE category = ?', (cat_name,)).fetchall()
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
    msg = bot.send_message(message.chat.id, f'📦 <b>{cat_name}</b> kategoriyasidagi mahsulotni tanlang:', parse_mode='HTML', reply_markup=markup)
    bot.register_next_step_handler(msg, ask_quantity_or_navigation)

def ask_quantity_or_navigation(message):
    if message.text == '🔙 Kategoriyalarga qaytish':
        show_categories_to_agent_again(message)
        return
    if message.text == '✅ Buyurtmani yakunlash':
        ask_order_comment(message)
        return
    conn = get_db_connection()
    p_check = conn.execute('SELECT id FROM products WHERE name = ?', (message.text,)).fetchone()
    conn.close()
    if not p_check:
        bot.send_message(message.chat.id, "❌ Bunday mahsulot topilmadi. Iltimos, tugmalardan tanlang:")
        choose_product_category(message)
        return
    user_steps[message.from_user.id]['current_product'] = message.text
    msg = bot.send_message(message.chat.id, f"🔢 <b>{message.text}</b> miqdorini kiriting:", parse_mode='HTML', reply_markup=types.ReplyKeyboardRemove())
    bot.register_next_step_handler(msg, add_to_cart)

def show_categories_to_agent_again(message):
    uid = message.from_user.id
    conn = get_db_connection()
    categories = conn.execute('SELECT DISTINCT category FROM products').fetchall()
    conn.close()
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for c in categories:
        cat_name = c['category'] if c['category'] else 'Boshqa'
        markup.add(types.KeyboardButton(f'📁 {cat_name}'))
    markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))
    msg = bot.send_message(message.chat.id, '📁 Kategoriyani tanlang:', reply_markup=markup)
    bot.register_next_step_handler(msg, choose_product_category)

def add_to_cart(message):
    uid = message.from_user.id
    try:
        qty = float(message.text)
        p_name = user_steps[uid]['current_product']
        user_steps[uid]['cart'][p_name] = qty
        bot.send_message(message.chat.id, f"📥 Qo'shildi: {p_name} - {qty}")
        cat_name = user_steps[uid].get('current_category', 'Boshqa')
        conn = get_db_connection()
        prods = conn.execute('SELECT name FROM products WHERE category = ?', (cat_name,)).fetchall()
        conn.close()
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        for p in prods:
            markup.add(types.KeyboardButton(p['name']))
        markup.add(types.KeyboardButton('🔙 Kategoriyalarga qaytish'))
        markup.add(types.KeyboardButton('✅ Buyurtmani yakunlash'))
        msg = bot.send_message(message.chat.id, 'Yana mahsulot qo\'shasizmi yoki buyurtmani yakunlaysizmi?', reply_markup=markup)
        bot.register_next_step_handler(msg, ask_quantity_or_navigation)
    except:
        bot.send_message(message.chat.id, '❌ Xatolik! Faqat raqam kiriting.')
        show_categories_to_agent_again(message)

def ask_order_comment(message):
    uid = message.from_user.id
    data = user_steps.get(uid)
    if not data or not data.get('cart'):
        bot.send_message(message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu('agent'))
        return
    msg = bot.send_message(message.chat.id, "📝 Buyurtma uchun **kommentariya (izoh)** yozing (masalan: <i>'Ertalabga yetkazilsin'</i>):", parse_mode='HTML', reply_markup=types.ReplyKeyboardRemove())
    bot.register_next_step_handler(msg, finish_order_with_comment)

def finish_order_with_comment(message):
    uid = message.from_user.id
    comment = message.text if message.text else ''
    data = user_steps.get(uid)
    if not data or not data['cart']:
        bot.send_message(message.chat.id, "Savat bo'sh!", reply_markup=get_main_menu('agent'))
        return
    
    conn = get_db_connection()
    total_sum, items_text, p_type = 0, '', data['price_type']
    excel_cart_items = []
    
    for p_name, qty in data['cart'].items():
        prod = conn.execute('SELECT optom_price, chakana_price, stock, id FROM products WHERE name = ?', (p_name,)).fetchone()
        optom_p = prod['optom_price'] if prod['optom_price'] is not None else 0
        chakana_p = prod['chakana_price'] if prod['chakana_price'] is not None else 0
        stock_p = prod['stock'] if prod['stock'] is not None else 0
        
        price = optom_p if p_type == 'optom' else chakana_p
        summa = price * qty
        total_sum += summa
        items_text += f"{p_name} - {qty}x = {summa:,.0f} so'm\n"
        conn.execute('UPDATE products SET stock = ? WHERE id = ?', (stock_p - qty, prod['id']))
        excel_cart_items.append({'name': p_name, 'qty': qty, 'price': price})
        
    # Do'kon qarzini yangilash
    shop_name = data['shop_name']
    shop_res = conn.execute('SELECT debt FROM shops WHERE name = ?', (shop_name,)).fetchone()
    current_debt = shop_res['debt'] if shop_res and shop_res['debt'] is not None else 0
    new_debt = current_debt + total_sum
    conn.execute('UPDATE shops SET debt = ? WHERE name = ?', (new_debt, shop_name))
    
    agent_res = conn.execute('SELECT name FROM users WHERE tg_id = ?', (uid,)).fetchone()
    agent_name = agent_res['name'] if agent_res else 'Nomalum'
    bugun = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    cursor = conn.cursor()
    cursor.execute('INSERT INTO orders (shop_name, agent_name, total_sum, items_text, status, date, price_type, comment) VALUES (?, ?, ?, ?, ?, ?, ?, ?)', 
                   (shop_name, agent_name, total_sum, items_text, 'Yangi', bugun, p_type, comment))
    order_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    excel_file = create_excel_invoice(order_id, shop_name, agent_name, bugun, p_type, excel_cart_items, comment)
    
    bot.send_message(message.chat.id, f'✅ Buyurtma qabul qilindi! Jami: {total_sum:,.0f} so\'m', reply_markup=get_main_menu('agent'))
    bot.send_message(ADMIN_ID, f"🔔 YANGI BUYURTMA (#{order_id}):\n\nDo'kon: {shop_name}\nSumma: {total_sum:,.0f} so'm\nIzoh: {comment}")
    with open(excel_file, 'rb') as doc:
        bot.send_document(ADMIN_ID, doc, caption=f'📄 Nakladnoy (#{order_id})')
    if os.path.exists(excel_file): os.remove(excel_file)
    
    try:
        group_text = f"📝 <b>YANGI BUYURTMA (#{order_id})</b>\n🏪 <b>Do'kon:</b> {shop_name}\n👤 <b>Agent:</b> {agent_name}\n💬 <b>Izoh:</b> {comment}\n💰 <b>Jami summa:</b> {total_sum:,.0f} so'm\n\n<b>Mahsulotlar:</b>\n{items_text}"
        bot.send_message(SEX_GROUP_ID, group_text, parse_mode='HTML')
        with open(create_excel_invoice(order_id, shop_name, agent_name, bugun, p_type, excel_cart_items, comment), 'rb') as doc_group:
            bot.send_document(SEX_GROUP_ID, doc_group, caption=f'📄 Nakladnoy (#{order_id})')
    except Exception as e:
        print('Guruhga yuborish xatosi:', e)
        
    if uid in user_steps: del user_steps[uid]

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
    conn.execute("INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, ?, ?, 'pending')", (message.from_user.id, name, message.text))
    conn.commit()
    conn.close()
    bot.send_message(message.chat.id, "⏳ Admin tasdig'i kutilmoqda.")
    bot.send_message(ADMIN_ID, f"🔔 Yangi ro'yxatdan o'tgan: {name} ({message.text})")

@bot.message_handler(func=lambda message: message.text == '📜 Mening Buyurtmalarim')
def my_orders(message):
    conn = get_db_connection()
    res = conn.execute('SELECT name FROM users WHERE tg_id = ?', (message.from_user.id,)).fetchone()
    agent_name = res['name'] if res else ''
    orders = conn.execute('SELECT shop_name, total_sum, status, date FROM orders WHERE agent_name = ? ORDER BY id DESC LIMIT 5', (agent_name,)).fetchall()
    conn.close()
    text = "<b>📜 Oxirgi buyurtmalar:</b>\n\n"
    for o in orders:
        t_sum = o['total_sum'] if o['total_sum'] is not None else 0
        text += f"🏪 {o['shop_name']} | {t_sum:,.0f} so'm | {o['status']} | {o['date']}\n\n"
    bot.send_message(message.chat.id, text, parse_mode='HTML')

@bot.message_handler(func=lambda message: message.text == "💰 Qarz/To'lov yozish")
def pay_debt_start(message):
    conn = get_db_connection()
    shops = conn.execute('SELECT name FROM shops').fetchall()
    conn.close()
    if not shops:
        bot.send_message(message.chat.id, "❌ Do'konlar yo'q.")
        return
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for s in shops:
        markup.add(types.KeyboardButton(s['name']))
    markup.add(types.KeyboardButton('🔄 Rolni almashtirish (Sex / Agent)'))
    msg = bot.send_message(message.chat.id, 'Do\'konni tanlang:', reply_markup=markup)
    bot.register_next_step_handler(msg, ask_payment_amount)

def ask_payment_amount(message):
    if message.text == '🔄 Rolni almashtirish (Sex / Agent)':
        switch_role_menu(message)
        return
    shop_name = message.text
    msg = bot.send_message(message.chat.id, f"'{shop_name}' qancha to'lov kirdi (summa):", reply_markup=types.ReplyKeyboardRemove())
    bot.register_next_step_handler(msg, process_payment, shop_name)

def process_payment(message, shop_name):
    try:
        amount = float(message.text)
        today = datetime.now().strftime('%Y-%m-%d %H:%M')
        conn = get_db_connection()
        s_res = conn.execute('SELECT debt FROM shops WHERE name = ?', (shop_name,)).fetchone()
        curr_debt = s_res['debt'] if s_res and s_res['debt'] is not None else 0
        conn.execute('UPDATE shops SET debt = ? WHERE name = ?', (curr_debt - amount, shop_name))
        conn.execute('INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)', (f"Qarz to'lovi ({shop_name})", amount, today))
        conn.commit()
        conn.close()
        bot.send_message(message.chat.id, f"✅ To\'lov yozildi: {amount:,.0f} so\'m chegirildi.", reply_markup=get_main_menu('agent'))
    except:
        bot.send_message(message.chat.id, '❌ Faqat raqam kiriting.', reply_markup=get_main_menu('agent'))

# Callback handlers for admin
@bot.callback_query_handler(func=lambda call: call.data.startswith('st_'))
def change_status_logic(call):
    _, mode, order_id = call.data.split('_')
    order_id = int(order_id)
    status_map = {'yangi': 'Yangi', 'otgruzka': 'Otgruzka', 'done': 'Yetkazildi', 'otmen': 'Bekor'}
    new_status = status_map[mode]
    conn = get_db_connection()
    conn.execute('UPDATE orders SET status = ? WHERE id = ?', (new_status, order_id))
    conn.commit()
    conn.close()
    bot.send_message(call.message.chat.id, f"✅ Status: <b>{new_status}</b>", parse_mode='HTML')

@bot.callback_query_handler(func=lambda call: call.data.startswith('approve_'))
def approve_agent_cb(call):
    agent_id = int(call.data.split('_')[-1])
    conn = get_db_connection()
    conn.execute("UPDATE users SET role = 'agent' WHERE tg_id = ?", (agent_id,))
    conn.commit()
    conn.close()
    bot.answer_callback_query(call.id, "Foydalanuvchi tasdiqlandi!")
    try:
        bot.send_message(agent_id, "🎉 Sizning so'rovingiz tasdiqlandi! /start bosing", reply_markup=get_main_menu('agent'))
    except: pass
