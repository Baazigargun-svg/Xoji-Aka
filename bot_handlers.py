import telebot
from telebot import types
from datetime import datetime
from database import get_db_connection, ADMIN_ID

BOT_TOKEN = "BOT_TOKEN_SHU_YERGA_YOZING"
SEX_GROUP_ID = -1003936599812  

bot = telebot.TeleBot(BOT_TOKEN)

def get_main_menu(role):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    if role == 'admin':
        markup.row('📦 Sklad & Mahsulotlar', '📊 Kunlik Hisobot')
        markup.row('👥 Agentlar boshqaruvi', "🏪 AKB (Do'konlar & Qarz)")
    elif 'agent' in role:
        markup.row('🛒 Yangi Buyurtma Urish', "🏪 Do'kon qo'shish")
        markup.row("💰 Qarz/To'lov yozish", '📜 Mening Buyurtmalarim')
    elif 'sex' in role:
        markup.row('📦 Skladga Kirim Qilish', '📋 Ombordagi Qoldiqlar')
    elif 'ekspeditor' in role:
        markup.row("💰 Kassaga pul topshirish", "📜 Menga biriktirilganlar")
    else:
        markup.add("📝 Ro'yxatdan o'tish")
    return markup

@bot.message_handler(commands=['start'])
def start_command(message):
    tg_id = message.from_user.id
    if tg_id == ADMIN_ID:
        bot.send_message(message.chat.id, 'Xoji aka, xush kelibsiz!', reply_markup=get_main_menu('admin'))
        return
    conn = get_db_connection()
    user = conn.execute('SELECT role, name FROM users WHERE tg_id = ?', (tg_id,)).fetchone()
    conn.close()
    if user:
        bot.send_message(message.chat.id, f"Salom {user['name']}!", reply_markup=get_main_menu(user['role']))
    else:
        bot.send_message(message.chat.id, "Ro'yxatdan o'ting.", reply_markup=get_main_menu('guest'))

@bot.message_handler(func=lambda m: m.text == "💰 Kassaga pul topshirish")
def expeditor_kassa_start(message):
    msg = bot.send_message(message.chat.id, "Kassaga qancha pul topshiryapsiz? (Faqat raqam yozing):")
    bot.register_next_step_handler(msg, process_kassa_amount)

def process_kassa_amount(message):
    try:
        amount = float(message.text)
        exp_id = message.from_user.id
        markup = types.InlineKeyboardMarkup()
        markup.row(
            types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"kassa_approve_{exp_id}_{amount}"),
            types.InlineKeyboardButton("❌ Bekor qilish", callback_data=f"kassa_reject_{exp_id}_{amount}")
        )
        bot.send_message(ADMIN_ID, f"🔔 **Ekspeditor topshiruvi:**\nXodim ID: {exp_id}\nSumma: {amount:,.0f} so'm", reply_markup=markup, parse_mode='Markdown')
        bot.send_message(message.chat.id, "⏳ Adminga tasdiqlash uchun yuborildi. Kuting...")
    except:
        bot.send_message(message.chat.id, "Xato! Faqat raqam kiriting.")

@bot.callback_query_handler(func=lambda call: call.data.startswith('kassa_'))
def kassa_decision(call):
    parts = call.data.split('_')
    action, exp_id, amount = parts[1], int(parts[2]), float(parts[3])
    
    if action == 'approve':
        today = datetime.now().strftime('%Y-%m-%d %H:%M')
        conn = get_db_connection()
        conn.execute('INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)', (f"Ekspeditor (ID: {exp_id})", amount, today))
        conn.commit()
        conn.close()
        bot.edit_message_text(f"✅ {amount:,.0f} so'm kassaga qabul qilindi.", chat_id=call.message.chat.id, message_id=call.message.message_id)
        try:
            bot.send_message(exp_id, f"✅ Admin tasdiqladi! {amount:,.0f} so'm kassaga kiritildi.")
        except: pass
    else:
        bot.edit_message_text(f"❌ {amount:,.0f} so'm qabul qilinmadi.", chat_id=call.message.chat.id, message_id=call.message.message_id)
        try:
            bot.send_message(exp_id, "❌ Admin pul topshirishni bekor qildi.")
        except: pass

@bot.message_handler(func=lambda m: m.text == "📝 Ro'yxatdan o'tish")
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