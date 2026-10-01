import threading
from flask import Flask

from database import init_web_db, ADMIN_ID
from bot_handlers import bot
from web_routes import web_bp

app = Flask(__name__)
app.secret_key = 'xoji_aka_maxfiy_kalit_2026'

# Flask Blueprint orqali web marshrularni ulaymiz
app.register_blueprint(web_bp)

if __name__ == '__main__':
    # 1. Bazani rejimda tekshirib, jadval va adminni yaratamiz
    init_web_db(ADMIN_ID)
    
    # 2. Telegram Botni orqa fonda (background thread) yoqamiz
    bot_thread = threading.Thread(target=bot.polling, kwargs={'none_stop': True})
    bot_thread.daemon = True
    bot_thread.start()
    
    # 3. Web serverni yurgizamiz
    app.run(host='0.0.0.0', port=5000)
