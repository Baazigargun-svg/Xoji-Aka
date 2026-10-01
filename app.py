import io
import threading
import os
from flask import Flask

from database import init_web_db, ADMIN_ID
from bot_handlers import bot
from web_routes import web_bp

app = Flask(__name__)
app.secret_key = 'xoji_aka_maxfiy_kalit_2026'

# Flask Blueprint orqali web marshrularni ulaymiz
app.register_blueprint(web_bp)

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
