import subprocess
import sys

# 1. Telegram botni ishga tushirish
bot_process = subprocess.Popen([sys.executable, "main.py"])

# 2. Veb-saytni (Flask) ishga tushirish
try:
    subprocess.run([sys.executable, "app.py"])
except KeyboardInterrupt:
    bot_process.terminate()