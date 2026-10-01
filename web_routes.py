import os
import re
from datetime import datetime
from flask import Blueprint, render_template_string, request, session, redirect, url_for, send_file

from database import get_db_connection
from excel_utils import create_svodka_excel
from bot_handlers import bot

web_bp = Blueprint('web_bp', __name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@web_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form.get('password') == 'admin123':
            session['logged_in'] = True
            return redirect(url_for('web_bp.operator_dashboard'))
        return "<h3 style='color:red; text-align:center;'>Xato parol! <a href='/login'>Qaytadan urinish</a></h3>"
    
    return '''
    <!DOCTYPE html>
    <html><head><title>Login - Xoji Aka ERP</title><meta name="viewport" content="width=device-width, initial-scale=1"></head>
    <body style="display:flex; justify-content:center; align-items:center; height:100vh; background:#f0f2f5; font-family:sans-serif;">
        <form method="POST" style="background:#fff; padding:30px; border-radius:10px; box-shadow:0 4px 10px rgba(0,0,0,0.1); width:300px;">
            <h3 style="text-align:center; margin-bottom:20px;">Xoji Aka ERP</h3>
            <input type="password" name="password" placeholder="Parol..." required style="width:100%; padding:10px; margin-bottom:15px; border:1px solid #ccc; border-radius:5px; box-sizing:border-box;">
            <button type="submit" style="width:100%; padding:10px; background:#2563eb; color:#fff; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">Kirish</button>
        </form>
    </body></html>
    '''

@web_bp.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('web_bp.login'))

@web_bp.route('/')
def operator_dashboard():
    if not session.get('logged_in'):
        return redirect(url_for('web_bp.login'))
        
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    
    conn = get_db_connection()
    query = "SELECT * FROM orders WHERE 1=1"
    params = []
    
    if start_date:
        query += " AND date >= ?"
        params.append(f"{start_date} 00:00")
    if end_date:
        query += " AND date <= ?"
        params.append(f"{end_date} 23:59")
        
    query += " ORDER BY id DESC"
    
    orders = conn.execute(query, params).fetchall()
    users = conn.execute("SELECT * FROM users").fetchall()
    conn.close()
    
    return render_template_string(HTML_TEMPLATE, orders=orders, users=users, start_date=start_date, end_date=end_date)

@web_bp.route('/export_expeditor', methods=['POST'])
def export_expeditor():
    if not session.get('logged_in'): 
        return redirect(url_for('web_bp.login'))
    
    order_ids = request.form.getlist('order_ids')
    if not order_ids:
        return "<script>alert('Hech qanday buyurtma tanlanmadi!'); window.location.href='/';</script>"
    
    conn = get_db_connection()
    placeholders = ','.join('?' for _ in order_ids)
    orders = conn.execute(f"SELECT * FROM orders WHERE id IN ({placeholders})", order_ids).fetchall()
    
    aggregated = {}
    total_summa = 0
    shop_list = []
    
    for o in orders:
        shop_list.append(o['shop_name'])
        total_summa += o['total_sum']
        lines = str(o['items_text']).strip().split('\n')
        for line in lines:
            m = re.match(r'^(.*?)\s*-\s*([\d\.]+)x', line.strip())
            if m:
                p_name = m.group(1).strip()
                qty = float(m.group(2))
                aggregated[p_name] = aggregated.get(p_name, 0) + qty

    file_path = os.path.join(UPLOAD_FOLDER, f"Svodka_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    create_svodka_excel(aggregated, file_path)
    
    text_msg = "🚚 *YIG'MA BUYURTMALAR (SVODKA)*\n\n"
    for p_name, qty in aggregated.items():
        text_msg += f"🔹 {p_name}: {qty}\n"
    text_msg += f"\n📍 Do'konlar: {', '.join(set(shop_list))}\n💰 Jami: {total_summa:,.0f} so'm"
    
    expeditors = conn.execute("SELECT tg_id FROM users WHERE role LIKE '%ekspeditor%'").fetchall()
    conn.close()
    
    for exp in expeditors:
        try:
            bot.send_message(exp['tg_id'], text_msg, parse_mode='Markdown')
            with open(file_path, 'rb') as f:
                bot.send_document(exp['tg_id'], f)
        except Exception as e:
            print(f"Xatolik ekspeditorga yuborishda: {e}")
        
    return send_file(file_path, as_attachment=True)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Xoji Aka ERP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f1f5f9; font-family: sans-serif; }
        .sidebar { background: #1e293b; min-height: 100vh; padding: 20px; color: white;}
        .card-custom { background: #ffffff; border-radius: 12px; padding:20px; margin-bottom:20px; border:none; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    </style>
</head>
<body>
    <div class="d-flex">
        <div class="sidebar" style="width: 260px;">
            <h4>Xoji Aka ERP</h4>
            <hr>
            <a href="/" class="text-white text-decoration-none d-block mb-3 fw-bold">📦 Buyurtmalar</a>
            <a href="/logout" class="text-danger text-decoration-none d-block fw-bold">🚪 Chiqish</a>
        </div>
        
        <div class="p-4" style="flex-grow: 1;">
            <div class="card-custom">
                <h5 class="mb-3">📅 Buyurtmalarni Filtrlash (Dan - Gacha)</h5>
                <form method="GET" action="/" class="row g-2 mb-3">
                    <div class="col-md-4">
                        <input type="date" name="start_date" class="form-control" value="{{ start_date }}">
                    </div>
                    <div class="col-md-4">
                        <input type="date" name="end_date" class="form-control" value="{{ end_date }}">
                    </div>
                    <div class="col-md-4 d-flex gap-2">
                        <button type="submit" class="btn btn-primary w-100">Filtrlash</button>
                        <a href="/" class="btn btn-outline-secondary">Tozalash</a>
                    </div>
                </form>
            </div>

            <div class="card-custom">
                <form method="POST" action="/export_expeditor">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h5 class="m-0">📦 Buyurtmalar Ro'yxati</h5>
                        <button type="submit" class="btn btn-success">🚚 Tanlanganlarni Ekspeditorga Yuborish (Svodka)</button>
                    </div>
                    <table class="table table-hover align-middle">
                        <thead class="table-light">
                            <tr>
                                <th><input type="checkbox" id="checkAll"></th>
                                <th>ID</th>
                                <th>Do'kon</th>
                                <th>Summa</th>
                                <th>Sana</th>
                                <th>Holat</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for o in orders %}
                            <tr>
                                <td><input type="checkbox" name="order_ids" value="{{ o.id }}" class="order-check"></td>
                                <td><b>#{{ o.id }}</b></td>
                                <td>{{ o.shop_name }}</td>
                                <td>{{ "{:,.0f}".format(o.total_sum) }} so'm</td>
                                <td><small class="text-muted">{{ o.date }}</small></td>
                                <td><span class="badge bg-info text-dark">{{ o.status }}</span></td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </form>
            </div>
            
            <div class="card-custom">
                <h5 class="mb-3">👥 Xodimlar va Foydalanuvchilar</h5>
                <table class="table table-bordered">
                    <thead class="table-light">
                        <tr><th>TG ID</th><th>Ism</th><th>Telefon</th><th>Rol</th></tr>
                    </thead>
                    <tbody>
                        {% for u in users %}
                        <tr>
                            <td>{{ u.tg_id }}</td>
                            <td>{{ u.name }}</td>
                            <td>{{ u.phone }}</td>
                            <td><span class="badge bg-secondary">{{ u.role }}</span></td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <script>
        document.getElementById('checkAll').addEventListener('change', function() {
            let checkboxes = document.querySelectorAll('.order-check');
            checkboxes.forEach(cb => cb.checked = this.checked);
        });
    </script>
</body>
</html>
"""