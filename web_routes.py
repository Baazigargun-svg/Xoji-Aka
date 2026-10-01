import os
import re
from datetime import datetime
from flask import Blueprint, render_template_string, request, session, redirect, url_for, send_file
import pandas as pd

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
    shops = conn.execute("SELECT * FROM shops").fetchall()
    products = conn.execute("SELECT * FROM products").fetchall()
    conn.close()
    
    return render_template_string(HTML_TEMPLATE, orders=orders, users=users, shops=shops, products=products, start_date=start_date, end_date=end_date)

@web_bp.route('/add_shop', methods=['POST'])
def add_shop():
    name, phone, visit_days, region, landmark, inventory = (
        request.form['name'], request.form['phone'],
        request.form.get('visit_days', ''), request.form.get('region', ''),
        request.form.get('landmark', ''), request.form.get('inventory', '')
    )
    conn = get_db_connection()
    try:
        conn.execute('INSERT INTO shops (name, phone, debt, visit_days, region, landmark, inventory) VALUES (?, ?, 0, ?, ?, ?, ?)',
                     (name, phone, visit_days, region, landmark, inventory))
        conn.commit()
    except: pass
    conn.close()
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/update_shop/<int:shop_id>', methods=['POST'])
def update_shop(shop_id):
    name, phone, region, landmark, visit_days, inventory = (
        request.form['name'], request.form['phone'], request.form.get('region', ''),
        request.form.get('landmark', ''), request.form.get('visit_days', ''), request.form.get('inventory', '')
    )
    conn = get_db_connection()
    try:
        conn.execute('UPDATE shops SET name = ?, phone = ?, region = ?, landmark = ?, visit_days = ?, inventory = ? WHERE id = ?',
                     (name, phone, region, landmark, visit_days, inventory, shop_id))
        conn.commit()
    except Exception as e: print(e)
    conn.close()
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/update_product/<int:prod_id>', methods=['POST'])
def update_product(prod_id):
    name, category, stock, cost_price, optom_price, chakana_price = (
        request.form['name'], request.form.get('category', 'Boshqa'),
        float(request.form.get('stock', 0)), float(request.form.get('cost_price', 0)),
        float(request.form.get('optom_price', 0)), float(request.form.get('chakana_price', 0))
    )
    conn = get_db_connection()
    try:
        conn.execute('UPDATE products SET name = ?, category = ?, stock = ?, cost_price = ?, optom_price = ?, chakana_price = ? WHERE id = ?',
                     (name, category, stock, cost_price, optom_price, chakana_price, prod_id))
        conn.commit()
    except Exception as e: print('Mahsulotni tahrirlash xatosi:', e)
    conn.close()
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/adjust_stock_web', methods=['POST'])
def adjust_stock_web():
    prod_id = request.form.get('prod_id')
    action = request.form.get('action')
    qty = float(request.form.get('qty', 0))
    conn = get_db_connection()
    if action == 'plus':
        conn.execute('UPDATE products SET stock = stock + ? WHERE id = ?', (qty, prod_id))
    elif action == 'minus':
        conn.execute('UPDATE products SET stock = stock - ? WHERE id = ?', (qty, prod_id))
    conn.commit()
    conn.close()
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/add_agent_web', methods=['POST'])
def add_agent_web():
    name, phone, tg_id, role = request.form.get('name'), request.form.get('phone'), request.form.get('tg_id'), request.form.get('role', 'agent')
    if tg_id:
        try: tg_id = int(tg_id)
        except: tg_id = None
    if name and tg_id:
        conn = get_db_connection()
        try:
            conn.execute('INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, ?, ?, ?)', (tg_id, name, phone or '', role))
            conn.commit()
        except Exception as e: print('Agent qo\'shish xatosi:', e)
        conn.close()
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/import_shops_excel', methods=['POST'])
def import_shops_excel():
    if 'excel_file' not in request.files: return redirect(url_for('web_bp.operator_dashboard'))
    file = request.files['excel_file']
    if file.filename == '': return redirect(url_for('web_bp.operator_dashboard'))
    try:
        df = pd.read_excel(file)
        conn = get_db_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            name = str(row.get('name', row.get("Do'kon nomi", ''))).strip()
            if not name or name == 'nan': continue
            phone = str(row.get('phone', row.get('Telefon', ''))).strip()
            region = str(row.get('region', row.get('Hudud', ''))).strip()
            landmark = str(row.get('landmark', row.get("Mo'ljal", ''))).strip()
            visit_days = str(row.get('visit_days', row.get('Tashrif kunlari', ''))).strip()
            inventory = str(row.get('inventory', row.get('Inventar', ''))).strip()
            debt_val = 0
            for d_key in ['debt', 'Qarz', 'Balans']:
                if d_key in row and pd.notna(row[d_key]):
                    try: debt_val = float(row[d_key]); break
                    except: pass
            cursor.execute('''INSERT INTO shops (name, phone, debt, visit_days, region, landmark, inventory) VALUES (?, ?, ?, ?, ?, ?, ?)
                              ON CONFLICT(name) DO UPDATE SET phone=excluded.phone, region=excluded.region, landmark=excluded.landmark, visit_days=excluded.visit_days, inventory=excluded.inventory''',
                           (name, phone if phone != 'nan' else '', debt_val, visit_days if visit_days != 'nan' else '', region if region != 'nan' else '', landmark if landmark != 'nan' else '', inventory if inventory != 'nan' else ''))
        conn.commit()
        conn.close()
    except Exception as e: print('Excel import xatosi:', e)
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/import_products_excel', methods=['POST'])
def import_products_excel():
    if 'excel_file' not in request.files: return redirect(url_for('web_bp.operator_dashboard'))
    file = request.files['excel_file']
    if file.filename == '': return redirect(url_for('web_bp.operator_dashboard'))
    try:
        df = pd.read_excel(file)
        conn = get_db_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            name = str(row.get('name', row.get('Mahsulot', ''))).strip()
            if not name or name == 'nan': continue
            category = str(row.get('category', row.get('Kategoriya', 'Boshqa'))).strip()
            stock_val = 0
            for s_key in ['stock', 'Qoldiq', 'Soni']:
                if s_key in row and pd.notna(row[s_key]):
                    try: stock_val = float(row[s_key]); break
                    except: pass
            cost_val = 0
            for c_key in ['cost_price', 'Tannarx']:
                if c_key in row and pd.notna(row[c_key]):
                    try: cost_val = float(row[c_key]); break
                    except: pass
            optom_val = 0
            for o_key in ['optom_price', 'Optom']:
                if o_key in row and pd.notna(row[o_key]):
                    try: optom_val = float(row[o_key]); break
                    except: pass
            chakana_val = optom_val
            for ch_key in ['chakana_price', 'Chakana']:
                if ch_key in row and pd.notna(row[ch_key]):
                    try: chakana_val = float(row[ch_key]); break
                    except: pass
            cursor.execute('''INSERT INTO products (name, category, stock, cost_price, optom_price, chakana_price) VALUES (?, ?, ?, ?, ?, ?)
                              ON CONFLICT(name) DO UPDATE SET category=excluded.category, stock=excluded.stock, cost_price=excluded.cost_price, optom_price=excluded.optom_price, chakana_price=excluded.chakana_price''',
                           (name, category, stock_val, cost_val, optom_val, chakana_val))
        conn.commit()
        conn.close()
    except Exception as e: print('Sklad Excel import xatosi:', e)
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/approve_agent/<int:tg_id>', methods=['POST'])
def web_approve_agent(tg_id):
    conn = get_db_connection()
    conn.execute("UPDATE users SET role = 'agent' WHERE tg_id = ?", (tg_id,))
    conn.commit()
    conn.close()
    try:
        bot.send_message(tg_id, "🎉 Sizning so'rovingiz tasdiqlandi! /start bosing", reply_markup=get_main_menu('agent'))
    except: pass
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/delete_agent/<int:tg_id>', methods=['POST'])
def web_delete_agent(tg_id):
    conn = get_db_connection()
    conn.execute('DELETE FROM users WHERE tg_id = ?', (tg_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('web_bp.operator_dashboard'))

@web_bp.route('/export_expeditor', methods=['POST'])
def export_expeditor():
    if not session.get('logged_in'): return redirect(url_for('web_bp.login'))
    order_ids = request.form.getlist('order_ids')
    if not order_ids: return "<script>alert('Hech qanday buyurtma tanlanmadi!'); window.location.href='/';</script>"
    
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
    for p_name, qty in aggregated.items(): text_msg += f"🔹 {p_name}: {qty}\n"
    text_msg += f"\n📍 Do'konlar: {', '.join(set(shop_list))}\n💰 Jami: {total_summa:,.0f} so'm"
    
    expeditors = conn.execute("SELECT tg_id FROM users WHERE role LIKE '%ekspeditor%'").fetchall()
    conn.close()
    for exp in expeditors:
        try:
            bot.send_message(exp['tg_id'], text_msg, parse_mode='Markdown')
            with open(file_path, 'rb') as f: bot.send_document(exp['tg_id'], f)
        except Exception as e: print(f"Xatolik ekspeditorga yuborishda: {e}")
    return send_file(file_path, as_attachment=True)

# HTML Shablon (to'liq tugatilgan va xatolarsiz)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Xoji Aka ERP — Premium Boshqaruv</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root { --bg-main: #f8fafc; --sidebar-bg: #0f172a; --primary-color: #3b82f6; }
        body { background-color: var(--bg-main); font-family: 'Inter', sans-serif; color: #1e293b; }
        .sidebar { width: 150px; background: var(--sidebar-bg); min-height: 100vh; box-shadow: 4px 0 20px rgba(0,0,0,0.05); }
        .sidebar .nav-link { text-align: center; padding: 12px 6px; color: #94a3b8; font-size: 12px; font-weight: 500; border-radius: 10px; margin: 6px 8px; transition: all 0.25s ease; }
        .sidebar .nav-link:hover { background: rgba(255, 255, 255, 0.08); color: #ffffff; }
        .sidebar .nav-link.active { background: var(--primary-color); color: #ffffff; font-weight: 600; }
        .sidebar .nav-link i { font-size: 20px; display: block; margin-bottom: 4px; }
        .brand-logo-container { text-align: center; padding: 16px 5px; border-bottom: 1px solid rgba(255, 255, 255, 0.1); margin-bottom: 10px; }
        .brand-xa { font-family: 'Georgia', serif; font-weight: 900; font-size: 34px; line-height: 0.8; color: #ffffff; font-style: italic; }
        .brand-line { height: 3px; background-color: #ef4444; width: 60px; margin: 6px auto; border-radius: 2px; }
        .brand-name { font-weight: 800; font-size: 9px; letter-spacing: 2px; text-transform: uppercase; color: #cbd5e1; }
        .card-custom { background: #ffffff; border-radius: 16px; padding: 20px; margin-bottom: 20px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02); }
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
        <button class="nav-link active" data-bs-toggle="pill" data-bs-target="#tab-orders" type="button"><i class="bi bi-cart-fill"></i>Buyurtmalar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-shops" type="button"><i class="bi bi-shop"></i>Do'konlar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-products" type="button"><i class="bi bi-box-seam"></i>Mahsulotlar</button>
        <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-users" type="button"><i class="bi bi-people-fill"></i>Xodimlar</button>
        <a href="/logout" class="nav-link text-danger mt-auto"><i class="bi bi-box-arrow-right"></i>Chiqish</a>
    </div>
    
    <div class="p-4" style="flex-grow: 1;">
        <div class="tab-content">
            <!-- BUYURTMALAR TAB -->
            <div class="tab-pane fade show active" id="tab-orders">
                <div class="card-custom">
                    <h5 class="mb-3">📅 Buyurtmalarni Filtrlash</h5>
                    <form method="GET" action="/" class="row g-2 mb-3">
                        <div class="col-md-4"><input type="date" name="start_date" class="form-control" value="{{ start_date }}"></div>
                        <div class="col-md-4"><input type="date" name="end_date" class="form-control" value="{{ end_date }}"></div>
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
                            <button type="submit" class="btn btn-success">🚚 Tanlanganlarni Svodka Qilib Yuborish</button>
                        </div>
                        <div class="table-responsive">
                            <table class="table table-hover align-middle">
                                <thead class="table-light">
                                    <tr><th><input type="checkbox" id="checkAll"></th><th>ID</th><th>Do'kon</th><th>Agent</th><th>Summa</th><th>Izoh</th><th>Sana</th><th>Holat</th></tr>
                                </thead>
                                <tbody>
                                    {% for o in orders %}
                                    <tr>
                                        <td><input type="checkbox" name="order_ids" value="{{ o.id }}" class="order-check"></td>
                                        <td><b>#{{ o.id }}</b></td>
                                        <td>{{ o.shop_name }}</td>
                                        <td>{{ o.agent_name }}</td>
                                        <td>{{ "{:,.0f}".format(o.total_sum) }} so'm</td>
                                        <td><small class="text-muted">{{ o.comment }}</small></td>
                                        <td><small class="text-muted">{{ o.date }}</small></td>
                                        <td><span class="badge bg-info text-dark">{{ o.status }}</span></td>
                                    </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </form>
                </div>
            </div>

            <!-- DO'KONLAR TAB -->
            <div class="tab-pane fade" id="tab-shops">
                <div class="card-custom">
                    <h5 class="mb-3">🏪 Do'konlar va Qarzlar</h5>
                    <div class="table-responsive">
                        <table class="table table-bordered">
                            <thead class="table-light"><tr><th>ID</th><th>Do'kon nomi</th><th>Telefon</th><th>Hudud</th><th>Qarz</th></tr></thead>
                            <tbody>
                                {% for s in shops %}
                                <tr>
                                    <td>{{ s.id }}</td>
                                    <td><b>{{ s.name }}</b></td>
                                    <td>{{ s.phone }}</td>
                                    <td>{{ s.region }}</td>
                                    <td class="text-danger fw-bold">{{ "{:,.0f}".format(s.debt) }} so'm</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- MAHSULOTLAR TAB -->
            <div class="tab-pane fade" id="tab-products">
                <div class="card-custom">
                    <h5 class="mb-3">📦 Mahsulotlar va Sklad Qoldiqlari</h5>
                    <div class="table-responsive">
                        <table class="table table-bordered">
                            <thead class="table-light"><tr><th>Mahsulot</th><th>Kategoriya</th><th>Qoldiq</th><th>Optom Narx</th><th>Chakana Narx</th></tr></thead>
                            <tbody>
                                {% for p in products %}
                                <tr>
                                    <td><b>{{ p.name }}</b></td>
                                    <td>{{ p.category }}</td>
                                    <td><span class="badge bg-success">{{ p.stock }}</span></td>
                                    <td>{{ "{:,.0f}".format(p.optom_price) }} so'm</td>
                                    <td>{{ "{:,.0f}".format(p.chakana_price) }} so'm</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- XODIMLAR TAB -->
            <div class="tab-pane fade" id="tab-users">
                <div class="card-custom">
                    <h5 class="mb-3">👥 Foydalanuvchilar va Xodimlar</h5>
                    <table class="table table-bordered">
                        <thead class="table-light"><tr><th>TG ID</th><th>Ism</th><th>Telefon</th><th>Rol</th><th>Amal</th></tr></thead>
                        <tbody>
                            {% for u in users %}
                            <tr>
                                <td>{{ u.tg_id }}</td>
                                <td>{{ u.name }}</td>
                                <td>{{ u.phone }}</td>
                                <td><span class="badge bg-secondary">{{ u.role }}</span></td>
                                <td>
                                    {% if u.role == 'pending' %}
                                    <form method="POST" action="/approve_agent/{{ u.tg_id }}" style="display:inline;"><button class="btn btn-sm btn-success">Tasdiqlash</button></form>
                                    {% endif %}
                                    <form method="POST" action="/delete_agent/{{ u.tg_id }}" style="display:inline;"><button class="btn btn-sm btn-danger">O'chirish</button></form>
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
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    document.getElementById('checkAll')?.addEventListener('change', function() {
        let checkboxes = document.querySelectorAll('.order-check');
        checkboxes.forEach(cb => cb.checked = this.checked);
    });
</script>
</body>
</html>
"""
