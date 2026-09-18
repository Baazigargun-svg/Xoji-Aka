from datetime import datetime
from flask import Flask, redirect, render_template_string, request, url_for, send_file, session
import sqlite3
import pandas as pd
import io

app = Flask(__name__)
DB_NAME = 'xoji_aka_factory.db'
app.secret_key = 'xoji_aka_maxfiy_kalit_2026'


def get_db_connection():
  conn = sqlite3.connect(DB_NAME)
  conn.row_factory = sqlite3.Row
  return conn


def init_web_db():
  conn = get_db_connection()
  cursor = conn.cursor()

  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS expenses 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, reason TEXT, amount REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS incomes 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, source TEXT, amount REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS product_incomes 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT, qty REAL, cost_price REAL, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS orders 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, shop_name TEXT, agent_name TEXT, items_text TEXT, total_sum REAL, discount REAL DEFAULT 0, status TEXT, date TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS order_status_history 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER, status TEXT, changed_at TEXT)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS products 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, category TEXT, stock REAL, cost_price REAL, optom_price REAL)'''
  )
  cursor.execute(
      '''CREATE TABLE IF NOT EXISTS shops 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, phone TEXT, debt REAL DEFAULT 0, visit_days TEXT)'''
  )

  try:
    cursor.execute("ALTER TABLE products ADD COLUMN category TEXT DEFAULT 'Boshqa'")
  except:
    pass
  try:
    cursor.execute('ALTER TABLE products ADD COLUMN cost_price REAL DEFAULT 0')
  except:
    pass
  try:
    cursor.execute('ALTER TABLE products ADD COLUMN optom_price REAL DEFAULT 0')
  except:
    pass
  try:
    cursor.execute('ALTER TABLE products ADD COLUMN stock REAL DEFAULT 0')
  except:
    pass
  try:
    cursor.execute("ALTER TABLE shops ADD COLUMN visit_days TEXT DEFAULT ''")
  except:
    pass
  try:
    cursor.execute('ALTER TABLE orders ADD COLUMN discount REAL DEFAULT 0')
  except:
    pass

  conn.commit()
  conn.close()


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
        :root {
            --bg-main: #f8fafc;
            --sidebar-bg: #0f172a;
            --primary-color: #3b82f6;
        }
        body { background-color: var(--bg-main); font-family: 'Inter', system-ui, -apple-system, sans-serif; color: #1e293b; }
        
        .sidebar { width: 150px; background: var(--sidebar-bg); min-height: 100vh; box-shadow: 4px 0 20px rgba(0,0,0,0.05); }
        .sidebar .nav-link { 
            text-align: center; padding: 12px 6px; color: #94a3b8; font-size: 12px; font-weight: 500; 
            border-radius: 10px; margin: 6px 8px; transition: all 0.25s ease;
        }
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
                <!-- BOSH PANEL -->
                <div class="tab-pane fade show active" id="tab-dashboard">
                    <div class="row g-4 mb-4">
                        <div class="col-md-4">
                            <div class="stat-box stat-blue shadow-sm">
                                <div class="small text-white-50">Tanlangan Sana Tushumi</div>
                                <h2 class="fw-bold mt-1 mb-0">{{ "{:,.0f}".format(daily_sum) }} <span class="fs-6">so'm</span></h2>
                                <i class="bi bi-currency-exchange position-absolute bottom-0 end-0 p-3 fs-1 opacity-25"></i>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="stat-box stat-green shadow-sm">
                                <div class="small text-white-50">Kassadagi Naqd Pul</div>
                                <h2 class="fw-bold mt-1 mb-0">{{ "{:,.0f}".format(kassa_balance) }} <span class="fs-6">so'm</span></h2>
                                <i class="bi bi-cash-stack position-absolute bottom-0 end-0 p-3 fs-1 opacity-25"></i>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="stat-box stat-red shadow-sm">
                                <div class="small text-white-50">Umumiy Nasiya (Qarzlar)</div>
                                <h2 class="fw-bold mt-1 mb-0">{{ "{:,.0f}".format(total_debt) }} <span class="fs-6">so'm</span></h2>
                                <i class="bi bi-journal-x position-absolute bottom-0 end-0 p-3 fs-1 opacity-25"></i>
                            </div>
                        </div>
                    </div>

                    <div class="row g-4">
                        <div class="col-md-5">
                            <div class="card-glass p-4 h-100 d-flex flex-column align-items-center justify-content-center">
                                <h6 class="fw-bold mb-3 text-secondary w-100"><i class="bi bi-pie-chart-fill text-primary me-2"></i>Mahsulot Kategoriyalari Ulushi</h6>
                                <div style="width: 260px; height: 260px;">
                                    <canvas id="categoryDonutChart"></canvas>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-7">
                            <div class="card-glass p-4 h-100">
                                <h6 class="fw-bold mb-3 text-secondary"><i class="bi bi-table me-2"></i>Kategoriyalar bo'yicha hisobot jadvali</h6>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle">
                                        <thead>
                                            <tr>
                                                <th>Kategoriya</th>
                                                <th>Ulush (%)</th>
                                                <th>Summa (so'm)</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {% for row in cat_table_data %}
                                            <tr>
                                                <td>
                                                    <span class="d-inline-block rounded-circle me-2" style="width: 10px; height: 10px; background-color: {{ row['color'] }};"></span>
                                                    <b>{{ row['name'] }}</b>
                                                </td>
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

                <!-- BUYURTMALAR -->
                <div class="tab-pane fade" id="tab-orders">
                    <div class="row g-4">
                        <div class="col-md-5">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold text-primary mb-3"><i class="bi bi-cart-plus me-2"></i>Yangi Buyurtma Kiritish</h5>
                                <form action="/add_order" method="POST">
                                    <div class="mb-3">
                                        <label class="form-label small fw-bold">Do'kon:</label>
                                        <select name="shop_name" class="form-select" required>
                                            <option value="">-- Do'konni tanlang --</option>
                                            {% for s in shops %}<option value="{{ s['name'] }}">{{ s['name'] }}</option>{% endfor %}
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

                                    <hr>
                                    <label class="form-label small fw-bold text-secondary mb-2"><i class="bi bi-basket me-1"></i>Mahsulotlar Savatchasi:</label>
                                    
                                    <div id="order-items-container">
                                        <div class="row g-2 mb-2 order-item-row align-items-center">
                                            <div class="col-7">
                                                <select name="product_name" class="form-select form-select-sm" required>
                                                    <option value="">-- Tovar tanlang --</option>
                                                    {% for p in products %}<option value="{{ p['name'] }}">{{ p['name'] }} (Sklad: {{ p['stock'] }}) - {{ "{:,.0f}".format(p['optom_price']) }} so'm</option>{% endfor %}
                                                </select>
                                            </div>
                                            <div class="col-4">
                                                <input type="number" step="any" name="qty" class="form-control form-control-sm" placeholder="Miqdor" value="1" required>
                                            </div>
                                            <div class="col-1 text-center">
                                                <button type="button" class="btn btn-sm text-danger p-0" onclick="removeRow(this)" title="O'chirish"><i class="bi bi-trash fs-6"></i></button>
                                            </div>
                                        </div>
                                    </div>

                                    <button type="button" class="btn btn-sm btn-outline-secondary w-100 mb-3 border-dashed" onclick="addItemRow()">
                                        <i class="bi bi-plus-circle me-1"></i> Yana mahsulot qo'shish
                                    </button>

                                    <button type="submit" class="btn btn-primary w-100 py-2 fw-bold shadow-sm">Buyurtmani Tasdiqlash</button>
                                </form>
                            </div>
                        </div>

                        <div class="col-md-7">
                            <div class="card-glass p-4 mb-3">
                                <h6 class="fw-bold mb-2"><i class="bi bi-printer me-1"></i>Ko'p zakazlarni chiqarish</h6>
                                <form action="/print_nakladnoy" method="GET" target="_blank" class="d-flex gap-2">
                                    <input type="text" name="ids" class="form-control form-control-sm" placeholder="Zakaz ID lari (Masalan: 1,2,3)" required>
                                    <button type="submit" class="btn btn-sm btn-dark text-nowrap">Chop etish</button>
                                </form>
                            </div>

                            <div class="card-glass p-4">
                                <div class="d-flex justify-content-between align-items-center mb-3">
                                    <h5 class="fw-bold mb-0"><i class="bi bi-list-check me-2"></i>Barcha Buyurtmalar</h5>
                                    <input type="text" id="orderSearch" class="form-control form-control-sm" placeholder="Buyurtma qidirish..." style="width: 200px;" onkeyup="filterOrders()">
                                </div>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover align-middle" id="ordersTable">
                                        <thead><tr><th>ID / Vaqt</th><th>Do'kon</th><th>Tafsilot</th><th>Status & Tarix</th><th>Amallar</th></tr></thead>
                                        <tbody>
                                            {% for o in orders %}
                                            <tr>
                                                <td><b>#{{ o['id'] }}</b><br><small class="text-muted"><i class="bi bi-clock me-1"></i>{{ o['date'] }}</small></td>
                                                <td><b>{{ o['shop_name'] }}</b><br><small class="text-muted">Agent: {{ o['agent_name'] }}</small></td>
                                                <td>
                                                    <code style="font-size: 11px; white-space: pre-line;" class="text-dark d-block">{{ o['items_text'] }}</code>
                                                    {% if o['discount'] and o['discount'] > 0 %}<small class="text-danger">Skidka: -{{ "{:,.0f}".format(o['discount']) }} so'm</small><br>{% endif %}
                                                    <div class="fw-bold text-primary mt-1">Jami: {{ "{:,.0f}".format(o['total_sum']) }} so'm</div>
                                                </td>
                                                <td>
                                                    <span class="badge {% if o['status'] == 'Bekor' %}bg-danger{% elif o['status'] == 'Yetkazildi' %}bg-success{% else %}bg-secondary{% endif %} mb-1">{{ o['status'] }}</span>
                                                    <div style="font-size: 11px; color: #64748b;" class="mt-1">
                                                        {% if o['history'] %}
                                                            {% for h in o['history'] %}
                                                                <div><b>{{ h['status'] }}:</b> <span class="text-muted">{{ h['changed_at'] }}</span></div>
                                                            {% endfor %}
                                                        {% else %}
                                                            <div>Yaratildi: {{ o['date'] }}</div>
                                                        {% endif %}
                                                    </div>
                                                </td>
                                                <td>
                                                    <form action="/update_status/{{ o['id'] }}" method="POST" class="d-flex gap-1 mb-1">
                                                        <select name="status" class="form-select form-select-sm" style="font-size: 11px; width: 90px;">
                                                            <option value="Yangi" {% if o['status'] == 'Yangi' %}selected{% endif %}>Yangi</option>
                                                            <option value="Otgruzka" {% if o['status'] == 'Otgruzka' %}selected{% endif %}>Otgruzka</option>
                                                            <option value="Yetkazildi" {% if o['status'] == 'Yetkazildi' %}selected{% endif %}>Yetkazildi</option>
                                                            <option value="Bekor" {% if o['status'] == 'Bekor' %}selected{% endif %}>Bekor</option>
                                                        </select>
                                                        <button type="submit" class="btn btn-sm btn-success px-2" title="Statusni yangilash"><i class="bi bi-check"></i></button>
                                                    </form>
                                                    <a href="/print_nakladnoy/{{ o['id'] }}" target="_blank" class="btn btn-sm btn-outline-dark py-0 w-100 mb-1" style="font-size: 11px;"><i class="bi bi-printer me-1"></i>Nakladnoy (A4)</a>
                                                    <button class="btn btn-sm btn-outline-primary py-0 w-100" style="font-size: 11px;" data-bs-toggle="modal" data-bs-target="#editModal{{ o['id'] }}">Tahrirlash</button>

                                                    <div class="modal fade" id="editModal{{ o['id'] }}" tabindex="-1">
                                                        <div class="modal-dialog">
                                                            <div class="modal-content">
                                                                <form action="/edit_order/{{ o['id'] }}" method="POST">
                                                                    <div class="modal-header">
                                                                        <h5 class="modal-title fs-6 fw-bold">Buyurtmani Tahrirlash (#{{ o['id'] }})</h5>
                                                                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                                                                    </div>
                                                                    <div class="modal-body text-start">
                                                                        <div class="mb-3">
                                                                            <label class="form-label small fw-bold text-muted">Mahsulotlar satri:</label>
                                                                            <textarea name="items_text" class="form-control font-monospace" rows="4" required>{{ o['items_text'] }}</textarea>
                                                                        </div>
                                                                        <div class="mb-3">
                                                                            <label class="form-label small fw-bold text-muted">Chegirma (Skidka):</label>
                                                                            <input type="number" step="any" name="discount" class="form-control" value="{{ o['discount'] if o['discount'] else 0 }}">
                                                                        </div>
                                                                        <div class="mb-3">
                                                                            <label class="form-label small fw-bold text-muted">Yangi Umumiy Summa (so'm):</label>
                                                                            <input type="number" step="any" name="total_sum" class="form-control fw-bold text-primary" value="{{ o['total_sum'] }}" required>
                                                                        </div>
                                                                    </div>
                                                                    <div class="modal-footer">
                                                                        <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Yopish</button>
                                                                        <button type="submit" class="btn btn-primary btn-sm">O'zgarishlarni Saqlash</button>
                                                                    </div>
                                                                </form>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </td>
                                            </tr>
                                            {% else %}
                                            <tr><td colspan="5" class="text-center text-muted py-4">Buyurtmalar yo'q</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- SKLAD -->
                <div class="tab-pane fade" id="tab-inventory">
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
                                        <div class="mb-2"><label class="form-label small text-primary fw-bold">Yangi nom:</label><input type="text" name="new_product_name" id="new_p_name" class="form-control" autocomplete="off"></div>
                                        <div><label class="form-label small text-primary fw-bold">Kategoriya:</label><input type="text" name="new_product_category" class="form-control" placeholder="Masalan: Ichimliklar" autocomplete="off"></div>
                                    </div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Miqdor:</label><input type="number" step="any" name="qty" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Tan narxi:</label><input type="number" step="any" name="cost_price" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Optom narxi:</label><input type="number" step="any" name="optom_price" class="form-control" required></div>
                                    <button type="submit" class="btn btn-primary w-100">Omborga Saqlash</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-7">
                            <div class="card-glass p-4">
                                <div class="d-flex justify-content-between align-items-center mb-3">
                                    <h5 class="fw-bold mb-0"><i class="bi bi-boxes me-2"></i>Ombordagi Qoldiqlar</h5>
                                    <input type="text" id="prodSearch" class="form-control form-control-sm" placeholder="Tovar qidirish..." style="width: 200px;" onkeyup="filterProducts()">
                                </div>
                                <div class="table-responsive">
                                    <table class="table table-custom table-hover" id="productsTable">
                                        <thead><tr><th>Mahsulot</th><th>Kategoriya</th><th>Qoldiq</th><th>Tan narx</th><th>Optom narx</th><th>Amallar</th></tr></thead>
                                        <tbody>
                                            {% for p in products %}
                                            <tr>
                                                <td><b>{{ p['name'] }}</b></td>
                                                <td><span class="badge bg-light text-dark border">{{ p['category'] if 'category' in p.keys() and p['category'] else 'Boshqa' }}</span></td>
                                                <td><span class="badge bg-success">{{ p['stock'] if 'stock' in p.keys() else 0 }}</span></td>
                                                <td>{{ "{:,.0f}".format(p['cost_price'] if 'cost_price' in p.keys() and p['cost_price'] else 0) }}</td>
                                                <td>{{ "{:,.0f}".format(p['optom_price'] if 'optom_price' in p.keys() and p['optom_price'] else 0) }}</td>
                                                <td>
                                                    <form action="/delete_product/{{ p['id'] }}" method="POST" onsubmit="return confirm('Haqiqatan ham bu mahsulotni o\\'chirmoqchimisiz?');" style="display:inline;">
                                                        <button type="submit" class="btn btn-sm btn-outline-danger py-0" style="font-size: 11px;"><i class="bi bi-trash"></i></button>
                                                    </form>
                                                </td>
                                            </tr>
                                            {% else %}
                                            <tr><td colspan="6" class="text-center text-muted py-3">Mahsulotlar mavjud emas</td></tr>
                                            {% endfor %}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- DO'KONLAR -->
                <div class="tab-pane fade" id="tab-clients">
                    <div class="row g-4">
                        <div class="col-md-4">
                            <div class="card-glass p-4">
                                <h5 class="fw-bold mb-3"><i class="bi bi-shop-window me-2 text-success"></i>Yangi Do'kon</h5>
                                <form action="/add_shop" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">Nomi:</label><input type="text" name="name" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Telefon:</label><input type="text" name="phone" class="form-control" required></div>
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
                                    <h5 class="fw-bold mb-0"><i class="bi bi-people me-2"></i>Do'konlar & Qarzni To'lash</h5>
                                    <input type="text" id="shopSearch" class="form-control form-control-sm" placeholder="Do'kon qidirish..." style="width: 200px;" onkeyup="filterShops()">
                                </div>
                                <table class="table table-custom table-hover" id="shopsTable">
                                    <thead><tr><th>Do'kon</th><th>Tel</th><th>Tashrif kuni</th><th>Qarzdorlik</th><th>To'lov</th></tr></thead>
                                    <tbody>
                                        {% for s in shops %}
                                        <tr>
                                            <td><b>{{ s['name'] }}</b></td>
                                            <td>{{ s['phone'] }}</td>
                                            <td><span class="badge bg-light text-dark border">{{ s['visit_days'] if s['visit_days'] else 'Belgilanmagan' }}</span></td>
                                            <td><b class="text-danger">{{ "{:,.0f}".format(s['debt']) }} so'm</b></td>
                                            <td>
                                                <form action="/pay_debt" method="POST" class="d-flex gap-1">
                                                    <input type="hidden" name="shop_id" value="{{ s['id'] }}">
                                                    <input type="number" name="amount" class="form-control form-control-sm" placeholder="Summa" required style="width: 100px;">
                                                    <button type="submit" class="btn btn-sm btn-success">Prixod</button>
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

                <!-- KASSA -->
                <div class="tab-pane fade" id="tab-kassa">
                    <div class="row g-4">
                        <div class="col-md-6">
                            <div class="card-glass p-4 border-start border-success border-4">
                                <h5 class="fw-bold text-success mb-3">Kirim qilish</h5>
                                <form action="/add_income" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">Manba:</label><input type="text" name="source" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Summa:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-success w-100">Kirim</button>
                                </form>
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="card-glass p-4 border-start border-danger border-4">
                                <h5 class="fw-bold text-danger mb-3">Chiqim (Rasxod)</h5>
                                <form action="/add_expense" method="POST">
                                    <div class="mb-3"><label class="form-label small fw-bold">Sabab:</label><input type="text" name="reason" class="form-control" required></div>
                                    <div class="mb-3"><label class="form-label small fw-bold">Summa:</label><input type="number" name="amount" class="form-control" required></div>
                                    <button type="submit" class="btn btn-danger w-100">Chiqim</button>
                                </form>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- HISOBOT -->
                <div class="tab-pane fade" id="tab-reports">
                    <div class="card-glass p-4 mb-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-file-earmark-bar-graph me-2 text-primary"></i>Moliyaviy Hisobot</h5>
                        <div class="row g-3">
                            <div class="col-md-3 bg-light p-3 rounded"><small class="text-muted">Umumiy Savdo</small><h4 class="fw-bold text-primary mb-0">{{ "{:,.0f}".format(total_revenue) }} so'm</h4></div>
                            <div class="col-md-3 bg-light p-3 rounded"><small class="text-muted">Tannarx</small><h4 class="fw-bold text-secondary mb-0">{{ "{:,.0f}".format(total_cost) }} so'm</h4></div>
                            <div class="col-md-3 bg-light p-3 rounded"><small class="text-muted">Rasxod</small><h4 class="fw-bold text-danger mb-0">{{ "{:,.0f}".format(total_expense) }} so'm</h4></div>
                            <div class="col-md-3 bg-success bg-opacity-10 p-3 rounded border border-success"><small class="text-success fw-bold">Sof Foyda</small><h4 class="fw-bold text-success mb-0">{{ "{:,.0f}".format(net_profit) }} so'm</h4></div>
                        </div>
                    </div>

                    <div class="card-glass p-4">
                        <h5 class="fw-bold mb-3"><i class="bi bi-person-badge me-2 text-primary"></i>Agentlar Kesimida Savdo Statistikasi</h5>
                        <div class="table-responsive">
                            <table class="table table-custom table-hover">
                               <thead><tr><th>Agent / Operator</th><th>Buyurtmalar Soni</th><th>Umumiy Savdo Summasi</th></tr></thead>
                               <tbody>
                                   {% for ag in agent_stats %}
                                   <tr>
                                       <td><b>{{ ag['agent_name'] }}</b></td>
                                       <td><span class="badge bg-secondary">{{ ag['count'] }} ta</span></td>
                                       <td><b class="text-primary">{{ "{:,.0f}".format(ag['sum']) }} so'm</b></td>
                                   </tr>
                                   {% else %}
                                   <tr><td colspan="3" class="text-center text-muted">Ma'lumotlar yo'q</td></tr>
                                   {% endfor %}
                               </tbody>
                            </table>
                        </div>
                    </div>
                </div>

            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    function toggleNewProd() {
        const sel = document.getElementById('p_sel');
        const div = document.getElementById('new_prod_div');
        const input = document.getElementById('new_p_name');
        if(sel.value === 'NEW') { 
            div.classList.remove('d-none'); 
            input.required = true; 
        } else { 
            div.classList.add('d-none'); 
            input.required = false; 
            input.value = '';
        }
    }

    function toggleDay(el, day) {
        el.classList.toggle('selected');
        let selectedDays = [];
        document.querySelectorAll('.day-badge.selected').forEach(badge => {
            selectedDays.push(badge.innerText);
        });
        document.getElementById('visit_days_input').value = selectedDays.join(', ');
    }

    function addItemRow() {
        const container = document.getElementById('order-items-container');
        const firstRow = container.querySelector('.order-item-row');
        const newRow = firstRow.cloneNode(true);
        newRow.querySelector('input[name="qty"]').value = '1';
        newRow.querySelector('select[name="product_name"]').value = '';
        container.appendChild(newRow);
    }

    function removeRow(btn) {
        const container = document.getElementById('order-items-container');
        if (container.querySelectorAll('.order-item-row').length > 1) {
            btn.closest('.order-item-row').remove();
        } else {
            alert("Kamida bitta mahsulot bo'lishi kerak!");
        }
    }

    function filterOrders() {
        let input = document.getElementById('orderSearch').value.toLowerCase();
        let rows = document.querySelectorAll('#ordersTable tbody tr');
        rows.forEach(row => {
            let text = row.innerText.toLowerCase();
            row.style.display = text.includes(input) ? '' : 'none';
        });
    }

    function filterProducts() {
        let input = document.getElementById('prodSearch').value.toLowerCase();
        let rows = document.querySelectorAll('#productsTable tbody tr');
        rows.forEach(row => {
            let text = row.innerText.toLowerCase();
            row.style.display = text.includes(input) ? '' : 'none';
        });
    }

    function filterShops() {
        let input = document.getElementById('shopSearch').value.toLowerCase();
        let rows = document.querySelectorAll('#shopsTable tbody tr');
        rows.forEach(row => {
            let text = row.innerText.toLowerCase();
            row.style.display = text.includes(input) ? '' : 'none';
        });
    }

    document.addEventListener("DOMContentLoaded", function() {
        const ctx = document.getElementById('categoryDonutChart').getContext('2d');
        const chartLabels = {{ chart_labels | safe }};
        const chartData = {{ chart_data | safe }};
        const chartColors = {{ chart_colors | safe }};

        new Chart(ctx, {
            type: 'doughnut',
            data: { 
                labels: chartLabels, 
                datasets: [{ 
                    data: chartData, 
                    backgroundColor: chartColors, 
                    borderWidth: 2, 
                    borderColor: '#ffffff' 
                }] 
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { 
                    legend: { 
                        position: 'bottom', 
                        labels: { boxWidth: 12, font: { size: 11 } } 
                    } 
                }, 
                cutout: '65%' 
            }
        });
    });
</script>
</body>
</html>
"""

NAKLADNOY_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <title>Nakladnoylar</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        @page { size: A4 portrait; margin: 5mm; }
        body { background: #fff; font-family: sans-serif; font-size: 10px; color: #000; margin: 0; padding: 0; }
        
        /* Har bir zakaz bloki */
        .order-block {
            width: 100%;
            page-break-inside: avoid;
            margin-bottom: 8mm;
            padding-bottom: 5mm;
            border-bottom: 1px dashed #666;
        }

        /* 1- va 2-nusxani yonma-yon joylashtirish */
        .copies-wrapper {
            display: flex;
            justify-content: space-between;
            gap: 4mm;
        }

        /* Bitta nakladnoy qutisi (A4 enining yarmidan kamroq) */
        .nakladnoy-box {
            width: 49%;
            border: 1px solid #000;
            padding: 4px;
            box-sizing: border-box;
            background: #fff;
        }

        table { width: 100%; border-collapse: collapse; margin-top: 3px; }
        th, td { border: 1px solid #000; padding: 2px 3px; font-size: 9.5px; text-align: left; }
        th { background: #f8f9fa; text-align: center; font-weight: bold; }
        
        .info-table td { border: 1px solid #000; padding: 2px 3px; font-size: 9.5px; }
        .signatures { margin-top: 5px; font-size: 9.5px; font-weight: bold; }

        @media print {
            body { padding: 0; margin: 0; }
            .no-print { display: none; }
            .order-block:last-child { border-bottom: none; }
        }
    </style>
</head>
<body onload="window.print()">
    <div class="no-print text-center py-2 bg-light border-bottom mb-3">
        <button onclick="window.print()" class="btn btn-primary btn-sm">Chop etish (Print)</button>
        <a href="/" class="btn btn-secondary btn-sm">Asosiy sahifaga qaytish</a>
    </div>

    {% for order, items_parsed in orders_data %}
    <div class="order-block">
        <div class="copies-wrapper">
            
            <!-- 1-NUSXA (CHAPDA) -->
            <div class="nakladnoy-box">
                <table class="info-table">
                    <tr>
                        <td colspan="2"><b>Buyurtmachi:</b> {{ order['shop_name'] }}</td>
                        <td><b>Tel:</b> {{ order.get('phone', '50714194') }}</td>
                    </tr>
                    <tr>
                        <td colspan="2"><b>Murojaat:</b> +998935075540</td>
                        <td><b>Sana:</b> {{ order['date'] }}</td>
                    </tr>
                </table>
                
                <table>
                    <thead>
                        <tr>
                            <th>Mahsulot nomi</th>
                            <th style="width: 35px; text-align: center;">Soni</th>
                            <th style="width: 50px; text-align: center;">narxi</th>
                            <th style="width: 60px; text-align: center;">summasi</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for item in items_parsed %}
                        <tr>
                            <td>{{ item.name }}</td>
                            <td style="text-align: center;">{{ item.qty }}</td>
                            <td style="text-align: right;">{{ "{:,.0f}".format(item.sum / float(item.qty) if item.qty and float(item.qty)>0 else 0) }}</td>
                            <td style="text-align: right;">{{ "{:,.0f}".format(item.sum) }}</td>
                        </tr>
                        {% endfor %}
                        <tr>
                            <td colspan="3" style="text-align: right; font-weight: bold;">Jami:</td>
                            <td style="text-align: right; font-weight: bold;">{{ "{:,.0f}".format(order['total_sum']) }}</td>
                        </tr>
                    </tbody>
                </table>
                
                <div class="signatures">
                    Berildi: ________________________
                </div>
            </div>

            <!-- 2-NUSXA (O'NGDA) -->
            <div class="nakladnoy-box">
                <table class="info-table">
                    <tr>
                        <td colspan="2"><b>Buyurtmachi:</b> {{ order['shop_name'] }}</td>
                        <td><b>Tel:</b> {{ order.get('phone', '50714194') }}</td>
                    </tr>
                    <tr>
                        <td colspan="2"><b>Murojaat:</b> +998935075540</td>
                        <td><b>Sana:</b> {{ order['date'] }}</td>
                    </tr>
                </table>
                
                <table>
                    <thead>
                        <tr>
                            <th>Mahsulot nomi</th>
                            <th style="width: 35px; text-align: center;">Soni</th>
                            <th style="width: 50px; text-align: center;">narxi</th>
                            <th style="width: 60px; text-align: center;">summasi</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for item in items_parsed %}
                        <tr>
                            <td>{{ item.name }}</td>
                            <td style="text-align: center;">{{ item.qty }}</td>
                            <td style="text-align: right;">{{ "{:,.0f}".format(item.sum / float(item.qty) if item.qty and float(item.qty)>0 else 0) }}</td>
                            <td style="text-align: right;">{{ "{:,.0f}".format(item.sum) }}</td>
                        </tr>
                        {% endfor %}
                        <tr>
                            <td colspan="3" style="text-align: right; font-weight: bold;">Jami:</td>
                            <td style="text-align: right; font-weight: bold;">{{ "{:,.0f}".format(order['total_sum']) }}</td>
                        </tr>
                    </tbody>
                </table>
                
                <div class="signatures">
                    Berildi: ________________________
                </div>
            </div>

        </div>
    </div>
    {% endfor %}
</body>
</html>
"""


@app.route('/')
def operator_dashboard():
  selected_date = request.args.get('filter_date', '')
  conn = get_db_connection()

  orders_raw = conn.execute('SELECT * FROM orders ORDER BY id DESC').fetchall()
  orders = []
  for o in orders_raw:
    o_dict = dict(o)
    history = conn.execute(
        'SELECT status, changed_at FROM order_status_history WHERE order_id = ? ORDER BY id ASC',
        (o['id'],)
    ).fetchall()
    o_dict['history'] = history
    orders.append(o_dict)

  products = conn.execute('SELECT * FROM products').fetchall()
  shops = conn.execute('SELECT * FROM shops').fetchall()

  total_debt = conn.execute('SELECT SUM(debt) FROM shops').fetchone()[0] or 0
  total_income = conn.execute('SELECT SUM(amount) FROM incomes').fetchone()[0] or 0
  total_expense = (
      conn.execute('SELECT SUM(amount) FROM expenses').fetchone()[0] or 0
  )
  kassa_balance = total_income - total_expense

  prod_cost_map = {}
  prod_cat_map = {}
  for p in products:
    p_keys = p.keys()
    prod_cost_map[p['name']] = (
        p['cost_price'] if 'cost_price' in p_keys and p['cost_price'] else 0
    )
    prod_cat_map[p['name']] = (
        p['category'] if 'category' in p_keys and p['category'] else 'Boshqa'
    )

  total_revenue, total_cost, daily_sum = 0, 0, 0
  cat_stats = {}
  agent_revenue_map = {}
  agent_count_map = {}

  for o in orders:
    a_name = o['agent_name'] or 'Nomaʼlum'
    if o['status'] != 'Bekor':
      order_rev = o['total_sum'] or 0
      order_cost = 0
      raw_date = str(o['date']) if o['date'] else ''
      order_date = (
          raw_date.split(' ')[0]
          if ' ' in raw_date
          else raw_date.split('T')[0]
      )

      if o['status'] == 'Yetkazildi':
        if not selected_date or order_date == selected_date:
          daily_sum += order_rev
          agent_revenue_map[a_name] = agent_revenue_map.get(a_name, 0) + order_rev
          agent_count_map[a_name] = agent_count_map.get(a_name, 0) + 1

      if o['items_text']:
        lines = [l.strip() for l in o['items_text'].split('\n') if l.strip()]
        for line in lines:
          try:
            parts = line.split('-')
            p_name = parts[0].strip()
            item_sum = 0
            if '=' in line:
              sum_str = (
                  line.split('=')[-1]
                  .replace("so'm", '')
                  .replace(' ', '')
                  .strip()
              )
              item_sum = float(sum_str)

            cost_p = prod_cost_map.get(p_name, 0)
            qty = 1.0
            if 'x' in line:
              qty_str = (
                  line.split('x')[1]
                  .split('=')[0]
                  .replace('kg', '')
                  .replace('dona', '')
                  .strip()
              )
              qty = float(qty_str)
            order_cost += qty * cost_p

            if o['status'] == 'Yetkazildi':
              if not selected_date or order_date == selected_date:
                cat = prod_cat_map.get(p_name, 'Boshqa')
                cat_stats[cat] = cat_stats.get(cat, 0) + item_sum
          except:
            pass
      if o['status'] == 'Yetkazildi':
        total_revenue += order_rev
        total_cost += order_cost

  net_profit = (total_revenue - total_cost) - total_expense

  agent_stats = []
  for ag_name in agent_revenue_map:
    agent_stats.append({
        'agent_name': ag_name,
        'count': agent_count_map.get(ag_name, 0),
        'sum': agent_revenue_map[ag_name]
    })

  if not cat_stats:
    chart_labels = ["Ma'lumot yo'q"]
    chart_data = [1]
    chart_colors = ['#e2e8f0']
  else:
    palette = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4']
    chart_labels = list(cat_stats.keys())
    chart_data = list(cat_stats.values())
    chart_colors = [
        palette[i % len(palette)] for i in range(len(chart_labels))
    ]

  cat_table_data = []
  for i, cat_name in enumerate(cat_stats.keys()):
    amt = cat_stats[cat_name]
    pct = round((amt / daily_sum * 100), 1) if daily_sum > 0 else 0
    cat_table_data.append({
        'name': cat_name,
        'amount': amt,
        'percent': pct,
        'color': palette[i % len(palette)],
    })

  conn.close()
  return render_template_string(
      HTML_TEMPLATE,
      orders=orders,
      products=products,
      shops=shops,
      daily_sum=daily_sum,
      total_debt=total_debt,
      kassa_balance=kassa_balance,
      total_income=total_income,
      total_expense=total_expense,
      total_revenue=total_revenue,
      total_cost=total_cost,
      net_profit=net_profit,
      cat_table_data=cat_table_data,
      agent_stats=agent_stats,
      filter_date=selected_date,
      chart_labels=chart_labels,
      chart_data=chart_data,
      chart_colors=chart_colors,
  )


@app.route('/print_nakladnoy')
@app.route('/print_nakladnoy/<int:order_id>')
def print_nakladnoy(order_id=None):
  ids_param = request.args.get('ids')
  if order_id:
    order_ids = [order_id]
  elif ids_param:
    order_ids = [int(i.strip()) for i in ids_param.split(',') if i.strip().isdigit()]
  else:
    return "Iltimos, zakaz ID larini ko'rsating!", 400

  conn = get_db_connection()
  orders_data = []
  
  for oid in order_ids:
    order = conn.execute('SELECT * FROM orders WHERE id = ?', (oid,)).fetchone()
    if order:
      items_parsed = []
      if order['items_text']:
        for line in order['items_text'].split('\n'):
          if not line.strip():
            continue
          try:
            parts = line.split('-')
            p_name = parts[0].strip()
            qty = "1"
            if 'x' in line:
              qty = line.split('x')[1].split('=')[0].strip()
            item_sum = 0
            if '=' in line:
              sum_str = line.split('=')[-1].replace("so'm", '').replace(',', '').strip()
              item_sum = float(sum_str)
            items_parsed.append({'name': p_name, 'qty': qty, 'sum': item_sum})
          except:
            pass
      orders_data.append((dict(order), items_parsed))
      
  conn.close()
  if not orders_data:
    return "Buyurtma(lar) topilmadi", 404

  return render_template_string(NAKLADNOY_TEMPLATE, orders_data=orders_data)


@app.route('/add_order', methods=['POST'])
def add_order():
  shop_name = request.form['shop_name']
  agent_name = request.form['agent_name']
  discount = float(request.form.get('discount', 0) or 0)
  
  product_names = request.form.getlist('product_name')
  quantities = request.form.getlist('qty')
  
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  cursor = conn.cursor()

  for i in range(len(product_names)):
    p_name = product_names[i]
    if not p_name:
      continue
    try:
      qty = float(quantities[i])
    except:
      qty = 1.0
    
    prod_db = cursor.execute('SELECT stock FROM products WHERE name = ?', (p_name,)).fetchone()
    current_stock = prod_db['stock'] if prod_db else 0
    if current_stock < qty:
      conn.close()
      return f"<script>alert('Xatolik: Omborda \"{p_name}\" yetarli emas! Mavjud qoldiq: {current_stock}'); window.history.back();</script>"

  total_sum = 0
  items_lines = []

  for i in range(len(product_names)):
    p_name = product_names[i]
    if not p_name:
      continue
    try:
      qty = float(quantities[i])
    except:
      qty = 1.0

    p_info = cursor.execute(
        'SELECT optom_price FROM products WHERE name = ?', (p_name,)
    ).fetchone()
    optom_price = p_info['optom_price'] if p_info else 0
    
    line_sum = qty * optom_price
    total_sum += line_sum
    items_lines.append(f"{p_name} - {qty}x = {line_sum:,.0f} so'm")

    cursor.execute('UPDATE products SET stock = stock - ? WHERE name = ?', (qty, p_name))

  final_sum = max(0, total_sum - discount)
  items_text = "\n".join(items_lines)

  cursor.execute(
      'INSERT INTO orders (shop_name, agent_name, items_text, total_sum, discount, status, date) VALUES (?, ?, ?, ?, ?, \'Yangi\', ?)',
      (shop_name, agent_name, items_text, final_sum, discount, today),
  )
  order_id = cursor.lastrowid

  cursor.execute(
      'INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, ?, ?)',
      (order_id, 'Yangi', today)
  )

  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/edit_order/<int:order_id>', methods=['POST'])
def edit_order(order_id):
  new_items_text = request.form['items_text']
  new_total_sum = float(request.form['total_sum'])
  new_discount = float(request.form.get('discount', 0) or 0)
  
  conn = get_db_connection()
  cursor = conn.cursor()
  old_order = cursor.execute(
      'SELECT shop_name, total_sum, status FROM orders WHERE id = ?',
      (order_id,),
  ).fetchone()
  if old_order:
    shop_name, old_sum = old_order['shop_name'], old_order['total_sum'] or 0

    if old_order['status'] == 'Yetkazildi':
      diff = new_total_sum - old_sum
      cursor.execute(
          'UPDATE shops SET debt = debt + ? WHERE name = ?', (diff, shop_name)
      )

    cursor.execute(
        'UPDATE orders SET items_text = ?, total_sum = ?, discount = ? WHERE id = ?',
        (new_items_text, new_total_sum, new_discount, order_id),
    )
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


@app.route('/pay_debt', methods=['POST'])
def pay_debt():
  shop_id, amount = request.form['shop_id'], float(request.form['amount'])
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  cursor = conn.cursor()
  shop = cursor.execute(
      'SELECT name, debt FROM shops WHERE id = ?', (shop_id,)
  ).fetchone()
  if shop:
    new_debt = max(0.0, shop['debt'] - amount)
    cursor.execute('UPDATE shops SET debt = ? WHERE id = ?', (new_debt, shop_id))
    cursor.execute(
        'INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)',
        (f"Qarz to'lovi ({shop['name']})", amount, today),
    )
    conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/add_stock', methods=['POST'])
def add_stock():
  p_select = request.form.get('product_select')
  qty = float(request.form.get('qty', 0))
  cost_price = float(request.form.get('cost_price', 0))
  optom_price = float(request.form.get('optom_price', 0))
  today = datetime.now().strftime('%Y-%m-%d %H:%M')

  conn = get_db_connection()
  cursor = conn.cursor()

  if p_select == 'NEW':
    p_name = request.form.get('new_product_name', '').strip()
    category = (
        request.form.get('new_product_category', 'Boshqa').strip() or 'Boshqa'
    )
    if p_name:
      cursor.execute(
          'INSERT OR REPLACE INTO products (name, category, stock, cost_price,'
          ' optom_price) VALUES (?, ?, ?, ?, ?)',
          (p_name, category, qty, cost_price, optom_price),
      )
  else:
    p_name = p_select
    cursor.execute(
        'UPDATE products SET stock = stock + ?, cost_price = ?, optom_price ='
        ' ? WHERE name = ?',
        (qty, cost_price, optom_price, p_name),
    )

  if p_name:
    cursor.execute(
        'INSERT INTO product_incomes (product_name, qty, cost_price, date)'
        ' VALUES (?, ?, ?, ?)',
        (p_name, qty, cost_price, today),
    )

  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/delete_product/<int:product_id>', methods=['POST'])
def delete_product(product_id):
  conn = get_db_connection()
  cursor = conn.cursor()
  cursor.execute('DELETE FROM products WHERE id = ?', (product_id,))
  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/add_shop', methods=['POST'])
def add_shop():
  name, phone = request.form['name'], request.form['phone']
  visit_days = request.form.get('visit_days', '')
  conn = get_db_connection()
  cursor = conn.cursor()
  try:
    cursor.execute(
        'INSERT INTO shops (name, phone, debt, visit_days) VALUES (?, ?, 0, ?)',
        (name, phone, visit_days),
    )
    conn.commit()
  except:
    pass
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/add_income', methods=['POST'])
def add_income():
  source, amount = request.form['source'], float(request.form['amount'])
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  cursor = conn.cursor()
  cursor.execute(
      'INSERT INTO incomes (source, amount, date) VALUES (?, ?, ?)',
      (source, amount, today),
  )
  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/add_expense', methods=['POST'])
def add_expense():
  reason, amount = request.form['reason'], float(request.form['amount'])
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  cursor = conn.cursor()
  cursor.execute(
      'INSERT INTO expenses (reason, amount, date) VALUES (?, ?, ?)',
      (reason, amount, today),
  )
  conn.commit()
  conn.close()
  return redirect(url_for('operator_dashboard'))


@app.route('/update_status/<int:order_id>', methods=['POST'])
def update_status(order_id):
  new_status = request.form['status']
  today = datetime.now().strftime('%Y-%m-%d %H:%M')
  conn = get_db_connection()
  cursor = conn.cursor()

  order = cursor.execute(
      'SELECT shop_name, total_sum, status, items_text FROM orders WHERE id = ?',
      (order_id,),
  ).fetchone()
  
  if order:
    shop_name = order['shop_name']
    total_sum = order['total_sum'] or 0
    old_status = order['status']

    if new_status == 'Yetkazildi' and old_status != 'Yetkazildi':
      cursor.execute(
          'UPDATE shops SET debt = debt + ? WHERE name = ?',
          (total_sum, shop_name),
      )
    elif old_status == 'Yetkazildi' and new_status != 'Yetkazildi':
      cursor.execute(
          'UPDATE shops SET debt = MAX(0, debt - ?) WHERE name = ?',
          (total_sum, shop_name),
      )

    if new_status == 'Bekor' and old_status != 'Bekor':
      if order['items_text']:
        for line in order['items_text'].split('\n'):
          try:
            parts = line.split('-')
            p_name = parts[0].strip()
            qty = 1.0
            if 'x' in line:
              qty = float(line.split('x')[1].split('=')[0].replace('kg', '').replace('dona', '').strip())
            cursor.execute('UPDATE products SET stock = stock + ? WHERE name = ?', (qty, p_name))
          except:
            pass
    elif old_status == 'Bekor' and new_status != 'Bekor':
      if order['items_text']:
        for line in order['items_text'].split('\n'):
          try:
            parts = line.split('-')
            p_name = parts[0].strip()
            qty = 1.0
            if 'x' in line:
              qty = float(line.split('x')[1].split('=')[0].replace('kg', '').replace('dona', '').strip())
            cursor.execute('UPDATE products SET stock = stock - ? WHERE name = ?', (qty, p_name))
          except:
            pass

    cursor.execute(
        'UPDATE orders SET status = ? WHERE id = ?', (new_status, order_id)
    )
    cursor.execute(
        'INSERT INTO order_status_history (order_id, status, changed_at) VALUES (?, ?, ?)',
        (order_id, new_status, today)
    )
    conn.commit()

  conn.close()
  return redirect(url_for('operator_dashboard'))

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
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
        
        {% if error %}
        <div class="alert alert-danger py-2 small text-center">{{ error }}</div>
        {% endif %}

        <form method="POST">
            <div class="mb-3">
                <label class="form-label small fw-bold">Login:</label>
                <input type="text" name="username" class="form-control" required autocomplete="off">
            </div>
            <div class="mb-4">
                <label class="form-label small fw-bold">Parol:</label>
                <input type="password" name="password" class="form-control" required>
            </div>
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

if __name__ == '__main__':
  init_web_db()
  app.run(host='0.0.0.0', port=5000, debug=True)
