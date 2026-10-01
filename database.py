import sqlite3

DB_NAME = 'xoji_aka_factory.db'
ADMIN_ID = 6851851908

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_web_db(admin_id=ADMIN_ID):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
        tg_id INTEGER PRIMARY KEY, 
        name TEXT, 
        phone TEXT, 
        role TEXT DEFAULT 'pending'
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        reason TEXT, 
        amount REAL, 
        date TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS incomes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        source TEXT, 
        amount REAL, 
        date TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS product_incomes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        product_name TEXT, 
        qty REAL, 
        cost_price REAL, 
        date TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        shop_name TEXT, 
        agent_name TEXT, 
        items_text TEXT, 
        total_sum REAL, 
        discount REAL DEFAULT 0, 
        status TEXT, 
        date TEXT, 
        price_type TEXT, 
        comment TEXT DEFAULT ''
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS order_status_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        order_id INTEGER, 
        status TEXT, 
        changed_at TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        name TEXT UNIQUE, 
        category TEXT DEFAULT 'Boshqa', 
        stock REAL DEFAULT 0, 
        cost_price REAL DEFAULT 0, 
        optom_price REAL DEFAULT 0, 
        chakana_price REAL DEFAULT 0
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS shops (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        name TEXT UNIQUE, 
        phone TEXT, 
        debt REAL DEFAULT 0, 
        visit_days TEXT, 
        region TEXT DEFAULT '', 
        landmark TEXT DEFAULT '', 
        inventory TEXT DEFAULT ''
    )''')

    migrations = [
        ("ALTER TABLE products ADD COLUMN category TEXT DEFAULT 'Boshqa'", "category"),
        ("ALTER TABLE products ADD COLUMN cost_price REAL DEFAULT 0", "cost_price"),
        ("ALTER TABLE products ADD COLUMN optom_price REAL DEFAULT 0", "optom_price"),
        ("ALTER TABLE products ADD COLUMN chakana_price REAL DEFAULT 0", "chakana_price"),
        ("ALTER TABLE products ADD COLUMN stock REAL DEFAULT 0", "stock"),
        ("ALTER TABLE shops ADD COLUMN visit_days TEXT DEFAULT ''", "visit_days"),
        ("ALTER TABLE shops ADD COLUMN region TEXT DEFAULT ''", "region"),
        ("ALTER TABLE shops ADD COLUMN landmark TEXT DEFAULT ''", "landmark"),
        ("ALTER TABLE shops ADD COLUMN inventory TEXT DEFAULT ''", "inventory"),
        ("ALTER TABLE orders ADD COLUMN discount REAL DEFAULT 0", "discount"),
        ("ALTER TABLE orders ADD COLUMN price_type TEXT", "price_type"),
        ("ALTER TABLE orders ADD COLUMN comment TEXT DEFAULT ''", "comment"),
    ]

    for query, col in migrations:
        try:
            cursor.execute(query)
        except:
            pass

    cursor.execute("INSERT OR REPLACE INTO users (tg_id, name, phone, role) VALUES (?, 'Admin', '', 'admin')", (admin_id,))
    conn.commit()
    conn.close()
