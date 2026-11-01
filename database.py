import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'orders.db')

def get_connection():
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    # 1. 상품 및 재고 마스터 (Products / SKU)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT UNIQUE,
        product_name TEXT NOT NULL,
        option_name TEXT DEFAULT '',
        english_name TEXT DEFAULT '',
        hs_code TEXT DEFAULT '',
        unit_price REAL DEFAULT 0.0,
        stock_quantity INTEGER DEFAULT 0,
        safety_stock INTEGER DEFAULT 10,
        location TEXT DEFAULT 'A-01',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 기존 테이블 컬럼 마이그레이션 (english_name, hs_code, unit_price 추가)
    cur.execute("PRAGMA table_info(products)")
    existing_cols = [row[1] for row in cur.fetchall()]
    if 'english_name' not in existing_cols:
        cur.execute("ALTER TABLE products ADD COLUMN english_name TEXT DEFAULT ''")
    if 'hs_code' not in existing_cols:
        cur.execute("ALTER TABLE products ADD COLUMN hs_code TEXT DEFAULT ''")
    if 'unit_price' not in existing_cols:
        cur.execute("ALTER TABLE products ADD COLUMN unit_price REAL DEFAULT 0.0")


    # 2. 세트/옵션 분할 규칙 (Bundle Rules)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS bundle_rules (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rule_name TEXT,
        source_product_name TEXT NOT NULL,
        source_option_name TEXT NOT NULL,
        target_product_name TEXT NOT NULL,
        target_option_name TEXT NOT NULL,
        target_sku TEXT,
        multiplier INTEGER DEFAULT 1,
        is_active INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 3. 원본 주문 마스터 (Orders)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_no TEXT NOT NULL,
        courier TEXT DEFAULT 'CJ대한통운',
        tracking_no TEXT DEFAULT '',
        shipping_method TEXT DEFAULT '택배',
        original_product_name TEXT,
        original_option_name TEXT,
        original_quantity INTEGER DEFAULT 1,
        recipient_name TEXT,
        recipient_phone1 TEXT,
        recipient_phone2 TEXT,
        customs_id TEXT,
        postal_code TEXT,
        address TEXT,
        shipping_memo TEXT,
        status TEXT DEFAULT '주문접수',
        upload_batch TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 4. 분할된 출고 아이템 (Order Items)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS order_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        order_no TEXT NOT NULL,
        sku TEXT,
        item_product_name TEXT NOT NULL,
        item_option_name TEXT NOT NULL,
        quantity INTEGER DEFAULT 1,
        stock_deducted INTEGER DEFAULT 0,
        status TEXT DEFAULT '준비중',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE
    )
    ''')

    # 5. 입출고 재고 이력 로그 (Inventory Logs)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS inventory_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        sku TEXT,
        product_name TEXT,
        option_name TEXT,
        change_type TEXT NOT NULL,
        change_qty INTEGER NOT NULL,
        resulting_stock INTEGER NOT NULL,
        reference_no TEXT,
        memo TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 6. 택배사 및 발송자 설정 (Courier Settings)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS courier_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        courier_name TEXT DEFAULT 'CJ대한통운',
        customer_code TEXT DEFAULT '12345678',
        api_key TEXT DEFAULT 'MOCK_API_KEY_SECURE_2026',
        sender_name TEXT DEFAULT '프리비 물류센터',
        sender_phone TEXT DEFAULT '02-1234-5678',
        sender_address TEXT DEFAULT '서울특별시 강남구 테헤란로 123 물류빌딩 3층',
        is_mock INTEGER DEFAULT 1,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 초기 기본 택배사 설정 확인 및 삽입
    cur.execute('SELECT COUNT(*) as cnt FROM courier_settings')
    if cur.fetchone()['cnt'] == 0:
        cur.execute('''
        INSERT INTO courier_settings (courier_name, customer_code, api_key, sender_name, sender_phone, sender_address, is_mock)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('CJ대한통운', '12345678', 'MOCK_API_KEY_CJ', '프리비 물류센터', '02-1234-5678', '서울특별시 강남구 테헤란로 123 물류센터', 1))

    conn.commit()
    conn.close()
    print("Database initialized successfully at", DB_PATH)

if __name__ == '__main__':
    init_db()
