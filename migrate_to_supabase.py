import os
import sqlite3
import psycopg2
from psycopg2.extras import DictCursor

PG_CONFIG = {
    "host": "aws-0-ap-south-1.pooler.supabase.com",
    "port": 5432,
    "user": "postgres.qazeqzbajdlrddfvmtxy",
    "password": "-WHkmP9pnRbT+V%",
    "dbname": "postgres",
    "sslmode": "require"
}

def migrate():
    print("1. Connecting to PostgreSQL (Supabase)...")
    pg_conn = psycopg2.connect(**PG_CONFIG)
    pg_cur = pg_conn.cursor()

    print("2. Creating schema on PostgreSQL...")
    pg_cur.execute('''
    CREATE TABLE IF NOT EXISTS products (
        id SERIAL PRIMARY KEY,
        sku TEXT UNIQUE,
        product_name TEXT NOT NULL,
        option_name TEXT DEFAULT '',
        english_name TEXT DEFAULT '',
        hs_code TEXT DEFAULT '',
        unit_price REAL DEFAULT 0.0,
        stock_quantity INTEGER DEFAULT 0,
        product_url TEXT DEFAULT '',
        safety_stock INTEGER DEFAULT 10,
        location TEXT DEFAULT 'A-01',
        dispatched_quantity INTEGER DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS product_options (
        id SERIAL PRIMARY KEY,
        product_id INTEGER REFERENCES products(id) ON DELETE CASCADE,
        sku TEXT NOT NULL,
        option_name TEXT NOT NULL,
        unit_price REAL DEFAULT 0.0,
        dispatched_quantity INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS unregistered_products (
        id SERIAL PRIMARY KEY,
        product_name TEXT NOT NULL,
        option_name TEXT DEFAULT '',
        combined_name TEXT UNIQUE NOT NULL,
        order_count INTEGER DEFAULT 1,
        status TEXT DEFAULT 'pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS bundle_rules (
        id SERIAL PRIMARY KEY,
        rule_name TEXT,
        source_product_name TEXT NOT NULL,
        source_option_name TEXT NOT NULL,
        target_product_name TEXT NOT NULL,
        target_option_name TEXT NOT NULL,
        target_sku TEXT,
        multiplier INTEGER DEFAULT 1,
        is_active INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS orders (
        id SERIAL PRIMARY KEY,
        order_no TEXT NOT NULL,
        courier TEXT DEFAULT '나르다',
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
        dispatched_at TEXT DEFAULT '',
        weight REAL DEFAULT 0.0,
        customs_validation_status TEXT DEFAULT '',
        customs_validation_msg TEXT DEFAULT '',
        invoice_pdf_url TEXT DEFAULT '',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS order_items (
        id SERIAL PRIMARY KEY,
        order_id INTEGER REFERENCES orders(id) ON DELETE CASCADE,
        order_no TEXT NOT NULL,
        sku TEXT,
        item_product_name TEXT NOT NULL,
        item_option_name TEXT NOT NULL,
        option_quantity INTEGER DEFAULT 1,
        quantity INTEGER DEFAULT 1,
        stock_deducted INTEGER DEFAULT 0,
        status TEXT DEFAULT '준비중',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS inventory_logs (
        id SERIAL PRIMARY KEY,
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
    );

    CREATE TABLE IF NOT EXISTS courier_settings (
        id SERIAL PRIMARY KEY,
        courier_name TEXT DEFAULT '나르다',
        customer_code TEXT DEFAULT 'TOA',
        api_key TEXT DEFAULT '5a65132fee9b4780c350b1942db432afd3232b30',
        narda_id TEXT DEFAULT 'calroid',
        sender_name TEXT DEFAULT 'TOA',
        sender_phone TEXT DEFAULT '02-1234-5678',
        sender_address TEXT DEFAULT '서울특별시 강남구 테헤란로 123 물류빌딩 3층',
        is_mock INTEGER DEFAULT 1,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    ''')
    pg_conn.commit()
    print("Schema created successfully!")

    # 3. Transfer data from SQLite orders.db
    sqlite_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'orders.db')
    if not os.path.exists(sqlite_path):
        print("No local orders.db found, skipping data copy.")
        pg_conn.close()
        return

    print("3. Migrating data from local orders.db...")
    sq_conn = sqlite3.connect(sqlite_path)
    sq_conn.row_factory = sqlite3.Row
    sq_cur = sq_conn.cursor()

    # Tables to copy
    tables = [
        'products', 'product_options', 'unregistered_products',
        'bundle_rules', 'orders', 'order_items', 'inventory_logs', 'courier_settings'
    ]

    for tbl in tables:
        sq_cur.execute(f"PRAGMA table_info({tbl})")
        cols = [r['name'] for r in sq_cur.fetchall()]
        if not cols:
            continue

        sq_cur.execute(f"SELECT * FROM {tbl}")
        rows = sq_cur.fetchall()
        print(f" - Copying {len(rows)} rows from {tbl}...")
        if rows:
            # Clear PG table first to prevent duplicate keys on re-run
            pg_cur.execute(f"TRUNCATE TABLE {tbl} CASCADE")
            col_names = ', '.join(cols)
            placeholders = ', '.join(['%s'] * len(cols))
            insert_sql = f"INSERT INTO {tbl} ({col_names}) VALUES ({placeholders})"
            for r in rows:
                pg_cur.execute(insert_sql, tuple(r[c] for c in cols))
            pg_conn.commit()

            # Reset serial sequence
            if 'id' in cols:
                pg_cur.execute(f"SELECT setval('{tbl}_id_seq', COALESCE((SELECT MAX(id) FROM {tbl}), 1))")
                pg_conn.commit()

    sq_conn.close()
    pg_conn.close()
    print("MIGRATION COMPLETED SUCCESSFULLY!")

if __name__ == '__main__':
    migrate()
