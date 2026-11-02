import os
import re
import sqlite3
import threading
from datetime import datetime

# Supabase PostgreSQL 연결 설정
SUPABASE_CONFIG = {
    "host": os.environ.get("PGHOST", "aws-0-ap-south-1.pooler.supabase.com"),
    "port": int(os.environ.get("PGPORT", 5432)),
    "user": os.environ.get("PGUSER", "postgres.qazeqzbajdlrddfvmtxy"),
    "password": os.environ.get("PGPASSWORD", "-WHkmP9pnRbT+V%"),
    "dbname": os.environ.get("PGDATABASE", "postgres"),
    "sslmode": "require"
}

USE_POSTGRES = True
try:
    import psycopg2
    import psycopg2.extras
    from psycopg2.pool import ThreadedConnectionPool
except ImportError:
    USE_POSTGRES = False
    ThreadedConnectionPool = None

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'orders.db')
_pg_pool = None
_pg_pool_lock = threading.Lock()

class PgCursorWrapper:
    def __init__(self, raw_cursor):
        self._cur = raw_cursor
        self.lastrowid = None

    def execute(self, sql, params=None):
        clean_sql = sql.strip()

        # 1. PRAGMA table_info(xxx) 변환
        m = re.match(r"PRAGMA\s+table_info\((['\"]?(\w+)['\"]?)\)", clean_sql, re.IGNORECASE)
        if m:
            tbl_name = m.group(2).lower()
            trans_sql = "SELECT column_name AS name FROM information_schema.columns WHERE table_name = %s ORDER BY ordinal_position"
            self._cur.execute(trans_sql, (tbl_name,))
            return self

        # 2. psycopg2 문자열 포맷팅 충돌 방지: 기존 %를 %%로 이스케이프 후 ? -> %s 변환
        trans_sql = clean_sql.replace('%', '%%').replace('?', '%s')

        # 3. INSERT 시 lastrowid 자동 지원을 위한 RETURNING id 처리
        is_insert = trans_sql.upper().startswith('INSERT INTO')
        has_returning = 'RETURNING' in trans_sql.upper()

        if is_insert and not has_returning:
            trans_sql_ret = trans_sql + ' RETURNING id'
            try:
                if params:
                    self._cur.execute(trans_sql_ret, tuple(params))
                else:
                    self._cur.execute(trans_sql_ret)
                ret_row = self._cur.fetchone()
                if ret_row:
                    self.lastrowid = ret_row['id'] if (isinstance(ret_row, dict) or hasattr(ret_row, '__getitem__')) else ret_row[0]
                return self
            except Exception:
                pass

        if params:
            self._cur.execute(trans_sql, tuple(params))
        else:
            self._cur.execute(trans_sql)
        return self

    def executemany(self, sql, seq_of_params):
        trans_sql = sql.replace('%', '%%').replace('?', '%s')
        self._cur.executemany(trans_sql, seq_of_params)
        return self

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    def fetchmany(self, size=None):
        return self._cur.fetchmany(size) if size else self._cur.fetchmany()

    @property
    def rowcount(self):
        return self._cur.rowcount

    def close(self):
        self._cur.close()

class PgConnectionWrapper:
    def __init__(self, pg_conn, pool=None):
        self._conn = pg_conn
        self._pool = pool

    def cursor(self):
        return PgCursorWrapper(self._conn.cursor(cursor_factory=psycopg2.extras.DictCursor))

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        if self._pool:
            self._pool.putconn(self._conn)
        else:
            self._conn.close()

def _get_postgres_pool():
    global _pg_pool
    if _pg_pool is None:
        with _pg_pool_lock:
            if _pg_pool is None:
                database_url = os.environ.get('DATABASE_URL')
                try:
                    if database_url:
                        _pg_pool = ThreadedConnectionPool(1, 12, database_url, sslmode='require')
                    else:
                        _pg_pool = ThreadedConnectionPool(1, 12, **SUPABASE_CONFIG)
                except Exception:
                    if not database_url:
                        raise
                    _pg_pool = ThreadedConnectionPool(1, 12, **SUPABASE_CONFIG)
    return _pg_pool

def get_connection():
    if USE_POSTGRES:
        try:
            pool = _get_postgres_pool()
            try:
                return PgConnectionWrapper(pool.getconn(), pool)
            except Exception:
                # 풀의 모든 연결이 사용 중일 때는 SQLite로 전환하지 않고 임시 연결을 사용합니다.
                database_url = os.environ.get('DATABASE_URL')
                if database_url:
                    try:
                        return PgConnectionWrapper(psycopg2.connect(database_url, sslmode='require'))
                    except Exception:
                        pass
                return PgConnectionWrapper(psycopg2.connect(**SUPABASE_CONFIG))
        except Exception as e:
            print(f"[DB Warning] PostgreSQL connection failed ({e}), falling back to SQLite.")

    # SQLite fallback
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cur = conn.cursor()

    if isinstance(conn, PgConnectionWrapper):
        cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS api_dispatch_status TEXT DEFAULT ''")
        cur.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS api_dispatch_message TEXT DEFAULT ''")
        cur.execute('CREATE INDEX IF NOT EXISTS idx_orders_status_id ON orders (status, id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_order_items_order_id_id ON order_items (order_id, id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_order_items_sku ON order_items (sku)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_inventory_logs_created_at ON inventory_logs (created_at)')
        conn.commit()
        conn.close()
        print("Connected to Supabase PostgreSQL cloud database successfully.")
        return

    conn.execute('PRAGMA journal_mode=WAL')

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
        product_url TEXT DEFAULT '',
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
    if 'product_url' not in existing_cols:
        cur.execute("ALTER TABLE products ADD COLUMN product_url TEXT DEFAULT ''")
    if 'dispatched_quantity' not in existing_cols:
        cur.execute("ALTER TABLE products ADD COLUMN dispatched_quantity INTEGER DEFAULT 0")



    # 1-1. 상품 하위 옵션별 출고/단가 마스터 (Product Options)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS product_options (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        sku TEXT NOT NULL,
        option_name TEXT NOT NULL,
        unit_price REAL DEFAULT 0.0,
        dispatched_quantity INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE
    )
    ''')

    # 1-2. 주문 엑셀에서 발견된 미등록 신규 상품 (Unregistered Products)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS unregistered_products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_name TEXT NOT NULL,
        option_name TEXT DEFAULT '',
        combined_name TEXT NOT NULL,
        order_count INTEGER DEFAULT 1,
        status TEXT DEFAULT 'pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(combined_name)
    )
    ''')

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
        api_dispatch_status TEXT DEFAULT '',
        api_dispatch_message TEXT DEFAULT '',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    cur.execute("PRAGMA table_info(orders)")
    order_cols = [row[1] for row in cur.fetchall()]
    if 'dispatched_at' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN dispatched_at TEXT DEFAULT ''")
    if 'weight' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN weight REAL DEFAULT 0.0")
    if 'customs_validation_status' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN customs_validation_status TEXT DEFAULT ''")
    if 'customs_validation_msg' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN customs_validation_msg TEXT DEFAULT ''")
    if 'invoice_pdf_url' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN invoice_pdf_url TEXT DEFAULT ''")
    if 'api_dispatch_status' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN api_dispatch_status TEXT DEFAULT ''")
    if 'api_dispatch_message' not in order_cols:
        cur.execute("ALTER TABLE orders ADD COLUMN api_dispatch_message TEXT DEFAULT ''")

    # 기존 출고대기 상태 주문접수로 통일 정규화
    cur.execute("UPDATE orders SET status = '주문접수' WHERE status = '출고대기'")


    # 4. 분할된 출고 아이템 (Order Items)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS order_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        order_no TEXT NOT NULL,
        sku TEXT,
        item_product_name TEXT NOT NULL,
        item_option_name TEXT NOT NULL,
        option_quantity INTEGER DEFAULT 1,
        quantity INTEGER DEFAULT 1,
        stock_deducted INTEGER DEFAULT 0,
        status TEXT DEFAULT '준비중',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE
    )
    ''')

    cur.execute("PRAGMA table_info(order_items)")
    item_cols = [row[1] for row in cur.fetchall()]
    if 'option_quantity' not in item_cols:
        cur.execute("ALTER TABLE order_items ADD COLUMN option_quantity INTEGER DEFAULT 1")

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
        courier_name TEXT DEFAULT '나르다',
        customer_code TEXT DEFAULT '정마무역',
        api_key TEXT DEFAULT 'e1df1b2a432106cd6f4c886dfb9c3ca24fbf27e9',
        sender_name TEXT DEFAULT 'TOA',
        sender_phone TEXT DEFAULT '02-1234-5678',
        sender_address TEXT DEFAULT '서울특별시 강남구 테헤란로 123 물류빌딩 3층',
        is_mock INTEGER DEFAULT 0,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # courier_settings 테이블 컬럼 마이그레이션 (narda_id 추가)
    cur.execute("PRAGMA table_info(courier_settings)")
    settings_cols = [row[1] for row in cur.fetchall()]
    if 'narda_id' not in settings_cols:
        cur.execute("ALTER TABLE courier_settings ADD COLUMN narda_id TEXT DEFAULT 'calroid'")

    # 초기 기본 택배사 설정 확인 및 삽입
    cur.execute('SELECT COUNT(*) as cnt FROM courier_settings')
    if cur.fetchone()['cnt'] == 0:
        cur.execute('''
        INSERT INTO courier_settings (courier_name, customer_code, api_key, narda_id, sender_name, sender_phone, sender_address, is_mock)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', ('나르다', 'TOA', '5a65132fee9b4780c350b1942db432afd3232b30', 'calroid', 'TOA', '02-1234-5678', '서울특별시 강남구 테헤란로 123 물류센터', 1))

    # 7. 옵션 자동 분할 구분 기호 설정 (Split Delimiters)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS split_delimiters (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        symbol TEXT UNIQUE NOT NULL
    )
    ''')
    try:
        cur.execute('SELECT COUNT(*) as cnt FROM split_delimiters')
        if cur.fetchone()['cnt'] == 0:
            for s in ['+', '-', '/', ',', '?', '&']:
                try:
                    cur.execute('INSERT INTO split_delimiters (symbol) VALUES (?)', (s,))
                except Exception:
                    pass
    except Exception:
        pass

    cur.execute('CREATE INDEX IF NOT EXISTS idx_orders_status_id ON orders (status, id)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_order_items_order_id_id ON order_items (order_id, id)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_order_items_sku ON order_items (sku)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_inventory_logs_created_at ON inventory_logs (created_at)')

    conn.commit()
    conn.close()
    print("Database initialized successfully at", DB_PATH)

if __name__ == '__main__':
    init_db()
