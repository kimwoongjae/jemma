import re
import os
import pandas as pd
from database import get_connection

HEADER_MAP = {
    'order_no': ['주문번호', '주문 번호', 'order_no', 'Order No'],
    'courier': ['택배사', '배송업체', '배송사', 'courier'],
    'tracking_no': ['송장번호', '운송장번호', 'tracking_no'],
    'shipping_method': ['배송방법', '배송방식', 'shipping_method'],
    'product_name': ['상품명', '주문상품명', '품목명', 'product_name'],
    'option_name': ['옵션이름', '옵션명', '선택사항', 'option_name'],
    'quantity': ['상품수량', '수량', '주문수량', 'qty', 'quantity'],
    'recipient_name': ['수취인이름', '수취인', '받는사람', '받는분', '고객명', 'recipient_name'],
    'recipient_phone1': ['수취인전화번호1', '수취인 전화번호1', '전화번호1', '연락처1', '수취인연락처', 'recipient_phone1'],
    'recipient_phone2': ['수취인전화번호2', '수취인 전화번호2', '전화번호2', '연락처2', 'recipient_phone2'],
    'customs_id': ['개인통관부호', '개인통관고유부호', '통관부호', 'customs_id'],
    'postal_code': ['수취인우편번호', '우편번호', 'postal_code', 'zipcode'],
    'address': ['수취인주소', '주소', '배송지', '배송주소', 'address'],
    'shipping_memo': ['배송메세지', '배송요청사항', '배송메모', '배송요청', 'shipping_memo']
}

def detect_columns(df_columns):
    mapping = {}
    col_list = [str(c).strip() for c in df_columns]
    for key, candidates in HEADER_MAP.items():
        found = None
        for cand in candidates:
            for actual in col_list:
                if cand.lower() == actual.lower():
                    found = actual
                    break
            if found:
                break
        mapping[key] = found
    return mapping

def generate_sku(product_name, option_name):
    # 영문/숫자/한글 축약으로 고유 SKU 생성
    clean_p = re.sub(r'\[.*?\]|\(.*?\)', '', product_name).strip()
    clean_p = re.sub(r'[^a-zA-Z0-9가-힣]', '', clean_p)[:6]
    clean_o = re.sub(r'[^a-zA-Z0-9가-힣]', '', option_name)[:6]
    sku = f"SKU-{clean_p.upper()}-{clean_o.upper()}"
    if len(sku) <= 5:
        sku = f"SKU-{abs(hash(product_name + option_name)) % 100000:05d}"
    return sku

def split_order_items(product_name, option_name, quantity, conn=None):
    product_name = str(product_name).strip() if pd.notna(product_name) else '상품명미상'
    option_name = str(option_name).strip() if pd.notna(option_name) else '단일옵션'
    try:
        qty = int(quantity) if pd.notna(quantity) and str(quantity).isdigit() else 1
    except:
        qty = 1

    # 1. DB의 bundle_rules에서 매핑 룰 탐색
    if conn:
        cur = conn.cursor()
        cur.execute('''
            SELECT target_product_name, target_option_name, target_sku, multiplier
            FROM bundle_rules
            WHERE is_active = 1
              AND (source_product_name = ? OR source_product_name = '*')
              AND source_option_name = ?
        ''', (product_name, option_name))
        rules = cur.fetchall()
        if rules:
            items = []
            for r in rules:
                sub_p = r['target_product_name'] or product_name
                sub_o = r['target_option_name'] or option_name
                sub_sku = r['target_sku'] or generate_sku(sub_p, sub_o)
                sub_qty = qty * (r['multiplier'] or 1)
                items.append({
                    'item_product_name': sub_p,
                    'item_option_name': sub_o,
                    'sku': sub_sku,
                    'option_quantity': r['multiplier'] or 1,
                    'quantity': sub_qty
                })
            return items

    # 2. 규칙이 없으면 '+' 기호로 분할
    items = []
    if '+' in option_name:
        parts = [p.strip() for p in option_name.split('+') if p.strip()]
        for p in parts:
            sku = generate_sku(product_name, p)
            items.append({
                'item_product_name': product_name,
                'item_option_name': p,
                'sku': sku,
                'option_quantity': 1,
                'quantity': qty
            })
    else:
        # 단일 품목 (옵션명에 숫자가 있는 경우 옵션 수량 인식, 예: '4', '4봉', '4개')
        import re
        opt_qty = 1
        num_match = re.match(r'^(\d+)\s*(개|봉|입|세트|box|박스)?$', option_name)
        if num_match:
            try:
                opt_qty = int(num_match.group(1))
            except:
                opt_qty = 1

        sku = generate_sku(product_name, option_name)
        items.append({
            'item_product_name': product_name,
            'item_option_name': option_name,
            'sku': sku,
            'option_quantity': opt_qty,
            'quantity': qty
        })

    return items

def process_excel_file(file_path, batch_name=None):
    if not batch_name:
        batch_name = os.path.basename(file_path)

    df = pd.read_excel(file_path)
    col_map = detect_columns(df.columns)

    conn = get_connection()
    parsed_orders = []
    preview_data = {
        'batch_name': batch_name,
        'total_excel_rows': len(df),
        'detected_columns': {k: v for k, v in col_map.items() if v},
        'orders': []
    }

    for idx, row in df.iterrows():
        def val(k):
            col = col_map.get(k)
            if col and col in row and pd.notna(row[col]):
                return str(row[col]).strip()
            return ''

        order_no = val('order_no') or f"ORD-{batch_name[:4]}-{idx+1:04d}"
        courier = val('courier') or 'CJ대한통운'
        tracking_no = val('tracking_no')
        shipping_method = val('shipping_method') or '택배'
        product_name = val('product_name')
        option_name = val('option_name')
        
        try:
            quantity = int(float(val('quantity'))) if val('quantity') else 1
        except:
            quantity = 1

        recipient_name = val('recipient_name')
        recipient_phone1 = val('recipient_phone1')
        recipient_phone2 = val('recipient_phone2')
        customs_id = val('customs_id')
        postal_code = val('postal_code')
        address = val('address')
        shipping_memo = val('shipping_memo')

        # 분할 처리
        split_items = split_order_items(product_name, option_name, quantity, conn)

        order_obj = {
            'order_no': order_no,
            'courier': courier,
            'tracking_no': tracking_no,
            'shipping_method': shipping_method,
            'original_product_name': product_name,
            'original_option_name': option_name,
            'original_quantity': quantity,
            'recipient_name': recipient_name,
            'recipient_phone1': recipient_phone1,
            'recipient_phone2': recipient_phone2,
            'customs_id': customs_id,
            'postal_code': postal_code,
            'address': address,
            'shipping_memo': shipping_memo,
            'split_items': split_items,
            'is_split': len(split_items) > 1
        }
        parsed_orders.append(order_obj)

    conn.close()
    preview_data['orders'] = parsed_orders
    preview_data['total_split_items'] = sum(len(o['split_items']) for o in parsed_orders)
    return preview_data

def save_orders_to_db(parsed_data, batch_name):
    conn = get_connection()
    cur = conn.cursor()
    saved_count = 0
    item_count = 0

    for ord_data in parsed_data['orders']:
        cur.execute('''
            INSERT INTO orders (
                order_no, courier, tracking_no, shipping_method,
                original_product_name, original_option_name, original_quantity,
                recipient_name, recipient_phone1, recipient_phone2,
                customs_id, postal_code, address, shipping_memo,
                status, upload_batch
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            ord_data['order_no'], ord_data['courier'], ord_data['tracking_no'], ord_data['shipping_method'],
            ord_data['original_product_name'], ord_data['original_option_name'], ord_data['original_quantity'],
            ord_data['recipient_name'], ord_data['recipient_phone1'], ord_data['recipient_phone2'],
            ord_data['customs_id'], ord_data['postal_code'], ord_data['address'], ord_data['shipping_memo'],
            '주문접수', batch_name
        ))
        order_id = cur.lastrowid
        saved_count += 1

        for it in ord_data['split_items']:
            # 상품 마스터(products)에 SKU 없으면 자동 등록 (기본재고 100으로 편의 등록 가능)
            cur.execute('SELECT id, stock_quantity FROM products WHERE sku = ?', (it['sku'],))
            prod = cur.fetchone()
            if not prod:
                cur.execute('''
                    INSERT INTO products (sku, product_name, option_name, stock_quantity, safety_stock)
                    VALUES (?, ?, ?, ?, ?)
                ''', (it['sku'], it['item_product_name'], it['item_option_name'], 100, 10))
                prod_id = cur.lastrowid
                # 초기 입고 로그
                cur.execute('''
                    INSERT INTO inventory_logs (product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (prod_id, it['sku'], it['item_product_name'], it['item_option_name'], '신규등록', 100, 100, '주문 업로드 시 자동 등록'))

            cur.execute('''
                INSERT INTO order_items (
                    order_id, order_no, sku, item_product_name, item_option_name, option_quantity, quantity, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                order_id, ord_data['order_no'], it['sku'], it['item_product_name'], it['item_option_name'], it.get('option_quantity', 1), it['quantity'], '준비중'
            ))
            item_count += 1

    conn.commit()
    conn.close()
    return {'saved_orders': saved_count, 'saved_items': item_count}
