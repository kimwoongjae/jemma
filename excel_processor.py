import re
import os
import pandas as pd
from database import get_connection

HEADER_MAP = {
    'order_no': ['주문번호', '주문 번호', 'order_no', 'Order No', '발주번호', '주문관리코드', '주문식별자'],
    'courier': ['택배사', '배송업체', '배송사', 'courier', '택배사명'],
    'tracking_no': ['송장번호', '운송장번호', 'tracking_no', '운송장 번호', '송장'],
    'shipping_method': ['배송방법', '배송방식', 'shipping_method'],
    'product_name': ['상품명', '주문상품명', '품목명', 'product_name', '상품', '판매상품명'],
    'option_name': ['옵션이름', '옵션명', '선택사항', 'option_name', '옵션정보', '옵션', '상품옵션', '선택옵션'],
    'quantity': ['상품수량', '수량', '주문수량', 'qty', 'quantity', '구매수량'],
    'recipient_name': ['수취인이름', '수취인', '받는사람', '받는분', '고객명', 'recipient_name', '수취인명'],
    'recipient_phone1': ['수취인전화번호1', '수취인 전화번호1', '전화번호1', '연락처1', '수취인연락처', 'recipient_phone1', '수취인휴대폰', '휴대폰번호', '휴대전화', '수취인연락처1'],
    'recipient_phone2': ['수취인전화번호2', '수취인 전화번호2', '전화번호2', '연락처2', 'recipient_phone2', '기타연락처', '전화번호', '수취인연락처2'],
    'customs_id': ['개인통관부호', '개인통관고유부호', '통관부호', 'customs_id', '통관고유부호', '개인통관고유번호', '통관번호'],
    'postal_code': ['수취인우편번호', '우편번호', 'postal_code', 'zipcode', '우편 번호', '수취인 우편번호'],
    'address': ['수취인주소', '주소', '배송지', '배송주소', 'address', '수취인 기본주소', '기본주소', '배송지주소', '전체주소'],
    'shipping_memo': ['배송메세지', '배송요청사항', '배송메모', '배송요청', 'shipping_memo', '배송메시지']
}

def detect_columns(df_columns):
    mapping = {}
    col_list = [str(c).strip() for c in df_columns]
    for key, candidates in HEADER_MAP.items():
        found = None
        for cand in candidates:
            cand_clean = re.sub(r'\s+', '', cand).lower()
            for actual in col_list:
                actual_clean = re.sub(r'\s+', '', actual).lower()
                if cand_clean == actual_clean:
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

def find_matching_product_info(product_name, option_name, conn=None, opt_rows=None, prod_rows=None, exact_only=False):
    """
    재고관리(product_options 및 products) 테이블에서 상품명/옵션명 또는 합친 명칭과 일치하는 바코드 및 출고수량 배수를 조회합니다.
    """
    combined = f"{product_name} {option_name}".strip() if (option_name and option_name != '단일옵션') else product_name
    clean_combined = re.sub(r'\s+', '', combined)
    clean_p = re.sub(r'\s+', '', product_name)
    clean_o = re.sub(r'\s+', '', option_name) if (option_name and option_name != '단일옵션') else ''

    close_after = False
    if opt_rows is None or prod_rows is None:
        if not conn:
            from database import get_connection
            conn = get_connection()
            close_after = True

    try:
        if opt_rows is None:
            cur = conn.cursor()
            try:
                cur.execute('SELECT sku, option_name, unit_price, dispatched_quantity FROM product_options')
                opt_rows = [dict(r) for r in cur.fetchall()]
            except Exception:
                opt_rows = []

        if prod_rows is None:
            cur = conn.cursor()
            cur.execute('SELECT sku, product_name, option_name, unit_price, dispatched_quantity FROM products')
            prod_rows = [dict(r) for r in cur.fetchall()]

        # 1-1. 옵션명 완전 일치
        for r in opt_rows:
            opt_name = (r['option_name'] or '').strip()
            if opt_name == combined or opt_name == product_name or (clean_o and opt_name == option_name):
                return {
                    'sku': r['sku'],
                    'dispatched_quantity': r['dispatched_quantity'] or 1,
                    'unit_price': r.get('unit_price', 0.0),
                    'matched': True,
                    'matched_product_name': None,
                    'matched_option_name': r.get('option_name')
                }

        # 1-2. 공백 제거 후 일치
        for r in opt_rows:
            opt_name = re.sub(r'\s+', '', r['option_name'] or '')
            if opt_name == clean_combined or opt_name == clean_p:
                return {
                    'sku': r['sku'],
                    'dispatched_quantity': r['dispatched_quantity'] or 1,
                    'unit_price': r.get('unit_price', 0.0),
                    'matched': True,
                    'matched_product_name': None,
                    'matched_option_name': r.get('option_name')
                }

        # 2-1. 완전 일치 (합친 명칭 or 상품명)
        for r in prod_rows:
            p_name = r['product_name'] or ''
            p_opt = r['option_name'] or ''
            p_combined = f"{p_name} {p_opt}".strip() if p_opt else p_name
            if p_combined == combined or p_name == combined:
                return {
                    'sku': r['sku'],
                    'dispatched_quantity': r.get('dispatched_quantity') or 1,
                    'unit_price': r.get('unit_price', 0.0),
                    'matched': True,
                    'matched_product_name': r.get('product_name'),
                    'matched_option_name': r.get('option_name')
                }

        # 2-2. product_name + option_name 일치
        for r in prod_rows:
            p_name = r['product_name'] or ''
            p_opt = r['option_name'] or ''
            if p_name == product_name and (p_opt == option_name or not p_opt):
                return {
                    'sku': r['sku'],
                    'dispatched_quantity': r.get('dispatched_quantity') or 1,
                    'unit_price': r.get('unit_price', 0.0),
                    'matched': True,
                    'matched_product_name': r.get('product_name'),
                    'matched_option_name': r.get('option_name')
                }

        # 2-3. 공백 제거 후 일치
        for r in prod_rows:
            p_name = re.sub(r'\s+', '', r['product_name'] or '')
            p_opt = re.sub(r'\s+', '', r['option_name'] or '')
            p_comb = (p_name + p_opt) if p_opt else p_name
            if p_comb == clean_combined or p_name == clean_combined:
                return {
                    'sku': r['sku'],
                    'dispatched_quantity': r.get('dispatched_quantity') or 1,
                    'unit_price': r.get('unit_price', 0.0),
                    'matched': True,
                    'matched_product_name': r.get('product_name'),
                    'matched_option_name': r.get('option_name')
                }

        # 2-4. 키워드 부분 일치 (exact_only=False 일 때만 수행)
        if not exact_only:
            for r in prod_rows:
                p_name = re.sub(r'\s+', '', r['product_name'] or '')
                if len(p_name) >= 3 and (p_name in clean_combined or clean_combined in p_name):
                    if clean_o:
                        if clean_o in p_name or (r['option_name'] and re.sub(r'\s+', '', r['option_name']) in clean_o):
                            return {
                                'sku': r['sku'],
                                'dispatched_quantity': r.get('dispatched_quantity') or 1,
                                'unit_price': r.get('unit_price', 0.0),
                                'matched': True,
                                'matched_product_name': r.get('product_name'),
                                'matched_option_name': r.get('option_name')
                            }
                    else:
                        return {
                            'sku': r['sku'],
                            'dispatched_quantity': r.get('dispatched_quantity') or 1,
                            'unit_price': r.get('unit_price', 0.0),
                            'matched': True,
                            'matched_product_name': r.get('product_name'),
                            'matched_option_name': r.get('option_name')
                        }

    finally:
        if close_after and conn:
            conn.close()

    return {'sku': None, 'dispatched_quantity': 1, 'unit_price': 0.0, 'matched': False}

def find_matching_barcode(product_name, option_name, conn=None):
    res = find_matching_product_info(product_name, option_name, conn)
    return res['sku'] if res['matched'] else None

def sync_order_items_sku(conn=None):
    """모든 order_items의 상품명/옵션명에 맞추어 재고관리(products 및 product_options)의 바코드(SKU) 및 옵션출고수량으로 최신 동기화"""
    close_after = False
    if not conn:
        from database import get_connection
        conn = get_connection()
        close_after = True

    try:
        cur = conn.cursor()
        try:
            cur.execute('SELECT sku, option_name, unit_price, dispatched_quantity FROM product_options')
            opt_rows = [dict(r) for r in cur.fetchall()]
        except Exception:
            opt_rows = []

        cur.execute('SELECT sku, product_name, option_name, unit_price, dispatched_quantity FROM products')
        prod_rows = [dict(r) for r in cur.fetchall()]

        cur.execute('SELECT id, item_product_name, item_option_name, sku, option_quantity FROM order_items')
        items = cur.fetchall()
        for it in items:
            info = find_matching_product_info(it['item_product_name'], it['item_option_name'], conn, opt_rows=opt_rows, prod_rows=prod_rows)
            if info['matched']:
                new_sku = info['sku']
                new_opt_qty = info['dispatched_quantity'] or 1
                if new_sku != it['sku'] or (it['option_quantity'] == 1 and new_opt_qty > 1):
                    cur.execute('UPDATE order_items SET sku = ?, option_quantity = ? WHERE id = ?', (new_sku, new_opt_qty, it['id']))
        conn.commit()
    finally:
        if close_after and conn:
            conn.close()

def split_order_items(product_name, option_name, quantity, conn=None, opt_rows=None, prod_rows=None, rules_cache=None, delimiters=None):
    product_name = str(product_name).strip() if pd.notna(product_name) else '상품명미상'
    option_name = str(option_name).strip() if pd.notna(option_name) else '단일옵션'
    try:
        qty = int(quantity) if pd.notna(quantity) and str(quantity).isdigit() else 1
    except:
        qty = 1

    # 1. DB의 bundle_rules에서 매핑 룰 탐색 (메모리 캐시 우선 조회로 초고속 처리)
    rules = []
    if rules_cache is not None:
        rules = [
            r for r in rules_cache
            if (r.get('source_product_name') == product_name or r.get('source_product_name') == '*')
            and r.get('source_option_name') == option_name
        ]
    elif conn:
        cur = conn.cursor()
        cur.execute('''
            SELECT target_product_name, target_option_name, target_sku, multiplier, source_product_name, source_option_name
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
            info = find_matching_product_info(sub_p, sub_o, conn, opt_rows=opt_rows, prod_rows=prod_rows)
            sub_sku = r['target_sku'] or info['sku'] or generate_sku(sub_p, sub_o)
            opt_multiplier = (r['multiplier'] or 1) * (info['dispatched_quantity'] if info['matched'] else 1)
            sub_qty = qty * opt_multiplier
            items.append({
                'item_product_name': sub_p,
                'item_option_name': sub_o,
                'sku': sub_sku,
                'product_quantity': qty,
                'option_quantity': opt_multiplier,
                'quantity': sub_qty,
                'matched': info['matched'] or bool(r['target_sku'])
            })
        return items

    # 2. 먼저 전체 옵션명이 재고관리에 등록된 단일 상품과 완전히 일치하는지 확인 (완전 일치만 허용)
    direct_info = find_matching_product_info(product_name, option_name, conn, opt_rows=opt_rows, prod_rows=prod_rows, exact_only=True)
    if direct_info['matched']:
        opt_qty = direct_info['dispatched_quantity'] or 1
        return [{
            'item_product_name': direct_info.get('matched_product_name') or product_name,
            'item_option_name': direct_info.get('matched_option_name') or option_name,
            'sku': direct_info['sku'],
            'product_quantity': qty,
            'option_quantity': opt_qty,
            'quantity': qty * opt_qty,
            'matched': True
        }]

    # 3. 구분 기호(+, -, ,, /, ?, & 등)로 자동 다중 품목 분할
    items = []
    if not delimiters:
        delimiters = ['+', '-', '/', ',', '?', '&']
    delimiters_pattern = '|'.join(re.escape(s.strip()) for s in delimiters if s and s.strip())

    parts = [p.strip() for p in re.split(delimiters_pattern, option_name) if p.strip()] if delimiters_pattern else [option_name]

    # 옵션명이 비어있거나 '단일옵션'인데 상품명에 구분 기호가 있는 경우 상품명 분할 지원
    is_product_split = False
    if len(parts) <= 1 and (not option_name or option_name in ('단일옵션', '단일상품')):
        prod_parts = [p.strip() for p in re.split(delimiters_pattern, product_name) if p.strip()] if delimiters_pattern else [product_name]
        if len(prod_parts) > 1:
            parts = prod_parts
            is_product_split = True

    if len(parts) > 1:
        for p in parts:
            lookup_p = '' if is_product_split else product_name
            info = find_matching_product_info(lookup_p, p, conn, opt_rows=opt_rows, prod_rows=prod_rows)
            sku = info['sku'] if info['matched'] else generate_sku(product_name, p)
            opt_qty = info['dispatched_quantity'] if info['matched'] else 1
            items.append({
                'item_product_name': info.get('matched_product_name') or (product_name if not is_product_split else p),
                'item_option_name': info.get('matched_option_name') or (p if not is_product_split else '단일옵션'),
                'sku': sku,
                'product_quantity': qty,
                'option_quantity': opt_qty,
                'quantity': qty * opt_qty,
                'matched': info['matched']
            })
    else:
        # 단일 품목 (부분 일치 포함 조회 및 숫자 수량 인식)
        info = find_matching_product_info(product_name, option_name, conn, opt_rows=opt_rows, prod_rows=prod_rows, exact_only=False)
        num_match = re.match(r'^(\d+)\s*(개|봉|입|세트|box|박스)?$', option_name)
        opt_qty = (info['dispatched_quantity'] if info['matched'] else int(num_match.group(1))) if (info['matched'] or num_match) else 1
        sku = info['sku'] if info['matched'] else generate_sku(product_name, option_name)

        items.append({
            'item_product_name': info.get('matched_product_name') or product_name,
            'item_option_name': info.get('matched_option_name') or option_name,
            'sku': sku,
            'product_quantity': qty,
            'option_quantity': opt_qty,
            'quantity': qty * opt_qty if (num_match or info['matched']) else qty,
            'matched': info['matched']
        })

    return items

def process_excel_file(file_path, batch_name=None):
    if not batch_name:
        batch_name = os.path.basename(file_path)

    df = pd.read_excel(file_path, dtype=str)
    col_map = detect_columns(df.columns)

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT sku, option_name, unit_price, dispatched_quantity FROM product_options')
        opt_rows = [dict(r) for r in cur.fetchall()]
    except Exception:
        opt_rows = []
    cur.execute('SELECT sku, product_name, option_name, unit_price, dispatched_quantity FROM products')
    prod_rows = [dict(r) for r in cur.fetchall()]

    try:
        cur.execute('SELECT target_product_name, target_option_name, target_sku, multiplier, source_product_name, source_option_name FROM bundle_rules WHERE is_active = 1')
        rules_cache = [dict(r) for r in cur.fetchall()]
    except Exception:
        rules_cache = []

    try:
        cur.execute('SELECT symbol FROM split_delimiters ORDER BY id ASC')
        delimiters = [r['symbol'] if (isinstance(r, dict) or hasattr(r, '__getitem__')) else r[0] for r in cur.fetchall()]
        if not delimiters:
            delimiters = ['+', '-', '/', ',', '?', '&']
    except Exception:
        delimiters = ['+', '-', '/', ',', '?', '&']

    parsed_orders = []
    preview_data = {
        'batch_name': batch_name,
        'total_excel_rows': len(df),
        'detected_columns': {k: v for k, v in col_map.items() if v},
        'orders': []
    }

    unregistered_dict = {}

    for idx, row in df.iterrows():
        def val(k):
            col = col_map.get(k)
            if col and col in row and pd.notna(row[col]):
                v = str(row[col]).strip()
                return '' if v.lower() == 'nan' else v
            return ''

        order_no = val('order_no') or f"ORD-{batch_name[:4]}-{idx+1:04d}"
        courier = val('courier') or '나르다'
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
        if recipient_phone1.endswith('.0'):
            recipient_phone1 = recipient_phone1[:-2].strip()
        if recipient_phone2.endswith('.0'):
            recipient_phone2 = recipient_phone2[:-2].strip()

        customs_id = val('customs_id')
        postal_code = val('postal_code')
        if postal_code:
            if postal_code.endswith('.0'):
                postal_code = postal_code[:-2].strip()
            # 5자리 한국 새우편번호(0xxxx)에서 앞자리 0이 누락된 경우 0 채움 복원 (예: 7668 -> 07668)
            if postal_code.isdigit() and len(postal_code) < 5:
                postal_code = postal_code.zfill(5)

        address = val('address')
        shipping_memo = val('shipping_memo')

        # 분할 처리 (캐시 사용으로 초고속)
        split_items = split_order_items(product_name, option_name, quantity, conn, opt_rows=opt_rows, prod_rows=prod_rows, rules_cache=rules_cache, delimiters=delimiters)

        for it in split_items:
            if not it.get('matched', False):
                comb = f"{it['item_product_name']} {it['item_option_name']}".strip() if (it['item_option_name'] and it['item_option_name'] != '단일옵션') else it['item_product_name']
                if comb not in unregistered_dict:
                    unregistered_dict[comb] = {
                        'product_name': it['item_product_name'],
                        'option_name': it['item_option_name'],
                        'combined_name': comb,
                        'count': 0
                    }
                unregistered_dict[comb]['count'] += (it.get('quantity') or 1)

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
    preview_data['unregistered_items'] = list(unregistered_dict.values())
    return preview_data

def save_orders_to_db(parsed_data, batch_name):
    conn = get_connection()
    cur = conn.cursor()
    saved_count = 0
    item_count = 0

    # 1. 미등록 상품들 unregistered_products 테이블에 등록/갱신
    for unreg in parsed_data.get('unregistered_items', []):
        try:
            cur.execute('''
                INSERT INTO unregistered_products (product_name, option_name, combined_name, order_count, status)
                VALUES (?, ?, ?, ?, 'pending')
                ON CONFLICT(combined_name) DO UPDATE SET
                    order_count = unregistered_products.order_count + EXCLUDED.order_count,
                    status = 'pending'
            ''', (unreg['product_name'], unreg['option_name'], unreg['combined_name'], unreg['count']))
        except Exception:
            pass

    orders = parsed_data.get('orders', [])
    all_order_items = []

    # 2. 주문 일괄 청크 삽입 (50건씩 묶어 쿼리 수 98% 절감)
    CHUNK_SIZE = 50
    for i in range(0, len(orders), CHUNK_SIZE):
        chunk = orders[i:i + CHUNK_SIZE]
        val_placeholders = []
        params = []
        for ord_data in chunk:
            val_placeholders.append('(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
            params.extend([
                ord_data['order_no'], ord_data['courier'], ord_data['tracking_no'], ord_data['shipping_method'],
                ord_data['original_product_name'], ord_data['original_option_name'], ord_data['original_quantity'],
                ord_data['recipient_name'], ord_data['recipient_phone1'], ord_data['recipient_phone2'],
                ord_data['customs_id'], ord_data['postal_code'], ord_data['address'], ord_data['shipping_memo'],
                '주문접수', batch_name
            ])
        
        insert_sql = f'''
            INSERT INTO orders (
                order_no, courier, tracking_no, shipping_method,
                original_product_name, original_option_name, original_quantity,
                recipient_name, recipient_phone1, recipient_phone2,
                customs_id, postal_code, address, shipping_memo,
                status, upload_batch
            ) VALUES {', '.join(val_placeholders)}
            RETURNING id
        '''
        cur.execute(insert_sql, params)
        returned_rows = cur.fetchall()
        for idx_c, ret in enumerate(returned_rows):
            oid = ret['id'] if (isinstance(ret, dict) or hasattr(ret, '__getitem__')) else ret[0]
            ord_data = chunk[idx_c]
            saved_count += 1
            for it in ord_data.get('split_items', []):
                all_order_items.append((
                    oid, ord_data['order_no'], it['sku'], it['item_product_name'],
                    it['item_option_name'], it.get('option_quantity', 1), it['quantity'], '준비중'
                ))

    # 3. 분할 품목(order_items) 일괄 청크 삽입 (100건씩 묶어 초고속 삽입)
    ITEM_CHUNK_SIZE = 100
    for i in range(0, len(all_order_items), ITEM_CHUNK_SIZE):
        chunk_items = all_order_items[i:i + ITEM_CHUNK_SIZE]
        val_placeholders = ['(?, ?, ?, ?, ?, ?, ?, ?)' for _ in chunk_items]
        item_params = []
        for it in chunk_items:
            item_params.extend(it)
        cur.execute(f'''
            INSERT INTO order_items (
                order_id, order_no, sku, item_product_name, item_option_name, option_quantity, quantity, status
            ) VALUES {', '.join(val_placeholders)}
        ''', item_params)
        item_count += len(chunk_items)

    conn.commit()
    conn.close()
    return {
        'saved_orders': saved_count,
        'saved_items': item_count,
        'unregistered_items': parsed_data.get('unregistered_items', [])
    }
