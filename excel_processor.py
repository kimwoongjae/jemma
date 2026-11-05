import re
import os
import pandas as pd
import json
from openpyxl import load_workbook
from database import get_connection

NARDA_MANIFEST_PROFILE = '__narda_manifest__'
NARDA_MANIFEST_HEADERS = [
    '서비스', '주문일', '배송구분', '대', '보험', '송장번호', '주문번호', '송하인',
    '받는분', '주문자e-mail', '통관고유부호/사업자번호', '전화번호', '휴대폰번호',
    '우편번호', '주소', '기타주소', '실무게(kg)', '부피무게(kg)', '가로(cm)',
    '세로(cm)', '높이(cm)', 'CTNS', 'HS Code 1', 'Item Name', '아이템상세정보',
    'Unit Price', 'Piece', 'Amount', '상품 URL', '원산지'
]

HEADER_MAP = {
    'order_no': ['주문번호', '주문 번호', 'order_no', 'Order No', '발주번호', '주문관리코드', '주문식별자'],
    'courier': ['택배사', '배송업체', '배송사', 'courier', '택배사명'],
    'tracking_no': ['송장번호', '운송장번호', 'tracking_no', '운송장 번호', '송장'],
    'shipping_method': ['배송방법', '배송방식', 'shipping_method'],
    'product_name': ['상품명', '주문상품명', '품목명', 'product_name', '상품', '판매상품명'],
    'option_name': ['옵션이름', '옵션명', '선택사항', 'option_name', '옵션정보', '옵션', '상품옵션', '선택옵션'],
    'sku': ['바코드', '상품바코드', '바코드번호', 'SKU', '상품코드', '옵션코드', 'sku'],
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


def resolve_upload_column(df_columns, reference):
    """Resolve an Excel column letter (A, H, AA) or an exact header name."""
    reference = str(reference or '').strip()
    if not reference:
        return None

    if re.fullmatch(r'[A-Za-z]+', reference):
        column_index = 0
        for char in reference.upper():
            column_index = column_index * 26 + (ord(char) - ord('A') + 1)
        column_index -= 1
        return df_columns[column_index] if 0 <= column_index < len(df_columns) else None

    normalized_reference = re.sub(r'\s+', '', reference).lower()
    for column in df_columns:
        if re.sub(r'\s+', '', str(column)).lower() == normalized_reference:
            return column
    return None


def normalize_sku(value):
    sku = str(value or '').strip()
    if sku.endswith('.0') and sku[:-2].isdigit():
        sku = sku[:-2]
    return sku


def find_product_info_by_sku(sku, opt_rows, prod_rows):
    """Find inventory option/product by exact barcode and return its default pack quantity."""
    normalized_sku = normalize_sku(sku)
    if not normalized_sku:
        return {'sku': None, 'dispatched_quantity': 1, 'unit_price': 0.0, 'matched': False}

    products_by_id = {str(row.get('id')): row for row in prod_rows if row.get('id') is not None}
    for row in opt_rows:
        if normalize_sku(row.get('sku')) != normalized_sku:
            continue
        parent = products_by_id.get(str(row.get('product_id')))
        try:
            option_quantity = max(1, int(row.get('dispatched_quantity') or 1))
        except (TypeError, ValueError):
            option_quantity = 1
        return {
            'sku': row.get('sku') or normalized_sku,
            'dispatched_quantity': option_quantity,
            'unit_price': row.get('unit_price') or (parent.get('unit_price', 0.0) if parent else 0.0),
            'english_name': parent.get('english_name', '') if parent else '',
            'hs_code': parent.get('hs_code', '') if parent else '',
            'product_url': parent.get('product_url', '') if parent else '',
            'matched': True,
            'matched_product_name': parent.get('product_name') if parent else None,
            'matched_option_name': row.get('option_name') or (parent.get('option_name') if parent else None)
        }

    for row in prod_rows:
        if normalize_sku(row.get('sku')) != normalized_sku:
            continue
        try:
            option_quantity = max(1, int(row.get('dispatched_quantity') or 1))
        except (TypeError, ValueError):
            option_quantity = 1
        return {
            'sku': row.get('sku') or normalized_sku,
            'dispatched_quantity': option_quantity,
            'unit_price': row.get('unit_price', 0.0),
            'english_name': row.get('english_name', ''),
            'hs_code': row.get('hs_code', ''),
            'product_url': row.get('product_url', ''),
            'matched': True,
            'matched_product_name': row.get('product_name'),
            'matched_option_name': row.get('option_name')
        }

    return {
        'sku': normalized_sku,
        'dispatched_quantity': 1,
        'unit_price': 0.0,
        'matched': False,
        'matched_product_name': None,
        'matched_option_name': None
    }

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

        # 상품명+옵션명 전체 조합이 일치하는 옵션을 먼저 찾는다.
        for r in opt_rows:
            opt_name = (r['option_name'] or '').strip()
            if opt_name == combined:
                return {
                    'sku': r['sku'],
                    'dispatched_quantity': r['dispatched_quantity'] or 1,
                    'unit_price': r.get('unit_price', 0.0),
                    'matched': True,
                    'matched_product_name': None,
                    'matched_option_name': r.get('option_name')
                }

        # 조합 일치가 없을 때만 단일 상품명/옵션명으로 제한적으로 찾는다.
        for r in opt_rows:
            opt_name = (r['option_name'] or '').strip()
            if (not clean_o and opt_name == product_name) or (clean_o and opt_name == option_name):
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
            if opt_name == clean_combined or (not clean_o and opt_name == clean_p):
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
            if p_name == product_name and (p_opt == option_name or (not clean_o and not p_opt)):
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

def sync_order_items_sku(conn=None, tenant_id=None):
    """기존 바코드를 우선 유지하고, 없는 경우 상품명/옵션명으로 SKU와 기본 옵션수량을 동기화합니다."""
    close_after = False
    if not conn:
        from database import get_connection
        conn = get_connection()
        close_after = True

    try:
        if tenant_id is None:
            raise ValueError('업체 정보가 없어 주문 상품 동기화를 진행할 수 없습니다.')
        cur = conn.cursor()
        try:
            cur.execute('SELECT product_id, sku, option_name, unit_price, dispatched_quantity FROM product_options WHERE tenant_id = ?', (tenant_id,))
            opt_rows = [dict(r) for r in cur.fetchall()]
        except Exception:
            opt_rows = []

        cur.execute('SELECT id, sku, product_name, option_name, unit_price, dispatched_quantity FROM products WHERE tenant_id = ?', (tenant_id,))
        prod_rows = [dict(r) for r in cur.fetchall()]

        cur.execute('SELECT id, item_product_name, item_option_name, sku, option_quantity FROM order_items WHERE tenant_id = ?', (tenant_id,))
        items = cur.fetchall()
        for it in items:
            # 업로드에서 바코드로 지정된 품목은 이후 상품명 동기화가 다른 SKU로 덮어쓰지 않게 합니다.
            info = find_product_info_by_sku(it['sku'], opt_rows, prod_rows) if it['sku'] else None
            if not info or not info['matched']:
                info = find_matching_product_info(it['item_product_name'], it['item_option_name'], conn, opt_rows=opt_rows, prod_rows=prod_rows)
            if info['matched']:
                new_sku = info['sku']
                new_opt_qty = info['dispatched_quantity'] or 1
                if new_sku != it['sku'] or (it['option_quantity'] == 1 and new_opt_qty > 1):
                    cur.execute('UPDATE order_items SET sku = ?, option_quantity = ? WHERE id = ? AND tenant_id = ?', (new_sku, new_opt_qty, it['id'], tenant_id))
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

    # 업체 설정에 등록된 기호로 나눌 수 있으면 전체 문자열 일치보다 분할을 우선한다.
    # 그렇지 않으면 재고 마스터에 합쳐진 옵션명이 있을 때 분할되지 않는 문제가 생긴다.
    if not delimiters:
        delimiters = ['+', '-', '/', ',', '?', '&']
    delimiters_pattern = '|'.join(re.escape(str(s).strip()) for s in delimiters if str(s).strip())
    option_parts = [p.strip() for p in re.split(delimiters_pattern, option_name) if p.strip()] if delimiters_pattern else [option_name]
    product_parts = [p.strip() for p in re.split(delimiters_pattern, product_name) if p.strip()] if delimiters_pattern else [product_name]
    option_is_quantity = bool(re.fullmatch(r'\d+\s*(개|봉|입|세트|box|박스)?', option_name, flags=re.IGNORECASE))
    # 공동구매 파일처럼 상품명과 옵션명 양쪽이 같은 구분 기호로 나뉜 경우
    # 같은 순번끼리 짝을 지어 품목별 상품명/옵션명을 보존합니다.
    paired_product_option_parts = (
        list(zip(product_parts, option_parts))
        if len(product_parts) > 1 and len(product_parts) == len(option_parts)
        else []
    )
    is_product_split = len(option_parts) <= 1 and len(product_parts) > 1 and (
        not option_name or option_name in ('단일옵션', '단일상품') or option_is_quantity
    )
    has_configured_split = len(option_parts) > 1 or is_product_split or bool(paired_product_option_parts)

    # 2. 설정된 기호로 나눌 품목이 없을 때만 전체 옵션명 완전 일치를 확인한다.
    direct_info = {'matched': False}
    if not has_configured_split:
        direct_info = find_matching_product_info(product_name, option_name, conn, opt_rows=opt_rows, prod_rows=prod_rows, exact_only=True)
    if direct_info.get('matched'):
        opt_qty = direct_info['dispatched_quantity'] or 1
        return [{
            # 화면에는 재고 마스터 이름이 아니라 업로드된 상품/옵션 원문을 보존한다.
            'item_product_name': product_name,
            'item_option_name': option_name,
            'sku': direct_info['sku'],
            'product_quantity': qty,
            'option_quantity': opt_qty,
            'quantity': qty * opt_qty,
            'matched': True
        }]

    # 3. 구분 기호(+, -, ,, /, ?, & 등)로 자동 다중 품목 분할
    items = []
    parts = product_parts if is_product_split else option_parts

    if paired_product_option_parts:
        for product_part, option_part in paired_product_option_parts:
            info = find_matching_product_info(product_part, option_part, conn, opt_rows=opt_rows, prod_rows=prod_rows)
            opt_qty = info['dispatched_quantity'] if info['matched'] else 1
            items.append({
                'item_product_name': product_part,
                'item_option_name': option_part,
                'sku': info['sku'] if info['matched'] else generate_sku(product_part, option_part),
                'product_quantity': qty,
                'option_quantity': opt_qty,
                'quantity': qty * opt_qty,
                'matched': info['matched']
            })
    elif len(parts) > 1:
        for p in parts:
            lookup_p = '' if is_product_split else product_name
            info = find_matching_product_info(lookup_p, p, conn, opt_rows=opt_rows, prod_rows=prod_rows)
            sku = info['sku'] if info['matched'] else generate_sku(product_name, p)
            opt_qty = info['dispatched_quantity'] if info['matched'] else 1
            items.append({
                'item_product_name': product_name if not is_product_split else p,
                'item_option_name': p if not is_product_split else '단일옵션',
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
            'item_product_name': product_name,
            'item_option_name': option_name,
            'sku': sku,
            'product_quantity': qty,
            'option_quantity': opt_qty,
            'quantity': qty * opt_qty if (num_match or info['matched']) else qty,
            'matched': info['matched']
        })

    return items

def process_excel_file(file_path, batch_name=None, column_mapping=None, header_row=1, tenant_id=None):
    if not batch_name:
        batch_name = os.path.basename(file_path)

    try:
        header_row = max(0, int(header_row))
    except (TypeError, ValueError):
        header_row = 1
    df = pd.read_excel(file_path, dtype=str, header=None if header_row == 0 else header_row - 1)
    col_map = detect_columns(df.columns)
    invalid_profile_columns = []
    for key, reference in (column_mapping or {}).items():
        if key not in HEADER_MAP or not str(reference or '').strip():
            continue
        resolved = resolve_upload_column(df.columns, reference)
        if resolved is None:
            invalid_profile_columns.append(f'{key}={reference}')
        else:
            col_map[key] = resolved
    if invalid_profile_columns:
        raise ValueError('업로드 형식에서 지정한 열을 파일에서 찾을 수 없습니다: ' + ', '.join(invalid_profile_columns))

    if tenant_id is None:
        raise ValueError('업체 정보가 없어 주문 파일을 처리할 수 없습니다.')
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT product_id, sku, option_name, unit_price, dispatched_quantity FROM product_options WHERE tenant_id = ?', (tenant_id,))
        opt_rows = [dict(r) for r in cur.fetchall()]
    except Exception:
        opt_rows = []
    cur.execute('SELECT id, sku, product_name, option_name, english_name, hs_code, product_url, unit_price, dispatched_quantity FROM products WHERE tenant_id = ?', (tenant_id,))
    prod_rows = [dict(r) for r in cur.fetchall()]

    try:
        cur.execute('SELECT target_product_name, target_option_name, target_sku, multiplier, source_product_name, source_option_name FROM bundle_rules WHERE tenant_id = ? AND is_active = 1', (tenant_id,))
        rules_cache = [dict(r) for r in cur.fetchall()]
    except Exception:
        rules_cache = []

    try:
        cur.execute('SELECT symbol FROM split_delimiters WHERE tenant_id = ? ORDER BY id ASC', (tenant_id,))
        delimiters = [r['symbol'] if (isinstance(r, dict) or hasattr(r, '__getitem__')) else r[0] for r in cur.fetchall()]
        if not delimiters:
            delimiters = ['+', '-', '/', ',', '?', '&']
    except Exception:
        delimiters = ['+', '-', '/', ',', '?', '&']

    parsed_orders = []
    preview_data = {
        'batch_name': batch_name,
        'total_excel_rows': len(df),
        'detected_columns': {k: str(v) for k, v in col_map.items() if v is not None and v != ''},
        'orders': []
    }

    unregistered_dict = {}

    for idx, row in df.iterrows():
        def val(k):
            col = col_map.get(k)
            if col is not None and col in row and pd.notna(row[col]):
                v = str(row[col]).strip()
                return '' if v.lower() == 'nan' else v
            return ''

        order_no = val('order_no') or f"ORD-{batch_name[:4]}-{idx+1:04d}"
        courier = val('courier') or '나르다'
        tracking_no = val('tracking_no')
        shipping_method = val('shipping_method') or '택배'
        product_name = val('product_name')
        option_name = val('option_name')
        barcode = normalize_sku(val('sku'))
        
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

        # 업로드 파일에 바코드가 있으면 상품명/옵션명 분할보다 바코드를 우선합니다.
        if barcode:
            barcode_info = find_product_info_by_sku(barcode, opt_rows, prod_rows)
            try:
                option_quantity = max(1, int(barcode_info.get('dispatched_quantity') or 1))
            except (TypeError, ValueError):
                option_quantity = 1
            split_items = [{
                'item_product_name': product_name or barcode_info.get('matched_product_name') or '상품명미상',
                'item_option_name': option_name or barcode_info.get('matched_option_name') or '단일옵션',
                'sku': barcode_info.get('sku') or barcode,
                'product_quantity': quantity,
                'option_quantity': option_quantity,
                'quantity': quantity * option_quantity,
                'matched': barcode_info.get('matched', False)
            }]
        else:
            # 바코드가 없으면 기존 상품명/옵션명 분할 및 자동 매칭을 사용합니다.
            split_items = split_order_items(product_name, option_name, quantity, conn, opt_rows=opt_rows, prod_rows=prod_rows, rules_cache=rules_cache, delimiters=delimiters)

        # 자동 열 인식/저장 형식은 기존 재고 매칭 결과에서 통관용 상품 정보를 보완합니다.
        # 나르다 기본형식 전용 파서는 이 경로에 들어오지 않아 업로드 값을 그대로 유지합니다.
        for item in split_items:
            info = find_product_info_by_sku(item.get('sku'), opt_rows, prod_rows) if item.get('sku') else {}
            if info.get('matched'):
                item['english_name'] = info.get('english_name') or ''
                item['hs_code'] = info.get('hs_code') or ''
                item['unit_price'] = info.get('unit_price') or 0
                item['product_url'] = info.get('product_url') or ''

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


def process_narda_manifest_file(file_path, batch_name=None, tenant_id=None):
    """Parse the fixed NARDA ManifestTemplate layout without inventory enrichment."""
    if not batch_name:
        batch_name = os.path.basename(file_path)
    if tenant_id is None:
        raise ValueError('업체 정보가 없어 주문 파일을 처리할 수 없습니다.')

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    sheet = workbook.active
    if sheet.max_column < 30:
        workbook.close()
        raise ValueError('나르다 기본형식 파일의 열 구성이 올바르지 않습니다. ManifestTemplate 파일인지 확인해주세요.')

    def clean(value):
        if value is None:
            return ''
        text = str(value).strip()
        if text.lower() in ('none', 'nan'):
            return ''
        if text.endswith('.0') and text[:-2].isdigit():
            text = text[:-2]
        return text

    header_values = [clean(sheet.cell(1, col).value) for col in range(1, 31)]
    if header_values[6] != '주문번호' or header_values[22] not in ('HS Code', 'HS Code 1') or header_values[23] != 'Item Name':
        workbook.close()
        raise ValueError('나르다 기본형식은 ManifestTemplate 파일만 지원합니다. 업로드 파일 형식을 확인해주세요.')

    # ManifestTemplate2의 열 순서는 고정되어 있어 제목 자동 인식이나 재고 검색을 하지 않습니다.
    # 주문 단위 정보는 다음 품목 행에서 비어 있을 수 있으므로 주문 그룹 안에서만 이어받습니다.
    order_level_indices = set(range(22))  # A:V
    item_name_idx = 23                    # X: Item Name
    order_groups = {}
    current_order_no = ''
    current_order_values = [''] * 30
    meaningful_rows = 0

    for row in sheet.iter_rows(min_row=2, max_col=30, values_only=True):
        values = [clean(value) for value in row]
        if len(values) < 30:
            values.extend([''] * (30 - len(values)))

        explicit_order_no = values[6]
        if explicit_order_no:
            current_order_no = explicit_order_no
            current_order_values = [''] * 30
        if not current_order_no:
            continue

        for col_idx in order_level_indices:
            if values[col_idx]:
                current_order_values[col_idx] = values[col_idx]

        # 서식만 적용된 빈 행이나 데이터가 끊긴 꼬리 행은 무시합니다.
        if not values[item_name_idx] and not values[24] and not values[22]:
            continue

        order_no = current_order_no
        row_fields = {}
        for col_idx, header in enumerate(NARDA_MANIFEST_HEADERS):
            cell_value = values[col_idx]
            if col_idx in order_level_indices and not cell_value:
                cell_value = current_order_values[col_idx]
            row_fields[header] = cell_value

        if not row_fields['Item Name']:
            continue

        try:
            piece = max(1, int(float(row_fields['Piece'] or 1)))
        except (TypeError, ValueError):
            piece = 1
        try:
            option_price = float(row_fields['Unit Price'] or 0)
        except (TypeError, ValueError):
            option_price = 0
        try:
            order_weight = float(row_fields['실무게(kg)'].replace(',', '') or 0)
        except (TypeError, ValueError):
            order_weight = 0

        order = order_groups.get(order_no)
        if order is None:
            postal = row_fields['우편번호']
            if postal.isdigit() and len(postal) < 5:
                postal = postal.zfill(5)
            order = {
                'order_no': order_no,
                'courier': '나르다',
                'tracking_no': row_fields['송장번호'],
                'shipping_method': row_fields['배송구분'] or '택배',
                'original_product_name': '',
                'original_option_name': '',
                'original_quantity': 0,
                'weight': order_weight,
                'recipient_name': row_fields['받는분'],
                'recipient_phone1': row_fields['휴대폰번호'] or row_fields['전화번호'],
                'recipient_phone2': row_fields['전화번호'] if row_fields['휴대폰번호'] else '',
                'customs_id': row_fields['통관고유부호/사업자번호'],
                'postal_code': postal,
                'address': ' '.join(part for part in (row_fields['주소'], row_fields['기타주소']) if part),
                'shipping_memo': '',
                'split_items': [],
                'is_split': False,
                'manifest_data': {header: row_fields[header] for header in NARDA_MANIFEST_HEADERS[:22]},
            }
            order_groups[order_no] = order

        item = {
            'item_product_name': row_fields['Item Name'],
            'item_option_name': row_fields['아이템상세정보'],
            'sku': '',
            'product_quantity': piece,
            'option_quantity': 1,
            'quantity': piece,
            'matched': False,
            'english_name': row_fields['Item Name'],
            'hs_code': row_fields['HS Code 1'],
            'unit_price': option_price,
            'manifest_data': row_fields,
        }
        order['split_items'].append(item)
        order['original_quantity'] += piece
        meaningful_rows += 1

    workbook.close()
    if not order_groups:
        raise ValueError('나르다 기본형식 파일에서 주문번호와 Item Name이 포함된 품목을 찾지 못했습니다.')

    for order in order_groups.values():
        items = order['split_items']
        order['original_product_name'] = ' / '.join(dict.fromkeys(item['item_product_name'] for item in items))
        order['original_option_name'] = ' / '.join(dict.fromkeys(item['item_option_name'] for item in items if item['item_option_name']))
        order['is_split'] = len(items) > 1

    parsed_orders = list(order_groups.values())
    return {
        'batch_name': batch_name,
        'upload_profile': NARDA_MANIFEST_PROFILE,
        'manifest_mode': True,
        'total_excel_rows': meaningful_rows,
        'total_split_items': meaningful_rows,
        'detected_columns': {f'{chr(65 + i)}': header for i, header in enumerate(NARDA_MANIFEST_HEADERS)},
        'orders': parsed_orders,
        'unregistered_items': [],
    }

def save_orders_to_db(parsed_data, batch_name, tenant_id=None):
    if tenant_id is None:
        raise ValueError('업체 정보가 없어 주문을 저장할 수 없습니다.')
    conn = get_connection()
    cur = conn.cursor()
    saved_count = 0
    item_count = 0

    # 1. 미등록 상품들 unregistered_products 테이블에 등록/갱신
    for unreg in parsed_data.get('unregistered_items', []):
        try:
            cur.execute('''
                INSERT INTO unregistered_products (tenant_id, product_name, option_name, combined_name, order_count, status)
                VALUES (?, ?, ?, ?, ?, 'pending')
                ON CONFLICT(tenant_id, combined_name) DO UPDATE SET
                    order_count = unregistered_products.order_count + EXCLUDED.order_count,
                    status = 'pending'
            ''', (tenant_id, unreg['product_name'], unreg['option_name'], unreg['combined_name'], unreg['count']))
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
            val_placeholders.append('(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
            params.extend([
                tenant_id,
                ord_data['order_no'], ord_data['courier'], ord_data['tracking_no'], ord_data['shipping_method'],
                ord_data['original_product_name'], ord_data['original_option_name'], ord_data['original_quantity'],
                ord_data['recipient_name'], ord_data['recipient_phone1'], ord_data['recipient_phone2'],
                ord_data['customs_id'], ord_data['postal_code'], ord_data['address'], ord_data['shipping_memo'],
                '주문접수', batch_name, float(ord_data.get('weight') or 0),
                json.dumps(ord_data.get('manifest_data') or {}, ensure_ascii=False)
            ])
        
        insert_sql = f'''
            INSERT INTO orders (
                tenant_id, order_no, courier, tracking_no, shipping_method,
                original_product_name, original_option_name, original_quantity,
                recipient_name, recipient_phone1, recipient_phone2,
                customs_id, postal_code, address, shipping_memo,
                status, upload_batch, weight, manifest_data
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
                    it['item_option_name'], it.get('product_quantity', ord_data.get('original_quantity', 1)),
                    it.get('option_quantity', 1), it['quantity'], '준비중',
                    json.dumps(it.get('manifest_data') or {}, ensure_ascii=False)
                ))

    # 3. 분할 품목(order_items) 일괄 청크 삽입 (100건씩 묶어 초고속 삽입)
    ITEM_CHUNK_SIZE = 100
    for i in range(0, len(all_order_items), ITEM_CHUNK_SIZE):
        chunk_items = all_order_items[i:i + ITEM_CHUNK_SIZE]
        val_placeholders = ['(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)' for _ in chunk_items]
        item_params = []
        for it in chunk_items:
            item_params.extend([tenant_id] + list(it))
        cur.execute(f'''
            INSERT INTO order_items (
                tenant_id, order_id, order_no, sku, item_product_name, item_option_name, product_quantity, option_quantity, quantity, status, manifest_data
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
