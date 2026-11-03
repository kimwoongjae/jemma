import random
import time
import re
import requests
import json
import concurrent.futures
import threading
from datetime import datetime
import openpyxl
from io import BytesIO
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from database import get_connection

# 스레드별 HTTP 세션을 사용해 동시 요청 사이의 세션/쿠키 공유를 방지합니다.
_narda_thread_local = threading.local()


def get_narda_session():
    session = getattr(_narda_thread_local, 'session', None)
    if session is None:
        session = requests.Session()
        adapter = HTTPAdapter(pool_connections=4, pool_maxsize=4, max_retries=Retry(total=2, backoff_factor=0.3))
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        _narda_thread_local.session = session
    return session

COURIER_PREFIX = {
    '나르다': '3054',
    'CJ대한통운': '6892',
    '우체국택배': '6098',
    '한진택배': '5081',
    '롯데택배': '4023',
    '로젠택배': '9012'
}

def generate_mock_tracking_number(courier_name='나르다'):
    prefix = COURIER_PREFIX.get(courier_name, '3054')
    # 12자리 유효 형태의 송장번호 생성
    rand_mid = f"{random.randint(1000000, 9999999):07d}"
    base_no = prefix + rand_mid
    # 검증번호 (modulo 7 알고리즘 등)
    check_digit = sum(int(c) for c in base_no) % 7
    return f"{base_no}{check_digit}"

def call_narda_regist_order(ord_row, items, settings):
    """
    나르다(NARDA JP) 주문 등록 및 송장 발급 API 호출
    http://api.narda-jp.com/api/registOrder
    """
    # 테스트 주문 전송 원천 차단
    order_no_check = str(ord_row.get('order_no') or '').strip().upper()
    recipient_check = str(ord_row.get('recipient_name') or '').strip()
    if order_no_check.startswith('TEST') or recipient_check in ('홍길동', '테스트'):
        return {"success": False, "error": "테스트 데이터는 나르다 API로 전송되지 않습니다."}

    url = "http://api.narda-jp.com/api/registOrder"

    # 전화번호 하이픈 제거
    tel = re.sub(r'[^0-9]', '', str(ord_row.get('recipient_phone1') or ''))
    htel = re.sub(r'[^0-9]', '', str(ord_row.get('recipient_phone2') or '')) or tel
    weight = float(ord_row['weight']) if (ord_row.get('weight') and ord_row['weight'] > 0) else 0.5

    api_items = []
    for it in items:
        # 상품수량(엑셀 원본) * 옵션수량(세트 낱개수) = 최종 구매/출고 수량
        opt_qty = int(it.get('option_quantity') or 1)
        prod_qty = int(it.get('product_quantity') or ord_row.get('original_quantity') or 1)
        qty = opt_qty * prod_qty

        price = float(it.get('unit_price') or 10.0)
        hs = str(it.get('hs_code') or '630690').replace('.', '').strip() or '630690'
        eng_name = (it.get('english_name') or '').strip() or it.get('item_product_name') or 'Product'
        item_url = it.get('product_url') or ''

        api_items.append({
            "hs_code": hs,
            "item_name": eng_name,
            "item_price": price,
            "item_cnt": qty,
            "item_total_price": round(price * qty, 2),
            "item_detail": None,
            "item_url": item_url,
            "item_origin": "JP"
        })

    if not api_items:
        api_items.append({
            "hs_code": "630690",
            "item_name": ord_row.get('original_product_name') or "Item",
            "item_price": 10.0,
            "item_cnt": ord_row.get('original_quantity') or 1,
            "item_total_price": 10.0 * (ord_row.get('original_quantity') or 1),
            "item_detail": None,
            "item_url": "",
            "item_origin": "JP"
        })

    # 송하인은 택배설정의 고객사 코드(customer_code) 값으로 전송
    sender = (settings.get('customer_code') or '').strip() or "정마무역"
    narda_id = (settings.get('narda_id') or '').strip() or "feelcoach33"
    narda_api_key = (settings.get('api_key') or '').strip() or "e1df1b2a432106cd6f4c886dfb9c3ca24fbf27e9"

    payload = {
        "order_no": ord_row['order_no'],
        "sender": sender,
        "order_email": "",
        "recipient": ord_row['recipient_name'],
        "customs_no": (ord_row.get('customs_id') or '').strip().upper(),
        "customs_type": "p",
        "tel": tel,
        "recipient_htel": htel,
        "zip_no": ord_row.get('postal_code') or "",
        "address1": ord_row.get('address') or "",
        "address2": ord_row.get('shipping_memo') or "문앞에 두고 가세요",
        "ctns": 1,
        "service": "NARDA AIR",
        "volume_list": [
            {
                "weight": weight,
                "volume_width": 0,
                "volume_height": 0,
                "volume_length": 0
            }
        ],
        "items": api_items,
        "narda_id": narda_id,
        "narda_api_key": narda_api_key
    }

    try:
        resp = get_narda_session().post(
            url,
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=25
        )
        res = resp.json()
        if res.get("resultCode") == 0 and res.get("result"):
            r = res["result"][0]
            if r.get("result") == "success":
                deli_invoice = r.get("deli_invoice")
                pdf_urls = r.get("url", [])
                pdf_url = pdf_urls[0] if pdf_urls else ""
                return {"success": True, "tracking_no": deli_invoice, "pdf_url": pdf_url}
            else:
                return {"success": False, "error": r.get("error_msg") or r.get("errorMsg") or "나르다 등록 실패"}
        else:
            return {"success": False, "error": res.get("errorMsg") or res.get("error_msg") or "나르다 API 오류"}
    except Exception as e:
        return {"success": False, "error": str(e)}

def issue_tracking_numbers(order_ids, courier_name=None):
    conn = get_connection()
    cur = conn.cursor()

    # 택배사 설정 조회
    cur.execute('SELECT * FROM courier_settings LIMIT 1')
    settings_row = cur.fetchone()
    settings = dict(settings_row) if settings_row else {}

    target_courier = courier_name or settings.get('courier_name') or '나르다'
    is_mock = settings.get('is_mock', 0)
    updated_orders = []
    failed_orders = []
    aggregated_deductions = {}

    # 전체 상품 마스터 캐시 (SKU 및 상품명 매칭용)
    cur.execute('SELECT id, sku, product_name, option_name, english_name, hs_code, unit_price, product_url, stock_quantity FROM products')
    prod_rows = [dict(r) for r in cur.fetchall()]
    prod_by_sku = {r['sku']: r for r in prod_rows if r['sku']}
    prod_by_name = {r['product_name']: r for r in prod_rows if r['product_name']}

    # 1. 선택 주문과 품목을 묶음 조회해 원격 DB 왕복을 줄입니다.
    order_ids = list(dict.fromkeys(int(oid) for oid in order_ids if str(oid).isdigit()))
    orders_data = []
    if not order_ids:
        conn.close()
        return {'updated_orders': [], 'failed_orders': []}
    placeholders = ','.join(['?'] * len(order_ids))
    cur.execute(f'SELECT * FROM orders WHERE id IN ({placeholders})', order_ids)
    orders_by_id = {int(row['id']): dict(row) for row in cur.fetchall()}
    cur.execute(f'SELECT * FROM order_items WHERE order_id IN ({placeholders}) ORDER BY id ASC', order_ids)
    items_by_order = {}
    for row in cur.fetchall():
        row_dict = dict(row)
        items_by_order.setdefault(int(row_dict['order_id']), []).append(row_dict)

    for oid in order_ids:
        ord_dict = orders_by_id.get(oid)
        if not ord_dict:
            continue
        raw_items = items_by_order.get(oid, [])
        existing_tracking = ord_dict.get('tracking_no') or ''
        existing_pdf = ord_dict.get('invoice_pdf_url') or ''

        if target_courier == '나르다' and is_mock == 0:
            if existing_tracking and existing_pdf and existing_tracking.startswith('30'):
                orders_data.append({
                    'oid': oid,
                    'ord_dict': ord_dict,
                    'items': [],
                    'deduct_items': raw_items,
                    'is_cached': True,
                    'tracking_no': existing_tracking,
                    'pdf_url': existing_pdf
                })
                continue

            items = []
            for raw_item in raw_items:
                it = dict(raw_item)
                matched_prod = prod_by_sku.get(it.get('sku')) or prod_by_name.get(it.get('item_product_name'))
                if matched_prod:
                    it['english_name'] = matched_prod.get('english_name') or it.get('item_product_name')
                    it['hs_code'] = matched_prod.get('hs_code') or '630690'
                    it['unit_price'] = matched_prod.get('unit_price') or 10.0
                    it['product_url'] = matched_prod.get('product_url') or ''
                else:
                    it['english_name'] = it.get('item_product_name') or 'Product'
                    it['hs_code'] = '630690'
                    it['unit_price'] = 10.0
                    it['product_url'] = ''
                items.append(it)

            orders_data.append({
                'oid': oid,
                'ord_dict': ord_dict,
                'items': items,
                'deduct_items': raw_items,
                'is_cached': False
            })
        else:
            orders_data.append({
                'oid': oid,
                'ord_dict': ord_dict,
                'items': [],
                'deduct_items': raw_items,
                'is_cached': False,
                'is_mock': True
            })

    stock_only_dispatch = str(settings.get('stock_only_dispatch') or '0').strip().lower() in ('1', 'true', 't', 'yes', 'on')
    if stock_only_dispatch and orders_data:
        # 선택 주문 전체의 필요 수량을 주문 순서대로 예약 계산해, 같은 SKU를
        # 여러 주문이 함께 선택해도 보유 재고를 초과하는 주문은 API에 보내지 않습니다.
        remaining_stock = {
            int(p['id']): max(0, int(float(p.get('stock_quantity') or 0)))
            for p in prod_rows if p.get('id') is not None
        }
        dispatchable_orders = []
        for order_data in orders_data:
            order = order_data['ord_dict']
            requirements = {}
            product_labels = {}
            shortage_messages = []

            for item in order_data.get('deduct_items', []):
                if int(item.get('stock_deducted') or 0):
                    continue
                sku = str(item.get('sku') or '').strip()
                product = prod_by_sku.get(sku) or prod_by_name.get(item.get('item_product_name'))
                try:
                    needed = max(1, int(float(item.get('quantity') or 1)))
                except (TypeError, ValueError):
                    needed = 1
                name = str(item.get('item_product_name') or item.get('item_option_name') or sku or '상품').strip()

                if not product:
                    shortage_messages.append(f'{name} (재고 미등록, 필요 {needed}개)')
                    continue

                product_id = int(product['id'])
                requirements[product_id] = requirements.get(product_id, 0) + needed
                product_labels[product_id] = (name, sku or str(product.get('sku') or ''))

            for product_id, needed in requirements.items():
                available = remaining_stock.get(product_id, 0)
                if needed > available:
                    name, sku = product_labels[product_id]
                    barcode_note = f' / 바코드 {sku}' if sku else ''
                    shortage_messages.append(f'{name}{barcode_note} (필요 {needed}개, 현재고 {available}개)')

            if shortage_messages:
                shown = '; '.join(shortage_messages[:3])
                if len(shortage_messages) > 3:
                    shown += f' 외 {len(shortage_messages) - 3}개 품목'
                error = (
                    f'재고 부족으로 송장을 발급하지 않았습니다: {shown}. '
                    "출고가 필요하면 [설정]에서 '재고가 있을 때만 출고' 스위치를 끄고 다시 시도하세요."
                )
                cur.execute('''
                    UPDATE orders SET api_dispatch_status = '실패', api_dispatch_message = ?
                    WHERE id = ?
                ''', (error, order_data['oid']))
                failed_orders.append({
                    'order_id': order_data['oid'],
                    'order_no': order.get('order_no'),
                    'recipient_name': order.get('recipient_name'),
                    'error': error
                })
                continue

            for product_id, needed in requirements.items():
                remaining_stock[product_id] = remaining_stock.get(product_id, 0) - needed
            dispatchable_orders.append(order_data)

        orders_data = dispatchable_orders

    # 전송 중 상태를 먼저 저장합니다. API 실패가 나도 주문상태와 재고는 건드리지 않습니다.
    if orders_data:
        active_ids = [item['oid'] for item in orders_data]
        active_placeholders = ','.join(['?'] * len(active_ids))
        cur.execute(f'''
            UPDATE orders
            SET api_dispatch_status = '전송중', api_dispatch_message = '나르다 API 응답 대기 중'
            WHERE id IN ({active_placeholders})
        ''', active_ids)
    conn.commit()

    if not orders_data:
        conn.close()
        return {'updated_orders': [], 'failed_orders': failed_orders}

    # 실제 나르다 API 전송 10건을 동시 5건으로 확인했고 모두 성공하여,
    # 처리량을 높이되 요청 시간 초과 위험을 줄이도록 동시 요청을 5건으로 제한합니다.
    def call_api_single(item):
        if item.get('is_cached'):
            return (item['oid'], item['ord_dict'], {
                'success': True,
                'tracking_no': item['tracking_no'],
                'pdf_url': item['pdf_url']
            })
        if item.get('is_mock'):
            mock_no = generate_mock_tracking_number(target_courier)
            return (item['oid'], item['ord_dict'], {'success': True, 'tracking_no': mock_no, 'pdf_url': ''})

        api_res = call_narda_regist_order(item['ord_dict'], item['items'], settings)
        return (item['oid'], item['ord_dict'], api_res)

    workers_count = min(5, len(orders_data)) if orders_data else 1
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers_count) as executor:
        api_results = list(executor.map(call_api_single, orders_data))

    # 3. 결과 DB 일괄 반영 및 재고 차감 처리 (메인 스레드 안전 트랜잭션)
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    for oid, ord_dict, api_res in api_results:
        if not api_res.get('success'):
            err_msg = api_res.get('error') or '등록 실패'
            print(f"나르다 API 발급 실패 (주문번호 {ord_dict['order_no']}): {err_msg}")
            cur.execute('''
                UPDATE orders SET api_dispatch_status = '실패', api_dispatch_message = ?
                WHERE id = ?
            ''', (str(err_msg), oid))
            failed_orders.append({
                'order_id': oid,
                'order_no': ord_dict['order_no'],
                'recipient_name': ord_dict['recipient_name'],
                'error': err_msg
            })
            continue

        tracking_no = api_res.get('tracking_no')
        pdf_url = api_res.get('pdf_url', '')
        if not tracking_no or not str(tracking_no).strip():
            err_msg = '나르다 응답에 송장번호가 없어 발송 완료로 처리하지 않았습니다.'
            cur.execute('''
                UPDATE orders SET api_dispatch_status = '실패', api_dispatch_message = ?
                WHERE id = ?
            ''', (err_msg, oid))
            failed_orders.append({
                'order_id': oid,
                'order_no': ord_dict['order_no'],
                'recipient_name': ord_dict['recipient_name'],
                'error': err_msg
            })
            continue
        tracking_no = str(tracking_no).strip()

        cur.execute('''
            UPDATE orders
            SET tracking_no = ?, courier = ?, invoice_pdf_url = ?, status = '출고완료', dispatched_at = ?,
                api_dispatch_status = '성공', api_dispatch_message = ''
            WHERE id = ?
        ''', (tracking_no, target_courier, pdf_url, now_str, oid))

        # ==========================================
        # 품목별 재고 차감 및 바코드별 집계 (중복 차감 방지)
        # ==========================================
        items_to_deduct = [
            it for it in items_by_order.get(int(oid), [])
            if not int(it.get('stock_deducted') or 0)
        ]

        for it in items_to_deduct:
            item_sku = it.get('sku')
            matched_prod = prod_by_sku.get(item_sku) or prod_by_name.get(it.get('item_product_name'))
            if not matched_prod and item_sku:
                cur.execute('SELECT id, sku, product_name, option_name, stock_quantity FROM products WHERE sku = ?', (item_sku,))
                pr = cur.fetchone()
                if pr:
                    matched_prod = dict(pr)

            if matched_prod:
                deduct_qty = int(it.get('quantity') or 1)
                cur_stock = matched_prod.get('stock_quantity') or 0
                new_stock = cur_stock - deduct_qty

                cur.execute('UPDATE products SET stock_quantity = ? WHERE id = ?', (new_stock, matched_prod['id']))
                matched_prod['stock_quantity'] = new_stock

                # 바코드(SKU) 기준 일괄 집계 누적 (개별 주문마다 기록하지 않고 묶음당 1건만 기록)
                sku_key = matched_prod['sku'] or str(matched_prod['id'])
                log_key = (oid, sku_key)
                if log_key not in aggregated_deductions:
                    aggregated_deductions[log_key] = {
                        'product_id': matched_prod['id'],
                        'sku': matched_prod['sku'],
                        'product_name': matched_prod['product_name'],
                        'option_name': matched_prod.get('option_name') or '',
                        'order_no': ord_dict['order_no'],
                        'total_qty': 0,
                        'final_stock': new_stock,
                        'courier': target_courier
                    }
                aggregated_deductions[log_key]['total_qty'] += deduct_qty
                aggregated_deductions[log_key]['final_stock'] = new_stock

            cur.execute("UPDATE order_items SET stock_deducted = 1, status = '출고완료' WHERE id = ?", (it['id'],))

        updated_orders.append({
            'order_id': oid,
            'order_no': ord_dict['order_no'],
            'recipient_name': ord_dict['recipient_name'],
            'courier': target_courier,
            'tracking_no': tracking_no,
            'invoice_pdf_url': pdf_url
        })

    # ==========================================
    # 주문별·바코드별로 출고 내역을 남겨 처리 건을 추적할 수 있도록 기록
    # ==========================================
    now_log_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    for agg in aggregated_deductions.values():
        memo_text = f"주문 출고 ({agg['courier']} 송장발급)"
        cur.execute('''
            INSERT INTO inventory_logs (
                product_id, sku, product_name, option_name,
                change_type, change_qty, resulting_stock, reference_no, memo, created_at
            ) VALUES (?, ?, ?, ?, '주문출고', ?, ?, ?, ?, ?)
        ''', (
            agg['product_id'], agg['sku'], agg['product_name'], agg['option_name'],
            -agg['total_qty'], agg['final_stock'], agg['order_no'], memo_text, now_log_str
        ))

    conn.commit()
    conn.close()
    return {
        'updated_orders': updated_orders,
        'failed_orders': failed_orders
    }

def complete_dispatch(order_ids):
    """
    발송 완료 처리:
    1. 주문 상태를 '출고완료'로 변경
    2. 연결된 order_items 품목들의 재고를 실제 차감 (stock_deducted = 1)
    3. inventory_logs 에 바코드 기준 총수량 1건만 '주문출고' 로그 기록
    """
    conn = get_connection()
    cur = conn.cursor()
    dispatched_count = 0
    deducted_items = 0
    aggregated_deductions = {}

    for oid in order_ids:
        cur.execute('SELECT * FROM orders WHERE id = ?', (oid,))
        ord_row = cur.fetchone()
        if not ord_row:
            continue

        now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
        cur.execute("UPDATE orders SET status = '출고완료', dispatched_at = ? WHERE id = ?", (now_str, oid))
        dispatched_count += 1

        cur.execute('SELECT * FROM order_items WHERE order_id = ? AND stock_deducted = 0', (oid,))
        items = cur.fetchall()

        for it in items:
            cur.execute('SELECT * FROM products WHERE sku = ?', (it['sku'],))
            prod = cur.fetchone()
            if prod:
                qty = it['quantity']
                new_stock = prod['stock_quantity'] - qty
                cur.execute('UPDATE products SET stock_quantity = ? WHERE id = ?', (new_stock, prod['id']))
                
                sku_key = prod['sku'] or str(prod['id'])
                if sku_key not in aggregated_deductions:
                    aggregated_deductions[sku_key] = {
                        'product_id': prod['id'],
                        'sku': prod['sku'],
                        'product_name': prod['product_name'],
                        'option_name': prod.get('option_name') or '',
                        'total_qty': 0,
                        'final_stock': new_stock,
                        'order_nos': set()
                    }
                aggregated_deductions[sku_key]['total_qty'] += qty
                aggregated_deductions[sku_key]['final_stock'] = new_stock
                aggregated_deductions[sku_key]['order_nos'].add(ord_row['order_no'])

            cur.execute("UPDATE order_items SET stock_deducted = 1, status = '출고완료' WHERE id = ?", (it['id'],))
            deducted_items += 1

    # 바코드 기준 1건만 기록
    now_log_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    for agg in aggregated_deductions.values():
        cnt = len(agg['order_nos'])
        memo_text = f"총 {cnt}건 주문 일괄 출고"
        ref_no = f"{cnt}건 일괄출고"
        cur.execute('''
            INSERT INTO inventory_logs (
                product_id, sku, product_name, option_name,
                change_type, change_qty, resulting_stock, reference_no, memo, created_at
            ) VALUES (?, ?, ?, ?, '주문출고', ?, ?, ?, ?, ?)
        ''', (
            agg['product_id'], agg['sku'], agg['product_name'], agg['option_name'],
            -agg['total_qty'], agg['final_stock'], ref_no, memo_text, now_log_str
        ))

    conn.commit()
    conn.close()
    return {'dispatched_orders': dispatched_count, 'deducted_items': deducted_items}

def export_shipping_excel(order_ids=None, column_order=None):
    """
    발송 완료된 주문 목록을 원본 엑셀 형식에 송장번호가 채워진 파일로 다운로드
    """
    conn = get_connection()
    cur = conn.cursor()

    if order_ids and len(order_ids) > 0:
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT * FROM orders WHERE id IN ({placeholders}) ORDER BY id ASC', order_ids)
        orders = cur.fetchall()
    else:
        cur.execute('SELECT * FROM orders ORDER BY id DESC')
        orders = cur.fetchall()

    order_ids_for_items = [row['id'] for row in orders]
    items_by_order = {}
    if order_ids_for_items:
        placeholders = ','.join(['?'] * len(order_ids_for_items))
        cur.execute(f'SELECT * FROM order_items WHERE order_id IN ({placeholders}) ORDER BY id ASC', order_ids_for_items)
        for item in cur.fetchall():
            items_by_order.setdefault(item['order_id'], []).append(dict(item))
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "발송처리결과"

    def product_quantity(order, item):
        try:
            value = int(item.get('product_quantity') or order['original_quantity'] or 0)
            if value > 0:
                return value
        except (TypeError, ValueError):
            pass
        try:
            option_qty = max(1, int(item.get('option_quantity') or 1))
            buy_qty = max(1, int(item.get('quantity') or 1))
            return max(1, round(buy_qty / option_qty))
        except (TypeError, ValueError):
            return 1

    column_definitions = {
        'weight': ('무게 (kg)', lambda o, its: o['weight'] if ('weight' in o.keys() and o['weight']) else 0.0),
        'order_no': ('주문번호', lambda o, its: o['order_no'] or ''),
        'dispatched_at': ('발송 일시', lambda o, its: o['dispatched_at'] or ''),
        'recipient': ('수취인 / 연락처', lambda o, its: ' / '.join(value for value in [o['recipient_name'] or '', o['recipient_phone1'] or ''] if value)),
        'customs_id': ('통관부호', lambda o, its: o['customs_id'] or ''),
        'address': ('배송지 주소', lambda o, its: o['address'] or ''),
        'product': ('상품명 + 옵션명', lambda o, its: ' | '.join(
            ' '.join(value for value in [str(item.get('item_product_name') or '').strip(), str(item.get('item_option_name') or '').strip()] if value and value not in str(item.get('item_product_name') or ''))
            for item in its
        ) if its else ' '.join(value for value in [o['original_product_name'] or '', o['original_option_name'] or ''] if value)),
        'product_quantity': ('상품수량', lambda o, its: ' | '.join(str(product_quantity(o, item)) for item in its) if its else str(o['original_quantity'] or 1)),
        'option_quantity': ('옵션수량', lambda o, its: ' | '.join(str(item.get('option_quantity') or 1) for item in its) if its else '1'),
        'buy_quantity': ('구매수량', lambda o, its: ' | '.join(str(item.get('quantity') or 1) for item in its) if its else str(o['original_quantity'] or 1)),
        'courier_tracking': ('택배사 / 송장번호', lambda o, its: ' / '.join(value for value in [o['courier'] or '', o['tracking_no'] or ''] if value)),
        'status': ('상태', lambda o, its: o['status'] or '')
    }
    requested_order = column_order or list(column_definitions)
    ordered_keys = [key for key in requested_order if key in column_definitions]
    ordered_keys.extend(key for key in column_definitions if key not in ordered_keys)
    headers = [column_definitions[key][0] for key in ordered_keys]
    ws.append(headers)

    for o in orders:
        row_items = items_by_order.get(o['id'], [])
        ws.append([column_definitions[key][1](o, row_items) for key in ordered_keys])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
