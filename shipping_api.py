import random
import time
from datetime import datetime
import openpyxl
from io import BytesIO
from database import get_connection

COURIER_PREFIX = {
    '나르다': '7788',
    'CJ대한통운': '6892',
    '우체국택배': '6098',
    '한진택배': '5081',
    '롯데택배': '4023',
    '로젠택배': '9012'
}

def generate_mock_tracking_number(courier_name='CJ대한통운'):
    prefix = COURIER_PREFIX.get(courier_name, '6892')
    # 12자리 유효 형태의 송장번호 생성
    rand_mid = f"{random.randint(1000000, 9999999):07d}"
    base_no = prefix + rand_mid
    # 검증번호 (modulo 7 알고리즘 등)
    check_digit = sum(int(c) for c in base_no) % 7
    return f"{base_no}{check_digit}"

def issue_tracking_numbers(order_ids, courier_name=None):
    conn = get_connection()
    cur = conn.cursor()

    # 택배사 설정 조회
    cur.execute('SELECT * FROM courier_settings LIMIT 1')
    settings = cur.fetchone()
    if not courier_name:
        courier_name = settings['courier_name'] if settings else 'CJ대한통운'

    updated_orders = []

    for oid in order_ids:
        cur.execute('SELECT * FROM orders WHERE id = ?', (oid,))
        ord_row = cur.fetchone()
        if not ord_row:
            continue

        # 이미 송장번호가 있으면 유지하거나 새로 갱신
        tracking_no = ord_row['tracking_no']
        if not tracking_no:
            # 실제 API 연동 시 여기서 requests.post(courier_api_url, ...) 호출
            # settings['is_mock'] == 1 이면 가상 채번
            tracking_no = generate_mock_tracking_number(courier_name)

        cur.execute('''
            UPDATE orders
            SET tracking_no = ?, courier = ?, status = '송장발급'
            WHERE id = ?
        ''', (tracking_no, courier_name, oid))

        updated_orders.append({
            'order_id': oid,
            'order_no': ord_row['order_no'],
            'recipient_name': ord_row['recipient_name'],
            'courier': courier_name,
            'tracking_no': tracking_no
        })

    conn.commit()
    conn.close()
    return updated_orders

def complete_dispatch(order_ids):
    """
    발송 완료 처리:
    1. 주문 상태를 '출고완료'로 변경
    2. 연결된 order_items 품목들의 재고를 실제 차감 (stock_deducted = 1)
    3. inventory_logs 에 '주문출고' 로그 기록
    """
    conn = get_connection()
    cur = conn.cursor()
    dispatched_count = 0
    deducted_items = 0

    for oid in order_ids:
        cur.execute('SELECT * FROM orders WHERE id = ?', (oid,))
        ord_row = cur.fetchone()
        if not ord_row:
            continue

        # 주문 상태 업데이트 및 발송일시 기록
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
        cur.execute("UPDATE orders SET status = '출고완료', dispatched_at = ? WHERE id = ?", (now_str, oid))
        dispatched_count += 1

        # 품목별 재고 차감
        cur.execute('SELECT * FROM order_items WHERE order_id = ? AND stock_deducted = 0', (oid,))
        items = cur.fetchall()

        for it in items:
            cur.execute('SELECT * FROM products WHERE sku = ?', (it['sku'],))
            prod = cur.fetchone()
            if prod:
                new_stock = prod['stock_quantity'] - it['quantity']
                cur.execute('UPDATE products SET stock_quantity = ? WHERE id = ?', (new_stock, prod['id']))
                cur.execute('''
                    INSERT INTO inventory_logs (
                        product_id, sku, product_name, option_name,
                        change_type, change_qty, resulting_stock, reference_no, memo
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    prod['id'], prod['sku'], prod['product_name'], prod['option_name'],
                    '주문출고', -it['quantity'], new_stock, ord_row['order_no'], f"수취인: {ord_row['recipient_name']}"
                ))
            
            cur.execute("UPDATE order_items SET stock_deducted = 1, status = '출고완료' WHERE id = ?", (it['id'],))
            deducted_items += 1

    conn.commit()
    conn.close()
    return {'dispatched_orders': dispatched_count, 'deducted_items': deducted_items}

def export_shipping_excel(order_ids=None):
    """
    발송 완료된 주문 목록을 원본 엑셀 형식에 송장번호가 채워진 파일로 다운로드
    """
    conn = get_connection()
    cur = conn.cursor()

    if order_ids and len(order_ids) > 0:
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT * FROM orders WHERE id IN ({placeholders}) ORDER BY id ASC', order_ids)
    else:
        cur.execute('SELECT * FROM orders ORDER BY id DESC')

    orders = cur.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "발송처리결과"

    headers = [
        "주문번호", "발송일시", "택배사", "송장번호", "무게(kg)", "배송방법", "상품명", "옵션이름", "상품수량",
        "수취인이름", "수취인전화번호1", "수취인전화번호2", "개인통관부호", "수취인우편번호", "수취인주소", "배송메세지", "처리상태"
    ]
    ws.append(headers)

    for o in orders:
        ws.append([
            o['order_no'],
            o['dispatched_at'] or '',
            o['courier'] or 'CJ대한통운',
            o['tracking_no'] or '',
            o['weight'] if ('weight' in o.keys() and o['weight']) else 0.0,
            o['shipping_method'] or '택배',
            o['original_product_name'] or '',
            o['original_option_name'] or '',
            o['original_quantity'] or 1,
            o['recipient_name'] or '',
            o['recipient_phone1'] or '',
            o['recipient_phone2'] or '',
            o['customs_id'] or '',
            o['postal_code'] or '',
            o['address'] or '',
            o['shipping_memo'] or '',
            o['status'] or '주문접수'
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
