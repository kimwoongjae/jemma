import random
import time
import re
import requests
import json
from datetime import datetime
import openpyxl
from io import BytesIO
from database import get_connection

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
        # 옵션수량 * 구매수량
        opt_qty = int(it.get('option_quantity') or 1)
        buy_qty = int(it.get('quantity') or 1)
        qty = opt_qty * buy_qty

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
        resp = requests.post(
            url,
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=30
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

    for oid in order_ids:
        cur.execute('SELECT * FROM orders WHERE id = ?', (oid,))
        ord_row = cur.fetchone()
        if not ord_row:
            continue

        ord_dict = dict(ord_row)
        existing_tracking = ord_dict.get('tracking_no') or ''
        existing_pdf = ord_dict.get('invoice_pdf_url') or ''
        tracking_no = None
        pdf_url = ''

        # 나르다 실 API 연동 호출
        if target_courier == '나르다' and is_mock == 0:
            if existing_tracking and existing_pdf and existing_tracking.startswith('30'):
                tracking_no = existing_tracking
                pdf_url = existing_pdf
            else:
                cur.execute('SELECT * FROM order_items WHERE order_id = ?', (oid,))
                raw_items = [dict(r) for r in cur.fetchall()]

                # 재고 마스터의 영문명, HS코드, 단가 매칭 보완
                items = []
                for it in raw_items:
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

                api_res = call_narda_regist_order(ord_dict, items, settings)
                if api_res.get('success'):
                    tracking_no = api_res.get('tracking_no')
                    pdf_url = api_res.get('pdf_url', '')
                else:
                    err_msg = api_res.get('error') or '등록 실패'
                    print(f"나르다 API 발급 실패 (주문번호 {ord_dict['order_no']}): {err_msg}")
                    failed_orders.append({
                        'order_id': oid,
                        'order_no': ord_dict['order_no'],
                        'recipient_name': ord_dict['recipient_name'],
                        'error': err_msg
                    })
                    continue
        else:
            # 모의 발급 모드
            tracking_no = generate_mock_tracking_number(target_courier)
            pdf_url = ''

        now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
        cur.execute('''
            UPDATE orders
            SET tracking_no = ?, courier = ?, invoice_pdf_url = ?, status = '출고완료', dispatched_at = ?
            WHERE id = ?
        ''', (tracking_no, target_courier, pdf_url, now_str, oid))

        # ==========================================
        # 품목별 재고 차감 및 바코드별 집계 (중복 차감 방지)
        # ==========================================
        cur.execute('SELECT * FROM order_items WHERE order_id = ? AND stock_deducted = 0', (oid,))
        items_to_deduct = [dict(r) for r in cur.fetchall()]

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
                if sku_key not in aggregated_deductions:
                    aggregated_deductions[sku_key] = {
                        'product_id': matched_prod['id'],
                        'sku': matched_prod['sku'],
                        'product_name': matched_prod['product_name'],
                        'option_name': matched_prod.get('option_name') or '',
                        'total_qty': 0,
                        'final_stock': new_stock,
                        'order_nos': set(),
                        'courier': target_courier
                    }
                aggregated_deductions[sku_key]['total_qty'] += deduct_qty
                aggregated_deductions[sku_key]['final_stock'] = new_stock
                aggregated_deductions[sku_key]['order_nos'].add(ord_dict['order_no'])

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
    # 바코드 기준으로 총수량 1건만 입출고 내역(inventory_logs)에 기입
    # ==========================================
    now_log_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    for agg in aggregated_deductions.values():
        cnt = len(agg['order_nos'])
        memo_text = f"총 {cnt}건 주문 일괄 출고 ({agg['courier']} 송장발급)"
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
