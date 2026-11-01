import os
import json
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file

from werkzeug.utils import secure_filename

from database import init_db, get_connection
from excel_processor import process_excel_file, save_orders_to_db
from shipping_api import (
    issue_tracking_numbers, complete_dispatch,
    export_shipping_excel
)

app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024  # 30MB
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# DB 초기화
init_db()

@app.route('/')
def index():
    return render_template('index.html')

# ----------------- 1. 통계 요약 대시보드 -----------------
@app.route('/api/dashboard/stats', methods=['GET'])
def get_dashboard_stats():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute('SELECT COUNT(*) as cnt FROM orders')
    total_orders = cur.fetchone()['cnt']

    cur.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = '주문접수'")
    pending_tracking = cur.fetchone()['cnt']

    cur.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = '송장발급'")
    issued_tracking = cur.fetchone()['cnt']

    cur.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = '출고완료'")
    completed_dispatch = cur.fetchone()['cnt']

    cur.execute('SELECT COUNT(*) as cnt FROM products WHERE stock_quantity <= safety_stock')
    low_stock_count = cur.fetchone()['cnt']

    cur.execute('SELECT COUNT(*) as cnt FROM order_items')
    total_items = cur.fetchone()['cnt']

    conn.close()
    return jsonify({
        'total_orders': total_orders,
        'pending_tracking': pending_tracking,
        'issued_tracking': issued_tracking,
        'completed_dispatch': completed_dispatch,
        'low_stock_count': low_stock_count,
        'total_items': total_items
    })

# ----------------- 2. 엑셀 업로드 및 주문 분할 -----------------
@app.route('/api/upload/preview', methods=['POST'])
def upload_preview():
    # 단일 또는 복수 파일 지원 ('files' 또는 'file')
    uploaded_files = request.files.getlist('files')
    if not uploaded_files:
        uploaded_files = request.files.getlist('file')
    if not uploaded_files and 'file' in request.files:
        uploaded_files = [request.files['file']]

    # 빈 파일 제거
    valid_files = [f for f in uploaded_files if f.filename != '']
    if not valid_files:
        return jsonify({'error': '업로드할 파일이 없습니다.'}), 400

    merged_orders = []
    total_excel_rows = 0
    total_split_items = 0
    batch_names = []
    detected_cols = {}

    try:
        for idx, file in enumerate(valid_files):
            filename = secure_filename(file.filename)
            if not filename:
                filename = f"order_{int(os.times()[4])}_{idx}.xlsx"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            parsed = process_excel_file(filepath, batch_name=file.filename)
            total_excel_rows += parsed.get('total_excel_rows', 0)
            total_split_items += parsed.get('total_split_items', 0)
            batch_names.append(file.filename)
            detected_cols.update(parsed.get('detected_columns', {}))
            merged_orders.extend(parsed.get('orders', []))

        preview_data = {
            'batch_name': ', '.join(batch_names) if len(batch_names) <= 2 else f"{batch_names[0]} 외 {len(batch_names)-1}개 파일",
            'total_files': len(valid_files),
            'total_excel_rows': total_excel_rows,
            'total_split_items': total_split_items,
            'detected_columns': detected_cols,
            'orders': merged_orders
        }
        return jsonify(preview_data)
    except Exception as e:
        return jsonify({'error': f'엑셀 처리 중 오류: {str(e)}'}), 500


@app.route('/api/upload/confirm', methods=['POST'])
def upload_confirm():
    data = request.get_json()
    if not data or 'orders' not in data:
        return jsonify({'error': '저장할 주문 데이터가 없습니다.'}), 400

    batch_name = data.get('batch_name', '배치주문')
    try:
        result = save_orders_to_db(data, batch_name)
        return jsonify({
            'status': 'success',
            'message': f"{result['saved_orders']}건의 주문({result['saved_items']}개 품목)이 성공적으로 등록되었습니다.",
            'result': result
        })
    except Exception as e:
        return jsonify({'error': f'저장 오류: {str(e)}'}), 500

# ----------------- 3. 주문 관리 API -----------------
@app.route('/api/orders', methods=['GET'])
def get_orders():
    status = request.args.get('status')
    search = request.args.get('search', '').strip()
    limit = int(request.args.get('limit', 200))

    conn = get_connection()
    cur = conn.cursor()

    query = 'SELECT * FROM orders WHERE 1=1'
    params = []

    if status and status != '전체':
        query += ' AND status = ?'
        params.append(status)

    if search:
        query += ''' AND (
            order_no LIKE ? OR recipient_name LIKE ? OR recipient_phone1 LIKE ? 
            OR tracking_no LIKE ? OR customs_id LIKE ? OR address LIKE ?
        )'''
        wild = f'%{search}%'
        params.extend([wild, wild, wild, wild, wild, wild])

    query += ' ORDER BY id DESC LIMIT ?'
    params.append(limit)

    cur.execute(query, params)
    orders = [dict(row) for row in cur.fetchall()]

    # 각 주문의 분할 품목(order_items) 조회
    for o in orders:
        cur.execute('SELECT * FROM order_items WHERE order_id = ?', (o['id'],))
        o['items'] = [dict(it) for it in cur.fetchall()]

    conn.close()
    return jsonify(orders)

@app.route('/api/orders/delete', methods=['POST'])
def delete_orders():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    if not order_ids:
        return jsonify({'error': '삭제할 주문을 선택하세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(order_ids))
    cur.execute(f'DELETE FROM order_items WHERE order_id IN ({placeholders})', order_ids)
    cur.execute(f'DELETE FROM orders WHERE id IN ({placeholders})', order_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'deleted_count': len(order_ids)})

# ----------------- 4. 배송 & 송장 & 출고 API -----------------
@app.route('/api/shipping/issue-tracking', methods=['POST'])
def api_issue_tracking():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    courier_name = data.get('courier_name')
    if not order_ids:
        return jsonify({'error': '송장을 발급할 주문을 선택하세요.'}), 400

    res = issue_tracking_numbers(order_ids, courier_name)
    return jsonify({'status': 'success', 'updated': res})

@app.route('/api/shipping/dispatch', methods=['POST'])
def api_dispatch():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    if not order_ids:
        return jsonify({'error': '출고 완료할 주문을 선택하세요.'}), 400

    res = complete_dispatch(order_ids)
    return jsonify({
        'status': 'success',
        'message': f"{res['dispatched_orders']}건 발송 완료 처리 및 {res['deducted_items']}개 품목 재고가 자동 차감되었습니다.",
        'details': res
    })

@app.route('/api/shipping/export-excel', methods=['GET'])
def api_export_excel():
    order_ids_str = request.args.get('order_ids', '')
    order_ids = [int(x) for x in order_ids_str.split(',') if x.strip().isdigit()]
    excel_buf = export_shipping_excel(order_ids if order_ids else None)
    return send_file(
        excel_buf,
        as_attachment=True,
        download_name='발송처리결과_송장완료.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/api/shipping/labels', methods=['GET'])
def api_shipping_labels():
    order_ids_str = request.args.get('order_ids', '')
    if not order_ids_str:
        return jsonify([])

    order_ids = [int(x) for x in order_ids_str.split(',') if x.strip().isdigit()]
    conn = get_connection()
    cur = conn.cursor()

    # 발송자 정보
    cur.execute('SELECT * FROM courier_settings LIMIT 1')
    sender = cur.fetchone()
    sender_info = dict(sender) if sender else {
        'sender_name': '프리비 물류센터',
        'sender_phone': '02-1234-5678',
        'sender_address': '서울특별시 강남구 테헤란로 123 물류빌딩 3층'
    }

    placeholders = ','.join(['?'] * len(order_ids))
    cur.execute(f'SELECT * FROM orders WHERE id IN ({placeholders}) ORDER BY id ASC', order_ids)
    orders = [dict(row) for row in cur.fetchall()]

    for o in orders:
        cur.execute('SELECT * FROM order_items WHERE order_id = ?', (o['id'],))
        o['items'] = [dict(it) for it in cur.fetchall()]
        o['sender'] = sender_info

    conn.close()
    return jsonify(orders)

# ----------------- 5. 재고 관리 (WMS) 및 입고 API -----------------
@app.route('/api/inventory', methods=['GET'])
def get_inventory():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products ORDER BY id DESC')
    products = [dict(p) for p in cur.fetchall()]
    conn.close()
    return jsonify(products)

@app.route('/api/inventory/search', methods=['GET'])
def search_product_by_barcode():
    barcode = request.args.get('barcode', '').strip()
    if not barcode:
        return jsonify({'found': False, 'message': '바코드를 입력하세요.'})

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products WHERE sku = ? LIMIT 1', (barcode,))
    prod = cur.fetchone()
    conn.close()

    if prod:
        return jsonify({'found': True, 'product': dict(prod)})
    else:
        return jsonify({'found': False, 'message': '신규 상품입니다 (미등록 바코드)'})

@app.route('/api/inventory/barcode-stock-in', methods=['POST'])
def barcode_stock_in():
    data = request.get_json() or {}
    sku = data.get('sku', '').strip()
    product_name = data.get('product_name', '').strip()
    english_name = data.get('english_name', '').strip()
    hs_code = data.get('hs_code', '').strip()
    unit_price = float(data.get('unit_price', 0.0) or 0.0)
    qty = int(data.get('quantity', 0) or 0)
    memo = data.get('memo', '바코드 개별 입고')

    if not sku:
        return jsonify({'error': '바코드(SKU)는 필수입니다.'}), 400
    if qty <= 0:
        return jsonify({'error': '입고 수량은 1개 이상이어야 합니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products WHERE sku = ?', (sku,))
    prod = cur.fetchone()

    if prod:
        new_stock = prod['stock_quantity'] + qty
        cur.execute('''
            UPDATE products
            SET stock_quantity = ?, 
                product_name = CASE WHEN ? != '' THEN ? ELSE product_name END,
                english_name = CASE WHEN ? != '' THEN ? ELSE english_name END,
                hs_code = CASE WHEN ? != '' THEN ? ELSE hs_code END,
                unit_price = CASE WHEN ? > 0 THEN ? ELSE unit_price END,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (new_stock, product_name, product_name, english_name, english_name, hs_code, hs_code, unit_price, unit_price, prod['id']))
        prod_id = prod['id']
        p_name = product_name or prod['product_name']
    else:
        if not product_name:
            product_name = f"상품_{sku}"
        new_stock = qty
        cur.execute('''
            INSERT INTO products (sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (sku, product_name, '', english_name, hs_code, unit_price, new_stock))
        prod_id = cur.lastrowid
        p_name = product_name


    # 입고 이력 로그
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    cur.execute('''
        INSERT INTO inventory_logs (product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (prod_id, sku, p_name, '', '입고', qty, new_stock, memo, now_str))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'new_stock': new_stock, 'product_name': p_name})

@app.route('/api/inventory/batch-stock-in', methods=['POST'])
def batch_stock_in():
    data = request.get_json() or {}
    items = data.get('items', [])
    if not items:
        return jsonify({'error': '입고할 품목 목록이 비어있습니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    success_count = 0
    total_qty = 0
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

    for it in items:
        sku = str(it.get('sku', '')).strip()
        qty = int(it.get('quantity', 0) or 0)
        p_name = str(it.get('product_name', '')).strip()
        unit_price = float(it.get('unit_price', 0.0) or 0.0)

        if not sku or qty <= 0:
            continue

        cur.execute('SELECT * FROM products WHERE sku = ?', (sku,))
        prod = cur.fetchone()

        if prod:
            new_stock = prod['stock_quantity'] + qty
            cur.execute('''
                UPDATE products
                SET stock_quantity = ?,
                    product_name = CASE WHEN ? != '' THEN ? ELSE product_name END,
                    unit_price = CASE WHEN ? > 0 THEN ? ELSE unit_price END,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (new_stock, p_name, p_name, unit_price, unit_price, prod['id']))
            prod_id = prod['id']
            actual_name = p_name or prod['product_name']
        else:
            actual_name = p_name or f"상품_{sku}"
            new_stock = qty
            cur.execute('''
                INSERT INTO products (sku, product_name, option_name, unit_price, stock_quantity)
                VALUES (?, ?, ?, ?, ?)
            ''', (sku, actual_name, '', unit_price, new_stock))
            prod_id = cur.lastrowid


        cur.execute('''
            INSERT INTO inventory_logs (product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (prod_id, sku, actual_name, '', '입고', qty, new_stock, '일괄 바코드 스캔 입고', now_str))

        success_count += 1
        total_qty += qty

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'processed_items': success_count, 'total_quantity': total_qty})

@app.route('/api/inventory/upload-excel', methods=['POST'])
def upload_products_excel():
    if 'file' not in request.files:
        return jsonify({'error': '업로드할 엑셀 파일이 없습니다.'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '선택된 파일명이 없습니다.'}), 400

    try:
        import pandas as pd
        df = pd.read_excel(file)
        
        # 컬럼 매핑 탐색
        col_map = {}
        for c in df.columns:
            cs = str(c).strip().lower()
            if '바코드' in cs or 'sku' in cs: col_map['sku'] = c
            elif '상품명' in cs or '품목명' in cs: col_map['product_name'] = c
            elif '영문' in cs: col_map['english_name'] = c
            elif 'hs' in cs: col_map['hs_code'] = c
            elif '단가' in cs or 'price' in cs: col_map['unit_price'] = c
            elif '재고' in cs or '수량' in cs: col_map['stock_quantity'] = c

        if 'sku' not in col_map and 'product_name' not in col_map:
            return jsonify({'error': '엑셀에 [바코드(SKU)] 또는 [상품명] 컬럼이 필요합니다.'}), 400

        conn = get_connection()
        cur = conn.cursor()
        saved_count = 0
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

        for _, row in df.iterrows():
            sku = str(row.get(col_map.get('sku', ''), '')).strip()
            name = str(row.get(col_map.get('product_name', ''), '')).strip()
            eng = str(row.get(col_map.get('english_name', ''), '')).strip()
            hs = str(row.get(col_map.get('hs_code', ''), '')).strip()
            
            try:
                price = float(row.get(col_map.get('unit_price', ''), 0.0) or 0.0)
            except:
                price = 0.0
            try:
                stock = int(float(row.get(col_map.get('stock_quantity', ''), 0) or 0))
            except:
                stock = 0

            if not sku and not name:
                continue
            if not sku:
                sku = f"SKU-{abs(hash(name)) % 1000000:06d}"

            cur.execute('SELECT * FROM products WHERE sku = ?', (sku,))
            existing = cur.fetchone()

            if existing:
                cur.execute('''
                    UPDATE products
                    SET product_name = CASE WHEN ? != '' THEN ? ELSE product_name END,
                        english_name = CASE WHEN ? != '' THEN ? ELSE english_name END,
                        hs_code = CASE WHEN ? != '' THEN ? ELSE hs_code END,
                        unit_price = ?,
                        stock_quantity = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, name, eng, eng, hs, hs, price, stock, existing['id']))
            else:
                cur.execute('''
                    INSERT INTO products (sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (sku, name or sku, '', eng, hs, price, stock))
                prod_id = cur.lastrowid

                cur.execute('''
                    INSERT INTO inventory_logs (product_id, sku, product_name, change_type, change_qty, resulting_stock, memo, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (prod_id, sku, name or sku, '엑셀등록', stock, stock, '엑셀 상품 일괄 등록', now_str))

            saved_count += 1

        conn.commit()
        conn.close()
        return jsonify({'status': 'success', 'message': f"{saved_count}개 상품이 엑셀에서 성공적으로 등록/반영되었습니다."})
    except Exception as e:
        return jsonify({'error': f'엑셀 처리 중 오류: {str(e)}'}), 500

@app.route('/api/inventory/product', methods=['POST'])
def save_product():
    data = request.get_json() or {}
    sku = data.get('sku', '').strip()
    product_name = data.get('product_name', '').strip()
    english_name = data.get('english_name', '').strip()
    hs_code = data.get('hs_code', '').strip()
    unit_price = float(data.get('unit_price', 0.0) or 0.0)
    stock_quantity = int(data.get('stock_quantity', 0) or 0)
    prod_id = data.get('id')

    if not sku or not product_name:
        return jsonify({'error': 'SKU(바코드)와 상품명은 필수입니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    if prod_id:
        cur.execute('''
            UPDATE products
            SET sku = ?, product_name = ?, english_name = ?, hs_code = ?, unit_price = ?, stock_quantity = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (sku, product_name, english_name, hs_code, unit_price, stock_quantity, prod_id))
    else:
        cur.execute('''
            INSERT INTO products (sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (sku, product_name, '', english_name, hs_code, unit_price, stock_quantity))


    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/inventory/product/<int:prod_id>', methods=['DELETE'])
def delete_product(prod_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM products WHERE id = ?', (prod_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/inventory/logs', methods=['GET'])
def get_inventory_logs():
    limit = int(request.args.get('limit', 200))
    search = request.args.get('search', '').strip()
    change_type = request.args.get('type', '').strip()

    conn = get_connection()
    cur = conn.cursor()
    query = 'SELECT * FROM inventory_logs WHERE 1=1'
    params = []

    if change_type and change_type != '전체':
        query += ' AND change_type = ?'
        params.append(change_type)

    if search:
        query += ' AND (sku LIKE ? OR product_name LIKE ? OR memo LIKE ?)'
        wild = f'%{search}%'
        params.extend([wild, wild, wild])

    query += ' ORDER BY id DESC LIMIT ?'
    params.append(limit)

    cur.execute(query, params)
    logs = [dict(l) for l in cur.fetchall()]
    conn.close()
    return jsonify(logs)

@app.route('/api/inventory/export-logs', methods=['GET'])
def export_inventory_logs():
    ids_str = request.args.get('ids', '').strip()
    conn = get_connection()
    cur = conn.cursor()

    if ids_str:
        ids = [int(x) for x in ids_str.split(',') if x.strip().isdigit()]
        placeholders = ','.join(['?'] * len(ids))
        cur.execute(f'SELECT * FROM inventory_logs WHERE id IN ({placeholders}) ORDER BY id DESC', ids)
    else:
        cur.execute('SELECT * FROM inventory_logs ORDER BY id DESC LIMIT 1000')

    logs = cur.fetchall()
    conn.close()

    import openpyxl
    from io import BytesIO
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "입출고내역"

    headers = ["일시 (날짜)", "바코드 (SKU)", "상품명", "구분 (입고/출고)", "변동수량", "잔여재고", "참조/사유"]
    ws.append(headers)

    for l in logs:
        # created_at 형식 정규화 (YYYY-MM-DD HH:MM)
        date_str = str(l['created_at'])[:16]
        ws.append([
            date_str,
            l['sku'],
            l['product_name'],
            l['change_type'],
            l['change_qty'],
            l['resulting_stock'],
            l['memo'] or ''
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name='입출고내역리스트.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ----------------- 6. 세트상품/옵션 분할 규칙 API -----------------
@app.route('/api/bundles', methods=['GET'])
def get_bundles():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM bundle_rules ORDER BY id DESC')
    rules = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify(rules)

@app.route('/api/bundles', methods=['POST'])
def add_bundle():
    data = request.get_json() or {}
    rule_name = data.get('rule_name', '')
    source_p = data.get('source_product_name', '').strip()
    source_o = data.get('source_option_name', '').strip()
    target_p = data.get('target_product_name', '').strip()
    target_o = data.get('target_option_name', '').strip()
    target_sku = data.get('target_sku', '').strip()
    multiplier = int(data.get('multiplier', 1))

    if not source_o or not target_p or not target_o:
        return jsonify({'error': '원본 옵션명과 타겟 상품명/옵션명은 필수입니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        INSERT INTO bundle_rules (
            rule_name, source_product_name, source_option_name,
            target_product_name, target_option_name, target_sku, multiplier
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (rule_name or f"{source_o} -> {target_o}", source_p or '*', source_o, target_p, target_o, target_sku, multiplier))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/bundles/<int:rule_id>', methods=['DELETE'])
def delete_bundle(rule_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM bundle_rules WHERE id = ?', (rule_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# ----------------- 7. 택배사 & 발송자 설정 API -----------------
@app.route('/api/settings', methods=['GET'])
def get_settings():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM courier_settings LIMIT 1')
    row = cur.fetchone()
    conn.close()
    return jsonify(dict(row) if row else {})

@app.route('/api/settings', methods=['POST'])
def save_settings():
    data = request.get_json() or {}
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        UPDATE courier_settings
        SET courier_name = ?, customer_code = ?, api_key = ?,
            sender_name = ?, sender_phone = ?, sender_address = ?,
            is_mock = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = (SELECT id FROM courier_settings LIMIT 1)
    ''', (
        data.get('courier_name', 'CJ대한통운'),
        data.get('customer_code', ''),
        data.get('api_key', ''),
        data.get('sender_name', ''),
        data.get('sender_phone', ''),
        data.get('sender_address', ''),
        1 if data.get('is_mock') else 0
    ))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

import sys
import io
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 5050))
    print("=" * 60)
    print(f"[SERVER] Smart OMS / WMS System Running on port {PORT}...")
    print(f"[URL] Open your browser at: http://127.0.0.1:{PORT}")
    print("=" * 60)
    app.run(host='0.0.0.0', port=PORT, debug=False)



