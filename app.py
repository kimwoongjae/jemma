import os
import json
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

# ----------------- 5. 재고 관리 (WMS) API -----------------
@app.route('/api/inventory', methods=['GET'])
def get_inventory():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products ORDER BY stock_quantity ASC, product_name ASC')
    products = [dict(p) for p in cur.fetchall()]
    conn.close()
    return jsonify(products)

@app.route('/api/inventory/stock-in', methods=['POST'])
def stock_in():
    data = request.get_json() or {}
    product_id = data.get('product_id')
    qty = int(data.get('quantity', 0))
    memo = data.get('memo', '일반 입고')

    if not product_id or qty <= 0:
        return jsonify({'error': '유효한 상품 및 입고 수량을 입력하세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products WHERE id = ?', (product_id,))
    prod = cur.fetchone()
    if not prod:
        conn.close()
        return jsonify({'error': '상품을 찾을 수 없습니다.'}), 404

    new_stock = prod['stock_quantity'] + qty
    cur.execute('UPDATE products SET stock_quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_stock, product_id))
    cur.execute('''
        INSERT INTO inventory_logs (product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (prod['id'], prod['sku'], prod['product_name'], prod['option_name'], '입고', qty, new_stock, memo))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'new_stock': new_stock})

@app.route('/api/inventory/adjust', methods=['POST'])
def stock_adjust():
    data = request.get_json() or {}
    product_id = data.get('product_id')
    new_stock = int(data.get('new_stock', 0))
    memo = data.get('memo', '실사 재고 조정')

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products WHERE id = ?', (product_id,))
    prod = cur.fetchone()
    if not prod:
        conn.close()
        return jsonify({'error': '상품을 찾을 수 없습니다.'}), 404

    diff = new_stock - prod['stock_quantity']
    cur.execute('UPDATE products SET stock_quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', (new_stock, product_id))
    cur.execute('''
        INSERT INTO inventory_logs (product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (prod['id'], prod['sku'], prod['product_name'], prod['option_name'], '재고조정', diff, new_stock, memo))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'new_stock': new_stock})

@app.route('/api/inventory/product', methods=['POST'])
def save_product():
    data = request.get_json() or {}
    sku = data.get('sku', '').strip()
    product_name = data.get('product_name', '').strip()
    option_name = data.get('option_name', '').strip()
    stock_quantity = int(data.get('stock_quantity', 0))
    safety_stock = int(data.get('safety_stock', 10))
    location = data.get('location', 'A-01').strip()
    prod_id = data.get('id')

    if not sku or not product_name:
        return jsonify({'error': 'SKU와 상품명은 필수입니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    if prod_id:
        cur.execute('''
            UPDATE products
            SET sku = ?, product_name = ?, option_name = ?, stock_quantity = ?, safety_stock = ?, location = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (sku, product_name, option_name, stock_quantity, safety_stock, location, prod_id))
    else:
        cur.execute('''
            INSERT INTO products (sku, product_name, option_name, stock_quantity, safety_stock, location)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (sku, product_name, option_name, stock_quantity, safety_stock, location))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/inventory/logs', methods=['GET'])
def get_inventory_logs():
    limit = int(request.args.get('limit', 100))
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM inventory_logs ORDER BY id DESC LIMIT ?', (limit,))
    logs = [dict(l) for l in cur.fetchall()]
    conn.close()
    return jsonify(logs)

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



