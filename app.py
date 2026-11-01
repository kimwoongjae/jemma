import os
import json
import re
import uuid
import time
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

    cur.execute("SELECT COUNT(*) as cnt FROM orders WHERE status IN ('주문접수', '출고대기')")
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
    all_unregistered = {}

    try:
        for idx, file in enumerate(valid_files):
            orig_name = file.filename or f"order_{idx+1}.xlsx"
            ext = os.path.splitext(orig_name)[1].lower()
            if ext not in ['.xlsx', '.xls']:
                ext = '.xlsx'
            safe_fname = f"upload_{int(time.time())}_{uuid.uuid4().hex[:8]}{ext}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_fname)
            file.save(filepath)

            try:
                parsed = process_excel_file(filepath, batch_name=orig_name)
                total_excel_rows += parsed.get('total_excel_rows', 0)
                total_split_items += parsed.get('total_split_items', 0)
                batch_names.append(orig_name)
                detected_cols.update(parsed.get('detected_columns', {}))
                merged_orders.extend(parsed.get('orders', []))

                for unreg in parsed.get('unregistered_items', []):
                    comb = unreg['combined_name']
                    if comb not in all_unregistered:
                        all_unregistered[comb] = dict(unreg)
                    else:
                        all_unregistered[comb]['count'] += unreg.get('count', 1)
            finally:
                if os.path.exists(filepath):
                    try:
                        os.remove(filepath)
                    except Exception:
                        pass

        preview_data = {
            'batch_name': ', '.join(batch_names) if len(batch_names) <= 2 else f"{batch_names[0]} 외 {len(batch_names)-1}개 파일",
            'total_files': len(valid_files),
            'total_excel_rows': total_excel_rows,
            'total_split_items': total_split_items,
            'detected_columns': detected_cols,
            'orders': merged_orders,
            'unregistered_items': list(all_unregistered.values())
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
    status = request.args.get('status', '').strip()
    courier = request.args.get('courier', '').strip()
    customs_status = request.args.get('customs_status', '').strip()
    search = request.args.get('search', '').strip()
    search_field = request.args.get('search_field', 'all').strip()
    sort_by = request.args.get('sort_by', 'single_barcode_asc').strip()
    limit_param = request.args.get('limit', '').strip()

    conn = get_connection()
    cur = conn.cursor()

    query = 'SELECT * FROM orders WHERE 1=1'
    params = []

    if status == '출고대기':
        query += " AND status IN ('주문접수', '출고대기', '송장발급')"
    elif status == '주문접수':
        query += " AND status IN ('주문접수', '출고대기')"
    elif status and status != '전체':
        query += ' AND status = ?'
        params.append(status)

    if customs_status and customs_status != '전체':
        if customs_status == '오류':
            query += " AND (customs_validation_status = '오류' OR status = '통관오류')"
        elif customs_status == '일치':
            query += " AND customs_validation_status = '일치'"
        elif customs_status == '미검증':
            query += " AND (customs_validation_status = '' OR customs_validation_status IS NULL)"

    if courier and courier != '전체':
        query += ' AND courier = ?'
        params.append(courier)

    if search:
        wild = f'%{search}%'
        if search_field == 'product_name':
            query += ''' AND (
                original_product_name LIKE ? OR original_option_name LIKE ? 
                OR id IN (SELECT order_id FROM order_items WHERE item_product_name LIKE ? OR item_option_name LIKE ?)
            )'''
            params.extend([wild, wild, wild, wild])
        elif search_field == 'sku':
            query += ' AND (id IN (SELECT order_id FROM order_items WHERE sku LIKE ?))'
            params.append(wild)
        elif search_field == 'recipient_name':
            query += ' AND recipient_name LIKE ?'
            params.append(wild)
        elif search_field == 'order_no':
            query += ' AND order_no LIKE ?'
            params.append(wild)
        elif search_field == 'phone':
            query += ' AND (recipient_phone1 LIKE ? OR recipient_phone2 LIKE ?)'
            params.extend([wild, wild])
        elif search_field == 'customs_id':
            query += ' AND customs_id LIKE ?'
            params.append(wild)
        elif search_field == 'address':
            query += ' AND (address LIKE ? OR postal_code LIKE ?)'
            params.extend([wild, wild])
        else: # all
            query += ''' AND (
                order_no LIKE ? OR recipient_name LIKE ? OR recipient_phone1 LIKE ? 
                OR tracking_no LIKE ? OR customs_id LIKE ? OR address LIKE ? OR postal_code LIKE ?
                OR original_product_name LIKE ? OR original_option_name LIKE ?
                OR id IN (SELECT order_id FROM order_items WHERE item_product_name LIKE ? OR item_option_name LIKE ? OR sku LIKE ?)
            )'''
            params.extend([wild, wild, wild, wild, wild, wild, wild, wild, wild, wild, wild, wild])

    query += ' ORDER BY id DESC'
    if limit_param and limit_param.isdigit():
        query += ' LIMIT ?'
        params.append(int(limit_param))

    cur.execute(query, params)
    orders = [dict(row) for row in cur.fetchall()]

    # 모든 주문의 order_items를 1번의 쿼리로 일괄 조회 (N+1 문제 해결 -> 초고속 로딩)
    if orders:
        order_ids = [o['id'] for o in orders]
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT * FROM order_items WHERE order_id IN ({placeholders}) ORDER BY id ASC', order_ids)
        items_by_order = {}
        for it in cur.fetchall():
            oid = it['order_id']
            if oid not in items_by_order:
                items_by_order[oid] = []
            items_by_order[oid].append(dict(it))
        for o in orders:
            o['items'] = items_by_order.get(o['id'], [])
    else:
        for o in orders:
            o['items'] = []

    conn.close()

    # 정렬: 단일 상품 주문(items 개수 1개)의 바코드(SKU) 숫자 오름차순 우선 정렬
    if sort_by == 'single_barcode_asc':
        def extract_sku_sort_key(sku_str):
            if not sku_str:
                return (2, 999999999999, '')
            digits = re.findall(r'\d+', str(sku_str))
            if digits:
                return (0, int(digits[0]), str(sku_str))
            return (1, 0, str(sku_str))

        def order_sort_key(o):
            items = o.get('items', [])
            item_cnt = len(items)
            is_single = 0 if item_cnt == 1 else 1
            primary_sku = items[0]['sku'] if items else ''
            sku_type, sku_num, sku_raw = extract_sku_sort_key(primary_sku)
            # 단일 품목 우선 -> 바코드 숫자 오름차순 -> 바코드 문자열 오름차순 -> 주문 ID 역순
            return (is_single, sku_type, sku_num, sku_raw, -o['id'])

        orders.sort(key=order_sort_key)

    return jsonify(orders)

@app.route('/api/dispatched-orders', methods=['GET'])
def get_dispatched_orders():
    courier = request.args.get('courier', '').strip()
    search = request.args.get('search', '').strip()
    limit_param = request.args.get('limit', '').strip()

    conn = get_connection()
    cur = conn.cursor()

    query = "SELECT * FROM orders WHERE status = '출고완료'"
    params = []

    if courier and courier != '전체':
        query += ' AND courier = ?'
        params.append(courier)

    if search:
        query += ''' AND (
            order_no LIKE ? OR recipient_name LIKE ? OR recipient_phone1 LIKE ? 
            OR tracking_no LIKE ? OR customs_id LIKE ? OR address LIKE ?
            OR original_product_name LIKE ? OR original_option_name LIKE ?
            OR id IN (SELECT order_id FROM order_items WHERE item_product_name LIKE ? OR item_option_name LIKE ? OR sku LIKE ?)
        )'''
        wild = f'%{search}%'
        params.extend([wild, wild, wild, wild, wild, wild, wild, wild, wild, wild, wild])

    query += ' ORDER BY COALESCE(dispatched_at, created_at) DESC, id DESC'
    if limit_param and limit_param.isdigit():
        query += ' LIMIT ?'
        params.append(int(limit_param))

    cur.execute(query, params)
    orders = [dict(row) for row in cur.fetchall()]

    if orders:
        order_ids = [o['id'] for o in orders]
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT * FROM order_items WHERE order_id IN ({placeholders}) ORDER BY id ASC', order_ids)
        items_by_order = {}
        for it in cur.fetchall():
            oid = it['order_id']
            if oid not in items_by_order:
                items_by_order[oid] = []
            items_by_order[oid].append(dict(it))
        for o in orders:
            o['items'] = items_by_order.get(o['id'], [])
    else:
        for o in orders:
            o['items'] = []

    conn.close()
    return jsonify(orders)

@app.route('/api/orders/<int:order_id>/weight', methods=['POST'])
def update_order_weight(order_id):
    data = request.get_json() or {}
    try:
        weight = float(data.get('weight', 0.0) or 0.0)
    except:
        weight = 0.0

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('UPDATE orders SET weight = ? WHERE id = ?', (weight, order_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'weight': weight})

@app.route('/api/orders/<int:order_id>/edit', methods=['POST'])
def edit_order(order_id):
    data = request.get_json() or {}
    recipient_name = data.get('recipient_name', '').strip()
    recipient_phone1 = data.get('recipient_phone1', '').strip()
    recipient_phone2 = data.get('recipient_phone2', '').strip()
    customs_id = data.get('customs_id', '').strip().upper()
    postal_code = data.get('postal_code', '').strip()
    address = data.get('address', '').strip()
    shipping_memo = data.get('shipping_memo', '').strip()
    try:
        weight = float(data.get('weight', 0.0) or 0.0)
    except:
        weight = 0.0

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        UPDATE orders 
        SET recipient_name = ?, recipient_phone1 = ?, recipient_phone2 = ?,
            customs_id = ?, postal_code = ?, address = ?, shipping_memo = ?, weight = ?
        WHERE id = ?
    ''', (recipient_name, recipient_phone1, recipient_phone2, customs_id, postal_code, address, shipping_memo, weight, order_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': '주문 정보가 성공적으로 수정되었습니다.'})

@app.route('/api/orders/check-dispatch-availability', methods=['POST'])
def check_dispatch_availability():
    """
    재고관리 테이블(products)의 현재 재고를 기준으로 목록 상단부터 순차 가상 차감하여
    각 주문별 출고 가능 여부('출고가능', '부분부족', '재고부족')를 판별합니다.
    """
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])

    conn = get_connection()
    cur = conn.cursor()

    cur.execute('SELECT sku, stock_quantity FROM products')
    stock_map = {r['sku']: (r['stock_quantity'] or 0) for r in cur.fetchall() if r['sku']}

    if order_ids:
        placeholders = ','.join('?' for _ in order_ids)
        cur.execute(f"SELECT * FROM orders WHERE id IN ({placeholders})", order_ids)
        raw_orders = {r['id']: dict(r) for r in cur.fetchall()}
        orders = [raw_orders[oid] for oid in order_ids if oid in raw_orders]
    else:
        cur.execute("SELECT * FROM orders WHERE status IN ('주문접수', '출고대기', '송장발급') ORDER BY id DESC")
        orders = [dict(r) for r in cur.fetchall()]

    items_by_order = {}
    if orders:
        o_ids = [o['id'] for o in orders]
        placeholders = ','.join(['?'] * len(o_ids))
        cur.execute(f'SELECT * FROM order_items WHERE order_id IN ({placeholders}) ORDER BY id ASC', o_ids)
        for it in cur.fetchall():
            oid = it['order_id']
            if oid not in items_by_order:
                items_by_order[oid] = []
            items_by_order[oid].append(dict(it))

    availability_results = {}
    for o in orders:
        items = items_by_order.get(o['id'], [])

        all_available = True
        any_available = False
        item_details = []

        for it in items:
            sku = it['sku']
            opt_qty = it.get('option_quantity') or 1
            buy_qty = it.get('quantity') or 1
            needed = opt_qty * buy_qty

            current_stock = stock_map.get(sku, 0)
            if current_stock >= needed:
                stock_map[sku] -= needed
                any_available = True
                item_details.append({
                    'sku': sku,
                    'needed': needed,
                    'available': True,
                    'remaining': stock_map[sku]
                })
            else:
                all_available = False
                item_details.append({
                    'sku': sku,
                    'needed': needed,
                    'available': False,
                    'remaining': current_stock
                })

        if not items:
            disp_status = '출고불가'
        elif all_available:
            disp_status = '출고가능'
        elif any_available:
            disp_status = '부분부족'
        else:
            disp_status = '재고부족'

        availability_results[o['id']] = {
            'status': disp_status,
            'items': item_details
        }

    conn.close()
    return jsonify({'status': 'success', 'results': availability_results})

@app.route('/api/orders/validate-customs', methods=['POST'])
def validate_customs_batch():
    """
    gsiexpress.com/pcc_chk.php 관세청 연동 서비스를 통해 
    개인통관고유부호 유효성(성명, 부호, 연락처, 우편번호 일치 여부)을 일괄 검증합니다.
    """
    import requests
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])

    conn = get_connection()
    cur = conn.cursor()

    try:
        if order_ids:
            placeholders = ','.join('?' for _ in order_ids)
            cur.execute(f'''
                SELECT id, order_no, recipient_name, customs_id, recipient_phone1, postal_code
                FROM orders WHERE id IN ({placeholders})
            ''', order_ids)
        else:
            cur.execute('''
                SELECT id, order_no, recipient_name, customs_id, recipient_phone1, postal_code
                FROM orders WHERE status IN ('주문접수', '출고대기', '송장발급')
            ''')

        target_orders = [dict(r) for r in cur.fetchall()]
        if not target_orders:
            conn.close()
            return jsonify({'status': 'success', 'message': '검증할 주문이 없습니다.', 'total_checked': 0, 'results': []})

        results = []
        batch_size = 50

        for i in range(0, len(target_orders), batch_size):
            batch = target_orders[i:i + batch_size]
            lines = []
            order_map_full = {}  # (name, cid) -> list of oids
            order_map_cid = {}   # cid -> list of oids

            for o in batch:
                name = (o['recipient_name'] or '').strip()
                cid = (o['customs_id'] or '').strip().upper()
                phone = (o['recipient_phone1'] or '').strip()
                pcode = (o['postal_code'] or '').strip()

                line = f"{name}/{cid}/{phone}/{pcode}"
                lines.append(line)

                key_full = (name, cid)
                if key_full not in order_map_full:
                    order_map_full[key_full] = []
                order_map_full[key_full].append(o['id'])

                if cid not in order_map_cid:
                    order_map_cid[cid] = []
                order_map_cid[cid].append(o['id'])

            payload = {
                'action_type': 'query',
                'chk_data': '\n'.join(lines)
            }

            try:
                resp = requests.post(
                    'https://www.gsiexpress.com/pcc_chk.php',
                    data=payload,
                    timeout=25,
                    verify=False,
                    headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                )

                # 정규식 기반 HTML 파싱 (bs4 설치 여부와 무관하게 100% 안전하게 동작)
                html_text = resp.content.decode('utf-8', 'ignore')
                row_matches = re.findall(r'<tr>([\s\S]*?)</tr>', html_text)
                
                found_ids = set()
                for row_html in row_matches:
                    td_matches = re.findall(r'<td[^>]*>([\s\S]*?)</td>', row_html)
                    td_clean = [re.sub(r'<[^>]+>', '', t).strip() for t in td_matches]
                    # 테이블 컬럼 형식: [이름, 통관고유부호, 핸드폰번호, 우편번호, 검증 결과, 검증 결과 메시지]
                    if len(td_clean) >= 5:
                        res_name = td_clean[0]
                        res_cid = td_clean[1].upper()
                        res_status = td_clean[4]  # '정상', '오류', '장애'
                        res_msg = td_clean[5] if len(td_clean) > 5 else ''

                        matched_ids = order_map_full.get((res_name, res_cid))
                        if not matched_ids:
                            matched_ids = order_map_cid.get(res_cid, [])

                        for oid in matched_ids:
                            cur.execute('''
                                UPDATE orders 
                                SET customs_validation_status = ?, customs_validation_msg = ?
                                WHERE id = ?
                            ''', (res_status, res_msg, oid))
                            results.append({
                                'order_id': oid,
                                'recipient_name': res_name,
                                'customs_id': res_cid,
                                'status': res_status,
                                'message': res_msg
                            })
                            found_ids.add(oid)

                # 검증 결과 매칭되지 않은 주문 처리
                for o in batch:
                    if o['id'] not in found_ids:
                        cid = (o['customs_id'] or '').strip().upper()
                        if not cid.startswith('P') or len(cid) != 13:
                            fail_status = '오류'
                            fail_msg = '13자리수 오류 또는 P시작 형식 오류' if len(cid) != 13 else 'P시작 형식 오류'
                        else:
                            fail_status = '오류'
                            fail_msg = '검증 결과 불일치 또는 미응답'

                        cur.execute('''
                            UPDATE orders 
                            SET customs_validation_status = ?, customs_validation_msg = ?
                            WHERE id = ?
                        ''', (fail_status, fail_msg, o['id']))
                        results.append({
                            'order_id': o['id'],
                            'recipient_name': o['recipient_name'],
                            'customs_id': o['customs_id'],
                            'status': fail_status,
                            'message': fail_msg
                        })

            except Exception as e:
                for o in batch:
                    cur.execute('''
                        UPDATE orders 
                        SET customs_validation_status = ?, customs_validation_msg = ?
                        WHERE id = ?
                    ''', ('장애', f'통신 오류: {str(e)}', o['id']))
                    results.append({
                        'order_id': o['id'],
                        'status': '장애',
                        'message': str(e)
                    })

        conn.commit()
        return jsonify({
            'status': 'success',
            'total_checked': len(results),
            'results': results
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': f'통관부호 검증 처리 중 오류: {str(e)}'}), 500
    finally:
        conn.close()

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

@app.route('/api/orders/bulk-weight', methods=['POST'])
def bulk_update_weight():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    try:
        weight = float(data.get('weight', 0.0) or 0.0)
    except (ValueError, TypeError):
        return jsonify({'error': '유효한 무게(kg) 숫자를 입력해주세요.'}), 400

    if not order_ids:
        return jsonify({'error': '무게를 적용할 주문을 선택하세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(order_ids))
    cur.execute(f'UPDATE orders SET weight = ? WHERE id IN ({placeholders})', [weight] + order_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'updated_count': len(order_ids), 'weight': weight})

# ----------------- 4. 배송 & 송장 & 출고 API -----------------
@app.route('/api/shipping/issue-tracking', methods=['POST'])
def api_issue_tracking():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    courier_name = data.get('courier_name')
    if not order_ids:
        return jsonify({'error': '송장을 발급할 주문을 선택하세요.'}), 400

    res = issue_tracking_numbers(order_ids, courier_name)
    if isinstance(res, dict):
        updated = res.get('updated_orders', [])
        failed = res.get('failed_orders', [])
    else:
        updated = res
        failed = []

    return jsonify({
        'status': 'success',
        'updated': updated,
        'failed': failed,
        'success_count': len(updated),
        'failed_count': len(failed)
    })

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

    if orders:
        order_ids = [o['id'] for o in orders]
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT * FROM order_items WHERE order_id IN ({placeholders}) ORDER BY id ASC', order_ids)
        items_by_order = {}
        for it in cur.fetchall():
            oid = it['order_id']
            if oid not in items_by_order:
                items_by_order[oid] = []
            items_by_order[oid].append(dict(it))
        for o in orders:
            o['items'] = items_by_order.get(o['id'], [])
            o['sender'] = sender_info
    else:
        for o in orders:
            o['items'] = []
            o['sender'] = sender_info

    conn.close()
    return jsonify(orders)

@app.route('/api/shipping/narda-pdf-bundle', methods=['GET'])
def api_narda_pdf_bundle():
    order_ids_str = request.args.get('order_ids', '')
    if not order_ids_str:
        return "출력할 주문이 지정되지 않았습니다.", 400

    order_ids = [int(x) for x in order_ids_str.split(',') if x.strip().isdigit()]
    if not order_ids:
        return "유효한 주문 ID가 없습니다.", 400

    conn = get_connection()
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(order_ids))
    cur.execute(f"SELECT id, order_no, invoice_pdf_url FROM orders WHERE id IN ({placeholders}) AND invoice_pdf_url != ''", order_ids)
    rows = cur.fetchall()
    conn.close()

    if not rows:
        return "선택한 주문에 발급된 나르다 송장 PDF가 없습니다. 먼저 송장을 발급해주세요.", 404

    from pypdf import PdfWriter
    import io

    writer = PdfWriter()
    merged_count = 0

    for r in rows:
        pdf_url = r['invoice_pdf_url']
        try:
            resp = requests.get(pdf_url, timeout=15)
            if resp.status_code == 200 and resp.content:
                pdf_stream = io.BytesIO(resp.content)
                writer.append(pdf_stream)
                merged_count += 1
        except Exception as e:
            print(f"PDF 다운로드 실패 ({r['order_no']}, {pdf_url}): {e}")

    if merged_count == 0:
        return "PDF 파일을 다운로드하여 병합하지 못했습니다.", 500

    output_stream = io.BytesIO()
    writer.write(output_stream)
    output_stream.seek(0)

    return send_file(
        output_stream,
        mimetype='application/pdf',
        as_attachment=False,
        download_name=f'NARDA_통합송장_{merged_count}건.pdf'
    )

# ----------------- 5. 재고 관리 (WMS) 및 입고 API -----------------
@app.route('/api/inventory', methods=['GET'])
def get_inventory():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        SELECT p.*,
            (COALESCE(p.dispatched_quantity, 0) + 
             COALESCE((
                 SELECT SUM(oi.quantity * oi.option_quantity) 
                 FROM order_items oi 
                 JOIN orders o ON oi.order_id = o.id 
                 WHERE oi.sku = p.sku AND o.status = '출고완료'
             ), 0)
            ) AS total_dispatched_quantity
        FROM products p 
        ORDER BY p.id DESC
    ''')
    products = [dict(p) for p in cur.fetchall()]
    for p in products:
        cur.execute('SELECT * FROM product_options WHERE sku = ? OR product_id = ? ORDER BY id ASC', (p['sku'], p['id']))
        p['options'] = [dict(row) for row in cur.fetchall()]
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
    product_url = data.get('product_url', '').strip()
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
                product_url = CASE WHEN ? != '' THEN ? ELSE product_url END,
                unit_price = CASE WHEN ? > 0 THEN ? ELSE unit_price END,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ''', (new_stock, product_name, product_name, english_name, english_name, hs_code, hs_code, product_url, product_url, unit_price, unit_price, prod['id']))
        prod_id = prod['id']
        p_name = product_name or prod['product_name']
    else:
        if not product_name:
            product_name = f"상품_{sku}"
        new_stock = qty
        cur.execute('''
            INSERT INTO products (sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity, product_url)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (sku, product_name, '', english_name, hs_code, unit_price, new_stock, product_url))
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

@app.route('/api/inventory/template', methods=['GET'])
def download_product_template():
    import openpyxl
    from io import BytesIO
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "상품등록양식"

    headers = ["바코드", "상품명", "영문명", "HS코드", "단가", "재고수량", "출고수량", "상품URL"]
    ws.append(headers)

    samples = [
        ["8801234567890", "[프리비] 엘모사 기미스틱", "ELMOSA BLEMISH STICK", "3304.99.1000", 12.50, 100, 0, "https://smartstore.naver.com/example"],
        ["8809876543210", "[일본] 브라질산 그린프로폴리스 캔디", "GREEN PROPOLIS CANDY", "2106.90.9099", 8.00, 50, 0, "https://smartstore.naver.com/example2"]
    ]
    for s in samples:
        ws.append(s)

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 35

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name='상품등록양식.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

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
            elif '출고' in cs: col_map['dispatched_quantity'] = c
            elif '재고' in cs or '수량' in cs: col_map['stock_quantity'] = c
            elif 'url' in cs or '링크' in cs: col_map['product_url'] = c

        if 'sku' not in col_map and 'product_name' not in col_map:
            return jsonify({'error': '엑셀에 [바코드(SKU)] 또는 [상품명] 컬럼이 필요합니다.'}), 400

        # 병합된 셀(바코드, 영문명, HS코드, 재고수량) 처리
        if col_map.get('sku'):
            df[col_map['sku']] = df[col_map['sku']].fillna('').astype(str).str.strip().replace('nan', '')
            if col_map.get('english_name'):
                eng_c = col_map['english_name']
                df[eng_c] = df[eng_c].fillna('').astype(str).str.strip().replace('nan', '')
                df['_grp'] = (df[eng_c] != '').cumsum()
                df[col_map['sku']] = df.groupby('_grp')[col_map['sku']].transform(lambda s: s.replace('', None).bfill().ffill().fillna(''))
            else:
                df[col_map['sku']] = df[col_map['sku']].replace('', None).ffill().fillna('')

        conn = get_connection()
        cur = conn.cursor()
        saved_count = 0
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

        for _, row in df.iterrows():
            sku = str(row.get(col_map.get('sku', ''), '')).strip()
            name = str(row.get(col_map.get('product_name', ''), '')).strip()
            eng = str(row.get(col_map.get('english_name', ''), '')).strip()
            hs = str(row.get(col_map.get('hs_code', ''), '')).strip()
            url = str(row.get(col_map.get('product_url', ''), '')).strip()
            if url.lower() == 'nan': url = ''
            
            try:
                price = float(row.get(col_map.get('unit_price', ''), 0.0) or 0.0)
            except:
                price = 0.0
            try:
                stock = int(float(row.get(col_map.get('stock_quantity', ''), 0) or 0))
            except:
                stock = 0
            try:
                dispatched = int(float(row.get(col_map.get('dispatched_quantity', ''), 1) or 1))
            except:
                dispatched = 1

            if not sku and not name:
                continue
            if not sku:
                sku = f"SKU-{abs(hash(name)) % 1000000:06d}"

            cur.execute('SELECT * FROM products WHERE sku = ?', (sku,))
            existing = cur.fetchone()

            if existing:
                cur.execute('''
                    UPDATE products
                    SET product_name = CASE WHEN ? != '' AND (product_name = '' OR product_name LIKE '상품_%') THEN ? ELSE product_name END,
                        english_name = CASE WHEN ? != '' THEN ? ELSE english_name END,
                        hs_code = CASE WHEN ? != '' THEN ? ELSE hs_code END,
                        product_url = CASE WHEN ? != '' THEN ? ELSE product_url END,
                        stock_quantity = CASE WHEN ? > 0 THEN ? ELSE stock_quantity END,
                        unit_price = CASE WHEN ? > 0 THEN ? ELSE unit_price END,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, name, eng, eng, hs, hs, url, url, stock, stock, price, price, existing['id']))
                prod_id = existing['id']
            else:
                cur.execute('''
                    INSERT INTO products (sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (sku, name or f"상품_{sku}", '', eng, hs, price, stock, dispatched, url))
                prod_id = cur.lastrowid

                cur.execute('''
                    INSERT INTO inventory_logs (product_id, sku, product_name, change_type, change_qty, resulting_stock, memo, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (prod_id, sku, name or sku, '엑셀등록', stock, stock, '엑셀 상품 일괄 등록', now_str))

            # 옵션 마스터(product_options)에 옵션 등록
            if name:
                cur.execute('SELECT id FROM product_options WHERE sku = ? AND option_name = ?', (sku, name))
                opt_row = cur.fetchone()
                if opt_row:
                    cur.execute('''
                        UPDATE product_options
                        SET unit_price = ?, dispatched_quantity = ?, updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                    ''', (price, dispatched, opt_row['id']))
                else:
                    cur.execute('''
                        INSERT INTO product_options (product_id, sku, option_name, unit_price, dispatched_quantity)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (prod_id, sku, name, price, dispatched))

                # 미등록 상품 목록에 있으면 등록 완료 처리
                cur.execute("UPDATE unregistered_products SET status = 'registered' WHERE combined_name = ? OR product_name = ?", (name, name))

            saved_count += 1

        from excel_processor import sync_order_items_sku
        sync_order_items_sku(conn)

        conn.commit()
        conn.close()
        return jsonify({'status': 'success', 'message': f"{saved_count}개 상품/옵션이 엑셀에서 성공적으로 등록/반영되었습니다."})
    except Exception as e:
        return jsonify({'error': f'엑셀 처리 중 오류: {str(e)}'}), 500

@app.route('/api/inventory/product', methods=['POST'])
def save_product():
    try:
        data = request.get_json() or {}
        sku = data.get('sku', '').strip()
        product_name = data.get('product_name', '').strip()
        english_name = data.get('english_name', '').strip()
        hs_code = data.get('hs_code', '').strip()
        unit_price = float(data.get('unit_price', 0.0) or 0.0)
        stock_quantity = int(data.get('stock_quantity', 0) or 0)
        dispatched_quantity = int(data.get('dispatched_quantity', 0) or 0)
        product_url = data.get('product_url', '').strip()
        prod_id = data.get('id')
        options = data.get('options', [])

        if not sku or not product_name:
            return jsonify({'error': 'SKU(바코드)와 상품명은 필수입니다.'}), 400

        conn = get_connection()
        cur = conn.cursor()
        if prod_id:
            cur.execute('''
                UPDATE products
                SET sku = ?, product_name = ?, english_name = ?, hs_code = ?, unit_price = ?, stock_quantity = ?, dispatched_quantity = ?, product_url = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (sku, product_name, english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url, prod_id))
        else:
            cur.execute('''
                INSERT INTO products (sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (sku, product_name, '', english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url))
            prod_id = cur.lastrowid

        # 옵션 목록 갱신 (전달된 옵션이 있는 경우)
        if options is not None:
            cur.execute('DELETE FROM product_options WHERE sku = ? OR product_id = ?', (sku, prod_id))
            for opt in options:
                opt_name = str(opt.get('option_name', '')).strip()
                if not opt_name:
                    continue
                try:
                    opt_price = float(opt.get('unit_price', 0.0) or 0.0)
                except:
                    opt_price = 0.0
                try:
                    opt_qty = int(opt.get('dispatched_quantity', 1) or 1)
                except:
                    opt_qty = 1

                cur.execute('''
                    INSERT INTO product_options (product_id, sku, option_name, unit_price, dispatched_quantity)
                    VALUES (?, ?, ?, ?, ?)
                ''', (prod_id, sku, opt_name, opt_price, opt_qty))

                cur.execute("UPDATE unregistered_products SET status = 'registered' WHERE combined_name = ? OR product_name = ?", (opt_name, opt_name))

        # 미등록 상품 테이블 상태 업데이트
        cur.execute("UPDATE unregistered_products SET status = 'registered' WHERE combined_name = ? OR product_name = ?", (product_name, product_name))

        # 주문 품목들의 바코드 동기화
        from excel_processor import sync_order_items_sku
        sync_order_items_sku(conn)

        conn.commit()
        conn.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'error': f'상품 저장 중 오류: {str(e)}'}), 500

@app.route('/api/inventory/products/delete', methods=['POST'])
def delete_products_batch():
    data = request.get_json() or {}
    product_ids = data.get('product_ids', [])
    if not product_ids:
        return jsonify({'error': '삭제할 상품을 선택하세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(product_ids))
    cur.execute(f'DELETE FROM product_options WHERE product_id IN ({placeholders})', product_ids)
    cur.execute(f'DELETE FROM inventory_logs WHERE product_id IN ({placeholders})', product_ids)
    cur.execute(f'DELETE FROM products WHERE id IN ({placeholders})', product_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'deleted_count': len(product_ids)})

@app.route('/api/inventory/product/<int:prod_id>', methods=['DELETE'])
def delete_product(prod_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM product_options WHERE product_id = ?', (prod_id,))
    cur.execute('DELETE FROM inventory_logs WHERE product_id = ?', (prod_id,))
    cur.execute('DELETE FROM products WHERE id = ?', (prod_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/inventory/unregistered-products', methods=['GET'])
def get_unregistered_products():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM unregistered_products WHERE status = 'pending' ORDER BY order_count DESC, id DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify(rows)

@app.route('/api/inventory/unregistered-products/<int:item_id>/ignore', methods=['POST'])
def ignore_unregistered_product(item_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("UPDATE unregistered_products SET status = 'ignored' WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/orders/move-status', methods=['POST'])
def move_orders_status():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    target_status = data.get('target_status', '통관오류')
    if not order_ids:
        return jsonify({'error': '이동할 주문을 선택하세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(order_ids))
    cur.execute(f'UPDATE orders SET status = ? WHERE id IN ({placeholders})', [target_status] + order_ids)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'moved_count': len(order_ids), 'target_status': target_status})


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

@app.route('/api/inventory/logs/delete', methods=['POST'])
def delete_inventory_logs():
    data = request.get_json() or {}
    log_ids = data.get('log_ids', [])
    if not log_ids:
        return jsonify({'error': '삭제할 입출고 내역을 선택하세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(log_ids))
    cur.execute(f'DELETE FROM inventory_logs WHERE id IN ({placeholders})', log_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'deleted_count': len(log_ids)})

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
        SET courier_name = ?, customer_code = ?, api_key = ?, narda_id = ?,
            sender_name = ?, sender_phone = ?, sender_address = ?,
            is_mock = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = (SELECT id FROM courier_settings LIMIT 1)
    ''', (
        data.get('courier_name', '나르다'),
        data.get('customer_code', ''),
        data.get('api_key', ''),
        data.get('narda_id', 'calroid'),
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



