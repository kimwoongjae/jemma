import os
import json
import math
import re
import uuid
import time
import secrets
import requests
from datetime import datetime
from email.utils import parsedate_to_datetime
from flask import Flask, request, jsonify, render_template, send_file, redirect, session, g
from werkzeug.security import generate_password_hash, check_password_hash

from werkzeug.utils import secure_filename

from database import init_db, get_connection
from excel_processor import NARDA_MANIFEST_PROFILE, process_excel_file, process_narda_manifest_file, save_orders_to_db
from shipping_api import (
    issue_tracking_numbers, complete_dispatch,
    export_shipping_excel
)

SHIPPING_WEIGHT_STEPS = [i / 2 for i in range(1, 61)] + list(range(35, 101, 5))


def complete_shipping_rate_rows(rows):
    """Return the configured weight grid, filling unregistered prices in memory."""
    existing = {round(float(row['weight_kg']), 1): dict(row) for row in rows}
    return [
        existing.get(weight, {'weight_kg': weight, 'air_rate': 0, 'sea_rate': 0})
        for weight in SHIPPING_WEIGHT_STEPS
    ]

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or os.urandom(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=os.environ.get('SESSION_COOKIE_SECURE', '0') == '1',
    PERMANENT_SESSION_LIFETIME=60 * 60 * 12
)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024  # 30MB
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# DB 초기화
init_db()


@app.before_request
def require_account_session():
    public_endpoints = {'auth_page', 'auth_login', 'auth_register', 'auth_join', 'auth_logout'}
    if request.endpoint == 'static' or request.endpoint in public_endpoints:
        return None

    role = session.get('account_role')
    if role == 'admin':
        if request.endpoint == 'admin_page' or request.endpoint and request.endpoint.startswith('admin_'):
            return None
        admin_company_id = session.get('admin_tenant_id')
        if admin_company_id:
            try:
                g.company_id = int(admin_company_id)
            except (TypeError, ValueError):
                session.pop('admin_tenant_id', None)
                session.pop('company_id', None)
                session.pop('company_name', None)
                return redirect('/admin')
            return None
        if request.path.startswith('/api/'):
            return jsonify({'error': '업체 계정으로 로그인해주세요.'}), 403
        return redirect('/admin')

    company_id = session.get('company_id')
    if not company_id:
        if request.path.startswith('/api/'):
            return jsonify({'error': '로그인이 필요합니다.'}), 401
        return redirect('/login')
    try:
        g.company_id = int(company_id)
    except (TypeError, ValueError):
        session.clear()
        return jsonify({'error': '로그인 정보가 올바르지 않습니다.'}), 401
    if request.endpoint == 'admin_page' or request.endpoint and request.endpoint.startswith('admin_'):
        return jsonify({'error': '관리자 계정으로 로그인해주세요.'}), 403
    return None


@app.route('/login', methods=['GET'])
def auth_page():
    return render_template('auth.html')


@app.route('/api/auth/login', methods=['POST'])
def auth_login():
    data = request.get_json() or {}
    username = str(data.get('username') or '').strip()
    password = str(data.get('password') or '')
    login_type = str(data.get('login_type') or 'user').strip().lower()
    admin_username = os.environ.get('ADMIN_USERNAME', '').strip()
    admin_password = os.environ.get('ADMIN_PASSWORD', '')
    if login_type == 'admin' and not (admin_username and admin_password):
        return jsonify({'error': '관리자 계정이 설정되지 않았습니다. Render의 ADMIN_USERNAME과 ADMIN_PASSWORD를 먼저 설정해주세요.'}), 503
    if admin_username and username == admin_username and admin_password and secrets.compare_digest(password, admin_password):
        session.clear()
        session.permanent = True
        session['account_role'] = 'admin'
        session['username'] = username
        return jsonify({'status': 'success', 'role': 'admin', 'redirect': '/admin'})
    if login_type == 'admin':
        return jsonify({'error': '관리자 아이디 또는 비밀번호가 올바르지 않습니다.'}), 401

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''SELECT users.id, users.username, users.password_hash, users.company_id, users.role,
                          companies.name AS company_name, companies.active AS company_active
                   FROM users JOIN companies ON companies.id = users.company_id
                   WHERE users.username = ? AND users.active = 1''', (username,))
    user = cur.fetchone()
    conn.close()
    if not user or not user['company_active'] or not check_password_hash(user['password_hash'], password):
        return jsonify({'error': '아이디 또는 비밀번호가 올바르지 않습니다.'}), 401
    session.clear()
    session.permanent = True
    session['account_role'] = 'user'
    session['user_id'] = user['id']
    session['username'] = user['username']
    session['company_id'] = user['company_id']
    session['company_name'] = user['company_name']
    return jsonify({'status': 'success', 'role': 'user', 'redirect': '/'})


def _create_company_account(company_name, username, password, join_code=None):
    company_name = str(company_name or '').strip()
    username = str(username or '').strip()
    if not re.fullmatch(r'[A-Za-z0-9_.-]{3,40}', username):
        return None, '아이디는 영문·숫자·밑줄·마침표·하이픈으로 3~40자 입력해주세요.'
    reserved_admin = os.environ.get('ADMIN_USERNAME', '').strip()
    if reserved_admin and username.casefold() == reserved_admin.casefold():
        return None, '관리자 아이디는 업체 계정에서 사용할 수 없습니다.'
    if len(password) < 10:
        return None, '비밀번호는 10자 이상 입력해주세요.'
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id FROM users WHERE LOWER(username) = LOWER(?)', (username,))
    if cur.fetchone():
        conn.close()
        return None, '이미 사용 중인 아이디입니다.'
    if join_code:
        cur.execute('SELECT id, name FROM companies WHERE invite_code = ? AND active = 1', (str(join_code).strip().upper(),))
        company = cur.fetchone()
        if not company:
            conn.close()
            return None, '초대 코드가 올바르지 않거나 사용할 수 없습니다.'
        company_id, company_name = company['id'], company['name']
    else:
        if not company_name or len(company_name) > 100:
            return None, '업체명을 1~100자로 입력해주세요.'
        cur.execute('SELECT id FROM companies WHERE LOWER(name) = LOWER(?)', (company_name,))
        if cur.fetchone():
            conn.close()
            return None, '이미 등록된 업체명입니다. 해당 업체의 초대 코드로 가입해주세요.'
        invite_code = secrets.token_urlsafe(9).replace('-', '').replace('_', '').upper()[:12]
        cur.execute('INSERT INTO companies (name, invite_code) VALUES (?, ?)', (company_name, invite_code))
        company_id = cur.lastrowid

        api_key = os.environ.get('NARDA_API_KEY', '')
        customer_code = os.environ.get('NARDA_CUSTOMER_CODE', 'TOA')
        narda_id = os.environ.get('NARDA_ID', 'calroid')
        cur.execute('''INSERT INTO courier_settings
            (tenant_id, courier_name, customer_code, api_key, narda_id, sender_name, sender_phone, sender_address, is_mock, stock_only_dispatch)
            VALUES (?, '나르다', ?, ?, ?, '', '', '', 0, 0)''', (company_id, customer_code, api_key, narda_id))
        cur.execute('INSERT INTO split_delimiters (tenant_id, symbol) VALUES (?, ?)', (company_id, '+'))
        for symbol in ['-', '/', ',', '?', '&']:
            cur.execute('INSERT INTO split_delimiters (tenant_id, symbol) VALUES (?, ?)', (company_id, symbol))

    role = 'owner' if not join_code else 'user'
    cur.execute('INSERT INTO users (username, password_hash, company_id, role) VALUES (?, ?, ?, ?)',
                (username, generate_password_hash(password), company_id, role))
    user_id = cur.lastrowid
    conn.commit()
    cur.execute('SELECT invite_code FROM companies WHERE id = ?', (company_id,))
    code_row = cur.fetchone()
    invite_code = code_row['invite_code'] if code_row else ''
    conn.close()
    return {'user_id': user_id, 'company_id': company_id, 'company_name': company_name, 'username': username, 'invite_code': invite_code}, None


@app.route('/api/auth/register', methods=['POST'])
def auth_register():
    data = request.get_json() or {}
    account, error = _create_company_account(data.get('company_name'), data.get('username'), str(data.get('password') or ''))
    if error:
        return jsonify({'error': error}), 400
    session.clear()
    session.permanent = True
    session['account_role'] = 'user'
    session['user_id'] = account['user_id']
    session['username'] = account['username']
    session['company_id'] = account['company_id']
    session['company_name'] = account['company_name']
    return jsonify({'status': 'success', 'redirect': '/', 'invite_code': account['invite_code']})


@app.route('/api/auth/join', methods=['POST'])
def auth_join():
    data = request.get_json() or {}
    account, error = _create_company_account('', data.get('username'), str(data.get('password') or ''), data.get('invite_code'))
    if error:
        return jsonify({'error': error}), 400
    session.clear()
    session.permanent = True
    session['account_role'] = 'user'
    session['user_id'] = account['user_id']
    session['username'] = account['username']
    session['company_id'] = account['company_id']
    session['company_name'] = account['company_name']
    return jsonify({'status': 'success', 'redirect': '/'})


@app.route('/api/auth/logout', methods=['POST'])
def auth_logout():
    session.clear()
    return jsonify({'status': 'success', 'redirect': '/login'})


@app.route('/api/account/password', methods=['POST'])
def change_account_password():
    if session.get('account_role') != 'user' or not session.get('user_id'):
        return jsonify({'error': '업체 계정으로 로그인해야 비밀번호를 변경할 수 있습니다.'}), 403
    data = request.get_json() or {}
    current_password = str(data.get('current_password') or '')
    new_password = str(data.get('new_password') or '')
    if not current_password:
        return jsonify({'error': '현재 비밀번호를 입력해주세요.'}), 400
    if len(new_password) < 10:
        return jsonify({'error': '새 비밀번호는 10자 이상 입력해주세요.'}), 400
    if current_password == new_password:
        return jsonify({'error': '현재 비밀번호와 다른 비밀번호를 입력해주세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT id, password_hash FROM users WHERE id = ? AND company_id = ? AND active = 1',
                    (session['user_id'], session.get('company_id')))
        user = cur.fetchone()
        if not user or not check_password_hash(user['password_hash'], current_password):
            return jsonify({'error': '현재 비밀번호가 올바르지 않습니다.'}), 400
        cur.execute('UPDATE users SET password_hash = ? WHERE id = ? AND company_id = ?',
                    (generate_password_hash(new_password), session['user_id'], session.get('company_id')))
        conn.commit()
        return jsonify({'status': 'success', 'message': '비밀번호를 변경했습니다.'})
    except Exception:
        conn.rollback()
        app.logger.exception('Account password change failed')
        return jsonify({'error': '비밀번호 변경에 실패했습니다. 다시 시도해주세요.'}), 500
    finally:
        conn.close()


@app.route('/admin', methods=['GET'])
def admin_page():
    if session.get('account_role') != 'admin':
        return redirect('/login')
    return redirect('/admin/companies')


@app.route('/admin/companies', methods=['GET'])
def admin_companies_page():
    if session.get('account_role') != 'admin':
        return redirect('/login')
    return render_template('admin.html', active_page='companies')


@app.route('/admin/shipments', methods=['GET'])
def admin_shipments_page():
    if session.get('account_role') != 'admin':
        return redirect('/login')
    return render_template('admin.html', active_page='shipments')


@app.route('/admin/inventory', methods=['GET'])
def admin_inventory_page():
    if session.get('account_role') != 'admin':
        return redirect('/login')
    return render_template('admin.html', active_page='inventory')


def _admin_inventory_company_exists(cur, company_id):
    cur.execute('SELECT id FROM companies WHERE id = ? AND active = 1', (company_id,))
    return cur.fetchone()


@app.route('/api/admin/companies/<int:company_id>/inventory/search', methods=['GET'])
def admin_inventory_search(company_id):
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    barcode = str(request.args.get('barcode') or '').strip()
    if not barcode:
        return jsonify({'error': '바코드를 입력해주세요.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    try:
        if not _admin_inventory_company_exists(cur, company_id):
            return jsonify({'error': '활성 거래처를 찾을 수 없습니다.'}), 404
        cur.execute('''SELECT id, sku, product_name, option_name, english_name, hs_code,
                              unit_price, stock_quantity, product_url
                       FROM products WHERE tenant_id = ? AND sku = ? LIMIT 1''',
                    (company_id, barcode))
        product = cur.fetchone()
        if not product:
            return jsonify({'found': False, 'message': '해당 거래처 재고에 등록되지 않은 바코드입니다.'})
        return jsonify({'found': True, 'product': dict(product)})
    finally:
        conn.close()


@app.route('/api/admin/companies/<int:company_id>/inventory/movement', methods=['POST'])
def admin_inventory_movement(company_id):
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    data = request.get_json(silent=True) or {}
    barcode = str(data.get('barcode') or '').strip()
    movement = str(data.get('movement') or '').strip()
    memo = str(data.get('memo') or '').strip()[:250]
    try:
        raw_quantity = data.get('quantity')
        numeric_quantity = float(raw_quantity)
        if not math.isfinite(numeric_quantity) or not numeric_quantity.is_integer():
            raise ValueError('quantity must be an integer')
        quantity = int(numeric_quantity)
    except (TypeError, ValueError):
        return jsonify({'error': '수량은 1 이상의 정수로 입력해주세요.'}), 400
    if not barcode:
        return jsonify({'error': '바코드를 입력해주세요.'}), 400
    if quantity <= 0:
        return jsonify({'error': '수량은 1 이상이어야 합니다.'}), 400
    if movement not in ('입고', '출고'):
        return jsonify({'error': '입고 또는 출고를 선택해주세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        if not _admin_inventory_company_exists(cur, company_id):
            return jsonify({'error': '활성 거래처를 찾을 수 없습니다.'}), 404
        cur.execute('SELECT id, sku, product_name, option_name, stock_quantity FROM products WHERE tenant_id = ? AND sku = ? LIMIT 1',
                    (company_id, barcode))
        product = cur.fetchone()
        if not product:
            return jsonify({'error': '해당 거래처 재고에서 바코드 상품을 찾을 수 없습니다.'}), 404
        current_stock = int(product['stock_quantity'] or 0)
        if movement == '출고' and quantity > current_stock:
            return jsonify({'error': f'재고가 부족합니다. 현재 재고는 {current_stock}개입니다.'}), 409
        change_qty = quantity if movement == '입고' else -quantity
        new_stock = current_stock + change_qty
        cur.execute('UPDATE products SET stock_quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ? AND tenant_id = ?',
                    (new_stock, product['id'], company_id))
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur.execute('''INSERT INTO inventory_logs
                       (tenant_id, product_id, sku, product_name, option_name, change_type, change_qty,
                        resulting_stock, memo, created_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                    (company_id, product['id'], product['sku'], product['product_name'],
                     product['option_name'] or '', movement, change_qty, new_stock,
                     memo or f'관리자 바코드 {movement}', now_str))
        conn.commit()
        return jsonify({'status': 'success', 'movement': movement, 'quantity': quantity,
                        'stock_quantity': new_stock, 'product_name': product['product_name'],
                        'sku': product['sku'], 'created_at': now_str})
    except Exception:
        conn.rollback()
        app.logger.exception('Admin inventory movement failed for company %s', company_id)
        return jsonify({'error': '입출고 저장에 실패했습니다. 다시 시도해주세요.'}), 500
    finally:
        conn.close()


@app.route('/admin/companies/<int:company_id>/shipping-rates', methods=['GET'])
def admin_company_shipping_rates_page(company_id):
    if session.get('account_role') != 'admin':
        return redirect('/login')
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id, name FROM companies WHERE id = ? AND active = 1', (company_id,))
    company = cur.fetchone()
    conn.close()
    if not company:
        return redirect('/admin/companies')
    return render_template('admin_shipping_rates.html', company=dict(company))


def _admin_company_for_shipping_rates(company_id):
    if session.get('account_role') != 'admin':
        return None
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id FROM companies WHERE id = ? AND active = 1', (company_id,))
    company = cur.fetchone()
    conn.close()
    return company


@app.route('/api/admin/companies/<int:company_id>/shipping-rates', methods=['GET'])
def admin_get_company_shipping_rates(company_id):
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT id FROM companies WHERE id = ? AND active = 1', (company_id,))
        if not cur.fetchone():
            return jsonify({'error': '활성 거래처를 찾을 수 없습니다.'}), 404
        cur.execute('SELECT weight_kg, air_rate, sea_rate FROM shipping_rates WHERE tenant_id = ? ORDER BY weight_kg ASC', (company_id,))
        rows = complete_shipping_rate_rows(cur.fetchall())
        cur.execute('SELECT additional_package_rate, box_package_rate FROM shipping_packaging_rates WHERE tenant_id = ?', (company_id,))
        packaging = cur.fetchone()
        return jsonify({'rates': [dict(row) for row in rows], 'packaging_fees': dict(packaging) if packaging else {'additional_package_rate': 0, 'box_package_rate': 0}})
    finally:
        conn.close()


@app.route('/api/admin/companies/<int:company_id>/shipping-rates', methods=['POST'])
def admin_save_company_shipping_rates(company_id):
    if not _admin_company_for_shipping_rates(company_id):
        return jsonify({'error': '관리자 권한 또는 활성 거래처 확인이 필요합니다.'}), 403
    g.company_id = company_id
    return save_shipping_rates()


@app.route('/api/admin/companies/<int:company_id>/shipping-rates/upload', methods=['POST'])
def admin_upload_company_shipping_rates(company_id):
    if not _admin_company_for_shipping_rates(company_id):
        return jsonify({'error': '관리자 권한 또는 활성 거래처 확인이 필요합니다.'}), 403
    g.company_id = company_id
    return upload_shipping_rates_excel()


@app.route('/api/admin/companies/<int:company_id>/shipping-rates/template', methods=['GET'])
def admin_download_company_shipping_rates_template(company_id):
    if not _admin_company_for_shipping_rates(company_id):
        return jsonify({'error': '관리자 권한 또는 활성 거래처 확인이 필요합니다.'}), 403
    return download_shipping_rates_template()


@app.route('/api/admin/companies/<int:company_id>/access', methods=['POST'])
def admin_access_company(company_id):
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id, name FROM companies WHERE id = ?', (company_id,))
    company = cur.fetchone()
    conn.close()
    if not company:
        return jsonify({'error': '업체를 찾을 수 없습니다.'}), 404
    session['admin_tenant_id'] = company['id']
    session['company_id'] = company['id']
    session['company_name'] = company['name']
    return jsonify({'status': 'success', 'redirect': '/'})


@app.route('/admin/return', methods=['POST'])
def admin_return_to_dashboard():
    if session.get('account_role') != 'admin':
        return redirect('/login')
    session.pop('admin_tenant_id', None)
    session.pop('company_id', None)
    session.pop('company_name', None)
    return redirect('/admin')


@app.route('/api/admin/companies', methods=['GET'])
def admin_companies():
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''SELECT c.id, c.name, c.invite_code, c.active, c.created_at,
                          COUNT(DISTINCT o.id) AS order_count,
                          COUNT(DISTINCT CASE WHEN o.status = '출고완료' THEN o.id END) AS dispatched_count,
                          COUNT(DISTINCT u.id) AS user_count
                   FROM companies c
                   LEFT JOIN orders o ON o.tenant_id = c.id
                   LEFT JOIN users u ON u.company_id = c.id
                   GROUP BY c.id, c.name, c.invite_code, c.active, c.created_at
                   ORDER BY c.id ASC''')
    companies = [dict(row) for row in cur.fetchall()]
    conn.close()
    return jsonify(companies)


@app.route('/api/admin/companies/<int:company_id>/delete', methods=['POST'])
def admin_delete_company(company_id):
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT id, name FROM companies WHERE id = ?', (company_id,))
        company = cur.fetchone()
        if not company:
            return jsonify({'error': '거래처를 찾을 수 없습니다.'}), 404

        # Delete only records owned by this company.
        for table in ('order_items', 'orders', 'product_options', 'inventory_logs', 'products',
                      'courier_settings', 'upload_profiles', 'bundle_rules', 'split_delimiters',
                      'unregistered_products', 'shipping_rates'):
            cur.execute(f'DELETE FROM {table} WHERE tenant_id = ?', (company_id,))
        cur.execute('DELETE FROM users WHERE company_id = ?', (company_id,))
        if company['name'] == '젬마 (기존 데이터)':
            cur.execute('''INSERT INTO system_flags (flag_key, flag_value) VALUES (?, ?)
                           ON CONFLICT(flag_key) DO UPDATE SET flag_value = excluded.flag_value''',
                        ('legacy_gemma_deleted', '1'))
        cur.execute('DELETE FROM companies WHERE id = ?', (company_id,))
        conn.commit()
        if session.get('admin_tenant_id') == company_id:
            session.pop('admin_tenant_id', None)
            session.pop('company_id', None)
            session.pop('company_name', None)
        return jsonify({'status': 'success', 'deleted_company': company['name']})
    except Exception:
        conn.rollback()
        app.logger.exception('Company deletion failed')
        return jsonify({'error': '거래처 삭제에 실패했습니다. 변경 내용은 저장되지 않았습니다.'}), 500
    finally:
        conn.close()


@app.route('/api/admin/companies/<int:target_company_id>/copy-legacy-catalog', methods=['POST'])
def admin_copy_legacy_catalog(target_company_id):
    """Copy the legacy Jemma product catalog and upload profiles to a new company.

    Existing destination catalog, upload, and split settings are preserved, so retries are safe.
    Orders and shipping data are intentionally not included.
    """
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT id, name FROM companies WHERE name = ?", ('젬마 (기존 데이터)',))
        source = cur.fetchone()
        cur.execute('SELECT id, name FROM companies WHERE id = ?', (target_company_id,))
        target = cur.fetchone()
        if not source:
            return jsonify({'error': '기존 데이터 업체(젬마 (기존 데이터))를 찾을 수 없습니다.'}), 404
        if not target:
            return jsonify({'error': '대상 업체를 찾을 수 없습니다.'}), 404
        source_id = int(source['id'])
        if source_id == target_company_id:
            return jsonify({'error': '기존 업체와 대상 업체가 같습니다.'}), 400

        cur.execute('SELECT * FROM products WHERE tenant_id = ? ORDER BY id', (source_id,))
        source_products = cur.fetchall()
        copied_products = skipped_products = copied_options = skipped_options = 0
        product_id_map = {}
        product_columns = (
            'sku', 'product_name', 'option_name', 'english_name', 'hs_code', 'unit_price',
            'stock_quantity', 'product_url', 'safety_stock', 'location', 'dispatched_quantity',
            'created_at', 'updated_at'
        )
        for product in source_products:
            sku = product['sku']
            if sku:
                cur.execute('SELECT id FROM products WHERE tenant_id = ? AND sku = ?', (target_company_id, sku))
            else:
                cur.execute('''SELECT id FROM products WHERE tenant_id = ? AND COALESCE(sku, '') = ''
                               AND product_name = ? AND COALESCE(option_name, '') = COALESCE(?, '')''',
                            (target_company_id, product['product_name'], product['option_name']))
            existing = cur.fetchone()
            if existing:
                product_id_map[product['id']] = existing['id']
                skipped_products += 1
                continue
            values = [product[column] for column in product_columns]
            cur.execute(
                f"INSERT INTO products (tenant_id, {', '.join(product_columns)}) "
                f"VALUES ({', '.join(['?'] * (len(product_columns) + 1))})",
                [target_company_id] + values
            )
            product_id_map[product['id']] = cur.lastrowid
            copied_products += 1

        cur.execute('SELECT * FROM product_options WHERE tenant_id = ? ORDER BY id', (source_id,))
        source_options = cur.fetchall()
        for option in source_options:
            target_product_id = product_id_map.get(option['product_id'])
            if target_product_id is None:
                cur.execute('SELECT id FROM products WHERE tenant_id = ? AND sku = ?', (target_company_id, option['sku']))
                product = cur.fetchone()
                target_product_id = product['id'] if product else None
            if target_product_id is None:
                skipped_options += 1
                continue
            cur.execute(
                'SELECT id FROM product_options WHERE tenant_id = ? AND product_id = ? AND sku = ? AND option_name = ?',
                (target_company_id, target_product_id, option['sku'], option['option_name'])
            )
            if cur.fetchone():
                skipped_options += 1
                continue
            cur.execute('''INSERT INTO product_options
                (tenant_id, product_id, sku, option_name, unit_price, dispatched_quantity, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (target_company_id, target_product_id, option['sku'], option['option_name'],
                 option['unit_price'], option['dispatched_quantity'], option['created_at'], option['updated_at']))
            copied_options += 1

        cur.execute('SELECT profile_name, mapping_json, header_row, sort_priority, sort_field, sort_direction, sort_rules_json, created_at FROM upload_profiles WHERE tenant_id = ? ORDER BY profile_name', (source_id,))
        source_profiles = cur.fetchall()
        copied_profiles = skipped_profiles = 0
        for profile in source_profiles:
            cur.execute('SELECT profile_name FROM upload_profiles WHERE tenant_id = ? AND profile_name = ?',
                        (target_company_id, profile['profile_name']))
            if cur.fetchone():
                skipped_profiles += 1
                continue
            cur.execute('''INSERT INTO upload_profiles (tenant_id, profile_name, mapping_json, header_row, sort_priority, sort_field, sort_direction, sort_rules_json, created_at)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                        (target_company_id, profile['profile_name'], profile['mapping_json'],
                         profile['header_row'], profile['sort_priority'], profile['sort_field'], profile['sort_direction'], profile['sort_rules_json'], profile['created_at']))
            copied_profiles += 1

        cur.execute('SELECT symbol FROM split_delimiters WHERE tenant_id = ? ORDER BY id', (source_id,))
        source_delimiters = cur.fetchall()
        copied_delimiters = skipped_delimiters = 0
        for delimiter in source_delimiters:
            cur.execute('SELECT id FROM split_delimiters WHERE tenant_id = ? AND symbol = ?',
                        (target_company_id, delimiter['symbol']))
            if cur.fetchone():
                skipped_delimiters += 1
                continue
            cur.execute('INSERT INTO split_delimiters (tenant_id, symbol) VALUES (?, ?)',
                        (target_company_id, delimiter['symbol']))
            copied_delimiters += 1

        cur.execute('''SELECT rule_name, source_product_name, source_option_name, target_product_name,
                              target_option_name, target_sku, multiplier, is_active, created_at
                       FROM bundle_rules WHERE tenant_id = ? ORDER BY id''', (source_id,))
        source_rules = cur.fetchall()
        copied_rules = skipped_rules = 0
        for rule in source_rules:
            cur.execute('''SELECT id FROM bundle_rules WHERE tenant_id = ? AND COALESCE(rule_name, '') = COALESCE(?, '')
                           AND source_product_name = ? AND source_option_name = ?
                           AND target_product_name = ? AND target_option_name = ?
                           AND COALESCE(target_sku, '') = COALESCE(?, '')
                           AND multiplier = ? AND is_active = ?''',
                        (target_company_id, rule['rule_name'], rule['source_product_name'], rule['source_option_name'],
                         rule['target_product_name'], rule['target_option_name'], rule['target_sku'],
                         rule['multiplier'], rule['is_active']))
            if cur.fetchone():
                skipped_rules += 1
                continue
            cur.execute('''INSERT INTO bundle_rules
                (tenant_id, rule_name, source_product_name, source_option_name, target_product_name,
                 target_option_name, target_sku, multiplier, is_active, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                (target_company_id, rule['rule_name'], rule['source_product_name'], rule['source_option_name'],
                 rule['target_product_name'], rule['target_option_name'], rule['target_sku'],
                 rule['multiplier'], rule['is_active'], rule['created_at']))
            copied_rules += 1

        conn.commit()
        return jsonify({
            'status': 'success', 'source_company': source['name'], 'target_company': target['name'],
            'copied_products': copied_products, 'skipped_products': skipped_products,
            'copied_options': copied_options, 'skipped_options': skipped_options,
            'copied_profiles': copied_profiles, 'skipped_profiles': skipped_profiles,
            'copied_delimiters': copied_delimiters, 'skipped_delimiters': skipped_delimiters,
            'copied_rules': copied_rules, 'skipped_rules': skipped_rules
        })
    except Exception as error:
        conn.rollback()
        app.logger.exception('Legacy product catalog copy failed')
        return jsonify({'error': '데이터 복사에 실패했습니다. 변경 내용은 저장되지 않았습니다.'}), 500
    finally:
        conn.close()


@app.route('/api/admin/orders', methods=['GET'])
def admin_company_orders():
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    company_id = request.args.get('company_id', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to = request.args.get('date_to', '').strip()
    view = request.args.get('view', 'all').strip()
    if date_from and not re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_from):
        return jsonify({'error': '시작 날짜 형식이 올바르지 않습니다.'}), 400
    if date_to and not re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_to):
        return jsonify({'error': '종료 날짜 형식이 올바르지 않습니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    date_expr = "SUBSTR(CAST(o.created_at AS TEXT), 1, 10)"
    conditions = ['1=1']
    params = []
    if view == 'shipments':
        conditions.append("(o.status IN ('출고완료', '송장발급') OR TRIM(COALESCE(o.tracking_no, '')) != '')")
    if company_id.isdigit():
        conditions.append('o.tenant_id = ?')
        params.append(int(company_id))
    if date_from:
        conditions.append(f'{date_expr} >= ?')
        params.append(date_from)
    if date_to:
        conditions.append(f'{date_expr} <= ?')
        params.append(date_to)
    cur.execute(f'''SELECT o.*, c.name AS company_name FROM orders o
                    JOIN companies c ON c.id = o.tenant_id
                    WHERE {' AND '.join(conditions)}
                    ORDER BY o.created_at DESC, o.id DESC
                    LIMIT 1000''', params)
    orders = [dict(row) for row in cur.fetchall()]
    conn.close()
    return jsonify(orders)


@app.route('/api/admin/shipping/issue-tracking', methods=['POST'])
def admin_issue_tracking():
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    data = request.get_json() or {}
    company_id = data.get('company_id')
    order_ids = data.get('order_ids') or []
    if not str(company_id or '').isdigit():
        return jsonify({'error': '거래처를 먼저 선택하세요.'}), 400
    company_id = int(company_id)
    if not isinstance(order_ids, list) or not order_ids:
        return jsonify({'error': '송장을 발급할 주문을 선택하세요.'}), 400
    try:
        order_ids = list(dict.fromkeys(int(order_id) for order_id in order_ids))
    except (TypeError, ValueError):
        return jsonify({'error': '주문 선택 정보가 올바르지 않습니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id FROM companies WHERE id = ? AND active = 1', (company_id,))
    company = cur.fetchone()
    if company:
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT COUNT(*) AS order_count FROM orders WHERE tenant_id = ? AND id IN ({placeholders})', [company_id] + order_ids)
        owned_count = cur.fetchone()['order_count']
    else:
        owned_count = 0
    conn.close()
    if not company:
        return jsonify({'error': '활성 상태의 거래처를 찾을 수 없습니다.'}), 404
    if owned_count != len(order_ids):
        return jsonify({'error': '선택한 주문 중 해당 거래처에 속하지 않는 주문이 있습니다. 목록을 새로고침해주세요.'}), 400

    result = issue_tracking_numbers(order_ids, tenant_id=company_id)
    updated = result.get('updated_orders', []) if isinstance(result, dict) else result
    failed = result.get('failed_orders', []) if isinstance(result, dict) else []
    return jsonify({
        'status': 'success' if updated else 'error',
        'updated': updated,
        'failed': failed,
        'success_count': len(updated),
        'failed_count': len(failed),
        'message': f"송장 발급 {len(updated)}건 성공, {len(failed)}건 실패"
    })


@app.route('/api/admin/orders/<int:order_id>/edit', methods=['POST'])
def admin_edit_order(order_id):
    if session.get('account_role') != 'admin':
        return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
    data = request.get_json() or {}
    company_id = data.get('company_id')
    if not str(company_id or '').isdigit():
        return jsonify({'error': '거래처 정보가 올바르지 않습니다.'}), 400
    company_id = int(company_id)
    order_no = str(data.get('order_no') or '').strip()
    recipient_name = str(data.get('recipient_name') or '').strip()
    recipient_phone1 = str(data.get('recipient_phone1') or '').strip()
    recipient_phone2 = str(data.get('recipient_phone2') or '').strip()
    customs_id = str(data.get('customs_id') or '').strip().upper()
    postal_code = str(data.get('postal_code') or '').strip()
    address = str(data.get('address') or '').strip()
    shipping_memo = str(data.get('shipping_memo') or '').strip()
    try:
        weight = max(0.0, float(data.get('weight') or 0))
    except (TypeError, ValueError):
        return jsonify({'error': '무게는 0 이상의 숫자로 입력해주세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT id FROM companies WHERE id = ?', (company_id,))
        if not cur.fetchone():
            return jsonify({'error': '거래처를 찾을 수 없습니다.'}), 404
        cur.execute('SELECT id FROM orders WHERE id = ? AND tenant_id = ?', (order_id, company_id))
        if not cur.fetchone():
            return jsonify({'error': '해당 거래처의 주문을 찾을 수 없습니다.'}), 404
        cur.execute('''UPDATE orders
            SET order_no = ?, recipient_name = ?, recipient_phone1 = ?, recipient_phone2 = ?, customs_id = ?,
                postal_code = ?, address = ?, shipping_memo = ?, weight = ?
            WHERE id = ? AND tenant_id = ?''',
            (order_no, recipient_name, recipient_phone1, recipient_phone2, customs_id, postal_code,
             address, shipping_memo, weight, order_id, company_id))
        conn.commit()
        return jsonify({'status': 'success', 'message': '주문 정보가 수정되었습니다.'})
    except Exception:
        conn.rollback()
        app.logger.exception('Admin order edit failed')
        return jsonify({'error': '주문 정보를 저장하지 못했습니다.'}), 500
    finally:
        conn.close()


@app.route('/api/company', methods=['GET'])
def get_current_company():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id, name, invite_code FROM companies WHERE id = ?', (session.get('company_id'),))
    company = cur.fetchone()
    conn.close()
    return jsonify(dict(company) if company else {})


def format_inventory_log_datetime(value):
    """Normalize SQLite strings and PostgreSQL datetime values for the logs screen/export."""
    if value is None:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d %H:%M')
    raw = str(value).strip()
    if not raw:
        return ''
    try:
        return datetime.fromisoformat(raw.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M')
    except (TypeError, ValueError):
        try:
            return parsedate_to_datetime(raw).strftime('%Y-%m-%d %H:%M')
        except (TypeError, ValueError, OverflowError):
            return raw[:16]

@app.route('/')
def index():
    return render_template('index.html')

# ----------------- 1. 통계 요약 대시보드 -----------------
@app.route('/api/dashboard/stats', methods=['GET'])
def get_dashboard_stats():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        SELECT
            (SELECT COUNT(*) FROM orders WHERE tenant_id = ?) AS total_orders,
            (SELECT COUNT(*) FROM orders WHERE tenant_id = ? AND status IN ('주문접수', '출고대기')) AS pending_tracking,
            (SELECT COUNT(*) FROM orders WHERE tenant_id = ? AND status = '송장발급') AS issued_tracking,
            (SELECT COUNT(*) FROM orders WHERE tenant_id = ? AND status = '출고완료') AS completed_dispatch,
            (SELECT COUNT(*) FROM products WHERE tenant_id = ? AND stock_quantity <= safety_stock) AS low_stock_count,
            (SELECT COUNT(*) FROM order_items WHERE tenant_id = ?) AS total_items,
            (SELECT COUNT(*) FROM orders WHERE tenant_id = ? AND status = '통관오류') AS customs_error_count,
            (SELECT COUNT(*) FROM products WHERE tenant_id = ?) AS product_count
    ''', (g.company_id,) * 8)
    stats = dict(cur.fetchone())
    conn.close()
    return jsonify(stats)

# ----------------- 2. 엑셀 업로드 및 주문 분할 -----------------
def extract_sku_sort_key(sku_str):
    if not sku_str or sku_str == '미지정':
        return (2, 999999999999, '')
    digits = re.findall(r'\d+', str(sku_str))
    if digits:
        return (0, int(digits[0]), str(sku_str))
    return (1, 0, str(sku_str))

def get_order_sort_key(o):
    items = o.get('items') or o.get('split_items') or []
    item_cnt = len(items) if items else 1
    # 1. 단품 주문(품목 1개) 우선 (0), 그 밑에 2개부터 합포장 주문 (1)
    is_multi = 0 if item_cnt == 1 else 1
    # 2. 대표 바코드(SKU) 숫자/문자 오름차순
    primary_sku = items[0]['sku'] if items and items[0].get('sku') else ''
    sku_type, sku_num, sku_raw = extract_sku_sort_key(primary_sku)
    # 3. 출고수량 (옵션수량 * 구매수량) 오름차순
    if items:
        dispatch_qty = sum(
            (int(it.get('option_quantity') or 1)) * (int(it.get('product_quantity') or o.get('original_quantity') or 1))
            for it in items
        )
    else:
        dispatch_qty = int(o.get('original_quantity') or 1)
    return (
        is_multi,
        item_cnt,
        sku_type,
        sku_num,
        sku_raw,
        dispatch_qty,
        str(o.get('order_no') or ''),
        -o.get('id', 0)
    )


@app.route('/api/upload-profiles', methods=['GET'])
def get_upload_profiles():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT profile_name, mapping_json, header_row, sort_priority, sort_field, sort_direction, sort_rules_json FROM upload_profiles WHERE tenant_id = ? ORDER BY profile_name ASC', (g.company_id,))
    profiles = []
    for row in cur.fetchall():
        profile = dict(row)
        try:
            profile['mapping'] = json.loads(profile.pop('mapping_json') or '{}')
        except (TypeError, ValueError):
            profile['mapping'] = {}
        try:
            profile['sort_rules'] = json.loads(profile.pop('sort_rules_json') or '[]')
            if not isinstance(profile['sort_rules'], list):
                profile['sort_rules'] = []
            profile['sort_rules'] = [
                ({'field': 'order_no', 'direction': 'count_asc'} if rule.get('field') == 'order_item_count' else rule)
                for rule in profile['sort_rules'] if isinstance(rule, dict)
            ]
        except (TypeError, ValueError):
            profile['sort_rules'] = []
        profiles.append(profile)
    conn.close()
    return jsonify(profiles)


@app.route('/api/upload-profiles', methods=['POST'])
def save_upload_profile():
    data = request.get_json() or {}
    profile_name = str(data.get('profile_name') or '').strip()
    mapping = data.get('mapping') or {}
    sort_field = str(data.get('sort_field') or 'upload_order').strip()
    sort_direction = str(data.get('sort_direction') or 'asc').strip().lower()
    try:
        header_row = int(data.get('header_row', 1))
        sort_priority = int(data.get('sort_priority', 100))
    except (TypeError, ValueError):
        return jsonify({'error': '헤더 행과 정렬 우선순위는 숫자로 입력해주세요.'}), 400
    if not profile_name:
        return jsonify({'error': '업로드 형식 이름을 입력해주세요.'}), 400
    if len(profile_name) > 80:
        return jsonify({'error': '업로드 형식 이름은 80자 이내로 입력해주세요.'}), 400
    if not isinstance(mapping, dict):
        return jsonify({'error': '열 연결 설정 형식이 올바르지 않습니다.'}), 400
    if header_row < 0 or header_row > 100:
        return jsonify({'error': '헤더 행은 0부터 100 사이로 입력해주세요.'}), 400
    if sort_priority < 1 or sort_priority > 9999:
        return jsonify({'error': '정렬 우선순위는 1부터 9999 사이로 입력해주세요.'}), 400
    allowed_sort_fields = {'upload_order', 'order_no', 'order_item_count', 'recipient', 'product', 'barcode', 'quantity'}
    if sort_field not in allowed_sort_fields or sort_direction not in {'asc', 'desc', 'count_asc', 'count_desc'}:
        return jsonify({'error': '정렬 기준 설정이 올바르지 않습니다.'}), 400

    allowed_fields = {
        'order_no', 'courier', 'tracking_no', 'shipping_method', 'product_name', 'sku',
        'option_name', 'quantity', 'recipient_name', 'recipient_phone1',
        'recipient_phone2', 'customs_id', 'postal_code', 'address', 'shipping_memo'
    }
    clean_mapping = {
        key: str(value).strip()[:120]
        for key, value in mapping.items()
        if key in allowed_fields and str(value or '').strip()
    }

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        INSERT INTO upload_profiles (tenant_id, profile_name, mapping_json, header_row, sort_priority, sort_field, sort_direction)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(tenant_id, profile_name) DO UPDATE SET
            mapping_json = excluded.mapping_json,
            header_row = excluded.header_row,
            sort_priority = excluded.sort_priority,
            sort_field = excluded.sort_field,
            sort_direction = excluded.sort_direction
        RETURNING profile_name
    ''', (g.company_id, profile_name, json.dumps(clean_mapping, ensure_ascii=False), header_row,
          sort_priority, sort_field, sort_direction))
    cur.fetchone()
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'profile_name': profile_name})


@app.route('/api/upload-profiles/sort', methods=['POST'])
def save_upload_profile_sort():
    data = request.get_json() or {}
    profile_name = str(data.get('profile_name') or '').strip()
    try:
        sort_priority = int(data.get('sort_priority', 100))
    except (TypeError, ValueError):
        return jsonify({'error': '표시 순서는 숫자로 입력해주세요.'}), 400
    allowed_sort_fields = {'upload_order', 'order_no', 'recipient', 'product', 'barcode', 'quantity'}
    allowed_sort_directions = {'asc', 'desc', 'count_asc', 'count_desc'}
    sort_rules = data.get('sort_rules')
    if sort_rules is None:
        sort_rules = [{'field': data.get('sort_field') or 'upload_order',
                       'direction': data.get('sort_direction') or 'asc'}]
    if not isinstance(sort_rules, list) or not 1 <= len(sort_rules) <= 10:
        return jsonify({'error': '정렬 기준은 1개 이상 10개 이하로 설정해주세요.'}), 400
    clean_sort_rules = []
    for rule in sort_rules:
        if not isinstance(rule, dict):
            return jsonify({'error': '정렬 기준 형식이 올바르지 않습니다.'}), 400
        field = str(rule.get('field') or '').strip()
        direction = str(rule.get('direction') or 'asc').strip().lower()
        if field == 'order_item_count':
            field, direction = 'order_no', 'count_asc'
        if field not in allowed_sort_fields or direction not in allowed_sort_directions:
            return jsonify({'error': '정렬 기준 또는 방향이 올바르지 않습니다.'}), 400
        clean_sort_rules.append({'field': field, 'direction': direction})
    sort_field = clean_sort_rules[0]['field']
    sort_direction = clean_sort_rules[0]['direction']
    if len(profile_name) > 80:
        return jsonify({'error': '업로드 형식 이름은 80자 이내여야 합니다.'}), 400
    if sort_priority < 1 or sort_priority > 9999:
        return jsonify({'error': '표시 순서는 1부터 9999 사이로 입력해주세요.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT profile_name FROM upload_profiles WHERE tenant_id = ? AND profile_name = ?',
                (g.company_id, profile_name))
    exists = cur.fetchone()
    if not exists and profile_name not in ('', NARDA_MANIFEST_PROFILE):
        conn.close()
        return jsonify({'error': '업로드 형식을 찾을 수 없습니다. 새로고침 후 다시 시도해주세요.'}), 404

    cur.execute('''
        INSERT INTO upload_profiles (tenant_id, profile_name, mapping_json, header_row, sort_priority, sort_field, sort_direction, sort_rules_json)
        VALUES (?, ?, '{}', 1, ?, ?, ?, ?)
        ON CONFLICT(tenant_id, profile_name) DO UPDATE SET
            sort_priority = excluded.sort_priority,
            sort_field = excluded.sort_field,
            sort_direction = excluded.sort_direction,
            sort_rules_json = excluded.sort_rules_json
        RETURNING profile_name
    ''', (g.company_id, profile_name, sort_priority, sort_field, sort_direction,
          json.dumps(clean_sort_rules, ensure_ascii=False)))
    cur.fetchone()
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'profile_name': profile_name, 'sort_priority': sort_priority,
                    'sort_field': sort_field, 'sort_direction': sort_direction, 'sort_rules': clean_sort_rules})


@app.route('/api/upload-profiles/delete', methods=['POST'])
def delete_upload_profile():
    data = request.get_json() or {}
    profile_name = str(data.get('profile_name') or '').strip()
    if not profile_name:
        return jsonify({'error': '삭제할 업로드 형식을 선택해주세요.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM upload_profiles WHERE profile_name = ? AND tenant_id = ?', (profile_name, g.company_id))
    deleted = cur.rowcount > 0
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'deleted': deleted})


@app.route('/api/upload/template', methods=['GET'])
def download_order_upload_template():
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    headers = [
        '주문번호', '택배사', '송장번호', '배송방법', '상품명', '옵션명', '바코드', '상품수량',
        '수취인', '수취인전화번호1', '수취인전화번호2', '개인통관고유부호',
        '우편번호', '주소', '배송요청사항'
    ]
    descriptions = [
        ('주문번호', '주문을 구분하는 번호입니다. 비워두면 사이트에서 임시 번호를 만듭니다.'),
        ('택배사', '택배사명입니다. 예: 나르다'),
        ('송장번호', '기존 송장번호가 있는 경우 입력합니다. 신규 송장 발급 예정이면 비워둡니다.'),
        ('배송방법', '배송 방식입니다. 예: 택배'),
        ('상품명', '주문 상품명입니다. 재고 상품명과 맞춰 입력해주세요.'),
        ('옵션명', '상품 옵션명입니다. 옵션이 없으면 비워도 됩니다.'),
        ('바코드', '재고관리의 바코드(SKU)와 정확히 같아야 합니다. 입력하면 상품명/옵션명 대신 바코드로 재고 상품을 찾습니다.'),
        ('상품수량', '주문한 상품 수량입니다.'),
        ('수취인', '받는 분 성명입니다.'),
        ('수취인전화번호1', '받는 분 연락처입니다.'),
        ('수취인전화번호2', '보조 연락처입니다. 없으면 비워도 됩니다.'),
        ('개인통관고유부호', '해외 배송에 필요한 개인통관고유부호입니다.'),
        ('우편번호', '배송지 우편번호입니다.'),
        ('주소', '배송지 전체 주소입니다.'),
        ('배송요청사항', '배송 요청 또는 문 앞 요청사항입니다.')
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = '주문업로드양식'
    ws.append(headers)
    header_fill = PatternFill(fill_type='solid', fgColor='24476B')
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(name='맑은 고딕', bold=True, color='FFFFFF')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{ws.cell(row=1, column=len(headers)).column_letter}1'
    ws.row_dimensions[1].height = 32
    for index, header in enumerate(headers, start=1):
        ws.column_dimensions[ws.cell(row=1, column=index).column_letter].width = max(16, min(26, len(header) + 6))

    guide = wb.create_sheet('작성안내')
    guide.append(['엑셀 열 제목', '작성 안내'])
    for row in descriptions:
        guide.append(list(row))
    for cell in guide[1]:
        cell.fill = header_fill
        cell.font = Font(name='맑은 고딕', bold=True, color='FFFFFF')
    for row in guide.iter_rows(min_row=1):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    guide.column_dimensions['A'].width = 28
    guide.column_dimensions['B'].width = 85
    guide.freeze_panes = 'A2'

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name='주문업로드_자동열인식_기본양식.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/api/upload/narda-basic-template', methods=['GET'])
def download_narda_basic_template():
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data', '기본양식.xlsx')
    if not os.path.isfile(template_path):
        return jsonify({'error': '기본 양식 파일을 찾을 수 없습니다.'}), 404
    return send_file(
        template_path,
        as_attachment=True,
        download_name='기본양식.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/api/upload/preview', methods=['POST'])
def upload_preview():
    # 단일 또는 복수 파일 지원 ('files' 또는 'file')
    uploaded_files = request.files.getlist('files')
    if not uploaded_files:
        uploaded_files = request.files.getlist('file')
    if not uploaded_files and 'file' in request.files:
        uploaded_files = [request.files['file']]

    # 빈 파일 제외
    valid_files = [f for f in uploaded_files if f.filename != '']
    if not valid_files:
        return jsonify({'error': '업로드할 파일이 없습니다.'}), 400

    merged_orders = []
    total_excel_rows = 0
    total_split_items = 0
    batch_names = []
    detected_cols = {}
    all_unregistered = {}
    profile_name = str(request.form.get('upload_profile') or '').strip()
    profile_mapping = {}
    profile_header_row = 1

    if profile_name and profile_name != NARDA_MANIFEST_PROFILE:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute('SELECT mapping_json, header_row FROM upload_profiles WHERE profile_name = ? AND tenant_id = ?', (profile_name, g.company_id))
        profile_row = cur.fetchone()
        conn.close()
        if not profile_row:
            return jsonify({'error': '선택한 업로드 형식을 찾을 수 없습니다. 형식 목록을 새로고침한 뒤 다시 선택해주세요.'}), 400
        try:
            profile_mapping = json.loads(profile_row['mapping_json'] or '{}')
        except (TypeError, ValueError):
            profile_mapping = {}
        profile_header_row = profile_row['header_row'] if profile_row['header_row'] is not None else 1

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
                if profile_name == NARDA_MANIFEST_PROFILE:
                    parsed = process_narda_manifest_file(filepath, batch_name=orig_name, tenant_id=g.company_id)
                else:
                    parsed = process_excel_file(
                        filepath,
                        batch_name=orig_name,
                        tenant_id=g.company_id,
                        column_mapping=profile_mapping,
                        header_row=profile_header_row
                    )
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
            'upload_profile': profile_name,
            'manifest_mode': profile_name == NARDA_MANIFEST_PROFILE,
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
        result = save_orders_to_db(data, batch_name, tenant_id=g.company_id)
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
    dispatch_filter = request.args.get('dispatch_filter', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to = request.args.get('date_to', '').strip()
    sort_by = request.args.get('sort_by', 'single_barcode_asc').strip()
    limit_param = request.args.get('limit', '').strip()

    conn = get_connection()
    cur = conn.cursor()

    query = 'SELECT * FROM orders WHERE tenant_id = ?'
    params = [g.company_id]

    # 등록일(created_at) 기준 기간 검색. ISO 날짜만 받아 SQL 조건에 사용합니다.
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_from):
        query += " AND SUBSTR(CAST(created_at AS TEXT), 1, 10) >= ?"
        params.append(date_from)
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_to):
        query += " AND SUBSTR(CAST(created_at AS TEXT), 1, 10) <= ?"
        params.append(date_to)

    if status == '출고대기':
        query += " AND status IN ('주문접수', '출고대기')"
    elif status == '주문접수':
        query += " AND status = '주문접수'"
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
                OR id IN (SELECT order_id FROM order_items WHERE tenant_id = orders.tenant_id AND (item_product_name LIKE ? OR item_option_name LIKE ?))
            )'''
            params.extend([wild, wild, wild, wild])
        elif search_field == 'sku':
            query += ' AND (id IN (SELECT order_id FROM order_items WHERE tenant_id = orders.tenant_id AND sku LIKE ?))'
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
                OR id IN (SELECT order_id FROM order_items WHERE tenant_id = orders.tenant_id AND (item_product_name LIKE ? OR item_option_name LIKE ? OR sku LIKE ?))
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
        cur.execute(f'SELECT * FROM order_items WHERE tenant_id = ? AND order_id IN ({placeholders}) ORDER BY id ASC', [g.company_id] + order_ids)
        items_by_order = {}
        for it in cur.fetchall():
            oid = it['order_id']
            if oid not in items_by_order:
                items_by_order[oid] = []
            items_by_order[oid].append(dict(it))
        for o in orders:
            o['items'] = items_by_order.get(o['id'], [])

        # 기존 주문 중 분할 품목명에 원본의 합쳐진 문자열이 저장된 경우에도,
        # 바코드가 재고 옵션에 연결되어 있으면 품목별 재고 상품명/옵션명을 화면에 제공합니다.
        item_skus = list(dict.fromkeys(
            str(item.get('sku') or '').strip()
            for order in orders for item in order.get('items', [])
            if str(item.get('sku') or '').strip()
        ))
        if item_skus:
            sku_placeholders = ','.join(['?'] * len(item_skus))
            display_names_by_sku = {}
            cur.execute(f'''SELECT po.sku, p.product_name, po.option_name
                FROM product_options po
                LEFT JOIN products p ON p.id = po.product_id AND p.tenant_id = po.tenant_id
                WHERE po.tenant_id = ? AND po.sku IN ({sku_placeholders})
                ORDER BY po.id ASC''', [g.company_id] + item_skus)
            for row in cur.fetchall():
                display_names_by_sku.setdefault(str(row['sku'] or '').strip(), {
                    'display_product_name': row['product_name'] or '',
                    'display_option_name': row['option_name'] or ''
                })
            cur.execute(f'''SELECT sku, product_name, option_name FROM products
                WHERE tenant_id = ? AND sku IN ({sku_placeholders})''', [g.company_id] + item_skus)
            for row in cur.fetchall():
                display_names_by_sku.setdefault(str(row['sku'] or '').strip(), {
                    'display_product_name': row['product_name'] or '',
                    'display_option_name': row['option_name'] or ''
                })
            for order in orders:
                if int(order.get('reship_clear_items') or 0):
                    continue
                for item in order.get('items', []):
                    display_names = display_names_by_sku.get(str(item.get('sku') or '').strip())
                    if display_names:
                        item.update(display_names)
    else:
        for o in orders:
            o['items'] = []

    # 정렬: 1) 주문번호 한개씩(품목 1개 단품) 우선 -> 2) 2개부터 합포장 정렬 -> 3) 바코드 오름차순 -> 4) 출고수량 오름차순
    if sort_by == 'single_barcode_asc':
        orders.sort(key=get_order_sort_key)

    # 재고 필터는 같은 요청에서 계산해 브라우저와 DB 왕복을 줄입니다.
    if dispatch_filter in ('출고가능', '재고부족') and orders:
        cur.execute("SELECT sku, stock_quantity FROM products WHERE tenant_id = ? AND sku IS NOT NULL AND sku != ''", (g.company_id,))
        stock_map = {row['sku']: (row['stock_quantity'] or 0) for row in cur.fetchall()}

        for order in orders:
            all_available = True
            any_available = False
            item_details = []
            items = order.get('items') or []

            for item in items:
                sku = item.get('sku')
                option_qty = int(item.get('option_quantity') or 1)
                product_qty = int(item.get('product_quantity') or order.get('original_quantity') or 1)
                needed = option_qty * product_qty
                current_stock = stock_map.get(sku, 0)
                if current_stock >= needed:
                    stock_map[sku] = current_stock - needed
                    any_available = True
                    item_details.append({'sku': sku, 'needed': needed, 'available': True, 'remaining': stock_map[sku]})
                else:
                    all_available = False
                    item_details.append({'sku': sku, 'needed': needed, 'available': False, 'remaining': current_stock})

            if not items:
                availability_status = '출고불가'
            elif all_available:
                availability_status = '출고가능'
            elif any_available:
                availability_status = '부분부족'
            else:
                availability_status = '재고부족'

            order['_dispatchStatus'] = availability_status
            order['_dispatchDetails'] = item_details

        if dispatch_filter == '출고가능':
            orders = [order for order in orders if order['_dispatchStatus'] == '출고가능']
        else:
            orders = [order for order in orders if order['_dispatchStatus'] != '출고가능']

    conn.close()

    return jsonify(orders)

@app.route('/api/picking-list', methods=['GET'])
def get_picking_list():
    """Aggregate pending order items by barcode for warehouse picking."""
    conn = get_connection()
    cur = conn.cursor()
    try:
        def combined_item_name(product, option):
            product = re.sub(r'\s*\(출고\s*품목\)\s*', ' ', str(product or '')).strip()
            option = re.sub(r'\s*\(출고\s*품목\)\s*', ' ', str(option or '')).strip()
            if option in ('단일옵션', 'nan', 'None'):
                option = ''
            if not option or option == product:
                return product or option
            if product and product in option:
                return option
            if option and option in product:
                return product
            return ' '.join(part for part in (product, option) if part).strip()

        cur.execute('''
            SELECT o.id AS order_id, o.original_quantity, oi.product_quantity, oi.sku,
                   oi.item_product_name, oi.item_option_name,
                   oi.option_quantity, oi.quantity
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            WHERE o.tenant_id = ? AND o.status IN ('주문접수', '출고대기')
            ORDER BY oi.sku ASC, oi.item_product_name ASC, oi.item_option_name ASC, o.id ASC
        ''', (g.company_id,))

        per_order_items = {}
        for row in cur.fetchall():
            item = dict(row)
            sku = str(item.get('sku') or '').strip()
            product = str(item.get('item_product_name') or '').strip()
            option = str(item.get('item_option_name') or '').strip()
            combined_name = combined_item_name(product, option)

            order_id = item.get('order_id')
            try:
                product_quantity = max(1, int(float(item.get('product_quantity') or item.get('original_quantity') or 1)))
            except (TypeError, ValueError):
                product_quantity = 1
            try:
                option_quantity = max(1, int(float(item.get('option_quantity') or 1)))
            except (TypeError, ValueError):
                option_quantity = 1
            # option_quantity가 나중에 동기화된 과거 주문은 quantity 값이 예전 수량일 수 있어
            # 주문의 상품수량 × 현재 옵션의 실제 출고수량으로 피킹 수량을 계산합니다.
            purchase_quantity = product_quantity * option_quantity
            # 먼저 주문 안에서 같은 바코드·상품 옵션의 분할 행을 합칩니다.
            barcode_key = ('sku', sku) if sku else ('name', combined_name)
            order_item_key = (barcode_key, order_id, combined_name or '(상품명 없음)')
            per_order_items[order_item_key] = per_order_items.get(order_item_key, 0) + purchase_quantity

        grouped = {}
        for (barcode_key, order_id, combined_name), purchase_quantity in per_order_items.items():
            sku = barcode_key[1] if barcode_key[0] == 'sku' else ''
            entry = grouped.setdefault(barcode_key, {
                'sku': sku,
                'order_ids': set(),
                'buckets': {}
            })
            entry['order_ids'].add(order_id)
            bucket = entry['buckets'].setdefault(str(purchase_quantity), {
                'purchase_quantity': purchase_quantity,
                'descriptions': {}
            })
            description = bucket['descriptions'].setdefault(combined_name, {
                'orders': set(),
                'total_quantity': 0
            })
            description['orders'].add(order_id)
            description['total_quantity'] += purchase_quantity

        groups = []
        for entry in grouped.values():
            buckets = [
                {
                    'purchase_quantity': bucket['purchase_quantity'],
                    'product_name': name,
                    'order_count': len(description['orders']),
                    'total_quantity': description['total_quantity']
                }
                for bucket in sorted(entry['buckets'].values(), key=lambda b: b['purchase_quantity'])
                for name, description in sorted(bucket['descriptions'].items())
            ]
            groups.append({
                'sku': entry['sku'],
                'order_count': len(entry['order_ids']),
                'total_quantity': sum(bucket['total_quantity'] for bucket in buckets),
                'buckets': buckets
            })
        groups.sort(key=lambda group: (not bool(group['sku']), group['sku'].casefold()))
        row_count = sum(len(group['buckets']) for group in groups)
        return jsonify({
            'groups': groups,
            'sku_count': len(grouped),
            'row_count': row_count,
            'order_count': len({order_id for entry in grouped.values() for order_id in entry['order_ids']}),
            'total_quantity': sum(group['total_quantity'] for group in groups)
        })
    finally:
        conn.close()

@app.route('/api/dispatched-orders', methods=['GET'])
def get_dispatched_orders():
    courier = request.args.get('courier', '').strip()
    search = request.args.get('search', '').strip()
    selected_date = request.args.get('date', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to = request.args.get('date_to', '').strip()
    limit_param = request.args.get('limit', '').strip()

    conn = get_connection()
    cur = conn.cursor()

    query = """SELECT * FROM orders
        WHERE tenant_id = ? AND (status IN ('출고완료', '송장발급')
           OR (tracking_no IS NOT NULL AND TRIM(tracking_no) != ''))"""
    params = [g.company_id]

    if courier and courier != '전체':
        query += ' AND courier = ?'
        params.append(courier)

    if search:
        query += ''' AND (
            order_no LIKE ? OR recipient_name LIKE ? OR recipient_phone1 LIKE ? 
            OR tracking_no LIKE ? OR customs_id LIKE ? OR address LIKE ?
            OR original_product_name LIKE ? OR original_option_name LIKE ?
            OR id IN (SELECT order_id FROM order_items WHERE tenant_id = orders.tenant_id AND (item_product_name LIKE ? OR item_option_name LIKE ? OR sku LIKE ?))
        )'''
        wild = f'%{search}%'
        params.extend([wild, wild, wild, wild, wild, wild, wild, wild, wild, wild, wild])

    dispatch_date_expr = "COALESCE(NULLIF(SUBSTR(CAST(dispatched_at AS TEXT), 1, 10), ''), SUBSTR(CAST(created_at AS TEXT), 1, 10))"
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', selected_date):
        query += f' AND {dispatch_date_expr} = ?'
        params.append(selected_date)
    else:
        if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_from):
            query += f' AND {dispatch_date_expr} >= ?'
            params.append(date_from)
        if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_to):
            query += f' AND {dispatch_date_expr} <= ?'
            params.append(date_to)

    # PostgreSQL에서는 dispatched_at(TEXT)와 created_at(TIMESTAMP)의
    # COALESCE가 타입 오류를 내므로 양쪽을 TEXT로 맞춰 정렬합니다.
    query += " ORDER BY COALESCE(NULLIF(CAST(dispatched_at AS TEXT), ''), CAST(created_at AS TEXT)) DESC, id DESC"
    if limit_param and limit_param.isdigit():
        query += ' LIMIT ?'
        params.append(int(limit_param))

    cur.execute(query, params)
    orders = [dict(row) for row in cur.fetchall()]

    if orders:
        order_ids = [o['id'] for o in orders]
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'SELECT * FROM order_items WHERE tenant_id = ? AND order_id IN ({placeholders}) ORDER BY id ASC', [g.company_id] + order_ids)
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

@app.route('/api/dispatched-orders/reship', methods=['POST'])
def reship_dispatched_orders():
    data = request.get_json(silent=True) or {}
    raw_ids = data.get('order_ids') or []
    clear_items = bool(data.get('clear_items'))
    try:
        order_ids = list(dict.fromkeys(int(value) for value in raw_ids if str(value).isdigit()))
    except (TypeError, ValueError):
        order_ids = []
    if not order_ids:
        return jsonify({'error': '재발송할 발송 건을 선택하세요.'}), 400
    if len(order_ids) > 500:
        return jsonify({'error': '한 번에 재발송할 수 있는 건은 최대 500건입니다.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'''SELECT * FROM orders
            WHERE tenant_id = ? AND id IN ({placeholders})
              AND (status IN ('출고완료', '송장발급')
                   OR (tracking_no IS NOT NULL AND TRIM(tracking_no) != ''))''',
            [g.company_id] + order_ids)
        selected = [dict(row) for row in cur.fetchall()]
        if len(selected) != len(order_ids):
            return jsonify({'error': '선택한 발송 건 중 일부를 찾지 못했습니다. 목록을 새로고침한 뒤 다시 선택하세요.'}), 409

        updated = []
        for order in selected:
            old_no = str(order.get('order_no') or '').strip()
            match = re.match(r'^(.*?)-(\d+)$', old_no)
            base_no = match.group(1) if match else old_no
            next_suffix = int(match.group(2)) + 1 if match else 1
            new_no = f'{base_no}-{next_suffix}'
            while True:
                cur.execute('''SELECT id FROM orders
                    WHERE tenant_id = ? AND order_no = ? AND id != ? LIMIT 1''',
                    (g.company_id, new_no, order['id']))
                if not cur.fetchone():
                    break
                next_suffix += 1
                new_no = f'{base_no}-{next_suffix}'

            original_manifest = order.get('manifest_data') or ''
            if clear_items:
                try:
                    manifest = json.loads(original_manifest) if original_manifest else {}
                    if not isinstance(manifest, dict):
                        manifest = {}
                except (TypeError, ValueError):
                    manifest = {}
                for key in ('Item Name', '아이템상세정보', 'Unit Price', 'Piece', 'Amount', 'HS Code 1'):
                    manifest[key] = ''
                original_manifest = json.dumps(manifest, ensure_ascii=False)

            # Preserve the completed shipment as history. Re-shipping creates
            # a fresh order with no inherited NARDA tracking number or PDF URL.
            cur.execute('''INSERT INTO orders (
                tenant_id, order_no, courier, tracking_no, shipping_method,
                original_product_name, original_option_name, original_quantity,
                recipient_name, recipient_phone1, recipient_phone2, customs_id,
                postal_code, address, shipping_memo, status, upload_batch,
                dispatched_at, weight, customs_validation_status, customs_validation_msg,
                api_dispatch_status, api_dispatch_message, manifest_data, upload_profile,
                packaging_type, reship_clear_items
            ) VALUES (?, ?, ?, '', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '주문접수', ?, '', ?, '', '', '', '', ?, ?, ?, ?)''',
                (g.company_id, new_no, order.get('courier') or '나르다', order.get('shipping_method') or '택배',
                 '' if clear_items else order.get('original_product_name'),
                 '' if clear_items else order.get('original_option_name'), order.get('original_quantity') or 1,
                 order.get('recipient_name'), order.get('recipient_phone1'), order.get('recipient_phone2'),
                 order.get('customs_id'), order.get('postal_code'), order.get('address'), order.get('shipping_memo'),
                 order.get('upload_batch'), order.get('weight') or 0, original_manifest,
                 order.get('upload_profile') or '', order.get('packaging_type') or '기본포장',
                 1 if clear_items else 0))
            new_order_id = cur.lastrowid

            cur.execute('SELECT * FROM order_items WHERE tenant_id = ? AND order_id = ? ORDER BY id ASC',
                        (g.company_id, order['id']))
            source_items = [dict(row) for row in cur.fetchall()]
            for item in source_items:
                item_manifest = item.get('manifest_data') or ''
                if clear_items:
                    try:
                        item_manifest_obj = json.loads(item_manifest) if item_manifest else {}
                        if not isinstance(item_manifest_obj, dict):
                            item_manifest_obj = {}
                    except (TypeError, ValueError):
                        item_manifest_obj = {}
                    for key in ('Item Name', '아이템상세정보', 'Unit Price', 'Piece', 'Amount', 'HS Code 1'):
                        item_manifest_obj[key] = ''
                    item_manifest = json.dumps(item_manifest_obj, ensure_ascii=False)
                cur.execute('''INSERT INTO order_items (
                    tenant_id, order_id, order_no, sku, item_product_name, item_option_name,
                    product_quantity, option_quantity, quantity, stock_deducted, status, manifest_data
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, '준비중', ?)''',
                    (g.company_id, new_order_id, new_no, item.get('sku'), item.get('item_product_name') or '',
                     item.get('item_option_name') or '', item.get('product_quantity') or 1,
                     item.get('option_quantity') or 1, item.get('quantity') or 1, item_manifest))
            updated.append({'id': new_order_id, 'source_id': order['id'], 'old_order_no': old_no, 'order_no': new_no})

        conn.commit()
        return jsonify({'status': 'success', 'updated_count': len(updated), 'items_cleared': clear_items,
                        'history_preserved': True, 'tracking_cleared': True, 'orders': updated})
    except Exception as e:
        conn.rollback()
        app.logger.exception('Reship failed')
        return jsonify({'error': f'재발송 처리 중 오류가 발생했습니다: {str(e)}'}), 500
    finally:
        conn.close()

@app.route('/api/dispatched-orders/dates', methods=['GET'])
def get_dispatched_order_dates():
    courier = request.args.get('courier', '').strip()
    search = request.args.get('search', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to = request.args.get('date_to', '').strip()
    conn = get_connection()
    cur = conn.cursor()
    query = """SELECT COALESCE(NULLIF(SUBSTR(CAST(dispatched_at AS TEXT), 1, 10), ''),
                    SUBSTR(CAST(created_at AS TEXT), 1, 10)) AS date, COUNT(*) AS count
               FROM orders
               WHERE tenant_id = ? AND (status IN ('출고완료', '송장발급')
                  OR (tracking_no IS NOT NULL AND TRIM(tracking_no) != ''))"""
    params = [g.company_id]
    if courier and courier != '전체':
        query += ' AND courier = ?'
        params.append(courier)
    if search:
        query += ''' AND (
            order_no LIKE ? OR recipient_name LIKE ? OR recipient_phone1 LIKE ?
            OR tracking_no LIKE ? OR customs_id LIKE ? OR address LIKE ?
            OR original_product_name LIKE ? OR original_option_name LIKE ?
            OR id IN (SELECT order_id FROM order_items WHERE tenant_id = orders.tenant_id AND (item_product_name LIKE ? OR item_option_name LIKE ? OR sku LIKE ?))
        )'''
        wild = f'%{search}%'
        params.extend([wild] * 11)
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_from):
        query += " AND COALESCE(NULLIF(SUBSTR(CAST(dispatched_at AS TEXT), 1, 10), ''), SUBSTR(CAST(created_at AS TEXT), 1, 10)) >= ?"
        params.append(date_from)
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', date_to):
        query += " AND COALESCE(NULLIF(SUBSTR(CAST(dispatched_at AS TEXT), 1, 10), ''), SUBSTR(CAST(created_at AS TEXT), 1, 10)) <= ?"
        params.append(date_to)
    query += ' GROUP BY date ORDER BY date DESC'
    cur.execute(query, params)
    dates = [dict(row) for row in cur.fetchall()]
    conn.close()
    return jsonify(dates)

@app.route('/api/orders/<int:order_id>/weight', methods=['POST'])
def update_order_weight(order_id):
    data = request.get_json() or {}
    try:
        weight = float(data.get('weight', 0.0) or 0.0)
    except:
        weight = 0.0

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('UPDATE orders SET weight = ? WHERE id = ? AND tenant_id = ?', (weight, order_id, g.company_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'weight': weight})


ORDER_PACKAGING_TYPES = {'기본포장', '추가포장', '박스포장', '추가+박스포장'}

@app.route('/api/orders/<int:order_id>/packaging', methods=['POST'])
def update_order_packaging(order_id):
    data = request.get_json(silent=True) or {}
    packaging_type = str(data.get('packaging_type') or '').strip()
    if packaging_type not in ORDER_PACKAGING_TYPES:
        return jsonify({'error': '포장 방식을 선택해주세요.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('UPDATE orders SET packaging_type = ? WHERE id = ? AND tenant_id = ?',
                    (packaging_type, order_id, g.company_id))
        if cur.rowcount == 0:
            return jsonify({'error': '주문을 찾을 수 없습니다.'}), 404
        conn.commit()
        return jsonify({'status': 'success', 'packaging_type': packaging_type})
    finally:
        conn.close()


@app.route('/api/orders/bulk-packaging', methods=['POST'])
def bulk_update_order_packaging():
    data = request.get_json(silent=True) or {}
    order_ids = data.get('order_ids') or []
    packaging_type = str(data.get('packaging_type') or '').strip()
    if packaging_type not in ORDER_PACKAGING_TYPES:
        return jsonify({'error': '포장 방식을 선택해주세요.'}), 400
    if not isinstance(order_ids, list) or not order_ids:
        return jsonify({'error': '포장 방식을 적용할 주문을 선택하세요.'}), 400
    ids = []
    try:
        ids = list(dict.fromkeys(int(order_id) for order_id in order_ids))
    except (TypeError, ValueError):
        return jsonify({'error': '주문 선택 정보가 올바르지 않습니다.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    try:
        marks = ','.join(['?'] * len(ids))
        cur.execute(f'UPDATE orders SET packaging_type = ? WHERE tenant_id = ? AND id IN ({marks})',
                    [packaging_type, g.company_id] + ids)
        updated_count = cur.rowcount
        conn.commit()
        return jsonify({'status': 'success', 'updated_count': updated_count, 'packaging_type': packaging_type})
    finally:
        conn.close()

@app.route('/api/orders/<int:order_id>/edit', methods=['POST'])
def edit_order(order_id):
    data = request.get_json() or {}
    recipient_name = str(data.get('recipient_name') or '').strip()
    recipient_phone1 = str(data.get('recipient_phone1') or '').strip()
    recipient_phone2 = str(data.get('recipient_phone2') or '').strip()
    customs_id = str(data.get('customs_id') or '').strip().upper()
    postal_code = str(data.get('postal_code') or '').strip()
    address = str(data.get('address') or '').strip()
    shipping_memo = str(data.get('shipping_memo') or '').strip()
    try:
        weight = max(0.0, float(data.get('weight', 0.0) or 0.0))
    except (TypeError, ValueError):
        return jsonify({'error': '무게는 0 이상의 숫자로 입력해주세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT * FROM orders WHERE id = ? AND tenant_id = ?', (order_id, g.company_id))
        order = cur.fetchone()
        if not order:
            return jsonify({'error': '주문을 찾을 수 없습니다.'}), 404

        manifest_items = data.get('manifest_items') or []
        if not isinstance(manifest_items, list):
            return jsonify({'error': '기본형식 품목 정보가 올바르지 않습니다.'}), 400
        order_manifest = order['manifest_data'] or '{}'
        try:
            order_manifest = json.loads(order_manifest) if isinstance(order_manifest, str) else dict(order_manifest)
        except (TypeError, ValueError):
            order_manifest = {}
        if not isinstance(order_manifest, dict):
            order_manifest = {}
        ctns = str(data.get('ctns') or '').strip()
        order_manifest['CTNS'] = ctns

        cur.execute('SELECT * FROM order_items WHERE tenant_id = ? AND order_id = ? ORDER BY id ASC',
                    (g.company_id, order_id))
        all_owned_items = [dict(row) for row in cur.fetchall()]
        manifest_owned = {}
        for row in all_owned_items:
            try:
                row_manifest = json.loads(row.get('manifest_data') or '{}')
            except (TypeError, ValueError):
                row_manifest = {}
            if isinstance(row_manifest, dict) and str(row_manifest.get('Item Name') or '').strip():
                manifest_owned[str(row['id'])] = (row, row_manifest)

        allowed_fields = {'HS Code 1', 'Item Name', '아이템상세정보', 'Unit Price', 'Piece', 'Amount', '상품 URL', '원산지'}
        item_updates = []
        item_inserts = []
        retained_ids = set()
        for submitted in manifest_items:
            if not isinstance(submitted, dict) or not isinstance(submitted.get('fields'), dict):
                return jsonify({'error': '수정할 상품 정보를 확인해주세요.'}), 400
            item_id = str(submitted.get('id') or '').strip()
            if item_id:
                if item_id not in manifest_owned:
                    return jsonify({'error': '수정할 상품을 확인할 수 없습니다. 새로고침 후 다시 시도해주세요.'}), 400
                row, fields = manifest_owned[item_id]
                fields = dict(fields)
            else:
                row = None
                fields = {}
            for field in allowed_fields:
                if field in submitted['fields']:
                    fields[field] = str(submitted['fields'][field] if submitted['fields'][field] is not None else '').strip()
            fields['CTNS'] = ctns
            if not str(fields.get('Item Name') or '').strip():
                return jsonify({'error': 'Item Name은 비워둘 수 없습니다. 상품 행을 삭제하거나 상품명을 입력해주세요.'}), 400
            try:
                piece_value = float(fields.get('Piece') or 0)
                if not math.isfinite(piece_value) or not piece_value.is_integer() or piece_value < 1:
                    raise ValueError
                piece = int(piece_value)
            except (TypeError, ValueError):
                return jsonify({'error': 'Piece는 1 이상의 정수로 입력해주세요.'}), 400
            for field in ('Unit Price', 'Amount'):
                value = str(fields.get(field) or '').replace(',', '').strip()
                if value:
                    try:
                        numeric_value = float(value)
                        if not math.isfinite(numeric_value) or numeric_value < 0:
                            raise ValueError
                    except ValueError:
                        return jsonify({'error': f'{field}는 0 이상의 숫자로 입력해주세요.'}), 400
                    fields[field] = value
            manifest_json = json.dumps(fields, ensure_ascii=False)
            entry = (manifest_json, fields['Item Name'], fields.get('아이템상세정보', ''), piece)
            if row:
                retained_ids.add(item_id)
                item_updates.append((*entry, item_id))
            else:
                item_inserts.append(entry)

        # 상품 목록에서 제거한 행은 기본형식 manifest 품목만 삭제합니다.
        deleted_ids = set(manifest_owned) - retained_ids
        for deleted_id in deleted_ids:
            cur.execute('DELETE FROM order_items WHERE id = ? AND order_id = ? AND tenant_id = ?',
                        (int(deleted_id), order_id, g.company_id))

        cur.execute('''UPDATE orders
            SET recipient_name = ?, recipient_phone1 = ?, recipient_phone2 = ?, customs_id = ?,
                postal_code = ?, address = ?, shipping_memo = ?, weight = ?, manifest_data = ?
            WHERE id = ? AND tenant_id = ?''',
            (recipient_name, recipient_phone1, recipient_phone2, customs_id, postal_code,
             address, shipping_memo, weight, json.dumps(order_manifest, ensure_ascii=False), order_id, g.company_id))
        for manifest_data, item_name, item_detail, piece, item_id in item_updates:
            cur.execute('''UPDATE order_items
                SET manifest_data = ?, item_product_name = ?, item_option_name = ?,
                    product_quantity = ?, quantity = ?
                WHERE id = ? AND order_id = ? AND tenant_id = ?''',
                (manifest_data, item_name, item_detail, piece, piece, int(item_id), order_id, g.company_id))
        for manifest_data, item_name, item_detail, piece in item_inserts:
            cur.execute('''INSERT INTO order_items (
                tenant_id, order_id, order_no, sku, item_product_name, item_option_name,
                product_quantity, option_quantity, quantity, stock_deducted, status, manifest_data
            ) VALUES (?, ?, ?, '', ?, ?, ?, 1, ?, 0, '준비중', ?)''',
                (g.company_id, order_id, order['order_no'], item_name, item_detail, piece, piece, manifest_data))
        if manifest_owned or item_inserts:
            cur.execute('''SELECT item_product_name, item_option_name, quantity FROM order_items
                WHERE order_id = ? AND tenant_id = ? ORDER BY id ASC''', (order_id, g.company_id))
            current_items = [dict(row) for row in cur.fetchall()]
            product_names = list(dict.fromkeys(str(row.get('item_product_name') or '').strip()
                                               for row in current_items
                                               if str(row.get('item_product_name') or '').strip()))
            option_names = list(dict.fromkeys(str(row.get('item_option_name') or '').strip()
                                              for row in current_items
                                              if str(row.get('item_option_name') or '').strip()))
            total_quantity = sum(int(row.get('quantity') or 0) for row in current_items)
            cur.execute('''UPDATE orders SET original_product_name = ?, original_option_name = ?, original_quantity = ?
                WHERE id = ? AND tenant_id = ?''',
                (' / '.join(product_names), ' / '.join(option_names), total_quantity, order_id, g.company_id))
        conn.commit()
        return jsonify({'status': 'success', 'message': '주문 및 기본형식 품목 정보가 수정되었습니다.'})
    except Exception:
        conn.rollback()
        app.logger.exception('Order edit failed')
        return jsonify({'error': '주문 정보를 저장하지 못했습니다. 다시 시도해주세요.'}), 500
    finally:
        conn.close()

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

    cur.execute('SELECT sku, stock_quantity FROM products WHERE tenant_id = ?', (g.company_id,))
    stock_map = {r['sku']: (r['stock_quantity'] or 0) for r in cur.fetchall() if r['sku']}

    if order_ids:
        placeholders = ','.join('?' for _ in order_ids)
        cur.execute(f"SELECT * FROM orders WHERE tenant_id = ? AND id IN ({placeholders})", [g.company_id] + order_ids)
        raw_orders = {r['id']: dict(r) for r in cur.fetchall()}
        orders = [raw_orders[oid] for oid in order_ids if oid in raw_orders]
    else:
        cur.execute("SELECT * FROM orders WHERE tenant_id = ? AND status IN ('주문접수', '출고대기', '송장발급') ORDER BY id DESC", (g.company_id,))
        orders = [dict(r) for r in cur.fetchall()]

    items_by_order = {}
    if orders:
        o_ids = [o['id'] for o in orders]
        placeholders = ','.join(['?'] * len(o_ids))
        cur.execute(f'SELECT * FROM order_items WHERE tenant_id = ? AND order_id IN ({placeholders}) ORDER BY id ASC', [g.company_id] + o_ids)
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
            opt_qty = int(it.get('option_quantity') or 1)
            prod_qty = int(it.get('product_quantity') or o.get('original_quantity') or 1)
            needed = opt_qty * prod_qty

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
                FROM orders WHERE tenant_id = ? AND id IN ({placeholders})
            ''', [g.company_id] + order_ids)
        else:
            cur.execute('''
                SELECT id, order_no, recipient_name, customs_id, recipient_phone1, postal_code
                FROM orders WHERE tenant_id = ? AND status IN ('주문접수', '출고대기', '송장발급')
            ''', (g.company_id,))

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
                if pcode.isdigit() and len(pcode) < 5:
                    pcode = pcode.zfill(5)

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
                                WHERE id = ? AND tenant_id = ?
                            ''', (res_status, res_msg, oid, g.company_id))
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
                            WHERE id = ? AND tenant_id = ?
                        ''', (fail_status, fail_msg, o['id'], g.company_id))
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
                        WHERE id = ? AND tenant_id = ?
                    ''', ('장애', f'통신 오류: {str(e)}', o['id'], g.company_id))
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
    cur.execute(f'DELETE FROM order_items WHERE tenant_id = ? AND order_id IN ({placeholders})', [g.company_id] + order_ids)
    cur.execute(f'DELETE FROM orders WHERE tenant_id = ? AND id IN ({placeholders})', [g.company_id] + order_ids)
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
    cur.execute(f'UPDATE orders SET weight = ? WHERE tenant_id = ? AND id IN ({placeholders})', [weight, g.company_id] + order_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'updated_count': len(order_ids), 'weight': weight})

# ----------------- 4. 배송 & 송장 & 출고 API -----------------
@app.route('/api/shipping/issue-tracking', methods=['POST'])
def api_issue_tracking():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    if not order_ids:
        return jsonify({'error': '송장을 발급할 주문을 선택하세요.'}), 400

    res = issue_tracking_numbers(order_ids, tenant_id=g.company_id)
    if isinstance(res, dict):
        updated = res.get('updated_orders', [])
        failed = res.get('failed_orders', [])
    else:
        updated = res
        failed = []

    return jsonify({
        'status': 'success' if updated else 'error',
        'updated': updated,
        'failed': failed,
        'success_count': len(updated),
        'failed_count': len(failed),
        'message': f"송장 발급 {len(updated)}건 성공, {len(failed)}건 실패"
    })

@app.route('/api/shipping/dispatch', methods=['POST'])
def api_dispatch():
    data = request.get_json() or {}
    order_ids = data.get('order_ids', [])
    if not order_ids:
        return jsonify({'error': '출고 완료할 주문을 선택하세요.'}), 400

    res = complete_dispatch(order_ids, tenant_id=g.company_id)
    return jsonify({
        'status': 'success',
        'message': f"{res['dispatched_orders']}건 발송 완료 처리 및 {res['deducted_items']}개 품목 재고가 자동 차감되었습니다.",
        'details': res
    })

@app.route('/api/shipping/export-excel', methods=['GET'])
def api_export_excel():
    order_ids_str = request.args.get('order_ids', '')
    order_ids = [int(x) for x in order_ids_str.split(',') if x.strip().isdigit()]
    column_order = [key.strip() for key in request.args.get('column_order', '').split(',') if key.strip()]
    excel_buf = export_shipping_excel(order_ids if order_ids else None, column_order=column_order or None, tenant_id=g.company_id)
    return send_file(
        excel_buf,
        as_attachment=True,
        download_name='발송처리결과_송장완료.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/api/shipping/batches', methods=['GET'])
def api_shipping_batches():
    """
    송장이 발급된 주문들을 날짜 및 발급 묶음별로 그룹핑하여 반환합니다.
    [송장 라벨 인쇄 센터] 좌측 목록 구성용
    """
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('''
            SELECT id, order_no, courier, tracking_no, invoice_pdf_url, recipient_name,
                   created_at, dispatched_at, status
            FROM orders
            WHERE tenant_id = ? AND LOWER(TRIM(COALESCE(courier, ''))) IN ('나르다', 'narda', 'narda air')
              AND invoice_pdf_url IS NOT NULL AND TRIM(invoice_pdf_url) != ''
            ORDER BY id DESC
        ''', (g.company_id,))
        orders = [dict(r) for r in cur.fetchall()]
        if not orders:
            return jsonify([])

        groups = {}
        for o in orders:
            dt_str = o.get('dispatched_at') or ''
            if dt_str and len(dt_str) >= 10:
                d_key = dt_str[:10]
            else:
                c_str = str(o.get('created_at') or '')
                d_key = c_str[:10] if len(c_str) >= 10 else datetime.now().strftime('%Y-%m-%d')

            courier = o.get('courier') or '나르다'
            group_key = f"{d_key}_{courier}"

            if group_key not in groups:
                groups[group_key] = {
                    'group_key': group_key,
                    'date': d_key,
                    'courier': courier,
                    'title': f"{d_key} {courier} 송장",
                    'order_ids': [],
                    'has_pdf_count': 0,
                    'orders_count': 0
                }
            groups[group_key]['order_ids'].append(o['id'])
            groups[group_key]['orders_count'] += 1
            groups[group_key]['has_pdf_count'] += 1

        batch_list = list(groups.values())
        return jsonify(batch_list)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/shipping/labels', methods=['GET'])
def api_shipping_labels():
    order_ids_str = request.args.get('order_ids', '')
    order_ids = [int(x) for x in order_ids_str.split(',') if x.strip().isdigit()]
    conn = get_connection()
    cur = conn.cursor()

    try:
        # 발송자 정보
        cur.execute('SELECT * FROM courier_settings WHERE tenant_id = ? ORDER BY id LIMIT 1', (g.company_id,))
        sender = cur.fetchone()
        sender_info = dict(sender) if sender else {
            'sender_name': 'TOA',
            'sender_phone': '02-1234-5678',
            'sender_address': '서울특별시 강남구 테헤란로 123 물류빌딩 3층'
        }

        if order_ids:
            placeholders = ','.join(['?'] * len(order_ids))
            cur.execute(f'''SELECT * FROM orders
                WHERE tenant_id = ? AND id IN ({placeholders})
                  AND LOWER(TRIM(COALESCE(courier, ''))) IN ('나르다', 'narda', 'narda air')
                  AND invoice_pdf_url IS NOT NULL AND TRIM(invoice_pdf_url) != ''
                ORDER BY id ASC''', [g.company_id] + order_ids)
        else:
            cur.execute('''
                SELECT * FROM orders
                WHERE tenant_id = ? AND LOWER(TRIM(COALESCE(courier, ''))) IN ('나르다', 'narda', 'narda air')
                  AND invoice_pdf_url IS NOT NULL AND TRIM(invoice_pdf_url) != ''
                ORDER BY id DESC LIMIT 150
            ''', (g.company_id,))

        orders = [dict(row) for row in cur.fetchall()]

        if orders:
            oids = [o['id'] for o in orders]
            placeholders = ','.join(['?'] * len(oids))
            cur.execute(f'SELECT * FROM order_items WHERE tenant_id = ? AND order_id IN ({placeholders}) ORDER BY id ASC', [g.company_id] + oids)
            items_by_order = {}
            for it in cur.fetchall():
                oid = it['order_id']
                if oid not in items_by_order:
                    items_by_order[oid] = []
                items_by_order[oid].append(dict(it))
            for o in orders:
                o['items'] = items_by_order.get(o['id'], [])
                o['sender'] = sender_info
        return jsonify(orders)
    finally:
        conn.close()

@app.route('/api/shipping/narda-pdf-bundle', methods=['GET'])
def api_narda_pdf_bundle():
    return _narda_pdf_bundle_for_tenant(g.company_id)


@app.route('/api/admin/shipping/narda-pdf-bundle', methods=['GET'])
def admin_narda_pdf_bundle():
    if session.get('account_role') != 'admin':
        return '관리자 권한이 필요합니다.', 403
    company_id = request.args.get('company_id', '').strip()
    if not company_id.isdigit():
        return '거래처를 선택해주세요.', 400
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT id FROM companies WHERE id = ? AND active = 1', (int(company_id),))
    company = cur.fetchone()
    conn.close()
    if not company:
        return '활성 상태의 거래처를 찾을 수 없습니다.', 404
    return _narda_pdf_bundle_for_tenant(int(company_id))


def _narda_pdf_bundle_for_tenant(tenant_id):
    order_ids_str = request.args.get('order_ids', '').strip()
    conn = get_connection()
    cur = conn.cursor()

    if order_ids_str:
        order_ids = [int(x) for x in order_ids_str.split(',') if x.strip().isdigit()]
        if not order_ids:
            conn.close()
            return "유효한 주문 ID가 없습니다.", 400
        placeholders = ','.join(['?'] * len(order_ids))
        cur.execute(f'''SELECT id, order_no, invoice_pdf_url FROM orders
            WHERE tenant_id = ? AND id IN ({placeholders})
              AND LOWER(TRIM(COALESCE(courier, ''))) IN ('나르다', 'narda', 'narda air')
              AND invoice_pdf_url IS NOT NULL AND TRIM(invoice_pdf_url) != '' ''', [tenant_id] + order_ids)
    else:
        # 지정되지 않은 경우 최근 발급된 나르다 PDF 송장 주문 최대 200건 조회
        cur.execute('''SELECT id, order_no, invoice_pdf_url FROM orders
            WHERE tenant_id = ? AND LOWER(TRIM(COALESCE(courier, ''))) IN ('나르다', 'narda', 'narda air')
              AND invoice_pdf_url IS NOT NULL AND TRIM(invoice_pdf_url) != ''
            ORDER BY id DESC LIMIT 200''', (tenant_id,))

    raw_rows = cur.fetchall()
    conn.close()

    rows = [dict(r) for r in raw_rows] if raw_rows else []
    if not rows:
        return "선택한 주문에 발급된 나르다 송장 PDF가 없습니다. 먼저 송장을 발급해주세요.", 404

    # 단 1건인 경우 해당 공식 PDF URL로 즉시 리다이렉트
    if len(rows) == 1:
        return redirect(rows[0]['invoice_pdf_url'])

    try:
        from pypdf import PdfReader, PdfWriter
        import io
        import concurrent.futures
        import threading
        from requests.adapters import HTTPAdapter
    except Exception as e:
        app.logger.exception('나르다 PDF 병합 라이브러리 로드 실패')
        return f'나르다 송장 병합 기능을 불러오지 못했습니다: {e}', 500

    pdf_session_local = threading.local()
    pdf_sessions = []
    pdf_sessions_lock = threading.Lock()

    def get_pdf_session():
        session = getattr(pdf_session_local, 'session', None)
        if session is None:
            session = requests.Session()
            adapter = HTTPAdapter(pool_connections=4, pool_maxsize=4, max_retries=0)
            session.mount('http://', adapter)
            session.mount('https://', adapter)
            pdf_session_local.session = session
            with pdf_sessions_lock:
                pdf_sessions.append(session)
        return session

    def fetch_single_pdf(item):
        url = (item.get('invoice_pdf_url') or '').strip()
        if not url or not url.startswith('http'):
            app.logger.warning('나르다 PDF URL이 비어 있거나 HTTP URL이 아닙니다 (order id=%s)', item.get('id'))
            return (item['id'], None)
        try:
            resp = get_pdf_session().get(
                url,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'},
                timeout=(8, 20)
            )
            if resp.status_code == 200 and resp.content.startswith(b'%PDF-'):
                return (item['id'], resp.content)
            app.logger.warning(
                '나르다 PDF 응답이 PDF 형식이 아닙니다 (order id=%s, HTTP %s, Content-Type=%s)',
                item.get('id'), resp.status_code, resp.headers.get('Content-Type', '')
            )
        except Exception as e:
            app.logger.warning('나르다 PDF 다운로드 실패 (order id=%s): %s', item.get('id'), e)
        return (item['id'], None)

    # 작업 스레드마다 HTTP 연결을 재사용하고, 공식 PDF를 최대 30개씩 병렬로 받습니다.
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(30, len(rows))) as executor:
            pdf_dict = dict(executor.map(fetch_single_pdf, rows))
    finally:
        for session in pdf_sessions:
            session.close()

    writer = PdfWriter()
    merged_count = 0
    failed_orders = []

    for r in rows:
        content = pdf_dict.get(r['id'])
        if not content:
            failed_orders.append(str(r.get('order_no') or r['id']))
            continue
        try:
            reader = PdfReader(io.BytesIO(content), strict=False)
            if reader.is_encrypted:
                try:
                    reader.decrypt('')
                except Exception:
                    pass
            if not reader.pages:
                raise ValueError('PDF 페이지가 없습니다')
            for page in reader.pages:
                writer.add_page(page)
            merged_count += 1
        except Exception as pe:
            failed_orders.append(str(r.get('order_no') or r['id']))
            app.logger.warning('나르다 송장 PDF 파싱 실패 (%s): %s', r.get('order_no'), pe)

    if merged_count == 0:
        return '선택한 나르다 공식 PDF를 가져오지 못했습니다. 송장 PDF URL 또는 서버의 외부 연결을 확인해주세요.', 502
    if failed_orders:
        failed_list = ', '.join(failed_orders[:20])
        return f'선택한 송장 중 {len(failed_orders)}건을 불러오지 못해 통합 출력을 중단했습니다. 실패 주문: {failed_list}', 502

    output_stream = io.BytesIO()
    try:
        writer.write(output_stream)
        output_stream.seek(0)
    except Exception as e:
        app.logger.exception('나르다 통합 PDF 생성 실패')
        return f'나르다 통합 PDF를 만들지 못했습니다: {e}', 500

    response = send_file(
        output_stream,
        mimetype='application/pdf',
        as_attachment=False,
        download_name=f'NARDA_통합송장_{merged_count}건.pdf',
        max_age=300
    )
    response.cache_control.private = True
    response.headers['X-Narda-Label-Count'] = str(merged_count)
    response.headers['X-Narda-Failed-Count'] = str(len(failed_orders))
    return response

# ----------------- 5. 재고 관리 (WMS) 및 입고 API -----------------
@app.route('/api/inventory', methods=['GET'])
def get_inventory():
    conn = get_connection()
    cur = conn.cursor()
    paged_request = 'page' in request.args or 'search' in request.args
    page_size = min(200, max(20, request.args.get('per_page', default=100, type=int)))
    page = max(0, request.args.get('page', default=0, type=int))
    search = request.args.get('search', '').strip()
    where_sql = 'p.tenant_id = ?'
    where_params = [g.company_id]
    if search:
        pattern = f'%{search}%'
        where_sql += ''' AND (
            p.sku LIKE ? OR p.product_name LIKE ? OR p.english_name LIKE ? OR p.hs_code LIKE ?
            OR EXISTS (SELECT 1 FROM product_options po WHERE po.tenant_id = p.tenant_id
                AND (po.product_id = p.id OR po.sku = p.sku) AND (po.option_name LIKE ? OR po.sku LIKE ?))
        )'''
        where_params.extend([pattern] * 6)
    if paged_request:
        cur.execute(f'SELECT COUNT(*) AS total FROM products p WHERE {where_sql}', where_params)
        total_products = int(cur.fetchone()['total'] or 0)
        page = min(page, max(0, (total_products - 1) // page_size))
        cur.execute(f'''SELECT p.*, COALESCE(p.dispatched_quantity, 0) AS total_dispatched_quantity
            FROM products p WHERE {where_sql} ORDER BY p.id DESC LIMIT ? OFFSET ?''',
            where_params + [page_size, page * page_size])
    else:
        total_products = 0
        cur.execute(f'''
            SELECT p.*,
                (COALESCE(p.dispatched_quantity, 0) + COALESCE(shipped.quantity, 0)) AS total_dispatched_quantity
            FROM products p
            LEFT JOIN (
                SELECT oi.tenant_id, oi.sku, SUM(oi.quantity) AS quantity
                FROM order_items oi
                JOIN orders o ON o.id = oi.order_id AND o.tenant_id = oi.tenant_id
                WHERE oi.tenant_id = ? AND o.status = '출고완료'
                GROUP BY oi.tenant_id, oi.sku
            ) shipped ON shipped.tenant_id = p.tenant_id AND shipped.sku = p.sku
            WHERE {where_sql} ORDER BY p.id DESC
        ''', [g.company_id] + where_params)
    import math
    raw_products = [dict(p) for p in cur.fetchall()]
    product_ids = [p['id'] for p in raw_products]
    product_skus = list(dict.fromkeys(str(p.get('sku') or '') for p in raw_products if p.get('sku')))
    if paged_request and product_skus:
        sku_placeholders = ','.join(['?'] * len(product_skus))
        cur.execute(f'''SELECT oi.sku, SUM(oi.quantity) AS quantity
            FROM order_items oi JOIN orders o ON o.id = oi.order_id AND o.tenant_id = oi.tenant_id
            WHERE oi.tenant_id = ? AND o.status = '출고완료' AND oi.sku IN ({sku_placeholders})
            GROUP BY oi.sku''', [g.company_id] + product_skus)
        dispatched_by_sku = {str(row['sku'] or ''): int(row['quantity'] or 0) for row in cur.fetchall()}
        for product in raw_products:
            product['total_dispatched_quantity'] = int(product.get('total_dispatched_quantity') or 0) + dispatched_by_sku.get(str(product.get('sku') or ''), 0)
    if product_ids:
        product_placeholders = ','.join(['?'] * len(product_ids))
        sku_clause = ''
        option_params = [g.company_id] + product_ids
        if product_skus:
            sku_placeholders = ','.join(['?'] * len(product_skus))
            sku_clause = f' OR sku IN ({sku_placeholders})'
            option_params.extend(product_skus)
        cur.execute(f'''SELECT * FROM product_options
            WHERE tenant_id = ? AND (product_id IN ({product_placeholders}){sku_clause}) ORDER BY id ASC''', option_params)
    else:
        cur.execute('SELECT * FROM product_options WHERE 1 = 0')
    options_by_product = {}
    options_by_sku = {}
    for row in cur.fetchall():
        option = dict(row)
        options_by_product.setdefault(option.get('product_id'), []).append(option)
        options_by_sku.setdefault(option.get('sku'), []).append(option)

    products = []
    for p in raw_products:
        raw_options_by_id = {
            opt['id']: opt
            for opt in options_by_product.get(p['id'], []) + options_by_sku.get(p.get('sku'), [])
        }
        raw_options = list(raw_options_by_id.values())
        cleaned_options = []
        for opt in raw_options:
            o_dict = dict(opt)
            for k in ['unit_price']:
                try:
                    val = float(o_dict.get(k, 0.0) or 0.0)
                    o_dict[k] = 0.0 if (math.isnan(val) or math.isinf(val)) else val
                except:
                    o_dict[k] = 0.0
            cleaned_options.append(o_dict)
        p['options'] = cleaned_options

        # NaN 및 None 방어 처리 (JSON 직렬화 및 브라우저 파싱 호환)
        for k in ['unit_price']:
            try:
                val = float(p.get(k, 0.0) or 0.0)
                p[k] = 0.0 if (math.isnan(val) or math.isinf(val)) else val
            except:
                p[k] = 0.0
        for k in ['stock_quantity', 'dispatched_quantity', 'total_dispatched_quantity']:
            try:
                p[k] = int(p.get(k, 0) or 0)
            except:
                p[k] = 0
        for k in ['english_name', 'hs_code', 'product_url']:
            v = str(p.get(k, '') or '').strip()
            p[k] = '' if v.lower() == 'nan' else v
        products.append(p)
    conn.close()
    if paged_request:
        return jsonify({'items': products, 'total': total_products, 'page': page, 'per_page': page_size})
    return jsonify(products)

@app.route('/api/inventory/search', methods=['GET'])
def search_product_by_barcode():
    barcode = request.args.get('barcode', '').strip()
    if not barcode:
        return jsonify({'found': False, 'message': '바코드를 입력하세요.'})

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM products WHERE tenant_id = ? AND sku = ? LIMIT 1', (g.company_id, barcode))
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
    cur.execute('SELECT * FROM products WHERE tenant_id = ? AND sku = ?', (g.company_id, sku))
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
            WHERE id = ? AND tenant_id = ?
        ''', (new_stock, product_name, product_name, english_name, english_name, hs_code, hs_code, product_url, product_url, unit_price, unit_price, prod['id'], g.company_id))
        prod_id = prod['id']
        p_name = product_name or prod['product_name']
    else:
        if not product_name:
            product_name = f"상품_{sku}"
        new_stock = qty
        cur.execute('''
            INSERT INTO products (tenant_id, sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity, product_url)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (g.company_id, sku, product_name, '', english_name, hs_code, unit_price, new_stock, product_url))
        prod_id = cur.lastrowid
        p_name = product_name



    # 입고 이력 로그
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    cur.execute('''
        INSERT INTO inventory_logs (tenant_id, product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (g.company_id, prod_id, sku, p_name, '', '입고', qty, new_stock, memo, now_str))

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

        cur.execute('SELECT * FROM products WHERE tenant_id = ? AND sku = ?', (g.company_id, sku))
        prod = cur.fetchone()

        if prod:
            new_stock = prod['stock_quantity'] + qty
            cur.execute('''
                UPDATE products
                SET stock_quantity = ?,
                    product_name = CASE WHEN ? != '' THEN ? ELSE product_name END,
                    unit_price = CASE WHEN ? > 0 THEN ? ELSE unit_price END,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ? AND tenant_id = ?
            ''', (new_stock, p_name, p_name, unit_price, unit_price, prod['id'], g.company_id))
            prod_id = prod['id']
            actual_name = p_name or prod['product_name']
        else:
            actual_name = p_name or f"상품_{sku}"
            new_stock = qty
            cur.execute('''
                INSERT INTO products (tenant_id, sku, product_name, option_name, unit_price, stock_quantity)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (g.company_id, sku, actual_name, '', unit_price, new_stock))
            prod_id = cur.lastrowid


        cur.execute('''
            INSERT INTO inventory_logs (tenant_id, product_id, sku, product_name, option_name, change_type, change_qty, resulting_stock, memo, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (g.company_id, prod_id, sku, actual_name, '', '입고', qty, new_stock, '일괄 바코드 스캔 입고', now_str))

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

        import math
        for _, row in df.iterrows():
            sku = str(row.get(col_map.get('sku', ''), '') or '').strip()
            name = str(row.get(col_map.get('product_name', ''), '') or '').strip()
            eng = str(row.get(col_map.get('english_name', ''), '') or '').strip()
            hs = str(row.get(col_map.get('hs_code', ''), '') or '').strip()
            url = str(row.get(col_map.get('product_url', ''), '') or '').strip()
            
            # 숫자형 바코드(SKU) 뒤에 붙은 .0 자동 제거 (예: 4531663030487.0 -> 4531663030487)
            if sku.endswith('.0') and sku[:-2].isdigit():
                sku = sku[:-2]
            if sku.lower() == 'nan': sku = ''
            if name.lower() == 'nan': name = ''
            if eng.lower() == 'nan': eng = ''
            if hs.lower() == 'nan': hs = ''
            if url.lower() == 'nan': url = ''
            
            # 단가(price) NaN 및 무한대 방어
            try:
                p_val = row.get(col_map.get('unit_price', ''), 0.0)
                if pd.isna(p_val) or str(p_val).strip().lower() == 'nan':
                    price = 0.0
                else:
                    price = float(p_val)
                if math.isnan(price) or math.isinf(price):
                    price = 0.0
            except:
                price = 0.0

            # 재고(stock) 수량 정수 변환 및 방어
            try:
                s_val = row.get(col_map.get('stock_quantity', ''), 0)
                if pd.isna(s_val) or str(s_val).strip().lower() == 'nan':
                    stock = 0
                else:
                    stock = int(float(s_val))
            except:
                stock = 0

            # 1회 출고 수량 정수 변환 및 방어
            try:
                d_val = row.get(col_map.get('dispatched_quantity', ''), 1)
                if pd.isna(d_val) or str(d_val).strip().lower() == 'nan':
                    dispatched = 1
                else:
                    dispatched = int(float(d_val))
            except:
                dispatched = 1

            if not sku and not name:
                continue
            if not sku:
                sku = f"SKU-{abs(hash(name)) % 1000000:06d}"

            cur.execute('SELECT * FROM products WHERE tenant_id = ? AND sku = ?', (g.company_id, sku))
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
                    WHERE id = ? AND tenant_id = ?
                ''', (name, name, eng, eng, hs, hs, url, url, stock, stock, price, price, existing['id'], g.company_id))
                prod_id = existing['id']
            else:
                cur.execute('''
                    INSERT INTO products (tenant_id, sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (g.company_id, sku, name or f"상품_{sku}", '', eng, hs, price, stock, dispatched, url))
                prod_id = cur.lastrowid

                cur.execute('''
                    INSERT INTO inventory_logs (tenant_id, product_id, sku, product_name, change_type, change_qty, resulting_stock, memo, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (g.company_id, prod_id, sku, name or sku, '엑셀등록', stock, stock, '엑셀 상품 일괄 등록', now_str))

            # 옵션 마스터(product_options)에 옵션 등록
            if name:
                cur.execute('SELECT id FROM product_options WHERE tenant_id = ? AND sku = ? AND option_name = ?', (g.company_id, sku, name))
                opt_row = cur.fetchone()
                if opt_row:
                    cur.execute('''
                        UPDATE product_options
                        SET unit_price = ?, dispatched_quantity = ?, updated_at = CURRENT_TIMESTAMP
                        WHERE id = ? AND tenant_id = ?
                    ''', (price, dispatched, opt_row['id'], g.company_id))
                else:
                    cur.execute('''
                        INSERT INTO product_options (tenant_id, product_id, sku, option_name, unit_price, dispatched_quantity)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (g.company_id, prod_id, sku, name, price, dispatched))

                # 미등록 상품 목록에 있으면 등록 완료 처리
                cur.execute("UPDATE unregistered_products SET status = 'registered' WHERE tenant_id = ? AND (combined_name = ? OR product_name = ?)", (g.company_id, name, name))

            saved_count += 1

        from excel_processor import sync_order_items_sku
        sync_order_items_sku(conn, tenant_id=g.company_id)

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
                WHERE id = ? AND tenant_id = ?
            ''', (sku, product_name, english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url, prod_id, g.company_id))
        else:
            cur.execute('''
                INSERT INTO products (tenant_id, sku, product_name, option_name, english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (g.company_id, sku, product_name, '', english_name, hs_code, unit_price, stock_quantity, dispatched_quantity, product_url))
            prod_id = cur.lastrowid

        # 옵션 목록 갱신 (전달된 옵션이 있는 경우)
        if options is not None:
            cur.execute('DELETE FROM product_options WHERE tenant_id = ? AND (sku = ? OR product_id = ?)', (g.company_id, sku, prod_id))
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
                    INSERT INTO product_options (tenant_id, product_id, sku, option_name, unit_price, dispatched_quantity)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (g.company_id, prod_id, sku, opt_name, opt_price, opt_qty))

                cur.execute("UPDATE unregistered_products SET status = 'registered' WHERE tenant_id = ? AND (combined_name = ? OR product_name = ?)", (g.company_id, opt_name, opt_name))

        # 미등록 상품 테이블 상태 업데이트
        cur.execute("UPDATE unregistered_products SET status = 'registered' WHERE tenant_id = ? AND (combined_name = ? OR product_name = ?)", (g.company_id, product_name, product_name))

        # 주문 품목들의 바코드 동기화
        from excel_processor import sync_order_items_sku
        sync_order_items_sku(conn, tenant_id=g.company_id)

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
    cur.execute(f'DELETE FROM product_options WHERE tenant_id = ? AND product_id IN ({placeholders})', [g.company_id] + product_ids)
    cur.execute(f'DELETE FROM inventory_logs WHERE tenant_id = ? AND product_id IN ({placeholders})', [g.company_id] + product_ids)
    cur.execute(f'DELETE FROM products WHERE tenant_id = ? AND id IN ({placeholders})', [g.company_id] + product_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'deleted_count': len(product_ids)})

@app.route('/api/inventory/product/<int:prod_id>', methods=['DELETE'])
def delete_product(prod_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM product_options WHERE tenant_id = ? AND product_id = ?', (g.company_id, prod_id))
    cur.execute('DELETE FROM inventory_logs WHERE tenant_id = ? AND product_id = ?', (g.company_id, prod_id))
    cur.execute('DELETE FROM products WHERE tenant_id = ? AND id = ?', (g.company_id, prod_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/inventory/unregistered-products', methods=['GET'])
def get_unregistered_products():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM unregistered_products WHERE tenant_id = ? AND status = 'pending' ORDER BY order_count DESC, id DESC", (g.company_id,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify(rows)

@app.route('/api/inventory/unregistered-products/<int:item_id>/ignore', methods=['POST'])
def ignore_unregistered_product(item_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("UPDATE unregistered_products SET status = 'ignored' WHERE tenant_id = ? AND id = ?", (g.company_id, item_id))
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
    cur.execute(f'UPDATE orders SET status = ? WHERE tenant_id = ? AND id IN ({placeholders})', [target_status, g.company_id] + order_ids)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'moved_count': len(order_ids), 'target_status': target_status})


@app.route('/api/inventory/logs', methods=['GET'])
def get_inventory_logs():
    selected_date = request.args.get('date', '').strip()
    limit = int(request.args.get('limit', 200))
    search = request.args.get('search', '').strip()
    change_type = request.args.get('type', '').strip()

    conn = get_connection()
    cur = conn.cursor()
    query = 'SELECT * FROM inventory_logs WHERE tenant_id = ?'
    params = [g.company_id]

    if change_type and change_type != '전체':
        if change_type == '출고':
            query += " AND change_type IN ('출고', '주문출고')"
        else:
            query += ' AND change_type = ?'
            params.append(change_type)

    if search:
        query += ' AND (sku LIKE ? OR product_name LIKE ? OR memo LIKE ?)'
        wild = f'%{search}%'
        params.extend([wild, wild, wild])

    has_valid_date = bool(re.fullmatch(r'\d{4}-\d{2}-\d{2}', selected_date))
    if has_valid_date:
        query += ' AND SUBSTR(CAST(created_at AS TEXT), 1, 10) = ?'
        params.append(selected_date)

    # 날짜를 선택하면 그 날짜의 전체 내역을 표시하고, 날짜를 선택하지 않은 경우 최근 내역만 표시합니다.
    if has_valid_date:
        query += ' ORDER BY id DESC'
    else:
        query += ' ORDER BY id DESC LIMIT ?'
        params.append(limit)

    cur.execute(query, params)
    logs = [dict(l) for l in cur.fetchall()]
    conn.close()
    logs.reverse()
    for log in logs:
        log['created_at'] = format_inventory_log_datetime(log.get('created_at'))
    return jsonify(logs)

@app.route('/api/inventory/logs/dates', methods=['GET'])
def get_inventory_log_dates():
    search = request.args.get('search', '').strip()
    change_type = request.args.get('type', '').strip()
    conn = get_connection()
    cur = conn.cursor()
    query = "SELECT SUBSTR(CAST(created_at AS TEXT), 1, 10) AS date, COUNT(*) AS count FROM inventory_logs WHERE tenant_id = ?"
    params = [g.company_id]
    if change_type and change_type != '전체':
        if change_type == '출고':
            query += " AND change_type IN ('출고', '주문출고')"
        else:
            query += ' AND change_type = ?'
            params.append(change_type)
    if search:
        query += ' AND (sku LIKE ? OR product_name LIKE ? OR memo LIKE ?)'
        wild = f'%{search}%'
        params.extend([wild, wild, wild])
    query += ' GROUP BY date ORDER BY date DESC'
    cur.execute(query, params)
    dates = [dict(row) for row in cur.fetchall()]
    conn.close()
    return jsonify(dates)

@app.route('/api/inventory/export-logs', methods=['GET'])
def export_inventory_logs():
    ids_str = request.args.get('ids', '').strip()
    selected_date = request.args.get('date', '').strip()
    search = request.args.get('search', '').strip()
    change_type = request.args.get('type', '').strip()
    conn = get_connection()
    cur = conn.cursor()

    if ids_str:
        ids = [int(x) for x in ids_str.split(',') if x.strip().isdigit()]
        placeholders = ','.join(['?'] * len(ids))
        cur.execute(f'SELECT * FROM inventory_logs WHERE tenant_id = ? AND id IN ({placeholders}) ORDER BY id ASC', [g.company_id] + ids)
    elif re.fullmatch(r'\d{4}-\d{2}-\d{2}', selected_date):
        query = 'SELECT * FROM inventory_logs WHERE tenant_id = ? AND SUBSTR(CAST(created_at AS TEXT), 1, 10) = ?'
        params = [g.company_id, selected_date]
        if change_type and change_type != '전체':
            if change_type == '출고':
                query += " AND change_type IN ('출고', '주문출고')"
            else:
                query += ' AND change_type = ?'
                params.append(change_type)
        if search:
            query += ' AND (sku LIKE ? OR product_name LIKE ? OR memo LIKE ?)'
            wild = f'%{search}%'
            params.extend([wild, wild, wild])
        cur.execute(query + ' ORDER BY id ASC', params)
    else:
        cur.execute('SELECT * FROM (SELECT * FROM inventory_logs WHERE tenant_id = ? ORDER BY id DESC LIMIT 1000) recent_logs ORDER BY id ASC', (g.company_id,))

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
        date_str = format_inventory_log_datetime(l['created_at'])
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
        download_name=f'입출고내역_{selected_date}.xlsx' if re.fullmatch(r'\d{4}-\d{2}-\d{2}', selected_date) else '입출고내역리스트.xlsx',
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
    cur.execute(f'DELETE FROM inventory_logs WHERE tenant_id = ? AND id IN ({placeholders})', [g.company_id] + log_ids)
    conn.commit()
    conn.close()

    return jsonify({'status': 'success', 'deleted_count': len(log_ids)})

# ----------------- 6. 세트상품/옵션 분할 규칙 API -----------------
@app.route('/api/bundles', methods=['GET'])
def get_bundles():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM bundle_rules WHERE tenant_id = ? ORDER BY id DESC', (g.company_id,))
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
            tenant_id, rule_name, source_product_name, source_option_name,
            target_product_name, target_option_name, target_sku, multiplier
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (g.company_id, rule_name or f"{source_o} -> {target_o}", source_p or '*', source_o, target_p, target_o, target_sku, multiplier))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/bundles/<int:rule_id>', methods=['DELETE'])
def delete_bundle(rule_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('DELETE FROM bundle_rules WHERE tenant_id = ? AND id = ?', (g.company_id, rule_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# ----------------- 6-2. 옵션 자동 분할 구분 기호 API -----------------
@app.route('/api/delimiters', methods=['GET'])
def get_delimiters():
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('SELECT symbol FROM split_delimiters WHERE tenant_id = ? ORDER BY id ASC', (g.company_id,))
        delimiters = [r['symbol'] if (isinstance(r, dict) or hasattr(r, '__getitem__')) else r[0] for r in cur.fetchall()]
        if not delimiters:
            delimiters = ['+', '-', '/', ',', '?', '&']
        return jsonify({'status': 'success', 'delimiters': delimiters})
    except Exception as e:
        return jsonify({'status': 'error', 'delimiters': ['+', '-', '/', ',', '?', '&'], 'error': str(e)})
    finally:
        conn.close()

@app.route('/api/delimiters', methods=['POST'])
def add_delimiter():
    data = request.get_json() or {}
    symbol = str(data.get('symbol', '')).strip()
    if not symbol:
        return jsonify({'status': 'error', 'message': '기호를 입력해주세요.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('INSERT INTO split_delimiters (tenant_id, symbol) VALUES (?, ?) ON CONFLICT DO NOTHING', (g.company_id, symbol))
        conn.commit()
        cur.execute('SELECT symbol FROM split_delimiters WHERE tenant_id = ? ORDER BY id ASC', (g.company_id,))
        delimiters = [r['symbol'] if (isinstance(r, dict) or hasattr(r, '__getitem__')) else r[0] for r in cur.fetchall()]
        return jsonify({'status': 'success', 'message': f"'{symbol}' 기호가 추가되었습니다.", 'delimiters': delimiters})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/delimiters/delete', methods=['POST'])
def delete_delimiter():
    data = request.get_json() or {}
    symbol = str(data.get('symbol', '')).strip()
    if not symbol:
        return jsonify({'status': 'error', 'message': '삭제할 기호를 지정해주세요.'}), 400
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('DELETE FROM split_delimiters WHERE tenant_id = ? AND symbol = ?', (g.company_id, symbol))
        conn.commit()
        cur.execute('SELECT symbol FROM split_delimiters WHERE tenant_id = ? ORDER BY id ASC', (g.company_id,))
        delimiters = [r['symbol'] if (isinstance(r, dict) or hasattr(r, '__getitem__')) else r[0] for r in cur.fetchall()]
        return jsonify({'status': 'success', 'message': f"'{symbol}' 기호가 삭제되었습니다.", 'delimiters': delimiters})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500
    finally:
        conn.close()

# ----------------- 7. 택배사 & 발송자 설정 API -----------------
@app.route('/api/shipping-rates/template', methods=['GET'])
def download_shipping_rates_template():
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '배송요금표'
    headers = ['무게 (kg)', '항공요금', '해상요금']
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='2563EB')
        cell.alignment = Alignment(horizontal='center')
    for weight in SHIPPING_WEIGHT_STEPS:
        sheet.append([weight, None, None])
        sheet.cell(sheet.max_row, 1).number_format = '0.0'
        sheet.cell(sheet.max_row, 2).number_format = '#,##0.##'
        sheet.cell(sheet.max_row, 3).number_format = '#,##0.##'
    sheet.column_dimensions['A'].width = 16
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False

    instructions = workbook.create_sheet('안내')
    instructions.append(['배송요금표 엑셀 업로드 안내'])
    instructions.append(['배송요금표 시트에서 항공요금과 해상요금 중 입력할 요금만 작성해 업로드하세요. 한쪽 요금만 입력해도 반영됩니다.'])
    instructions.append(['요금 칸을 비워두면 해당 무게의 기존 요금은 유지됩니다.'])
    instructions.append(['무게 기준은 양식에 있는 값을 그대로 사용해주세요.'])
    instructions.column_dimensions['A'].width = 85
    instructions['A1'].font = Font(bold=True, size=14, color='2563EB')
    for row in range(2, 5):
        instructions.cell(row, 1).alignment = Alignment(wrap_text=True, vertical='top')

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='배송요금표_기본양식.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/shipping-rates/upload', methods=['POST'])
def upload_shipping_rates_excel():
    if session.get('account_role') != 'admin':
        return jsonify({'error': '배송요금표는 관리자만 업로드할 수 있습니다.'}), 403
    from io import BytesIO
    import openpyxl

    uploaded = request.files.get('file')
    if not uploaded or not uploaded.filename:
        return jsonify({'error': '배송요금표 엑셀 파일을 선택해주세요.'}), 400
    if not uploaded.filename.lower().endswith('.xlsx'):
        return jsonify({'error': '.xlsx 엑셀 파일만 업로드할 수 있습니다.'}), 400

    try:
        workbook = openpyxl.load_workbook(BytesIO(uploaded.read()), data_only=True, read_only=True)
    except Exception:
        return jsonify({'error': '엑셀 파일을 읽을 수 없습니다. 기본 양식의 .xlsx 파일인지 확인해주세요.'}), 400

    def normalize_header(value):
        return ''.join(char for char in str(value or '').strip().lower() if char.isalnum() or '\uac00' <= char <= '\ud7a3')

    aliases = {
        'weight': {'무게', '무게kg', '중량', '중량kg', 'weight', 'weightkg', 'kg'},
        'air': {'항공', '항공요금', '항공요금표', 'air', 'airrate'},
        'sea': {'해상', '해상요금', '해상요금표', 'sea', 'searate'}
    }

    def match_header(key, name):
        if name in aliases[key]:
            return True
        if key == 'weight':
            return name.startswith(('무게', '중량', 'weight'))
        if key == 'air':
            return name.startswith('항공') or (name.startswith('air') and 'rate' in name)
        if key == 'sea':
            return name.startswith('해상') or (name.startswith('sea') and 'rate' in name)
        return False

    selected_sheet = next((workbook[name] for name in ('배송요금표', 'Shipping Rates') if name in workbook.sheetnames), workbook.active)
    # 엑셀에서 서식만 멀리 적용된 파일도 빠르게 처리하도록 실제 요금표 범위만 읽습니다.
    rows = selected_sheet.iter_rows(
        max_row=min(selected_sheet.max_row or 500, 500),
        max_col=min(selected_sheet.max_column or 50, 50),
        values_only=True
    )
    header_map = None
    data_rows = []
    for row_number, row in enumerate(rows, start=1):
        normalized = [normalize_header(value) for value in row]
        if header_map is None:
            found = {key: next((idx for idx, name in enumerate(normalized) if match_header(key, name)), None)
                     for key in aliases}
            # 무게 열과 항공/해상 중 하나의 요금 열만 있으면 업로드 가능하게 한다.
            if found['weight'] is not None and (found['air'] is not None or found['sea'] is not None):
                header_map = found
            continue
        data_rows.append((row_number, row))
    workbook.close()
    if header_map is None:
        return jsonify({'error': '무게 열과 항공요금 또는 해상요금 열을 찾지 못했습니다. 기본 양식의 제목을 확인해주세요.'}), 400

    allowed_weights = {round(float(weight), 1) for weight in SHIPPING_WEIGHT_STEPS}
    parsed = {}
    for row_number, row in data_rows:
        values = [
            row[header_map[key]] if header_map.get(key) is not None and header_map[key] < len(row) else None
            for key in ('weight', 'air', 'sea')
        ]
        if all(value is None or str(value).strip() == '' for value in values):
            continue
        try:
            weight_text = str(values[0]).strip().lower().replace('kg', '').strip()
            weight = round(float(weight_text), 1)
        except (TypeError, ValueError):
            return jsonify({'error': f'{row_number}행의 무게를 숫자로 읽을 수 없습니다.'}), 400
        if weight not in allowed_weights:
            return jsonify({'error': f'{row_number}행의 무게 {weight:g}kg는 배송요금표 기준에 없습니다.'}), 400
        if weight in parsed:
            return jsonify({'error': f'{row_number}행에 무게 {weight:.1f}kg가 중복되어 있습니다.'}), 400

        rate_values = {}
        for key, value in (('air_rate', values[1]), ('sea_rate', values[2])):
            if value is None or str(value).strip() == '':
                rate_values[key] = None
                continue
            try:
                amount_text = str(value).strip().replace(',', '').replace(' ', '')
                for unit in ('JPY', 'KRW', '￥', '¥', '₩', '원', '엔'):
                    amount_text = amount_text.replace(unit, '')
                amount = float(amount_text)
            except (TypeError, ValueError):
                return jsonify({'error': f'{row_number}행의 {"항공" if key == "air_rate" else "해상"}요금이 숫자가 아닙니다.'}), 400
            if not math.isfinite(amount) or amount < 0 or amount > 999999999:
                return jsonify({'error': f'{row_number}행의 요금은 0 이상 999,999,999 이하로 입력해주세요.'}), 400
            rate_values[key] = amount
        if rate_values['air_rate'] is None and rate_values['sea_rate'] is None:
            continue
        parsed[weight] = rate_values

    if not parsed:
        return jsonify({'error': '반영할 요금이 없습니다. 항공요금 또는 해상요금 칸에 값을 입력해주세요.'}), 400

    conn = get_connection()
    cur = conn.cursor()
    try:
        weights = list(parsed)
        placeholders = ','.join(['?'] * len(weights))
        cur.execute(f'''SELECT weight_kg, air_rate, sea_rate FROM shipping_rates
            WHERE tenant_id = ? AND weight_kg IN ({placeholders})''', [g.company_id] + weights)
        existing = {round(float(row['weight_kg']), 1): row for row in cur.fetchall()}
        normalized = []
        for weight, rate_values in parsed.items():
            old = existing.get(weight)
            air_rate = rate_values['air_rate'] if rate_values['air_rate'] is not None else (float(old['air_rate'] or 0) if old else 0)
            sea_rate = rate_values['sea_rate'] if rate_values['sea_rate'] is not None else (float(old['sea_rate'] or 0) if old else 0)
            normalized.append((g.company_id, weight, air_rate, sea_rate))
        cur.executemany('''
            INSERT INTO shipping_rates (tenant_id, weight_kg, air_rate, sea_rate, updated_at)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT (tenant_id, weight_kg) DO UPDATE SET
                air_rate = excluded.air_rate,
                sea_rate = excluded.sea_rate,
                updated_at = CURRENT_TIMESTAMP
        ''', normalized)
        conn.commit()
        return jsonify({'status': 'success', 'saved_count': len(normalized)})
    except Exception:
        conn.rollback()
        app.logger.exception('Shipping rate Excel upload failed')
        return jsonify({'error': '배송요금 엑셀 반영 중 오류가 발생했습니다.'}), 500
    finally:
        conn.close()


@app.route('/api/shipping-rates', methods=['GET'])
def get_shipping_rates():
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('''SELECT weight_kg, air_rate, sea_rate FROM shipping_rates
                       WHERE tenant_id = ? ORDER BY weight_kg ASC''', (g.company_id,))
        rates = complete_shipping_rate_rows(cur.fetchall())
        cur.execute('SELECT additional_package_rate, box_package_rate FROM shipping_packaging_rates WHERE tenant_id = ?', (g.company_id,))
        packaging = cur.fetchone()
        return jsonify({'rates': rates, 'packaging_fees': dict(packaging) if packaging else {'additional_package_rate': 0, 'box_package_rate': 0}})
    finally:
        conn.close()


@app.route('/api/shipping-rates', methods=['POST'])
def save_shipping_rates():
    if session.get('account_role') != 'admin':
        return jsonify({'error': '배송요금표는 관리자만 수정할 수 있습니다.'}), 403
    data = request.get_json(silent=True) or {}
    rates = data.get('rates')
    if not isinstance(rates, list):
        return jsonify({'error': '요금표 데이터 형식이 올바르지 않습니다.'}), 400

    packaging_fees = data.get('packaging_fees')
    normalized_packaging = None
    if packaging_fees is not None:
        if not isinstance(packaging_fees, dict):
            return jsonify({'error': '포장 요금 데이터 형식이 올바르지 않습니다.'}), 400
        try:
            additional_fee = float(packaging_fees.get('additional_package_rate') or 0)
            box_fee = float(packaging_fees.get('box_package_rate') or 0)
        except (TypeError, ValueError):
            return jsonify({'error': '포장 요금은 숫자로 입력해주세요.'}), 400
        if not all(math.isfinite(value) and 0 <= value <= 999999999 for value in (additional_fee, box_fee)):
            return jsonify({'error': '포장 요금은 0 이상 999,999,999 이하 숫자로 입력해주세요.'}), 400
        normalized_packaging = (g.company_id, additional_fee, box_fee)

    allowed_weights = {round(float(weight), 1) for weight in SHIPPING_WEIGHT_STEPS}
    normalized = []
    seen = set()
    for row in rates:
        if not isinstance(row, dict):
            return jsonify({'error': '요금표 행 형식이 올바르지 않습니다.'}), 400
        try:
            weight = round(float(row.get('weight_kg')), 1)
            air_rate = float(row.get('air_rate') or 0)
            sea_rate = float(row.get('sea_rate') or 0)
        except (TypeError, ValueError):
            return jsonify({'error': '무게와 요금은 숫자로 입력해주세요.'}), 400
        if weight not in allowed_weights or weight in seen:
            return jsonify({'error': '무게 기준이 중복되었거나 요금표에 없는 값입니다.'}), 400
        if not all(math.isfinite(value) and 0 <= value <= 999999999 for value in (air_rate, sea_rate)):
            return jsonify({'error': '요금은 0 이상 999,999,999 이하 숫자로 입력해주세요.'}), 400
        seen.add(weight)
        normalized.append((g.company_id, weight, air_rate, sea_rate))

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.executemany('''
            INSERT INTO shipping_rates (tenant_id, weight_kg, air_rate, sea_rate, updated_at)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT (tenant_id, weight_kg) DO UPDATE SET
                air_rate = excluded.air_rate,
                sea_rate = excluded.sea_rate,
                updated_at = CURRENT_TIMESTAMP
        ''', normalized)
        if normalized_packaging:
            cur.execute('''
                INSERT INTO shipping_packaging_rates (tenant_id, additional_package_rate, box_package_rate, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT (tenant_id) DO UPDATE SET
                    additional_package_rate = excluded.additional_package_rate,
                    box_package_rate = excluded.box_package_rate,
                    updated_at = CURRENT_TIMESTAMP
            ''', normalized_packaging)
        conn.commit()
        return jsonify({'status': 'success', 'saved_count': len(normalized), 'packaging_saved': normalized_packaging is not None})
    except Exception as e:
        conn.rollback()
        app.logger.exception('Shipping rate save failed')
        return jsonify({'error': '배송요금표 저장에 실패했습니다.'}), 500
    finally:
        conn.close()


@app.route('/api/settings', methods=['GET'])
def get_settings():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute('SELECT * FROM courier_settings WHERE tenant_id = ? ORDER BY id LIMIT 1', (g.company_id,))
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
        SET courier_name = '나르다', customer_code = ?, api_key = ?, narda_id = ?,
            is_mock = 0, stock_only_dispatch = ?, updated_at = CURRENT_TIMESTAMP
        WHERE tenant_id = ?
    ''', (
        data.get('customer_code', ''),
        data.get('api_key', ''),
        data.get('narda_id', 'calroid'),
        1 if data.get('stock_only_dispatch') else 0,
        g.company_id
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



