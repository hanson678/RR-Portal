# -*- coding: utf-8 -*-
"""ZURU 河源排期入单系统 — 云端版
功能：解析PO Excel + 与河源排期匹配 → 修改单明细 + 生成新单Excel
云端模式：排期文件通过网页上传，不依赖本地路径/WPS COM
"""
import os
import sys
import json
import time
import logging
import pickle
import re
import secrets
import subprocess
import threading
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, after_this_request
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix

from excel_po_parser import ExcelPOParser
from hy_schedule import analyze_orders, write_orders
from fy_schedule import generate_fy_export

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, 'data')
EXPORT_DIR = os.path.join(APP_DIR, 'exports')
UPLOAD_DIR = os.path.join(APP_DIR, 'uploads')
SCHEDULE_DIR = os.path.join(APP_DIR, 'uploads', 'schedules')

for d in (DATA_DIR, EXPORT_DIR, UPLOAD_DIR, SCHEDULE_DIR):
    os.makedirs(d, exist_ok=True)

app = Flask(__name__,
            template_folder=os.path.join(APP_DIR, 'templates'),
            static_folder=os.path.join(APP_DIR, 'static'))
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)
app.config['UPLOAD_FOLDER'] = UPLOAD_DIR
app.config['SCHEDULE_FOLDER'] = SCHEDULE_DIR
app.config['EXPORT_FOLDER'] = EXPORT_DIR
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

LOG_FILE = os.path.join(DATA_DIR, 'ops.log')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
logger = logging.getLogger('hy_schedule')
logger.setLevel(logging.INFO)
if not logger.handlers:
    _fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
    _fh.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
    logger.addHandler(_fh)


# ========== 排期文件状态 ==========

def _get_schedule_info():
    """读取排期文件上传状态"""
    try:
        files = [f for f in os.listdir(SCHEDULE_DIR)
                 if f.endswith('.xlsx') and not f.startswith('~$')]
    except Exception:
        files = []
    map_path = os.path.join(DATA_DIR, 'hy_item_map.json')
    item_count = 0
    if os.path.exists(map_path):
        try:
            with open(map_path, 'r', encoding='utf-8') as f:
                item_count = len(json.load(f))
        except Exception:
            pass
    return {
        'file_count': len(files),
        'files': sorted(files),
        'item_count': item_count,
    }


def _run_scan():
    """运行 scan_hy_items.py 全量重建货号映射表"""
    script = os.path.join(APP_DIR, 'scan_hy_items.py')
    # Windows上隐藏CMD窗口
    kwargs = {}
    if os.name == 'nt':
        kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
    try:
        result = subprocess.run(
            [sys.executable, script, SCHEDULE_DIR],
            capture_output=True, text=True, timeout=600, encoding='utf-8',
            **kwargs
        )
        logger.info(f'[河源] 全量重建映射表: rc={result.returncode}')
        return result.returncode == 0, result.stdout, result.stderr
    except Exception as e:
        logger.error(f'[河源] 扫描失败: {e}')
        return False, '', str(e)


def _load_item_map():
    """加载现有映射表"""
    map_path = os.path.join(DATA_DIR, 'hy_item_map.json')
    if os.path.exists(map_path):
        try:
            with open(map_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_item_map(mapping):
    """保存映射表"""
    map_path = os.path.join(DATA_DIR, 'hy_item_map.json')
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(map_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)


def _incremental_scan(filepaths):
    """增量扫描：只扫描指定文件，merge到现有映射表"""
    sys.path.insert(0, APP_DIR)
    from scan_hy_items import scan_single

    mapping = _load_item_map()
    total_new = 0
    for fpath in filepaths:
        fn = os.path.basename(fpath)
        try:
            file_map = scan_single(fpath)
            for k, entries in file_map.items():
                for entry in entries:
                    mapping.setdefault(k, [])
                    if entry not in mapping[k]:
                        mapping[k].append(entry)
                        total_new += 1
        except Exception as e:
            logger.warning(f'[河源] 增量扫描失败 {fn}: {e}')
    _save_item_map(mapping)
    logger.info(f'[河源] 增量扫描: {len(filepaths)}个文件, 新增{total_new}条')
    return len(mapping)


def _remove_file_from_map(filename):
    """从映射表中移除指定文件的所有条目"""
    mapping = _load_item_map()
    to_delete = []
    for k, entries in mapping.items():
        mapping[k] = [e for e in entries if e.get('file') != filename]
        if not mapping[k]:
            to_delete.append(k)
    for k in to_delete:
        del mapping[k]
    _save_item_map(mapping)
    return len(mapping)


# ========== 基于文件的待处理订单缓存（支持多worker/重启不丢失） ==========

SESSION_DIR = os.path.join(DATA_DIR, 'sessions')
os.makedirs(SESSION_DIR, exist_ok=True)
_sessions_lock = threading.Lock()


def _safe_path_under(base_dir, filename):
    """Resolve filename under base_dir, rejecting traversal and sibling-prefix tricks.
    Returns absolute path on success, None on rejection."""
    if not filename:
        return None
    abs_base = os.path.abspath(base_dir)
    abs_target = os.path.abspath(os.path.join(abs_base, filename))
    try:
        if os.path.commonpath([abs_base, abs_target]) != abs_base:
            return None
    except ValueError:
        return None
    return abs_target


def _session_path(session_id):
    if not session_id or not re.match(r'^[a-zA-Z0-9_-]+$', str(session_id)):
        return None
    return os.path.join(SESSION_DIR, f'{session_id}.pkl')


def _save_session(session_id, data):
    path = _session_path(session_id)
    if not path:
        return
    with _sessions_lock:
        with open(path, 'wb') as f:
            pickle.dump(data, f)


def _load_session(session_id):
    path = _session_path(session_id)
    if not path:
        return None
    with _sessions_lock:
        if not os.path.exists(path):
            return None
        try:
            with open(path, 'rb') as f:
                return pickle.load(f)
        except Exception:
            return None


def _delete_session(session_id):
    path = _session_path(session_id)
    if not path:
        return
    with _sessions_lock:
        if os.path.exists(path):
            os.remove(path)


def _cleanup_pending():
    now = datetime.now()
    with _sessions_lock:
        for fn in os.listdir(SESSION_DIR):
            if not fn.endswith('.pkl'):
                continue
            path = os.path.join(SESSION_DIR, fn)
            try:
                with open(path, 'rb') as f:
                    data = pickle.load(f)
                ts = data.get('timestamp')
                if isinstance(ts, datetime) and (now - ts).total_seconds() > 3600:
                    os.remove(path)
            except Exception:
                pass


# ========== 路由 ==========

@app.after_request
def _set_csrf_cookie(response):
    if 'hy_csrf_token' not in request.cookies:
        response.set_cookie('hy_csrf_token', secrets.token_urlsafe(24), samesite='Lax')
    return response


@app.before_request
def _csrf_protect():
    if request.method in ('GET', 'HEAD', 'OPTIONS', 'TRACE'):
        return
    cookie_token = request.cookies.get('hy_csrf_token', '')
    header_token = request.headers.get('X-CSRF-Token', '')
    if not cookie_token or cookie_token != header_token:
        return jsonify({'error': 'CSRF验证失败，请刷新页面后重试'}), 403


@app.route('/')
def index():
    return render_template('hy_master.html')


@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'service': 'hy-schedule-system'})


# ── 排期文件管理 ──

@app.route('/api/hy-info')
def hy_info():
    """排期文件状态"""
    info = _get_schedule_info()
    return jsonify({
        'exists': info['file_count'] > 0,
        'file_count': info['file_count'],
        'files': info['files'],
        'item_count': info['item_count'],
        'dir': SCHEDULE_DIR,
    })


@app.route('/api/hy-upload-schedules', methods=['POST'])
def upload_schedules():
    """上传排期文件（多文件），保存到 uploads/schedules/ 并自动重建映射表"""
    saved = []
    for f in request.files.getlist('files'):
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ('.xlsx', '.xls'):
            continue
        safe_name = secure_filename(f.filename) or f'schedule_{int(time.time())}{ext}'
        if not safe_name.lower().endswith(ext):
            safe_name += ext
        path = os.path.join(SCHEDULE_DIR, safe_name)
        f.save(path)
        saved.append(safe_name)

    if not saved:
        return jsonify({'error': '没有有效的Excel文件'}), 400

    # 增量扫描：只扫描新上传的文件（而非全量重建）
    new_paths = [os.path.join(SCHEDULE_DIR, fn) for fn in saved]
    item_count = _incremental_scan(new_paths)
    info = _get_schedule_info()

    logger.info(f'[河源] 上传排期文件(增量): {saved}')
    return jsonify({
        'ok': True,
        'uploaded': saved,
        'file_count': info['file_count'],
        'item_count': info['item_count'],
        'scan_ok': True,
        'msg': f'已上传{len(saved)}个排期文件，映射表已更新（{info["item_count"]}个货号）',
    })


@app.route('/api/hy-delete-schedule', methods=['POST'])
def delete_schedule():
    """删除指定排期文件"""
    filename = (request.json or {}).get('filename', '')
    if not filename:
        return jsonify({'error': '缺少文件名'}), 400
    filepath = _safe_path_under(SCHEDULE_DIR, filename)
    if not filepath:
        return jsonify({'error': '非法路径'}), 403
    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在'}), 404
    os.remove(filepath)
    # 增量移除：从映射表中清除该文件的条目（不全量重扫）
    _remove_file_from_map(filename)
    info = _get_schedule_info()
    logger.info(f'[河源] 删除排期文件(增量): {filename}')
    return jsonify({
        'ok': True,
        'msg': f'已删除 {filename}',
        'file_count': info['file_count'],
        'item_count': info['item_count'],
    })


@app.route('/api/hy-delete-all-schedules', methods=['POST'])
def delete_all_schedules():
    """一键清空所有排期文件"""
    # 先杀掉正在运行的scan子进程（释放文件锁）
    if os.name == 'nt':
        import signal
        try:
            out = subprocess.run(
                ['tasklist', '/FI', 'IMAGENAME eq python.exe', '/FO', 'CSV', '/NH'],
                capture_output=True, text=True, timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW
            ).stdout
            my_pid = os.getpid()
            for line in out.strip().split('\n'):
                if 'python' in line.lower():
                    parts = line.strip('"').split('","')
                    if len(parts) >= 2:
                        pid = int(parts[1])
                        if pid != my_pid:
                            try:
                                os.kill(pid, signal.SIGTERM)
                            except (ProcessLookupError, PermissionError):
                                pass
        except Exception:
            pass
        time.sleep(1)

    deleted = []
    failed = []
    for f in os.listdir(SCHEDULE_DIR):
        if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
            fpath = os.path.join(SCHEDULE_DIR, f)
            try:
                os.remove(fpath)
                deleted.append(f)
            except PermissionError:
                failed.append((f, fpath))

    # 重试被锁的文件（更长等待）
    if failed:
        time.sleep(3)
        still_failed = []
        for f, fpath in failed:
            try:
                os.remove(fpath)
                deleted.append(f)
            except Exception:
                still_failed.append(f)
        failed = still_failed

    # 清空映射表
    map_path = os.path.join(DATA_DIR, 'hy_item_map.json')
    if not failed:
        with open(map_path, 'w', encoding='utf-8') as fp:
            json.dump({}, fp)
    else:
        _run_scan()

    info = _get_schedule_info()
    logger.info(f'[河源] 一键清空排期: 删除{len(deleted)}个, 剩余{info["file_count"]}个')
    msg = f'已清空{len(deleted)}个排期文件'
    if failed:
        msg += f'（{len(failed)}个文件被占用，请稍后重试）'
    return jsonify({'ok': True, 'msg': msg, 'deleted': len(deleted),
                    'file_count': info['file_count'], 'item_count': info['item_count']})


@app.route('/api/hy-rescan', methods=['POST'])
def hy_rescan():
    """手动重建映射表"""
    ok, stdout, stderr = _run_scan()
    info = _get_schedule_info()
    return jsonify({
        'ok': ok,
        'item_count': info['item_count'],
        'stdout': stdout,
        'stderr': stderr,
    })


# ── PO分析 ──

@app.route('/api/hy-upload', methods=['POST'])
def hy_upload():
    """上传PO Excel，分析得到修改单/新单/重复货号"""
    _cleanup_pending()

    info = _get_schedule_info()
    if info['file_count'] == 0:
        return jsonify({'error': '请先上传排期文件（点击上方"管理排期文件"按钮）'}), 400

    saved = []
    for f in request.files.getlist('files'):
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ('.xlsx', '.xls'):
            continue
        safe_name = secure_filename(f.filename) or f'upload_{int(time.time())}{ext}'
        if not safe_name.endswith(ext):
            safe_name += ext
        path = os.path.join(UPLOAD_DIR, safe_name)
        f.save(path)
        saved.append((f.filename, path))

    if not saved:
        return jsonify({'error': '没有有效的Excel文件'}), 400

    orders = []
    errors = []
    for fname, path in saved:
        try:
            data = ExcelPOParser().parse(path)
            data['filename'] = fname
            orders.append(data)
        except Exception as e:
            errors.append(f'{fname}: {str(e)[:80]}')

    if not orders:
        return jsonify({'error': f'解析失败: {"; ".join(errors)}'}), 400

    try:
        analysis = analyze_orders(SCHEDULE_DIR, orders)
    except Exception:
        logger.exception('[河源] 分析失败')
        return jsonify({'error': '排期分析失败，请稍后重试或联系管理员'}), 500

    session_id = secrets.token_urlsafe(16)
    _save_session(session_id, {
        'orders': orders,
        'analysis': analysis,
        'timestamp': datetime.now(),
    })

    amb_for_ui = []
    for amb in analysis['ambiguous']:
        amb_for_ui.append({
            'key': f"{amb['order_idx']}_{amb['line_idx']}",
            'item': amb['item'],
            'po': amb['po'],
            'candidates': amb['candidates'],
        })

    return jsonify({
        'ok': True,
        'need_selection': True,
        'session_id': session_id,
        'ambiguous': amb_for_ui,
        'modifications_count': len(analysis['modifications']),
        'new_count': len(analysis['new_lines']),
        'unknown_count': len(analysis['unknown']),
        'errors': errors,
    })


@app.route('/api/hy-submit-selection', methods=['POST'])
def hy_submit_selection():
    """提交重复货号选择，生成分类Excel（云端不做COM写入）"""
    _cleanup_pending()
    data = request.json or {}
    session_id = data.get('session_id')
    selections = data.get('selections', {})

    session = _load_session(session_id)
    if not session:
        return jsonify({'error': '会话已过期，请重新上传'}), 400

    orders = session['orders']
    analysis = session['analysis']

    valid_map = {}
    for amb in analysis['ambiguous']:
        key = f"{amb['order_idx']}_{amb['line_idx']}"
        valid_map[key] = {(c['file'], c['sheet']) for c in amb['candidates']}

    for key, sel in selections.items():
        if key not in valid_map:
            return jsonify({'error': f'非法选择key: {key}'}), 400
        if not isinstance(sel, dict) or 'file' not in sel or 'sheet' not in sel:
            return jsonify({'error': f'选择数据格式错误: {key}'}), 400
        if (sel['file'], sel['sheet']) not in valid_map[key]:
            return jsonify({'error': f'选择的文件/Sheet不在候选范围内: {key}'}), 400

    try:
        result = write_orders(SCHEDULE_DIR, orders,
                              ambiguous_selections=selections,
                              export_dir=EXPORT_DIR)
    except Exception:
        logger.exception('[河源] 提交写入失败')
        return jsonify({'error': '处理失败，请稍后重试或联系管理员'}), 500
    finally:
        _delete_session(session_id)

    resp = {
        'ok': True, 'msg': result['msg'],
        'modified': result.get('modified', 0),
        'new_count': result.get('new_count', 0),
        'mod_details': result.get('mod_details', []),
        'new_details': result.get('new_details', []),
        'unknown': result.get('unknown', []),
    }
    if result.get('export_file'):
        resp['export_file'] = result['export_file']
    return jsonify(resp)


@app.route('/api/hy-export-only', methods=['POST'])
def hy_export_only():
    """仅生成分类Excel，不写入排期。ambiguous货号放入所有匹配sheet。"""
    from hy_schedule import _extract_header, _prepare_line_data, \
        _item_upper, _generate_new_excel

    data = request.json or {}
    session_id = data.get('session_id')

    if not session_id:
        return jsonify({'error': '缺少session_id'}), 400

    session = _load_session(session_id)
    if not session:
        return jsonify({'error': '会话已过期，请重新上传'}), 400

    orders = session['orders']
    try:
        analysis = analyze_orders(SCHEDULE_DIR, orders)
    except Exception:
        logger.exception('[河源] 分析失败(export-only)')
        return jsonify({'error': '排期分析失败，请稍后重试或联系管理员'}), 500

    import re
    for amb in analysis['ambiguous']:
        for cand in amb['candidates']:
            analysis['new_lines'].append({
                'order_idx': amb['order_idx'],
                'line_idx': amb['line_idx'],
                'item': amb['item'], 'po': amb['po'],
                'file': cand['file'], 'sheet': cand['sheet'],
                'type': cand.get('type', 'hy'),
            })

    cn_names = dict(analysis.get('cn_names', {}))
    try:
        cn_map_path = os.path.join(DATA_DIR, 'item_cn_name_map.json')
        if os.path.exists(cn_map_path):
            with open(cn_map_path, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            for k, v in raw.items():
                if not k.startswith('_') and k.upper() not in cn_names:
                    cn_names[k.upper()] = v.get('cn_name', '') if isinstance(v, dict) else str(v)
    except Exception:
        pass

    new_rows = []
    for nl in analysis['new_lines']:
        oi, li = nl['order_idx'], nl['line_idx']
        order = orders[oi]
        hdr = _extract_header(order)
        ln_data = _prepare_line_data(order, order['lines'][li], hdr['ship_dt'], hdr['full_note'],
                                     wb_name=nl['file'])
        item_base = re.match(r'(\d+[A-Za-z]*\d*)', _item_upper(nl['item']))
        cn_name = cn_names.get(item_base.group(1).upper(), '') if item_base else ''
        _has_outer = bool(ln_data['outer_qty'])
        _has_price = bool(ln_data['price'])
        new_rows.append({
            'target_file': nl['file'], 'target_sheet': nl['sheet'],
            'type': nl.get('type', 'hy'),
            'po_date': hdr['po_date_dt'], 'customer': hdr['customer'], 'dest': hdr['dest'],
            'po': hdr['po'], 'cpo': ln_data['customer_po'], 'sku_line': ln_data['f_sku'],
            'item': ln_data['sku_spec'], 'cn_name': cn_name,
            'qty': ln_data['qty'], 'inner': ln_data['inner_pcs'],
            'outer': ln_data['outer_qty'],
            'total_box': '__FORMULA_TOTAL_BOX__' if _has_outer else '',
            'ship_date': ln_data['line_ship_dt'], 'insp_date': ln_data['insp_dt'],
            'remark': hdr['full_note'],
            'from_person': hdr['from_person'].split('/')[0].strip() if hdr['from_person'] else '',
            'price': round(ln_data['price'], 4) if _has_price else '',
            'amount': '__FORMULA_AMOUNT__' if _has_price else '',
        })

    for m in analysis['modifications']:
        oi, li = m['order_idx'], m['line_idx']
        order = orders[oi]
        hdr = _extract_header(order)
        ln_data = _prepare_line_data(order, order['lines'][li], hdr['ship_dt'], hdr['full_note'],
                                     wb_name=m['file'])
        item_base = re.match(r'(\d+[A-Za-z]*\d*)', _item_upper(m['item']))
        cn_name = cn_names.get(item_base.group(1).upper(), '') if item_base else ''
        _has_outer = bool(ln_data['outer_qty'])
        _has_price = bool(ln_data['price'])
        new_rows.append({
            'target_file': '修改单', 'target_sheet': '',
            'po_date': hdr['po_date_dt'], 'customer': hdr['customer'], 'dest': hdr['dest'],
            'po': hdr['po'], 'cpo': ln_data['customer_po'], 'sku_line': ln_data['f_sku'],
            'item': ln_data['sku_spec'], 'cn_name': cn_name,
            'qty': ln_data['qty'], 'inner': ln_data['inner_pcs'],
            'outer': ln_data['outer_qty'],
            'total_box': '__FORMULA_TOTAL_BOX__' if _has_outer else '',
            'ship_date': ln_data['line_ship_dt'], 'insp_date': ln_data['insp_dt'],
            'remark': f"[修改] 目标:{m['file']}/{m['sheet']} R{m['row']}",
            'from_person': hdr['from_person'].split('/')[0].strip() if hdr['from_person'] else '',
            'price': round(ln_data['price'], 4) if _has_price else '',
            'amount': '__FORMULA_AMOUNT__' if _has_price else '',
        })

    for uk in analysis['unknown']:
        oi, li = uk['order_idx'], uk['line_idx']
        order = orders[oi]
        hdr = _extract_header(order)
        ln_data = _prepare_line_data(order, order['lines'][li], hdr['ship_dt'], hdr['full_note'],
                                     wb_name='')
        item_base = re.match(r'(\d+[A-Za-z]*\d*)', _item_upper(uk['item']))
        cn_name = cn_names.get(item_base.group(1).upper(), '') if item_base else ''
        _has_outer = bool(ln_data['outer_qty'])
        _has_price = bool(ln_data['price'])
        new_rows.append({
            'target_file': '未识别货号', 'target_sheet': '',
            'po_date': hdr['po_date_dt'], 'customer': hdr['customer'], 'dest': hdr['dest'],
            'po': hdr['po'], 'cpo': ln_data['customer_po'], 'sku_line': ln_data['f_sku'],
            'item': ln_data['sku_spec'], 'cn_name': cn_name,
            'qty': ln_data['qty'], 'inner': ln_data['inner_pcs'],
            'outer': ln_data['outer_qty'],
            'total_box': '__FORMULA_TOTAL_BOX__' if _has_outer else '',
            'ship_date': ln_data['line_ship_dt'], 'insp_date': ln_data['insp_dt'],
            'remark': hdr['full_note'],
            'from_person': hdr['from_person'].split('/')[0].strip() if hdr['from_person'] else '',
            'price': round(ln_data['price'], 4) if _has_price else '',
            'amount': '__FORMULA_AMOUNT__' if _has_price else '',
        })

    if not new_rows:
        return jsonify({'error': '没有可导出的数据'}), 400

    # 按排期类型分流导出
    hy_rows = [r for r in new_rows if r.get('type', 'hy') != 'fy']
    fy_rows_export = [r for r in new_rows if r.get('type') == 'fy']

    export_file = None
    fy_export_file = None

    if hy_rows:
        export_file = _generate_new_excel(hy_rows, EXPORT_DIR, schedule_dir=SCHEDULE_DIR)
    if fy_rows_export:
        fy_export_file = generate_fy_export(fy_rows_export, EXPORT_DIR, SCHEDULE_DIR)

    if not export_file and fy_export_file:
        export_file = fy_export_file
        fy_export_file = None

    if not export_file:
        return jsonify({'error': '生成Excel失败'}), 500

    _delete_session(session_id)

    msg_parts = []
    if hy_rows:
        msg_parts.append(f'河源{len(hy_rows)}行')
    if fy_rows_export:
        msg_parts.append(f'翻译{len(fy_rows_export)}行')
    msg = f'已生成分类Excel（{"+".join(msg_parts) or str(len(new_rows))+"行"}）'

    logger.info(f'[导出] {msg}, 文件={export_file}, fy={fy_export_file}')
    resp = {
        'ok': True,
        'msg': msg,
        'export_file': export_file,
        'new_count': len(analysis['new_lines']),
        'mod_count': len(analysis['modifications']),
        'unknown_count': len(analysis['unknown']),
        'ambiguous_count': len(analysis['ambiguous']),
    }
    if fy_export_file:
        resp['fy_export_file'] = fy_export_file
    return jsonify(resp)


@app.route('/api/hy-export-download/<filename>')
def hy_export_download(filename):
    filepath = _safe_path_under(EXPORT_DIR, filename)
    if not filepath:
        return jsonify({'error': '非法路径'}), 403
    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(filepath, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    port = int(os.environ.get('APP_PORT', 5008))
    print('=' * 50)
    print('  ZURU 河源排期入单系统（云端版）')
    print(f'  http://localhost:{port}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=os.environ.get('FLASK_DEBUG') == '1', use_reloader=False, threaded=True)
