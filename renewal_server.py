"""
FY26 Q3 續約總表管理系統 - Flask 後端
"""
from flask import Flask, request, jsonify, send_from_directory, session, Response
import sqlite3, re, os, json, hashlib, queue, threading
from datetime import datetime, date, timedelta

# Excel date serial → 'YYYY/MM/DD'
_DATE_KEYWORDS = ('日', 'date', 'Date', 'DATE', '日期', 'start', 'end', 'Start', 'End', '提醒', '週年', '周年')
_EXCEL_ORIGIN  = date(1899, 12, 30)

def maybe_excel_date(val_str, col_label):
    """Convert Excel date serial to readable string if applicable."""
    if not val_str:
        return val_str
    label = col_label or ''
    if not any(kw in label for kw in _DATE_KEYWORDS):
        return val_str
    try:
        n = float(val_str)
        if n != int(n):  # has decimals → datetime, ignore time part
            n = int(n)
        if 36526 <= n <= 54789:  # 2000-01-01 to 2049-12-31
            return (_EXCEL_ORIGIN + timedelta(days=int(n))).strftime('%Y/%m/%d')
    except (ValueError, TypeError):
        pass
    return val_str

# ─── SSE broadcast ────────────────────────────────────────────────────────────
_sse_clients = []
_sse_lock = threading.Lock()

def _sse_broadcast(msg: dict):
    payload = json.dumps(msg, ensure_ascii=False)
    with _sse_lock:
        dead = []
        for q in _sse_clients:
            try:
                q.put_nowait(payload)
            except queue.Full:
                dead.append(q)
        for q in dead:
            _sse_clients.remove(q)

app = Flask(__name__, static_folder='.')
app.config['JSON_SORT_KEYS'] = False
app.secret_key = 'fy26_renewal_2026_secret'
try:
    app.json.sort_keys = False
except AttributeError:
    pass
DB_PATH = os.path.join(os.path.dirname(__file__), 'renewal_data.db')
PERM_XLSX = os.path.join(os.path.dirname(__file__), '續約總表_權限表_20260526(1).xlsx')
GROUPS = ['Supervisor', 'Renew Team', 'Call out Team', 'BD Team', 'Adobe 原廠']

# ─── DB helpers ───────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def sanitize(name):
    if not name:
        return ''
    s = re.sub(r'[^\w一-鿿]', '_', str(name).strip())
    s = re.sub(r'_+', '_', s).strip('_')
    return s or 'col'

def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS import_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT, sheet TEXT, imported_at TEXT, row_count INTEGER
        );
        CREATE TABLE IF NOT EXISTS col_meta (
            table_name TEXT, col_index INTEGER, col_key TEXT, col_label TEXT,
            PRIMARY KEY (table_name, col_index)
        );
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, email TEXT, username TEXT UNIQUE,
            password TEXT, group_name TEXT, status TEXT DEFAULT '啟用', notes TEXT,
            last_login TEXT, last_logout TEXT
        );
        CREATE TABLE IF NOT EXISTS field_permissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_name TEXT, group_name TEXT, has_access INTEGER DEFAULT 1,
            UNIQUE(field_name, group_name)
        );
    """)
    conn.commit()
    # Seed from Excel if tables empty
    if conn.execute('SELECT COUNT(*) FROM accounts').fetchone()[0] == 0:
        _seed_from_excel(conn)
    # Seed tab/func permissions if not present
    _seed_tab_perms(conn)
    _seed_edit_perms(conn)
    conn.close()

# From the permission image:
# Columns: Call out Team, BD Team, Adobe 原廠, Renew Team, Supervisor
_TAB_PERMS = {
    '_tab_dash':       [0, 0, 1, 1, 1],
    '_tab_main':       [1, 1, 1, 1, 1],
    '_tab_rbob':       [0, 0, 1, 1, 1],
    '_tab_keyaccount': [0, 0, 1, 1, 1],
    '_tab_tongji':     [0, 0, 1, 1, 1],
    '_tab_close':      [0, 0, 1, 1, 1],
    '_tab_reseller':   [0, 0, 1, 1, 1],
    '_tab_admin':      [0, 0, 0, 0, 1],
    '_func_add_row':   [0, 0, 0, 1, 1],
}
_TAB_GROUPS = ['Call out Team', 'BD Team', 'Adobe 原廠', 'Renew Team', 'Supervisor']

# Default editable column fields (all groups) — stored as _edit_<field> in field_permissions
_EDIT_DEFAULTS = {
    'Close_Unit', 'Close Unit', '進度', '更新日期', '連絡狀況',
    'Attrition_Status', 'Call Name', 'Lost Reason',
    '授權管理人 (FY25)', '電話 (FY25)', 'Email (FY25)',
    '主要管理人 (FY26)', '授權管理人 (FY26)', '電話 (FY26)', 'Email (FY26)',
}

def _seed_edit_perms(conn):
    """Create/update _edit_<field> rows for every non-tab/non-func column field."""
    existing = {r[0] for r in conn.execute(
        "SELECT DISTINCT field_name FROM field_permissions"
        " WHERE substr(field_name,1,5) NOT IN ('_tab_','_func') AND field_name NOT LIKE '\\_edit\\_%' ESCAPE '\\'"
        " AND field_name NOT IN ('匯入 Excel','匯入 CSV')"
    ).fetchall()}
    for fname in existing:
        edit_key = '_edit_' + fname
        default_has = 1 if fname in _EDIT_DEFAULTS else 0
        for g in _TAB_GROUPS:
            has = 1 if g == 'Supervisor' else default_has
            conn.execute(
                'INSERT OR IGNORE INTO field_permissions(field_name,group_name,has_access) VALUES(?,?,?)',
                (edit_key, g, has)
            )
    # Supervisor must always be able to edit all fields — fix any rows seeded as 0
    conn.execute(
        "UPDATE field_permissions SET has_access=1"
        " WHERE substr(field_name,1,6)='_edit_' AND group_name='Supervisor'"
    )
    conn.commit()

def _seed_tab_perms(conn):
    for field, vals in _TAB_PERMS.items():
        for g, v in zip(_TAB_GROUPS, vals):
            # INSERT OR IGNORE so manual changes by Supervisor are preserved
            conn.execute(
                'INSERT OR IGNORE INTO field_permissions(field_name,group_name,has_access) VALUES(?,?,?)',
                (field, g, v)
            )
    conn.commit()
    # Ensure _func_add_row exists (for servers already running without it)
    for g, v in zip(_TAB_GROUPS, _TAB_PERMS['_func_add_row']):
        conn.execute(
            'INSERT OR IGNORE INTO field_permissions(field_name,group_name,has_access) VALUES(?,?,?)',
            ('_func_add_row', g, v)
        )
    conn.commit()

def _seed_from_excel(conn):
    try:
        import pandas as pd
        xl = pd.ExcelFile(PERM_XLSX)
        # 帳號明細
        df = pd.read_excel(xl, sheet_name='帳號明細')
        df.columns = ['name','email','username','password','group_name','status','notes']
        for _, r in df.iterrows():
            conn.execute(
                'INSERT OR IGNORE INTO accounts(name,email,username,password,group_name,status,notes) VALUES(?,?,?,?,?,?,?)',
                (str(r['name']).strip(), str(r['email']).strip(), str(r['username']).strip(),
                 hash_pw(str(r['password']).strip()), str(r['group_name']).strip(),
                 str(r['status']).strip() if str(r['status']) != 'nan' else '啟用',
                 str(r['notes']).strip() if str(r['notes']) != 'nan' else '')
            )
        # 權限說明表
        df2 = pd.read_excel(xl, sheet_name='權限說明表', header=None)
        # Find header row (欄位名稱)
        hdr_idx = None
        for i, row in df2.iterrows():
            if str(row.iloc[0]).strip() == '欄位名稱':
                hdr_idx = i; break
        if hdr_idx is not None:
            df2.columns = df2.iloc[hdr_idx]
            df2 = df2.iloc[hdr_idx+1:].reset_index(drop=True)
            groups_in_sheet = [c for c in df2.columns if c and str(c) != 'nan' and c != '欄位名稱']
            for _, row in df2.iterrows():
                field = str(row['欄位名稱']).strip() if str(row['欄位名稱']) != 'nan' else ''
                if not field: continue
                for g in groups_in_sheet:
                    val = str(row.get(g, '')).strip()
                    has = 0 if '❌' in val else 1
                    # Map sheet group names to DB group names
                    gmap = {'Adobe': 'Adobe 原廠'}
                    gname = gmap.get(g, g)
                    conn.execute(
                        'INSERT OR REPLACE INTO field_permissions(field_name,group_name,has_access) VALUES(?,?,?)',
                        (field, gname, has)
                    )
        conn.commit()
    except Exception as e:
        print(f'Seed warning: {e}')

init_db()

# ─── Import helpers ───────────────────────────────────────────────────────────
def import_sheet_rows(conn, rows_iter, table_name, header_row_idx, data_start_idx):
    """Import from an iterable of row-tuples (values only)."""
    rows_buf = []
    for i, row in enumerate(rows_iter):
        rows_buf.append(row)
        if i >= header_row_idx:  # we have the header, stop buffering early
            break

    if len(rows_buf) < header_row_idx + 1:
        return 0

    raw_headers = [str(v).strip() if v is not None else '' for v in rows_buf[header_row_idx - 1]]

    # Trim trailing empty cols
    while raw_headers and not raw_headers[-1]:
        raw_headers.pop()
    if not raw_headers:
        return 0

    # Build unique sanitized keys
    keys, seen = [], {}
    for i, h in enumerate(raw_headers):
        k = sanitize(h) or f'col_{i}'
        if k in seen:
            seen[k] += 1
            k = f'{k}_{seen[k]}'
        else:
            seen[k] = 0
        keys.append(k)

    # (Re)create table
    conn.execute(f'DROP TABLE IF EXISTS "{table_name}"')
    col_defs = ', '.join([f'"{k}" TEXT' for k in keys])
    conn.execute(f'CREATE TABLE "{table_name}" (id INTEGER PRIMARY KEY AUTOINCREMENT, {col_defs})')

    # Store column metadata
    conn.execute('DELETE FROM col_meta WHERE table_name=?', (table_name,))
    conn.executemany(
        'INSERT INTO col_meta VALUES (?,?,?,?)',
        [(table_name, i, keys[i], raw_headers[i]) for i in range(len(keys))]
    )

    placeholders = ','.join(['?' for _ in keys])
    col_str = ','.join([f'"{k}"' for k in keys])

    def insert_row(row_vals):
        vals, has_data = [], False
        for i in range(len(keys)):
            v = row_vals[i] if i < len(row_vals) else None
            s = str(v).strip() if v is not None else None
            if s and s not in ('None', 'nan'):
                has_data = True
                s = maybe_excel_date(s, raw_headers[i] if i < len(raw_headers) else '')
            else:
                s = None
            vals.append(s)
        if not has_data:
            return False
        conn.execute(f'INSERT INTO "{table_name}" ({col_str}) VALUES ({placeholders})', vals)
        return True

    n = 0
    # Insert rows already buffered after header
    for row in rows_buf[data_start_idx - 1:]:
        if insert_row(row):
            n += 1

    # Continue from iterator
    for row in rows_iter:
        if insert_row(row):
            n += 1

    conn.commit()
    return n


def import_sheet(conn, ws, table_name, header_row, data_start_row):
    """Wrap a worksheet into the row-based importer (read_only compatible)."""
    def row_gen():
        for row in ws.iter_rows(values_only=True):
            yield row
    return import_sheet_rows(conn, row_gen(), table_name, header_row, data_start_row)

# ─── Routes ───────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('.', 'renewal_app.html')

@app.route('/api/import', methods=['POST'])
def api_import():
    """Fast import using zipfile + ElementTree (avoids slow openpyxl full parse)."""
    import zipfile, xml.etree.ElementTree as ET

    if 'file' in request.files and request.files['file'].filename:
        f = request.files['file']
        tmp_path = os.path.join(os.path.dirname(__file__), '__tmp_import.xlsx')
        f.save(tmp_path)
        filename = f.filename
        cleanup = True
    else:
        tmp_path = os.path.join(os.path.dirname(__file__), 'FY26_Q3_續約總表_範例.xlsx')
        filename = os.path.basename(tmp_path)
        cleanup = False
        if not os.path.exists(tmp_path):
            return jsonify(error='找不到預設 Excel 檔案'), 404

    try:
        conn = get_db()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        results = {}

        zf = zipfile.ZipFile(tmp_path, 'r')

        # ── shared strings ──────────────────────────────────────────────────
        NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        ss_root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
        shared = []
        for si in ss_root.iter(f'{{{NS}}}si'):
            parts = [t.text or '' for t in si.iter(f'{{{NS}}}t')]
            shared.append(''.join(parts))

        # ── sheet name → file path map ──────────────────────────────────────
        wb_root = ET.fromstring(zf.read('xl/workbook.xml'))
        RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        sheet_rid = {}
        for s in wb_root.iter(f'{{{NS}}}sheet'):
            sheet_rid[s.get('name')] = s.get(f'{{{RNS}}}id')

        rels_root = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
        rid_path = {r.get('Id'): 'xl/' + r.get('Target')
                    for r in rels_root.iter(f'{{{PNS}}}Relationship')}

        def col_letters_to_idx(s):
            idx = 0
            for ch in s:
                idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
            return idx - 1  # 0-based

        def sheet_rows(sheet_name):
            """Stream rows from a worksheet using iterparse (memory efficient)."""
            rid = sheet_rid.get(sheet_name)
            if not rid:
                return
            path = rid_path.get(rid)
            if not path:
                return
            full_path = path  # rid_path already has 'xl/' prefix
            # Track current row cells
            cur_cells = {}
            cur_t = None
            cur_ref = None
            cur_in_v = False
            cur_v_text = ''

            with zf.open(full_path) as fh:
                for event, elem in ET.iterparse(fh, events=('start', 'end')):
                    tag = elem.tag.split('}', 1)[-1] if '}' in elem.tag else elem.tag

                    if event == 'start':
                        if tag == 'row':
                            cur_cells = {}
                        elif tag == 'c':
                            cur_ref = elem.get('r', '')
                            cur_t   = elem.get('t', '')
                            cur_v_text = ''
                            cur_in_v = False
                        elif tag == 'v':
                            cur_in_v = True
                            cur_v_text = elem.text or ''
                    elif event == 'end':
                        if tag == 'v' and cur_in_v:
                            cur_in_v = False
                        elif tag == 'c':
                            val = None
                            if cur_v_text:
                                if cur_t == 's':
                                    try:
                                        val = shared[int(cur_v_text)]
                                    except Exception:
                                        val = cur_v_text
                                else:
                                    val = cur_v_text
                            col_letters = ''.join(ch for ch in cur_ref if ch.isalpha())
                            if col_letters:
                                cur_cells[col_letters_to_idx(col_letters)] = val
                            elem.clear()
                        elif tag == 'row':
                            if cur_cells:
                                max_col = max(cur_cells.keys())
                                yield [cur_cells.get(i) for i in range(max_col + 1)]
                            elem.clear()

        def do_import(sheet_name, table_name, header_row, data_start_row):
            if sheet_name not in sheet_rid:
                return None
            row_iter = sheet_rows(sheet_name)
            n = import_sheet_rows(conn, row_iter, table_name, header_row, data_start_row)
            conn.execute(
                'INSERT INTO import_log (filename,sheet,imported_at,row_count) VALUES (?,?,?,?)',
                (filename, sheet_name, now, n)
            )
            conn.commit()
            return n

        # ── Import each sheet ───────────────────────────────────────────────
        # 總表(New): row1=group label, row2=headers, row3+=data
        n = do_import('總表(New)', 'main_table', 2, 3)
        if n is not None:
            results['總表(New)'] = n

        n = do_import('RBOB', 'rbob', 1, 2)
        if n is not None:
            results['RBOB'] = n

        n = do_import('Key account', 'key_account', 1, 2)
        if n is not None:
            results['Key account'] = n

        n = do_import('統整報表', 'tongji', 1, 2)
        if n is not None:
            results['統整報表'] = n

        n = do_import('close summary', 'close_summary', 1, 2)
        if n is not None:
            results['close summary'] = n

        n = do_import('Reseller Owner', 'reseller_owner', 1, 2)
        if n is not None:
            results['Reseller Owner'] = n

        zf.close()
        conn.close()

        if cleanup and os.path.exists(tmp_path):
            os.remove(tmp_path)

        return jsonify(ok=True, filename=filename, sheets=results, imported_at=now)

    except Exception as e:
        import traceback
        return jsonify(error=str(e), detail=traceback.format_exc()), 500


@app.route('/api/columns/<table>')
def api_columns(table):
    ALLOWED = {'main_table','rbob','key_account','tongji','close_summary','reseller_owner'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400
    conn = get_db()
    rows = conn.execute(
        'SELECT col_index, col_key, col_label FROM col_meta WHERE table_name=? ORDER BY col_index',
        (table,)
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/data/<table>')
def api_data(table):
    ALLOWED = {'main_table','rbob','key_account','tongji','close_summary','reseller_owner'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400

    page  = max(1, int(request.args.get('page', 1)))
    limit = min(500, int(request.args.get('limit', 100)))
    q     = request.args.get('q', '').strip()
    col   = request.args.get('col', '')   # filter by specific col
    val   = request.args.get('val', '')

    offset = (page - 1) * limit

    conn = get_db()

    # Check table exists
    exists = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone()
    if not exists:
        conn.close()
        return jsonify(rows=[], total=0, page=page, limit=limit)

    where_clauses, params = [], []

    if q:
        # Full text search across all text columns
        cols_info = conn.execute(
            'SELECT col_key FROM col_meta WHERE table_name=?', (table,)
        ).fetchall()
        col_keys = [r['col_key'] for r in cols_info]
        if col_keys:
            like_parts = ' OR '.join([f'"{k}" LIKE ?' for k in col_keys])
            where_clauses.append(f'({like_parts})')
            params.extend([f'%{q}%' for _ in col_keys])

    if col and val:
        where_clauses.append(f'"{sanitize(col)}" = ?')
        params.append(val)

    # Multi-column filters via f_ prefix params
    for k, v in request.args.items():
        if k.startswith('f_') and v:
            col_name = sanitize(k[2:])
            if col_name:
                where_clauses.append(f'"{col_name}" = ?')
                params.append(v)

    where_sql = f'WHERE {" AND ".join(where_clauses)}' if where_clauses else ''

    total = conn.execute(f'SELECT COUNT(*) FROM "{table}" {where_sql}', params).fetchone()[0]

    # Get ordered column keys from col_meta
    ordered_keys = [r['col_key'] for r in conn.execute(
        'SELECT col_key FROM col_meta WHERE table_name=? ORDER BY col_index', (table,)
    ).fetchall()]

    rows = conn.execute(
        f'SELECT rowid as _rid, * FROM "{table}" {where_sql} LIMIT ? OFFSET ?',
        params + [limit, offset]
    ).fetchall()

    def ordered_row(r):
        from collections import OrderedDict
        d = OrderedDict()
        d['_rid'] = r['_rid']
        for k in ordered_keys:
            d[k] = r[k] if k in r.keys() else None
        return d

    conn.close()
    return jsonify(rows=[ordered_row(r) for r in rows], total=total, page=page, limit=limit)


@app.route('/api/stats')
def api_stats():
    conn = get_db()
    stats = {}

    def tbl_exists(t):
        return conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (t,)
        ).fetchone()

    if tbl_exists('main_table'):
        stats['total'] = conn.execute('SELECT COUNT(*) FROM main_table').fetchone()[0]

        # 進度分佈
        rows = conn.execute(
            'SELECT "進度" AS s, COUNT(*) AS n FROM main_table WHERE "進度" IS NOT NULL GROUP BY "進度" ORDER BY n DESC'
        ).fetchall()
        stats['by_status'] = [dict(r) for r in rows]

        # Owner分佈
        rows = conn.execute(
            'SELECT "Owner" AS s, COUNT(*) AS n FROM main_table WHERE "Owner" IS NOT NULL GROUP BY "Owner" ORDER BY n DESC LIMIT 15'
        ).fetchall()
        stats['by_owner'] = [dict(r) for r in rows]

        # 部門別
        rows = conn.execute(
            'SELECT "部門別" AS s, COUNT(*) AS n FROM main_table WHERE "部門別" IS NOT NULL GROUP BY "部門別" ORDER BY n DESC'
        ).fetchall()
        stats['by_dept'] = [dict(r) for r in rows]

        # 市場別
        rows = conn.execute(
            'SELECT "市場別" AS s, COUNT(*) AS n FROM main_table WHERE "市場別" IS NOT NULL GROUP BY "市場別" ORDER BY n DESC'
        ).fetchall()
        stats['by_market'] = [dict(r) for r in rows]

        # 經銷商 top 15
        rows = conn.execute(
            'SELECT "經銷商" AS s, COUNT(*) AS n FROM main_table WHERE "經銷商" IS NOT NULL GROUP BY "經銷商" ORDER BY n DESC LIMIT 15'
        ).fetchall()
        stats['by_reseller'] = [dict(r) for r in rows]

        # Attrition status
        rows = conn.execute(
            'SELECT "Attrition_Status" AS s, COUNT(*) AS n FROM main_table WHERE "Attrition_Status" IS NOT NULL GROUP BY "Attrition_Status" ORDER BY n DESC'
        ).fetchall()
        stats['by_attrition'] = [dict(r) for r in rows]

        # Close ARR sum (try column name variations)
        for col in ['Close_ARR', 'Close_Unit', 'RBOB']:
            try:
                r = conn.execute(
                    f'SELECT SUM(CAST(REPLACE("{col}",",","") AS REAL)) FROM main_table WHERE "{col}" IS NOT NULL'
                ).fetchone()
                if r and r[0] is not None:
                    stats[f'sum_{col.lower()}'] = round(r[0], 2)
            except Exception:
                pass

    if tbl_exists('key_account'):
        stats['key_account_total'] = conn.execute('SELECT COUNT(*) FROM key_account').fetchone()[0]

    if tbl_exists('rbob'):
        stats['rbob_total'] = conn.execute('SELECT COUNT(*) FROM rbob').fetchone()[0]

    if tbl_exists('tongji'):
        stats['tongji_total'] = conn.execute('SELECT COUNT(*) FROM tongji').fetchone()[0]

    # Last import
    last = conn.execute(
        'SELECT filename, imported_at FROM import_log ORDER BY id DESC LIMIT 1'
    ).fetchone()
    if last:
        stats['last_import'] = dict(last)

    conn.close()
    return jsonify(stats)


@app.route('/api/distinct/<table>/<col>')
def api_distinct(table, col):
    ALLOWED = {'main_table','rbob','key_account','tongji','close_summary','reseller_owner'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400
    col = sanitize(col)
    conn = get_db()
    try:
        rows = conn.execute(
            f'SELECT DISTINCT "{col}" FROM "{table}" WHERE "{col}" IS NOT NULL ORDER BY "{col}" LIMIT 300'
        ).fetchall()
        return jsonify([r[0] for r in rows])
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        conn.close()


@app.route('/api/row/<table>', methods=['POST'])
def api_row_create(table):
    ALLOWED = {'main_table'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400
    data = request.get_json() or {}
    conn = get_db()
    try:
        cols = [r['col_key'] for r in conn.execute(
            'SELECT col_key FROM col_meta WHERE table_name=? ORDER BY col_index', (table,)
        ).fetchall()]
        col_set = set(cols)
        valid = {sanitize(k): v for k, v in data.items() if sanitize(k) in col_set}
        if not valid:
            return jsonify(error='no valid columns'), 400
        col_str = ', '.join(f'"{k}"' for k in valid)
        placeholders = ', '.join('?' for _ in valid)
        conn.execute(f'INSERT INTO "{table}" ({col_str}) VALUES ({placeholders})', list(valid.values()))
        conn.commit()
        rid = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
        return jsonify(ok=True, rowid=rid)
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        conn.close()


@app.route('/api/row/<table>/<int:row_id>', methods=['PUT', 'DELETE'])
def api_row(table, row_id):
    ALLOWED = {'main_table'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400
    conn = get_db()
    try:
        if request.method == 'DELETE':
            conn.execute(f'DELETE FROM "{table}" WHERE rowid=?', (row_id,))
            conn.commit()
            return jsonify(ok=True)
        data = request.get_json() or {}
        cols = [r['col_key'] for r in conn.execute(
            'SELECT col_key FROM col_meta WHERE table_name=? ORDER BY col_index', (table,)
        ).fetchall()]
        col_set = set(cols)
        valid = {sanitize(k): v for k, v in data.items() if sanitize(k) in col_set}
        if not valid:
            return jsonify(error='no valid columns'), 400
        set_str = ', '.join(f'"{k}"=?' for k in valid)
        now_ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        # Ensure _updated_at column exists
        tbl_cols = [r[1] for r in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
        if '_updated_at' not in tbl_cols:
            conn.execute(f'ALTER TABLE "{table}" ADD COLUMN _updated_at TEXT')
        conn.execute(f'UPDATE "{table}" SET {set_str}, "_updated_at"=? WHERE rowid=?',
                     list(valid.values()) + [now_ts, row_id])
        conn.commit()
        return jsonify(ok=True)
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        conn.close()


@app.route('/api/import_log')
def api_import_log():
    conn = get_db()
    rows = conn.execute('SELECT * FROM import_log ORDER BY id DESC LIMIT 20').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


# ─── Auth ─────────────────────────────────────────────────────────────────────
@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.get_json() or {}
    username = str(data.get('username', '')).strip()
    password = str(data.get('password', '')).strip()
    if not username or not password:
        return jsonify(error='請輸入帳號與密碼'), 400
    conn = get_db()
    row = conn.execute(
        'SELECT * FROM accounts WHERE username=? AND status=?', (username, '啟用')
    ).fetchone()
    conn.close()
    if not row or row['password'] != hash_pw(password):
        return jsonify(error='帳號或密碼錯誤'), 401
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn2 = get_db()
    # Migrate: add columns if missing
    existing_cols = [r[1] for r in conn2.execute('PRAGMA table_info(accounts)').fetchall()]
    if 'last_login' not in existing_cols:
        conn2.execute('ALTER TABLE accounts ADD COLUMN last_login TEXT')
    if 'last_logout' not in existing_cols:
        conn2.execute('ALTER TABLE accounts ADD COLUMN last_logout TEXT')
    conn2.execute('UPDATE accounts SET last_login=? WHERE id=?', (now, row['id']))
    conn2.commit()
    conn2.close()
    session['user'] = {'id': row['id'], 'name': row['name'], 'username': row['username'],
                       'group': row['group_name'], 'email': row['email'], 'last_login': now}
    return jsonify(ok=True, user=dict(session['user']))

@app.route('/api/logout', methods=['POST'])
def api_logout():
    u = session.get('user')
    if u:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = get_db()
        conn.execute('UPDATE accounts SET last_logout=? WHERE id=?', (now, u['id']))
        conn.commit()
        conn.close()
    session.clear()
    return jsonify(ok=True)


@app.route('/api/profile', methods=['GET'])
def api_profile():
    u = session.get('user')
    if not u:
        return jsonify(error='not_logged_in'), 401
    conn = get_db()
    existing_cols = [r[1] for r in conn.execute('PRAGMA table_info(accounts)').fetchall()]
    if 'last_login' not in existing_cols:
        conn.execute('ALTER TABLE accounts ADD COLUMN last_login TEXT')
        conn.commit()
    if 'last_logout' not in existing_cols:
        conn.execute('ALTER TABLE accounts ADD COLUMN last_logout TEXT')
        conn.commit()
    row = conn.execute('SELECT id,name,email,username,group_name,status,last_login,last_logout FROM accounts WHERE id=?', (u['id'],)).fetchone()
    conn.close()
    return jsonify(dict(row))


@app.route('/api/profile', methods=['PUT'])
def api_profile_update():
    u = session.get('user')
    if not u:
        return jsonify(error='not_logged_in'), 401
    d = request.get_json() or {}
    conn = get_db()
    try:
        if d.get('password'):
            conn.execute('UPDATE accounts SET email=?, password=? WHERE id=?',
                         (d.get('email', ''), hash_pw(d['password']), u['id']))
        else:
            conn.execute('UPDATE accounts SET email=? WHERE id=?',
                         (d.get('email', ''), u['id']))
        conn.commit()
        session['user'] = {**u, 'email': d.get('email', u['email'])}
        return jsonify(ok=True)
    except Exception as e:
        return jsonify(error=str(e)), 400
    finally:
        conn.close()

@app.route('/api/me')
def api_me():
    u = session.get('user')
    if not u:
        return jsonify(error='not_logged_in'), 401
    # Backfill last_login if missing (e.g. session survived server restart)
    if not u.get('last_login'):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = get_db()
        existing_cols = [r[1] for r in conn.execute('PRAGMA table_info(accounts)').fetchall()]
        if 'last_login' not in existing_cols:
            conn.execute('ALTER TABLE accounts ADD COLUMN last_login TEXT')
        if 'last_logout' not in existing_cols:
            conn.execute('ALTER TABLE accounts ADD COLUMN last_logout TEXT')
        conn.execute('UPDATE accounts SET last_login=? WHERE id=? AND last_login IS NULL', (now, u['id']))
        conn.commit()
        conn.close()
        u['last_login'] = now
        session['user'] = u
    return jsonify(u)

@app.route('/api/my_permissions')
def api_my_permissions():
    u = session.get('user')
    if not u:
        return jsonify(error='not_logged_in'), 401
    conn = get_db()
    rows = conn.execute(
        'SELECT field_name, has_access FROM field_permissions WHERE group_name=?', (u['group'],)
    ).fetchall()
    conn.close()
    return jsonify({r['field_name']: bool(r['has_access']) for r in rows})


# ─── Accounts CRUD ────────────────────────────────────────────────────────────
def _require_supervisor():
    u = session.get('user')
    if not u or u.get('group') != 'Supervisor':
        return jsonify(error='僅 Supervisor 可操作'), 403
    return None

@app.route('/api/accounts', methods=['GET'])
def api_accounts_list():
    err = _require_supervisor()
    if err: return err
    conn = get_db()
    rows = conn.execute('SELECT id,name,email,username,group_name,status,notes,last_login,last_logout FROM accounts ORDER BY id').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/accounts', methods=['POST'])
def api_accounts_create():
    err = _require_supervisor()
    if err: return err
    d = request.get_json() or {}
    conn = get_db()
    try:
        conn.execute(
            'INSERT INTO accounts(name,email,username,password,group_name,status,notes) VALUES(?,?,?,?,?,?,?)',
            (d.get('name',''), d.get('email',''), d.get('username',''),
             hash_pw(d.get('password','')), d.get('group_name',''), d.get('status','啟用'), d.get('notes',''))
        )
        conn.commit()
        return jsonify(ok=True, id=conn.execute('SELECT last_insert_rowid()').fetchone()[0])
    except Exception as e:
        return jsonify(error=str(e)), 400
    finally:
        conn.close()

@app.route('/api/accounts/<int:aid>', methods=['PUT'])
def api_accounts_update(aid):
    err = _require_supervisor()
    if err: return err
    d = request.get_json() or {}
    conn = get_db()
    try:
        pw_clause = ', password=?' if d.get('password') else ''
        params = [d.get('name',''), d.get('email',''), d.get('username',''),
                  d.get('group_name',''), d.get('status','啟用'), d.get('notes','')]
        if d.get('password'):
            params.insert(3, hash_pw(d['password']))
        params.append(aid)
        conn.execute(
            f'UPDATE accounts SET name=?,email=?,username=?{pw_clause},group_name=?,status=?,notes=? WHERE id=?', params
        )
        conn.commit()
        return jsonify(ok=True)
    except Exception as e:
        return jsonify(error=str(e)), 400
    finally:
        conn.close()

@app.route('/api/accounts/<int:aid>', methods=['DELETE'])
def api_accounts_delete(aid):
    err = _require_supervisor()
    if err: return err
    conn = get_db()
    conn.execute('DELETE FROM accounts WHERE id=?', (aid,))
    conn.commit()
    conn.close()
    return jsonify(ok=True)


# ─── Permissions table ────────────────────────────────────────────────────────
@app.route('/api/perm_table')
def api_perm_table():
    u = session.get('user')
    if not u: return jsonify(error='not_logged_in'), 401
    conn = get_db()
    rows = conn.execute('SELECT field_name, group_name, has_access FROM field_permissions ORDER BY id').fetchall()
    conn.close()
    # Build {field: {group: has_access}}
    result = {}
    for r in rows:
        result.setdefault(r['field_name'], {})[r['group_name']] = bool(r['has_access'])
    return jsonify(result)

@app.route('/api/table_version/<table>')
def api_table_version(table):
    ALLOWED = {'main_table'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400
    u = session.get('user')
    if not u:
        return jsonify(error='not_logged_in'), 401
    conn = get_db()
    try:
        tbl_cols = [r[1] for r in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
        if '_updated_at' not in tbl_cols:
            return jsonify(version='')
        row = conn.execute(f'SELECT MAX(_updated_at) as v FROM "{table}"').fetchone()
        return jsonify(version=row['v'] or '')
    finally:
        conn.close()


@app.route('/api/changes/<table>')
def api_changes(table):
    ALLOWED = {'main_table'}
    if table not in ALLOWED:
        return jsonify(error='invalid table'), 400
    u = session.get('user')
    if not u:
        return jsonify(error='not_logged_in'), 401
    since = request.args.get('since', '')
    conn = get_db()
    try:
        tbl_cols = [r[1] for r in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
        if '_updated_at' not in tbl_cols or not since:
            return jsonify(rows=[])
        rows = conn.execute(
            f'SELECT rowid as _rid, * FROM "{table}" WHERE _updated_at > ?'
            ' ORDER BY _updated_at DESC LIMIT 50', (since,)
        ).fetchall()
        ordered_keys = [r['col_key'] for r in conn.execute(
            'SELECT col_key FROM col_meta WHERE table_name=? ORDER BY col_index', (table,)
        ).fetchall()]
        from collections import OrderedDict
        def fmt(r):
            d = OrderedDict()
            d['_rid'] = r['_rid']
            for k in ordered_keys:
                d[k] = r[k] if k in r.keys() else None
            d['_updated_at'] = r['_updated_at']
            return d
        return jsonify(rows=[fmt(r) for r in rows])
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        conn.close()


@app.route('/api/perm_table/<group>/<path:field>', methods=['PUT'])
def api_perm_update(group, field):
    err = _require_supervisor()
    if err: return err
    d = request.get_json() or {}
    has = 1 if d.get('has_access') else 0
    conn = get_db()
    conn.execute(
        'INSERT OR REPLACE INTO field_permissions(field_name,group_name,has_access) VALUES(?,?,?)',
        (field, group, has)
    )
    conn.commit()
    conn.close()
    return jsonify(ok=True)


if __name__ == '__main__':
    print('╔══════════════════════════════════════════╗')
    print('║  FY26 Q3 續約總表管理系統  啟動中...     ║')
    print('║  http://127.0.0.1:5001                   ║')
    print('╚══════════════════════════════════════════╝')
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=False, host='0.0.0.0', port=port, use_reloader=False)
