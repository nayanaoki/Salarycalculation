# -*- coding: utf-8 -*-
"""データ保管 (SQLite) と設定・対応表の管理。"""
import sqlite3
import datetime
import json

import app_paths

# DB は書込可能なユーザー領域(%APPDATA%)に保存する。
# 配布版でインストール先が読取専用でも、更新で上書きされても消えない。
# 旧配置(実行ファイル同階層)に DB があれば自動移行される。
DB_PATH = app_paths.db_path()

# サービス内容 → (身体分, 生活分) 既定対応表
DEFAULT_MAPPING = {
    "訪問型独自サービス１１": (0, 60),
    "訪問型独自サービス１２": (0, 60),
    "訪問型独自サービス１３": (0, 60),
    "訪問型独自サービス２１": (0, 60),
    "生活２・Ⅱ": (0, 45),
    "生活３・Ⅱ": (0, 60),
    "身体１・Ⅱ": (30, 0),
    "身体２・Ⅱ": (60, 0),
    "身体３・Ⅱ": (90, 0),
    "身１生１・Ⅱ": (30, 20),
    "身１生２・Ⅱ": (30, 45),
    "身１生３・Ⅱ": (30, 70),
    "身２生１・Ⅱ": (60, 20),
    "身２生２・Ⅱ": (60, 45),
    "身２生３・Ⅱ": (60, 70),
}

# 資格の選択肢（従業員管理のプルダウン・資格手当の区分）
QUALIFICATIONS = ["初任者", "実務者", "介護福祉士", "その他"]

# 資格 → 資格手当(円/時)の設定キー
SHIKAKU_KEYS = {
    "初任者": "shikaku_shoninsha",
    "実務者": "shikaku_jitsumusha",
    "介護福祉士": "shikaku_kaigofukushishi",
    "その他": "shikaku_sonota",
}

DEFAULT_SETTINGS = {
    "rate_training": "1100",        # 研修時給
    "rate_living": "1200",          # 生活時給
    "rate_body": "1500",            # 身体時給
    "kotsu_unit": "100",            # 交通費単価 (1件あたり)
    "kotsu_first_included": "1",    # 1日の最初の1件を交通費対象に含める(1)/含めない(0)
    "rate_kaizen_living": "0",      # 処遇改善手当 (生活/時間)
    "rate_kaizen_body": "0",        # 処遇改善手当 (身体/時間)
    "default_shikaku": "0",         # 資格手当 初期値 (1時間あたりの支給額・各人で編集可能)
    "shikaku_shoninsha": "0",       # 資格手当(初任者) 円/時
    "shikaku_jitsumusha": "0",      # 資格手当(実務者) 円/時
    "shikaku_kaigofukushishi": "0",  # 資格手当(介護福祉士) 円/時
    "shikaku_sonota": "0",          # 資格手当(その他) 円/時
    "attend_start": "8:30",         # 個人出勤簿 始業時間
    "attend_end": "17:30",          # 個人出勤簿 終業時間
    "attend_shotei": "9:00",        # 個人出勤簿 所定内勤務時間
    "attend_post": "訪問介護職員",   # 個人出勤簿 所属
    "attend_employ": "パートスタッフ",  # 個人出勤簿 雇用形態
}


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _add_missing_columns(cursor, table, columns):
    """既存テーブルに不足している列を追加する(後方互換のためのマイグレーション)。"""
    existing = {r[1] for r in cursor.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, coltype in columns.items():
        if name not in existing:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {name} {coltype}")


def init_db():
    conn = connect()
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS settings(
        key TEXT PRIMARY KEY, value TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS mapping(
        service TEXT PRIMARY KEY, body INTEGER, living INTEGER)""")
    c.execute("""CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mgmt_no TEXT, name TEXT NOT NULL, qualification TEXT,
        created_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS work_data(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person TEXT, period TEXT, rows_json TEXT, kotsu_flags TEXT,
        created_at TEXT, updated_at TEXT,
        UNIQUE(person, period))""")
    c.execute("""CREATE TABLE IF NOT EXISTS records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person TEXT, period TEXT,
        visits INTEGER, total_min INTEGER,
        body_min INTEGER, living_min INTEGER, training_min INTEGER,
        rate_training REAL, rate_living REAL, rate_body REAL,
        amt_training INTEGER, amt_living INTEGER, amt_body INTEGER,
        kotsu INTEGER, kotsu_count INTEGER, shikaku INTEGER, rate_shikaku REAL,
        other1 INTEGER, other2 INTEGER,
        rate_kaizen_living REAL, rate_kaizen_body REAL, amt_kaizen INTEGER,
        total_amount INTEGER, created_at TEXT)""")
    # 既存DBへの列追加(マイグレーション)
    _add_missing_columns(c, "records", {
        "rate_kaizen_living": "REAL",
        "rate_kaizen_body": "REAL",
        "amt_kaizen": "INTEGER",
        "rate_shikaku": "REAL",
        "kotsu_count": "INTEGER",
    })
    _add_missing_columns(c, "work_data", {"kotsu_flags": "TEXT"})
    # 初期データ投入
    for k, v in DEFAULT_SETTINGS.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    if c.execute("SELECT COUNT(*) FROM mapping").fetchone()[0] == 0:
        for s, (b, l) in DEFAULT_MAPPING.items():
            c.execute("INSERT INTO mapping(service,body,living) VALUES(?,?,?)", (s, b, l))
    conn.commit()
    conn.close()


# --- 設定 ---
def get_settings():
    conn = connect()
    rows = conn.execute("SELECT key,value FROM settings").fetchall()
    conn.close()
    d = dict(DEFAULT_SETTINGS)
    d.update({r["key"]: r["value"] for r in rows})
    return d


def save_settings(d):
    conn = connect()
    for k, v in d.items():
        conn.execute("INSERT INTO settings(key,value) VALUES(?,?) "
                     "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))
    conn.commit()
    conn.close()


# --- 対応表 ---
def get_mapping():
    conn = connect()
    rows = conn.execute("SELECT service,body,living FROM mapping ORDER BY rowid").fetchall()
    conn.close()
    return {r["service"]: (r["body"], r["living"]) for r in rows}


def save_mapping(mapping):
    """mapping: {service: (body, living)} で全置換。"""
    conn = connect()
    conn.execute("DELETE FROM mapping")
    for s, (b, l) in mapping.items():
        conn.execute("INSERT INTO mapping(service,body,living) VALUES(?,?,?)",
                     (s, int(b), int(l)))
    conn.commit()
    conn.close()


# --- 記録 ---
def save_record(rec):
    rec = dict(rec)
    rec.setdefault("created_at", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    cols = ["person", "period", "visits", "total_min", "body_min", "living_min",
            "training_min", "rate_training", "rate_living", "rate_body",
            "amt_training", "amt_living", "amt_body", "kotsu", "kotsu_count", "shikaku",
            "rate_shikaku", "other1", "other2", "rate_kaizen_living", "rate_kaizen_body",
            "amt_kaizen", "total_amount", "created_at"]
    conn = connect()
    conn.execute(
        f"INSERT INTO records({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
        [rec.get(c) for c in cols])
    conn.commit()
    rid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return rid


def list_records():
    conn = connect()
    rows = conn.execute("SELECT * FROM records ORDER BY created_at DESC, id DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_record(rid):
    conn = connect()
    r = conn.execute("SELECT * FROM records WHERE id=?", (rid,)).fetchone()
    conn.close()
    return dict(r) if r else None


def delete_record(rid):
    conn = connect()
    conn.execute("DELETE FROM records WHERE id=?", (rid,))
    conn.commit()
    conn.close()


# --- 従業員 ---
def list_employees():
    conn = connect()
    rows = conn.execute(
        "SELECT id,mgmt_no,name,qualification FROM employees "
        "ORDER BY mgmt_no, id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_employee(mgmt_no, name, qualification):
    conn = connect()
    conn.execute(
        "INSERT INTO employees(mgmt_no,name,qualification,created_at) "
        "VALUES(?,?,?,?)",
        (mgmt_no, name, qualification,
         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def update_employee(eid, mgmt_no, name, qualification):
    conn = connect()
    conn.execute(
        "UPDATE employees SET mgmt_no=?, name=?, qualification=? WHERE id=?",
        (mgmt_no, name, qualification, eid))
    conn.commit()
    conn.close()


def delete_employee(eid):
    conn = connect()
    conn.execute("DELETE FROM employees WHERE id=?", (eid,))
    conn.commit()
    conn.close()


def get_employee_by_name(name):
    """氏名で従業員を1件取得(前後空白は無視)。未登録なら None。"""
    if not name:
        return None
    conn = connect()
    r = conn.execute(
        "SELECT id,mgmt_no,name,qualification FROM employees "
        "WHERE TRIM(name)=TRIM(?) LIMIT 1", (name,)).fetchone()
    conn.close()
    return dict(r) if r else None


def _to_float(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


# --- 取込データ(読み込んだ勤務実績の生データ) ---
def _cell_to_str(v):
    return "" if v is None else str(v)


def save_work_data(person, period, rows):
    """読み込んだ生データ(行の並び)を person+period 単位で保存(upsert)。

    セルは datetime.time 等になり得るため文字列化して JSON で保管する。
    再読込時は calc.calc_rows がそのまま集計できる。
    """
    norm = [[_cell_to_str(c) for c in row] for row in rows]
    payload = json.dumps(norm, ensure_ascii=False)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = connect()
    conn.execute(
        "INSERT INTO work_data(person,period,rows_json,created_at,updated_at) "
        "VALUES(?,?,?,?,?) "
        "ON CONFLICT(person,period) DO UPDATE SET "
        "rows_json=excluded.rows_json, updated_at=excluded.updated_at",
        (person, period, payload, now, now))
    conn.commit()
    conn.close()


def _row_to_work_data(r):
    """work_data の1行を dict 化し、rows / kotsu_flags を JSON から復元する。"""
    d = dict(r)
    try:
        d["rows"] = json.loads(d.pop("rows_json") or "[]")
    except (ValueError, TypeError):
        d["rows"] = []
    try:
        flags = json.loads(d.pop("kotsu_flags") or "null")
    except (ValueError, TypeError):
        flags = None
    d["kotsu_flags"] = flags if isinstance(flags, list) else None
    return d


def list_work_data():
    """保存済み取込データの一覧(生データ・交通費フラグ含む)。"""
    conn = connect()
    rows = conn.execute(
        "SELECT id,person,period,rows_json,kotsu_flags,updated_at FROM work_data "
        "ORDER BY updated_at DESC, id DESC").fetchall()
    conn.close()
    return [_row_to_work_data(r) for r in rows]


def get_work_data(wid):
    conn = connect()
    r = conn.execute(
        "SELECT id,person,period,rows_json,kotsu_flags,updated_at FROM work_data "
        "WHERE id=?", (wid,)).fetchone()
    conn.close()
    return _row_to_work_data(r) if r else None


def get_kotsu_flags(person, period):
    """person+period の交通費あり/なし手動設定を返す(未設定なら None)。"""
    conn = connect()
    r = conn.execute(
        "SELECT kotsu_flags FROM work_data WHERE person=? AND period=?",
        (person, period)).fetchone()
    conn.close()
    if not r or not r["kotsu_flags"]:
        return None
    try:
        flags = json.loads(r["kotsu_flags"])
    except (ValueError, TypeError):
        return None
    return flags if isinstance(flags, list) else None


def save_kotsu_flags(person, period, flags):
    """交通費のあり/なし手動設定(訪問ごとの True/False/None のリスト)を保存する。"""
    payload = json.dumps(flags, ensure_ascii=False)
    conn = connect()
    conn.execute(
        "UPDATE work_data SET kotsu_flags=? WHERE person=? AND period=?",
        (payload, person, period))
    conn.commit()
    conn.close()


def delete_work_data(wid):
    conn = connect()
    conn.execute("DELETE FROM work_data WHERE id=?", (wid,))
    conn.commit()
    conn.close()


def shikaku_rate_for(qualification, settings=None):
    """資格に対応する資格手当(円/時)を返す。

    その資格の手当が設定されている(0より大きい)ときはそれを適用し、
    未設定(0)や未登録の資格のときは default_shikaku(初期値)を適用する。
    """
    settings = settings or get_settings()
    key = SHIKAKU_KEYS.get(qualification or "")
    if key:
        rate = _to_float(settings.get(key))
        if rate > 0:
            return rate
    return _to_float(settings.get("default_shikaku", 0))
