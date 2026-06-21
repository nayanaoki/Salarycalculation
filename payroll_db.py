# -*- coding: utf-8 -*-
"""データ保管 (SQLite) と設定・対応表の管理。"""
import sqlite3
import datetime

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

DEFAULT_SETTINGS = {
    "rate_training": "1100",   # 研修時給
    "rate_living": "1200",     # 生活時給
    "rate_body": "1500",       # 身体時給
    "kotsu_unit": "100",       # 交通費 1件単価
}


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = connect()
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS settings(
        key TEXT PRIMARY KEY, value TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS mapping(
        service TEXT PRIMARY KEY, body INTEGER, living INTEGER)""")
    c.execute("""CREATE TABLE IF NOT EXISTS records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person TEXT, period TEXT,
        visits INTEGER, total_min INTEGER,
        body_min INTEGER, living_min INTEGER, training_min INTEGER,
        rate_training REAL, rate_living REAL, rate_body REAL,
        amt_training INTEGER, amt_living INTEGER, amt_body INTEGER,
        kotsu INTEGER, shikaku INTEGER, other1 INTEGER, other2 INTEGER,
        total_amount INTEGER, created_at TEXT)""")
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
            "amt_training", "amt_living", "amt_body", "kotsu", "shikaku",
            "other1", "other2", "total_amount", "created_at"]
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
