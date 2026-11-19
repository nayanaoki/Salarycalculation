# -*- coding: utf-8 -*-
"""データ保管 (SQLite) と設定・対応表の管理（4職種で1つのDB）。

DB: %APPDATA%\\給与自動計算\\payroll.db（v1.3.0 までの介護版と同じファイル）

  settings  共通の値（交通費など）は key のまま、職種ごとの値は "<職種>_<key>"
            例 kaigo_rate_body / nurse_rate_visit / pt_shikaku_q0
  mapping   (job, service) ごとに 介護=身体分・生活分 / 訪問看護=規定分
  employees 氏名・職種・資格
  work_data 取り込んだ実績（氏名＋対象月ごと。職種も一緒に持つ）
  records   保存した計算結果（全職種共通の列）

v1.4.0 で初めて起動したとき（settings.schema_version が無いとき）:
  1. いまの DB を payroll_backup_v13_<日時>.db として同じフォルダに控える
  2. 介護版の形（mapping に job が無い・単価キーに職種が無い 等）を新しい形へ移す
  3. 訪問看護給与（3アプリ）の DB から設定・対応表・従業員・労働記録・計算記録を取り込む
     （訪問看護給与のデータは読むだけで、消したり書き換えたりしない）
"""
import os
import shutil
import sqlite3
import datetime
import json

import app_paths
import jobs
import payroll_calc as calc

DB_PATH = app_paths.db_path()

# ---------------------------------------------------------------- 既定の対応表
# 介護: サービス内容 → (身体分, 生活分)
KAIGO_MAPPING = {
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

# 訪問看護: サービス内容 → 規定時間（分）。0 は「実時間で数える」。算定の根拠:
#   介護保険 訪問看護費（訪看Ⅰ／予防は予訪看Ⅰ）
#     Ⅰ１=20分未満 Ⅰ２=30分未満 Ⅰ３=30分以上1時間未満 Ⅰ４=1時間以上1時間30分未満
#     → 看護師等は区分の上限（20/30/60/90分）で数える
#     「・准」= 准看護師が行う場合（所定の90/100）。時間区分は看護師と同じ
#     Ⅰ５ = 理学療法士・作業療法士・言語聴覚士の訪問。1回20分（20分＝1単位）
#   定期巡回・随時対応型との連携（定期巡回訪看）は月単位の報酬 → 実時間
#   医療保険 訪問看護基本療養費(Ⅰ)・精神科訪問看護基本療養費(Ⅰ) → 実時間
_KUBUN = {"１": 20, "２": 30, "３": 60, "４": 90}
_PREFIXES = ("訪看Ⅰ", "予訪看Ⅰ")


def _nurse_mapping(suffix, iryo):
    m = {}
    for pre in _PREFIXES:
        for no, minutes in _KUBUN.items():
            m[f"{pre}{no}{suffix}"] = minutes
    for s in iryo:
        m[s] = 0
    return m


def _pt_mapping():
    m = {}
    for pre in _PREFIXES:
        m[f"{pre}５"] = 20
        m[f"{pre}５(×２)"] = 40
        m[f"{pre}５(×３)"] = 60
        m[f"{pre}５・２超"] = 20
    m["基本療養費Ⅰ・理学"] = 0
    return m


DEFAULT_MAPPINGS = {
    jobs.KAIGO: KAIGO_MAPPING,
    jobs.NURSE: _nurse_mapping("", ("基本療養費Ⅰ・正看", "精神基本療養費Ⅰ・正看",
                                    "定期巡回訪看", "定期巡回訪看・介５")),
    jobs.LPN: _nurse_mapping("・准", ("基本療養費Ⅰ・准看", "精神基本療養費Ⅰ・准看",
                                      "定期巡回訪看", "定期巡回訪看・介５")),
    jobs.PT: _pt_mapping(),
}

# 訪問看護 v1.0.0 で3職種に共通で入れていた表（取り込み時に他職種の行を見分ける）
_LEGACY_NURSE_COMMON = {}
for _pre in _PREFIXES:
    for _no, _m in _KUBUN.items():
        _LEGACY_NURSE_COMMON[f"{_pre}{_no}"] = _m
        _LEGACY_NURSE_COMMON[f"{_pre}{_no}・准"] = _m
    _LEGACY_NURSE_COMMON[f"{_pre}５"] = 20
    _LEGACY_NURSE_COMMON[f"{_pre}５(×２)"] = 40
    _LEGACY_NURSE_COMMON[f"{_pre}５(×３)"] = 60
for _s in ("基本療養費Ⅰ・正看", "基本療養費Ⅰ・准看", "基本療養費Ⅰ・理学",
           "精神基本療養費Ⅰ・正看", "精神基本療養費Ⅰ・准看",
           "定期巡回訪看", "定期巡回訪看・介５"):
    _LEGACY_NURSE_COMMON[_s] = 0


# ---------------------------------------------------------------- 設定
def jkey(job, key):
    """職種ごとの設定キー。"""
    return f"{job}_{key}"


def shikaku_key(job, i):
    """資格の並び順 → 設定キー（資格名を変えてもキーは変わらない）。"""
    return jkey(job, f"shikaku_q{i}")


DEFAULT_SETTINGS = {
    "kotsu_unit": "100",            # 交通費 1件あたり（全職種一律）
    "kotsu_first_included": "1",    # 1日の最初の1件を交通費に含める(1)/含めない(0)
    "default_shikaku": "0",         # 資格手当 初期値・未登録者 円/時
    "use_kitei": "1",               # 訪問看護の計上時間: 1=規定時間 / 0=実時間
    "attend_start": "8:30",
    "attend_end": "17:30",
    "attend_employ": "パートスタッフ",
}
for _job, _j in jobs.JOBS.items():
    for _k, _label, _v in _j["rates"]:
        DEFAULT_SETTINGS[jkey(_job, _k)] = _v
    for _i, _q in enumerate(_j["qualifications"]):
        DEFAULT_SETTINGS[shikaku_key(_job, _i)] = "0"
    DEFAULT_SETTINGS[jkey(_job, "attend_post")] = _j["attend_post"]
    DEFAULT_SETTINGS[jkey(_job, "attend_shotei")] = _j["attend_shotei"]

RECORD_COLS = ["job", "person", "period", "visits", "kotsu_count", "total_min",
               "training_min", "body_min", "living_min", "visit_min", "units",
               "rate_training", "rate_body", "rate_living", "rate_visit",
               "rate_kaizen", "rate_kaizen_body", "rate_kaizen_living",
               "amt_training", "amt_body", "amt_living", "amt_visit", "amt_kaizen",
               "kotsu", "rate_shikaku", "shikaku", "other1", "other2",
               "total_amount", "created_at"]
_REAL = {c for c in RECORD_COLS if c.startswith("rate_")}

# 初回起動時の取り込み結果（画面に1度だけ出す）
LEGACY_SUMMARY = []


def connect(path=None):
    conn = sqlite3.connect(path or DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


SCHEMA_VERSION = 2          # 2 = 4職種を1つの DB で持つ形（v1.4.0〜）


def _tables(c):
    return {r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}


def _columns(c, table):
    return {r[1] for r in c.execute(f"PRAGMA table_info({table})").fetchall()}


def _schema_version(c):
    if "settings" not in _tables(c):
        return 0
    r = c.execute("SELECT value FROM settings WHERE key='schema_version'").fetchone()
    try:
        return int(r[0]) if r else 1
    except (TypeError, ValueError):
        return 1


def _backup_before_migrate():
    """移行の前に、いまの DB を同じフォルダへ控える（失敗しても移行は止めない）。"""
    if not os.path.exists(DB_PATH) or os.path.getsize(DB_PATH) == 0:
        return None
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(os.path.dirname(DB_PATH), f"payroll_backup_v13_{stamp}.db")
    try:
        shutil.copy2(DB_PATH, dst)
        return dst
    except OSError:
        return None


def _migrate_kaigo_v13(c):
    """v1.3.0 までの介護版 DB を、4職種で持つ形へ移す（介護のデータはすべて残す）。"""
    tables = _tables(c)
    # 単価・資格手当・出勤簿の値 → 介護の職種キーへ（古いキーは残しても害はない）
    if "settings" in tables:
        old = {r[0]: r[1] for r in c.execute("SELECT key,value FROM settings")}
        _shared, own = _legacy_settings(jobs.KAIGO, old)
        for k, v in own.items():
            _set(c, k, v)
    # 対応表: (service, body, living) → (job, service, body, living, minutes)
    if "mapping" in tables and "job" not in _columns(c, "mapping"):
        c.execute("ALTER TABLE mapping RENAME TO mapping_v13")
        c.execute("""CREATE TABLE mapping(
            job TEXT NOT NULL, service TEXT NOT NULL,
            body INTEGER DEFAULT 0, living INTEGER DEFAULT 0, minutes INTEGER DEFAULT 0,
            PRIMARY KEY(job, service))""")
        c.execute("INSERT OR REPLACE INTO mapping(job,service,body,living,minutes) "
                  "SELECT 'kaigo', TRIM(service), COALESCE(body,0), COALESCE(living,0), 0 "
                  "FROM mapping_v13 ORDER BY rowid")
        c.execute("DROP TABLE mapping_v13")
    # 従業員・労働記録・計算記録: 職種の列を足し、いまある行は介護にする
    for table in ("employees", "work_data", "records"):
        if table in tables and "job" not in _columns(c, table):
            c.execute(f"ALTER TABLE {table} ADD COLUMN job TEXT")
        if table in tables:
            c.execute(f"UPDATE {table} SET job='kaigo' WHERE job IS NULL OR job=''")
    if "records" in tables:
        have = _columns(c, "records")
        for col in RECORD_COLS:
            if col not in have:
                kind = "REAL" if col in _REAL else "TEXT" if col in (
                    "job", "person", "period", "created_at") else "INTEGER"
                c.execute(f"ALTER TABLE records ADD COLUMN {col} {kind} DEFAULT 0")
        # 交通費件数の列が無かった頃の記録は、訪問件数で数えていた
        c.execute("UPDATE records SET kotsu_count=visits WHERE kotsu_count IS NULL")
        for col in ("visit_min", "units", "amt_visit", "rate_visit", "rate_kaizen",
                    "kotsu_count", "amt_kaizen", "rate_kaizen_living", "rate_kaizen_body",
                    "rate_shikaku"):
            c.execute(f"UPDATE records SET {col}=0 WHERE {col} IS NULL")


def init_db():
    conn = connect()
    c = conn.cursor()
    version = _schema_version(c)
    had_kaigo = bool(_tables(c) & {"mapping", "employees", "records", "work_data"})
    if version < SCHEMA_VERSION and had_kaigo:
        backup = _backup_before_migrate()
        _migrate_kaigo_v13(c)
        if backup:
            LEGACY_SUMMARY.append(f"移行前の控え: {os.path.basename(backup)}")
    c.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)")
    c.execute("""CREATE TABLE IF NOT EXISTS mapping(
        job TEXT NOT NULL, service TEXT NOT NULL,
        body INTEGER DEFAULT 0, living INTEGER DEFAULT 0, minutes INTEGER DEFAULT 0,
        PRIMARY KEY(job, service))""")
    c.execute("""CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mgmt_no TEXT, name TEXT NOT NULL, job TEXT, qualification TEXT,
        created_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS work_data(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person TEXT, period TEXT, job TEXT, rows_json TEXT, kotsu_flags TEXT,
        created_at TEXT, updated_at TEXT, UNIQUE(person, period))""")
    cols = ", ".join(f"{col} {'REAL' if col in _REAL else 'TEXT' if col in ('job', 'person', 'period', 'created_at') else 'INTEGER'}"
                     for col in RECORD_COLS)
    c.execute(f"CREATE TABLE IF NOT EXISTS records("
              f"id INTEGER PRIMARY KEY AUTOINCREMENT, {cols})")
    for k, v in DEFAULT_SETTINGS.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    for job, m in DEFAULT_MAPPINGS.items():
        if c.execute("SELECT COUNT(*) FROM mapping WHERE job=?", (job,)).fetchone()[0] == 0:
            _write_mapping(c, job, m)
    if version < SCHEMA_VERSION:
        # 介護のデータがある端末では、交通費など共通の値は介護の値を残す
        keep = set(_SHARED) if had_kaigo else set()
        try:
            LEGACY_SUMMARY.extend(_import_legacy(c, keep))
        except Exception as ex:                    # 取り込めなくても起動は続ける
            LEGACY_SUMMARY.append(f"訪問看護給与のデータの取り込みに失敗しました: {ex}")
        _set(c, "schema_version", SCHEMA_VERSION)
        _set(c, "legacy_imported", _now())
    conn.commit()
    conn.close()


# ---------------------------------------------------------------- 統合前データの取り込み
_KAIGO_SHIKAKU = ["shikaku_shoninsha", "shikaku_jitsumusha",
                  "shikaku_kaigofukushishi", "shikaku_sonota"]
_SHARED = ("kotsu_unit", "kotsu_first_included", "default_shikaku",
           "attend_start", "attend_end", "attend_employ")


def _legacy_settings(job, old):
    """統合前の設定 → 統合版のキーへ。(共通の値, 職種の値) を返す。"""
    shared = {k: old[k] for k in _SHARED if k in old}
    own = {}
    if job == jobs.KAIGO:
        for k in ("rate_training", "rate_living", "rate_body",
                  "rate_kaizen_living", "rate_kaizen_body"):
            if k in old:
                own[jkey(job, k)] = old[k]
        for i, k in enumerate(_KAIGO_SHIKAKU):
            if k in old:
                own[shikaku_key(job, i)] = old[k]
    else:
        for k in ("rate_training", "rate_visit", "rate_kaizen"):
            if k in old:
                own[jkey(job, k)] = old[k]
        for i in range(len(jobs.qualifications(job))):
            k = f"shikaku_q{i}"
            if k in old:
                own[shikaku_key(job, i)] = old[k]
        if "use_kitei" in old:
            shared["use_kitei"] = old["use_kitei"]
    for k in ("attend_post", "attend_shotei"):
        if k in old:
            own[jkey(job, k)] = old[k]
    return shared, own


def _legacy_mapping(job, src):
    """統合前の対応表（手で直した行も含めてそのまま）。"""
    if job == jobs.KAIGO:
        return {r["service"]: (int(r["body"] or 0), int(r["living"] or 0))
                for r in src.execute("SELECT service,body,living FROM mapping")}
    rows = {r["service"]: int(r["minutes"] or 0)
            for r in src.execute("SELECT service,minutes FROM mapping")}
    ver = src.execute("SELECT value FROM settings WHERE key='mapping_version'").fetchone()
    if not ver or str(ver[0]) < "2":
        # v1.0.0 の共通表のまま → 他職種の既定行を除き、この職種の既定行を足す
        default = DEFAULT_MAPPINGS[job]
        rows = {s: m for s, m in rows.items()
                if s in default or _LEGACY_NURSE_COMMON.get(s) != m}
        for s, m in default.items():
            rows.setdefault(s, m)
    return rows


def _import_legacy(c, keep_shared=()):
    """訪問看護給与（3アプリ）の DB から、設定・対応表・従業員・労働記録・計算記録を取り込む。

    keep_shared に入っている共通の値（交通費など）は上書きしない。
    """
    summary = []
    shared_done = set(keep_shared)
    for job in jobs.ORDER:
        path = app_paths.legacy_db_path(job)
        if not path:
            continue
        src = connect(path)
        try:
            tables = {r[0] for r in src.execute(
                "SELECT name FROM sqlite_master WHERE type='table'")}
            n_emp = n_wd = n_rec = 0
            if "settings" in tables:
                old = {r["key"]: r["value"] for r in
                       src.execute("SELECT key,value FROM settings")}
                shared, own = _legacy_settings(job, old)
                for k, v in own.items():
                    _set(c, k, v)
                for k, v in shared.items():
                    if k not in shared_done:
                        _set(c, k, v)
                        shared_done.add(k)
            if "mapping" in tables:
                m = _legacy_mapping(job, src)
                if m:
                    c.execute("DELETE FROM mapping WHERE job=?", (job,))
                    _write_mapping(c, job, m)
            if "employees" in tables:
                have = {(_name_key(r[0]), r[1]) for r in
                        c.execute("SELECT name,job FROM employees")}
                for e in src.execute("SELECT mgmt_no,name,qualification,created_at "
                                     "FROM employees ORDER BY id"):
                    if (_name_key(e["name"]), job) in have:
                        continue
                    c.execute("INSERT INTO employees(mgmt_no,name,job,qualification,"
                              "created_at) VALUES(?,?,?,?,?)",
                              (e["mgmt_no"], e["name"], job, e["qualification"],
                               e["created_at"] or _now()))
                    n_emp += 1
            if "work_data" in tables:
                for w in src.execute("SELECT * FROM work_data ORDER BY id"):
                    w = dict(w)
                    cur = c.execute(
                        "INSERT OR IGNORE INTO work_data(person,period,job,rows_json,"
                        "kotsu_flags,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
                        (w.get("person"), w.get("period"), job, w.get("rows_json"),
                         w.get("kotsu_flags"), w.get("created_at"), w.get("updated_at")))
                    n_wd += cur.rowcount
            if "records" in tables:
                for r in src.execute("SELECT * FROM records ORDER BY id"):
                    r = dict(r)
                    r["job"] = job
                    for k in ("body_min", "living_min", "visit_min", "units",
                              "amt_body", "amt_living", "amt_visit", "amt_kaizen",
                              "kotsu_count", "shikaku", "other1", "other2"):
                        r[k] = r.get(k) or 0
                    for k in _REAL:
                        r[k] = r.get(k) or 0
                    if r.get("kotsu_count") == 0 and r.get("kotsu"):
                        r["kotsu_count"] = r.get("visits") or 0
                    c.execute(f"INSERT INTO records({','.join(RECORD_COLS)}) "
                              f"VALUES({','.join('?' * len(RECORD_COLS))})",
                              [r.get(col) for col in RECORD_COLS])
                    n_rec += 1
            summary.append(f"{jobs.label(job)}：従業員 {n_emp}名・労働記録 {n_wd}件・"
                           f"計算記録 {n_rec}件")
        finally:
            src.close()
    return summary


def _set(c, k, v):
    c.execute("INSERT INTO settings(key,value) VALUES(?,?) "
              "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))


# ---------------------------------------------------------------- 設定
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
        _set(conn, k, v)
    conn.commit()
    conn.close()


def _to_float(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def rates_for(job, settings=None):
    """職種の単価 dict（calc.amounts が読む形）。"""
    s = settings or get_settings()
    g = lambda k: _to_float(s.get(jkey(job, k)))   # noqa: E731
    if job == jobs.KAIGO:
        return {"training": g("rate_training"), "living": g("rate_living"),
                "body": g("rate_body"), "kaizen_living": g("rate_kaizen_living"),
                "kaizen_body": g("rate_kaizen_body")}
    return {"training": g("rate_training"), "visit": g("rate_visit"),
            "kaizen": g("rate_kaizen")}


def shikaku_rate_for(job, qualification, settings=None):
    """資格に応じた資格手当(円/時)。未設定(0)や未登録なら初期値を使う。"""
    settings = settings or get_settings()
    quals = jobs.qualifications(job)
    if qualification in quals:
        rate = _to_float(settings.get(shikaku_key(job, quals.index(qualification))))
        if rate > 0:
            return rate
    return _to_float(settings.get("default_shikaku", 0))


# ---------------------------------------------------------------- 対応表
def _write_mapping(c, job, mapping):
    for s, v in mapping.items():
        if job == jobs.KAIGO:
            s = str(s).strip()
            b, l = v
            c.execute("INSERT OR REPLACE INTO mapping(job,service,body,living,minutes) "
                      "VALUES(?,?,?,?,0)", (job, s, int(b or 0), int(l or 0)))
        else:
            s = calc.normalize_service(s)
            c.execute("INSERT OR REPLACE INTO mapping(job,service,body,living,minutes) "
                      "VALUES(?,?,0,0,?)", (job, s, int(v or 0)))


def get_mapping(job):
    """計算用。介護 {内容: (身体分, 生活分)} / 訪問看護 {正規化した内容: 規定分}。"""
    conn = connect()
    rows = conn.execute("SELECT service,body,living,minutes FROM mapping WHERE job=? "
                        "ORDER BY rowid", (job,)).fetchall()
    conn.close()
    if job == jobs.KAIGO:
        return {r["service"]: (int(r["body"] or 0), int(r["living"] or 0)) for r in rows}
    return {calc.normalize_service(r["service"]): int(r["minutes"] or 0) for r in rows}


def get_mapping_rows(job):
    """画面表示用。介護 [(内容, 身体分, 生活分)] / 訪問看護 [(内容, 規定分)]。"""
    conn = connect()
    rows = conn.execute("SELECT service,body,living,minutes FROM mapping WHERE job=? "
                        "ORDER BY rowid", (job,)).fetchall()
    conn.close()
    if job == jobs.KAIGO:
        return [(r["service"], int(r["body"] or 0), int(r["living"] or 0)) for r in rows]
    return [(r["service"], int(r["minutes"] or 0)) for r in rows]


def save_mapping(job, mapping):
    """その職種の対応表を全置換。"""
    conn = connect()
    conn.execute("DELETE FROM mapping WHERE job=?", (job,))
    _write_mapping(conn, job, {s: v for s, v in mapping.items() if str(s).strip()})
    conn.commit()
    conn.close()


# ---------------------------------------------------------------- 記録
def save_record(rec):
    rec = dict(rec)
    rec.setdefault("created_at", _now())
    conn = connect()
    conn.execute(
        f"INSERT INTO records({','.join(RECORD_COLS)}) "
        f"VALUES({','.join('?' * len(RECORD_COLS))})",
        [rec.get(c) for c in RECORD_COLS])
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


# ---------------------------------------------------------------- 従業員
def list_employees():
    conn = connect()
    rows = conn.execute("SELECT id,mgmt_no,name,job,qualification FROM employees "
                        "ORDER BY mgmt_no, id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_employee(mgmt_no, name, job, qualification):
    conn = connect()
    conn.execute("INSERT INTO employees(mgmt_no,name,job,qualification,created_at) "
                 "VALUES(?,?,?,?,?)", (mgmt_no, name, job, qualification, _now()))
    conn.commit()
    conn.close()


def update_employee(eid, mgmt_no, name, job, qualification):
    conn = connect()
    conn.execute("UPDATE employees SET mgmt_no=?, name=?, job=?, qualification=? "
                 "WHERE id=?", (mgmt_no, name, job, qualification, eid))
    conn.commit()
    conn.close()


def delete_employee(eid):
    conn = connect()
    conn.execute("DELETE FROM employees WHERE id=?", (eid,))
    conn.commit()
    conn.close()


def _name_key(name):
    return "".join(str(name or "").split())


def get_employee_by_name(name):
    """氏名で従業員を1件取得（空白の有無は無視。PDF は「山田 太郎」と空白入り）。"""
    key = _name_key(name)
    if not key:
        return None
    for e in list_employees():
        if _name_key(e["name"]) == key:
            return e
    return None


# ---------------------------------------------------------------- 取込データ
def save_work_data(person, period, rows, job=None):
    """取込データを person+period 単位で保存（上書き）。交通費の手動設定は残す。"""
    norm = [["" if c is None else str(c) for c in row] for row in rows]
    payload = json.dumps(norm, ensure_ascii=False)
    now = _now()
    conn = connect()
    conn.execute(
        "INSERT INTO work_data(person,period,job,rows_json,created_at,updated_at) "
        "VALUES(?,?,?,?,?,?) ON CONFLICT(person,period) DO UPDATE SET "
        "job=excluded.job, rows_json=excluded.rows_json, updated_at=excluded.updated_at",
        (person, period, job, payload, now, now))
    conn.commit()
    conn.close()


def set_work_data_job(person, period, job):
    conn = connect()
    conn.execute("UPDATE work_data SET job=? WHERE person=? AND period=?",
                 (job, person, period))
    conn.commit()
    conn.close()


def _row_to_work_data(r):
    d = dict(r)
    try:
        d["rows"] = [calc.clean_row(x) for x in json.loads(d.pop("rows_json") or "[]")]
    except (ValueError, TypeError):
        d["rows"] = []
    try:
        flags = json.loads(d.pop("kotsu_flags") or "null")
    except (ValueError, TypeError):
        flags = None
    d["kotsu_flags"] = flags if isinstance(flags, list) else None
    return d


_WD_COLS = "id,person,period,job,rows_json,kotsu_flags,updated_at"


def list_work_data():
    conn = connect()
    rows = conn.execute(f"SELECT {_WD_COLS} FROM work_data "
                        "ORDER BY updated_at DESC, id DESC").fetchall()
    conn.close()
    return [_row_to_work_data(r) for r in rows]


def get_work_data(wid):
    conn = connect()
    r = conn.execute(f"SELECT {_WD_COLS} FROM work_data WHERE id=?", (wid,)).fetchone()
    conn.close()
    return _row_to_work_data(r) if r else None


def get_kotsu_flags(person, period):
    conn = connect()
    r = conn.execute("SELECT kotsu_flags FROM work_data WHERE person=? AND period=?",
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
    conn = connect()
    conn.execute("UPDATE work_data SET kotsu_flags=? WHERE person=? AND period=?",
                 (json.dumps(flags, ensure_ascii=False), person, period))
    conn.commit()
    conn.close()


def delete_work_data(wid):
    conn = connect()
    conn.execute("DELETE FROM work_data WHERE id=?", (wid,))
    conn.commit()
    conn.close()


def db_exists():
    return os.path.exists(DB_PATH)
