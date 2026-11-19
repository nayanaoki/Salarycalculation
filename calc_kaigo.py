# -*- coding: utf-8 -*-
"""訪問介護の勤務実績の集計・給与計算ロジック（統合前の「給与自動計算」から移植）。"""
import re
from collections import defaultdict

# 列インデックス (0始まり): A実施日 B曜日 C開始 D終了 E時間 F利用者名 G区分 H内容
COL_DATE, COL_DOW, COL_START, COL_END, COL_DUR, COL_USER, COL_KUBUN, COL_SERVICE = range(8)
HEADER_ROWS = 2


def _to_minutes_of_day(cell):
    """時刻セルを 0時からの分に変換。datetime.time / 'H:MM' 文字列に対応。"""
    if cell is None:
        return None
    hour = getattr(cell, "hour", None)
    if hour is not None:
        return hour * 60 + getattr(cell, "minute", 0)
    m = re.match(r"\s*(\d{1,2})\s*[:：]\s*(\d{1,2})", str(cell))
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def parse_minutes(dur_cell, start_cell, end_cell):
    """列E(例 '75分')から分。無ければ開始/終了から算出。

    Excel読込では時刻は datetime.time、貼り付けでは 'H:MM' 文字列になるため
    どちらでも算出できるようにする。
    """
    if dur_cell is not None and str(dur_cell).strip() != "":
        m = re.search(r"(\d+)", str(dur_cell))
        if m:
            return int(m.group(1))
    s = _to_minutes_of_day(start_cell)
    e = _to_minutes_of_day(end_cell)
    if s is not None and e is not None and e >= s:
        return e - s
    return 0


def is_plain_body(service):
    """番号なしの「身体」かどうか。

    「身体１・Ⅱ」「身体2」など番号付きは対応表の固定値で計算するため除外し、
    番号の付かない「身体」「身体・Ⅱ」等のみ実時間計算の対象とする。
    """
    return bool(re.match(r"^身体(?![0-9０-９])", service))


def apportion_service(body, living, minutes):
    """対応表(身体分,生活分)と実提供分から、実際に計上する(身体,生活)分を返す。

    身体分+生活分 をそのサービス内容の「規定提供時間」とみなす。
    実提供時間(minutes)が規定提供時間未満のときは、身体:生活の比率を保ったまま
    実提供時間を按分する（例: 身1生1=30+20=50分・比率60%:40%、実45分なら身27/生18）。
    実提供時間が規定以上のとき、または比率が定まらない/時間不明のときは固定値を用いる。

    戻り: (body_min, living_min, adjusted)  adjusted=按分したか
    """
    body = int(body or 0)
    living = int(living or 0)
    kitei = body + living
    if kitei > 0 and minutes and 0 < minutes < kitei:
        b = int(round(minutes * body / kitei))
        l = minutes - b            # 合計が実提供時間と一致するよう残りを生活へ
        return b, l, True
    return body, living, False


def _kotsu_auto_flags(visit_times, first_included):
    """各訪問が交通費対象かどうかの自動判定を bool リストで返す。

    visit_times: [(日キー, 開始分, 終了分), ...] 全訪問分（元の並び順）。
    戻りは visit_times と同じ並び順の bool リスト。
    - 直前サービスの終了時刻と当該サービスの開始時刻が同じ時刻なら、移動が無い
      ものとして交通費対象外にする（選択式ではなく全利用者共通）。
    - first_included=False のとき、各日の最初の1件を交通費対象外にする。
    日ごとに開始時刻順へ並べ替えて隣接判定する。
    """
    by_day = defaultdict(list)
    for order, (day, s, e) in enumerate(visit_times):
        by_day[day].append((s, e, order))
    flags = [False] * len(visit_times)
    for items in by_day.values():
        # 開始時刻順(不明は末尾)。同時刻は元の並び順で安定化。
        items.sort(key=lambda t: (t[0] is None, t[0] if t[0] is not None else 0, t[2]))
        prev_end = None
        for i, (s, e, order) in enumerate(items):
            chargeable = True
            if i == 0 and not first_included:
                chargeable = False            # その日の1件目は対象外
            elif i > 0 and s is not None and prev_end is not None and s == prev_end:
                chargeable = False            # 前サービス終了＝開始 → 移動なし
            flags[order] = chargeable
            prev_end = e
    return flags


def apply_kotsu_overrides(auto_flags, overrides):
    """自動判定(auto_flags)に手動の交通費あり/なし(overrides)を上書きした結果を返す。

    overrides は各訪問に対応するリスト。要素が True/False なら強制的にその値を、
    None（または範囲外・未指定）なら自動判定をそのまま用いる。
    """
    result = []
    for i, auto in enumerate(auto_flags):
        ov = overrides[i] if overrides and i < len(overrides) else None
        result.append(auto if ov is None else bool(ov))
    return result


def _kotsu_count(visit_times, first_included, overrides=None):
    """交通費を支給する件数を求める（手動上書きがあれば反映）。"""
    flags = apply_kotsu_overrides(
        _kotsu_auto_flags(visit_times, first_included), overrides)
    return sum(1 for f in flags if f)


def is_data_row(row):
    """勤務実績の「データ行」らしいかどうか。

    データ行は実施日(先頭列)に日付の数字を持ち、開始/終了/時間のいずれかに
    時刻・時間がある。見出し行(「実施日」「曜日」…)やタイトル行(「◯年◯月度」等)は
    実施日に数字が無い、あるいは時刻が無いため False になる。
    """
    if not row or len(row) <= COL_SERVICE:
        return False
    date = row[COL_DATE]
    if date is None or not re.search(r"\d", str(date)):
        return False
    return (_to_minutes_of_day(row[COL_START]) is not None
            or _to_minutes_of_day(row[COL_END]) is not None
            or parse_minutes(row[COL_DUR], row[COL_START], row[COL_END]) > 0)


def sheet_rows(ws):
    """1シートから先頭の見出し行を除いたデータ行(タプルの並び)を返す。

    見出し行の数はExcelによって異なる（0〜2行など）。先頭から最初の
    「データ行」までを見出しとみなして自動的に読み飛ばす。これにより
    見出しが無いExcel（1行目からデータ）でも全件を取り込める。
    """
    rows = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(rows):
        if is_data_row(row):
            return rows[i:]
    # データ行が判定できないときは従来どおり固定の見出し行数で切る（保険）
    return rows[HEADER_ROWS:]


def calc_sheet(ws, mapping, kotsu_first_included=True, kotsu_overrides=None):
    """1シート(=人物1人)を集計。mapping: {service:(body,living)}"""
    return calc_rows(sheet_rows(ws), mapping, kotsu_first_included, kotsu_overrides)


def calc_rows(rows, mapping, kotsu_first_included=True, kotsu_overrides=None):
    """行タプルの並び(=人物1人)を集計。mapping: {service:(body,living)}

    rows は各要素が少なくとも COL_SERVICE(8列)まで持つタプル/リスト。
    Excelシート・貼り付けデータのどちらからでも同じロジックで集計する。
    kotsu_first_included=False のとき、各日の最初の1件を交通費対象外にする。
    kotsu_overrides は各訪問(visit)の交通費あり/なしを手動で上書きするリスト
    （True/False/None、訪問と同じ並び順）。
    """
    body_min = living_min = training_min = 0
    visits = 0
    apportioned = 0
    errors = []
    visit_times = []                       # 交通費件数の判定用 (日,開始分,終了分)
    for row in rows:
        if not row or len(row) <= COL_SERVICE or row[COL_DATE] in (None, ""):
            continue
        kubun = row[COL_KUBUN] or ""
        service = row[COL_SERVICE] or ""
        if isinstance(kubun, str):
            kubun = kubun.strip()
        if isinstance(service, str):
            service = service.strip()
        minutes = parse_minutes(row[COL_DUR], row[COL_START], row[COL_END])
        visits += 1
        visit_times.append((str(row[COL_DATE]),
                            _to_minutes_of_day(row[COL_START]),
                            _to_minutes_of_day(row[COL_END])))

        if kubun == "同行":
            training_min += minutes        # 同行 → 実時間を研修時給
        elif service == "家事":
            living_min += minutes          # 家事 → 実時間を生活時給
        elif is_plain_body(service):
            body_min += minutes            # 番号なし「身体」→ 実時間を身体時給
        elif service in mapping:
            # 対応表 → 規定提供時間(身体+生活)。実提供時間が規定未満なら按分。
            b, l = mapping[service]
            b, l, adj = apportion_service(b, l, minutes)
            body_min += b
            living_min += l
            if adj:
                apportioned += 1
        else:
            errors.append(service)         # 対応表外 → エラー(計算しない)

    return {
        "visits": visits,
        "kotsu_count": _kotsu_count(visit_times, kotsu_first_included, kotsu_overrides),
        "apportioned": apportioned,
        "body_min": body_min,
        "living_min": living_min,
        "training_min": training_min,
        "total_min": body_min + living_min + training_min,
        "errors": errors,
    }


def row_detail(row, mapping):
    """1行を労働記録(勤務実績)出力用に分解する。空行なら None。

    戻り: dict(date,dow,user,service,start,end,minutes,body,living,training,error)
    区分・内容の判定は calc_rows と同じ（同行→研修/家事→生活/番号なし身体→身体/
    対応表→固定分数/それ以外→error）。
    """
    if not row or len(row) <= COL_SERVICE or row[COL_DATE] in (None, ""):
        return None
    kubun = row[COL_KUBUN] or ""
    service = row[COL_SERVICE] or ""
    if isinstance(kubun, str):
        kubun = kubun.strip()
    if isinstance(service, str):
        service = service.strip()
    minutes = parse_minutes(row[COL_DUR], row[COL_START], row[COL_END])
    body = living = training = 0
    error = False
    adjusted = False
    if kubun == "同行":
        training = minutes
    elif service == "家事":
        living = minutes
    elif is_plain_body(service):
        body = minutes
    elif service in mapping:
        b, l = mapping[service]
        body, living, adjusted = apportion_service(b, l, minutes)
    else:
        error = True
    return {
        "date": row[COL_DATE], "dow": row[COL_DOW], "user": row[COL_USER],
        "service": service, "start": row[COL_START], "end": row[COL_END],
        "minutes": minutes, "body": body, "living": living,
        "training": training, "error": error, "adjusted": adjusted,
    }


def visit_details(rows, mapping, kotsu_first_included=True, kotsu_overrides=None):
    """労働記録の訪問(visit)を1件ずつ明細化し、交通費の対象/対象外も付ける。

    戻り: list of dict。row_detail の各キーに加え
      - visit_index: 訪問の連番(0始まり)
      - kotsu_auto:  自動判定での交通費対象か
      - kotsu:       手動上書きを反映した最終的な交通費対象か
    calc_rows と同じ並び順・同じ判定なので、visit_index はそのまま
    kotsu_overrides の添字として使える。
    """
    details = []
    visit_times = []
    for row in rows:
        d = row_detail(row, mapping)
        if d is None:
            continue
        visit_times.append((str(row[COL_DATE]),
                            _to_minutes_of_day(row[COL_START]),
                            _to_minutes_of_day(row[COL_END])))
        details.append(d)
    auto = _kotsu_auto_flags(visit_times, kotsu_first_included)
    final = apply_kotsu_overrides(auto, kotsu_overrides)
    for i, d in enumerate(details):
        d["visit_index"] = i
        d["kotsu_auto"] = auto[i] if i < len(auto) else False
        d["kotsu"] = final[i] if i < len(final) else False
    return details


def parse_pasted(text):
    """Excelから貼り付けたタブ区切りテキストを行タプルの並びに変換。

    見出し行(実施日など日付でない行)や空行は自動的に読み飛ばす。
    A実施日 B曜日 C開始 D終了 E時間 F利用者名 G区分 H内容 の並びを想定。
    """
    rows = []
    for line in text.splitlines():
        if not line.strip():
            continue
        cells = line.split("\t")
        # 8列に満たない貼り付けにも耐えるようパディング
        if len(cells) < COL_SERVICE + 1:
            cells = cells + [""] * (COL_SERVICE + 1 - len(cells))
        cells = [c.strip() for c in cells]
        # 実施日に数字が無い行(見出し・注記など)は読み飛ばす
        if not re.search(r"\d", cells[COL_DATE] or ""):
            continue
        rows.append(tuple(cells))
    return rows


def yen(minutes, rate):
    """分 ÷ 60 × 時給 (四捨五入)。"""
    return int(round(minutes / 60.0 * rate))


def calc_kaizen(res, rates):
    """処遇改善手当 = 生活時間×(生活/時間) + 身体時間×(身体/時間)。"""
    return (yen(res["living_min"], rates.get("kaizen_living", 0))
            + yen(res["body_min"], rates.get("kaizen_body", 0)))


def calc_shikaku(res, shikaku_rate):
    """資格手当 = (身体+生活)時間 × 1時間あたりの支給額(四捨五入)。

    研修時給が適用される時間(同行)は資格手当の対象外とするため、
    total_min ではなく body_min + living_min で計算する。
    """
    return yen(res["body_min"] + res["living_min"], float(shikaku_rate or 0))


def build_record(person, period, res, rates, kotsu_unit,
                 shikaku_rate=0, other1=0, other2=0):
    """集計結果と各種単価から1人分の給与レコードを組み立てる。

    shikaku_rate は資格手当の「1時間あたりの支給額」。稼働時間に応じて
    自動計算し、金額(shikaku)を求める。
    """
    amt_training = yen(res["training_min"], rates["training"])
    amt_living = yen(res["living_min"], rates["living"])
    amt_body = yen(res["body_min"], rates["body"])
    amt_kaizen = calc_kaizen(res, rates)
    kotsu_count = res.get("kotsu_count", res["visits"])
    kotsu = int(round(kotsu_count * kotsu_unit))
    shikaku_rate = float(shikaku_rate or 0)
    shikaku = calc_shikaku(res, shikaku_rate)
    other1 = int(other1 or 0)
    other2 = int(other2 or 0)
    total = (amt_training + amt_living + amt_body + amt_kaizen
             + kotsu + shikaku + other1 + other2)
    return {
        "person": person, "period": period,
        "visits": res["visits"], "kotsu_count": kotsu_count,
        "total_min": res["total_min"],
        "body_min": res["body_min"], "living_min": res["living_min"],
        "training_min": res["training_min"],
        "rate_training": rates["training"], "rate_living": rates["living"],
        "rate_body": rates["body"],
        "rate_kaizen_living": rates.get("kaizen_living", 0),
        "rate_kaizen_body": rates.get("kaizen_body", 0),
        "amt_training": amt_training, "amt_living": amt_living, "amt_body": amt_body,
        "amt_kaizen": amt_kaizen,
        "kotsu": kotsu, "rate_shikaku": shikaku_rate, "shikaku": shikaku,
        "other1": other1, "other2": other2,
        "total_amount": total,
    }
