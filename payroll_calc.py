# -*- coding: utf-8 -*-
"""勤務実績(Excel)の集計・給与計算ロジック。"""
import re

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


def sheet_rows(ws):
    """1シートから見出し行を除いたデータ行(タプルの並び)を返す。"""
    return list(ws.iter_rows(values_only=True))[HEADER_ROWS:]


def calc_sheet(ws, mapping):
    """1シート(=人物1人)を集計。mapping: {service:(body,living)}"""
    return calc_rows(sheet_rows(ws), mapping)


def calc_rows(rows, mapping):
    """行タプルの並び(=人物1人)を集計。mapping: {service:(body,living)}

    rows は各要素が少なくとも COL_SERVICE(8列)まで持つタプル/リスト。
    Excelシート・貼り付けデータのどちらからでも同じロジックで集計する。
    """
    body_min = living_min = training_min = 0
    visits = 0
    errors = []
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

        if kubun == "同行":
            training_min += minutes        # 同行 → 実時間を研修時給
        elif service == "家事":
            living_min += minutes          # 家事 → 実時間を生活時給
        elif is_plain_body(service):
            body_min += minutes            # 番号なし「身体」→ 実時間を身体時給
        elif service in mapping:
            b, l = mapping[service]        # 身体1・生活など → 対応表の固定値
            body_min += b
            living_min += l
        else:
            errors.append(service)         # 対応表外 → エラー(計算しない)

    return {
        "visits": visits,
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
    if kubun == "同行":
        training = minutes
    elif service == "家事":
        living = minutes
    elif is_plain_body(service):
        body = minutes
    elif service in mapping:
        body, living = mapping[service]
    else:
        error = True
    return {
        "date": row[COL_DATE], "dow": row[COL_DOW], "user": row[COL_USER],
        "service": service, "start": row[COL_START], "end": row[COL_END],
        "minutes": minutes, "body": body, "living": living,
        "training": training, "error": error,
    }


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
    kotsu = int(round(res["visits"] * kotsu_unit))
    shikaku_rate = float(shikaku_rate or 0)
    shikaku = calc_shikaku(res, shikaku_rate)
    other1 = int(other1 or 0)
    other2 = int(other2 or 0)
    total = (amt_training + amt_living + amt_body + amt_kaizen
             + kotsu + shikaku + other1 + other2)
    return {
        "person": person, "period": period,
        "visits": res["visits"], "total_min": res["total_min"],
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
