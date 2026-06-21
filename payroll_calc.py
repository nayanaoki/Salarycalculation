# -*- coding: utf-8 -*-
"""勤務実績(Excel)の集計・給与計算ロジック。"""
import re

# 列インデックス (0始まり): A実施日 B曜日 C開始 D終了 E時間 F利用者名 G区分 H内容
COL_DATE, COL_DOW, COL_START, COL_END, COL_DUR, COL_USER, COL_KUBUN, COL_SERVICE = range(8)
HEADER_ROWS = 2


def parse_minutes(dur_cell, start_cell, end_cell):
    """列E(例 '75分')から分。無ければ開始/終了から算出。"""
    if dur_cell is not None:
        m = re.search(r"(\d+)", str(dur_cell))
        if m:
            return int(m.group(1))
    try:
        s = start_cell.hour * 60 + start_cell.minute
        e = end_cell.hour * 60 + end_cell.minute
        if e >= s:
            return e - s
    except AttributeError:
        pass
    return 0


def is_plain_body(service):
    """番号なしの「身体」かどうか。

    「身体１・Ⅱ」「身体2」など番号付きは対応表の固定値で計算するため除外し、
    番号の付かない「身体」「身体・Ⅱ」等のみ実時間計算の対象とする。
    """
    return bool(re.match(r"^身体(?![0-9０-９])", service))


def calc_sheet(ws, mapping):
    """1シート(=人物1人)を集計。mapping: {service:(body,living)}"""
    body_min = living_min = training_min = 0
    visits = 0
    errors = []
    for row in list(ws.iter_rows(values_only=True))[HEADER_ROWS:]:
        if not row or row[COL_DATE] in (None, ""):
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


def yen(minutes, rate):
    """分 ÷ 60 × 時給 (四捨五入)。"""
    return int(round(minutes / 60.0 * rate))


def build_record(person, period, res, rates, kotsu_unit,
                 shikaku=0, other1=0, other2=0):
    """集計結果と各種単価から1人分の給与レコードを組み立てる。"""
    amt_training = yen(res["training_min"], rates["training"])
    amt_living = yen(res["living_min"], rates["living"])
    amt_body = yen(res["body_min"], rates["body"])
    kotsu = int(round(res["visits"] * kotsu_unit))
    shikaku = int(shikaku or 0)
    other1 = int(other1 or 0)
    other2 = int(other2 or 0)
    total = amt_training + amt_living + amt_body + kotsu + shikaku + other1 + other2
    return {
        "person": person, "period": period,
        "visits": res["visits"], "total_min": res["total_min"],
        "body_min": res["body_min"], "living_min": res["living_min"],
        "training_min": res["training_min"],
        "rate_training": rates["training"], "rate_living": rates["living"],
        "rate_body": rates["body"],
        "amt_training": amt_training, "amt_living": amt_living, "amt_body": amt_body,
        "kotsu": kotsu, "shikaku": shikaku, "other1": other1, "other2": other2,
        "total_amount": total,
    }
