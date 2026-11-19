# -*- coding: utf-8 -*-
"""職種ごとの計算を1つの入口にまとめる（給与計算タブはここだけを呼ぶ）。

  - 介護            → calc_kaigo（区分・内容／身体分・生活分／按分）
  - 看護師・准看・PT → calc_nurse（サービス内容／規定時間／★同行）

実績の行は、どの職種でも 8 列にそろえて持つ（日付・曜日・開始・終了・分・利用者名
までは共通。7・8 列目は 介護＝区分・内容／訪問看護＝サービス内容・備考）。
同じ行でも職種によって読み方が変わるため、計算には必ず職種を渡す。

金額は職種によらず次の式で組み立てる（使わない時間は 0 なので結果は統合前と同じ）:
  基本給   = 身体分×身体時給 + 生活分×生活時給 + 訪問分×訪問時給 + 同行分×研修時給
  処遇改善 = 身体分×(身体) + 生活分×(生活) + 訪問分×(訪問)    ※同行は対象外
  資格手当 = (身体分+生活分+訪問分) × 円/時                   ※同行は対象外
"""
import re
from collections import Counter

import jobs
import calc_kaigo as kaigo
import calc_nurse as nurse

# 列の位置（日付〜利用者名は全職種で共通）
COL_DATE, COL_DOW, COL_START, COL_END, COL_DUR, COL_USER = range(6)
NCOLS = nurse.NCOLS

# 共通の道具（日付・行の整形・対象月）
clean_row = nurse.clean_row
parse_date = nurse.parse_date
month_of = nurse.month_of
filter_month = nurse.filter_month
main_month = nurse.main_month
is_data_row = nurse.is_data_row
sheet_rows = nurse.sheet_rows
parse_minutes = nurse.parse_minutes
parse_pasted = nurse.parse_pasted
normalize_service = nurse.normalize_service
yen = nurse.yen


# ---------------------------------------------------------------- 職種の自動判定
_KAIGO_KUBUN = {"介護予防", "同行", "家事", "身体介護", "身体生活", "生活援助", "生活"}
_KAIGO_NAIYOU = re.compile(r"^(身体|生活|身[0-9０-９]|家事|訪問型独自)")
_LPN = re.compile(r"准")
_PT = re.compile(r"Ⅰ[５5]|理学|作業療法|言語聴覚")
_NURSE = re.compile(r"訪看Ⅰ[１-４1-4]|正看")


def _vote(row):
    """1行から職種の手がかりを読む（決め手が無い行は None）。"""
    c7 = str(row[6] or "").strip()          # 介護=区分 / 訪問看護=サービス内容
    c8 = str(row[7] or "").strip()          # 介護=内容 / 訪問看護=備考
    s = normalize_service(c7)
    if _LPN.search(s):
        return jobs.LPN
    if _PT.search(s):
        return jobs.PT
    if _NURSE.search(s):
        return jobs.NURSE
    if c7 in _KAIGO_KUBUN or _KAIGO_NAIYOU.match(c8):
        return jobs.KAIGO
    return None


def detect_job(rows):
    """実績の中身から職種を推定する。

    戻り: (職種キー or None, 説明文)。行ごとの手がかりで多数決をとり、
    同数や手がかりなし（定期巡回だけ等）のときは None。
    """
    votes = Counter()
    for r in rows or []:
        r = clean_row(r)
        if r[COL_DATE] in (None, ""):
            continue
        v = _vote(r)
        if v:
            votes[v] += 1
    if not votes:
        return None, "実績から職種を判定できませんでした"
    (top, n), *rest = votes.most_common()
    if rest and rest[0][1] == n:
        return None, "実績に複数の職種の記録が同じ数ずつあります"
    total = sum(votes.values())
    return top, f"実績から判定（{total}件中{n}件が{jobs.label(top)}の記録）"


# ---------------------------------------------------------------- 集計
_EMPTY = {"visits": 0, "kotsu_count": 0, "total_min": 0, "training_min": 0,
          "body_min": 0, "living_min": 0, "visit_min": 0, "units": 0,
          "slots": {}, "real_count": 0, "apportioned": 0, "errors": []}


def calc_rows(job, rows, mapping, kotsu_first_included=True, kotsu_overrides=None,
              use_kitei=True):
    """1人分を職種のきまりで集計する。キーは全職種で同じ（使わないものは 0）。"""
    rows = [clean_row(r) for r in rows]
    out = dict(_EMPTY, slots={}, errors=[])
    if job == jobs.KAIGO:
        r = kaigo.calc_rows(rows, mapping, kotsu_first_included, kotsu_overrides)
    elif jobs.is_nursing(job):
        r = nurse.calc_rows(rows, mapping, kotsu_first_included, kotsu_overrides,
                            use_kitei=use_kitei, unit_min=jobs.unit_min(job))
    else:
        return out
    out.update(r)
    out["errors"] = [str(s).strip() or "（空欄）" for s in out["errors"]]
    out["total_min"] =(out["body_min"] + out["living_min"] + out["visit_min"]
                        + out["training_min"])
    return out


def _kaigo_kind(d):
    if d["error"]:
        return "対応表外"
    if d["training"]:
        return "同行"
    parts = []
    if d["body"]:
        parts.append(f"身{d['body']}")
    if d["living"]:
        parts.append(f"生{d['living']}")
    text = "・".join(parts) or "-"
    return text + ("（按分）" if d.get("adjusted") else "")


def visit_details(job, rows, mapping, kotsu_first_included=True, kotsu_overrides=None,
                  use_kitei=True):
    """訪問を1件ずつ明細化する（労働記録タブ・勤務実績PDF用）。

    共通キー: date/dow/user/service/start/end/minutes/training/pay/error/label/
              visit_index/kotsu_auto/kotsu（介護は body/living/adjusted も持つ）
    """
    rows = [clean_row(r) for r in rows]
    if job == jobs.KAIGO:
        details = kaigo.visit_details(rows, mapping, kotsu_first_included,
                                      kotsu_overrides)
        for d in details:
            d["pay"] = d["body"] + d["living"]
            d["label"] = _kaigo_kind(d)
        return details
    if jobs.is_nursing(job):
        details = nurse.visit_details(rows, mapping, kotsu_first_included,
                                      kotsu_overrides, use_kitei=use_kitei)
        for d in details:
            if d["error"]:
                d["label"] = "対応表外"
            elif d["training"]:
                d["label"] = f"同行 {d['training']}分"
            elif d["kind"] == "実時間":
                d["label"] = f"実時間 {d['pay']}分"
            else:
                d["label"] = d["kind"] or "-"
        return details
    return []


def slot_text(res, job):
    """訪問看護の時間の内訳（「30分×12 / 実時間×3」）。介護は空。"""
    if not jobs.is_nursing(job):
        return ""
    return nurse.slot_text(res, jobs.unit_min(job))


# ---------------------------------------------------------------- 金額
def amounts(res, rates, kotsu_unit):
    """集計結果と単価 → 各項目の金額。"""
    g = lambda k: float(rates.get(k) or 0)    # noqa: E731
    kotsu_count = res.get("kotsu_count", res["visits"])
    return {
        "amt_training": yen(res["training_min"], g("training")),
        "amt_body": yen(res["body_min"], g("body")),
        "amt_living": yen(res["living_min"], g("living")),
        "amt_visit": yen(res["visit_min"], g("visit")),
        "amt_kaizen": (yen(res["body_min"], g("kaizen_body"))
                       + yen(res["living_min"], g("kaizen_living"))
                       + yen(res["visit_min"], g("kaizen"))),
        "kotsu_count": kotsu_count,
        "kotsu": int(round(kotsu_count * float(kotsu_unit or 0))),
    }


def shikaku_minutes(res):
    """資格手当の対象時間（同行を除く）。"""
    return res["body_min"] + res["living_min"] + res["visit_min"]


def calc_shikaku(res, rate):
    return yen(shikaku_minutes(res), float(rate or 0))


def build_record(job, person, period, res, rates, kotsu_unit,
                 shikaku_rate=0, other1=0, other2=0):
    """集計結果と単価から1人分の給与レコードを組み立てる（全職種共通の形）。"""
    a = amounts(res, rates, kotsu_unit)
    shikaku_rate = float(shikaku_rate or 0)
    shikaku = calc_shikaku(res, shikaku_rate)
    other1 = int(other1 or 0)
    other2 = int(other2 or 0)
    total = (a["amt_training"] + a["amt_body"] + a["amt_living"] + a["amt_visit"]
             + a["amt_kaizen"] + a["kotsu"] + shikaku + other1 + other2)
    g = lambda k: float(rates.get(k) or 0)    # noqa: E731
    return {
        "job": job, "person": person, "period": period,
        "visits": res["visits"], "kotsu_count": a["kotsu_count"],
        "total_min": res["total_min"], "training_min": res["training_min"],
        "body_min": res["body_min"], "living_min": res["living_min"],
        "visit_min": res["visit_min"], "units": res.get("units", 0),
        "rate_training": g("training"), "rate_body": g("body"),
        "rate_living": g("living"), "rate_visit": g("visit"),
        "rate_kaizen": g("kaizen"), "rate_kaizen_body": g("kaizen_body"),
        "rate_kaizen_living": g("kaizen_living"),
        "amt_training": a["amt_training"], "amt_body": a["amt_body"],
        "amt_living": a["amt_living"], "amt_visit": a["amt_visit"],
        "amt_kaizen": a["amt_kaizen"], "kotsu": a["kotsu"],
        "rate_shikaku": shikaku_rate, "shikaku": shikaku,
        "other1": other1, "other2": other2, "total_amount": total,
    }


def base_pay(rec):
    """基本給（身体・生活・訪問。同行は別）。総計の表で使う。"""
    return (rec.get("amt_body") or 0) + (rec.get("amt_living") or 0) \
        + (rec.get("amt_visit") or 0)


def extra_pay(rec):
    """手当等（処遇改善・資格手当・その他）。"""
    return ((rec.get("amt_kaizen") or 0) + (rec.get("shikaku") or 0)
            + (rec.get("other1") or 0) + (rec.get("other2") or 0))
