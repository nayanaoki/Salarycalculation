# -*- coding: utf-8 -*-
"""職種（介護・看護師・准看護師・理学療法士）の違いを1か所にまとめる。

統合前は職種ごとに別アプリだった。統合版では1つのアプリ・1つの給与計算タブで
全員を計算し、人ごとに職種を決めて、その職種の計算のきまりと単価を当てる。

職種の決め方（上から順に優先。payroll_app が使う）:
  1. 従業員管理に登録した職種（氏名で照合）
  2. 実績の中身から自動判定（payroll_calc.detect_job）
  3. どちらでも決まらなければ「未判定」→ 画面で選んでもらう
"""

KAIGO = "kaigo"
NURSE = "nurse"
LPN = "lpn"
PT = "pt"

ORDER = [KAIGO, NURSE, LPN, PT]

JOBS = {
    KAIGO: {
        "label": "介護",
        "long": "訪問介護",
        "qualifications": ["初任者", "実務者", "介護福祉士", "その他"],
        "unit_min": None,
        "time_rule": "対応表の身体分・生活分で数えます（実時間が規定より短いときは按分）。",
        "attend_post": "訪問介護職員",
        "attend_shotei": "9:00",
        # 単価の項目（設定キー, 表示名, 既定値）
        "rates": [("rate_training", "研修時給（同行）", "1100"),
                  ("rate_living", "生活時給", "1200"),
                  ("rate_body", "身体時給", "1500"),
                  ("rate_kaizen_living", "処遇改善（生活）", "0"),
                  ("rate_kaizen_body", "処遇改善（身体）", "0")],
    },
    NURSE: {
        "label": "看護師",
        "long": "訪問看護（看護師）",
        "qualifications": ["看護師", "保健師", "助産師"],
        "unit_min": None,
        "time_rule": "20分・30分・60分・90分の区分で数えます。",
        "attend_post": "訪問看護（看護師）",
        "attend_shotei": "8:00",
        "rates": [("rate_training", "研修時給（★同行）", "1100"),
                  ("rate_visit", "訪問時給", "1800"),
                  ("rate_kaizen", "処遇改善手当", "0")],
    },
    LPN: {
        "label": "准看護師",
        "long": "訪問看護（准看護師）",
        "qualifications": ["准看護師"],
        "unit_min": None,
        "time_rule": "20分・30分・60分・90分の区分で数えます。",
        "attend_post": "訪問看護（准看護師）",
        "attend_shotei": "8:00",
        "rates": [("rate_training", "研修時給（★同行）", "1100"),
                  ("rate_visit", "訪問時給", "1600"),
                  ("rate_kaizen", "処遇改善手当", "0")],
    },
    PT: {
        "label": "理学療法士",
        "long": "訪問リハビリ（理学療法士）",
        "qualifications": ["理学療法士", "作業療法士", "言語聴覚士"],
        "unit_min": 20,
        "time_rule": "20分＝1単位。2単位（40分）が基本です。",
        "attend_post": "訪問看護（理学療法士）",
        "attend_shotei": "8:00",
        "rates": [("rate_training", "研修時給（★同行）", "1100"),
                  ("rate_visit", "訪問時給", "2000"),
                  ("rate_kaizen", "処遇改善手当", "0")],
    },
}

NURSING = (NURSE, LPN, PT)          # 訪問看護系（実績の形式・計算のきまりが共通）


def label(job):
    return JOBS[job]["label"] if job in JOBS else "未判定"


def from_label(text):
    """表示名 → 職種キー（見つからなければ None）。"""
    for k, j in JOBS.items():
        if text == j["label"]:
            return k
    return None


def labels():
    return [JOBS[k]["label"] for k in ORDER]


def is_nursing(job):
    return job in NURSING


def unit_min(job):
    return JOBS[job]["unit_min"] if job in JOBS else None


def qualifications(job):
    return list(JOBS[job]["qualifications"]) if job in JOBS else []


def job_of_qualification(qual):
    """資格から職種を推定（看護師→看護師、介護福祉士→介護 など）。"""
    for k in ORDER:
        if qual in JOBS[k]["qualifications"] and qual != "その他":
            return k
    return None
