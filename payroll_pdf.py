# -*- coding: utf-8 -*-
"""給与明細・全員総計・勤務実績(労働記録)・個人出勤簿の PDF 出力。"""
import os
import re
import calendar
import datetime

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

import jobs
import payroll_calc as calc

FONT = "HeiseiKakuGo-W5"
_registered = False
_mincho = None

DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]
GRAY = colors.HexColor("#D9D9D9")   # 見出し・総計行の塗り


# ---------------------------------------------------------------- フォント
def _ensure_font():
    global _registered
    if not _registered:
        pdfmetrics.registerFont(UnicodeCIDFont(FONT))
        _registered = True


def _ensure_mincho():
    """游明朝を登録して名前を返す（埋め込まれるので他のPCでも同じ見た目）。"""
    global _mincho
    if _mincho:
        return _mincho
    win = os.environ.get("WINDIR", r"C:\Windows")
    for fname in ("yumin.ttf", "yuminl.ttf", "yumindb.ttf"):
        path = os.path.join(win, "Fonts", fname)
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("YuMincho", path))
                _mincho = "YuMincho"
                return _mincho
            except Exception:
                pass
    _ensure_font()
    _mincho = FONT
    return _mincho


# ---------------------------------------------------------------- 値の解釈
def parse_day(v):
    d = calc.parse_date(v)
    if d:
        return d.day
    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})", str(v or ""))
    if m:
        return int(m.group(2))
    m = re.search(r"^\s*(\d{1,2})\s*日?\s*$", str(v or ""))
    return int(m.group(1)) if m else None


def parse_time(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime.time):
        return v
    m = re.match(r"\s*(\d{1,2})\s*[:：]\s*(\d{1,2})", str(v))
    if m:
        return datetime.time(int(m.group(1)) % 24, int(m.group(2)))
    return None


def parse_year_month(period):
    m = re.search(r"(\d{4}).*?(\d{1,2})", str(period or ""))
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    t = datetime.date.today()
    return t.year, t.month


def _hm(minutes):
    return f"{minutes // 60}時間{minutes % 60}分"


def _clock(minutes):
    minutes = int(minutes)
    return f"{minutes // 60}:{minutes % 60:02d}"


def _fmt_time(t):
    return "" if t is None else f"{t.hour}:{t.minute:02d}"


def _cell(c, x, top, w, h, text="", size=9, align="center", border=True,
          fill=None, font=None):
    """左上(x, top)・幅w・高さh の枠にテキストを描く（yは下からの座標系）。"""
    y = top - h
    if fill is not None:
        c.setFillColor(fill)
        c.rect(x, y, w, h, stroke=1 if border else 0, fill=1)
        c.setFillColor(colors.black)
    elif border:
        c.rect(x, y, w, h, stroke=1, fill=0)
    if text != "" and text is not None:
        c.setFont(font or FONT, size)
        ty = y + h / 2 - size * 0.35
        if align == "center":
            c.drawCentredString(x + w / 2, ty, str(text))
        elif align == "left":
            c.drawString(x + 1.5 * mm, ty, str(text))
        else:
            c.drawRightString(x + w - 1.5 * mm, ty, str(text))


def _fit(c, text, font, size, width):
    """幅に収まるよう末尾を「…」で切る。"""
    text = str(text or "")
    if pdfmetrics.stringWidth(text, font, size) <= width:
        return text
    while text and pdfmetrics.stringWidth(text + "…", font, size) > width:
        text = text[:-1]
    return text + "…"


# ================================================================ 給与明細
def export_record_pdf(rec, path):
    """1人分の給与明細 PDF。"""
    fn = _ensure_mincho()
    c = canvas.Canvas(path, pagesize=A4)
    w, h = A4
    x = 25 * mm
    y = h - 25 * mm

    job = rec.get("job")
    unit = jobs.unit_min(job)
    c.setFont(fn, 18)
    c.drawString(x, y, f"給与明細（{jobs.label(job)}）")
    y -= 10 * mm
    c.setFont(fn, 11)
    c.drawString(x, y, f"氏名: {rec['person']}      対象期間: {rec.get('period') or '-'}")
    y -= 6 * mm
    c.drawString(x, y, f"作成日時: {rec.get('created_at', '')}")
    y -= 12 * mm

    c.setFont(fn, 13)
    c.drawString(x, y, "■ 稼働")
    y -= 8 * mm
    c.setFont(fn, 11)
    c.drawString(x + 5 * mm, y,
                 f"合計時間: {_hm(rec['total_min'])}  ({rec['total_min'] / 60:.2f} 時間)")
    y -= 6 * mm
    line = f"訪問件数: {rec['visits']} 件"
    if unit and rec.get("units"):
        line += f"      単位数: {rec['units']} 単位（{unit}分＝1単位）"
    c.drawString(x + 5 * mm, y, line)
    y -= 12 * mm

    c.setFont(fn, 13)
    c.drawString(x, y, "■ 給与")
    y -= 9 * mm
    r = lambda k: rec.get(k) or 0    # noqa: E731
    if job == jobs.KAIGO:
        rows = [
            ("研修時給（同行）", _hm(r("training_min")), f"@{r('rate_training'):.0f}",
             r("amt_training")),
            ("生活時給", _hm(r("living_min")), f"@{r('rate_living'):.0f}", r("amt_living")),
            ("身体時給", _hm(r("body_min")), f"@{r('rate_body'):.0f}", r("amt_body")),
            ("処遇改善手当", "",
             f"生@{r('rate_kaizen_living'):.0f}/身@{r('rate_kaizen_body'):.0f}",
             r("amt_kaizen")),
        ]
    else:
        rows = [
            ("訪問時給", _hm(r("visit_min")), f"@{r('rate_visit'):.0f}", r("amt_visit")),
            ("研修時給（同行）", _hm(r("training_min")), f"@{r('rate_training'):.0f}",
             r("amt_training")),
            ("処遇改善手当", "", f"@{r('rate_kaizen'):.0f}", r("amt_kaizen")),
        ]
    rows += [
        ("交通費", f"{rec.get('kotsu_count', rec['visits'])}件", "", rec["kotsu"]),
        ("資格手当", "", f"@{rec.get('rate_shikaku') or 0:.0f}", rec["shikaku"]),
        ("その他", "", "", rec["other1"]),
        ("その他2", "", "", rec["other2"]),
    ]
    col_item, col_time, col_rate, col_amt = x + 5 * mm, x + 50 * mm, x + 90 * mm, x + 150 * mm
    c.setFont(fn, 10)
    c.drawString(col_item, y, "項目")
    c.drawString(col_time, y, "時間")
    c.drawString(col_rate, y, "単価")
    c.drawRightString(col_amt, y, "金額")
    y -= 2 * mm
    c.line(x, y, x + 150 * mm, y)
    y -= 7 * mm
    c.setFont(fn, 11)
    for name, t, rate, amt in rows:
        c.drawString(col_item, y, name)
        c.drawString(col_time, y, t)
        c.drawString(col_rate, y, rate)
        c.drawRightString(col_amt, y, f"{amt:,} 円")
        y -= 7 * mm
    y -= 1 * mm
    c.line(x, y, x + 150 * mm, y)
    y -= 9 * mm
    c.setFont(fn, 14)
    c.drawString(col_item, y, "合計支給額")
    c.drawRightString(col_amt, y, f"{rec['total_amount']:,} 円")
    c.showPage()
    c.save()
    return path


# ================================================================ 全員総計
def export_summary_pdf(records, path, title=None):
    fn = _ensure_mincho()
    title = title or "全員総計"
    c = canvas.Canvas(path, pagesize=A4)
    w, h = A4
    x = 14 * mm
    y = h - 22 * mm
    c.setFont(fn, 18)
    c.drawString(x, y, title)
    y -= 14 * mm

    # 職種 / 氏名 は左寄せ、数字は右寄せ（xs は数字列の右端）
    headers = ["職種", "氏名", "訪問", "時間", "基本給", "研修", "交通費", "手当等", "支給額"]
    xs = [x, x + 20 * mm, x + 66 * mm, x + 84 * mm, x + 110 * mm, x + 130 * mm,
          x + 150 * mm, x + 168 * mm, x + 182 * mm]
    right = xs[-1]

    def head():
        c.setFont(fn, 9)
        for i, (hx, t) in enumerate(zip(xs, headers)):
            (c.drawString if i < 2 else c.drawRightString)(hx, y, t)

    def line(job_text, name, t, size=9):
        c.setFont(fn, size)
        c.drawString(xs[0], y, job_text)
        c.drawString(xs[1], y, name)
        c.drawRightString(xs[2], y, f"{t['visits']}件")
        c.drawRightString(xs[3], y, f"{t['total_min'] / 60:.1f}h")
        c.drawRightString(xs[4], y, f"{t['base']:,}")
        c.drawRightString(xs[5], y, f"{t['amt_training']:,}")
        c.drawRightString(xs[6], y, f"{t['kotsu']:,}")
        c.drawRightString(xs[7], y, f"{t['extra']:,}")
        c.drawRightString(xs[8], y, f"{t['total_amount']:,}")

    def zero():
        return dict(visits=0, total_min=0, base=0, amt_training=0, kotsu=0,
                    extra=0, total_amount=0)

    def add(t, r):
        t["visits"] += r.get("visits") or 0
        t["total_min"] += r.get("total_min") or 0
        t["base"] += calc.base_pay(r)
        t["amt_training"] += r.get("amt_training") or 0
        t["kotsu"] += r.get("kotsu") or 0
        t["extra"] += calc.extra_pay(r)
        t["total_amount"] += r.get("total_amount") or 0

    def next_row(step=6 * mm):
        nonlocal y
        y -= step
        if y < 25 * mm:
            c.showPage()
            y = h - 22 * mm
            head()
            y -= 8 * mm

    head()
    y -= 2 * mm
    c.line(x, y, right, y)
    y -= 6 * mm
    order = {k: i for i, k in enumerate(jobs.ORDER)}
    records = sorted(records, key=lambda r: (order.get(r.get("job"), 9),
                                             str(r.get("person"))))
    tot, sub = zero(), {}
    for r in records:
        one = zero()
        add(one, r)
        line(jobs.label(r.get("job")), _fit(c, r["person"], fn, 9, 44 * mm), one)
        add(tot, r)
        add(sub.setdefault(r.get("job"), zero()), r)
        next_row()
    if len(sub) > 1:                             # 職種が2つ以上なら小計も出す
        y -= 1 * mm
        c.line(x, y, right, y)
        next_row(6 * mm)
        for job in sorted(sub, key=lambda k: order.get(k, 9)):
            line(jobs.label(job), "小計", sub[job])
            next_row()
    y -= 1 * mm
    c.line(x, y, right, y)
    y -= 7 * mm
    line("", "総計", tot, size=11)
    y -= 12 * mm
    c.setFont(fn, 14)
    c.drawString(x, y, f"総支給額合計: {tot['total_amount']:,} 円")
    c.showPage()
    c.save()
    return path


# ================================================================ 勤務実績(労働記録)
# 先頭6列は共通: 日付 / 曜日 / 利用者様 / サービス内容 / 開始 / 終了。
# 「時間」の列は職種で変わる:
#   介護     提供 / 研修 / 身体 / 生活    （統合前の介護版と同じ）
#   訪問看護 実時間 / 計上 / 同行          （統合前の訪問看護版と同じ）
_WR_HEAD = [("日付", 12), ("曜日", 9), ("利用者様", 30), ("サービス内容", 50),
            ("開始", 13), ("終了", 13)]
_WR_TIME = {
    "kaigo": [("提供", 13, lambda d: d["minutes"]), ("研修", 11, lambda d: d["training"]),
              ("身体", 11, lambda d: d["body"]), ("生活", 11, lambda d: d["living"])],
    "nursing": [("実時間", 14, lambda d: d["minutes"]),
                ("計上", 14, lambda d: 0 if d["error"] else d["pay"]),
                ("同行", 13, lambda d: d["training"])],
}


def export_work_record_pdf(person, period, rows, mapping, path, job, use_kitei=True):
    """取込データを「勤務実績」書式で PDF 出力する（=労働記録）。"""
    fn = _ensure_mincho()
    details = calc.visit_details(job, [tuple(r) for r in rows], mapping,
                                 use_kitei=use_kitei)
    year, month = parse_year_month(period)
    tcols = _WR_TIME["kaigo" if job == jobs.KAIGO else "nursing"]
    head = list(_WR_HEAD)
    if job == jobs.KAIGO:
        head[3] = ("提供内容", 44)
    cols = head + [(t, w) for t, w, _f in tcols]
    nt = len(tcols)
    err_col = 6 + (2 if job == jobs.KAIGO else 1)     # 対応表外と書く列

    c = canvas.Canvas(path, pagesize=A4)
    pw, ph = A4
    widths = [w * mm for _, w in cols]
    x0 = (pw - sum(widths)) / 2
    xs = [x0]
    for w in widths:
        xs.append(xs[-1] + w)
    row_h, head_h = 5.6 * mm, 6 * mm
    top_margin, bottom_margin = ph - 26 * mm, 15 * mm
    tot = [0] * nt

    def page_header():
        c.setFont(fn, 16)
        c.drawCentredString(pw / 2, ph - 15 * mm, f"勤務実績（{jobs.label(job)}）")
        c.setFont(fn, 11)
        c.drawRightString(xs[-1], ph - 14 * mm, f"{year}年{month}月度")
        c.drawRightString(xs[-1], ph - 20 * mm, str(person))

    def table_header(top):
        two = head_h * 2
        for i in range(4):
            _cell(c, xs[i], top, widths[i], two, cols[i][0], size=10, fill=GRAY, font=fn)
        _cell(c, xs[4], top, widths[4] + widths[5], head_h, "時刻", size=10, fill=GRAY, font=fn)
        _cell(c, xs[6], top, sum(widths[6:]), head_h, "時間（分）",
              size=10, fill=GRAY, font=fn)
        for i in range(4, len(cols)):
            _cell(c, xs[i], top - head_h, widths[i], head_h, cols[i][0], size=10,
                  fill=GRAY, font=fn)
        return top - two

    idx, n = 0, len(details)
    while True:
        page_header()
        y = table_header(top_margin)
        drew = False
        while idx < n and (y - row_h) > bottom_margin:
            d = details[idx]
            day = parse_day(d["date"])
            nums = [f(d) or 0 for _t, _w, f in tcols]
            vals = [f"{day}日" if day else str(d["date"] or ""), d["dow"] or "",
                    _fit(c, d["user"], fn, 9, widths[2] - 3 * mm),
                    _fit(c, d["service"], fn, 9, widths[3] - 3 * mm),
                    _fmt_time(parse_time(d["start"])), _fmt_time(parse_time(d["end"]))]
            vals += [v or "" for v in nums]
            if d["error"]:
                vals[err_col] = "対応表外"
            aligns = ["center", "center", "left", "left"] + ["center"] * (2 + nt)
            for i, (v, al) in enumerate(zip(vals, aligns)):
                small = d["error"] and i == err_col
                _cell(c, xs[i], y, widths[i], row_h, v, size=7 if small else 9,
                      align=al, font=fn)
            for i, v in enumerate(nums):
                tot[i] += v
            y -= row_h
            idx += 1
            drew = True
        if idx >= n:
            _cell(c, xs[0], y, widths[0] + widths[1], row_h, "提供回数", size=9, fill=GRAY, font=fn)
            _cell(c, xs[2], y, widths[2], row_h, n, size=9, fill=GRAY, font=fn)
            _cell(c, xs[3], y, widths[3], row_h, "", size=9, fill=GRAY, font=fn)
            _cell(c, xs[4], y, widths[4] + widths[5], row_h, "総計", size=9, fill=GRAY, font=fn)
            for i, v in enumerate(tot):
                _cell(c, xs[6 + i], y, widths[6 + i], row_h, v, size=9, fill=GRAY, font=fn)
            c.showPage()
            break
        c.showPage()
        if not drew:
            break
    c.save()
    return path


# ================================================================ 個人出勤簿
_AB_ROWS = ["始業時間", "終業時間", "訪問時間", "訪問以外時間", "休憩時間",
            "所定内勤務時間", "超過勤務時間", "深夜割増時間", "休日", "欠勤"]


def export_attendance_pdf(person, period, rows, path, start="8:30", end="17:30",
                          shotei="8:00", post="", employ="パートスタッフ", job=None):
    """取込データの日別訪問時間（実時間）から「個人出勤簿」書式で PDF 出力する。"""
    fn = _ensure_mincho()
    year, month = parse_year_month(period)
    ndays = calendar.monthrange(year, month)[1]
    reiwa = year - 2018

    per_day = {}
    for row in rows:
        row = calc.clean_row(row)
        if row[calc.COL_DATE] in (None, ""):
            continue
        day = parse_day(row[calc.COL_DATE])
        if day is None:
            continue
        per_day[day] = per_day.get(day, 0) + calc.parse_minutes(
            row[calc.COL_DUR], row[calc.COL_START], row[calc.COL_END])

    st_time = parse_time(start) or datetime.time(8, 30)
    en_time = parse_time(end) or datetime.time(17, 30)
    shotei_t = parse_time(shotei) or datetime.time(8, 0)
    shotei_min = shotei_t.hour * 60 + shotei_t.minute

    c = canvas.Canvas(path, pagesize=landscape(A4))
    pw, ph = landscape(A4)
    left, right = 12 * mm, pw - 12 * mm
    usable = right - left

    c.setFont(fn, 18)
    c.drawCentredString(pw / 2, ph - 16 * mm, f"令和{reiwa}年{month}月出勤簿")

    info_top = ph - 24 * mm
    ih = 7 * mm
    lw, vw = 26 * mm, 50 * mm
    _cell(c, left, info_top, lw, ih, "従業員氏名", size=10, font=fn)
    _cell(c, left + lw, info_top, vw, ih, person, size=10, font=fn)
    _cell(c, left + lw + vw, info_top, lw, ih, "所属", size=10, font=fn)
    _cell(c, left + 2 * lw + vw, info_top, vw, ih, post, size=10, font=fn)
    info2 = info_top - ih
    _cell(c, left, info2, lw, ih, "職種", size=10, font=fn)
    _cell(c, left + lw, info2, vw, ih, jobs.label(job), size=10, font=fn)
    _cell(c, left + lw + vw, info2, lw, ih, "雇用形態", size=10, font=fn)
    _cell(c, left + 2 * lw + vw, info2, vw, ih, employ, size=10, font=fn)
    grid_top = info2 - ih - 10 * mm

    label_w, last_w = 26 * mm, 16 * mm
    day_w = (usable - label_w - last_w) / ndays
    hdr_h, row_h = 6 * mm, 7.2 * mm
    xD, xLast = left + label_w, right - last_w

    def day_x(day):
        return xD + (day - 1) * day_w

    _cell(c, left, grid_top, label_w, hdr_h * 2, "", size=8, font=fn)
    for day in range(1, ndays + 1):
        wd = datetime.date(year, month, day).weekday()
        _cell(c, day_x(day), grid_top, day_w, hdr_h, day, size=8, font=fn)
        _cell(c, day_x(day), grid_top - hdr_h, day_w, hdr_h, DOW_JP[wd], size=8, font=fn)
    _cell(c, xLast, grid_top, last_w, hdr_h * 2, "出勤日数", size=9, font=fn)

    row_top = grid_top - hdr_h * 2
    for ri, label in enumerate(_AB_ROWS):
        top = row_top - ri * row_h
        _cell(c, left, top, label_w, row_h, label, size=9, font=fn)
        for day in range(1, ndays + 1):
            work = per_day.get(day, 0)
            text = ""
            if work > 0:
                if label == "始業時間":
                    text = _fmt_time(st_time)
                elif label == "終業時間":
                    text = _fmt_time(en_time)
                elif label == "訪問時間":
                    text = _clock(min(work, 1439))
                elif label == "訪問以外時間":
                    text = _clock(max(shotei_min - work, 0))
                elif label == "所定内勤務時間":
                    text = _fmt_time(shotei_t)
            _cell(c, day_x(day), top, day_w, row_h, text, size=6.5, font=fn)
        _cell(c, xLast, top, last_w, row_h, "", size=9, font=fn)

    worked = sum(1 for v in per_day.values() if v > 0)
    total = sum(per_day.values())
    c.setFont(fn, 11)
    c.drawCentredString(xLast + last_w / 2, row_top - row_h + row_h / 2 - 4, str(worked))
    c.setFont(fn, 9)
    c.drawCentredString(xLast + last_w / 2, row_top - 3 * row_h + row_h / 2 - 3, _clock(total))

    sign_top = row_top - len(_AB_ROWS) * row_h - 12 * mm
    c.setFont(fn, 10)
    c.drawRightString(right - 44 * mm, sign_top,
                      "勤務時間及び超過勤務時間に相違なく、上記内容に同意致します。")
    bw = 20 * mm
    bx = right - 2 * bw
    _cell(c, bx, sign_top - 3 * mm, bw, 6 * mm, "本人", size=9, font=fn)
    _cell(c, bx, sign_top - 9 * mm, bw, 14 * mm, "", size=9, font=fn)
    _cell(c, bx + bw, sign_top - 3 * mm, bw, 6 * mm, "管理者", size=9, font=fn)
    _cell(c, bx + bw, sign_top - 9 * mm, bw, 14 * mm, "", size=9, font=fn)
    c.showPage()
    c.save()
    return path
