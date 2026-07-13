# -*- coding: utf-8 -*-
"""給与計算結果・労働記録(勤務実績)・個人出勤簿の PDF 出力。"""
import os
import calendar
import datetime

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

import payroll_calc as calc
import payroll_export as pex

FONT = "HeiseiKakuGo-W5"
_registered = False
_mincho = None

DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]
GRAY = colors.HexColor("#D9D9D9")   # 見出し・総計行の塗り


def _ensure_font():
    global _registered
    if not _registered:
        pdfmetrics.registerFont(UnicodeCIDFont(FONT))
        _registered = True


def _ensure_mincho():
    """游明朝(Yu Mincho)を登録して名前を返す。無ければ明朝系CIDにフォールバック。"""
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


def _hm(minutes):
    return f"{minutes // 60}時間{minutes % 60}分"


def _fmt_time(t):
    """datetime.time → 'H:MM'。None は空。"""
    if t is None:
        return ""
    return f"{t.hour}:{t.minute:02d}"


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


def export_record_pdf(rec, path):
    """1人分の給与明細 PDF を path に出力。"""
    _ensure_font()
    c = canvas.Canvas(path, pagesize=A4)
    w, h = A4
    x = 25 * mm
    y = h - 25 * mm

    c.setFont(FONT, 18)
    c.drawString(x, y, "給与明細")
    y -= 10 * mm
    c.setFont(FONT, 11)
    c.drawString(x, y, f"氏名: {rec['person']}      対象期間: {rec.get('period') or '-'}")
    y -= 6 * mm
    c.drawString(x, y, f"作成日時: {rec.get('created_at', '')}")
    y -= 12 * mm

    # 稼働
    c.setFont(FONT, 13)
    c.drawString(x, y, "■ 稼働")
    y -= 8 * mm
    c.setFont(FONT, 11)
    c.drawString(x + 5 * mm, y, f"合計勤務時間: {_hm(rec['total_min'])}  ({rec['total_min']/60:.2f} 時間)")
    y -= 6 * mm
    c.drawString(x + 5 * mm, y, f"訪問件数: {rec['visits']} 件")
    y -= 12 * mm

    # 給与
    c.setFont(FONT, 13)
    c.drawString(x, y, "■ 給与")
    y -= 9 * mm

    rk_living = rec.get("rate_kaizen_living") or 0
    rk_body = rec.get("rate_kaizen_body") or 0
    amt_kaizen = rec.get("amt_kaizen") or 0
    rows = [
        ("研修時給", _hm(rec["training_min"]), f"@{rec['rate_training']:.0f}", rec["amt_training"]),
        ("生活時給", _hm(rec["living_min"]), f"@{rec['rate_living']:.0f}", rec["amt_living"]),
        ("身体時給", _hm(rec["body_min"]), f"@{rec['rate_body']:.0f}", rec["amt_body"]),
        ("処遇改善手当", "", f"生@{rk_living:.0f}/身@{rk_body:.0f}", amt_kaizen),
        ("交通費", f"{rec.get('kotsu_count', rec['visits'])}件", "", rec["kotsu"]),
        ("資格手当", "", f"@{rec.get('rate_shikaku') or 0:.0f}", rec["shikaku"]),
        ("その他", "", "", rec["other1"]),
        ("その他2", "", "", rec["other2"]),
    ]
    col_item = x + 5 * mm
    col_time = x + 50 * mm
    col_rate = x + 90 * mm
    col_amt = x + 150 * mm

    c.setFont(FONT, 10)
    c.drawString(col_item, y, "項目")
    c.drawString(col_time, y, "時間")
    c.drawString(col_rate, y, "単価")
    c.drawRightString(col_amt, y, "金額")
    y -= 2 * mm
    c.line(x, y, x + 150 * mm, y)
    y -= 7 * mm

    c.setFont(FONT, 11)
    for name, t, rate, amt in rows:
        c.drawString(col_item, y, name)
        c.drawString(col_time, y, t)
        c.drawString(col_rate, y, rate)
        c.drawRightString(col_amt, y, f"{amt:,} 円")
        y -= 7 * mm

    y -= 1 * mm
    c.line(x, y, x + 150 * mm, y)
    y -= 9 * mm
    c.setFont(FONT, 14)
    c.drawString(col_item, y, "合計支給額")
    c.setFillColor(colors.HexColor("#003366"))
    c.drawRightString(col_amt, y, f"{rec['total_amount']:,} 円")
    c.setFillColor(colors.black)

    c.showPage()
    c.save()
    return path


def export_summary_pdf(records, path, title="全員総計"):
    """全員分の一覧と総計を PDF 出力。"""
    _ensure_font()
    c = canvas.Canvas(path, pagesize=A4)
    w, h = A4
    x = 18 * mm
    y = h - 22 * mm

    c.setFont(FONT, 18)
    c.drawString(x, y, title)
    y -= 14 * mm

    headers = ["氏名", "訪問", "勤務時間", "研修", "生活", "身体", "交通費", "支給額"]
    xs = [x, x + 38 * mm, x + 55 * mm, x + 80 * mm, x + 102 * mm,
          x + 124 * mm, x + 146 * mm, x + 172 * mm]
    c.setFont(FONT, 9)
    for hx, head in zip(xs, headers):
        if head in ("氏名",):
            c.drawString(hx, y, head)
        else:
            c.drawRightString(hx, y, head)
    y -= 2 * mm
    c.line(x, y, x + 172 * mm, y)
    y -= 6 * mm

    tot = dict(visits=0, amt_training=0, amt_living=0, amt_body=0, kotsu=0, total_amount=0, total_min=0)
    c.setFont(FONT, 9)
    for r in records:
        c.drawString(xs[0], y, str(r["person"])[:10])
        c.drawRightString(xs[1], y, f"{r['visits']}件")
        c.drawRightString(xs[2], y, f"{r['total_min']/60:.1f}h")
        c.drawRightString(xs[3], y, f"{r['amt_training']:,}")
        c.drawRightString(xs[4], y, f"{r['amt_living']:,}")
        c.drawRightString(xs[5], y, f"{r['amt_body']:,}")
        c.drawRightString(xs[6], y, f"{r['kotsu']:,}")
        c.drawRightString(xs[7], y, f"{r['total_amount']:,}")
        for k in tot:
            tot[k] += r[k]
        y -= 6 * mm
        if y < 25 * mm:
            c.showPage()
            y = h - 22 * mm
            c.setFont(FONT, 9)

    y -= 1 * mm
    c.line(x, y, x + 172 * mm, y)
    y -= 7 * mm
    c.setFont(FONT, 11)
    c.drawString(xs[0], y, "総計")
    c.drawRightString(xs[1], y, f"{tot['visits']}件")
    c.drawRightString(xs[2], y, f"{tot['total_min']/60:.1f}h")
    c.drawRightString(xs[3], y, f"{tot['amt_training']:,}")
    c.drawRightString(xs[4], y, f"{tot['amt_living']:,}")
    c.drawRightString(xs[5], y, f"{tot['amt_body']:,}")
    c.drawRightString(xs[6], y, f"{tot['kotsu']:,}")
    c.drawRightString(xs[7], y, f"{tot['total_amount']:,}")
    y -= 12 * mm
    c.setFont(FONT, 14)
    c.setFillColor(colors.HexColor("#003366"))
    c.drawString(x, y, f"総支給額合計: {tot['total_amount']:,} 円")
    c.setFillColor(colors.black)

    c.showPage()
    c.save()
    return path


# ================================================================ 労働記録(勤務実績)
# 列: 日付 / 曜日 / 利用者様 / 提供内容 / 開始 / 終了 / 提供 / 研修 / 身体 / 生活
_WR_COLS = [("日付", 13), ("曜日", 9), ("利用者様", 30), ("提供内容", 44),
            ("開始", 13), ("終了", 13), ("提供", 13),
            ("研修", 11), ("身体", 11), ("生活", 11)]


def export_work_record_pdf(person, period, rows, mapping, path):
    """取込データを「勤務実績」書式に成型して PDF 出力する（=労働記録）。"""
    fn = _ensure_mincho()
    rows = [tuple(r) for r in rows]
    details = [d for d in (calc.row_detail(r, mapping) for r in rows) if d]
    year, month = pex.parse_year_month(period)

    c = canvas.Canvas(path, pagesize=A4)
    pw, ph = A4
    widths = [w * mm for _, w in _WR_COLS]
    table_w = sum(widths)
    x0 = (pw - table_w) / 2
    xs = [x0]
    for w in widths:
        xs.append(xs[-1] + w)

    row_h = 5.6 * mm
    head_h = 6 * mm
    top_margin = ph - 26 * mm
    bottom_margin = 15 * mm

    tot = {"prov": 0, "training": 0, "body": 0, "living": 0}

    def draw_page_header():
        c.setFont(fn, 16)
        c.drawCentredString(pw / 2, ph - 15 * mm, "勤務実績")
        c.setFont(fn, 11)
        c.drawRightString(xs[-1], ph - 14 * mm, f"{year}年{month}月度")
        c.drawRightString(xs[-1], ph - 20 * mm, str(person))

    def draw_table_header(top):
        # 上段: 日付/曜日/利用者様/提供内容 は2段ぶち抜き。時刻・時間は結合見出し
        # 見出し行はグレーで塗る
        two = head_h * 2
        for i in range(4):
            _cell(c, xs[i], top, widths[i], two, _WR_COLS[i][0], size=10,
                  fill=GRAY, font=fn)
        jikoku_w = widths[4] + widths[5]
        _cell(c, xs[4], top, jikoku_w, head_h, "時刻", size=10, fill=GRAY, font=fn)
        jikan_w = widths[6] + widths[7] + widths[8] + widths[9]
        _cell(c, xs[6], top, jikan_w, head_h, "時間", size=10, fill=GRAY, font=fn)
        low = top - head_h
        for i in (4, 5, 6, 7, 8, 9):
            _cell(c, xs[i], low, widths[i], head_h, _WR_COLS[i][0], size=10,
                  fill=GRAY, font=fn)
        return top - two

    idx = 0
    n = len(details)
    while True:
        draw_page_header()
        y = draw_table_header(top_margin)
        drew_any = False
        while idx < n and (y - row_h) > bottom_margin:
            d = details[idx]
            day = pex.parse_day(d["date"])
            st = pex.parse_time(d["start"]); en = pex.parse_time(d["end"])
            prov = d["minutes"] or (d["body"] + d["living"] + d["training"])
            vals = [f"{day}日" if day else str(d["date"] or ""),
                    d["dow"] or "", d["user"] or "", d["service"],
                    _fmt_time(st), _fmt_time(en), prov or "",
                    d["training"] or "", d["body"] or "", d["living"] or ""]
            aligns = ["center", "center", "left", "left", "center", "center",
                      "center", "center", "center", "center"]
            for i, (v, al) in enumerate(zip(vals, aligns)):
                _cell(c, xs[i], y, widths[i], row_h, v, size=9, align=al, font=fn)
            tot["prov"] += prov
            tot["training"] += d["training"]
            tot["body"] += d["body"]
            tot["living"] += d["living"]
            y -= row_h
            idx += 1
            drew_any = True

        if idx >= n:
            # 総計行もグレーで塗る
            _cell(c, xs[0], y, widths[0] + widths[1], row_h, "提供回数", size=9,
                  fill=GRAY, font=fn)
            _cell(c, xs[2], y, widths[2], row_h, n, size=9, fill=GRAY, font=fn)
            _cell(c, xs[3], y, widths[3], row_h, "", size=9, fill=GRAY, font=fn)
            _cell(c, xs[4], y, widths[4] + widths[5], row_h, "総計", size=9,
                  fill=GRAY, font=fn)
            _cell(c, xs[6], y, widths[6], row_h, tot["prov"], size=9, fill=GRAY, font=fn)
            _cell(c, xs[7], y, widths[7], row_h, tot["training"], size=9, fill=GRAY, font=fn)
            _cell(c, xs[8], y, widths[8], row_h, tot["body"], size=9, fill=GRAY, font=fn)
            _cell(c, xs[9], y, widths[9], row_h, tot["living"], size=9, fill=GRAY, font=fn)
            c.showPage()
            break
        c.showPage()
        if not drew_any:
            break

    c.save()
    return path


# ================================================================ 個人出勤簿
_AB_ROWS = ["始業時間", "終業時間", "ケア時間", "ケア以外時間", "休憩時間",
            "所定内勤務時間", "超過勤務時間", "深夜割増時間", "休日", "欠勤"]


def export_attendance_pdf(person, period, rows, path,
                          start="8:30", end="17:30", shotei="9:00",
                          post="訪問介護職員", employ="パートスタッフ"):
    """取込データの日別ケア時間から「個人出勤簿」書式で PDF 出力する。"""
    fn = _ensure_mincho()
    rows = [tuple(r) for r in rows]
    year, month = pex.parse_year_month(period)
    ndays = calendar.monthrange(year, month)[1]
    reiwa = year - 2018

    per_day = {}
    for row in rows:
        if not row or len(row) <= calc.COL_SERVICE or row[calc.COL_DATE] in (None, ""):
            continue
        day = pex.parse_day(row[calc.COL_DATE])
        if day is None:
            continue
        minutes = calc.parse_minutes(row[calc.COL_DUR], row[calc.COL_START],
                                     row[calc.COL_END])
        per_day[day] = per_day.get(day, 0) + minutes

    st_time = pex.parse_time(start) or datetime.time(8, 30)
    en_time = pex.parse_time(end) or datetime.time(17, 30)
    shotei_t = pex.parse_time(shotei) or datetime.time(9, 0)
    shotei_min = shotei_t.hour * 60 + shotei_t.minute

    c = canvas.Canvas(path, pagesize=landscape(A4))
    pw, ph = landscape(A4)
    left = 12 * mm
    right = pw - 12 * mm
    usable = right - left

    c.setFont(fn, 18)
    c.drawCentredString(pw / 2, ph - 16 * mm, f"令和{reiwa}年{month}月出勤簿")

    # 情報ブロック(氏名・所属・職種・雇用形態) ── ひとつの固まり
    info_top = ph - 24 * mm
    ih = 7 * mm
    lw, vw = 26 * mm, 46 * mm
    _cell(c, left, info_top, lw, ih, "従業員氏名", size=10, font=fn)
    _cell(c, left + lw, info_top, vw, ih, person, size=10, font=fn)
    _cell(c, left + lw + vw, info_top, lw, ih, "所属", size=10, font=fn)
    _cell(c, left + 2 * lw + vw, info_top, vw, ih, post, size=10, font=fn)
    info2 = info_top - ih
    _cell(c, left, info2, lw, ih, "職種", size=10, font=fn)
    _cell(c, left + lw, info2, vw, ih, "-", size=10, font=fn)
    _cell(c, left + lw + vw, info2, lw, ih, "雇用形態", size=10, font=fn)
    _cell(c, left + 2 * lw + vw, info2, vw, ih, employ, size=10, font=fn)
    info_bottom = info2 - ih

    # 情報ブロックと勤務グリッドの間に明確な余白をあける
    grid_top = info_bottom - 10 * mm
    label_w = 26 * mm
    last_w = 16 * mm
    day_w = (usable - label_w - last_w) / ndays
    hdr_h = 6 * mm
    row_h = 7.2 * mm

    xL = left
    xD = left + label_w
    xLast = right - last_w

    def day_x(day):
        return xD + (day - 1) * day_w

    _cell(c, xL, grid_top, label_w, hdr_h * 2, "", size=8, font=fn)
    for day in range(1, ndays + 1):
        wd = datetime.date(year, month, day).weekday()
        _cell(c, day_x(day), grid_top, day_w, hdr_h, day, size=8, font=fn)
        _cell(c, day_x(day), grid_top - hdr_h, day_w, hdr_h, DOW_JP[wd], size=8, font=fn)
    _cell(c, xLast, grid_top, last_w, hdr_h * 2, "出勤日数", size=9, font=fn)

    row_top = grid_top - hdr_h * 2
    for ri, label in enumerate(_AB_ROWS):
        top = row_top - ri * row_h
        _cell(c, xL, top, label_w, row_h, label, size=9, font=fn)
        for day in range(1, ndays + 1):
            care = per_day.get(day, 0)
            text = ""
            if care > 0:
                if label == "始業時間":
                    text = _fmt_time(st_time)
                elif label == "終業時間":
                    text = _fmt_time(en_time)
                elif label == "ケア時間":
                    text = _fmt_time(pex._minutes_to_time(min(care, 1439)))
                elif label == "ケア以外時間":
                    text = _fmt_time(pex._minutes_to_time(min(max(shotei_min - care, 0), 1439)))
                elif label == "所定内勤務時間":
                    text = _fmt_time(shotei_t)
            # 時刻(17:30等)がセル幅からはみ出さないよう小さめのフォント
            _cell(c, day_x(day), top, day_w, row_h, text, size=6.5, font=fn)
        _cell(c, xLast, top, last_w, row_h, "", size=9, font=fn)

    worked = sum(1 for d in per_day.values() if d > 0)
    total_care = sum(per_day.values())
    c.setFont(fn, 11)
    c.drawCentredString(xLast + last_w / 2, row_top - row_h + row_h / 2 - 4, str(worked))
    th, tm = divmod(int(total_care), 60)
    c.setFont(fn, 9)
    c.drawCentredString(xLast + last_w / 2, row_top - 3 * row_h + row_h / 2 - 3,
                        f"{th}:{tm:02d}")

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
