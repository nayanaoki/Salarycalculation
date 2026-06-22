# -*- coding: utf-8 -*-
"""給与計算結果の PDF 出力。"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

FONT = "HeiseiKakuGo-W5"
_registered = False


def _ensure_font():
    global _registered
    if not _registered:
        pdfmetrics.registerFont(UnicodeCIDFont(FONT))
        _registered = True


def _hm(minutes):
    return f"{minutes // 60}時間{minutes % 60}分"


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
        ("交通費", f"{rec['visits']}件", "", rec["kotsu"]),
        ("資格手当", "", "", rec["shikaku"]),
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
