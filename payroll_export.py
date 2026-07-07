# -*- coding: utf-8 -*-
"""労働記録(勤務実績)・個人出勤簿の Excel 出力。

取込データ(手動入力・貼り付けした勤務実績)を、事業所の
「勤務実績」書式および「個人出勤簿」書式に成型して .xlsx で出力する。
"""
import re
import calendar
import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import payroll_calc as calc

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


# ---------------------------------------------------------------- 値の解釈
def parse_day(v):
    """各種表現の日付セルから「日」を取り出す。取れなければ None。"""
    if v is None:
        return None
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.day
    s = str(v)
    for pat, idx in [(r"(\d{4})-(\d{1,2})-(\d{1,2})", 3),
                     (r"(\d{1,2})\s*月\s*(\d{1,2})", 2),
                     (r"(\d{1,2})/(\d{1,2})", 2),
                     (r"(\d{1,2})\s*日", 1),
                     (r"^\s*(\d{1,2})\s*$", 1)]:
        m = re.search(pat, s)
        if m:
            return int(m.group(idx))
    return None


def parse_time(v):
    """time / 'H:MM(:SS)' 文字列を datetime.time に。取れなければ None。"""
    if v is None or v == "":
        return None
    if isinstance(v, datetime.time):
        return v
    m = re.match(r"\s*(\d{1,2})\s*[:：]\s*(\d{1,2})", str(v))
    if m:
        return datetime.time(int(m.group(1)), int(m.group(2)))
    return None


def parse_year_month(period):
    """期間文字列(例 '2026年06月')から (year, month)。不明なら当月。"""
    m = re.search(r"(\d{4}).*?(\d{1,2})", str(period or ""))
    if m:
        mon = int(m.group(2))
        if 1 <= mon <= 12:
            return int(m.group(1)), mon
    t = datetime.date.today()
    return t.year, t.month


def _minutes_to_time(minutes):
    minutes = int(round(minutes))
    return datetime.time(minutes // 60, minutes % 60)


# ---------------------------------------------------------------- 勤務実績
def export_work_record(person, period, rows, mapping, path):
    """取込データを「勤務実績」書式で .xlsx 出力する（=労働記録）。"""
    rows = [tuple(r) for r in rows]
    details = [d for d in (calc.row_detail(r, mapping) for r in rows) if d]

    wb = openpyxl.Workbook()
    ws = wb.active
    year, month = parse_year_month(period)
    ws.title = f"{year}年{month:02d}月"

    ws.merge_cells("A1:J1")
    ws["A1"] = "勤務実績"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = CENTER
    ws.merge_cells("H3:J3")
    ws["H3"] = f"{year}年{month}月度"
    ws["H3"].alignment = CENTER
    ws.merge_cells("H4:J4")
    ws["H4"] = person
    ws["H4"].font = Font(size=12, bold=True)
    ws["H4"].alignment = CENTER

    # 見出し(2段)
    ws.merge_cells("A6:A7"); ws["A6"] = "日付"
    ws.merge_cells("B6:B7"); ws["B6"] = "曜日"
    ws.merge_cells("C6:C7"); ws["C6"] = "利用者様"
    ws.merge_cells("D6:D7"); ws["D6"] = "提供内容"
    ws.merge_cells("E6:F6"); ws["E6"] = "時刻"
    ws["E7"] = "開始"; ws["F7"] = "終了"
    ws.merge_cells("G6:J6"); ws["G6"] = "時間"
    ws["G7"] = "提供"; ws["H7"] = "研修"; ws["I7"] = "身体"; ws["J7"] = "生活"
    for r in (6, 7):
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
            cell.font = Font(bold=True)

    tot = {"minutes": 0, "training": 0, "body": 0, "living": 0}
    r = 8
    for d in details:
        day = parse_day(d["date"])
        ws.cell(row=r, column=1, value=(f"{day}日" if day else (str(d["date"]) or "")))
        ws.cell(row=r, column=2, value=(d["dow"] or ""))
        ws.cell(row=r, column=3, value=(d["user"] or ""))
        ws.cell(row=r, column=4, value=d["service"])
        st = parse_time(d["start"]); en = parse_time(d["end"])
        if st:
            c = ws.cell(row=r, column=5, value=st); c.number_format = "h:mm"
        if en:
            c = ws.cell(row=r, column=6, value=en); c.number_format = "h:mm"
        prov = d["minutes"] or (d["body"] + d["living"] + d["training"])
        ws.cell(row=r, column=7, value=prov)
        if d["training"]:
            ws.cell(row=r, column=8, value=d["training"])
        if d["body"]:
            ws.cell(row=r, column=9, value=d["body"])
        if d["living"]:
            ws.cell(row=r, column=10, value=d["living"])
        tot["minutes"] += prov
        tot["training"] += d["training"]
        tot["body"] += d["body"]
        tot["living"] += d["living"]
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = LEFT if c in (3, 4) else CENTER
        r += 1

    # 総計行
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(row=r, column=1, value="提供回数")
    ws.cell(row=r, column=3, value=len(details))
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    ws.cell(row=r, column=5, value="総計")
    ws.cell(row=r, column=7, value=tot["minutes"])
    ws.cell(row=r, column=8, value=tot["training"])
    ws.cell(row=r, column=9, value=tot["body"])
    ws.cell(row=r, column=10, value=tot["living"])
    for c in range(1, 11):
        cell = ws.cell(row=r, column=c)
        cell.border = BORDER
        cell.alignment = CENTER
        cell.font = Font(bold=True)

    widths = {"A": 7, "B": 5, "C": 14, "D": 20, "E": 7, "F": 7,
              "G": 7, "H": 6, "I": 6, "J": 6}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    wb.save(path)
    return path


# ---------------------------------------------------------------- 個人出勤簿
def export_attendance_book(person, period, rows, path,
                           start="8:30", end="17:30", shotei="9:00",
                           post="訪問介護職員", employ="パートスタッフ"):
    """取込データの日別ケア時間から「個人出勤簿」書式で .xlsx 出力する。"""
    rows = [tuple(r) for r in rows]
    year, month = parse_year_month(period)
    ndays = calendar.monthrange(year, month)[1]
    reiwa = year - 2018

    # 日ごとのケア(提供)分の合計
    per_day = {}
    for row in rows:
        # 提供(ケア)実時間は分類に依らず minutes を日別に合算する
        if not row or len(row) <= calc.COL_SERVICE or row[calc.COL_DATE] in (None, ""):
            continue
        day = parse_day(row[calc.COL_DATE])
        if day is None:
            continue
        minutes = calc.parse_minutes(row[calc.COL_DUR], row[calc.COL_START],
                                     row[calc.COL_END])
        per_day[day] = per_day.get(day, 0) + minutes

    shotei_min = _time_to_minutes(shotei, 540)
    st_time = parse_time(start) or datetime.time(8, 30)
    en_time = parse_time(end) or datetime.time(17, 30)
    shotei_time = _minutes_to_time(shotei_min)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{month}月"

    last_col = 2 + ndays          # 日列は C(3)..(2+ndays)
    label_col = 34                # AH: 出勤日数など右端ラベル
    end_col = max(last_col, label_col)

    ws.merge_cells(start_row=4, start_column=1, end_row=4,
                   end_column=max(last_col, label_col - 1))
    ws.cell(row=4, column=1, value=f"令和{reiwa}年{month}月出勤簿")
    ws.cell(row=4, column=1).font = Font(size=14, bold=True)
    ws.cell(row=4, column=1).alignment = CENTER

    ws.cell(row=7, column=1, value="従業員氏名")
    ws.cell(row=7, column=3, value=person)
    ws.cell(row=7, column=8, value="所属")
    ws.cell(row=7, column=10, value=post)
    ws.cell(row=8, column=1, value="職種")
    ws.cell(row=8, column=3, value="-")
    ws.cell(row=8, column=8, value="雇用形態")
    ws.cell(row=8, column=10, value=employ)

    # 日付・曜日の行
    ws.cell(row=10, column=label_col, value="出勤日数").alignment = CENTER
    worked = 0
    total_care = 0
    for day in range(1, ndays + 1):
        col = 2 + day
        wd = datetime.date(year, month, day).weekday()
        ws.cell(row=10, column=col, value=day).alignment = CENTER
        ws.cell(row=11, column=col, value=DOW_JP[wd]).alignment = CENTER
        care = per_day.get(day, 0)
        if care > 0:
            worked += 1
            total_care += care
            c = ws.cell(row=12, column=col, value=st_time); c.number_format = "h:mm"
            c = ws.cell(row=13, column=col, value=en_time); c.number_format = "h:mm"
            c = ws.cell(row=14, column=col, value=_minutes_to_time(min(care, 1439)))
            c.number_format = "h:mm"
            other = max(shotei_min - care, 0)
            c = ws.cell(row=15, column=col, value=_minutes_to_time(min(other, 1439)))
            c.number_format = "h:mm"
            c = ws.cell(row=18, column=col, value=shotei_time); c.number_format = "h:mm"
        for r in (10, 11, 12, 13, 14, 15, 17, 18):
            ws.cell(row=r, column=col).alignment = CENTER
            ws.cell(row=r, column=col).border = BORDER

    labels = {12: "始業時間", 13: "終業時間", 14: "ケア時間",
              15: "ケア以外時間", 17: "休憩時間", 18: "所定内勤務時間",
              19: "超過勤務時間", 20: "深夜割増時間", 22: "休日", 23: "欠勤"}
    for r, lab in labels.items():
        ws.cell(row=r, column=1, value=lab).alignment = LEFT
        ws.cell(row=r, column=1).border = BORDER

    ws.cell(row=12, column=label_col, value=worked).alignment = CENTER
    th, tm = divmod(int(total_care), 60)
    ws.cell(row=14, column=label_col, value=f"{th}:{tm:02d}").alignment = CENTER
    ws.cell(row=30, column=21,
            value="勤務時間及び超過勤務時間に相違なく、上記内容に同意致します。")
    ws.cell(row=31, column=30, value="本人")
    ws.cell(row=31, column=32, value="管理者")

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 4
    for day in range(1, ndays + 1):
        ws.column_dimensions[get_column_letter(2 + day)].width = 5
    ws.column_dimensions[get_column_letter(label_col)].width = 10

    wb.save(path)
    return path


def _time_to_minutes(v, default):
    t = parse_time(v)
    if t:
        return t.hour * 60 + t.minute
    return default
