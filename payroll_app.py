# -*- coding: utf-8 -*-
"""
給与自動計算 専用システム
  - Excel(勤務実績)読込 → 人物ごとに給与計算
  - 計算記録を人物ごとに保管 / 全員の総計
  - 結果を PDF 出力
  - 対応表(身体・生活の割当)編集画面
  - その他項目(時給・交通費単価)設定画面
  - 交通費は 1件当たり単価 × 訪問件数
"""
import os
import sys
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import openpyxl
    import reportlab  # noqa: F401
except ImportError as ex:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("ライブラリ不足",
                         f"必要なライブラリが不足しています: {ex.name}\n\n"
                         "コマンドプロンプトで次を実行してください:\n"
                         "    pip install openpyxl reportlab")
    sys.exit(1)

import payroll_db as db
import payroll_calc as calc
import payroll_pdf as pdf
import updater
from version import __version__

FONT = ("Meiryo UI", 10)
FONT_B = ("Meiryo UI", 10, "bold")


def parse_int(s, default=0):
    try:
        return int(round(float(str(s).replace(",", "").strip())))
    except (ValueError, AttributeError):
        return default


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        db.init_db()
        self.title(f"給与自動計算 システム  v{__version__}")
        self.geometry("900x680")
        self.option_add("*Font", FONT)

        self.workbook = None
        self.person_tabs = {}   # person -> dict(state)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self.main_nb = nb

        self.tab_calc = ttk.Frame(nb, padding=8)
        self.tab_records = ttk.Frame(nb, padding=8)
        self.tab_totals = ttk.Frame(nb, padding=8)
        self.tab_mapping = ttk.Frame(nb, padding=8)
        self.tab_settings = ttk.Frame(nb, padding=8)
        nb.add(self.tab_calc, text="給与計算")
        nb.add(self.tab_records, text="計算記録")
        nb.add(self.tab_totals, text="全員総計")
        nb.add(self.tab_mapping, text="対応表編集")
        nb.add(self.tab_settings, text="その他項目設定")

        self._build_calc()
        self._build_records()
        self._build_totals()
        self._build_mapping()
        self._build_settings()

        nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # 起動時に最新版を確認（更新がある時だけ通知）
        self.after(1000, lambda: updater.check_for_updates(self, silent=True))

    # ============================================================ 給与計算
    def _build_calc(self):
        t = self.tab_calc
        top = ttk.Frame(t)
        top.pack(fill="x")
        ttk.Button(top, text="Excelを選択...", command=self.open_file).pack(side="left")
        self.file_label = ttk.Label(top, text="ファイル未選択", foreground="#666")
        self.file_label.pack(side="left", padx=10)

        s = db.get_settings()
        rate = ttk.LabelFrame(t, text="時給設定 (円)  ※対象月の時給を入力", padding=8)
        rate.pack(fill="x", pady=6)
        self.var_train = tk.StringVar(value=s["rate_training"])
        self.var_living = tk.StringVar(value=s["rate_living"])
        self.var_body = tk.StringVar(value=s["rate_body"])
        self._rate_in(rate, "研修時給", self.var_train, 0)
        self._rate_in(rate, "生活時給", self.var_living, 1)
        self._rate_in(rate, "身体時給", self.var_body, 2)
        ttk.Label(rate, text="交通費単価:").grid(row=0, column=6, sticky="e", padx=(16, 2))
        self.var_kotsu = tk.StringVar(value=s["kotsu_unit"])
        e_kotsu = ttk.Entry(rate, textvariable=self.var_kotsu, width=6, justify="right")
        e_kotsu.grid(row=0, column=7, sticky="w")
        e_kotsu.bind("<Return>", lambda _e: self.calculate())
        ttk.Label(rate, text="円/件").grid(row=0, column=8, sticky="w", padx=(2, 0))
        ttk.Label(rate, text="期間:").grid(row=0, column=9, sticky="e", padx=(16, 2))
        self.var_period = tk.StringVar(value=datetime.date.today().strftime("%Y年%m月"))
        ttk.Entry(rate, textvariable=self.var_period, width=10).grid(row=0, column=10)

        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=4)
        ttk.Button(bar, text="計算する", command=self.calculate).pack(side="left")
        ttk.Button(bar, text="全員分を記録保存", command=self.save_all_records).pack(side="left", padx=6)

        self.calc_nb = ttk.Notebook(t)
        self.calc_nb.pack(fill="both", expand=True, pady=6)

        self.calc_status = ttk.Label(t, text="Excelを選択してください。", foreground="#060")
        self.calc_status.pack(fill="x")

    def _rate_in(self, parent, label, var, col):
        ttk.Label(parent, text=label).grid(row=0, column=col * 2, sticky="e", padx=(8, 2))
        e = ttk.Entry(parent, textvariable=var, width=7, justify="right")
        e.grid(row=0, column=col * 2 + 1, sticky="w")
        e.bind("<Return>", lambda _e: self.calculate())

    def open_file(self):
        path = filedialog.askopenfilename(
            title="勤務実績Excelを選択",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("すべて", "*.*")])
        if not path:
            return
        try:
            self.workbook = openpyxl.load_workbook(path, data_only=True)
        except Exception as ex:
            messagebox.showerror("読込エラー", f"Excelの読込に失敗しました。\n{ex}")
            return
        self.file_label.config(text=path, foreground="#000")
        self.calculate()

    def _rates(self):
        return {"training": float(parse_int(self.var_train.get())),
                "living": float(parse_int(self.var_living.get())),
                "body": float(parse_int(self.var_body.get()))}

    def calculate(self):
        if self.workbook is None:
            self.calc_status.config(text="先にExcelを選択してください。", foreground="#a00")
            return
        for tab in self.calc_nb.tabs():
            self.calc_nb.forget(tab)
        self.person_tabs.clear()

        mapping = db.get_mapping()
        rates = self._rates()
        kotsu_unit = parse_int(self.var_kotsu.get())

        errors = []
        for ws in self.workbook.worksheets:
            res = calc.calc_sheet(ws, mapping)
            self._add_person_tab(ws.title, res, rates, kotsu_unit)
            if res["errors"]:
                errors.append(f"[{ws.title}] {', '.join(sorted(set(res['errors'])))}")

        if errors:
            self.calc_status.config(
                text="⛔ 対応表外のサービスがあります（計算に含めていません）。"
                     "「対応表編集」タブで追加してください → " + " / ".join(errors),
                foreground="#a00")
        else:
            self.calc_status.config(text="計算完了。資格手当・その他を入力後、記録保存/PDF出力できます。",
                                    foreground="#060")

    def _add_person_tab(self, person, res, rates, kotsu_unit):
        frame = ttk.Frame(self.calc_nb, padding=10)
        self.calc_nb.add(frame, text=person)
        st = {"person": person, "res": res, "rates": rates, "kotsu_unit": kotsu_unit}
        self.person_tabs[person] = st

        # 稼働
        work = ttk.LabelFrame(frame, text="稼働に関する内容", padding=8)
        work.pack(fill="x")
        h, m = divmod(res["total_min"], 60)
        ttk.Label(work, text=f"合計勤務時間:  {h}時間{m}分  ({res['total_min']/60:.2f} 時間)").grid(
            row=0, column=0, sticky="w", pady=2)
        ttk.Label(work, text=f"訪問件数:  {res['visits']} 件").grid(row=1, column=0, sticky="w")

        # 給与
        pay = ttk.LabelFrame(frame, text="給与に関する内容", padding=8)
        pay.pack(fill="both", expand=True, pady=6)
        for c, head in enumerate(["項目", "時間", "金額"]):
            ttk.Label(pay, text=head, font=FONT_B).grid(row=0, column=c, sticky="w", padx=10, pady=(0, 4))

        def disp(r, name, minutes, amount):
            ttk.Label(pay, text=name, width=12, anchor="w").grid(row=r, column=0, sticky="w", padx=10, pady=2)
            tt = "" if minutes is None else f"{minutes//60}時間{minutes%60}分"
            ttk.Label(pay, text=tt, width=12, anchor="e").grid(row=r, column=1, sticky="e", padx=10)
            lbl = ttk.Label(pay, text=f"{amount:,} 円", anchor="e", width=14)
            lbl.grid(row=r, column=2, sticky="e", padx=10)
            return lbl

        amt_t = calc.yen(res["training_min"], rates["training"])
        amt_l = calc.yen(res["living_min"], rates["living"])
        amt_b = calc.yen(res["body_min"], rates["body"])
        kotsu = res["visits"] * kotsu_unit
        st.update(amt_t=amt_t, amt_l=amt_l, amt_b=amt_b, kotsu=kotsu)

        disp(1, "研修時給", res["training_min"], amt_t)
        disp(2, "生活時給", res["living_min"], amt_l)
        disp(3, "身体時給", res["body_min"], amt_b)
        disp(4, f"交通費 ({res['visits']}件×{kotsu_unit}円)", None, kotsu)

        # 編集可能項目
        st["var_shikaku"] = tk.StringVar(value="0")
        st["var_other1"] = tk.StringVar(value="0")
        st["var_other2"] = tk.StringVar(value="0")
        self._edit_row(pay, 5, "資格手当", st["var_shikaku"], person)
        self._edit_row(pay, 6, "その他", st["var_other1"], person)
        self._edit_row(pay, 7, "その他2", st["var_other2"], person)

        ttk.Separator(pay, orient="horizontal").grid(row=8, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(pay, text="合計支給額", font=("Meiryo UI", 13, "bold")).grid(row=9, column=0, sticky="w", padx=10)
        st["lbl_total"] = ttk.Label(pay, text="", font=("Meiryo UI", 13, "bold"), foreground="#003366")
        st["lbl_total"].grid(row=9, column=2, sticky="e", padx=10)

        btns = ttk.Frame(frame)
        btns.pack(fill="x")
        ttk.Button(btns, text="この人物を記録保存", command=lambda p=person: self.save_record(p)).pack(side="left")
        ttk.Button(btns, text="PDF出力", command=lambda p=person: self.export_pdf(p)).pack(side="left", padx=6)

        self._update_total(person)

    def _edit_row(self, pay, r, name, var, person):
        ttk.Label(pay, text=name, width=12, anchor="w").grid(row=r, column=0, sticky="w", padx=10, pady=2)
        ttk.Label(pay, text="(入力)", foreground="#888").grid(row=r, column=1, sticky="e", padx=10)
        e = ttk.Entry(pay, textvariable=var, width=14, justify="right")
        e.grid(row=r, column=2, sticky="e", padx=10)
        var.trace_add("write", lambda *_a, p=person: self._update_total(p))

    def _update_total(self, person):
        st = self.person_tabs[person]
        shikaku = parse_int(st["var_shikaku"].get())
        other1 = parse_int(st["var_other1"].get())
        other2 = parse_int(st["var_other2"].get())
        total = st["amt_t"] + st["amt_l"] + st["amt_b"] + st["kotsu"] + shikaku + other1 + other2
        st["lbl_total"].config(text=f"{total:,} 円")

    def _record_from_tab(self, person):
        st = self.person_tabs[person]
        rec = calc.build_record(
            person, self.var_period.get(), st["res"], st["rates"], st["kotsu_unit"],
            shikaku=parse_int(st["var_shikaku"].get()),
            other1=parse_int(st["var_other1"].get()),
            other2=parse_int(st["var_other2"].get()))
        return rec

    def save_record(self, person):
        rec = self._record_from_tab(person)
        db.save_record(rec)
        self.calc_status.config(text=f"{person} の記録を保存しました。", foreground="#060")
        self._refresh_records()

    def save_all_records(self):
        if not self.person_tabs:
            return
        for person in self.person_tabs:
            db.save_record(self._record_from_tab(person))
        self.calc_status.config(text=f"全員（{len(self.person_tabs)}名）の記録を保存しました。",
                                foreground="#060")
        self._refresh_records()

    def export_pdf(self, person):
        rec = self._record_from_tab(person)
        rec["created_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        default = f"給与明細_{person}_{self.var_period.get()}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(
            title="PDF保存先", defaultextension=".pdf", initialfile=default,
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        try:
            pdf.export_record_pdf(rec, path)
        except Exception as ex:
            messagebox.showerror("PDF出力エラー", str(ex))
            return
        if messagebox.askyesno("完了", f"PDFを出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    # ============================================================ 計算記録
    def _build_records(self):
        t = self.tab_records
        ttk.Label(t, text="保存された計算記録（人物ごと）", font=FONT_B).pack(anchor="w")
        cols = ("id", "person", "period", "visits", "total_min", "total_amount", "created_at")
        heads = {"id": "ID", "person": "氏名", "period": "期間", "visits": "訪問",
                 "total_min": "勤務時間", "total_amount": "支給額", "created_at": "保存日時"}
        widths = {"id": 40, "person": 110, "period": 90, "visits": 60,
                  "total_min": 90, "total_amount": 100, "created_at": 150}
        tv = ttk.Treeview(t, columns=cols, show="headings", height=16)
        for cn in cols:
            tv.heading(cn, text=heads[cn])
            tv.column(cn, width=widths[cn], anchor="center")
        tv.pack(fill="both", expand=True, pady=6)
        self.rec_tree = tv

        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Button(bar, text="PDFで出力", command=self._rec_pdf).pack(side="left")
        ttk.Button(bar, text="削除", command=self._rec_delete).pack(side="left", padx=6)
        ttk.Button(bar, text="再読み込み", command=self._refresh_records).pack(side="left")
        self._refresh_records()

    def _refresh_records(self):
        if not hasattr(self, "rec_tree"):
            return
        self.rec_tree.delete(*self.rec_tree.get_children())
        for r in db.list_records():
            self.rec_tree.insert("", "end", values=(
                r["id"], r["person"], r["period"], f"{r['visits']}件",
                f"{r['total_min']/60:.2f}h", f"{r['total_amount']:,}円", r["created_at"]))

    def _selected_rec_id(self):
        sel = self.rec_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "記録を選択してください。")
            return None
        return int(self.rec_tree.item(sel[0])["values"][0])

    def _rec_pdf(self):
        rid = self._selected_rec_id()
        if rid is None:
            return
        rec = db.get_record(rid)
        default = f"給与明細_{rec['person']}_{rec.get('period') or ''}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default,
                                            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        pdf.export_record_pdf(rec, path)
        if messagebox.askyesno("完了", f"PDFを出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    def _rec_delete(self):
        rid = self._selected_rec_id()
        if rid is None:
            return
        if messagebox.askyesno("確認", "選択した記録を削除しますか？"):
            db.delete_record(rid)
            self._refresh_records()

    # ============================================================ 全員総計
    def _build_totals(self):
        t = self.tab_totals
        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Label(bar, text="集計対象:").pack(side="left")
        self.total_mode = tk.StringVar(value="current")
        ttk.Radiobutton(bar, text="今回の計算結果", variable=self.total_mode,
                        value="current", command=self._refresh_totals).pack(side="left")
        ttk.Radiobutton(bar, text="保存済み記録すべて", variable=self.total_mode,
                        value="records", command=self._refresh_totals).pack(side="left")
        ttk.Button(bar, text="更新", command=self._refresh_totals).pack(side="left", padx=8)
        ttk.Button(bar, text="総計をPDF出力", command=self._totals_pdf).pack(side="left")

        cols = ("person", "visits", "total_min", "amt_training", "amt_living",
                "amt_body", "kotsu", "extra", "total_amount")
        heads = {"person": "氏名", "visits": "訪問", "total_min": "勤務時間",
                 "amt_training": "研修", "amt_living": "生活", "amt_body": "身体",
                 "kotsu": "交通費", "extra": "手当等", "total_amount": "支給額"}
        tv = ttk.Treeview(t, columns=cols, show="headings", height=15)
        for cn in cols:
            tv.heading(cn, text=heads[cn])
            tv.column(cn, width=90, anchor="e")
        tv.column("person", width=110, anchor="w")
        tv.pack(fill="both", expand=True, pady=6)
        self.tot_tree = tv
        self.tot_label = ttk.Label(t, text="", font=("Meiryo UI", 13, "bold"), foreground="#003366")
        self.tot_label.pack(anchor="e")

    def _collect_total_records(self):
        if self.total_mode.get() == "records":
            return db.list_records()
        return [self._record_from_tab(p) for p in self.person_tabs]

    def _refresh_totals(self):
        recs = self._collect_total_records()
        self.tot_tree.delete(*self.tot_tree.get_children())
        tot = dict(visits=0, total_min=0, amt_training=0, amt_living=0,
                   amt_body=0, kotsu=0, total_amount=0)
        for r in recs:
            extra = r["shikaku"] + r["other1"] + r["other2"]
            self.tot_tree.insert("", "end", values=(
                r["person"], f"{r['visits']}件", f"{r['total_min']/60:.2f}h",
                f"{r['amt_training']:,}", f"{r['amt_living']:,}", f"{r['amt_body']:,}",
                f"{r['kotsu']:,}", f"{extra:,}", f"{r['total_amount']:,}"))
            for k in tot:
                tot[k] += r[k]
        self.tot_tree.insert("", "end", values=(
            "【総計】", f"{tot['visits']}件", f"{tot['total_min']/60:.2f}h",
            f"{tot['amt_training']:,}", f"{tot['amt_living']:,}", f"{tot['amt_body']:,}",
            f"{tot['kotsu']:,}", "", f"{tot['total_amount']:,}"))
        self.tot_label.config(text=f"総支給額合計: {tot['total_amount']:,} 円  （{len(recs)}名）")

    def _totals_pdf(self):
        recs = self._collect_total_records()
        if not recs:
            messagebox.showinfo("データなし", "集計対象がありません。")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                            initialfile="全員総計.pdf", filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        pdf.export_summary_pdf(recs, path)
        if messagebox.askyesno("完了", f"PDFを出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    # ============================================================ 対応表編集
    def _build_mapping(self):
        t = self.tab_mapping
        ttk.Label(t, text="サービス内容ごとの身体・生活の割当（分）。編集して保存できます。",
                  font=FONT_B).pack(anchor="w")
        ttk.Label(t, text="※「家事」「同行」「身体」（番号なし）は実時間で計算するため一覧には含まれません。",
                  foreground="#666").pack(anchor="w")
        ttk.Label(t, text="※対応表外のサービスはエラーになります。ここに追加してください。",
                  foreground="#a00").pack(anchor="w")
        cols = ("service", "body", "living")
        tv = ttk.Treeview(t, columns=cols, show="headings", height=15)
        tv.heading("service", text="サービス内容")
        tv.heading("body", text="身体(分)")
        tv.heading("living", text="生活(分)")
        tv.column("service", width=260, anchor="w")
        tv.column("body", width=100, anchor="center")
        tv.column("living", width=100, anchor="center")
        tv.pack(fill="both", expand=True, pady=6)
        tv.bind("<Double-1>", self._map_edit_cell)
        self.map_tree = tv

        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Button(bar, text="行を追加", command=self._map_add).pack(side="left")
        ttk.Button(bar, text="選択行を削除", command=self._map_del).pack(side="left", padx=6)
        ttk.Button(bar, text="保存", command=self._map_save).pack(side="left")
        ttk.Button(bar, text="既定に戻す", command=self._map_reset).pack(side="left", padx=6)
        self.map_status = ttk.Label(t, text="", foreground="#060")
        self.map_status.pack(anchor="w")
        self._refresh_mapping()

    def _refresh_mapping(self):
        self.map_tree.delete(*self.map_tree.get_children())
        for s, (b, l) in db.get_mapping().items():
            self.map_tree.insert("", "end", values=(s, b, l))

    def _map_edit_cell(self, event):
        tv = self.map_tree
        row = tv.identify_row(event.y)
        col = tv.identify_column(event.x)
        if not row:
            return
        col_idx = int(col[1:]) - 1
        x, y, w, hgt = tv.bbox(row, col)
        cur = tv.item(row)["values"][col_idx]
        ent = ttk.Entry(tv)
        ent.place(x=x, y=y, width=w, height=hgt)
        ent.insert(0, str(cur))
        ent.focus_set()

        def commit(_e=None):
            vals = list(tv.item(row)["values"])
            v = ent.get().strip()
            if col_idx in (1, 2):
                v = parse_int(v)
            vals[col_idx] = v
            tv.item(row, values=vals)
            ent.destroy()
        ent.bind("<Return>", commit)
        ent.bind("<FocusOut>", commit)

    def _map_add(self):
        self.map_tree.insert("", "end", values=("新規サービス", 0, 0))

    def _map_del(self):
        for s in self.map_tree.selection():
            self.map_tree.delete(s)

    def _map_save(self):
        mapping = {}
        for iid in self.map_tree.get_children():
            s, b, l = self.map_tree.item(iid)["values"]
            s = str(s).strip()
            if s:
                mapping[s] = (parse_int(b), parse_int(l))
        db.save_mapping(mapping)
        self.map_status.config(text="対応表を保存しました。次回の計算から反映されます。", foreground="#060")

    def _map_reset(self):
        if messagebox.askyesno("確認", "対応表を既定値に戻しますか？（現在の編集は失われます）"):
            db.save_mapping(db.DEFAULT_MAPPING)
            self._refresh_mapping()
            self.map_status.config(text="既定値に戻しました。", foreground="#060")

    # ============================================================ その他項目設定
    def _build_settings(self):
        t = self.tab_settings
        ttk.Label(t, text="その他項目・各種単価の設定", font=FONT_B).pack(anchor="w", pady=(0, 8))
        s = db.get_settings()
        fr = ttk.Frame(t)
        fr.pack(anchor="w")
        self.set_vars = {}
        items = [("rate_training", "研修時給 (円)"), ("rate_living", "生活時給 (円)"),
                 ("rate_body", "身体時給 (円)"), ("kotsu_unit", "交通費 1件単価 (円)")]
        for i, (key, label) in enumerate(items):
            ttk.Label(fr, text=label, width=20, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
            v = tk.StringVar(value=s[key])
            self.set_vars[key] = v
            ttk.Entry(fr, textvariable=v, width=12, justify="right").grid(row=i, column=1, sticky="w")

        ttk.Label(t, text="※時給は給与計算画面でも対象月ごとに変更できます。\n"
                          "※交通費は『単価 × 訪問件数』で全員一律計算されます。",
                  foreground="#666").pack(anchor="w", pady=8)
        ttk.Button(t, text="保存", command=self._settings_save).pack(anchor="w")
        self.set_status = ttk.Label(t, text="", foreground="#060")
        self.set_status.pack(anchor="w", pady=4)

        ttk.Separator(t, orient="horizontal").pack(fill="x", pady=12)
        upd = ttk.Frame(t)
        upd.pack(anchor="w")
        ttk.Label(upd, text=f"バージョン: v{__version__}", foreground="#444").pack(side="left")
        ttk.Button(upd, text="更新を確認",
                   command=lambda: updater.check_for_updates(self, silent=False)).pack(side="left", padx=12)

    def _settings_save(self):
        d = {k: str(parse_int(v.get())) for k, v in self.set_vars.items()}
        db.save_settings(d)
        # 計算画面へ反映
        self.var_train.set(d["rate_training"])
        self.var_living.set(d["rate_living"])
        self.var_body.set(d["rate_body"])
        self.var_kotsu.set(d["kotsu_unit"])
        self.set_status.config(text="設定を保存しました。", foreground="#060")

    # ============================================================ 共通
    def _on_tab_change(self, _e):
        cur = self.main_nb.tab(self.main_nb.select(), "text")
        if cur == "全員総計":
            self._refresh_totals()
        elif cur == "計算記録":
            self._refresh_records()


if __name__ == "__main__":
    App().mainloop()
