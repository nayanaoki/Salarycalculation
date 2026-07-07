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
import theme
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
        self.geometry("960x700")
        self.option_add("*Font", FONT)
        theme.apply(self)   # ライト系テックのテーマを適用

        self.workbook = None
        self.person_tabs = {}   # person -> dict(state)
        self._tab_outer = {}    # inner frame -> outer(tab) frame

        self._build_header()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.main_nb = nb

        # 各タブは縦横スクロール可能な内側フレーム。ウィンドウが小さくても
        # 内容が見切れずスクロールで参照できる。
        self.tab_employees = self._scrollable_tab(nb, "従業員管理")
        self.tab_calc = self._scrollable_tab(nb, "給与計算")
        self.tab_workdata = self._scrollable_tab(nb, "労働記録")
        self.tab_records = self._scrollable_tab(nb, "計算記録")
        self.tab_totals = self._scrollable_tab(nb, "全員総計")
        self.tab_mapping = self._scrollable_tab(nb, "対応表編集")
        self.tab_settings = self._scrollable_tab(nb, "その他項目設定")

        self._build_employees()
        self._build_calc()
        self._build_workdata()
        self._build_records()
        self._build_totals()
        self._build_mapping()
        self._build_settings()

        nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # 起動時に最新版を確認（更新がある時だけ通知）
        self.after(1000, lambda: updater.check_for_updates(self, silent=True))

    # ============================================================ ヘッダー
    def _build_header(self):
        """上部の濃紺ヘッダーバー（アクセント＋タイトル＋バージョン）。"""
        header = tk.Frame(self, bg=theme.HEADER_BG, height=46)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        tk.Frame(header, bg=theme.ACCENT, width=6).pack(side="left", fill="y")
        tk.Label(header, text="給与自動計算 システム", bg=theme.HEADER_BG,
                 fg=theme.HEADER_FG, font=("Meiryo UI", 14, "bold")).pack(
            side="left", padx=14)
        tk.Label(header, text=f"v{__version__}", bg=theme.HEADER_BG,
                 fg=theme.HEADER_SUB, font=("Meiryo UI", 10)).pack(
            side="right", padx=16)

    # ============================================================ スクロール枠
    def _scrollable_tab(self, nb, text):
        """縦横スクロール可能な内側フレームを持つタブを追加し、内側を返す。"""
        outer = ttk.Frame(nb)
        nb.add(outer, text=text)
        canvas = tk.Canvas(outer, highlightthickness=0, bg=theme.SURFACE)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        inner = ttk.Frame(canvas, padding=8)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner(_e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_inner)

        def _on_canvas(e):
            # 内容がキャンバスより小さいときは幅を合わせて余白を無くす
            canvas.itemconfigure(win, width=max(e.width, inner.winfo_reqwidth()))
        canvas.bind("<Configure>", _on_canvas)

        # マウスホイールはカーソルが乗っているキャンバスに対して効かせる
        def _wheel(e):
            canvas.yview_scroll(int(-e.delta / 120), "units")
        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        self._tab_outer[inner] = outer
        return inner

    def _select_tab(self, inner):
        self.main_nb.select(self._tab_outer[inner])

    # ============================================================ 従業員管理
    def _build_employees(self):
        t = self.tab_employees
        ttk.Label(t, text="従業員の登録・編集", font=FONT_B).pack(anchor="w")
        ttk.Label(t, text="※氏名は勤務実績のシート名(または貼り付け時の氏名)と一致させると、"
                          "資格に応じた資格手当が自動で適用されます。",
                  foreground="#666").pack(anchor="w", pady=(0, 6))

        cols = ("id", "mgmt_no", "name", "qualification")
        heads = {"id": "ID", "mgmt_no": "管理番号", "name": "氏名", "qualification": "資格"}
        widths = {"id": 40, "mgmt_no": 120, "name": 200, "qualification": 140}
        tv = ttk.Treeview(t, columns=cols, show="headings", height=13)
        for cn in cols:
            tv.heading(cn, text=heads[cn])
            tv.column(cn, width=widths[cn],
                      anchor="w" if cn in ("mgmt_no", "name", "qualification") else "center")
        tv.pack(fill="both", expand=True, pady=6)
        tv.bind("<<TreeviewSelect>>", self._emp_on_select)
        self.emp_tree = tv

        form = ttk.LabelFrame(t, text="入力", padding=8)
        form.pack(fill="x", pady=4)
        ttk.Label(form, text="管理番号").grid(row=0, column=0, sticky="e", padx=(0, 4))
        self.emp_mgmt = tk.StringVar()
        ttk.Entry(form, textvariable=self.emp_mgmt, width=14).grid(row=0, column=1, padx=(0, 12))
        ttk.Label(form, text="氏名 *").grid(row=0, column=2, sticky="e", padx=(0, 4))
        self.emp_name = tk.StringVar()
        ttk.Entry(form, textvariable=self.emp_name, width=18).grid(row=0, column=3, padx=(0, 12))
        ttk.Label(form, text="資格 *").grid(row=0, column=4, sticky="e", padx=(0, 4))
        self.emp_qual = tk.StringVar()
        ttk.Combobox(form, textvariable=self.emp_qual, values=db.QUALIFICATIONS,
                     width=12, state="readonly").grid(row=0, column=5)

        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Button(bar, text="新規追加", command=self._emp_add,
                   style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="選択を更新", command=self._emp_update).pack(side="left", padx=6)
        ttk.Button(bar, text="選択を削除", command=self._emp_delete).pack(side="left")
        ttk.Button(bar, text="入力をクリア", command=self._emp_clear).pack(side="left", padx=6)
        self.emp_status = ttk.Label(t, text="", foreground="#060")
        self.emp_status.pack(anchor="w", pady=4)
        self._refresh_employees()

    def _refresh_employees(self):
        self.emp_tree.delete(*self.emp_tree.get_children())
        for e in db.list_employees():
            self.emp_tree.insert("", "end", values=(
                e["id"], e["mgmt_no"] or "", e["name"], e["qualification"] or ""))

    def _emp_on_select(self, _e=None):
        sel = self.emp_tree.selection()
        if not sel:
            return
        _id, mgmt, name, qual = self.emp_tree.item(sel[0])["values"]
        self.emp_mgmt.set(mgmt)
        self.emp_name.set(name)
        self.emp_qual.set(qual)

    def _emp_form_values(self):
        name = self.emp_name.get().strip()
        qual = self.emp_qual.get().strip()
        if not name:
            messagebox.showwarning("入力不足", "氏名は必須です。")
            return None
        if not qual:
            messagebox.showwarning("入力不足", "資格を選択してください。")
            return None
        return self.emp_mgmt.get().strip(), name, qual

    def _emp_add(self):
        v = self._emp_form_values()
        if not v:
            return
        db.add_employee(*v)
        self._refresh_employees()
        self._refresh_paste_employees()
        self._emp_clear()
        self.emp_status.config(text="従業員を追加しました。", foreground="#060")

    def _emp_update(self):
        sel = self.emp_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "更新する従業員を一覧から選択してください。")
            return
        v = self._emp_form_values()
        if not v:
            return
        eid = int(self.emp_tree.item(sel[0])["values"][0])
        db.update_employee(eid, *v)
        self._refresh_employees()
        self._refresh_paste_employees()
        self.emp_status.config(text="従業員情報を更新しました。", foreground="#060")

    def _emp_delete(self):
        sel = self.emp_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "削除する従業員を選択してください。")
            return
        eid = int(self.emp_tree.item(sel[0])["values"][0])
        if messagebox.askyesno("確認", "選択した従業員を削除しますか？"):
            db.delete_employee(eid)
            self._refresh_employees()
            self._refresh_paste_employees()
            self._emp_clear()
            self.emp_status.config(text="削除しました。", foreground="#060")

    def _emp_clear(self):
        self.emp_mgmt.set("")
        self.emp_name.set("")
        self.emp_qual.set("")

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

        # 2段目: 処遇改善手当(時間単価) と 資格手当(円/時・初期値)
        self.var_kaizen_living = tk.StringVar(value=s["rate_kaizen_living"])
        self.var_kaizen_body = tk.StringVar(value=s["rate_kaizen_body"])
        self.var_shikaku_default = tk.StringVar(value=s["default_shikaku"])
        self._rate_in(rate, "処遇改善(生活/時)", self.var_kaizen_living, 0, row=1)
        self._rate_in(rate, "処遇改善(身体/時)", self.var_kaizen_body, 1, row=1)
        self._rate_in(rate, "資格手当(円/時)", self.var_shikaku_default, 2, row=1)

        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=4)
        ttk.Button(bar, text="計算する", command=self.calculate,
                   style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="データを貼り付けて追加",
                   command=self.open_paste_dialog).pack(side="left", padx=6)
        ttk.Button(bar, text="全員分を記録保存", command=self.save_all_records).pack(side="left", padx=6)

        self.calc_nb = ttk.Notebook(t)
        self.calc_nb.pack(fill="both", expand=True, pady=6)

        self.calc_status = ttk.Label(t, text="Excelを選択してください。", foreground="#060")
        self.calc_status.pack(fill="x")

    def _rate_in(self, parent, label, var, col, row=0):
        pady = (8, 0) if row else 0
        ttk.Label(parent, text=label).grid(row=row, column=col * 2, sticky="e",
                                           padx=(8, 2), pady=pady)
        e = ttk.Entry(parent, textvariable=var, width=7, justify="right")
        e.grid(row=row, column=col * 2 + 1, sticky="w", pady=pady)
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
                "body": float(parse_int(self.var_body.get())),
                "kaizen_living": float(parse_int(self.var_kaizen_living.get())),
                "kaizen_body": float(parse_int(self.var_kaizen_body.get()))}

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
            rows = calc.sheet_rows(ws)
            res = calc.calc_rows(rows, mapping)
            # 読み込んだ生データをそのまま保存(person+period で upsert)
            db.save_work_data(ws.title, self.var_period.get(), rows)
            self._add_person_tab(ws.title, res, rates, kotsu_unit, rows=rows)
            if res["errors"]:
                errors.append(f"[{ws.title}] {', '.join(sorted(set(res['errors'])))}")
        self._refresh_workdata()

        if errors:
            self.calc_status.config(
                text="⛔ 対応表外のサービスがあります（計算に含めていません）。"
                     "「対応表編集」タブで追加してください → " + " / ".join(errors),
                foreground="#a00")
        else:
            self.calc_status.config(text="計算完了。資格手当・その他を入力後、記録保存/PDF出力できます。",
                                    foreground="#060")

    def _add_person_tab(self, person, res, rates, kotsu_unit, rows=None):
        frame = ttk.Frame(self.calc_nb, padding=10)
        self.calc_nb.add(frame, text=person)
        st = {"person": person, "res": res, "rates": rates, "kotsu_unit": kotsu_unit,
              "rows": rows}
        self.person_tabs[person] = st

        # 従業員登録から資格を照合し、資格手当(円/時)の初期値を決める
        emp = db.get_employee_by_name(person)
        qual = (emp.get("qualification") if emp else None)
        settings = db.get_settings()
        settings["default_shikaku"] = str(parse_int(self.var_shikaku_default.get()))
        shikaku_rate = db.shikaku_rate_for(qual, settings)
        st["qual"] = qual

        # 稼働
        work = ttk.LabelFrame(frame, text="稼働に関する内容", padding=8)
        work.pack(fill="x")
        h, m = divmod(res["total_min"], 60)
        ttk.Label(work, text=f"合計勤務時間:  {h}時間{m}分  ({res['total_min']/60:.2f} 時間)").grid(
            row=0, column=0, sticky="w", pady=2)
        ttk.Label(work, text=f"訪問件数:  {res['visits']} 件").grid(row=1, column=0, sticky="w")
        qual_text = f"資格:  {qual}（資格手当 {int(round(shikaku_rate)):,} 円/時 を適用）" \
            if qual else "資格:  従業員未登録（既定の資格手当を適用）"
        ttk.Label(work, text=qual_text, foreground="#036").grid(row=2, column=0, sticky="w", pady=(2, 0))

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
        amt_k = calc.calc_kaizen(res, rates)
        kotsu = res["visits"] * kotsu_unit
        st.update(amt_t=amt_t, amt_l=amt_l, amt_b=amt_b, amt_k=amt_k, kotsu=kotsu)

        disp(1, "研修時給", res["training_min"], amt_t)
        disp(2, "生活時給", res["living_min"], amt_l)
        disp(3, "身体時給", res["body_min"], amt_b)
        disp(4, "処遇改善手当", None, amt_k)
        disp(5, f"交通費 ({res['visits']}件×{kotsu_unit}円)", None, kotsu)

        # 編集可能項目 (資格手当は保有資格に応じた円/時をプリセット・上書き可)
        st["var_shikaku"] = tk.StringVar(value=str(int(round(shikaku_rate))))
        st["var_other1"] = tk.StringVar(value="0")
        st["var_other2"] = tk.StringVar(value="0")
        self._shikaku_row(pay, 6, person)
        self._edit_row(pay, 7, "その他", st["var_other1"], person)
        self._edit_row(pay, 8, "その他2", st["var_other2"], person)

        ttk.Separator(pay, orient="horizontal").grid(row=9, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(pay, text="合計支給額", font=("Meiryo UI", 13, "bold")).grid(row=10, column=0, sticky="w", padx=10)
        st["lbl_total"] = ttk.Label(pay, text="", font=("Meiryo UI", 13, "bold"), foreground="#003366")
        st["lbl_total"].grid(row=10, column=2, sticky="e", padx=10)

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

    def _shikaku_row(self, pay, r, person):
        """資格手当行: 円/時を入力→稼働時間に応じて金額を自動計算・表示。"""
        st = self.person_tabs[person]
        ttk.Label(pay, text="資格手当", width=12, anchor="w").grid(
            row=r, column=0, sticky="w", padx=10, pady=2)
        rate_frame = ttk.Frame(pay)
        rate_frame.grid(row=r, column=1, sticky="e", padx=10)
        ttk.Entry(rate_frame, textvariable=st["var_shikaku"], width=6,
                  justify="right").pack(side="left")
        ttk.Label(rate_frame, text="円/時").pack(side="left")
        st["lbl_shikaku"] = ttk.Label(pay, text="0 円", anchor="e", width=14)
        st["lbl_shikaku"].grid(row=r, column=2, sticky="e", padx=10)
        st["var_shikaku"].trace_add("write", lambda *_a, p=person: self._update_total(p))

    def _update_total(self, person):
        st = self.person_tabs[person]
        shikaku = calc.calc_shikaku(st["res"], parse_int(st["var_shikaku"].get()))
        st["amt_shikaku"] = shikaku
        if "lbl_shikaku" in st:
            st["lbl_shikaku"].config(text=f"{shikaku:,} 円")
        other1 = parse_int(st["var_other1"].get())
        other2 = parse_int(st["var_other2"].get())
        total = (st["amt_t"] + st["amt_l"] + st["amt_b"] + st["amt_k"]
                 + st["kotsu"] + shikaku + other1 + other2)
        st["lbl_total"].config(text=f"{total:,} 円")

    def _record_from_tab(self, person):
        st = self.person_tabs[person]
        rec = calc.build_record(
            person, self.var_period.get(), st["res"], st["rates"], st["kotsu_unit"],
            shikaku_rate=parse_int(st["var_shikaku"].get()),
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

    # ------------------------------------------------------------ 貼り付け入力
    def open_paste_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("データを貼り付けて追加")
        dlg.geometry("660x540")
        dlg.transient(self)
        dlg.configure(bg=theme.SURFACE)

        ttk.Label(dlg, text="従業員を選択（または氏名を入力）", font=FONT_B).pack(
            anchor="w", padx=10, pady=(10, 2))
        row = ttk.Frame(dlg)
        row.pack(fill="x", padx=10)
        self._paste_name = tk.StringVar()
        self._paste_combo = ttk.Combobox(row, textvariable=self._paste_name, width=30)
        self._paste_combo.pack(side="left")
        self._refresh_paste_employees()

        ttk.Label(dlg, text="勤務実績の表（実施日〜内容の列）をExcelからコピーし、下欄に貼り付けてください。\n"
                            "見出し行や空行は自動で読み飛ばします。時給・交通費単価は給与計算画面の設定を使用します。",
                  foreground="#666").pack(anchor="w", padx=10, pady=(8, 2))
        txt = tk.Text(dlg, height=18, width=80, bg=theme.SURFACE, fg=theme.TEXT,
                      relief="solid", borderwidth=1, highlightthickness=1,
                      highlightbackground=theme.BORDER, highlightcolor=theme.ACCENT,
                      insertbackground=theme.TEXT, padx=6, pady=6)
        txt.pack(fill="both", expand=True, padx=10, pady=4)
        txt.focus_set()

        bar = ttk.Frame(dlg)
        bar.pack(fill="x", padx=10, pady=8)

        def do_add():
            name = self._paste_name.get().strip()
            if not name:
                messagebox.showwarning("入力不足", "氏名を入力または選択してください。", parent=dlg)
                return
            rows = calc.parse_pasted(txt.get("1.0", "end"))
            if not rows:
                messagebox.showwarning(
                    "データなし", "貼り付けデータから有効な行が読み取れませんでした。", parent=dlg)
                return
            mapping = db.get_mapping()
            rates = self._rates()
            kotsu_unit = parse_int(self.var_kotsu.get())
            res = calc.calc_rows(rows, mapping)
            # 貼り付けた生データもそのまま保存
            db.save_work_data(name, self.var_period.get(), rows)
            self._remove_person_tab(name)
            self._add_person_tab(name, res, rates, kotsu_unit, rows=rows)
            self._refresh_workdata()
            self.calc_nb.select(self.calc_nb.tabs()[-1])
            msg = f"{name} を追加しました（{res['visits']}件）。"
            color = "#060"
            if res["errors"]:
                msg += " 対応表外: " + ", ".join(sorted(set(res["errors"])))
                color = "#a00"
            self.calc_status.config(text=msg, foreground=color)
            dlg.destroy()

        ttk.Button(bar, text="この内容で追加", command=do_add,
                   style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="閉じる", command=dlg.destroy).pack(side="left", padx=6)

    def _refresh_paste_employees(self):
        combo = getattr(self, "_paste_combo", None)
        if combo is None:
            return
        try:
            if not combo.winfo_exists():
                return
            combo["values"] = [e["name"] for e in db.list_employees()]
        except tk.TclError:
            pass

    def _remove_person_tab(self, name):
        if name not in self.person_tabs:
            return
        for tab in self.calc_nb.tabs():
            if self.calc_nb.tab(tab, "text") == name:
                self.calc_nb.forget(tab)
                break
        self.person_tabs.pop(name, None)

    # ============================================================ 労働記録
    def _build_workdata(self):
        t = self.tab_workdata
        ttk.Label(t, text="労働記録（手動入力・貼り付けした勤務実績の保存データ）",
                  font=FONT_B).pack(anchor="w")
        ttk.Label(t, text="※Excel読込・貼り付けの都度、氏名＋期間ごとに自動保存されます。"
                          "アプリを閉じても残ります。労働記録(勤務実績)Excel・個人出勤簿を出力できます。",
                  foreground="#666").pack(anchor="w", pady=(0, 6))

        cols = ("id", "person", "period", "lines", "updated_at")
        heads = {"id": "ID", "person": "氏名", "period": "期間",
                 "lines": "明細行数", "updated_at": "保存日時"}
        widths = {"id": 40, "person": 140, "period": 110, "lines": 80, "updated_at": 160}
        tv = ttk.Treeview(t, columns=cols, show="headings", height=15)
        for cn in cols:
            tv.heading(cn, text=heads[cn])
            tv.column(cn, width=widths[cn],
                      anchor="w" if cn in ("person", "period") else "center")
        tv.pack(fill="both", expand=True, pady=6)
        self.wd_tree = tv

        bar = ttk.Frame(t)
        bar.pack(fill="x")
        ttk.Button(bar, text="労働記録(勤務実績)を出力",
                   command=self._wd_export_record,
                   style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="個人出勤簿をダウンロード",
                   command=self._wd_export_attendance,
                   style="Accent.TButton").pack(side="left", padx=6)
        self.wd_status = ttk.Label(t, text="", foreground="#060")
        self.wd_status.pack(anchor="w", pady=(6, 4))

        bar2 = ttk.Frame(t)
        bar2.pack(fill="x")
        ttk.Button(bar2, text="選択を給与計算へ再読込",
                   command=self._wd_reload_selected).pack(side="left")
        ttk.Button(bar2, text="すべて再読込",
                   command=self._wd_reload_all).pack(side="left", padx=6)
        ttk.Button(bar2, text="選択を削除", command=self._wd_delete).pack(side="left")
        ttk.Button(bar2, text="再表示", command=self._refresh_workdata).pack(side="left", padx=6)
        self._refresh_workdata()

    def _refresh_workdata(self):
        if not hasattr(self, "wd_tree"):
            return
        self.wd_tree.delete(*self.wd_tree.get_children())
        for d in db.list_work_data():
            self.wd_tree.insert("", "end", values=(
                d["id"], d["person"], d["period"] or "", f"{len(d['rows'])}行",
                d["updated_at"]))

    def _wd_load_into_calc(self, d):
        """保存済み取込データ1件を給与計算タブへ再読込する。"""
        rows = [tuple(r) for r in d["rows"]]
        mapping = db.get_mapping()
        rates = self._rates()
        kotsu_unit = parse_int(self.var_kotsu.get())
        res = calc.calc_rows(rows, mapping)
        self._remove_person_tab(d["person"])
        self._add_person_tab(d["person"], res, rates, kotsu_unit, rows=rows)
        return res

    def _wd_reload_selected(self):
        sel = self.wd_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "再読込するデータを選択してください。")
            return
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        d = db.get_work_data(wid)
        if not d:
            return
        if d.get("period"):
            self.var_period.set(d["period"])
        self._wd_load_into_calc(d)
        self._select_tab(self.tab_calc)
        self.calc_nb.select(self.calc_nb.tabs()[-1])
        self.calc_status.config(text=f"{d['person']} の取込データを再読込しました。",
                                foreground="#060")

    def _wd_reload_all(self):
        data = db.list_work_data()
        if not data:
            messagebox.showinfo("データなし", "保存された取込データがありません。")
            return
        for tab in self.calc_nb.tabs():
            self.calc_nb.forget(tab)
        self.person_tabs.clear()
        for d in data:
            self._wd_load_into_calc(d)
        self._select_tab(self.tab_calc)
        if self.calc_nb.tabs():
            self.calc_nb.select(self.calc_nb.tabs()[0])
        self.calc_status.config(text=f"保存された取込データ {len(data)} 件を再読込しました。",
                                foreground="#060")

    def _wd_delete(self):
        sel = self.wd_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "削除するデータを選択してください。")
            return
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        if messagebox.askyesno("確認", "選択した労働記録を削除しますか？"):
            db.delete_work_data(wid)
            self._refresh_workdata()
            self.wd_status.config(text="削除しました。", foreground="#060")

    def _wd_selected(self):
        sel = self.wd_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "出力するデータを一覧から選択してください。")
            return None
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        return db.get_work_data(wid)

    def _wd_export_record(self):
        """選択した労働記録を「勤務実績」書式の PDF で出力（元デザイン準拠）。"""
        d = self._wd_selected()
        if not d:
            return
        default = f"勤務実績_{d['person']}_{d.get('period') or ''}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(
            title="労働記録(勤務実績)の保存先", defaultextension=".pdf",
            initialfile=default, filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        try:
            pdf.export_work_record_pdf(d["person"], d.get("period"), d["rows"],
                                       db.get_mapping(), path)
        except Exception as ex:
            messagebox.showerror("出力エラー", str(ex))
            return
        self.wd_status.config(text=f"労働記録(勤務実績)を出力しました: {path}",
                              foreground="#060")
        if messagebox.askyesno("完了", f"労働記録を出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    def _wd_export_attendance(self):
        """選択した労働記録から「個人出勤簿」書式の PDF をダウンロード（元デザイン準拠）。"""
        d = self._wd_selected()
        if not d:
            return
        default = f"個人出勤簿_{d['person']}_{d.get('period') or ''}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(
            title="個人出勤簿の保存先", defaultextension=".pdf",
            initialfile=default, filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        s = db.get_settings()
        try:
            pdf.export_attendance_pdf(
                d["person"], d.get("period"), d["rows"], path,
                start=s.get("attend_start", "8:30"),
                end=s.get("attend_end", "17:30"),
                shotei=s.get("attend_shotei", "9:00"),
                post=s.get("attend_post", "訪問介護職員"),
                employ=s.get("attend_employ", "パートスタッフ"))
        except Exception as ex:
            messagebox.showerror("出力エラー", str(ex))
            return
        self.wd_status.config(text=f"個人出勤簿をダウンロードしました: {path}",
                              foreground="#060")
        if messagebox.askyesno("完了", f"個人出勤簿を出力しました。\n{path}\n\n開きますか？"):
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
        ttk.Button(bar, text="PDFで出力", command=self._rec_pdf,
                   style="Accent.TButton").pack(side="left")
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
        ttk.Button(bar, text="総計をPDF出力", command=self._totals_pdf,
                   style="Accent.TButton").pack(side="left")

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
            extra = r["shikaku"] + r["other1"] + r["other2"] + (r["amt_kaizen"] or 0)
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
        ttk.Button(bar, text="保存", command=self._map_save,
                   style="Accent.TButton").pack(side="left")
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
                 ("rate_body", "身体時給 (円)"), ("kotsu_unit", "交通費 1件単価 (円)"),
                 ("rate_kaizen_living", "処遇改善手当 (生活/時間)"),
                 ("rate_kaizen_body", "処遇改善手当 (身体/時間)"),
                 ("default_shikaku", "資格手当 初期値/未登録者 (円/時)"),
                 ("shikaku_shoninsha", "資格手当（初任者） (円/時)"),
                 ("shikaku_jitsumusha", "資格手当（実務者） (円/時)"),
                 ("shikaku_kaigofukushishi", "資格手当（介護福祉士） (円/時)"),
                 ("shikaku_sonota", "資格手当（その他） (円/時)")]
        for i, (key, label) in enumerate(items):
            ttk.Label(fr, text=label, width=28, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
            v = tk.StringVar(value=s[key])
            self.set_vars[key] = v
            ttk.Entry(fr, textvariable=v, width=12, justify="right").grid(row=i, column=1, sticky="w")

        ttk.Label(t, text="※時給は給与計算画面でも対象月ごとに変更できます。\n"
                          "※交通費は『単価 × 訪問件数』で全員一律計算されます。",
                  foreground="#666").pack(anchor="w", pady=8)

        # 個人出勤簿の既定値(文字列・時刻はそのまま保存)
        ttk.Label(t, text="個人出勤簿の既定値", font=FONT_B).pack(anchor="w", pady=(8, 4))
        afr = ttk.Frame(t)
        afr.pack(anchor="w")
        self.attend_vars = {}
        aitems = [("attend_start", "始業時間 (例 8:30)"),
                  ("attend_end", "終業時間 (例 17:30)"),
                  ("attend_shotei", "所定内勤務時間 (例 9:00)"),
                  ("attend_post", "所属"),
                  ("attend_employ", "雇用形態")]
        for i, (key, label) in enumerate(aitems):
            ttk.Label(afr, text=label, width=28, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
            v = tk.StringVar(value=s.get(key, ""))
            self.attend_vars[key] = v
            ttk.Entry(afr, textvariable=v, width=16).grid(row=i, column=1, sticky="w")
        ttk.Label(t, text="※ケア時間はデータから日別集計、ケア以外時間＝所定内−ケア時間で算出します。",
                  foreground="#666").pack(anchor="w", pady=(6, 0))

        ttk.Button(t, text="保存", command=self._settings_save,
                   style="Accent.TButton").pack(anchor="w", pady=(8, 0))
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
        # 個人出勤簿の既定値は文字列のまま保存
        for k, v in self.attend_vars.items():
            d[k] = v.get().strip()
        db.save_settings(d)
        # 計算画面へ反映
        self.var_train.set(d["rate_training"])
        self.var_living.set(d["rate_living"])
        self.var_body.set(d["rate_body"])
        self.var_kotsu.set(d["kotsu_unit"])
        self.var_kaizen_living.set(d["rate_kaizen_living"])
        self.var_kaizen_body.set(d["rate_kaizen_body"])
        self.var_shikaku_default.set(d["default_shikaku"])
        self.set_status.config(text="設定を保存しました。", foreground="#060")

    # ============================================================ 共通
    def _on_tab_change(self, _e):
        cur = self.main_nb.tab(self.main_nb.select(), "text")
        if cur == "全員総計":
            self._refresh_totals()
        elif cur == "計算記録":
            self._refresh_records()
        elif cur == "労働記録":
            self._refresh_workdata()


if __name__ == "__main__":
    App().mainloop()
