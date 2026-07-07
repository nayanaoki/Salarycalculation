# -*- coding: utf-8 -*-
"""ライト系テックのUIテーマ（標準ttkのみ・依存追加なし）。

配色と ttk.Style の設定を一元管理する。App 生成直後に apply(root) を呼ぶ。
"""
from tkinter import ttk

# --- 配色 -------------------------------------------------------------
BG = "#EDF0F4"          # 画面のすき間(ガター)・スクロール領域
SURFACE = "#FFFFFF"     # パネル/カード/入力の背景
TEXT = "#1F2937"        # 文字
MUTED = "#6B7280"       # 補助文字
ACCENT = "#2563EB"      # 主色(藍)
ACCENT_DARK = "#1D4ED8"  # ホバー/押下
ACCENT_LIGHT = "#DBEAFE"  # 選択行など
HEADER_BG = "#1E293B"   # 上部ヘッダーバー(濃紺)
HEADER_FG = "#F8FAFC"
HEADER_SUB = "#94A3B8"
BORDER = "#E5E7EB"      # 枠線/区切り
HEAD_CELL = "#EEF2F7"   # Treeview 見出し
TAB_IDLE = "#E4E9F0"    # 非選択タブ
SUCCESS = "#059669"
DANGER = "#DC2626"

FONT = ("Meiryo UI", 10)
FONT_B = ("Meiryo UI", 10, "bold")
FONT_H = ("Meiryo UI", 13, "bold")


def apply(root):
    """root(Tk) に対して ttk テーマ・配色を適用する。"""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")   # 色のカスタムに必要な素直なテーマ
    except Exception:
        pass
    root.configure(bg=BG)

    style.configure(".", background=SURFACE, foreground=TEXT, font=FONT)

    # フレーム/ラベル（パネルは白で統一し、色ズレを防ぐ）
    style.configure("TFrame", background=SURFACE)
    style.configure("Gutter.TFrame", background=BG)
    style.configure("TLabel", background=SURFACE, foreground=TEXT)
    style.configure("Muted.TLabel", background=SURFACE, foreground=MUTED)
    style.configure("Head.TLabel", background=SURFACE, foreground=TEXT, font=FONT_H)
    style.configure("Success.TLabel", background=SURFACE, foreground=SUCCESS)
    style.configure("Danger.TLabel", background=SURFACE, foreground=DANGER)

    # 通常ボタン
    style.configure("TButton", background="#F3F5F9", foreground=TEXT,
                    bordercolor=BORDER, relief="flat", padding=(10, 5))
    style.map("TButton",
              background=[("active", "#E6EBF2"), ("pressed", "#DCE3EC")],
              bordercolor=[("active", ACCENT)])
    # 主ボタン(アクセント)
    style.configure("Accent.TButton", background=ACCENT, foreground="#FFFFFF",
                    bordercolor=ACCENT, relief="flat", padding=(12, 6))
    style.map("Accent.TButton",
              background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK)],
              foreground=[("disabled", "#C7D2FE")])

    # ノートブック(タブ)
    style.configure("TNotebook", background=BG, borderwidth=0,
                    tabmargins=(6, 4, 6, 0))
    style.configure("TNotebook.Tab", background=TAB_IDLE, foreground=MUTED,
                    padding=(14, 7), borderwidth=0, font=FONT)
    style.map("TNotebook.Tab",
              background=[("selected", SURFACE)],
              foreground=[("selected", ACCENT)])

    # ラベルフレーム(カード風)
    style.configure("TLabelframe", background=SURFACE, bordercolor=BORDER,
                    relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=SURFACE, foreground=ACCENT,
                    font=FONT_B)

    # 入力
    style.configure("TEntry", fieldbackground=SURFACE, foreground=TEXT,
                    bordercolor=BORDER, padding=3)
    style.map("TEntry", bordercolor=[("focus", ACCENT)])
    style.configure("TCombobox", fieldbackground=SURFACE, background=SURFACE,
                    foreground=TEXT, bordercolor=BORDER, arrowcolor=TEXT, padding=3)
    style.map("TCombobox", bordercolor=[("focus", ACCENT)],
              fieldbackground=[("readonly", SURFACE)])

    # 一覧(Treeview)
    style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE,
                    foreground=TEXT, rowheight=24, bordercolor=BORDER,
                    borderwidth=1, relief="solid")
    style.configure("Treeview.Heading", background=HEAD_CELL, foreground=TEXT,
                    relief="flat", font=FONT_B, padding=4)
    style.map("Treeview.Heading", background=[("active", "#E2E8F0")])
    style.map("Treeview", background=[("selected", ACCENT_LIGHT)],
              foreground=[("selected", TEXT)])

    # ラジオ/セパレータ/スクロールバー
    style.configure("TRadiobutton", background=SURFACE, foreground=TEXT)
    style.map("TRadiobutton", foreground=[("active", ACCENT)])
    style.configure("TSeparator", background=BORDER)
    style.configure("Vertical.TScrollbar", background="#CBD5E1",
                    troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)
    style.configure("Horizontal.TScrollbar", background="#CBD5E1",
                    troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)
    return style
