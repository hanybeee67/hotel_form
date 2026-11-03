# -*- coding: utf-8 -*-
"""거성호텔 나이트 근무 마감 자동화 툴 (GUI)
변환 로직은 converters.py에 있고, 이 파일은 화면만 담당한다."""
import datetime as dt
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import converters as cv

VERSION = "v1.2"
COPYRIGHT = "© 2026 한희. All rights reserved."  # 저작권자 표시 (사용자 지시 2026-11-02, 눈에 띄지 않게 창 오른쪽 아래)
WEEKDAYS = "월화수목금토일"
FILETYPES = [("엑셀/표 파일", "*.xlsx *.xlsm *.xls *.csv *.htm *.html"), ("모든 파일", "*.*")]

# 서류 카드: (종류, 제목, 설명, ERP 메뉴 위치)
DOCS = [
    ("income", "수입일보", "객실 순 · 묶음별 계 (마지막 계 = ERP 계)\n누계 = 현재 잔액 · 선/후 자동 채움",
     "영업보고서 → 마감일보 → 수입일보 → 엑셀전환"),
    ("card", "카드 · 현금영수증 승인내역", "승인 '성공' 건만 · 객실 자동 채움\n한벽루 빈 줄 4개 · 현금영수증 · 계",
     "① 객실관리 → 카드결제 → 승인 내역조회\n② (선택) 객실관리 → 조회 및 출력 → 객실조회"),
    ("credit", "대외후불 리스트", "객실 순 · 금액 합계\n원본 합계와 자동 대조",
     "영업보고서 → 입금관련 → 대외 후불 리스트"),
    ("breakfast", "조식명단", "'조식/조포' 객실만 (룸온리 제외)\n객실 순 · 맨 아래 인원 합계",
     "영업보고서 → 객실/예약 → Clerk → IN HOUSE LIST"),
]
# 5번: 조식 인원수 — 파일 두 개(IN HOUSE LIST + 객실 예약 리스트)를 합쳐 내일 아침 인원만 센다 (2026-10-27)
BF_COUNT = ("bf_count", "조식 인원수 (내일 아침)",
            "체크인한 손님 + 오늘 도착할 손님 → 내일 아침 조식 인원과 조식명단",
            "IN HOUSE LIST → 객실조회(또는 객실 예약 리스트) 두 파일을 차례로 선택")

# ---------------------------------------------------------------- 디자인
C = {
    "bg": "#F3F5F9", "card": "#FFFFFF", "line": "#D5DBE5", "head": "#1F2A44", "head_sub": "#D3DAEA",
    "text": "#111827", "sub": "#374151", "hint": "#5B6474", "accent": "#2F6FEB", "accent_hover": "#2559C4",
    "soft": "#EAF1FE", "soft_hover": "#D7E5FD", "ok": "#15803D", "warn": "#B45309", "err": "#B91C1C",
    "ok_bg": "#E8F6EC", "warn_bg": "#FEF3E2", "err_bg": "#FDECEC", "disabled": "#C7CDD8",
}
BADGE = {"income": "#2F6FEB", "card": "#7C3AED", "credit": "#0E9F6E", "breakfast": "#E08A00", "bf_count": "#D9480F"}

if sys.platform.startswith("win"):
    FONT = "맑은 고딕"
    MONO = "Consolas"
else:
    FONT = "WenQuanYi Zen Hei"
    MONO = "DejaVu Sans Mono"


# 글자 크기 배율 (사용자 요청 2026-11-02: 글씨가 작아 읽기 힘듦 → 기본 1.35배, 화면 위 [가-]/[가+]로 조절·기억)
UI_SCALE = max(1.0, min(1.9, float(cv.load_user_settings().get("ui_scale", 1.35))))


def S(size):
    """배율을 적용한 글자 크기 (가장 작은 글씨도 11pt 이상)"""
    return max(11, round(size * UI_SCALE))


def F(size, bold=False, *extra):
    style = " ".join((["bold"] if bold else []) + list(extra))
    return (FONT, S(size), style) if style else (FONT, S(size))


def enable_dpi_awareness():
    """Windows 고해상도 화면에서 글자가 흐리게 보이지 않게 한다."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def open_path(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class FlatButton(tk.Label):
    """색·마우스 올림 효과를 직접 그리는 평면 버튼 (Windows 기본 버튼보다 깔끔하게 보이도록)."""

    def __init__(self, master, text, command, kind="primary", size=10, padx=16, pady=8, **kw):
        styles = {
            "primary": (C["accent"], "#FFFFFF", C["accent_hover"]),
            "soft": (C["soft"], C["accent"], C["soft_hover"]),
            "ghost": (C["card"], C["sub"], C["bg"]),
            "head": (C["head"], "#FFFFFF", "#33405F"),
        }
        self.bg, self.fg, self.hover = styles[kind]
        super().__init__(master, text=text, bg=self.bg, fg=self.fg, font=F(size, True),
                         padx=padx, pady=pady, cursor="hand2", **kw)
        self.command = command
        self.enabled = True
        self.bind("<Enter>", lambda e: self.enabled and self.configure(bg=self.hover))
        self.bind("<Leave>", lambda e: self.configure(bg=self.bg if self.enabled else C["disabled"]))
        self.bind("<ButtonRelease-1>", self._click)

    def _click(self, _e):
        if self.enabled and self.command:
            self.command()

    def set_enabled(self, on):
        self.enabled = on
        self.configure(bg=self.bg if on else C["disabled"], fg=self.fg if on else "#FFFFFF",
                       cursor="hand2" if on else "arrow")


class DocCard(tk.Frame):
    """서류 한 종류를 나타내는 카드: 번호 배지 · 제목 · 설명 · 상태 · 버튼."""

    def __init__(self, master, index, kind, title, desc, menu, on_run, on_open, compact=False):
        super().__init__(master, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        self.kind, self.on_open, self.output = kind, on_open, None
        tk.Frame(self, bg=BADGE[kind], width=5).pack(side="left", fill="y")
        body = tk.Frame(self, bg=C["card"], padx=16, pady=8 if compact else 14)
        body.pack(side="left", fill="both", expand=True)

        top = tk.Frame(body, bg=C["card"])
        top.pack(fill="x")
        tk.Label(top, text=str(index), bg=BADGE[kind], fg="#FFFFFF", font=F(10, True),
                 width=2).pack(side="left")
        tk.Label(top, text=title, bg=C["card"], fg=C["text"], font=F(12, True)).pack(side="left", padx=(8, 0))

        tk.Label(body, text=desc, bg=C["card"], fg=C["sub"], font=F(9), justify="left",
                 wraplength=int((300 if kind != "bf_count" else 760) * UI_SCALE),
                 anchor="w").pack(fill="x", pady=(8, 2))
        if not compact:  # 작은 화면에서는 ERP 메뉴 안내 줄을 숨긴다 (오늘 마감 확인 창에도 있음)
            tk.Label(body, text=("" if kind == "bf_count" else "ERP  ") + menu, bg=C["card"], fg=C["hint"], font=F(8),
                     anchor="w", justify="left", wraplength=int((300 if kind != "bf_count" else 760) * UI_SCALE)).pack(fill="x")

        bottom = tk.Frame(body, bg=C["card"])
        bottom.pack(fill="x", pady=(12, 0))
        self.status = tk.Label(bottom, text="대기 중", bg=C["bg"], fg=C["sub"], font=F(9), padx=8, pady=3)
        self.status.pack(side="left")
        self.open_link = tk.Label(bottom, text="결과 열기", bg=C["card"], fg=C["accent"],
                                  font=F(9, False, "underline"), cursor="hand2")
        self.open_link.bind("<Button-1>", lambda e: self.output and self.on_open(self.output))
        self.button = FlatButton(bottom, "파일 선택", lambda: on_run(kind), kind="soft", size=9, padx=14, pady=6)
        self.button.pack(side="right")

    def set_state(self, state, text, output=None):
        colors = {"idle": (C["bg"], C["sub"]), "run": (C["soft"], C["accent"]), "ok": (C["ok_bg"], C["ok"]),
                  "warn": (C["warn_bg"], C["warn"]), "err": (C["err_bg"], C["err"])}
        bg, fg = colors[state]
        self.status.configure(text=text, bg=bg, fg=fg)
        self.output = output
        if output:
            self.open_link.pack(side="left", padx=(10, 0))
        else:
            self.open_link.pack_forget()


class DailyDialog(tk.Toplevel):
    """오늘 마감 ERP 파일 확인 창: 파일마다 ERP 메뉴·찾은 파일(✔/✖)·[선택], 아래에 만들어질 서류 1~5번"""

    def __init__(self, app, folder):
        super().__init__(app, bg=C["bg"])
        self.app, self.folder = app, folder
        self.title("오늘 마감 — ERP 파일 확인")
        self.transient(app)
        self.resizable(True, False)
        names = app.cfg.get("daily", {}).get("names", {})
        self.mmdd = f"{app.bdate:%m%d}"
        head = tk.Frame(self, bg=C["head"], padx=18, pady=12)
        head.pack(fill="x")
        tk.Label(head, text=f"영업일자 {app.bdate:%Y-%m-%d} ({WEEKDAYS[app.bdate.weekday()]}) 마감 파일", bg=C["head"],
                 fg="#FFFFFF", font=F(13, True)).pack(anchor="w")
        tk.Label(head, text=f"ERP에서 엑셀전환한 파일을 ERP 원본 폴더에 '이름{self.mmdd}'로 저장하면 [전체 불러오기]로 한 번에 찾습니다. "
                            "없거나 다른 파일이면 [선택]", bg=C["head"], fg=C["head_sub"], font=F(9)).pack(anchor="w", pady=(2, 0))
        # ERP 원본 폴더 + [전체 불러오기] (사용자 요청 2026-11-03: 공유폴더에서 그날 파일을 한 번에 불러와 버튼 하나로)
        fbar = tk.Frame(self, bg=C["card"], padx=16, pady=10, highlightthickness=1, highlightbackground=C["line"])
        fbar.pack(fill="x", padx=14, pady=(12, 0))
        tk.Label(fbar, text="ERP 원본 폴더", bg=C["card"], fg=C["text"], font=F(10, True)).pack(side="left")
        self.folder_lbl = tk.Label(fbar, text="", bg=C["card"], font=F(9), anchor="w")
        self.folder_lbl.pack(side="left", padx=(10, 0), fill="x", expand=True)
        FlatButton(fbar, "폴더 바꾸기", self.change_folder, kind="soft", size=9, padx=12, pady=4).pack(side="right")
        FlatButton(fbar, f"전체 불러오기 ({self.mmdd})", self.load_all, size=10, padx=14, pady=5).pack(side="right", padx=(0, 8))
        self.folder_warn = tk.Label(self, text="", bg=C["bg"], fg=C["err"], font=F(9, True), anchor="w", padx=18)
        self.folder_warn.pack(fill="x")
        body = tk.Frame(self, bg=C["card"], padx=16, pady=10, highlightthickness=1, highlightbackground=C["line"])
        body.pack(fill="x", padx=14, pady=(12, 0))
        body.columnconfigure(2, weight=1)
        self.rows = {}
        for i, (key, label, menu, required, use) in enumerate(cv.DAILY_SLOTS):
            icon = tk.Label(body, text="", bg=C["card"], font=F(13, True), width=2)
            icon.grid(row=i * 2, column=0, rowspan=2, sticky="n", pady=(6, 0))
            nm = f"{names.get(key, '')}{self.mmdd}" if key != "prev_report" else f"일보{(app.bdate - dt.timedelta(days=1)):%m%d}"
            tk.Label(body, text=f"{label}", bg=C["card"], fg=C["text"], font=F(10, True), anchor="w").grid(
                row=i * 2, column=1, sticky="w", pady=(6, 0))
            file_lbl = tk.Label(body, text="", bg=C["card"], font=F(9), anchor="w")
            file_lbl.grid(row=i * 2, column=2, sticky="w", padx=(10, 0), pady=(6, 0))
            tk.Label(body, text=f"파일 이름 '{nm}'  ·  ERP {menu}  →  {use}", bg=C["card"], fg=C["hint"], font=F(8),
                     anchor="w").grid(row=i * 2 + 1, column=1, columnspan=2, sticky="w")
            FlatButton(body, "선택", lambda k=key, l=label: self.pick(k, l), kind="soft", size=9, padx=12, pady=4).grid(
                row=i * 2, column=3, rowspan=2, padx=(10, 0), pady=(6, 0))
            self.rows[key] = (icon, file_lbl, required)
        i = len(cv.DAILY_SLOTS)
        self.folio_icon = tk.Label(body, text="", bg=C["card"], font=F(13, True), width=2)
        self.folio_icon.grid(row=i * 2, column=0, rowspan=2, sticky="n", pady=(6, 0))
        tk.Label(body, text="객실별 거래내역 (선택)", bg=C["card"], fg=C["text"], font=F(10, True)).grid(
            row=i * 2, column=1, sticky="w", pady=(6, 0))
        self.folio_lbl = tk.Label(body, text="", bg=C["card"], font=F(9), anchor="w")
        self.folio_lbl.grid(row=i * 2, column=2, sticky="w", padx=(10, 0), pady=(6, 0))
        tk.Label(body, text="중간 정산한 장기 투숙 객실만 · 이름 자유 (폴더에 넣으면 내용으로 자동 인식)  →  ① 누계(잔액)",
                 bg=C["card"], fg=C["hint"], font=F(8), anchor="w").grid(row=i * 2 + 1, column=1, columnspan=2, sticky="w")
        FlatButton(body, "선택", self.pick_folios, kind="soft", size=9, padx=12, pady=4).grid(
            row=i * 2, column=3, rowspan=2, padx=(10, 0), pady=(6, 0))
        self.out_lbl = tk.Label(self, text="", bg=C["bg"], fg=C["text"], font=F(9), justify="left", anchor="w", padx=18,
                                wraplength=int(760 * UI_SCALE))
        self.out_lbl.pack(fill="x", pady=(10, 0))
        bar = tk.Frame(self, bg=C["bg"], padx=14, pady=12)
        bar.pack(fill="x")
        FlatButton(bar, "ERP원본 폴더 열기", app.open_erp_folder, kind="soft", size=10, pady=8).pack(side="left")
        self.go = FlatButton(bar, "변환 시작  ▶", self.start, size=11, pady=8, padx=24)
        self.go.pack(side="right")
        FlatButton(bar, "닫기", self.destroy, kind="ghost", size=10, pady=8).pack(side="right", padx=(0, 8))
        self.refresh()
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()

    def load_all(self):
        """폴더에서 영업일자 파일을 전부 다시 불러온다 (직접 고른 파일은 지우고 폴더 것으로)"""
        self.app.daily_overrides.clear()
        self.refresh()
        found = sum(1 for k, *_ in cv.DAILY_SLOTS if self.plan.get(k)) + len(self.plan["folios"])
        self.app._log(f"전체 불러오기: {self.folder} 에서 {self.mmdd} 파일 {found}개", "ok" if found else "warn")

    def change_folder(self):
        folder = filedialog.askdirectory(parent=self, title="ERP 원본 파일이 있는 폴더를 고르세요 (예: 공유폴더의 erp 원본)",
                                         initialdir=self.folder if os.path.isdir(self.folder) else None)
        if not folder:
            return
        self.folder = os.path.normpath(folder)
        try:
            cv.save_user_setting("erp_folder", self.folder)
        except cv.ConversionError as e:
            messagebox.showerror("ERP 원본 폴더", str(e), parent=self)
        self.load_all()

    def pick(self, key, label):
        path = filedialog.askopenfilename(parent=self, title=f"[{label}] 파일 선택", filetypes=FILETYPES)
        if path:
            self.app.daily_overrides[key] = path
            self.refresh()

    def pick_folios(self):
        paths = filedialog.askopenfilenames(parent=self, title="객실별 거래내역 (Ctrl+클릭으로 여러 개)", filetypes=FILETYPES)
        if paths:
            self.app.daily_overrides["folios"] = list(paths)
            self.refresh()

    def refresh(self):
        plan = cv.daily_plan(self.folder, f"{self.app.bdate:%Y-%m-%d}", self.app.cfg, self.app.daily_overrides)
        self.plan = plan
        ok = os.path.isdir(self.folder)
        self.folder_lbl.configure(text=self.folder, fg=C["text"] if ok else C["err"])
        self.folder_warn.configure(text="" if ok else "✖ 이 폴더를 열 수 없습니다 — 공유폴더 PC가 켜져 있는지·네트워크 연결을 확인하거나 [폴더 바꾸기]로 고르세요")
        for key, (icon, lbl, required) in self.rows.items():
            path = plan.get(key)
            if path:
                mine = self.app.daily_overrides.get(key) == path
                icon.configure(text="✔", fg=C["ok"])
                lbl.configure(text=os.path.basename(path) + ("  (직접 선택)" if mine else ""), fg=C["ok"])
            else:
                icon.configure(text="✖" if required else "–", fg=C["err"] if required else C["sub"])
                lbl.configure(text="없음" + (" — 폴더에 저장하거나 [선택]" if required else " (없어도 됨)"),
                              fg=C["err"] if required else C["sub"])
        n = len(plan["folios"])
        self.folio_icon.configure(text="✔" if n else "–", fg=C["ok"] if n else C["sub"])
        self.folio_lbl.configure(text=f"{n}개" if n else "없음 (없어도 됨)", fg=C["ok"] if n else C["sub"])
        has = lambda k: bool(plan.get(k))
        lines = [
            ("① 수입일보", has("income"), "어제 일보 이어받기" if has("prev_report") else "어제 일보 없음 → 도착일부터"),
            ("② 카드 승인내역", has("card"), "객실 자동 채우기: " + " + ".join(
                x for x, ok in (("객실조회", has("rooms")), ("수입일보 결제", has("income"))) if ok) if has("rooms") or has("income")
             else "객실 칸은 직접 입력"),
            ("③ 대외후불 리스트", has("credit"), ""),
            ("④·⑤ 조식명단 · 조식 인원수", has("inhouse"), "IN HOUSE + 객실조회" if has("rooms") else "IN HOUSE만 (도착 예정 빠짐)"),
        ]
        self.out_lbl.configure(text="만들어질 서류\n" + "\n".join(
            f"   {'✔' if ok else '✖  (파일 없음)'} {name}" + (f"  —  {note}" if ok and note else "") for name, ok, note in lines))
        self.go.set_enabled(any(ok for _, ok, _ in lines))

    def start(self):
        if not self.go.enabled:
            return
        self.destroy()
        self.app.start_daily(self.folder, self.plan)


class PrintDialog(tk.Toplevel):
    """마감서류 인쇄 (사용자 요청 2026-11-03): 결과 폴더에서 그날 서류를 찾아 순서의 역순으로 인쇄 →
    인쇄물을 쌓으면 대외 → 카드 → 수입 → 일일업무일지 순서가 된다."""

    def __init__(self, app):
        super().__init__(app, bg=C["bg"])
        self.app = app
        self.folder = cv.print_folder(app.cfg)
        self.pcfg = app.cfg.get("print") or {}
        self.title("마감서류 인쇄")
        self.transient(app)
        self.resizable(True, False)
        head = tk.Frame(self, bg=C["head"], padx=18, pady=12)
        head.pack(fill="x")
        tk.Label(head, text=f"영업일자 {app.bdate:%Y-%m-%d} ({WEEKDAYS[app.bdate.weekday()]}) 마감서류 인쇄", bg=C["head"],
                 fg="#FFFFFF", font=F(13, True)).pack(anchor="w")
        tk.Label(head, text="파일 이름은 제목 앞 두 글자 + 날짜 (대외1103·카드1103·수입1103·일일1103). "
                            "아래 순서의 거꾸로 인쇄해서, 다 나온 종이를 그대로 쌓으면 이 순서가 됩니다.",
                 bg=C["head"], fg=C["head_sub"], font=F(9), justify="left",
                 wraplength=int(720 * UI_SCALE)).pack(anchor="w", pady=(2, 0))
        fbar = tk.Frame(self, bg=C["bg"], padx=18)
        fbar.pack(fill="x", pady=(10, 0))
        tk.Label(fbar, text="마감서류 폴더", bg=C["bg"], fg=C["text"], font=F(10, True)).pack(side="left")
        self.folder_lbl = tk.Label(fbar, text="", bg=C["bg"], font=F(9), anchor="w")
        self.folder_lbl.pack(side="left", padx=(10, 0), fill="x", expand=True)
        FlatButton(fbar, "폴더 바꾸기", self.change_folder, kind="soft", size=9, padx=12, pady=4).pack(side="right")
        self.folder_warn = tk.Label(self, text="", bg=C["bg"], fg=C["err"], font=F(9, True), anchor="w", padx=18)
        self.folder_warn.pack(fill="x")
        self.body = tk.Frame(self, bg=C["card"], padx=16, pady=10, highlightthickness=1, highlightbackground=C["line"])
        self.body.pack(fill="x", padx=14, pady=(6, 0))
        self.info = tk.Label(self, text="", bg=C["bg"], fg=C["text"], font=F(9), anchor="w", justify="left", padx=18)
        self.info.pack(fill="x", pady=(8, 0))
        bar = tk.Frame(self, bg=C["bg"], padx=14, pady=12)
        bar.pack(fill="x")
        FlatButton(bar, "폴더 열기", self.open_folder, kind="soft", size=10, pady=8).pack(side="left")
        FlatButton(bar, "다시 찾기", self.refresh, kind="soft", size=10, pady=8).pack(side="left", padx=(8, 0))
        self.go = FlatButton(bar, "인쇄 시작  🖨", self.start, size=11, pady=8, padx=24)
        self.go.pack(side="right")
        FlatButton(bar, "닫기", self.destroy, kind="ghost", size=10, pady=8).pack(side="right", padx=(0, 8))
        self.refresh()
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()

    def _short(self, i):
        order = self.pcfg.get("order") or []
        item = order[i] if i < len(order) else {}
        return item.get("prefix", "") if isinstance(item, dict) else ""

    def open_folder(self):
        try:
            open_path(self.folder)
        except OSError as e:
            messagebox.showwarning("마감서류 폴더", f"폴더를 열 수 없습니다.\n{self.folder}\n({e})", parent=self)

    def change_folder(self):
        folder = filedialog.askdirectory(parent=self, title="인쇄할 마감서류가 있는 폴더를 고르세요 (예: 공유폴더의 마감파일)",
                                         initialdir=self.folder if os.path.isdir(self.folder) else None)
        if not folder:
            return
        self.folder = os.path.normpath(folder)
        try:
            cv.save_user_setting("print_folder", self.folder)
        except cv.ConversionError as e:
            messagebox.showerror("마감서류 폴더", str(e), parent=self)
        self.refresh()

    def refresh(self):
        ok = os.path.isdir(self.folder)
        self.folder_lbl.configure(text=self.folder, fg=C["text"] if ok else C["err"])
        self.folder_warn.configure(text="" if ok else "✖ 이 폴더를 열 수 없습니다 — 공유폴더 PC·네트워크를 확인하거나 [폴더 바꾸기]로 고르세요")
        for w in self.body.winfo_children():
            w.destroy()
        self.plan = cv.print_plan(self.folder, f"{self.app.bdate:%Y-%m-%d}", self.app.cfg)
        for i, (label, path) in enumerate(self.plan):
            tk.Label(self.body, text="✔" if path else "✖", bg=C["card"], fg=C["ok"] if path else C["err"],
                     font=F(13, True), width=2).grid(row=i, column=0, pady=4)
            tk.Label(self.body, text=f"{i + 1}. {label}", bg=C["card"], fg=C["text"], font=F(11, True), anchor="w",
                     width=14).grid(row=i, column=1, sticky="w")
            tk.Label(self.body, text=os.path.basename(path) if path else
                     f"없음 — '{self._short(i)}{self.app.bdate:%m%d}' 파일이 없어 빼고 인쇄", bg=C["card"],
                     fg=C["ok"] if path else C["err"], font=F(9), anchor="w").grid(row=i, column=2, sticky="w", padx=(10, 0))
        todo = [(l, p) for l, p in self.plan if p]
        order = list(reversed(todo)) if self.pcfg.get("reverse", True) else todo
        self.info.configure(text=("인쇄 순서: " + " → ".join(l for l, _ in order)) if order else
                            "인쇄할 서류가 없습니다. 영업일자와 폴더를 확인해 주세요.")
        self.go.set_enabled(bool(order))
        self.order = order

    def start(self):
        if not self.go.enabled:
            return
        order, delay = self.order, int(float(self.pcfg.get("delay_seconds", 12)) * 1000)
        self.destroy()
        app = self.app
        app._log(f"마감서류 인쇄 ({len(order)}개): " + " → ".join(l for l, _ in order), "bold")

        def step(i):
            if i >= len(order):
                app._log("   인쇄 명령을 모두 보냈습니다. 프린터에서 다 나왔는지 확인해 주세요.", "ok")
                return
            label, path = order[i]
            try:
                cv.print_file(path)
                app._log(f"   {i + 1}/{len(order)} {label}: {os.path.basename(path)}", "sub")
            except cv.ConversionError as e:
                app._log(f"   {label}: {e}", "err")
            app.after(delay, lambda: step(i + 1))  # 앞 서류가 프린터로 넘어간 뒤 다음 서류 (순서가 뒤섞이지 않게)

        step(0)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"거성호텔 나이트 마감 도우미 {VERSION}")
        self.configure(bg=C["bg"])
        # Windows 화면 배율(125%·150%)에 맞춰 창 크기도 키운다
        k = max(1.0, self.winfo_fpixels("1i") / 96)
        # 작은 화면(노트북 768px 등)에서도 아래가 잘리지 않게 화면 높이에 맞춘다
        w = min(int(1080 * UI_SCALE * k), self.winfo_screenwidth() - int(40 * k))
        h = min(int(760 * UI_SCALE * k), self.winfo_screenheight() - int(80 * k))
        self.geometry(f"{w}x{h}")
        self.minsize(min(w, int(900 * k)), min(h, int(640 * k)))
        self.last_dir = None
        self.busy = False
        self.jobs = queue.Queue()
        try:
            self.cfg = cv.load_config()
        except cv.ConversionError as e:
            messagebox.showerror("설정 오류", str(e))
            self.destroy()
            return
        self.bdate = dt.datetime.strptime(cv.default_business_date(self.cfg), "%Y-%m-%d").date()
        self._build()
        self._refresh_date()
        self._log("프로그램을 시작했습니다. 영업일자를 확인한 뒤 서류를 선택하세요.", "sub")
        self.last_dir = cv.default_out_dir(None, self.cfg)
        self._log(f"결과 저장 위치: {self.last_dir}  ('저장 위치 변경'으로 바꿀 수 있습니다)", "sub")
        names = self.cfg.get("daily", {}).get("names", {})
        self._log(f"ERP 원본 폴더: {cv.erp_folder(self.cfg)}  — 파일 이름은 "
                  + ", ".join(f"{v}{self.bdate:%m%d}" for k, v in names.items() if k not in ("reserve", "prev_report"))
                  + " 처럼 저장 후 [오늘 마감 한 번에]", "sub")
        self.after(100, self._poll)

    # ------------------------------------------------------------ 화면 구성
    def _build(self):
        # 화면 높이가 낮으면(노트북 768px 등) 간단히 보기: 카드의 ERP 메뉴 줄·여백을 줄여 작업 기록이 보이게
        compact = self.winfo_screenheight() < 820 * UI_SCALE * max(1.0, self.winfo_fpixels("1i") / 96)
        head = tk.Frame(self, bg=C["head"], padx=24, pady=8 if compact else 16)
        head.pack(fill="x")
        left = tk.Frame(head, bg=C["head"])
        left.pack(side="left")
        tk.Label(left, text="거성호텔 나이트 마감 도우미", bg=C["head"], fg="#FFFFFF",
                 font=F(17, True)).pack(anchor="w")
        if not compact:
            tk.Label(left, text="오아시스 ERP 엑셀 → 호텔 마감 양식으로 자동 변환", bg=C["head"],
                     fg=C["head_sub"], font=F(9)).pack(anchor="w", pady=(2, 0))
        right = tk.Frame(head, bg=C["head"])
        right.pack(side="right", anchor="center")
        tk.Label(right, text=f"글자 크기 {round(UI_SCALE * 100)}%", bg=C["head"], fg=C["head_sub"], font=F(9)).pack(side="left")
        FlatButton(right, "가−", lambda: self.change_scale(-0.1), kind="head", size=10, padx=10, pady=4).pack(
            side="left", padx=(8, 0))
        FlatButton(right, "가+", lambda: self.change_scale(0.1), kind="head", size=12, padx=10, pady=2).pack(
            side="left", padx=(4, 0))
        tk.Label(right, text=VERSION, bg=C["head"], fg=C["head_sub"], font=F(9)).pack(side="left", padx=(14, 0))

        # 영업일자
        bar = tk.Frame(self, bg=C["card"], padx=24, pady=4 if compact else 12, highlightthickness=1,
                       highlightbackground=C["line"])
        bar.pack(fill="x", padx=20, pady=(8 if compact else 16, 0))
        tk.Label(bar, text="영업일자", bg=C["card"], fg=C["text"], font=F(11, True)).pack(side="left")
        FlatButton(bar, "◀", lambda: self._shift(-1), kind="ghost", size=11, padx=8, pady=2).pack(
            side="left", padx=(14, 0))
        self.date_lbl = tk.Label(bar, text="", bg=C["card"], fg=C["accent"], font=F(15, True), width=15)
        self.date_lbl.pack(side="left")
        FlatButton(bar, "▶", lambda: self._shift(1), kind="ghost", size=11, padx=8, pady=2).pack(side="left")
        FlatButton(bar, "직접 입력", self._ask_date, kind="soft", size=9, padx=10, pady=4).pack(
            side="left", padx=(12, 0))
        tk.Label(bar, text="정오 전에는 전날이 자동으로 선택됩니다", bg=C["card"], fg=C["sub"],
                 font=F(9)).pack(side="right")

        # 서류 카드 2×2
        grid = tk.Frame(self, bg=C["bg"])
        grid.pack(fill="x", padx=20, pady=(8 if compact else 14, 0))
        grid.columnconfigure((0, 1, 2), weight=1, uniform="col")
        self.cards = {}
        # 글씨를 키우면서 한 줄에 3개 (①②③ / ④ + ⑤ 넓게) — 화면 높이를 덜 차지하도록
        for n, (kind, title, desc, menu) in enumerate(DOCS):
            card = DocCard(grid, n + 1, kind, title, desc, menu, self.run_single, open_path, compact)
            card.grid(row=n // 3, column=n % 3, sticky="nsew", padx=(0 if n % 3 == 0 else 6, 0 if n % 3 == 2 else 6),
                      pady=(0, 12))
            self.cards[kind] = card
        kind, title, desc, menu = BF_COUNT
        card = DocCard(grid, len(DOCS) + 1, kind, title, desc, menu, lambda _k: self.run_bf_count(), open_path, compact)
        card.grid(row=1, column=1, columnspan=2, sticky="nsew", padx=(6, 0), pady=(0, 12))
        card.button.configure(text="파일 2개 선택")
        self.cards[kind] = card

        # 일괄 변환
        act = tk.Frame(self, bg=C["bg"])
        act.pack(fill="x", padx=20)
        # 오늘 마감 한 번에: 바탕화면 ERP원본 폴더의 '수입1029'처럼 이름 붙인 파일을 찾아 모두 변환 (2026-10-29)
        self.daily_btn = FlatButton(act, "오늘 마감 한 번에  ▶", self.run_daily, size=11, pady=11)
        self.daily_btn.pack(side="left", fill="x", expand=True)
        FlatButton(act, "마감서류 인쇄 🖨", self.run_print, kind="soft", size=10, pady=11).pack(side="left", padx=(10, 0))
        self.all_btn = FlatButton(act, "파일 골라서 변환", self.run_all, kind="soft", size=10, pady=11)
        self.all_btn.pack(side="left", padx=(10, 0))
        self.folder_btn = FlatButton(act, "결과 폴더 열기", self.open_folder, kind="soft", size=10, pady=11)
        self.folder_btn.pack(side="left", padx=(10, 0))
        FlatButton(act, "저장 위치 변경", self.choose_out_dir, kind="soft", size=10, pady=11).pack(
            side="left", padx=(10, 0))

        # 상태 표시줄 (맨 아래 고정)
        foot = tk.Frame(self, bg=C["bg"])
        foot.pack(side="bottom", fill="x")
        tk.Label(foot, text=COPYRIGHT, bg=C["bg"], fg="#A3ACBA", font=("맑은 고딕", 8), padx=14).pack(side="right")
        self.status = tk.Label(foot, text="대기 중", bg=C["bg"], fg=C["sub"], font=F(9), anchor="w", padx=22,
                               pady=3 if compact else 8)
        self.status.pack(side="left", fill="x", expand=True)

        # 작업 기록
        logbox = tk.Frame(self, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        logbox.pack(fill="both", expand=True, padx=20, pady=(8 if compact else 14, 0))
        lt = tk.Frame(logbox, bg=C["card"], padx=14, pady=8)
        lt.pack(fill="x")
        tk.Label(lt, text="작업 기록", bg=C["card"], fg=C["text"], font=F(10, True)).pack(side="left")
        clear = tk.Label(lt, text="지우기", bg=C["card"], fg=C["sub"], font=F(9), cursor="hand2")
        clear.pack(side="right")
        clear.bind("<Button-1>", lambda e: self._clear_log())
        tk.Frame(logbox, bg=C["line"], height=1).pack(fill="x")
        inner = tk.Frame(logbox, bg=C["card"])
        inner.pack(fill="both", expand=True)
        self.log = tk.Text(inner, height=7, state="disabled", wrap="word", bg=C["card"], fg=C["text"],
                           relief="flat", font=F(9), padx=14, pady=8, spacing1=2, spacing3=2,
                           highlightthickness=0, borderwidth=0)
        sb = tk.Scrollbar(inner, command=self.log.yview, relief="flat", borderwidth=0)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for tag in ("ok", "warn", "err", "sub"):
            self.log.tag_configure(tag, foreground=C[tag])
        self.log.tag_configure("time", foreground=C["hint"], font=(MONO, S(8)))
        self.log.tag_configure("bold", font=F(9, True))

    def change_scale(self, delta):
        """글자 크기 조절 (화면 위 [가−]/[가+]) — 바로 적용하고 다음 실행 때도 기억"""
        global UI_SCALE
        new = round(max(1.0, min(1.9, UI_SCALE + delta)), 2)
        if self.busy or new == UI_SCALE:
            return
        UI_SCALE = new
        try:
            cv.save_user_setting("ui_scale", new)
        except cv.ConversionError:
            pass
        text = self.log.get("1.0", "end-1c")
        for w in self.winfo_children():
            w.destroy()
        self._build()
        self._refresh_date()
        self.log.configure(state="normal")
        self.log.insert("end", text + ("\n" if text else ""))
        self.log.see("end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------ 영업일자
    def _refresh_date(self):
        d = self.bdate
        self.date_lbl.configure(text=f"{d:%Y-%m-%d} ({WEEKDAYS[d.weekday()]})")

    def _shift(self, days):
        if not self.busy:
            self.bdate += dt.timedelta(days=days)
            self._refresh_date()

    def _ask_date(self):
        if self.busy:
            return
        win = tk.Toplevel(self, bg=C["card"], padx=20, pady=16)
        win.title("영업일자 입력")
        win.resizable(False, False)
        win.transient(self)
        tk.Label(win, text="영업일자 (예: 2026-10-23)", bg=C["card"], fg=C["text"], font=F(10)).pack(anchor="w")
        var = tk.StringVar(value=f"{self.bdate:%Y-%m-%d}")
        ent = tk.Entry(win, textvariable=var, font=F(12), width=14, relief="solid", borderwidth=1)
        ent.pack(pady=8, fill="x")
        ent.select_range(0, "end")
        ent.focus_set()
        msg = tk.Label(win, text="", bg=C["card"], fg=C["err"], font=F(9))
        msg.pack(anchor="w")

        def ok(_e=None):
            try:
                self.bdate = dt.datetime.strptime(var.get().strip(), "%Y-%m-%d").date()
            except ValueError:
                msg.configure(text="YYYY-MM-DD 형식으로 입력해 주세요.")
                return
            self._refresh_date()
            win.destroy()

        FlatButton(win, "확인", ok, size=10, pady=6).pack(fill="x", pady=(6, 0))
        ent.bind("<Return>", ok)
        win.bind("<Escape>", lambda e: win.destroy())
        win.grab_set()

    # ------------------------------------------------------------ 작업 기록
    def _log(self, msg, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", dt.datetime.now().strftime("%H:%M  "), "time")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------ 변환 (작업 스레드)
    def _set_busy(self, on, text=""):
        self.busy = on
        for card in self.cards.values():
            card.button.set_enabled(not on)
        self.all_btn.set_enabled(not on)
        self.daily_btn.set_enabled(not on)
        self.configure(cursor="watch" if on else "")
        self.status.configure(text=text or "대기 중", fg=C["accent"] if on else C["sub"])

    def _start(self, jobs, room_source=None, unique=None):
        """jobs: [(kind 또는 None, 경로)] 또는 [(kind, 경로, 짝 파일)] — 화면이 멈추지 않도록 별도 스레드에서 변환.
        짝 파일: 카드 → 객실 자동 채우기용 객실조회, 조식 인원수(bf_count) → 객실 예약 리스트"""
        bdate = f"{self.bdate:%Y-%m-%d}"
        jobs = [j if len(j) == 3 else (j[0], j[1], room_source) for j in jobs]
        unique = len(jobs) > 1 if unique is None else unique
        self._set_busy(True, f"변환 중… (영업일자 {bdate})")
        for kind, _, _ in jobs:
            if kind:
                self.cards[kind].set_state("run", "변환 중…")

        def work():
            for kind, path, extra in jobs:
                if kind is None:
                    kind = cv.detect_type(path, self.cfg)
                    if kind is None:
                        self.jobs.put(("unknown", None, path, None))
                        continue
                    self.jobs.put(("start", kind, path, None))
                try:
                    out = cv.default_out_dir(path, self.cfg)
                    if kind == "bf_count":
                        res = cv.breakfast_count(path, extra, out, bdate, self.cfg)
                    elif isinstance(extra, dict):  # 오늘 마감: 파일끼리 연결 (객실조회·수입일보 결제·어제 일보·거래내역)
                        res = cv.convert_file(kind, path, out, bdate, self.cfg, unique=unique, **extra)
                    else:
                        res = cv.convert_file(kind, path, out, bdate, self.cfg, unique=unique,
                                              room_source=extra if kind == "card" else None,
                                              folios=extra if kind == "income" else None)
                    self.jobs.put(("ok", kind, path, res))
                except cv.ConversionError as e:
                    self.jobs.put(("err", kind, path, str(e)))
                except Exception as e:  # 예상치 못한 오류도 프로그램이 꺼지지 않도록
                    log = cv.write_error_log(f"변환 오류 [{kind}] {path}")
                    self.jobs.put(("err", kind, path, f"예상치 못한 오류: {type(e).__name__}: {e}"
                                   + (f"\n오류 내용이 저장된 파일: {log}" if log else "")))
            self.jobs.put(("done", None, None, None))

        self._summary = {"ok": 0, "fail": 0}
        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            while True:
                self._handle(*self.jobs.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _handle(self, ev, kind, path, data):
        name = os.path.basename(path) if path else ""
        if ev == "start":
            self.cards[kind].set_state("run", "변환 중…")
        elif ev == "unknown":
            self._summary["fail"] += 1
            self._log(f"{name} — 어떤 서류인지 알 수 없습니다. 해당 서류 카드의 '파일 선택'으로 변환해 주세요.", "err")
        elif ev == "ok":
            self._summary["ok"] += 1
            title = self._title(kind)
            self.last_dir = os.path.dirname(data["output"])
            self.folder_btn.set_enabled(True)
            if kind == "bf_count":
                self._log(f"[{title}] {data['bf_date'][5:].replace('-', '/')} 아침 조식 {data['total']}명 "
                          f"(투숙 중 {data['inhouse']}명 + 도착 예정 {data['arrivals']}명) → "
                          f"{os.path.basename(data['output'])}", "ok")
                for w in data["warnings"][1:]:
                    self._log("   확인: " + w, "warn")
                # 결과 열기 = 호텔 양식 조식명단 (없으면 집계표)
                self.cards[kind].set_state("ok" if data.get("sanity_ok", True) else "warn",
                                           f"조식 {data['total']}명" + ("" if data.get("sanity_ok", True) else " · 확인 필요"),
                                           data.get("list") or data["output"])
                if data.get("list"):
                    self._log(f"   조식명단(호텔 양식) → {os.path.basename(data['list'])}  ('결과 열기')", "ok")
                show = messagebox.showinfo if data.get("sanity_ok", True) else messagebox.showwarning
                self.after(50, lambda d=data, show=show: show(
                    "조식 인원수", f"{d['bf_date'][5:].replace('-', '/')} 아침 조식  {d['total']}명\n\n"
                    f"투숙 중(체크인 완료)  {d['inhouse']}명\n오늘 도착 예정(체크인 전)  {d['arrivals']}명\n\n"
                    + ("" if d.get("sanity_ok", True) else "⚠ 한 객실 인원이 너무 많은 곳이 있습니다. 작업 기록을 꼭 확인하세요.\n\n")
                    + "카드의 '결과 열기'로 호텔 양식 조식명단을 열 수 있습니다."))
                return
            self._log(f"[{title}] {data['count']}건 변환 완료 → {os.path.basename(data['output'])}", "ok")
            for w in data["warnings"]:
                self._log("   확인: " + w, "warn")
            if data["warnings"]:
                self.cards[kind].set_state("warn", f"완료 {data['count']}건 · 확인 {len(data['warnings'])}개",
                                           data["output"])
            else:
                self.cards[kind].set_state("ok", f"완료 {data['count']}건", data["output"])
        elif ev == "err":
            self._summary["fail"] += 1
            title = self._title(kind)
            self._log(f"[{title}] {name} 변환 실패", "err")
            for line in data.splitlines():
                self._log("   " + line, "err")
            self.cards[kind].set_state("err", "변환 실패 · 아래 기록 확인")
            self.bell()
        elif ev == "done":
            ok, fail = self._summary["ok"], self._summary["fail"]
            text = f"완료 — 성공 {ok}건" + (f", 실패 {fail}건" if fail else "") + \
                   ("  ·  '결과 폴더 열기'로 결과를 확인하세요" if ok else "")
            self._set_busy(False)
            self.status.configure(text=text, fg=C["err"] if fail and not ok else C["sub"])

    def _title(self, kind):
        if kind == "bf_count":
            return self.cfg["breakfast_count"]["title"]
        return self.cfg["converters"][kind]["title"]

    # ------------------------------------------------------------ 버튼 동작
    def run_bf_count(self):
        """조식 인원수: IN HOUSE LIST → 객실 예약 리스트 순서로 파일 두 개를 고른다."""
        if self.busy:
            return
        ih = filedialog.askopenfilename(title="[조식 인원수] 1/2  IN HOUSE LIST 파일 선택 (이미 체크인한 손님)",
                                        filetypes=FILETYPES)
        if not ih:
            return
        rv = filedialog.askopenfilename(title="[조식 인원수] 2/2  객실 예약 리스트 파일 선택 (오늘 도착 예정 손님)",
                                        filetypes=FILETYPES)
        if not rv:
            return
        self._log(f"[조식 인원수] {os.path.basename(ih)} + {os.path.basename(rv)}", "bold")
        self._start([("bf_count", ih, rv)], unique=False)

    def run_single(self, kind):
        if self.busy:
            return
        title = self.cfg["converters"][kind]["title"]
        path = filedialog.askopenfilename(title=f"[{title}] ERP에서 받은 엑셀 파일 선택", filetypes=FILETYPES)
        if not path:
            return
        room_source = None
        if kind == "income":
            # 장기 투숙 객실의 누계(현재 잔액)용 객실별 거래내역 (선택, 여러 개)
            room_source = list(filedialog.askopenfilenames(
                title="(선택) 장기 투숙 객실별 거래내역 파일들 — Ctrl+클릭으로 여러 개. 없으면 [취소]",
                filetypes=FILETYPES)) or None
        if kind == "card" and self.cfg["converters"]["card"].get("room_source"):
            # 객실 칸 자동 채우기 (선택): 객실 예약 리스트(상태 CHECK IN)의 Acct# → 객실#
            room_source = filedialog.askopenfilename(
                title="(선택) 객실 자동 채우기 — 객실관리 → 조회 및 출력 → 객실조회 파일. 없으면 [취소]", filetypes=FILETYPES) or None
        extra = (f" + 거래내역 {len(room_source)}개" if isinstance(room_source, list)
                 else f" + 객실 자동 채우기: {os.path.basename(room_source)}" if room_source else "")
        self._log(f"[{title}] {os.path.basename(path)}" + extra, "bold")
        self._start([(kind, path)], room_source)

    def run_daily(self):
        """오늘 마감 한 번에: ERP 파일 확인 창 (사용자 요청 2026-11-02 — 어떤 파일이 필요한지 한눈에, 없으면 직접 선택)"""
        if self.busy:
            return
        folder = cv.erp_folder(self.cfg)
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            pass
        if getattr(self, "_daily_date", None) != self.bdate:
            self.daily_overrides, self._daily_date = {}, self.bdate  # 영업일자가 바뀌면 직접 고른 파일은 초기화
        DailyDialog(self, folder)

    def start_daily(self, folder, plan=None):
        bdate = f"{self.bdate:%Y-%m-%d}"
        jobs, notes, found = cv.daily_jobs(folder, bdate, self.cfg, self.daily_overrides)
        self._log(f"오늘 마감 한 번에 (영업일자 {bdate})", "bold")
        for key, path in found.items():
            self._log(f"   {dict((k, l) for k, l, *_ in cv.DAILY_SLOTS).get(key, key)}: {os.path.basename(path)}", "sub")
        for n in notes:
            self._log("   " + n, "warn")
        if jobs:
            self._start(jobs, unique=False)

    def run_print(self):
        if not self.busy:
            PrintDialog(self)

    def open_erp_folder(self):
        folder = cv.erp_folder(self.cfg)
        try:
            os.makedirs(folder, exist_ok=True)
            open_path(folder)
        except OSError as e:
            messagebox.showwarning("ERP원본 폴더", f"폴더를 열 수 없습니다.\n{folder}\n({e})")

    def run_all(self):
        if self.busy:
            return
        paths = filedialog.askopenfilenames(
            title="마감 원본 파일을 모두 선택하세요 (Ctrl+클릭으로 여러 개)", filetypes=FILETYPES)
        if paths:
            self._log(f"일괄 변환: {len(paths)}개 파일", "bold")
            self._start([(None, p) for p in paths])

    def open_folder(self):
        folder = self.last_dir or cv.default_out_dir(None, self.cfg)
        try:
            os.makedirs(folder, exist_ok=True)
            open_path(folder)
        except OSError as e:
            messagebox.showwarning("결과 폴더", f"폴더를 열 수 없습니다.\n{folder}\n({e})")

    def choose_out_dir(self):
        """예: 프로그램은 바탕화면에서 실행하고, 결과는 공용폴더의 한희 폴더에 모으고 싶을 때"""
        if self.busy:
            return
        folder = filedialog.askdirectory(title="변환 결과를 저장할 폴더를 고르세요 (예: 공용폴더의 한희 폴더)",
                                         initialdir=self.last_dir or None)
        if not folder:
            return
        folder = os.path.normpath(folder)
        if os.path.basename(folder) not in ("마감결과", "마감파일"):  # 공유폴더 '마감파일'은 그대로 (2026-11-03)
            folder = os.path.join(folder, "마감결과")
        try:
            cv.save_user_setting("output_dir", folder)
        except cv.ConversionError as e:
            messagebox.showerror("저장 위치", str(e))
            return
        self.last_dir = folder
        self._log(f"결과 저장 위치를 바꿨습니다: {folder}", "ok")


def main():
    try:
        enable_dpi_awareness()
        App().mainloop()
    except Exception as e:  # 마지막 안전망: 창이 그냥 사라지지 않게 안내하고 오류 내용을 파일로 남긴다
        log = cv.write_error_log("프로그램 실행 오류")
        try:
            messagebox.showerror("오류", f"프로그램 실행 중 오류가 발생했습니다.\n{type(e).__name__}: {e}"
                                 + (f"\n\n오류 내용이 저장된 파일:\n{log}" if log else ""))
        except Exception:
            pass


if __name__ == "__main__":
    main()
