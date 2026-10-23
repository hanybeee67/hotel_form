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
WEEKDAYS = "월화수목금토일"
FILETYPES = [("엑셀/표 파일", "*.xlsx *.xlsm *.xls *.csv *.htm *.html"), ("모든 파일", "*.*")]

# 서류 카드: (종류, 제목, 설명, ERP 메뉴 위치)
DOCS = [
    ("income", "수입일보", "객실 순 정렬 · 매출(NET) · 누계 수식\n선/후 · 숙박수는 노란 칸에 직접 입력",
     "영업보고서 → 마감일보 → 수입일보 → 엑셀전환"),
    ("card", "카드 · 현금영수증 승인내역", "승인 '성공' 건만 · 현금영수증 합산 · 계\n객실 칸은 Acct#를 보고 직접 입력",
     "객실관리 → 카드결제 → 승인 내역조회 → 엑셀전환"),
    ("credit", "대외후불 리스트", "객실 순 정렬 · 금액 합계\n원본 합계와 자동 대조",
     "영업보고서 → 입금관련 → 대외 후불 리스트"),
    ("breakfast", "조식명단", "비고의 '조식/조포' 객실만 추출 ('룸온리' 제외)\n인원은 비고의 'N인' 우선",
     "영업보고서 → 객실/예약 → Clerk → IN HOUSE LIST"),
]

# ---------------------------------------------------------------- 디자인
C = {
    "bg": "#F3F5F9", "card": "#FFFFFF", "line": "#E2E6EE", "head": "#1F2A44", "head_sub": "#AEB8CF",
    "text": "#1F2937", "sub": "#6B7280", "accent": "#2F6FEB", "accent_hover": "#2559C4",
    "soft": "#EAF1FE", "soft_hover": "#D7E5FD", "ok": "#15803D", "warn": "#B45309", "err": "#B91C1C",
    "ok_bg": "#E8F6EC", "warn_bg": "#FEF3E2", "err_bg": "#FDECEC", "disabled": "#C7CDD8",
}
BADGE = {"income": "#2F6FEB", "card": "#7C3AED", "credit": "#0E9F6E", "breakfast": "#E08A00"}

if sys.platform.startswith("win"):
    FONT = "맑은 고딕"
    MONO = "Consolas"
else:
    FONT = "WenQuanYi Zen Hei"
    MONO = "DejaVu Sans Mono"


def F(size, bold=False):
    return (FONT, size, "bold") if bold else (FONT, size)


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

    def __init__(self, master, index, kind, title, desc, menu, on_run, on_open):
        super().__init__(master, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        self.kind, self.on_open, self.output = kind, on_open, None
        tk.Frame(self, bg=BADGE[kind], width=5).pack(side="left", fill="y")
        body = tk.Frame(self, bg=C["card"], padx=16, pady=14)
        body.pack(side="left", fill="both", expand=True)

        top = tk.Frame(body, bg=C["card"])
        top.pack(fill="x")
        tk.Label(top, text=str(index), bg=BADGE[kind], fg="#FFFFFF", font=F(10, True),
                 width=2).pack(side="left")
        tk.Label(top, text=title, bg=C["card"], fg=C["text"], font=F(12, True)).pack(side="left", padx=(8, 0))

        tk.Label(body, text=desc, bg=C["card"], fg=C["sub"], font=F(9), justify="left",
                 anchor="w").pack(fill="x", pady=(8, 2))
        tk.Label(body, text="ERP  " + menu, bg=C["card"], fg="#9AA3B2", font=F(8), anchor="w",
                 justify="left", wraplength=330).pack(fill="x")

        bottom = tk.Frame(body, bg=C["card"])
        bottom.pack(fill="x", pady=(12, 0))
        self.status = tk.Label(bottom, text="대기 중", bg=C["bg"], fg=C["sub"], font=F(9), padx=8, pady=3)
        self.status.pack(side="left")
        self.open_link = tk.Label(bottom, text="결과 열기", bg=C["card"], fg=C["accent"],
                                  font=(FONT, 9, "underline"), cursor="hand2")
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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"거성호텔 나이트 마감 도우미 {VERSION}")
        self.configure(bg=C["bg"])
        # Windows 화면 배율(125%·150%)에 맞춰 창 크기도 키운다
        k = max(1.0, self.winfo_fpixels("1i") / 96)
        self.geometry(f"{int(900 * k)}x{int(780 * k)}")
        self.minsize(int(820 * k), int(700 * k))
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
        self.after(100, self._poll)

    # ------------------------------------------------------------ 화면 구성
    def _build(self):
        head = tk.Frame(self, bg=C["head"], padx=24, pady=16)
        head.pack(fill="x")
        left = tk.Frame(head, bg=C["head"])
        left.pack(side="left")
        tk.Label(left, text="거성호텔 나이트 마감 도우미", bg=C["head"], fg="#FFFFFF",
                 font=F(17, True)).pack(anchor="w")
        tk.Label(left, text="오아시스 ERP 엑셀 → 호텔 마감 양식으로 자동 변환", bg=C["head"],
                 fg=C["head_sub"], font=F(9)).pack(anchor="w", pady=(2, 0))
        tk.Label(head, text=VERSION, bg=C["head"], fg=C["head_sub"], font=F(9)).pack(side="right", anchor="n")

        # 영업일자
        bar = tk.Frame(self, bg=C["card"], padx=24, pady=12, highlightthickness=1,
                       highlightbackground=C["line"])
        bar.pack(fill="x", padx=20, pady=(16, 0))
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
        grid.pack(fill="x", padx=20, pady=(14, 0))
        grid.columnconfigure((0, 1), weight=1, uniform="col")
        self.cards = {}
        for n, (kind, title, desc, menu) in enumerate(DOCS):
            card = DocCard(grid, n + 1, kind, title, desc, menu, self.run_single, open_path)
            card.grid(row=n // 2, column=n % 2, sticky="nsew", padx=(0 if n % 2 == 0 else 7, 7 if n % 2 == 0 else 0),
                      pady=(0, 14))
            self.cards[kind] = card

        # 일괄 변환
        act = tk.Frame(self, bg=C["bg"])
        act.pack(fill="x", padx=20)
        self.all_btn = FlatButton(act, "모든 서류 한 번에 변환  (여러 파일 선택)", self.run_all, size=11, pady=11)
        self.all_btn.pack(side="left", fill="x", expand=True)
        self.folder_btn = FlatButton(act, "결과 폴더 열기", self.open_folder, kind="soft", size=10, pady=11)
        self.folder_btn.pack(side="left", padx=(10, 0))
        self.folder_btn.set_enabled(False)

        # 상태 표시줄 (맨 아래 고정)
        self.status = tk.Label(self, text="대기 중", bg=C["bg"], fg=C["sub"], font=F(9), anchor="w", padx=22,
                               pady=8)
        self.status.pack(side="bottom", fill="x")

        # 작업 기록
        logbox = tk.Frame(self, bg=C["card"], highlightthickness=1, highlightbackground=C["line"])
        logbox.pack(fill="both", expand=True, padx=20, pady=(14, 0))
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
        self.log.tag_configure("time", foreground="#9AA3B2", font=(MONO, 8))
        self.log.tag_configure("bold", font=F(9, True))

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
        self.configure(cursor="watch" if on else "")
        self.status.configure(text=text or "대기 중", fg=C["accent"] if on else C["sub"])

    def _start(self, jobs):
        """jobs: [(kind 또는 None, 경로)] — 화면이 멈추지 않도록 별도 스레드에서 변환."""
        bdate = f"{self.bdate:%Y-%m-%d}"
        unique = len(jobs) > 1
        self._set_busy(True, f"변환 중… (영업일자 {bdate})")
        for kind, _ in jobs:
            if kind:
                self.cards[kind].set_state("run", "변환 중…")

        def work():
            for kind, path in jobs:
                if kind is None:
                    kind = cv.detect_type(path, self.cfg)
                    if kind is None:
                        self.jobs.put(("unknown", None, path, None))
                        continue
                    self.jobs.put(("start", kind, path, None))
                try:
                    res = cv.convert_file(kind, path, cv.default_out_dir(path), bdate, self.cfg, unique=unique)
                    self.jobs.put(("ok", kind, path, res))
                except cv.ConversionError as e:
                    self.jobs.put(("err", kind, path, str(e)))
                except Exception as e:  # 예상치 못한 오류도 프로그램이 꺼지지 않도록
                    self.jobs.put(("err", kind, path, f"예상치 못한 오류: {type(e).__name__}: {e}"))
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
            title = self.cfg["converters"][kind]["title"]
            self.last_dir = os.path.dirname(data["output"])
            self.folder_btn.set_enabled(True)
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
            title = self.cfg["converters"][kind]["title"]
            self._log(f"[{title}] {name} 변환 실패", "err")
            for line in data.splitlines():
                self._log("   " + line, "err")
            self.cards[kind].set_state("err", "변환 실패 · 아래 기록 확인")
            self.bell()
        elif ev == "done":
            ok, fail = self._summary["ok"], self._summary["fail"]
            text = f"완료 — 성공 {ok}건" + (f", 실패 {fail}건" if fail else "") + \
                   ("  ·  결과는 원본 폴더의 '마감결과' 폴더에 저장됩니다" if ok else "")
            self._set_busy(False)
            self.status.configure(text=text, fg=C["err"] if fail and not ok else C["sub"])

    # ------------------------------------------------------------ 버튼 동작
    def run_single(self, kind):
        if self.busy:
            return
        title = self.cfg["converters"][kind]["title"]
        path = filedialog.askopenfilename(title=f"[{title}] ERP에서 받은 엑셀 파일 선택", filetypes=FILETYPES)
        if path:
            self._log(f"[{title}] {os.path.basename(path)}", "bold")
            self._start([(kind, path)])

    def run_all(self):
        if self.busy:
            return
        paths = filedialog.askopenfilenames(
            title="마감 원본 파일을 모두 선택하세요 (Ctrl+클릭으로 여러 개)", filetypes=FILETYPES)
        if paths:
            self._log(f"일괄 변환: {len(paths)}개 파일", "bold")
            self._start([(None, p) for p in paths])

    def open_folder(self):
        if self.last_dir and os.path.isdir(self.last_dir):
            open_path(self.last_dir)


def main():
    enable_dpi_awareness()
    try:
        App().mainloop()
    except Exception as e:  # 마지막 안전망: 창이 그냥 사라지지 않게 안내
        try:
            messagebox.showerror("오류", f"프로그램 실행 중 오류가 발생했습니다.\n{type(e).__name__}: {e}")
        except Exception:
            pass


if __name__ == "__main__":
    main()
