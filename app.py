# -*- coding: utf-8 -*-
"""거성호텔 나이트 근무 마감 자동화 툴 v1.0 (GUI)"""
import os
import re
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import converters as cv

VERSION = "v1.2"
SECTIONS = [
    ("income", "1. 수입일보 자동 변환", [
        "오아시스 ERP에서 추출한 엑셀 파일을 선택하여 호텔 양식으로 변환.",
        "객실, 투숙객명, 거래처, 도착일, 출발일, 매출(NET) 자동 매핑, 객실번호 순 정렬.",
        "'선/후', '숙박수'는 노란 빈칸으로 남겨두어 근무자 최종 확인 유도."]),
    ("card", "2. 카드/현금영수증 승인 내역 변환", [
        "ERP 승인 내역 조회 엑셀을 선택하여 호텔 마감 양식의 표로 정리.",
        "카드번호, 사업자번호, 매입사, 가맹점번호, 구분, 승인번호, 승인금액을 올바른 칸에 자동 정렬 (승인 실패 건 제외).",
        "ERP에는 객실번호 대신 Acct#만 있어 '객실' 칸은 직접 입력 (Acct#는 오른쪽에 참고용 표시)."]),
    ("credit", "3. 대외 후불 리스트 정리", [
        "대외 후불 내역 원본 데이터를 읽어와 최종 보고 양식으로 변환.",
        "객실, 고객명, 도착일, 출발일, 거래처, 금액 추출 및 자동 합계 계산."]),
    ("breakfast", "4. 조식 명단 추출 및 정리", [
        "IN HOUSE LIST에서 비고에 '조식/조포'가 있는 객실만 추출 ('룸온리' 제외).",
        "객실, 고객명, 인원(비고의 'N인' 우선)을 호텔 조식명단 양식(Sign 칸 포함)으로 채우기."]),
]
FILETYPES = [("엑셀/표 파일", "*.xlsx *.xlsm *.xls *.csv"), ("모든 파일", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"거성호텔 나이트 근무 마감 자동화 툴 {VERSION}")
        self.geometry("760x780")
        self.minsize(720, 700)
        self.last_dir = None
        try:
            self.cfg = cv.load_config()
        except cv.ConversionError as e:
            messagebox.showerror("설정 오류", str(e))
            self.destroy()
            return
        self._build()

    # ------------------------------------------------------------ UI
    def _build(self):
        pad = {"padx": 14}
        ttk.Label(self, text=f"거성호텔 나이트 근무 마감 자동화 툴 {VERSION}",
                  font=("맑은 고딕", 15, "bold")).pack(pady=(14, 6))

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Label(top, text="영업일자").pack(side="left")
        self.date_var = tk.StringVar(value=cv.default_business_date(self.cfg))
        ttk.Entry(top, textvariable=self.date_var, width=12).pack(side="left", padx=6)
        ttk.Label(top, text="(YYYY-MM-DD · 정오 이전에는 전날이 자동 입력됩니다)", foreground="#666").pack(side="left")

        for key, title, lines in SECTIONS:
            box = ttk.LabelFrame(self, text=f" {title} ")
            box.pack(fill="x", padx=14, pady=5)
            for ln in lines:
                ttk.Label(box, text="• " + ln, wraplength=560, justify="left").pack(anchor="w", padx=8, pady=1)
            ttk.Button(box, text="파일 선택 후 변환 ▶",
                       command=lambda k=key: self.run_single(k)).pack(anchor="e", padx=8, pady=6)

        ttk.Separator(self).pack(fill="x", padx=14, pady=8)
        style = ttk.Style()
        style.configure("Big.TButton", font=("맑은 고딕", 11, "bold"), padding=8)
        ttk.Button(self, text="▶ (선택) 모든 마감 서류 일괄 변환 (한 번에 처리)",
                   style="Big.TButton", command=self.run_all).pack(fill="x", **pad)

        self.status = tk.StringVar(value="상태: 대기 중... 파일을 선택해주세요.")
        ttk.Label(self, textvariable=self.status, font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 2))

        logf = ttk.Frame(self)
        logf.pack(fill="both", expand=True, padx=14, pady=(0, 4))
        self.log = tk.Text(logf, height=8, state="disabled", wrap="word", background="#fafafa")
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.log.tag_configure("ok", foreground="#1a7f37")
        self.log.tag_configure("warn", foreground="#b35900")
        self.log.tag_configure("err", foreground="#c00000")

        self.open_btn = ttk.Button(self, text="결과 폴더 열기", state="disabled", command=self.open_folder)
        self.open_btn.pack(anchor="e", padx=14, pady=(0, 10))

    # ------------------------------------------------------------ 보조
    def _log(self, msg, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _date(self):
        d = self.date_var.get().strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
            messagebox.showwarning("영업일자 확인", "영업일자를 YYYY-MM-DD 형식으로 입력해 주세요.\n예) 2026-10-23")
            return None
        return d

    def _convert(self, kind, path, bdate, unique=False):
        out_dir = cv.default_out_dir(path)
        name = self.cfg["converters"][kind]["title"]
        self.status.set(f"상태: [{name}] 변환 중...")
        self._log(f"▶ [{name}] {os.path.basename(path)}")
        try:
            res = cv.convert_file(kind, path, out_dir, bdate, self.cfg, unique=unique)
        except cv.ConversionError as e:
            self._log(f"   ✖ {e}", "err")
            self.status.set(f"상태: [{name}] 변환 실패")
            return None
        except Exception as e:  # 예상치 못한 오류도 창이 꺼지지 않도록 안내
            self._log(f"   ✖ 예상치 못한 오류: {type(e).__name__}: {e}", "err")
            self.status.set(f"상태: [{name}] 변환 실패")
            return None
        self.last_dir = out_dir
        self.open_btn.configure(state="normal")
        self._log(f"   ✔ {res['count']}건 변환 완료 → {os.path.basename(res['output'])}", "ok")
        for w in res["warnings"]:
            self._log(f"   ⚠ {w}", "warn")
        return res

    # ------------------------------------------------------------ 동작
    def run_single(self, kind):
        bdate = self._date()
        if not bdate:
            return
        name = self.cfg["converters"][kind]["title"]
        path = filedialog.askopenfilename(title=f"[{name}] 원본 파일 선택", filetypes=FILETYPES)
        if not path:
            return
        res = self._convert(kind, path, bdate)
        if res:
            self.status.set(f"상태: [{name}] 변환 완료 ({res['count']}건) - 결과 폴더에서 확인하세요.")
            if res["warnings"]:
                messagebox.showinfo("확인 필요", "변환은 완료되었으나 확인할 사항이 있습니다.\n\n" + "\n".join(res["warnings"]))
        else:
            messagebox.showerror("변환 실패", "아래 로그의 안내를 확인해 주세요.")

    def run_all(self):
        bdate = self._date()
        if not bdate:
            return
        paths = filedialog.askopenfilenames(
            title="마감 원본 파일을 모두 선택하세요 (Ctrl+클릭으로 여러 개)", filetypes=FILETYPES)
        if not paths:
            return
        ok, fail, unknown = 0, 0, []
        for p in paths:
            kind = cv.detect_type(p, self.cfg)
            if not kind:
                unknown.append(os.path.basename(p))
                self._log(f"▶ {os.path.basename(p)}\n   ✖ 마감 서류 종류를 알 수 없습니다. 개별 버튼으로 변환해 주세요.", "err")
                continue
            if self._convert(kind, p, bdate, unique=True):
                ok += 1
            else:
                fail += 1
        total = len(paths)
        self.status.set(f"상태: 일괄 변환 종료 - 성공 {ok} / 실패 {fail + len(unknown)} (총 {total}개)")
        if ok:
            messagebox.showinfo("일괄 변환 완료", f"성공 {ok}건, 실패 {fail + len(unknown)}건\n결과 폴더에서 확인하세요.")
        else:
            messagebox.showerror("일괄 변환 실패", "변환된 파일이 없습니다. 로그를 확인해 주세요.")

    def open_folder(self):
        if not self.last_dir or not os.path.isdir(self.last_dir):
            return
        if sys.platform.startswith("win"):
            os.startfile(self.last_dir)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", self.last_dir])
        else:
            subprocess.Popen(["xdg-open", self.last_dir])


if __name__ == "__main__":
    App().mainloop()
