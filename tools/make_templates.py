# -*- coding: utf-8 -*-
"""호텔 실제 마감 파일 → 개인정보 없는 빈 양식(templates/*.xlsx) 만들기 (개발자용 도구)

호텔 파일에서 '서식'만 옮겨 새 통합 문서를 만든다.
  제목 행 · 머리글 행 · 데이터 서식 행 1줄 · 계 행 · (조식명단) 외부식사·안내 문구 행
값은 제목/머리글/계/안내 문구만 남기고 투숙객·카드 정보는 하나도 옮기지 않는다.

사용 예 (프로젝트 폴더에서):
  python tools/make_templates.py --income 수입일보.xls --card 카드승인.xls \
      --credit 대외후불.xls --breakfast 조식명단.xlsx
옵션 없이 준 서류만 만든다. 결과는 templates/ 에 저장.
"""
import argparse
import os
import sys
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = {"income": "수입일보.xlsx", "card": "카드승인내역.xlsx", "credit": "대외후불리스트.xlsx",
       "breakfast": "조식명단.xlsx"}
TOTAL = {"계", "합계"}
# 구형 .xls는 xlrd가 인쇄 설정을 읽지 못하므로, 호텔 파일을 엑셀/LibreOffice로 열어 확인한 값을 둔다 (2026-10-23)
PAGE = {
    "income": {"orientation": "landscape", "margins": (0.748, 0.748, 0.394, 0.394), "titles": True},
    "card": {"orientation": "landscape", "margins": (0.75, 0.75, 1.0, 1.0), "titles": False},
    "credit": {"orientation": "landscape", "margins": (0.75, 0.75, 1.0, 1.0), "titles": False},
}
XLS_LINE = {1: "thin", 2: "medium", 3: "dashed", 4: "dotted", 5: "thick", 6: "double", 7: "hair",
            8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot", 12: "mediumDashDotDot",
            13: "slantDashDot"}
XLS_H = {1: "left", 2: "center", 3: "right", 4: "fill", 5: "justify", 6: "centerContinuous"}
XLS_V = {0: "top", 1: "center", 2: "bottom", 3: "justify"}


def _norm(v):
    return "".join(str(v).split()) if v is not None else ""


# ---------------------------------------------------------------- 구형 .xls
def _xls_style(book, sh, r, c):
    xf = book.xf_list[sh.cell_xf_index(r, c)]
    f = book.font_list[xf.font_index]
    rgb = None
    if f.colour_index in book.colour_map and book.colour_map[f.colour_index]:
        rgb = "%02X%02X%02X" % book.colour_map[f.colour_index]
    font = Font(name=f.name, size=f.height / 20, bold=bool(f.bold), italic=bool(f.italic),
                underline="single" if f.underlined else None, color=rgb)
    b = xf.border
    side = lambda s: Side(style=XLS_LINE.get(s)) if s else Side()
    border = Border(left=side(b.left_line_style), right=side(b.right_line_style),
                    top=side(b.top_line_style), bottom=side(b.bottom_line_style))
    a = xf.alignment
    align = Alignment(horizontal=XLS_H.get(a.hor_align), vertical=XLS_V.get(a.vert_align),
                      wrap_text=bool(a.text_wrapped), shrink_to_fit=bool(a.shrink_to_fit))
    fill = PatternFill(fill_type=None)
    bg = xf.background
    if bg.fill_pattern == 1 and book.colour_map.get(bg.pattern_colour_index):
        fill = PatternFill("solid", fgColor="%02X%02X%02X" % book.colour_map[bg.pattern_colour_index])
    fmt = book.format_map[xf.format_key].format_str
    return font, border, align, fill, fmt


def from_xls(kind, path):
    import xlrd
    book = xlrd.open_workbook(path, formatting_info=True)
    sh = next((s for s in book.sheets() if s.sheet_selected), book.sheet_by_index(0))
    hrow = next(r for r in range(min(sh.nrows, 30)) if _norm(sh.cell_value(r, 0)) == "객실")
    # 머리글은 A열부터 테두리 있는 칸이 이어진 곳까지 (양식 옆에 적어 둔 메모 칸은 제외)
    def boxed(c):
        b = book.xf_list[sh.cell_xf_index(hrow, c)].border
        return _norm(sh.cell_value(hrow, c)) and (b.top_line_style or b.bottom_line_style)
    last = 0
    while last + 1 < sh.ncols and boxed(last + 1):
        last += 1
    for (r0, r1, c0, c1) in sh.merged_cells:
        if r0 <= hrow < r1 and c0 <= last < c1:
            last = c1 - 1
    foot = next((r for r in range(hrow + 1, sh.nrows)
                 if any(_norm(sh.cell_value(r, c)) in TOTAL for c in range(last + 1))), None)
    rows = [("title", 0), ("head", hrow), ("data", hrow + 1)] + ([("foot", foot)] if foot else [])

    wb = Workbook()
    ws = wb.active
    ws.title = "양식"
    for out_r, (role, src_r) in enumerate(rows, 1):
        for c in range(last + 1):
            cell = ws.cell(out_r, c + 1)
            cell.font, cell.border, cell.alignment, cell.fill, cell.number_format = _xls_style(book, sh, src_r, c)
            v = sh.cell_value(src_r, c)
            if role in ("title", "head") or (role == "foot" and _norm(v) in TOTAL):
                cell.value = v if v != "" else None
        if src_r in sh.rowinfo_map:
            ws.row_dimensions[out_r].height = sh.rowinfo_map[src_r].height / 20
    for (r0, r1, c0, c1) in sh.merged_cells:
        if r0 in (0, hrow) and c0 <= last:
            out_r = 1 if r0 == 0 else 2
            ws.merge_cells(start_row=out_r, start_column=c0 + 1, end_row=out_r + (r1 - r0) - 1,
                           end_column=min(c1, last + 1))
    for c in range(last + 1):
        if c in sh.colinfo_map:
            ws.column_dimensions[get_column_letter(c + 1)].width = round(sh.colinfo_map[c].width / 256, 2)
    p = PAGE[kind]
    ws.page_setup.orientation = p["orientation"]
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = p["margins"]
    if p["titles"]:
        ws.print_title_rows = "1:2"
    return wb


# ---------------------------------------------------------------- .xlsx (조식명단)
def from_xlsx(kind, path):
    src_wb = load_workbook(path)
    src = src_wb.active
    hrow = next(r for r in range(1, 31) if _norm(src.cell(r, 1).value) == "객실")
    last = 1
    while last < src.max_column and _norm(src.cell(hrow, last + 1).value):
        last += 1
    for mr in src.merged_cells.ranges:
        if mr.min_row <= hrow <= mr.max_row and mr.min_col <= last <= mr.max_col:
            last = mr.max_col
    # 데이터 영역 다음, 글자가 있는 행(외부식사·안내 문구)은 양식의 일부로 남긴다
    data_end = hrow + 1
    trailer = []
    for r in range(hrow + 1, src.max_row + 1):
        first = _norm(src.cell(r, 1).value)
        if first and not first.isdigit() and first not in TOTAL:
            trailer.append(r)
    foot = next((r for r in range(hrow + 1, src.max_row + 1)
                 if any(_norm(src.cell(r, c).value) in TOTAL for c in range(1, last + 1))), None)
    rows = [("title", 1), ("head", hrow), ("data", data_end)] + ([("foot", foot)] if foot else []) + \
           [("trailer", r) for r in trailer]

    wb = Workbook()
    ws = wb.active
    ws.title = "양식"
    rmap = {}
    for out_r, (role, src_r) in enumerate(rows, 1):
        rmap[src_r] = out_r
        for c in range(1, last + 1):
            s, d = src.cell(src_r, c), ws.cell(out_r, c)
            d.font, d.border, d.alignment = copy(s.font), copy(s.border), copy(s.alignment)
            d.fill, d.number_format, d.protection = copy(s.fill), s.number_format, copy(s.protection)
            if role in ("title", "head", "trailer") or (role == "foot" and _norm(s.value) in TOTAL):
                d.value = s.value
        h = src.row_dimensions[src_r].height
        if h:
            ws.row_dimensions[out_r].height = h
    for mr in src.merged_cells.ranges:
        if mr.min_row in rmap and mr.min_col <= last and (mr.min_row <= hrow or mr.min_row in trailer):
            r0 = rmap[mr.min_row]
            ws.merge_cells(start_row=r0, start_column=mr.min_col, end_row=r0 + mr.max_row - mr.min_row,
                           end_column=min(mr.max_col, last))
    for c in range(1, last + 1):
        L = get_column_letter(c)
        if L in src.column_dimensions and src.column_dimensions[L].width:
            ws.column_dimensions[L].width = src.column_dimensions[L].width
    ps, sps = ws.page_setup, src.page_setup
    ps.orientation, ps.paperSize, ps.scale = sps.orientation, sps.paperSize, sps.scale
    ps.fitToWidth, ps.fitToHeight = sps.fitToWidth, sps.fitToHeight
    if src.sheet_properties.pageSetUpPr:
        ws.sheet_properties.pageSetUpPr.fitToPage = src.sheet_properties.pageSetUpPr.fitToPage
    for k in ("left", "right", "top", "bottom", "header", "footer"):
        setattr(ws.page_margins, k, getattr(src.page_margins, k))
    ws.print_options.horizontalCentered = src.print_options.horizontalCentered
    ws.print_options.verticalCentered = src.print_options.verticalCentered
    if src.print_title_rows:
        ws.print_title_rows = f"1:{rmap.get(hrow, 2)}"
    if src.auto_filter.ref:
        ws.auto_filter.ref = f"A{rmap[hrow]}:{get_column_letter(last)}{rmap[hrow] + 1}"
    ws.sheet_view.zoomScale = src.sheet_view.zoomScale
    return wb


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    for k in OUT:
        ap.add_argument(f"--{k}")
    ap.add_argument("--out", default=os.path.join(ROOT, "templates"))
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    for kind in OUT:
        src = getattr(args, kind)
        if not src:
            continue
        wb = from_xlsx(kind, src) if src.lower().endswith((".xlsx", ".xlsm")) else from_xls(kind, src)
        wb.properties.creator = wb.properties.lastModifiedBy = "거성호텔 마감 도우미"
        dst = os.path.join(args.out, OUT[kind])
        wb.save(dst)
        print(f"{kind}: {dst}")


if __name__ == "__main__":
    sys.exit(main())
