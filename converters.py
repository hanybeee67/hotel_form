# -*- coding: utf-8 -*-
"""거성호텔 나이트 마감 자동화 - 변환 로직 (GUI와 분리되어 단독 테스트 가능)"""
import csv
import datetime as dt
import json
import os
import re
import time
import sys
from copy import copy
from html.parser import HTMLParser

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break


class ConversionError(Exception):
    """사용자에게 그대로 보여 줄 수 있는 오류 메시지를 담는다."""


# ---------------------------------------------------------------- 설정 로드
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bundled_dir():
    return getattr(sys, "_MEIPASS", app_dir())


def load_config():
    """exe 옆의 config.json을 먼저 읽고, 없으면 exe 안에 포함된 기본값을 사용."""
    for base in (app_dir(), bundled_dir()):
        path = os.path.join(base, "config.json")
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return json.load(f)
    raise ConversionError("config.json 파일을 찾을 수 없습니다.")


# ---------------------------------------------------------------- 파일 읽기
class _TableParser(HTMLParser):
    """ERP가 'xls'라는 이름으로 내보내는 HTML 표 파일을 읽기 위한 간단한 파서."""

    def __init__(self):
        super().__init__()
        self.tables, self._t, self._row, self._cell = [], None, None, None

    def handle_starttag(self, tag, attrs):
        if tag == "table" and self._t is None:
            self._t = []
            self.tables.append(self._t)
        elif tag == "tr" and self._t is not None:
            self._row = []
            self._t.append(self._row)
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr":
            self._row = None
        elif tag == "table":
            self._t = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _decode(raw):
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def read_table(path):
    """엑셀(xlsx/xls), HTML 위장 xls, CSV를 모두 2차원 리스트로 읽는다."""
    if not os.path.exists(path):
        raise ConversionError(f"파일을 찾을 수 없습니다: {path}")
    with open(path, "rb") as f:
        raw = f.read()
    head = raw[:2048]

    if head[:2] == b"PK":  # xlsx / xlsm
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows

    if head[:4] == b"\xd0\xcf\x11\xe0":  # 구형 xls
        try:
            import xlrd
        except ImportError:
            raise ConversionError("구형 .xls 파일을 읽으려면 xlrd가 필요합니다. "
                                  "엑셀에서 .xlsx로 저장한 뒤 다시 시도해 주세요.")
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        rows = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    row.append(xlrd.xldate.xldate_as_datetime(cell.value, book.datemode))
                else:
                    row.append(cell.value)
            rows.append(row)
        return rows

    text = _decode(raw)
    if head.lstrip()[:1] == b"<" or "<table" in text[:5000].lower():
        parser = _TableParser()
        parser.feed(text)
        if not parser.tables:
            raise ConversionError("파일에서 표를 찾지 못했습니다.")
        return max(parser.tables, key=len)

    dialect = csv.excel_tab if text.count("\t") > text.count(",") else csv.excel
    return [row for row in csv.reader(text.splitlines(), dialect)]


# ---------------------------------------------------------------- 값 정리
def norm(v):
    if v is None:
        return ""
    return re.sub(r"[\s_\-\(\)\[\]/\.\*:#]", "", str(v)).lower()


def clean_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def to_number(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else float(v)
    s = str(v).strip().replace(",", "").replace("원", "").replace("₩", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        n = float(s)
    except ValueError:
        return None
    n = -n if neg else n
    return int(n) if n.is_integer() else n


def to_date(v):
    if v is None or v == "":
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and 30000 < float(v) < 80000:  # 엑셀 일련번호
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.search(r"(\d{4})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})", s)
    if m:
        return f"{int(m[1]):04d}-{int(m[2]):02d}-{int(m[3]):02d}"
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s)
    if m:
        return f"{m[1]}-{m[2]}-{m[3]}"
    m = re.fullmatch(r"(\d{2})[-./](\d{1,2})[-./](\d{1,2})", s)
    if m:
        return f"20{m[1]}-{int(m[2]):02d}-{int(m[3]):02d}"
    return s


def mask_card(s):
    """마스킹되지 않은 카드번호만 앞4/뒤4자리만 남기고 가린다."""
    if "*" in s:
        return s
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 13:
        return f"{digits[:4]}-****-****-{digits[-4:]}"
    return s


# ---------------------------------------------------------------- 헤더 탐지
def find_header(rows, fields, max_scan=40):
    """상단 max_scan행 중 별칭이 가장 많이 일치하는 행을 헤더로 판단한다.
    1단계: 정확히 일치, 2단계: 포함 관계(남은 열/항목만)."""
    alias_norm = {k: [norm(a) for a in v] for k, v in fields.items()}
    best = (0, None, {})
    for i, row in enumerate(rows[:max_scan]):
        mapping, used = {}, set()
        for j, cell in enumerate(row):
            n = norm(cell)
            if not n:
                continue
            for key, al in alias_norm.items():
                if key not in mapping and n in al:
                    mapping[key] = j
                    used.add(j)
                    break
        for j, cell in enumerate(row):
            n = norm(cell)
            if not n or j in used:
                continue
            for key, al in alias_norm.items():
                if key in mapping:
                    continue
                if any(len(a) >= 2 and a in n for a in al):
                    mapping[key] = j
                    used.add(j)
                    break
        if len(mapping) > best[0]:
            best = (len(mapping), i, mapping)
    return best


_TOTAL_WORDS = {"합계", "소계", "총계", "계", "total", "subtotal", "grandtotal"}


def _is_total_row(row):
    # 수입일보는 매출구분 칸에 'ROOM 소계', '조식 소계'처럼 적혀 있다 (2026-10-28 실제 파일)
    return any(norm(c) in _TOTAL_WORDS or norm(c).endswith(("소계", "합계", "총계")) for c in row[:4])


def fmt_room(v, mode):
    s = clean_text(v)
    if re.fullmatch(r"\d{1,4}", s):
        return s.zfill(4) if mode == "pad4" else str(int(s))
    return s


def _room_key(rec):
    r = rec.get("room", "") or ""
    return (0, int(r)) if r.isdigit() else (1, r)


# ---------------------------------------------------------------- 원본 → 레코드
def _extract(conv, rows):
    """원본 표 → {field: value} 레코드 목록. (records, warnings, mapped_fields)"""
    fields = conv["source_fields"]
    score, hidx, mapping = find_header(rows, fields)
    warnings = []

    for tgt, alt in conv.get("fallback", {}).items():
        if tgt not in mapping and alt in mapping:
            mapping[tgt] = mapping[alt]
            warnings.append(f"'{fields[tgt][0]}' 열이 없어 '{fields[alt][0]}' 열로 대체했습니다.")

    missing_req = [k for k in conv["required"] if k not in mapping]
    if hidx is None or score == 0 or missing_req:
        names = ", ".join(fields[k][0] for k in missing_req) or "제목 행"
        found = [clean_text(c) for c in (rows[hidx] if hidx is not None else []) if clean_text(c)]
        raise ConversionError(
            f"원본에서 필수 항목({names})을 찾지 못했습니다.\n"
            f"선택한 파일이 맞는지 확인해 주세요.\n(인식된 제목: {', '.join(found[:12]) or '없음'})")

    for k in conv.get("warn_missing", []):
        if k not in mapping:
            warnings.append(f"원본에 '{fields[k][0]}' 열이 없어 빈칸으로 남깁니다.")

    types = conv.get("types", {})
    include = {f: {norm(v) for v in vals} for f, vals in conv.get("include", {}).items()}
    exclude = {f: {norm(v) for v in vals} for f, vals in conv.get("exclude", {}).items()}
    records, src_total = [], None
    lo, hi = conv.get("room_range") or (None, None)
    out_of_range = []
    for row in rows[hidx + 1:]:
        if not any(clean_text(c) for c in row):
            continue
        if _is_total_row(row):
            src_total = row
            continue
        rec = {}
        for key, j in mapping.items():
            val = row[j] if j < len(row) else None
            t = types.get(key)
            if t == "number":
                rec[key] = to_number(val)
            elif t == "date":
                rec[key] = to_date(val)
            else:
                rec[key] = clean_text(val)
        if any(f in rec and norm(rec[f]) not in vals for f, vals in include.items()):
            continue
        if any(f in rec and norm(rec[f]) in vals for f, vals in exclude.items()):
            continue
        if "room" in rec and conv.get("room_format"):
            rec["room"] = fmt_room(rec["room"], conv["room_format"])
        if conv.get("require_room") and not rec.get("room"):
            continue
        if lo is not None and rec.get("room"):  # 호텔 객실(301~713)이 아닌 번호(가상객실·단체 계정 등)는 뺀다
            n = int(rec["room"]) if rec["room"].isdigit() else None
            if n is None or not lo <= n <= hi:
                out_of_range.append(rec["room"])
                continue
        if not any(v not in ("", None) for v in rec.values()):
            continue
        records.append(rec)
    if out_of_range:
        shown = ", ".join(sorted(set(out_of_range))[:8]) + (" …" if len(set(out_of_range)) > 8 else "")
        warnings.append(f"객실 범위({lo}~{hi}) 밖의 번호 {len(out_of_range)}건은 뺐습니다: {shown}")

    # 원본에 합계 행이 있으면 우리가 계산한 합계와 대조
    if conv.get("check_total") and src_total is not None and conv.get("sum"):
        fld = conv["sum"][0]
        j = mapping.get(fld)
        if j is not None and j < len(src_total):
            tv = to_number(src_total[j])
            mine = sum(r.get(fld) or 0 for r in records)
            if tv is not None and abs(tv - mine) > 0.5:
                warnings.append(f"원본 합계({tv:,.0f})와 변환 합계({mine:,.0f})가 다릅니다. 원본을 확인해 주세요.")
    return records, warnings, set(mapping)


def _is_breakfast(val, no_values):
    if val is None:
        return False
    n = norm(val)
    return n not in {norm(x) for x in no_values} and (to_number(val) is None or to_number(val) > 0)


# ---------------------------------------------------------------- 엑셀 출력
_THIN = Side(style="thin", color="000000")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_BLANK_FILL = PatternFill("solid", fgColor="FFFFF2B3")  # 근무자가 직접 확인/입력할 칸
_NO_FILL = PatternFill(fill_type=None)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ONE_LINE = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)  # 호텔 양식처럼 한 줄 (넘치면 글자 축소)
_DEFAULT_FONT = {"name": "굴림", "title": 20, "header": 9, "body": 9}
_GREY = "666666"


def resolve_template(rel):
    """양식 파일은 exe 옆의 templates/를 먼저 찾고, 없으면 exe 안에 포함된 것을 찾는다."""
    if not rel:
        return None
    for base in (app_dir(), bundled_dir()):
        p = os.path.join(base, rel)
        if os.path.exists(p):
            return p
    return None


def _style_of(cell):
    return {"font": copy(cell.font), "border": copy(cell.border), "fill": copy(cell.fill),
            "alignment": copy(cell.alignment), "number_format": cell.number_format}


def _apply(cell, st, **over):
    st = {**st, **over}
    cell.font, cell.border, cell.fill = copy(st["font"]), copy(st["border"]), copy(st["fill"])
    cell.alignment = copy(st["alignment"])
    if st.get("number_format"):
        cell.number_format = st["number_format"]


def _has_style(cell):
    b = cell.border
    return any(getattr(b, s).style for s in ("left", "right", "top", "bottom")) or cell.fill.fill_type is not None


class _Sheet:
    """출력 시트와 서식 정보. 호텔 양식(템플릿)이 있으면 그 서식을, 없으면 config의 폰트를 쓴다."""

    def __init__(self):
        self.wb = self.ws = None
        self.hrow = 0
        self.colmap = {}        # out_cols 순번 → 시트 열 번호(1부터)
        self.data_st = {}       # 시트 열 → 데이터 칸 서식
        self.foot_st = {}       # 시트 열 → 합계 칸 서식
        self.data_height = None
        self.foot_height = None
        self.note_font = None
        self.from_template = False
        self.warnings = []
        self.extra_cols = []    # 양식 표 안이지만 값을 채우지 않는 열(예: 병합된 Sign 칸의 오른쪽 열)
        self.trailer = []       # 표 아래에 그대로 둘 행(예: 조식명단의 외부식사·안내 문구)
        self.frame_last = 0     # 양식 표의 마지막 열


def _fill(color):
    return PatternFill("solid", fgColor=color) if color else _NO_FILL


def _new_sheet(sheet_name, title_text, out_cols, font_cfg, row_height):
    """템플릿이 없을 때: config의 font 설정(폰트·크기·행 높이)으로 새 양식을 만든다."""
    f = {**_DEFAULT_FONT, **(font_cfg or {})}
    s = _Sheet()
    s.wb = Workbook()
    ws = s.ws = s.wb.active
    ws.title = sheet_name[:31]
    nprint = sum(1 for c in out_cols if c.get("print", True))

    t = ws.cell(1, 1, title_text)
    t.font = Font(name=f["name"], size=f["title"], bold=f.get("title_bold", True))
    t.alignment = _CENTER
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nprint)
    ws.row_dimensions[1].height = f.get("title_height", 44)
    s.hrow = 2
    ws.row_dimensions[2].height = f.get("header_height", 28)
    head_fill = _fill(f.get("header_fill", "E7EBF0"))
    sum_fill = _fill(f.get("sum_fill", "F2F2F2"))
    for i, col in enumerate(out_cols):
        c = i + 1
        s.colmap[i] = c
        color = _GREY if not col.get("print", True) else "000000"
        cell = ws.cell(2, c, col["header"])
        cell.font = Font(name=f["name"], size=f["header"], bold=f.get("header_bold", True), color=color)
        cell.fill, cell.border, cell.alignment = copy(head_fill), _BORDER, _CENTER
        ws.column_dimensions[get_column_letter(c)].width = col.get("width", 12)
        s.data_st[c] = {"font": Font(name=f["name"], size=f["body"], bold=f.get("body_bold", False), color=color),
                        "border": _BORDER, "fill": _NO_FILL, "alignment": _ONE_LINE, "number_format": None}
        s.foot_st[c] = {"font": Font(name=f["name"], size=f["body"], bold=f.get("sum_bold", True)),
                        "border": _BORDER, "fill": sum_fill, "alignment": _CENTER, "number_format": None}
    s.data_height = row_height or f.get("row_height")
    s.foot_height = s.data_height
    s.note_font = Font(name=f["name"], size=max(f["body"] - 1, 8), italic=True, color=_GREY)
    ws.freeze_panes = ws.cell(3, 1)
    ws.page_setup.orientation = f.get("orientation", "portrait")
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.print_title_rows = "2:2"
    return s


def _template_sheet(path, sheet_name, title_text, out_cols):
    """호텔 양식을 열어 머리글 행을 찾고, 그 아래 데이터·합계 행의 서식을 기억한 뒤 비운다.
    폰트·열 너비·제목 병합·인쇄 설정은 양식 그대로 유지된다."""
    try:
        wb = load_workbook(path)
    except Exception as e:
        raise ConversionError(f"호텔 양식 파일을 열 수 없습니다: {os.path.basename(path)}\n({e})")
    ws = wb.worksheets[0]
    heads = {}  # 양식 머리글 → out_cols 순번 (header와 aliases 모두 인정)
    for i, col in enumerate(out_cols):
        for h in [col["header"]] + col.get("aliases", []):
            heads.setdefault(norm(h), i)

    best = (0, None, {})
    for r in range(1, min(ws.max_row, 30) + 1):
        m = {}
        for c in range(1, ws.max_column + 1):
            i = heads.get(norm(ws.cell(r, c).value))
            if i is not None and i not in m:
                m[i] = c
        if len(m) > best[0]:
            best = (len(m), r, m)
    if best[0] < 2:
        raise ConversionError(f"호텔 양식({os.path.basename(path)})에서 머리글 행(객실 등)을 찾지 못했습니다.\n"
                              "양식의 열 제목이 config.json의 output 'header'와 같은지 확인해 주세요.")
    s = _Sheet()
    s.wb, s.ws, s.from_template = wb, ws, True
    s.hrow, s.colmap = best[1], dict(best[2])
    hrow = s.hrow

    # 양식에 없는 열(예: 인쇄 안 되는 Acct# 참고 열)은 오른쪽 끝에 덧붙인다
    nxt = max(max(s.colmap.values()), ws.max_column) + 1
    for i, col in enumerate(out_cols):
        if i in s.colmap:
            continue
        s.colmap[i] = nxt
        src = ws.cell(hrow, max(best[2].values()))
        cell = ws.cell(hrow, nxt, col["header"])
        _apply(cell, _style_of(src))
        cell.font = copy(src.font) if col.get("print", True) else Font(
            name=src.font.name, size=src.font.size, bold=src.font.bold, color=_GREY)
        ws.column_dimensions[get_column_letter(nxt)].width = col.get("width", 12)
        if col.get("print", True):
            s.warnings.append(f"호텔 양식에 '{col['header']}' 열이 없어 오른쪽 끝에 추가했습니다.")
        nxt += 1

    # 제목: 머리글 위쪽의 첫 글자 칸을 이번 날짜 제목으로 교체(서식은 그대로)
    done = False
    for r in range(1, hrow):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip():
                ws.cell(r, c).value = title_text
                done = True
                break
        if done:
            break
    if not done:
        s.warnings.append("호텔 양식에서 제목 칸을 찾지 못해 제목을 넣지 않았습니다.")

    # 데이터 행 서식: 머리글 바로 아래 행. 서식이 없으면 머리글 서식에서 굵게만 뺀다
    drow = hrow + 1
    use_data = drow <= ws.max_row and any(_has_style(ws.cell(drow, c)) for c in best[2].values())
    foot_row = None
    for r in range(hrow + 1, ws.max_row + 1):
        if any(norm(ws.cell(r, c).value) in ("계", "합계", "총계") for c in range(1, ws.max_column + 1)):
            foot_row = r
            break
    if foot_row == drow:
        use_data = False
    for i, c in s.colmap.items():
        if use_data:
            st = _style_of(ws.cell(drow, c))
        else:
            st = _style_of(ws.cell(hrow, c))
            st["font"] = copy(st["font"])
            st["font"].b = False
            st["fill"] = _NO_FILL
        if not out_cols[i].get("print", True):
            fnt = copy(st["font"])
            fnt.color = _GREY
            st["font"] = fnt
        s.data_st[c] = st
        if foot_row:
            s.foot_st[c] = _style_of(ws.cell(foot_row, c))
        else:
            fst = dict(st)
            fnt = copy(st["font"])
            fnt.b = True
            fst["font"] = fnt
            s.foot_st[c] = fst
    rd = ws.row_dimensions
    s.data_height = rd[drow].height if use_data and drow in rd else None
    s.foot_height = rd[foot_row].height if foot_row and foot_row in rd else s.data_height
    base_font = s.data_st[s.colmap[0]]["font"]
    s.note_font = Font(name=base_font.name, size=base_font.size, italic=True, color=_GREY)

    # 양식 표의 너비(머리글 병합 포함) 안에서 값을 넣지 않는 열도 테두리·글꼴을 똑같이 그린다
    s.frame_last = max(best[2].values())
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= hrow <= mr.max_row and mr.min_col <= s.frame_last:
            s.frame_last = max(s.frame_last, mr.max_col)
    for c in range(1, s.frame_last + 1):
        if c in s.colmap.values():
            continue
        s.extra_cols.append(c)
        s.data_st[c] = _style_of(ws.cell(drow if use_data else hrow, c))
        s.foot_st[c] = _style_of(ws.cell(foot_row, c)) if foot_row else s.data_st[c]

    # 표 아래 안내 행: 합계 행 뒤(합계가 없으면 데이터 서식 행 뒤)에 있고, 첫 칸이 객실번호가 아닌 글자인 행
    start = (foot_row or drow) + 1
    for r in range(start, ws.max_row + 1):
        first = clean_text(ws.cell(r, 1).value)
        if not first or first.isdigit() or norm(first) in _TOTAL_WORDS:
            continue
        cells = [(c, ws.cell(r, c).value, _style_of(ws.cell(r, c))) for c in range(1, s.frame_last + 1)]
        merges = [(mr.min_col, mr.max_col, mr.max_row - mr.min_row) for mr in ws.merged_cells.ranges
                  if mr.min_row == r]
        s.trailer.append((cells, rd[r].height if r in rd else None, merges))

    # 머리글 아래(예시 행·빈 행·기존 합계 행)를 비운다
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row > hrow:
            ws.unmerge_cells(str(mr))
    if ws.max_row > hrow:
        ws.delete_rows(hrow + 1, ws.max_row - hrow)
    for k in [k for k in rd if k > hrow]:
        del rd[k]
    if not ws.print_title_rows:
        ws.print_title_rows = f"{hrow}:{hrow}"
    ws.title = sheet_name[:31]
    return s


_LINE = "FF000000"  # 불투명 검정 (앞 FF = 투명도 없음. 00이면 일부 프로그램에서 선이 안 보일 수 있음)


def _is_blank_fill(cell):
    return cell.fill.fill_type == "solid" and (cell.fill.fgColor.rgb or "").upper().endswith("FFF2B3")


def _draw_grid(ws, r0, r1, c0, c1):
    """표 전체에 검은 선을 긋는다: 안쪽은 가는 선, 바깥 테두리(와 제목 아래)는 굵은 선 (어느 화면·인쇄에서도 보이게 색을 지정).
    노란 입력 칸도 다른 칸과 같은 가는 선 (사용자 지시 2026-11-02: 인쇄하면 노란색은 안 나오는데 그 칸 테두리만 진하게 나옴.
    2026-10-23에 넣었던 '노란 칸 둘레 굵은 선'은 제거)."""
    thin, thick = Side(style="thin", color=_LINE), Side(style="medium", color=_LINE)
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(r, c)
            if cell.fill.fill_type == "solid" and (cell.fill.fgColor.rgb or "").upper().endswith("FFFFFF"):
                cell.fill = PatternFill(fill_type=None)  # 흰색 채우기는 필요 없음 (노란 입력 칸만 색 유지)
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(r, c).border = Border(
                left=thick if c == c0 else thin,
                right=thick if c == c1 else thin,
                top=thick if r in (r0, r0 + 1) else thin,
                bottom=thick if r in (r1, r0) else thin)


def _fit_a4(ws, hrow, keep_margins=False):
    """A4 한 장 너비에 맞추고(세로는 필요한 만큼), 가운데 정렬·쪽마다 제목/머리글 반복·쪽 번호.
    keep_margins: 호텔 양식(샘플)의 여백을 그대로 둔다 (사용자 지시 2026-11-02: 간격·크기를 샘플과 똑같이)"""
    ps = ws.page_setup
    ps.paperSize = ws.PAPERSIZE_A4
    ps.scale = None
    ps.fitToWidth, ps.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    if not keep_margins:
        m = ws.page_margins
        m.left = m.right = 0.4
        m.top = m.bottom = 0.5
        m.header = m.footer = 0.3
    ws.print_title_rows = f"1:{hrow}"
    ws.oddFooter.center.text = "&P / &N"


def _col_points(width, px=7):
    """엑셀 열 너비(글자 수 단위) → 포인트 (기본 글꼴 최대 숫자 폭 px 기준, 보통 7px)"""
    return ((width or 8.43) * px + 5) * 0.75


# 호텔 PC 프로그램은 열을 7px 계산보다 넓게 그린다 (2026-10-27 사진: 조식명단 Sign 열이 잘려 4쪽으로 인쇄됨)
# → 꽉 채우는 조식명단은 넉넉하게 계산해 가로가 넘치지 않게 한다. 9px는 가로 여백이 많이 남아(같은 날 사용자) 8px로.
# config.json의 breakfast.width_px로 조절: 숫자를 줄이면 표가 가로로 커지고, 늘리면 작아진다
SAFE_PX = 8


def _scale_sheet(ws, last, right, f):
    """열 너비·글자 크기·행 높이를 같은 비율 f로 줄이거나 키운다 (표 모양은 그대로, 크기만)"""
    for c in range(1, right + 1):
        d = ws.column_dimensions[get_column_letter(c)]
        d.width = round((d.width or 8.43) * f, 2)
    for r in range(1, last + 1):
        if ws.row_dimensions[r].height:
            ws.row_dimensions[r].height = round(ws.row_dimensions[r].height * f, 2)
        for c in range(1, right + 1):
            cell = ws.cell(r, c)
            font = copy(cell.font)
            font.sz = max(6, round((cell.font.sz or 11) * f * 2) / 2)  # 0.5pt 단위
            cell.font = font


def _fit_pages(ws, hrow, first, last, right, pages, fill=False, min_ratio=0.75, width_px=None):
    """A4 'pages'장 안에 들어가도록: 너비는 1장, 세로는 최대 pages장.
    넘치면 글자 크기는 그대로 두고 행 높이(줄 간격)만 필요한 만큼 줄인다 (사용자 지시 2026-10-26).
    fill=True(조식명단, 2026-10-27): 1장 분량이면 1장, 넘치면 2장(최대 pages장)을 **꽉 채우도록** 행 높이를 늘리거나 줄인다.
    그래도 넘치면 엑셀의 '1쪽 너비 × N쪽 높이 맞춤'이 마지막으로 축소한다."""
    landscape = ws.page_setup.orientation == "landscape"
    page_w, page_h = (842, 595) if landscape else (595, 842)
    m = ws.page_margins
    usable_w = page_w - (m.left + m.right) * 72
    usable_h = page_h - (m.top + m.bottom) * 72
    widths = lambda px: sum(_col_points(ws.column_dimensions[get_column_letter(c)].width, px)
                            for c in range(1, right + 1))
    if fill:
        # 가로도 꽉 차게 (사용자 지시 2026-10-27): 넉넉한 폭 계산으로 A4 너비에 맞게 열 너비·글자 크기를 같이 조절,
        # 배율 100%로 인쇄 → 세로도 이 크기 기준으로 채운다
        f = max(0.5, min(1.3, usable_w * 0.98 / widths(width_px or SAFE_PX)))
        _scale_sheet(ws, last, right, f)
        total_w, scale = usable_w, 1.0
    else:
        total_w = widths(7)
        scale = min(1.0, usable_w / total_w) if total_w else 1.0
    height = lambda r: ws.row_dimensions[r].height or 15
    repeat = sum(height(r) for r in range(1, hrow + 1))  # 쪽마다 반복되는 제목·머리글
    per_page = usable_h / scale - repeat  # 한 쪽에서 표 본문이 쓸 수 있는 높이
    body = sum(height(r) for r in range(first, last + 1))
    n = pages
    if fill:
        # 양식 행 높이를 min_ratio(기본 75%, Sign 칸에 서명할 공간)까지 줄여서 들어가면 그 쪽 수로 — 최대 pages장
        n = next((k for k in range(1, pages + 1) if body * min_ratio <= k * per_page), pages)
        target = n * per_page * 0.97  # 행 단위로 쪽이 나뉘므로 약간 여유를 둔다
    else:
        target = pages * per_page
    if fill and n == 2 and _fill_two_equal_pages(ws, first, last, right, usable_h, repeat):
        return
    if body and (fill or body > target):
        k = target / body
        for r in range(first, last + 1):
            size = max((ws.cell(r, c).font.sz or 11) for c in range(1, right + 1))
            lines = max(str(ws.cell(r, c).value or "").count("\n") + 1 for c in range(1, right + 1))
            floor = size * 1.3 * lines  # 글자가 잘리지 않는 최소 높이
            ws.row_dimensions[r].height = round(min(409, max(floor, height(r) * k)), 2)  # 409 = 엑셀 행 높이 최대
    ps = ws.page_setup
    ps.fitToWidth, ps.fitToHeight = 1, n


def _wrapped_lines(ws, r, c, right):
    """칸 글자가 몇 줄로 그려지는지 어림 (병합 칸은 병합 너비, 한글은 숫자 약 1.8자 폭).
    2026-11-03: 조식명단 행이 많아 칸 높이가 낮아지자 표 아래 안내 문구(한/영, 줄바꿈 포함)가 잘림 → 줄바꿈된 줄 수까지 셈"""
    text = str(ws.cell(r, c).value or "")
    if text.startswith("="):  # 수식(합계 =SUM…)은 결과 숫자만 보인다
        return 1
    size = ws.cell(r, c).font.sz or 11
    c1 = next((m.max_col for m in ws.merged_cells.ranges if m.min_row == r and m.min_col == c), c)
    width = sum(ws.column_dimensions[get_column_letter(k)].width or 8.43 for k in range(c, min(c1, right) + 1))
    cap = max(1.0, width * 11 / size * 0.95)
    units = lambda t: sum(1.8 if ord(ch) > 0x2E80 else 1.0 for ch in t)
    return sum(max(1, -(-units(line) // cap)) for line in text.split("\n"))


def _row_floor(ws, r, right):
    """글자가 잘리지 않는 최소 행 높이"""
    size = max((ws.cell(r, c).font.sz or 11) for c in range(1, right + 1))
    lines = max(_wrapped_lines(ws, r, c, right) for c in range(1, right + 1))
    return size * 1.3 * lines


def _fill_two_equal_pages(ws, first, last, right, usable_h, repeat):
    """조식명단 2장 (사용자 지시 2026-10-27: 두 장을 나란히 놓으므로 표 크기가 똑같고 맨 아래 줄이 딱 맞아야 함).
    모든 행을 같은 높이(한 칸)로 만들고, 1쪽 k칸 / 2쪽 k칸이 되도록 수동 쪽 나눔을 넣는다.
    2쪽 마지막 행(안내 문구)은 남는 칸 수만큼 높이를 차지해 두 쪽의 맨 아래 선이 같아진다.
    '맞춤' 인쇄는 수동 쪽 나눔을 무시하므로 배율 100%로 인쇄한다(너비는 _scale_sheet가 이미 맞춤). 안 되면 False."""
    scale_pct = 100
    per_page = (usable_h - repeat) * 0.98  # 쪽마다 표 본문이 쓸 수 있는 높이 (여유 2%)
    others = last - first  # 마지막 행(안내 문구 또는 합계) 앞까지의 행 수
    floor_other = max(_row_floor(ws, r, right) for r in range(first, last))
    floor_last = _row_floor(ws, last, right)
    k = max(1, -(-(others + 1) // 2))
    for _ in range(10):  # 칸 높이 ↔ 마지막 행이 차지할 칸 수가 서로 맞을 때까지
        h = per_page / k
        m = max(1, -(-floor_last // h)) if h else 1
        k_new = max(1, -(-(others + int(m)) // 2))
        if k_new == k:
            break
        k = k_new
    h = per_page / k
    if h < floor_other or k > others:  # 2장 안에 안 들어가거나 1쪽을 채울 행이 모자라면 포기
        return False
    for r in range(first, last):
        ws.row_dimensions[r].height = round(h, 2)
    ws.row_dimensions[last].height = round(h * (2 * k - others), 2)
    ws.row_breaks.brk = []
    ws.row_breaks.append(Break(id=first + k - 1))  # 1쪽 = 머리글 + k칸
    ps = ws.page_setup
    ps.scale = scale_pct
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ps.fitToWidth, ps.fitToHeight = 1, 2
    return True


_LOOKUP_SHEET = "카드목록"


def _manual_card_formulas(ws, r, out_cols, idx, letter, look):
    """한벽루 줄 (사용자 지시 2026-11-01): 카드번호 앞 4자리 → 매입사, 사업자번호+매입사 → 가맹점번호, 구분 = 카드승인.
    사업자번호는 한벽루 번호가 기본이고 목록에서 호텔 번호로 바꿀 수 있다. 근무자가 칸에 직접 쓰면 그 값이 우선."""
    f = {"card": "card_no", "biz": "biz_no", "acq": "acquirer", "merchant": "merchant_no", "kind": "kind"}
    f.update(look.get("fields") or {})
    if not all(f[k] in idx for k in ("card", "biz", "acq", "merchant")):
        return
    L = {k: letter(idx[v]) for k, v in f.items() if v in idx}
    cell = lambda k: ws[f"{L[k]}{r}"]
    blank = f'{L["card"]}{r}=""'
    if look.get("other"):
        cell("biz").value = f'=IF({blank},"","{look["other"]}")'
    cell("acq").value = f'=IF({blank},"",IFERROR(VLOOKUP(LEFT({L["card"]}{r},4),{_LOOKUP_SHEET}!$A:$B,2,FALSE),""))'
    cell("merchant").value = (f'=IF({L["acq"]}{r}="","",IFERROR(VLOOKUP({L["biz"]}{r}&"|"&{L["acq"]}{r},'
                              f'{_LOOKUP_SHEET}!$D:$E,2,FALSE),""))')
    if "kind" in L:
        cell("kind").value = f'=IF({blank},"","카드승인")'
    choices = [b for b in (look.get("other"), look.get("hotel")) if b]
    if choices:
        dv = DataValidation(type="list", formula1='"' + ",".join(choices) + '"', allow_blank=True, showErrorMessage=False)
        dv.promptTitle, dv.prompt, dv.showInputMessage = "사업자번호", "한벽루 / 호텔 중 고르세요", True
        ws.add_data_validation(dv)
        dv.add(f"{L['biz']}{r}")


def _add_card_lookup_sheet(wb, look):
    """숨긴 시트 '카드목록': A·B = 카드번호 앞 4자리 → 매입사, D·E = '사업자번호|매입사' → 가맹점번호 (수식이 참고)"""
    ws = wb.create_sheet(_LOOKUP_SHEET)
    ws["A1"], ws["B1"], ws["D1"], ws["E1"] = "앞4자리", "매입사", "사업자번호|매입사", "가맹점번호"
    for i, (p, acq) in enumerate(sorted(look.get("prefix", {}).items()), 2):
        ws.cell(i, 1, p).number_format = "@"
        ws.cell(i, 2, acq)
    for i, (k, m) in enumerate(sorted(look.get("merchant", {}).items()), 2):
        ws.cell(i, 4, k)
        ws.cell(i, 5, m).number_format = "@"
    ws.sheet_state = "hidden"


def write_workbook(path, sheet_name, title_text, out_cols, records, sum_fields,
                   number_format="0", notes=None, cash=None, row_height=None,
                   template=None, font_cfg=None, sum_label="계", sum_label_field=None, drop_trailer=(),
                   fit_pages=None, fill_pages=False, width_px=None, card_lookup=None):
    """template(호텔 양식 경로)가 있으면 양식을 열어 값만 채우고, 없으면 font_cfg로 새로 만든다.
    반환: 경고 목록"""
    if template:
        s = _template_sheet(template, sheet_name, title_text, out_cols)
        if row_height and s.data_height is None:
            s.data_height = row_height
    else:
        s = _new_sheet(sheet_name, title_text, out_cols, font_cfg, row_height)
    ws, hrow, colmap = s.ws, s.hrow, s.colmap
    # 열 찾기: 데이터 열은 field로, 모든 열은 header로도 찾는다
    # (누계처럼 formula_of만 있는 열이 원래 열(sales)을 덮어쓰면 자기 자신을 가리키는 순환 수식이 된다)
    idx = {}
    for i, c in enumerate(out_cols):
        if c.get("field"):
            idx.setdefault(c["field"], i)
        idx.setdefault(c["header"], i)
    printable = sorted(colmap[i] for i, c in enumerate(out_cols) if c.get("print", True))
    letter = lambda i: get_column_letter(colmap[i])

    r = hrow
    sections = []  # 수입일보: 매출구분 묶음마다 계 줄 [(첫 행, 끝 행, 계 행)]
    sec_start = None
    for n, rec in enumerate(records):
        r += 1
        if sec_start is None:
            sec_start = r
        if s.data_height:
            ws.row_dimensions[r].height = s.data_height
        for i, col in enumerate(out_cols):
            cell = ws.cell(r, colmap[i])
            st = s.data_st[colmap[i]]
            _apply(cell, st)
            fmt = st["number_format"] if st["number_format"] not in (None, "General") else number_format
            if col.get("formula_of"):
                cell.value = f"={letter(idx[col['formula_of']])}{r}"
                # 수입일보 누계 = 매출 × 숙박수 (호텔 실제 일보 2026-10-28). 숙박수가 없는 줄(조식·연회 등)은 매출 그대로
                t = col.get("times")
                if rec.get("_cum_prev") is not None:  # 거래내역 잔액: 어제까지 잔액 + 오늘 매출
                    cell.value = f"={rec['_cum_prev']}+{letter(idx[col['formula_of']])}{r}"
                elif t in idx and isinstance(rec.get(t), (int, float)):
                    cell.value += f"*{letter(idx[t])}{r}"
                cell.number_format = fmt
                continue
            f = col.get("field")
            val = rec.get(f) if f else None
            if val not in ("", None):
                cell.value = val
            if f in rec.get("_yellow", ()):  # 자동으로 채웠지만 확인이 필요한 칸 (예: 카드 객실 후보가 여러 개)
                cell.fill = copy(_BLANK_FILL)
            elif (f is None and col.get("fill", True)) or (col.get("fill_if_empty") and val in ("", None)
                                                          and not rec.get("_no_fill")):
                cell.fill = copy(_BLANK_FILL)  # 근무자가 직접 입력할 칸
            if isinstance(val, (int, float)):
                cell.number_format = fmt
        for c in s.extra_cols:
            _apply(ws.cell(r, c), s.data_st[c])
        if rec.get("_manual"):  # 카드: ERP에 없는 한벽루 단말기 승인 건을 직접 적는 빈 줄 (2026-10-28)
            for i, col in enumerate(out_cols):
                cell = ws.cell(r, colmap[i])
                if col.get("print", True):
                    cell.fill = copy(_BLANK_FILL)
                    if col.get("field") == "amount" or col.get("aliases") and "금액" in col["aliases"]:
                        cell.number_format = number_format
                elif rec.get("_note"):
                    cell.value = rec["_note"]
            if card_lookup:
                _manual_card_formulas(ws, r, out_cols, idx, letter, card_lookup)
        nxt = records[n + 1] if n + 1 < len(records) else None
        if "_sec" in rec and nxt is not None and nxt.get("_sec") != rec["_sec"]:
            r += 1  # 묶음 끝 → 계 줄 자리 (서식·수식은 아래에서)
            sections.append((sec_start, r - 1, r))
            sec_start = None

    first, last = hrow + 1, r

    # 직접 입력 칸: 선택 목록(예: 선/후)이나 숫자만 입력되게 도움말을 붙인다
    manual = records and any((c.get("field") is None and c.get("fill", True)) or
                             (c.get("fill_if_empty") and any(rec.get(c["field"]) in ("", None) for rec in records))
                             for c in out_cols)
    for i, col in enumerate(out_cols):
        if not records or (col.get("field") is not None and not col.get("choices")):
            continue
        rng = f"{letter(i)}{first}:{letter(i)}{last}"
        if col.get("choices"):
            # choices_free: 목록에서 고르거나 '선(카드)10/27'처럼 직접 써도 됨 (호텔 실제 일보)
            dv = DataValidation(type="list", formula1='"' + ",".join(col["choices"]) + '"', allow_blank=True,
                                showErrorMessage=not col.get("choices_free"), errorTitle="입력 확인",
                                error=f"{' / '.join(col['choices'])} 중에서 골라 주세요.")
        elif col.get("input") == "int":
            dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                                showErrorMessage=True, errorTitle="입력 확인", error="숫자만 입력해 주세요.")
        else:
            continue
        dv.promptTitle, dv.prompt, dv.showInputMessage = col["header"], "직접 확인 후 입력하세요.", True
        ws.add_data_validation(dv)
        dv.add(rng)

    def _foot_row(row):
        if s.foot_height:
            ws.row_dimensions[row].height = s.foot_height
        for c in printable + s.extra_cols:
            _apply(ws.cell(row, c), s.foot_st[c])

    def _foot(row, i, value, fill=None):
        cell = ws.cell(row, colmap[i])
        cell.value = value
        if fill is not None:
            cell.fill = copy(fill)
        if not isinstance(value, str) or value.startswith("="):
            fmt = s.foot_st[colmap[i]]["number_format"]
            cell.number_format = fmt if fmt not in (None, "General") else number_format

    if records and cash is not None:  # 카드 양식: 현금영수증 줄 + 계
        r += 1
        _foot_row(r)
        # 라벨 칸: 구분 열이 있으면 구분, 없으면(11/01 샘플 양식) 승인번호 칸
        _foot(r, idx["kind"] if "kind" in idx else idx["approval_no"], "현금영수증")
        amt = idx["amount"]
        _foot(r, amt, cash["amount"], None if cash["amount"] is not None else _BLANK_FILL)
        cash_row = r
        r += 1
        _foot_row(r)
        L = letter(amt)
        _foot(r, idx["approval_no"], "계")
        _foot(r, amt, f"=SUM({L}{first}:{L}{last})+{L}{cash_row}")
    elif records and sum_fields:
        r += 1
        if sections or "_sec" in records[-1]:
            sections.append((sec_start or r, r - 1, r))
        sum_idx = [idx[f] for f in sum_fields if f in idx]
        label_col = None
        if sum_idx:
            label_i = min(sum_idx, key=lambda i: colmap[i])
            label_col = max(colmap[label_i] - 1, printable[0])
            if sum_label_field in idx:  # 예: 조식명단은 객실 칸에 '합계'
                label_col = colmap[idx[sum_label_field]]
        if sections:
            # 수입일보: 묶음(ROOM·Extra-Charge·조식·연회 …)마다 계 줄, 계는 앞 묶음까지 이어서 더한 값 (호텔 실제 일보)
            prev = None
            for a, b, fr in sections:
                for c in printable + s.extra_cols:  # 계 줄 칸의 데이터 서식(노란 칸 등)을 지우고 계 서식으로
                    ws.cell(fr, c).fill = copy(_NO_FILL)
                    ws.cell(fr, c).value = None
                _foot_row(fr)
                if label_col:
                    ws.cell(fr, label_col).value = sum_label
                for i in sum_idx:
                    L = letter(i)
                    _foot(fr, i, f"={L}{prev}+SUM({L}{a}:{L}{b})" if prev else f"=SUM({L}{a}:{L}{b})")
                prev = fr
            # 계 줄 사이 데이터 행에 걸린 입력 도움말 범위는 그대로 둬도 계 줄은 비어 있어 문제 없음
        else:
            _foot_row(r)
            if label_col:
                ws.cell(r, label_col).value = sum_label
            for i in sum_idx:
                L = letter(i)
                _foot(r, i, f"=SUM({L}{first}:{L}{last})")

    drop = {norm(x) for x in drop_trailer}
    for cells, height, merges in s.trailer:  # 양식의 표 아래 행(외부식사·안내 문구)을 그대로 붙인다
        if cells and norm(cells[0][1]) in drop:  # 예: 조식명단 '외부식사' 행은 합계 행으로 대체
            continue
        r += 1
        for c, v, st in cells:
            cell = ws.cell(r, c)
            cell.value = v
            _apply(cell, st)
        if height:
            ws.row_dimensions[r].height = height
        for c0, c1, extra in merges:
            ws.merge_cells(start_row=r, start_column=c0, end_row=r + extra, end_column=c1)

    table_end = r  # 여기까지가 표(제목~합계·안내 행). 아래 메모는 표 밖

    if notes:
        r += 2
        for n in notes:
            ws.cell(r, printable[0], n).font = copy(s.note_font)
            r += 1

    right = max(printable[-1], s.frame_last)
    _draw_grid(ws, 1, table_end, 1, right)
    ws.print_area = f"A1:{get_column_letter(right)}{r}"
    _fit_a4(ws, hrow, keep_margins=bool(template) and not fill_pages)
    if fit_pages:
        _fit_pages(ws, hrow, first, table_end, right, fit_pages, fill=bool(fill_pages),
                   min_ratio=fill_pages if isinstance(fill_pages, float) else 0.75, width_px=width_px)
    if ws.auto_filter.ref and records:
        ws.auto_filter.ref = f"A{hrow}:{get_column_letter(right)}{last}"
    if manual or (cash is not None and cash.get("amount") is None):
        # 노란 칸은 화면에서만 보이는 안내 → 인쇄는 흑백으로 해서 호텔 양식처럼 나오게
        ws.page_setup.blackAndWhite = True

    try:
        wb_out = _clean_copy(s.ws)
        if card_lookup:
            _add_card_lookup_sheet(wb_out, card_lookup)
        wb_out.save(path)
        _mark_styles_applied(path)
    except PermissionError:
        raise ConversionError(f"결과 파일이 엑셀에서 열려 있습니다. 닫은 뒤 다시 시도해 주세요.\n{path}")
    return s.warnings


def _mark_styles_applied(path):
    """셀 서식마다 '테두리·채우기·글꼴·표시 형식을 적용함' 표시(applyBorder 등)를 붙인다.
    openpyxl은 이 표시를 쓰지 않는데, 호텔 PC의 프로그램은 표시가 없으면 테두리를 그리지 않는다.
    사용자 확인(2026-10-23): 이 표시를 넣은 빌드 12·13에서는 선이 보였고, 뺀 빌드 16에서는 안 보였음 → 반드시 유지."""
    import re
    import shutil
    import tempfile
    import zipfile
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=os.path.dirname(os.path.abspath(path)))
    tmp.close()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                xml = data.decode("utf-8")
                m = re.search(r"<cellXfs.*?</cellXfs>", xml, re.S)
                if m:
                    def flag(x):
                        tag = x.group(0)
                        for attr in ("applyNumberFormat", "applyFont", "applyFill", "applyBorder"):
                            if attr not in tag:
                                tag = tag.replace("<xf ", f'<xf {attr}="1" ', 1)
                        return tag
                    xfs = re.sub(r"<xf [^>]*>", flag, m.group(0))
                    xml = xml[:m.start()] + xfs + xml[m.end():]
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp.name, path)


def _clean_copy(ws):
    """완성된 시트를 새 통합 문서로 옮겨 담는다 (값·글꼴·크기·테두리·정렬·표시 형식·열 너비·행 높이·병합·인쇄 설정만).
    호텔 .xls에서 넘어온 색상표·흰색 채우기 같은 군더더기를 모두 버리고, 엑셀이 가장 확실히 읽는 표준 서식만 남긴다
    (2026-10-23 '엑셀에서 표 선이 안 보인다' 지적 3회)."""
    wb = Workbook()
    out = wb.active
    out.title = ws.title

    def rgb(color):
        v = getattr(color, "rgb", None)
        return ("FF" + v[-6:]) if isinstance(v, str) and len(v) >= 6 else None

    def side(sd):
        return Side(style=sd.style, color=_LINE) if sd is not None and sd.style else Side()

    # 병합을 먼저 해 두어야 병합 칸 가장자리 선이 뒤에서 덮어쓰이지 않는다
    for mr in ws.merged_cells.ranges:
        out.merge_cells(str(mr))
    for row in ws.iter_rows():
        for c in row:
            if c.value is None and not c.has_style:
                continue
            d = out.cell(c.row, c.column)
            if not isinstance(d, MergedCell):
                d.value = c.value
            f = c.font
            d.font = Font(name=f.name or "굴림", size=f.sz or 11, bold=bool(f.b), italic=bool(f.i), color=rgb(f.color))
            b = c.border
            d.border = Border(left=side(b.left), right=side(b.right), top=side(b.top), bottom=side(b.bottom))
            a = c.alignment
            d.alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, wrap_text=bool(a.wrap_text),
                                    shrink_to_fit=bool(a.shrink_to_fit))
            if _is_blank_fill(c):
                d.fill = PatternFill("solid", fgColor="FFFFF2B3", bgColor="FFFFF2B3")
            d.number_format = c.number_format or "General"
    for key, dim in ws.column_dimensions.items():
        if dim.width:
            out.column_dimensions[key].width = dim.width
    for r, dim in ws.row_dimensions.items():
        if dim.height:
            out.row_dimensions[r].height = dim.height
    for dv in ws.data_validations.dataValidation:
        nd = DataValidation(type=dv.type, operator=dv.operator, formula1=dv.formula1, allow_blank=dv.allow_blank,
                            showErrorMessage=dv.showErrorMessage, errorTitle=dv.errorTitle, error=dv.error,
                            showInputMessage=dv.showInputMessage, promptTitle=dv.promptTitle, prompt=dv.prompt)
        out.add_data_validation(nd)
        nd.add(str(dv.sqref))
    if ws.auto_filter.ref:
        out.auto_filter.ref = ws.auto_filter.ref
    out.freeze_panes = ws.freeze_panes
    if ws.print_area:
        out.print_area = ws.print_area[0] if isinstance(ws.print_area, (list, tuple)) else ws.print_area.split("!")[-1]
    if ws.print_title_rows:
        out.print_title_rows = ws.print_title_rows.replace("$", "")
    ps, src = out.page_setup, ws.page_setup
    ps.orientation, ps.paperSize = src.orientation, src.paperSize
    ps.fitToWidth, ps.fitToHeight = src.fitToWidth, src.fitToHeight
    ps.blackAndWhite = src.blackAndWhite
    fit = ws.sheet_properties.pageSetUpPr.fitToPage if ws.sheet_properties.pageSetUpPr else True
    out.sheet_properties.pageSetUpPr.fitToPage = True if fit is None else bool(fit)
    if src.scale:
        ps.scale = src.scale
    for brk in ws.row_breaks.brk:  # 조식명단 2장: 정확히 반으로 나누는 수동 쪽 나눔
        out.row_breaks.append(Break(id=brk.id))
    for k in ("left", "right", "top", "bottom", "header", "footer"):
        setattr(out.page_margins, k, getattr(ws.page_margins, k))
    out.print_options.horizontalCentered = ws.print_options.horizontalCentered
    out.oddFooter.center.text = ws.oddFooter.center.text
    if ws.sheet_view.zoomScale:
        out.sheet_view.zoomScale = ws.sheet_view.zoomScale
    return wb


# ---------------------------------------------------------------- 공개 API
def unique_path(path):
    base, ext = os.path.splitext(path)
    n, p = 2, path
    while os.path.exists(p):
        p = f"{base}({n}){ext}"
        n += 1
    return p


def _parse_date(s):
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        raise ConversionError("영업일자를 YYYY-MM-DD 형식으로 입력해 주세요. 예) 2026-10-23")


# 근무자가 ERP 비고(또는 고객명)에 적는 단체 표시 '단체이름/방 갯수/인원수' (사용자 지시 2026-11-03)
# 예: '가상협회/17/30', '가상협회 / 17실 / 30명'. 날짜(2026/11/01)와 헷갈리지 않게 이름은 글자로 시작·끝나야 함
_GROUP_NOTE = re.compile(r"([가-힣A-Za-z(][^/\n\r]*?[가-힣A-Za-z)])\s*/\s*(\d{1,3})\s*(?:실|개|룸|객실|방)?\s*/\s*"
                         r"(\d{1,4})\s*(?:명|인)?(?![\d/~.\-])")


def group_note(*texts):
    """'단체이름/방 갯수/인원수' 표시를 찾으면 (이름, 방 수, 인원), 없으면 None"""
    for t in texts:
        for m in _GROUP_NOTE.finditer(str(t or "")):
            rooms, pax = int(m.group(2)), int(m.group(3))
            if 1 <= rooms <= 200 and rooms <= pax <= 1000:
                return m.group(1).strip(), rooms, pax
    return None


def _bf_people(remark, conv, guest=None):
    """비고 문구로 조식 대상 여부와 조식 인원을 판단. (대상여부, 인원 또는 None, 판단 이유)
    '조식불포함'·'조식X'·'룸온리'처럼 조식이 없는 표현을 먼저 걸러 낸다 (2026-10-26: 실제 47명인데 59명으로 집계된 문제)."""
    t = remark or ""
    if re.search(conv["bf_exclude_regex"], t, re.I):
        return False, None, "룸온리·조식 불포함"
    g = group_note(t, guest)
    if g:  # '단체이름/방 갯수/인원수' → 단체 전체 조식 인원 (객실별 인원은 호출한 쪽에서, 모자라면 '단체 추가' 줄)
        return True, None, f"조식 포함(단체 {g[2]}명, {g[1]}실, {g[0]})"
    if re.search(conv["bf_include_regex"], t, re.I):
        # 인원: '1인 조포', '2인 조식 포함', '조식 2인', '조식2명' → 조식 바로 앞뒤의 숫자만 쓴다.
        # 비고의 다른 'N인'은 세지 않는다: '2인 1실 조식 추가 비용 13,000원'은 추가요금 안내이지 조식 인원이 아님
        # (2026-10-26 실제 IN HOUSE LIST: 이 문구가 있는 16객실을 2명으로 세어 59명 → 규칙 수정 후 실제와 같은 47명)
        # 숫자가 없으면 호출한 쪽에서 성인+소인 수를 쓴다
        m = (re.search(r"(\d+)\s*(?:인|명)\s*(?:조식|조포|b\s*/?\s*f)", t, re.I)
             or re.search(r"(?:조식|조포|b\s*/?\s*f)[^\d\n]{0,6}(\d+)\s*(?:인|명)", t, re.I))
        n = int(m.group(1)) if m else None
        if n is not None and n > conv.get("max_pax_per_room", 4):
            # '조식 포함 (35명)'처럼 단체 전체 인원 → 객실 인원이 아님 (2026-10-29: 18실을 35명씩 세어 632명). 호출한 쪽에서 객실별로 센다
            return True, None, f"조식 포함(단체 {n}명)"
        return True, n, "조식 포함"
    return False, None, "조식 표시 없음"


def _names_count(guest):
    """투숙객명 '홍길동/김철'처럼 여러 이름이 적혀 있으면 이름 수 (1명이면 None)"""
    parts = [p for p in re.split(r"[/,，·]", guest or "") if p.strip()]
    return len(parts) if len(parts) > 1 else None


def _pax_fallback(rec):
    """비고에 객실 인원이 없을 때: 투숙객명에 적힌 이름 수 → 없으면 성인+소인. (인원, 추정 여부)"""
    k = _names_count(rec.get("guest"))
    a, c = rec.get("adults"), rec.get("children")
    ac = (a or 0) + (c or 0) if (a is not None or c is not None) else None
    if k and (ac is None or k > ac):
        return k, True
    return ac, ac is not None


def _pax_sanity(records, conv, warnings):
    """안전장치 (2026-10-29 '객실마다 35명' 재발 방지): 한 객실 인원이 객실 최대 인원보다 많으면
    맨 위에 '확인 필요' 경고를 띄운다. 어떤 경로로 인원을 셌든 마지막에 한 번 더 검사한다."""
    cap = conv.get("max_pax_per_room", 4)
    big = [f"{r.get('room') or '미배정'}({r.get('pax')}명)" for r in records
           if isinstance(r.get("pax"), (int, float)) and r["pax"] > cap and not r.get("_extra")]
    if big:
        warnings.insert(0, f"⚠ 확인 필요: 한 객실 인원이 {cap}명을 넘는 곳이 있습니다 — " + ", ".join(big)
                        + " (단체 전체 인원이 객실 인원으로 들어갔을 수 있습니다)")
    return not big


def _check_group_totals(records, warnings):
    """단체 비고의 전체 인원(N명)과 객실별로 센 인원 합을 비교해 작업 기록에 남긴다.
    객실별로 센 합이 비고 인원보다 적으면(2026-11-03 실제: 17실에 '30명 조식 포함', 객실마다 1명으로 셈)
    모자란 인원을 '단체 추가 인원' 한 줄로 돌려준다 → 명단 합계 = 비고 인원 (어느 방이 몇 명인지는 지어내지 않음)."""
    groups = {}
    for rec in records:
        m = re.search(r"단체 (\d+)명", rec.get("_why") or "")
        if m and rec.get("_bf", True):
            g = groups.setdefault((int(m.group(1)), rec.get("_why")), [0, 0, {}])
            g[0] += 1
            g[1] += rec.get("pax") or 0
            name = (rec.get("guest") or "").strip()
            g[2][name] = g[2].get(name, 0) + 1
    extras = []
    for (n, why), (rooms, pax, names) in groups.items():
        same = "일치" if pax == n else f"다름 — 확인 필요"
        name = max(names, key=names.get) if names else "단체"
        stated = re.search(r"(\d+)실, (.+)\)$", why or "")
        if stated:  # 근무자가 적은 '단체이름/방 갯수/인원수'
            name = stated.group(2)
            if int(stated.group(1)) != rooms:
                warnings.append(f"⚠ 단체 '{name}': 적힌 방 갯수 {stated.group(1)}실과 명단의 객실 {rooms}실이 다릅니다 — 확인해 주세요")
        if pax < n:
            extras.append({"room": "단체", "guest": f"{name} 추가", "pax": n - pax,
                           "_extra": True, "_bf": True, "_guess": False, "_why": f"단체 비고 {n}명 - 객실별 {pax}명"})
            same = f"모자란 {n - pax}명은 '단체 추가 인원' 줄로 넣어 합계 {n}명"
        warnings.append(f"단체 비고 '조식 포함 ({n}명)': {rooms}객실을 투숙객명·성인 수로 세어 {pax}명 ({same})")
    return extras


_HANGUL_NAME = re.compile(r"[가-힣]{2,4}")
_ERP_MASKED = re.compile(r"[가-힣][*＊]+[가-힣]?|[가-힣]{2}[*＊]+[가-힣]?")
# 4글자 한국 이름은 드묾 → 두 글자 성(남궁·황보 …)으로 시작할 때만 이름으로 본다 (2026-10-27: '씨티소재'가 '씨00재'로 가려짐)
_TWO_CHAR_SURNAMES = ("남궁", "황보", "제갈", "선우", "독고", "사공", "서문", "동방", "어금", "망절", "강전", "소봉")
# 2·3글자라도 회사·단체를 뜻하는 말이 들어 있으면 이름이 아님
_COMPANY_WORDS = re.compile(r"소재|그룹|상사|산업|건설|전자|물산|화학|제약|은행|증권|보험|여행|투어|항공|호텔|학교|대학|"
                            r"교회|협회|공사|회사|기업|재단|센터|병원|통상|무역|기획|개발|디자인|엔지|테크|코리아|"
                            r"유통|식품|컴퍼니|주식|유한|단체|팀$|시청|구청|군청|경찰|소방|부대")


def is_korean_person_name(name):
    """한국 사람 이름으로 볼 만한지: 3글자(대부분) · 2글자(외자) · 4글자는 두 글자 성일 때만. 회사 이름은 제외."""
    s = (name or "").strip()
    if not _HANGUL_NAME.fullmatch(s) or _COMPANY_WORDS.search(s):
        return False
    return len(s) < 4 or s.startswith(_TWO_CHAR_SURNAMES)


# 이름 앞에 붙은 근무조·직책 같은 꼬리표 (이름으로 보지 않음)
_LABEL_WORDS = {"주간", "야간", "오전", "오후", "새벽", "직원", "사원", "팀장", "과장", "부장", "차장", "대리", "실장", "기사",
                "인솔", "가이드", "장기", "단기", "객실", "조식", "본사", "현장", "외부"}


def korean_person_part(name):
    """'가상회사 주간-홍길동', '야간_김철', '(주간)홍길동'처럼 회사·근무조 꼬리표 뒤에 한국 이름이 붙어 있으면 그 이름만.
    사용자 지시 2026-11-03: 장기 투숙 회사 직원 이름이 '주간-000'으로 들어감 → 순수한 한국 이름만 조식명단에.
    마지막 낱말이 한국 이름(ERP가 *로 가린 이름 포함)일 때만, 아니면 None (진양그룹-Noah는 그대로)."""
    parts = [p for p in re.split(r"[\s\-_:()\[\]·.~]+", (name or "").strip()) if p]
    if len(parts) < 2:
        return None
    last = parts[-1]
    # 바로 앞 낱말이 근무조·직책 꼬리표일 때만 (단체명 뒤 '(경기도)' 같은 지역 이름은 이름이 아님, 2026-11-03)
    if last in _LABEL_WORDS or parts[-2] not in _LABEL_WORDS:
        return None
    if is_korean_person_name(last) or _ERP_MASKED.fullmatch(last):
        return last
    return None


def mask_korean_name(name, mark="0"):
    """개인정보 보호 (사용자 지시 2026-10-26): 한국인 이름만 가운데 글자를 0으로.
    3자 홍길동→홍0동, 외자 2자 김철→김0, 4자는 두 글자 성만 남궁민수→남00수.
    회사·단체 이름(씨티소재, 삼성전자), 외국인 이름·영문 섞인 이름(진양그룹-Noah)·5자 이상은 그대로 (2026-10-27).
    ERP가 이미 *로 가린 이름(정*영)은 *를 0으로 바꿔 표기를 통일."""
    s = (name or "").strip()
    if re.search(r"[/,，]", s):  # '홍길동/김철'처럼 여러 사람 → 각각 가린다 (2026-10-29)
        return re.sub(r"[^/,，]+", lambda m: m.group(0).replace(m.group(0).strip(), mask_korean_name(m.group(0).strip(), mark))
                      if m.group(0).strip() else m.group(0), s)
    if _ERP_MASKED.fullmatch(s):
        return re.sub(r"[*＊]", mark, s)
    if not is_korean_person_name(s):
        person = korean_person_part(s)
        if not person:
            return name
        s = person  # '가상회사 주간-홍길동' → '홍길동'만 (사용자 지시 2026-11-03)
    if _ERP_MASKED.fullmatch(s):
        return re.sub(r"[*＊]", mark, s)
    if len(s) == 2:
        return s[0] + mark
    return s[0] + mark * (len(s) - 2) + s[-1]


def _nights(arrival, departure):
    """숙박수(박) = 출발일 - 도착일. 날짜가 없거나 출발일이 도착일보다 빠르면 None"""
    try:
        a = dt.datetime.strptime(str(arrival)[:10], "%Y-%m-%d").date()
        d = dt.datetime.strptime(str(departure)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    n = (d - a).days
    return n if n >= 0 else None


def _nights_so_far(arrival, departure, bdate):
    """수입일보 숙박수 = 도착일부터 영업일자까지 묵은 날 수 (10/26 도착 → 10/28 일보에 3).
    사용자 지시 2026-11-01: 누계는 '이번 달 1일부터'가 아니라 **지금까지의 잔액** → 중간 정산한 장기 투숙은
    객실별 거래내역 파일(read_folio)로 잔액을 구하고, 파일이 없을 때만 이 값을 쓴다. 날짜를 모르면 None."""
    try:
        a = dt.date.fromisoformat(arrival)
    except (TypeError, ValueError):
        return None
    if a > bdate:
        return None
    start = a
    n = (bdate - start).days + 1
    try:  # 오늘 출발 손님이 일보에 남아 있으면 출발일 전날까지만
        d = dt.date.fromisoformat(departure)
        if d <= bdate:
            n = max(0, (d - start).days)
    except (TypeError, ValueError):
        pass
    return n


_FOLIO_SIGNATURE = ("영업일자", "room", "항목", "매출", "입금")


def read_folio(path):
    """ERP 객실별 거래내역(영업일자 · Bill# · Room# · 구분# · 항목 · 매출 · 입금 …) → (객실, [(날짜, 매출, 입금)]).
    거래내역 파일이 아니면 None"""
    try:
        rows = read_table(path)
    except Exception:
        return None
    for i, row in enumerate(rows[:10]):
        heads = [norm(c) for c in row]
        if all(any(h == k or h.startswith(k) for h in heads) for k in _FOLIO_SIGNATURE):
            col = lambda k: next(j for j, h in enumerate(heads) if h == k or h.startswith(k))
            cd, cr, cs, cp = col("영업일자"), col("room"), col("매출"), col("입금")
            room, out = None, []
            for r in rows[i + 1:]:
                d = to_date(r[cd]) if cd < len(r) else ""
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d or ""):
                    continue  # 맨 아래 합계 줄 등
                room = room or fmt_room(r[cr], "pad4")
                out.append((d, to_number(r[cs]) or 0, to_number(r[cp]) or 0))
            return (room, out) if room else None
    return None


def folio_balance(entries, bdate):
    """누계 = 어제까지의 잔액 + 오늘 매출 (사용자 지시 2026-11-01, 실제 10/28 일보로 검산:
    0406 3박 330,000 + 110,000 = 440,000 / 0602 8~9월 정산 뒤 10/1부터 27박 + 오늘 = 2,772,000 / 0604 매일 빌이체로 잔액 0 → 오늘 90,000).
    반환: (어제까지 잔액, 잔액이 0이 된 뒤 묵은 날 수 + 오늘)"""
    today = bdate.isoformat()
    by_day = {}
    for d, sale, pay in entries:
        if d < today:
            by_day[d] = by_day.get(d, 0) + sale - pay
    bal, nights = 0, 0
    for d in sorted(by_day):
        bal += by_day[d]
        nights = 0 if bal <= 0 else nights + 1  # 정산해서 0이 되면 다시 센다
    return max(bal, 0), nights + 1


def folio_reset_date(entries, bdate):
    """어제까지 중 잔액이 마지막으로 0(정산)이 된 날 — 없으면 None"""
    today = bdate.isoformat()
    by_day = {}
    for d, sale, pay in entries:
        if d < today:
            by_day[d] = by_day.get(d, 0) + sale - pay
    bal, last = 0, None
    for d in sorted(by_day):
        bal += by_day[d]
        if bal <= 0:
            last = d
    return last


def _report_rows(path, md):
    """수입일보 파일(프로그램 결과 .xlsx 또는 호텔 수작업 .xls/.xlsx)의 행 목록 [[(값, 수식)…]].
    시트가 여러 개면 제목에 MM/DD가 있는 시트 → 선택된 시트 → 첫 시트"""
    with open(path, "rb") as f:
        head = f.read(4)
    if head[:2] == b"PK":
        wb_v = load_workbook(path, data_only=True)
        wb_f = load_workbook(path)
        act = wb_v.active.title
        pick = act if md in str(wb_v[act]["A1"].value or "") else \
            next((n for n in reversed(wb_v.sheetnames) if md in str(wb_v[n]["A1"].value or "")), None)
        name = pick or act
        wv, wf = wb_v[name], wb_f[name]
        return [[(wv.cell(r, c).value, wf.cell(r, c).value) for c in range(1, wv.max_column + 1)]
                for r in range(1, wv.max_row + 1)]
    import xlrd
    book = xlrd.open_workbook(path)
    sheets = book.sheets()
    # 선택된(마지막으로 작업한) 시트가 그 날짜면 그것 — 같은 이름 시트가 여러 개일 수 있음 (2026-11-02 실제: '1101'과 '1101 ')
    sel = [x for x in sheets if getattr(x, "sheet_selected", 0)]
    sh = next((x for x in sel if x.nrows and md in str(x.cell_value(0, 0))), None) \
        or next((x for x in reversed(sheets) if x.nrows and md in str(x.cell_value(0, 0))), None) \
        or (sel[0] if sel else sheets[0])
    return [[(sh.cell_value(r, c), sh.cell_value(r, c)) for c in range(sh.ncols)] for r in range(sh.nrows)]


def read_prev_report(path, prev_date):
    """어제 수입일보에서 ROOM 묶음(첫 '계' 줄 위)의 객실별 (도착일, 숙박수, 매출, 누계)를 읽는다 (사용자 지시 2026-11-01:
    객실마다 거래내역을 받기 번거로움 → 어제 일보의 누계를 이어받는다). 근무자가 엑셀에서 고친 숙박수·누계도 반영. 못 읽으면 {}"""
    try:
        rows = _report_rows(path, prev_date.strftime("%m/%d"))
    except Exception:
        return {}
    hi = next((i for i, r in enumerate(rows[:10]) if {"객실", "숙박수", "누계"} <= {norm(v) for v, _ in r}), None)
    if hi is None:
        return {}
    h = [norm(v) for v, _ in rows[hi]]
    c_room, c_arr, c_n, c_s, c_cum, c_pp = (h.index(k) if k in h else None
                                            for k in ("객실", "도착일", "숙박수", "매출", "누계", norm("선/후")))
    if None in (c_room, c_n, c_s, c_cum):
        return {}
    from openpyxl.utils import column_index_from_string
    out = {}
    for r in rows[hi + 1:]:
        if any(norm(v) in _TOTAL_WORDS for v, _ in r[:8]):
            break  # ROOM 묶음 끝 (첫 계 줄)
        room = fmt_room(r[c_room][0], "pad4") if r[c_room][0] not in (None, "") else ""
        if not re.fullmatch(r"\d{4}", room):
            continue
        sales, n = to_number(r[c_s][0]), to_number(r[c_n][0])
        val, formula = r[c_cum]
        cum = to_number(val) if not (isinstance(val, str) and val.startswith("=")) else None
        if cum is None and isinstance(formula, str):  # 엑셀이 계산값을 저장하지 않은 프로그램 결과 파일
            row_of = lambda ref: r[column_index_from_string(re.match(r"[A-Z]+", ref).group()) - 1][0]
            m = re.fullmatch(r"=([A-Z]+\d+)\*([A-Z]+\d+)", formula)
            m2 = re.fullmatch(r"=(-?\d+(?:\.\d+)?)\+([A-Z]+\d+)", formula)
            m3 = re.fullmatch(r"=([A-Z]+\d+)", formula)
            if m:
                a, b = to_number(row_of(m.group(1))), to_number(row_of(m.group(2)))
                cum = a * b if a is not None and b is not None else None
            elif m2:
                cum = float(m2.group(1)) + (to_number(row_of(m2.group(2))) or 0)
            elif m3:
                cum = to_number(row_of(m3.group(1)))
        if cum is None or n is None:
            continue
        arr = to_date(r[c_arr][0]) if c_arr is not None else ""
        pp = clean_text(r[c_pp][0]) if c_pp is not None else ""
        out[room] = (arr, int(n), sales or 0, int(cum) if float(cum).is_integer() else cum, pp)
    return out


def _income_sections(records, bdate, conv, warnings, folios=None, partners=None, prev_rep=None):
    """수입일보를 호텔 실제 일보 모양으로 (사용자 지시 2026-10-28, 실제 10/28 일보로 확인).
    - 매출구분(ROOM · Extra-Charge · 1회용품 · 조식 · 연회 …) 묶음을 ERP 순서대로, 묶음마다 계 줄(앞 묶음까지 이어서 합산)
    - 매출 0 줄(CASH·카드·선수금 등)은 뺀다. 객실번호 범위와 상관없이 매출이 잡힌 줄은 모두 (1030 외부 조식 포함)
    - ROOM 줄: 객실마다 한 줄(같은 객실이 또 있으면 금액 합산), 숙박수 = 이번 달 묵은 날 수, 누계 = 매출 × 숙박수
    - 다른 묶음: 투숙객명 옆에 (EXTRA CHARGE)/(조식) 같은 표시, 숙박수 빈칸, 누계 = 매출"""
    room_kind = norm(conv.get("room_kind", "ROOM"))
    labels = {norm(k): v for k, v in conv.get("kind_labels", {}).items()}
    # 선/후 미리 채우기 (2026-11-02): 오늘 결제 줄(매출 0, 입금 있음)의 구분 → 카드 '카.선', CASH '현.선', 대외후불 '온라인결제'
    pay_map = {norm(k): v for k, v in (conv.get("prepost_by_payment") or {}).items()}
    paid = {}  # 객실 → [(구분, 결제 금액)]
    for rec in records:
        lab = pay_map.get(norm(rec.get("sales_kind")))
        if lab and rec.get("room") and (rec.get("payment") or 0) > 0 and not (rec.get("sales") or 0):
            paid.setdefault(rec["room"], []).append((lab, rec["payment"]))

    def paid_label(rec):
        """그 줄 매출 이상을 결제한 건이 있을 때만 선결제로 본다 (조식 13,000원 카드 결제로 객실까지 '카.선'이 되지 않게)"""
        return next((lab for lab, amt in paid.get(rec.get("room"), ()) if amt >= (rec.get("sales") or 0)), None)
    zero = [r for r in records if not (r.get("sales") or 0)]
    order, groups = [], {}
    for rec in records:
        if not (rec.get("sales") or 0):
            continue
        kind = rec.get("sales_kind") or "기타"
        if norm(kind) not in groups:
            order.append(norm(kind))
            groups[norm(kind)] = []
        groups[norm(kind)].append(rec)
    if room_kind in groups:  # ROOM 묶음을 맨 앞에
        order.remove(room_kind)
        order.insert(0, room_kind)
    out, merged, bad, used, long_no_folio, stale, odd, carried = [], [], [], [], [], [], [], []
    for sec, key in enumerate(order):
        rows = groups[key]
        if key == room_kind:
            by_room = {}
            for rec in rows:
                k = rec.get("room") or id(rec)
                if k in by_room:
                    by_room[k]["sales"] += rec["sales"]
                    merged.append(rec.get("room"))
                else:
                    by_room[k] = dict(rec)
            rows = list(by_room.values())
            for rec in rows:
                rec["nights"] = _nights_so_far(rec.get("arrival"), rec.get("departure"), bdate)
                if rec.get("room") in (folios or {}):
                    prev, n = folio_balance(folios[rec["room"]], bdate)
                    last = max((d for d, _, _ in folios[rec["room"]]), default="")
                    if last and last < (bdate - dt.timedelta(days=1)).isoformat():
                        stale.append(f"{rec['room']}(마지막 {last[5:]})")
                    rate = rec.get("sales") or 0
                    if rate and prev % rate == 0:
                        # 잔액이 매출의 정수배(일부 박만 정산·입금한 경우 포함) → 숙박수 = 남은 잔액의 박 수, 누계 = 매출×숙박수
                        n = int(prev // rate) + 1
                    else:
                        rec["_cum_prev"] = prev  # 요금 변경·부분 입금 → 누계 = 어제까지 잔액 + 오늘 매출 (숙박수는 정산 뒤 묵은 날)
                        odd.append(f"{rec['room']}(잔액 {prev + rate:,.0f}원, {n}박)")
                    rec["nights"] = n
                    used.append(f"{rec['room']}({prev + (rec.get('sales') or 0):,.0f}원)")
                elif rec.get("room") in (prev_rep or {}) and (not prev_rep[rec["room"]][0]
                                                               or prev_rep[rec["room"]][0] == rec.get("arrival")):
                    # 어제 일보 이어받기: 누계 = 어제 누계 + 오늘 매출, 숙박수 = 어제 + 1 (같은 손님 = 도착일이 같을 때만)
                    _, pn, _, pcum, pp = prev_rep[rec["room"]]
                    if pp:
                        rec["prepost"] = pp  # 계속 묵는 손님은 선/후도 어제 일보 그대로 (오늘 결제보다 우선)
                    rate = rec.get("sales") or 0
                    rec["nights"] = pn + 1
                    if pcum != rate * pn:
                        rec["_cum_prev"] = pcum  # 요금이 바뀌었거나 근무자가 누계를 고친 경우
                    carried.append(rec["room"])
                elif rec.get("room") in (partners or {}) and rec["nights"] is not None:
                    # 같은 회사 동료 객실: ERP가 대표 객실 거래내역을 주므로, 대표 객실이 정산한 다음 날부터 센다 (사용자 지시 2026-11-01)
                    rep_room, reset = partners[rec["room"]]
                    if reset:
                        start = max(dt.date.fromisoformat(rec["arrival"]),
                                    dt.date.fromisoformat(reset) + dt.timedelta(days=1))
                        rec["nights"] = max(1, (bdate - start).days + 1)
                    used.append(f"{rec['room']}(동료 {rep_room}호 정산 기준 {rec['nights']}박)")
                elif rec["nights"] is None:
                    bad.append(rec.get("room") or "객실 없음")
                elif rec["nights"] > conv.get("long_stay_days", 7):
                    long_no_folio.append(f"{rec['room']}({rec['nights']}박)")
        else:
            rows = [dict(r) for r in rows]
            for rec in rows:
                rec["nights"] = None
                rec["_no_fill"] = True
                if conv.get("extra_kind_in_company"):
                    # 11/01 수작업 일보(사용자가 정한 표준, 2026-11-02): 이름은 그대로, 거래처 칸에 매출구분(Extra-Charge·조식·연회)
                    rec["company"] = rec.get("sales_kind") or rec.get("company")
                    continue
                lab = labels.get(key, rec.get("sales_kind"))
                base = rec.get("guest") or ""
                if norm(lab) not in norm(base):  # '[[외부 조식]] 가상객실'처럼 이미 구분이 들어 있으면 그대로
                    if len(base) + len(lab) > 14:  # 실제 일보처럼 '가상건설협회 금속지붕' → '가상건설(EXTRA CHARGE)'
                        base = base[:4]
                    rec["guest"] = f"{base}({lab})"
                rec["nights"] = None
                rec["_no_fill"] = True
        for rec in rows:
            if not rec.get("prepost"):
                rec["prepost"] = paid_label(rec)
        rows.sort(key=_room_key)
        for rec in rows:
            rec["_sec"] = sec
        out += rows
    if zero:
        warnings.append(f"매출이 0인 줄 {len(zero)}건(CASH·카드·선수금 등)은 뺐습니다.")
    if merged:
        warnings.append("ROOM 매출이 두 줄 이상인 객실은 금액을 더해 한 줄로 넣었습니다: " + ", ".join(sorted(set(merged))))
    if bad:
        warnings.append("도착일을 알 수 없어 숙박수를 비워 둔 객실(노란 칸): " + ", ".join(bad))
    if used:
        warnings.append("객실별 거래내역 파일로 누계(현재 잔액)를 계산한 객실: " + ", ".join(used))
    if folios:
        extra = sorted(set(folios) - {r.get("room") for r in out})
        if extra:
            warnings.append("거래내역 파일은 있지만 오늘 수입일보에 없는 객실(퇴실 등): " + ", ".join(extra))
    if carried:
        warnings.append(f"어제 일보에서 누계를 이어받은 객실 {len(carried)}개 (누계 = 어제 누계 + 오늘 매출, 숙박수 = 어제 + 1)")
    if odd:
        warnings.append("잔액이 매출로 나누어떨어지지 않는 객실(요금 변경·부분 입금) — 누계는 거래내역 잔액 그대로, "
                        "숙박수는 마지막 정산 뒤 묵은 날입니다. 확인해 주세요: " + ", ".join(odd))
    if stale:
        warnings.append("⚠ 거래내역 파일이 오래되어(어제까지 내역이 없음) 누계가 모자랄 수 있습니다 — 오늘 다시 받아 주세요: " + ", ".join(stale))
    if long_no_folio:
        warnings.append("⚠ 장기 투숙인데 객실별 거래내역 파일이 없어 '도착일부터 매일 매출'로 누계를 셌습니다 — "
                        "중간 정산했으면 틀립니다. 거래내역을 받아 함께 넣어 주세요: " + ", ".join(long_no_folio))
    if len(order) > 1:
        names = [groups[k][0].get("sales_kind") for k in order]
        warnings.append("매출구분별로 묶고 묶음마다 계(앞 묶음까지 합산)를 넣었습니다: " + " → ".join(names))
    return out


def _group_key(name):
    """고객명이 완전히 같은 객실끼리만 단체로 본다.
    '진양그룹-Noah'처럼 단체명 옆에 개인 이름이 있으면 각자 한 줄 (사용자 지시 2026-10-23)"""
    return " ".join((name or "").split())


def _group_breakfast(records, conv):
    """같은 단체(고객명이 똑같은 객실이 group_min_rooms개 이상)를 한 줄로 합친다.
    반환: (개별 객실 + 단체 줄, [(단체명, 객실 수, 인원 합)])"""
    min_rooms = conv.get("group_min_rooms", 2)
    by_key = {}
    for rec in records:
        k = _group_key(rec.get("guest"))
        if k:
            by_key.setdefault(k, []).append(rec)
    singles, rows, groups = [], [], []
    for rec in records:
        k = _group_key(rec.get("guest"))
        if not k or len(by_key[k]) < min_rooms:
            singles.append(rec)
    for k, recs in by_key.items():
        if len(recs) < min_rooms:
            continue
        paxes = [r.get("pax") for r in recs]
        pax = sum(paxes) if all(p is not None for p in paxes) else None
        rows.append({"room": conv.get("group_room_label", "단체"), "guest": f"{k} ({len(recs)}실)", "pax": pax,
                     "_group": True})
        groups.append((k, len(recs), pax))
    return singles + rows, groups


def _room_lookup(path, conv):
    """객실 예약 리스트(상태 CHECK IN = 재실)에서 Acct# → [(객실, 일일요금×일수, 투숙객명)] (카드 객실 자동 채우기, 2026-10-28)"""
    rcfg = conv["room_source"]
    records, _, mapped = _extract({"source_fields": rcfg["source_fields"], "required": ["acct", "room"],
                                   "types": {"daily": "number", "nights": "number"}, "room_format": conv.get("room_format")},
                                  read_table(path))
    out = {}
    for rec in records:
        if not rec.get("acct") or not rec.get("room"):
            continue
        total = (rec.get("daily") or 0) * (rec.get("nights") or 0)
        out.setdefault(rec["acct"], []).append((rec["room"], total, rec.get("guest") or ""))
    return out


def _fill_card_rooms(records, lookup, warnings):
    """카드 승인 건의 Acct#로 객실을 채운다. 한 Acct#에 객실이 여러 개면(단체·여러 방 예약)
    승인금액 = 일일요금×일수인 객실을 고르고, 그래도 모르면 후보를 모두 적고 노란 칸으로 둔다."""
    filled, ambiguous, missing = 0, [], []
    for rec in records:
        acct = rec.get("acct")
        if not acct or rec.get("room"):
            continue
        cands = lookup.get(acct)
        if not cands:
            missing.append(f"{acct}({rec.get('amount') or 0:,.0f}원)")
            continue
        if len(cands) > 1:
            exact = [c for c in cands if c[1] and c[1] == rec.get("amount")]
            if len(exact) == 1:
                cands = exact
        if len(cands) == 1:
            rec["room"] = cands[0][0]
            filled += 1
        else:
            rec["room"] = "/".join(c[0] for c in cands)
            rec["_yellow"] = ("room",)
            ambiguous.append(f"{rec['room']} ({rec.get('amount') or 0:,.0f}원)")
    warnings.append(f"객실 예약 리스트로 객실 {filled}건을 자동으로 채웠습니다. 맞는지 확인해 주세요.")
    if ambiguous:
        warnings.append("한 예약에 객실이 여러 개라 정하지 못한 건(노란 칸, 후보 중 하나만 남기세요): " + ", ".join(ambiguous))
    if missing:
        warnings.append("예약 리스트에 없는 Acct#(이미 퇴실한 손님·조식 등 — 직접 입력): " + ", ".join(missing))


# ---------------------------------------------------------------- 카드 한벽루 줄 자동 채우기 (사용자 지시 2026-11-01)
def _prefix4(card_no):
    d = re.sub(r"\D", "", str(card_no or ""))[:4]
    return d if len(d) == 4 else None


def card_knowledge(erp_records, out_dir, cfg=None):
    """카드번호 앞 4자리 → 매입사, (사업자번호, 매입사) → 가맹점번호, 호텔·한벽루 사업자번호를 **호텔 PC에서** 익힌다.
    재료: 오늘 ERP 승인내역 + 결과 폴더의 지난 '카드 승인내역' 파일(근무자가 한벽루 줄에 적은 값 포함).
    사업자번호·가맹점번호는 저장소(GitHub)에 넣지 않는다는 규칙 → 내 PC 설정(AppData)에만 저장."""
    know = load_user_settings().get("card_learn") or {}
    prefixes = {k: dict(v) for k, v in (know.get("prefixes") or {}).items()}
    merchants = dict(know.get("merchants") or {})
    biz_count = dict(know.get("biz_count") or {})
    other_count = dict(know.get("other_biz_count") or {})

    def learn(card_no, biz, acq, merchant, other=False):
        p, biz, acq, merchant = _prefix4(card_no), clean_text(biz), clean_text(acq), clean_text(merchant)
        if p and acq:
            prefixes.setdefault(p, {})
            prefixes[p][acq] = prefixes[p].get(acq, 0) + 1
        if biz and acq and merchant:
            merchants[f"{biz}|{acq}"] = merchant
        if biz:
            (other_count if other else biz_count)[biz] = (other_count if other else biz_count).get(biz, 0) + 1

    for rec in erp_records:
        learn(rec.get("card_no"), rec.get("biz_no"), rec.get("acquirer"), rec.get("merchant_no"))
    hotel = max(biz_count, key=biz_count.get) if biz_count else None
    # 지난 결과 파일: 근무자가 직접 적은 한벽루 줄 (Acct# 참고 칸에 '한벽루' 표시)
    try:
        files = sorted((f for f in os.listdir(out_dir) if f.startswith("카드 승인내역") and f.endswith(".xlsx")),
                       reverse=True)[:30]
    except OSError:
        files = []
    for name in files:
        try:
            ws = load_workbook(os.path.join(out_dir, name), data_only=True, read_only=True).active
            rows = list(ws.iter_rows(values_only=True))
        except Exception:
            continue
        head = next((i for i, r in enumerate(rows[:10]) if r and "카드번호" in [clean_text(c) for c in r]), None)
        if head is None:
            continue
        h = [clean_text(c) for c in rows[head]]
        col = lambda k: h.index(k) if k in h else None
        cc, cb, ca, cm = col("카드번호"), col("사업자번호"), col("매입사"), col("가맹점번호")
        cn = next((j for j, x in enumerate(h) if x.startswith("Acct")), None)
        if None in (cc, cb, ca, cm):
            continue
        for r in rows[head + 1:]:
            get = lambda j: r[j] if j is not None and j < len(r) else None
            note = clean_text(get(cn))
            biz = clean_text(get(cb))
            learn(get(cc), biz, get(ca), get(cm), other="한벽루" in note or bool(hotel and biz and biz != hotel))
    other_count.pop(hotel, None)
    know = {"prefixes": prefixes, "merchants": merchants, "biz_count": biz_count, "other_biz_count": other_count}
    try:
        save_user_setting("card_learn", know)
    except ConversionError:
        pass
    return {"prefix": {p: max(v, key=v.get) for p, v in prefixes.items()}, "merchant": merchants,
            "hotel": hotel, "other": max(other_count, key=other_count.get) if other_count else None}


def find_prev_report(out_dir, prev_date, cfg=None):
    """어제 수입일보: ① ERP원본 폴더의 수작업 일보 '일보MMDD' (호텔에서 직접 만든 것이 우선) ② 결과 폴더의 '수입일보 YYYY-MM-DD.xlsx'"""
    cfg = cfg or {}
    mmdd = prev_date.strftime("%m%d")
    prefix = ((cfg.get("daily") or {}).get("names") or {}).get("prev_report", "일보")
    folder = erp_folder(cfg)
    try:
        pat = re.compile(rf"^{re.escape(norm(prefix))}\D*{mmdd}(?!\d)")
        cands = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.lower().endswith(_TABLE_EXT) and not f.startswith("~$") and pat.match(norm(os.path.splitext(f)[0]))]
        if cands:
            return max(cands, key=os.path.getmtime)
    except OSError:
        pass
    conv = (cfg.get("converters") or {}).get("income") or {}
    name = conv.get("file_pattern", "수입일보 {date}").format(date=prev_date.isoformat(), m=prev_date.month, d=prev_date.day)
    p = os.path.join(writable_out_dir(out_dir)[0], f"{name}.xlsx")
    return p if os.path.exists(p) else None


def _income_card_payments(path, cfg):
    """수입일보 ERP의 결제 줄(매출구분 '카드', 매출 0·입금 있음) → [(객실, 금액)] — 카드 승인 건의 객실 찾기용 (2026-11-02)"""
    conv = dict(cfg["converters"]["income"], required=["room"], require_room=True, room_range=None)
    try:
        records, _, _ = _extract(conv, read_table(path))
    except ConversionError:
        return []
    kinds = {norm(k) for k, v in (conv.get("prepost_by_payment") or {}).items() if v == "카.선"} or {norm("카드")}
    return [(r["room"], r.get("payment")) for r in records
            if norm(r.get("sales_kind")) in kinds and (r.get("payment") or 0) > 0 and not (r.get("sales") or 0)]


def _fill_card_rooms_by_payment(records, payments, warnings):
    """카드 승인 건 중 객실을 못 정한 건(빈칸·후보 여러 개)을 수입일보 ERP의 '카드' 결제 줄 금액으로 찾는다.
    같은 금액 결제가 한 객실뿐일 때만 채우고, 쓴 결제는 다시 쓰지 않는다 (예: 퇴실 손님 조식 추가 결제)"""
    pool = [(fmt_room(room, "strip"), amt) for room, amt in payments]
    for rec in records:  # 이미 객실이 정해진 건이 쓴 결제는 먼저 뺀다
        room = rec.get("room") or ""
        if room and "/" not in room and (room, rec.get("amount")) in pool:
            pool.remove((room, rec.get("amount")))
    filled = []
    todo = [r for r in records if not r.get("_manual") and (not r.get("room") or "/" in (r.get("room") or ""))]
    for amt in sorted({r.get("amount") for r in todo if r.get("amount")}):
        recs = [r for r in todo if r.get("amount") == amt]
        for rec in list(recs):
            allowed = set(rec["room"].split("/")) if rec.get("room") else None
            cands = [(r, a) for r, a in pool if a == amt and (allowed is None or r in allowed)]
            rooms = sorted({r for r, _ in cands})
            # 같은 금액 결제가 한 객실뿐이거나, 남은 승인 건 수와 결제 건 수가 같으면(예: 440,000원 2건 ↔ 0301·0310) 차례로 채운다
            if len(rooms) == 1 or (cands and len(cands) == len(recs)):
                pick = sorted(cands)[0]
                rec["room"] = pick[0]
                rec.pop("_yellow", None)
                pool.remove(pick)
                recs.remove(rec)
                filled.append(f"{rec['room']}({amt:,.0f}원)")
    if filled:
        warnings.append("수입일보 ERP의 카드 결제 금액으로 객실을 찾아 채웠습니다: " + ", ".join(filled))


def convert_file(kind, src_path, out_dir, business_date, cfg=None, unique=False, room_source=None, folios=None,
                 prev_report=None, income_source=None):
    """kind: income | card | credit | breakfast.
    반환: {'output': 경로, 'count': 건수, 'warnings': [...], 'title': 이름}"""
    cfg = cfg or load_config()
    conv = dict(cfg["converters"][kind])
    if "room_range" not in conv and cfg.get("valid_rooms"):
        conv["room_range"] = cfg["valid_rooms"]
    bdate = _parse_date(business_date)
    rows = read_table(src_path)
    records, warnings, mapped = _extract(conv, rows)
    notes = []
    cash = None

    card_lookup = None
    if kind == "income":
        folio_map, partners = {}, {}
        for path in folios or ():
            got = read_folio(path)
            if got:
                base = os.path.basename(path)
                m = re.match(r"\D*(\d{3,4})(?!\d)", base)
                named = m and fmt_room(m.group(1), "pad4")
                if named and named != got[0] and re.fullmatch(r"\d{4}", got[0]):
                    # 2026-11-01 실제: 같은 회사 동료 객실(612·613·705·710)의 거래내역을 받으면 대표 객실(602·604·406·510) 내역이 나옴
                    # → 동료 객실은 대표 객실이 마지막으로 정산한 다음 날부터 매일 매출로 센다 (사용자 지시 "반영해 줘")
                    reset = folio_reset_date(got[1], bdate)
                    if reset and reset >= (bdate - dt.timedelta(days=1)).isoformat():
                        reset = None  # 대표 객실이 어제까지 매일 빌이체(604처럼) → 동료 객실 정산과는 무관 (실제 10/28: 613은 2박)
                    partners[named] = (got[0], reset)
                    continue
                folio_map[got[0]] = got[1]
            else:
                warnings.append(f"객실별 거래내역으로 읽지 못한 파일: {os.path.basename(path)}")
        prev_date = bdate - dt.timedelta(days=1)
        prev_path = prev_report or find_prev_report(out_dir, prev_date, cfg)
        prev_map = read_prev_report(prev_path, prev_date) if prev_path else {}
        if prev_path and not prev_map:
            warnings.append(f"어제 일보({os.path.basename(prev_path)})에서 객실 누계를 읽지 못했습니다.")
        elif prev_path:
            warnings.append(f"어제 일보: {os.path.basename(prev_path)}")
        records = _income_sections(records, bdate, conv, warnings, folio_map, partners, prev_map)
        warnings.insert(0, "매출 열은 ERP의 '매출'(부가세 포함 금액)로 채웠습니다.")
        warnings.append("'선/후'는 노란 칸에 직접 입력하세요(카.선 / 현.선 / 후 …). "
                        "누계 = 현재 잔액(정산 뒤 묵은 날 × 매출, 수식) — 숙박수를 고치면 누계도 바뀝니다.")

    if kind == "card":
        if room_source and conv.get("room_source"):
            _fill_card_rooms(records, _room_lookup(room_source, conv), warnings)
        if income_source:
            _fill_card_rooms_by_payment(records, _income_card_payments(income_source, cfg), warnings)
        if conv.get("room_source"):  # 모든 보고서는 객실 순 (사용자 확인 2026-11-02) — 객실을 못 찾은 건은 맨 아래
            if conv.get("room_source", {}).get("sort_by_room"):
                records.sort(key=lambda r: (not r.get("room"), _room_key({"room": (r.get("room") or "").split("/")[0]})))
        if conv.get("cash_footer"):
            kws = conv.get("cash_keywords", ["현금"])
            cash_rows = [r for r in records if any(k in (r.get("kind") or "") for k in kws)]
            records = [r for r in records if r not in cash_rows]
            total_cash = sum(r.get("amount") or 0 for r in cash_rows)
            cash = {"amount": total_cash if cash_rows else None}
            if cash_rows:
                warnings.append(f"현금영수증 {len(cash_rows)}건({total_cash:,.0f}원)은 표에서 빼고 '현금영수증' 줄에 합산했습니다.")
            else:
                warnings.append("원본에 현금영수증 건이 없습니다. '현금영수증' 줄(노란 칸)은 필요하면 직접 입력하세요.")
        if cfg.get("mask_card_number", True):
            for rec in records:
                if rec.get("card_no"):
                    rec["card_no"] = mask_card(rec["card_no"])
        for rec in records:
            if rec.get("card_no"):
                rec["card_no"] = re.sub(r"x", "X", rec["card_no"], flags=re.I) if "x" in rec["card_no"].lower() else rec["card_no"]
        manual = conv.get("manual_rows") or {}
        if manual.get("count") and manual.get("auto_fill", True):
            card_lookup = card_knowledge([r for r in records if not r.get("_manual")], writable_out_dir(out_dir)[0], cfg)
            card_lookup["fields"] = manual.get("fields", {})
            if card_lookup["other"]:
                warnings.append("한벽루 줄: 카드번호 앞 4자리를 적으면 매입사·가맹점번호·구분이 자동으로 채워지고, "
                                "사업자번호는 한벽루 번호가 기본(목록에서 호텔 번호로 바꿀 수 있음)입니다.")
            else:
                warnings.append("한벽루 줄: 한벽루 사업자번호를 아직 모릅니다. 오늘 한 번 직접 적어 저장해 두면 다음부터 자동으로 채워집니다.")
        if manual.get("count"):
            records = records + [{"_manual": True, "_note": manual.get("note", "")} for _ in range(manual["count"])]
            warnings.append(f"아래 노란 빈 줄 {manual['count']}개는 {manual.get('label', 'ERP에 없는 승인 건')}을 "
                            "직접 적는 칸입니다. 계에 자동으로 더해집니다(안 쓰는 줄은 비워 두세요).")
        if "room" not in mapped and not (room_source or income_source):
            warnings.insert(0, "ERP 승인내역에는 객실번호가 없고 Acct#만 있습니다. '객실' 열은 비워 두었으니 "
                               "오른쪽 'Acct#' 열을 참고해 직접 입력해 주세요. (Acct#는 인쇄되지 않습니다)")

    if kind == "breakfast":
        has_col = "breakfast" in mapped and any(r.get("breakfast") not in ("", None) for r in records)
        before = len(records)
        kept, reasons, guessed = [], {}, []
        for rec in records:
            if has_col:
                ok, n, why = _is_breakfast(rec.get("breakfast"), conv["no_breakfast_values"]), None, "조식 열"
            else:
                ok, n, why = _bf_people(rec.get("remark"), conv, rec.get("guest"))
            if not ok:
                reasons[why] = reasons.get(why, 0) + 1
                continue
            rec["_why"] = why
            if n is not None:
                rec["pax"] = n
            elif rec.get("pax") is None:
                pax, guess = _pax_fallback(rec)
                if pax is not None:
                    rec["pax"] = pax
                    if guess and "단체" not in why:
                        guessed.append(rec.get("room"))
            kept.append(rec)
        # 같은 객실이 여러 줄이면(동숙자 등) 한 번만 센다
        by_room, dup = {}, []
        for rec in kept:
            room = rec.get("room")
            if room in by_room:
                dup.append(room)
                prev = by_room[room]
                if (rec.get("pax") or 0) > (prev.get("pax") or 0):
                    prev["pax"] = rec["pax"]
                continue
            by_room[room] = rec
        records = list(by_room.values())
        records += _check_group_totals(records, warnings)  # 단체 비고 인원보다 모자라면 '단체 추가 인원' 줄
        total = sum(r.get("pax") or 0 for r in records)
        excl = ", ".join(f"{k} {v}객실" for k, v in reasons.items())
        warnings.append(f"조식 대상 {sum(1 for r in records if not r.get('_extra'))}객실 {total}명 (전체 {before}줄 중" + (f" 제외: {excl}" if excl else "") + ").")
        if dup:
            warnings.append(f"같은 객실이 여러 줄이라 한 번만 셌습니다: {', '.join(sorted(set(dup)))}")
        if guessed:
            warnings.append("비고에 인원(N인)이 없어 성인+소인 수로 채운 객실 — 실제 조식 인원을 확인해 주세요: "
                            + ", ".join(guessed))
        _pax_sanity(records, conv, warnings)
        if conv.get("group_rows"):
            records, groups = _group_breakfast(records, conv)
            for name, n, pax in groups:
                warnings.append(f"단체 '{name}' {n}객실은 맨 아래 한 줄로 합쳤습니다 (인원 {pax if pax is not None else '확인 필요'}).")
        no_pax = sum(1 for r in records if r.get("pax") is None)
        if no_pax:
            warnings.append(f"조식 대상 중 {no_pax}객실은 인원을 알 수 없어 빈칸입니다. 직접 입력해 주세요.")
        if "remark" not in mapped and not has_col:
            warnings.append("원본에 '비고'(조식 표시) 열이 없어 조식 대상을 가려내지 못했습니다. 전체를 확인해 주세요.")
        warnings.append("비고에 '조식 포함/조포'가 있는 객실만 뽑았습니다('룸온리'·'조식불포함'·'조식X' 등은 제외). 비고가 다른 표현이면 알려 주세요.")

    if cfg.get("mask_korean_names", True) and conv.get("mask_names"):  # 조식명단에만 (사용자 지시 2026-10-26)
        for rec in records:
            if rec.get("guest"):
                rec["guest"] = rec["guest"] if rec.get("_extra") else mask_korean_name(rec["guest"])

    if conv.get("sort_by_room"):  # 단체 줄은 항상 맨 아래
        records.sort(key=lambda rec: (rec.get("_group", False), _room_key(rec)))

    if not records:
        raise ConversionError("변환할 데이터 행이 없습니다. 파일 내용이나 조건(조식 대상 등)을 확인해 주세요.")

    shown = bdate + dt.timedelta(days=conv.get("date_offset_days", 0))
    if conv.get("title_date") == "today":  # 조식명단: 영업일이 아니라 서류를 만드는 날(오늘) 날짜
        shown = dt.date.today()
    md = shown.strftime("%m/%d")
    title_text = conv["sheet_title"].format(md=md)
    fname = conv["file_pattern"].format(date=shown.strftime("%Y-%m-%d"), m=shown.month, d=shown.day)
    fname = re.sub(r'[\\/:*?"<>|]', "-", fname)

    out_dir, moved = writable_out_dir(out_dir)
    if moved:
        warnings.append(f"프로그램 폴더에 저장할 권한이 없어 '{out_dir}' 에 저장했습니다.")
    out_path = os.path.join(out_dir, f"{fname}.xlsx")
    if unique:
        out_path = unique_path(out_path)
    tpl_rel = conv.get("template")
    tpl = resolve_template(tpl_rel)
    if tpl_rel and not tpl:
        warnings.append(f"호텔 양식 파일({tpl_rel})이 없어 config.json의 기본 서식으로 만들었습니다.")
    warnings += write_workbook(out_path, shown.strftime("%m%d"), title_text, conv["output"], records,
                               conv.get("sum", []), conv.get("number_format", "0"), notes, cash,
                               conv.get("row_height"), template=tpl, font_cfg=conv.get("font"),
                               sum_label=conv.get("sum_label", "계"), sum_label_field=conv.get("sum_label_field"),
                               drop_trailer=conv.get("drop_trailer", ()),
                               fit_pages=conv.get("fit_pages", cfg.get("fit_pages")),
                               fill_pages=conv.get("fill_pages", False), width_px=conv.get("width_px"),
                               card_lookup=card_lookup)
    return {"output": out_path, "count": len(records), "warnings": warnings, "title": conv["title"]}


# ---------------------------------------------------------------- 조식 인원수 (내일 아침)
def _bf_rows(conv, bf_conv, path, what):
    """원본 → 조식 대상 레코드 (객실번호 범위 밖 제외, 비고로 조식 판단, 인원 채움)"""
    records, warnings, mapped = _extract(conv, read_table(path))
    # 조식 열이 있어도 전부 0·빈칸이면(예: 객실 예약 리스트의 '조식요금[총액]') 쓸모가 없으니 비고로 판단 (2026-10-28 실제 파일)
    has_col = "breakfast" in mapped and any(_is_breakfast(r.get("breakfast"), bf_conv["no_breakfast_values"])
                                            for r in records)
    for rec in records:
        if has_col:
            rec["_bf"], n, rec["_why"] = _is_breakfast(rec.get("breakfast"), bf_conv["no_breakfast_values"]), None, "조식 열"
        else:
            rec["_bf"], n, rec["_why"] = _bf_people(rec.get("remark"), bf_conv, rec.get("guest"))
        rec["_guess"] = False
        if n is not None:
            rec["pax"] = n
        elif rec.get("pax") is None:
            pax, guess = _pax_fallback(rec)
            if pax is not None:
                rec["pax"] = pax
                rec["_guess"] = guess and "단체" not in rec["_why"]
    if "remark" not in mapped and not has_col:
        warnings.append(f"{what}에 '비고'(조식 표시) 열이 없어 조식 대상을 가려내지 못했습니다.")
    return records, [f"[{what}] " + w for w in warnings], mapped


def _status_in(rec, words):
    st = norm(rec.get("status"))
    return bool(st) and any(norm(w) in st for w in words)


def breakfast_count(inhouse_path, reserve_path, out_dir, business_date, cfg=None):
    """내일 아침 조식 인원 = 투숙 중(IN HOUSE LIST, 오늘 밤 묵는 객실) + 오늘 도착 예정(객실 예약 리스트, 아직 체크인 전).
    저녁 8시 보고 때 미체크인 손님이 IN HOUSE LIST에 없어 빠지던 문제 (사용자 요청 2026-10-27).
    같은 손님이 두 파일에 다 있으면(이미 체크인) 한 번만 센다: 예약 상태 '투숙' · 같은 예약번호 · 같은 객실 · 같은 이름+도착일.
    반환: {'output', 'count'(객실), 'total'(명), 'inhouse', 'arrivals', 'warnings', 'title', 'bf_date'}"""
    cfg = cfg or load_config()
    conv = dict(cfg["breakfast_count"])
    bf_conv = cfg["converters"]["breakfast"]
    conv.setdefault("room_range", cfg.get("valid_rooms"))
    bdate = _parse_date(business_date)
    bf_date = bdate + dt.timedelta(days=1)
    today_s, bf_s = bdate.strftime("%Y-%m-%d"), bf_date.strftime("%Y-%m-%d")
    warnings = []

    # 1) 투숙 중: 오늘 밤 묵는 객실 (출발일이 내일 이후). 오늘 출발인데 아직 남은 손님은 내일 조식 없음
    ih_conv = dict(conv, required=["room"], require_room=True)
    ih, w, _ = _bf_rows(ih_conv, bf_conv, inhouse_path, "IN HOUSE LIST")
    warnings += w
    leaving = [r for r in ih if r.get("departure") and r["departure"] <= today_s]
    ih = [r for r in ih if r not in leaving]
    if leaving:
        warnings.append(f"오늘 출발(내일 조식 없음)이라 뺀 투숙 객실 {len(set(r['room'] for r in leaving))}개")
    by_room = {}
    for r in ih:
        if r["room"] in by_room:
            prev = by_room[r["room"]]
            prev["_bf"] = prev["_bf"] or r["_bf"]
            if r["_bf"] and (r.get("pax") or 0) > (prev.get("pax") or 0):
                prev["pax"] = r["pax"]
            continue
        by_room[r["room"]] = r
    ih_all = list(by_room.values())

    # 2) 오늘 도착 예정 (아직 체크인 전)
    rv_conv = dict(conv, required=["arrival"], require_room=False)
    rv, w, rv_mapped = _bf_rows(rv_conv, bf_conv, reserve_path, "객실 예약 리스트")
    warnings += w
    ih_rooms = {r["room"] for r in ih_all}
    ih_resv = {r.get("resv_no") for r in ih_all if r.get("resv_no")}
    ih_names = {(norm(r.get("guest")), r.get("arrival")) for r in ih_all if r.get("guest")}
    arrivals, skipped = [], {}

    def skip(why):
        skipped[why] = skipped.get(why, 0) + 1

    late = {}  # IN HOUSE LIST를 받은 뒤 체크인한 객실 (객실조회에는 CHECK IN, IN HOUSE LIST에는 없음) — 2026-10-29
    lo, hi = conv.get("room_range") or (None, None)
    for r in rv:
        room = r.get("room") or ""
        if r.get("departure") and r["departure"] <= today_s:
            skip("오늘 출발(내일 조식 없음)")
        elif _status_in(r, conv["cancel_status"]):
            skip("취소·노쇼")
        elif (r.get("resv_no") and r["resv_no"] in ih_resv) or (room and room in ih_rooms) \
                or (r.get("guest") and (norm(r["guest"]), r.get("arrival")) in ih_names):
            skip("이미 체크인(IN HOUSE LIST에 있음)")
        elif _status_in(r, conv["checked_in_status"]):
            if room and (lo is None or (room.isdigit() and lo <= int(room) <= hi)):
                late.setdefault(room, r)  # 투숙 중으로 센다
            else:
                skip("체크인했지만 객실번호 없음·범위 밖")
        elif r.get("arrival") != today_s:
            skip("오늘 도착이 아님")
        else:
            arrivals.append(r)
    if late:
        ih_all = ih_all + list(late.values())
        rooms = sorted(late)
        warnings.append(f"IN HOUSE LIST를 받은 뒤 체크인한 {len(late)}객실은 객실조회 파일에서 찾아 투숙 중으로 셌습니다 "
                        f"(조식 {sum(1 for r in late.values() if r['_bf'])}객실): "
                        + ", ".join(rooms[:10]) + (" …" if len(rooms) > 10 else ""))
    if not arrivals and rv and all(_status_in(r, conv["checked_in_status"]) for r in rv):
        warnings.append("예약 파일에 체크인 전 손님이 없습니다(모두 CHECK IN). 오늘 도착 예정 손님까지 세려면 "
                        "객실조회에서 상태를 '전체'(확정 포함)로 받아 주세요.")
    if "status" not in rv_mapped:
        warnings.append("예약 리스트에 '예약상태' 열이 없어 취소·노쇼를 가려내지 못했고, 이미 체크인한 손님은 "
                        "예약번호·객실번호·이름으로만 가려냈습니다. 취소 건이 섞이지 않았는지 확인해 주세요.")

    extras = _check_group_totals(ih_all + arrivals, warnings)  # 단체 비고 인원보다 모자라면 '단체 추가 인원'

    # 3) 집계
    def summary(rows):
        bf = [r for r in rows if r["_bf"]]
        return bf, sum(int(r.get("rooms") or 1) if not r.get("room") else 1 for r in bf), \
            sum(r.get("pax") or 0 for r in bf)

    ih_bf, ih_rooms_n, ih_pax = summary(ih_all)
    ar_bf, ar_rooms_n, ar_pax = summary(arrivals)
    extra_pax = sum(e["pax"] for e in extras)
    sanity_ok = _pax_sanity(ih_bf + ar_bf, bf_conv, warnings)
    total = ih_pax + ar_pax + extra_pax
    excl = lambda rows: sum(1 for r in rows if not r["_bf"])
    warnings.insert(0, f"{bf_date:%m/%d} 아침 조식 {total}명 = 투숙 중 {ih_pax}명({ih_rooms_n}객실) "
                       f"+ 오늘 도착 예정 {ar_pax}명({ar_rooms_n}객실)"
                       + (f" + 단체 추가 인원 {extra_pax}명" if extra_pax else ""))
    warnings.append(f"조식 아님(룸온리·표시 없음): 투숙 {excl(ih_all)}객실, 도착 예정 {excl(arrivals)}건")
    if skipped:
        warnings.append("예약 리스트에서 뺀 건: " + ", ".join(f"{k} {v}건" for k, v in skipped.items()))
    guess = [r.get("room") or r.get("guest") or "?" for r in ih_bf + ar_bf if r["_guess"]]
    if guess:
        warnings.append("비고에 조식 인원(N인)이 없어 성인+소인 수로 센 곳 — 확인해 주세요: " + ", ".join(guess))
    no_pax = [r.get("room") or r.get("guest") or "?" for r in ih_bf + ar_bf if r.get("pax") is None]
    if no_pax:
        warnings.append("인원을 알 수 없어 0명으로 센 곳 — 꼭 확인해 주세요: " + ", ".join(no_pax))
    multi = [r.get("guest") or "?" for r in ar_bf if not r.get("room") and (r.get("rooms") or 1) > 1]
    if multi:
        warnings.append("객실 여러 개를 한 줄로 예약한 건(객실 미배정) — 인원이 전체 인원인지 확인해 주세요: " + ", ".join(multi))

    # 4) 확인용 엑셀 (A4 세로 한 장 목록)
    mask = cfg.get("mask_korean_names", True)
    lines = [("투숙 중", r) for r in sorted(ih_bf, key=_room_key)] + \
            [("도착 예정", r) for r in sorted(ar_bf, key=lambda r: (r.get("room") or "9999", r.get("guest") or ""))] + \
            [("단체 추가", e) for e in extras]
    wb = Workbook()
    ws = wb.active
    ws.title = bf_date.strftime("%m%d")
    font = lambda size, b=False: Font(name="맑은 고딕", size=size, bold=b)
    ws.merge_cells("A1:E1")
    ws["A1"] = f"조식 인원 ({bf_date:%m/%d} 아침)"
    ws["A1"].font, ws["A1"].alignment = font(20, True), _CENTER
    ws.row_dimensions[1].height = 40
    info = [("합계", total, f"{ih_rooms_n + ar_rooms_n}객실"), ("투숙 중 (체크인 완료)", ih_pax, f"{ih_rooms_n}객실"),
            ("오늘 도착 예정 (체크인 전)", ar_pax, f"{ar_rooms_n}객실")]
    if extra_pax:
        info.append(("단체 추가 인원 (비고 인원 - 객실별)", extra_pax, f"{len(extras)}단체"))
    for i, (k, v, n) in enumerate(info, 2):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=2)
        ws.cell(i, 1, k).font = font(13, True)
        ws.cell(i, 3, v).font = font(16 if i == 2 else 14, True)
        ws.cell(i, 4, "명").font = font(12)
        ws.cell(i, 5, n).font = font(12)
        for c in range(1, 6):
            ws.cell(i, c).alignment = _ONE_LINE
        ws.row_dimensions[i].height = 30
    hrow = 2 + len(info) + 1
    for c, h in enumerate(["구분", "객실", "고객명", "인원", "조식 근거"], 1):
        ws.cell(hrow, c, h).font = font(11, True)
        ws.cell(hrow, c).alignment = _CENTER
    r = hrow
    for kind, rec in lines:
        r += 1
        guest = rec.get("guest") or ""
        vals = [kind, rec.get("room") or "미배정", mask_korean_name(guest) if mask and not rec.get("_extra") else guest,
                rec.get("pax"), rec.get("_why", "") + (" · 인원 확인" if rec["_guess"] or rec.get("pax") is None else "")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font, cell.alignment = font(11), _ONE_LINE
            if c == 4 and v is None:
                cell.fill = _BLANK_FILL
        ws.row_dimensions[r].height = 22
    for col, w in zip("ABCDE", (16, 14, 30, 8, 24)):
        ws.column_dimensions[col].width = w
    _draw_grid(ws, 1, 1 + len(info), 1, 5)
    if r > hrow:
        _draw_grid(ws, hrow, r, 1, 5)
    else:
        _draw_grid(ws, hrow, hrow, 1, 5)
    _fit_a4(ws, hrow)
    ws.print_area = f"A1:E{max(r, hrow)}"

    fname = re.sub(r'[\\/:*?"<>|]', "-", conv["file_pattern"].format(date=bf_s, m=bf_date.month, d=bf_date.day))
    out_dir, moved = writable_out_dir(out_dir)
    if moved:
        warnings.append(f"프로그램 폴더에 저장할 권한이 없어 '{out_dir}' 에 저장했습니다.")
    out_path = os.path.join(out_dir, f"{fname}.xlsx")
    try:
        _clean_copy(ws).save(out_path)
        _mark_styles_applied(out_path)
    except PermissionError:
        raise ConversionError(f"결과 파일이 엑셀에서 열려 있습니다. 닫은 뒤 다시 시도해 주세요.\n{out_path}")
    # 5) 호텔 양식 조식명단 (투숙 중 + 오늘 도착 예정, 사용자 요청 2026-10-29) — 4번 조식명단과 같은 양식·인쇄 설정
    list_path = None
    if ih_bf or ar_bf:
        recs = []
        for r0 in ih_bf + ar_bf:
            guest = r0.get("guest") or ""
            if cfg.get("mask_korean_names", True) and bf_conv.get("mask_names"):
                guest = mask_korean_name(guest)
            recs.append({"room": r0.get("room") or "미배정", "guest": guest, "pax": r0.get("pax")})
        recs.sort(key=lambda r0: (not r0["room"].isdigit(), _room_key(r0)))
        recs += [{"room": e["room"], "guest": e["guest"], "pax": e["pax"]} for e in extras]  # 맨 아래 (합계 바로 위)
        list_name = re.sub(r'[\\/:*?"<>|]', "-", conv.get("list_file_pattern", "조식명단(예약포함) {date}").format(
            date=bf_s, m=bf_date.month, d=bf_date.day))
        list_path = os.path.join(out_dir, f"{list_name}.xlsx")
        tpl = resolve_template(bf_conv.get("template"))
        warnings += write_workbook(list_path, bf_date.strftime("%m%d"),
                                   bf_conv["sheet_title"].format(md=bf_date.strftime("%m/%d")), bf_conv["output"], recs,
                                   bf_conv.get("sum", []), bf_conv.get("number_format", "0"), None, None,
                                   bf_conv.get("row_height"), template=tpl, font_cfg=bf_conv.get("font"),
                                   sum_label=bf_conv.get("sum_label", "계"), sum_label_field=bf_conv.get("sum_label_field"),
                                   drop_trailer=bf_conv.get("drop_trailer", ()),
                                   fit_pages=bf_conv.get("fit_pages", cfg.get("fit_pages")),
                                   fill_pages=bf_conv.get("fill_pages", False), width_px=bf_conv.get("width_px"))
        warnings.append(f"호텔 양식 조식명단({len(recs) - len(extras)}객실 {total}명)도 만들었습니다: {os.path.basename(list_path)}"
                        + (" — 미배정 객실은 객실번호를 직접 적어 주세요" if any(r0['room'] == '미배정' for r0 in recs) else ""))
    return {"output": out_path, "list": list_path, "sanity_ok": sanity_ok,
            "count": ih_rooms_n + ar_rooms_n, "total": total, "inhouse": ih_pax,
            "arrivals": ar_pax, "warnings": warnings, "title": conv["title"], "bf_date": bf_s}


def detect_type(path, cfg=None):
    """파일명 키워드 → 제목 행 고유 항목 순으로 마감 서류 종류를 추정. 실패 시 None."""
    cfg = cfg or load_config()
    name = os.path.basename(path).lower()
    for key, conv in cfg["converters"].items():
        if any(kw.lower() in name for kw in conv["keywords"]):
            return key
    try:
        rows = read_table(path)
    except Exception:
        return None
    best, best_key = (0, 0), None
    for key, conv in cfg["converters"].items():
        score, _, mapping = find_header(rows, conv["source_fields"])
        sig = sum(1 for s in conv.get("signature", []) if s in mapping)
        if sig > 0 and (sig, score) > best:
            best, best_key = (sig, score), key
    return best_key


def _can_write(folder):
    try:
        os.makedirs(folder, exist_ok=True)
        probe = os.path.join(folder, ".쓰기확인.tmp")
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
        return True
    except OSError:
        return False


def fallback_dir():
    """프로그램 폴더(공용폴더 등)에 쓸 수 없을 때 쓰는 내 PC 폴더: 문서\마감결과 → 바탕 화면\마감결과 → 임시 폴더"""
    home = os.path.expanduser("~")
    for base in (os.path.join(home, "Documents"), os.path.join(home, "문서"), os.path.join(home, "Desktop"),
                 os.environ.get("TEMP") or os.environ.get("TMP") or home):
        if base and os.path.isdir(base):
            return os.path.join(base, "마감결과")
    return os.path.join(home, "마감결과")


def writable_out_dir(folder):
    """결과 폴더에 실제로 쓸 수 있는지 확인하고, 안 되면 내 PC 폴더로 바꾼다. 반환: (폴더, 바꿨는지)"""
    if _can_write(folder):
        return folder, False
    alt = fallback_dir()
    if _can_write(alt):
        return alt, True
    raise ConversionError(f"결과를 저장할 폴더를 만들 수 없습니다.\n{folder}\n폴더 권한을 확인하거나 프로그램을 바탕화면에 복사해 실행해 주세요.")


def write_error_log(text):
    """예상치 못한 오류 내용을 '오류기록.txt'에 남긴다 (프로그램 폴더 → 안 되면 내 PC 폴더). 반환: 파일 경로 또는 None"""
    import traceback
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    body = f"[{stamp}] {text}\n{traceback.format_exc()}\n"
    for folder in (app_dir(), fallback_dir()):
        try:
            os.makedirs(folder, exist_ok=True)
            path = os.path.join(folder, "오류기록.txt")
            with open(path, "a", encoding="utf-8") as f:
                f.write(body)
            return path
        except OSError:
            continue
    return None


def _settings_path():
    """근무자가 화면에서 고른 설정(저장 위치 등)을 두는 곳: 내 PC의 AppData (공용폴더 권한과 무관)"""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "GeoseongClosingTool", "settings.json")


def load_user_settings():
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_user_setting(key, value):
    data = load_user_settings()
    data[key] = value
    path = _settings_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError as e:
        raise ConversionError(f"설정을 저장하지 못했습니다: {e}")


def default_out_dir(src_path=None, cfg=None):
    """결과 폴더 우선순위: ① 화면의 '저장 위치 변경'으로 고른 폴더 ② config.json의 output_dir
    ③ 프로그램(exe)이 있는 폴더 안의 '마감결과' (사용자 지시 2026-10-23: 한희 폴더 안에 두기).
    프로그램을 바탕화면에서 실행해도 ①로 한희 폴더를 골라 두면 결과는 한희 폴더에 모인다."""
    chosen = (load_user_settings().get("output_dir") or "").strip()
    if chosen:
        return chosen
    custom = ((cfg or {}).get("output_dir") or "").strip()
    if custom:
        return custom if os.path.isabs(custom) else os.path.join(app_dir(), custom)
    return os.path.join(app_dir(), "마감결과")


# ---------------------------------------------------------------- 오늘 마감 한 번에 (ERP원본 폴더, 사용자 요청 2026-10-29)
def desktop_dir():
    """바탕화면 폴더 (OneDrive로 옮겨진 바탕화면도 찾는다)"""
    home = os.path.expanduser("~")
    for name in (os.path.join("OneDrive", "바탕 화면"), os.path.join("OneDrive", "Desktop"), "바탕 화면", "Desktop"):
        p = os.path.join(home, name)
        if os.path.isdir(p):
            return p
    return os.path.join(home, "Desktop")


def erp_folder(cfg=None):
    """ERP 원본을 모아 두는 폴더: 화면에서 고른 폴더 > config.json daily.folder > 바탕화면의 ERP원본"""
    chosen = (load_user_settings().get("erp_folder") or "").strip()
    if chosen:
        return chosen
    custom = (((cfg or {}).get("daily") or {}).get("folder") or "").strip()
    return custom or os.path.join(desktop_dir(), "ERP원본")


_TABLE_EXT = (".xls", ".xlsx", ".xlsm", ".csv", ".htm", ".html")


def find_daily_files(folder, business_date, cfg=None):
    """파일 이름 = 약칭 + 영업일자 MMDD (예: 수입1029, 카드1029, 대외1029, 인하1029, 객실1029).
    같은 이름이 여러 개면(수입1029 (1).xls 등) 가장 최근 파일. 반환: {약칭 키: 경로}"""
    cfg = cfg or load_config()
    names = (cfg.get("daily") or {}).get("names", {})
    mmdd = _parse_date(business_date).strftime("%m%d")
    found = {}
    try:
        entries = os.listdir(folder)
    except OSError:
        return found
    for key, prefix in names.items():
        pat = re.compile(rf"^{re.escape(norm(prefix))}\D*{mmdd}(?!\d)")
        cands = [os.path.join(folder, f) for f in entries
                 if f.lower().endswith(_TABLE_EXT) and not f.startswith("~$") and pat.match(norm(os.path.splitext(f)[0]))]
        if cands:
            found[key] = max(cands, key=os.path.getmtime)
    return found


# 오늘 마감에 쓰는 ERP 파일 (사용자 요청 2026-11-02: 무엇을 넣으면 1~5번이 모두 만들어지는지 + 파일끼리 연결해 빈칸 채우기)
# (키, 화면 이름, ERP 메뉴, 필수 여부, 쓰는 곳)
DAILY_SLOTS = [
    ("income", "수입일보", "영업보고서 → 마감일보 → 수입일보 → 엑셀전환", True,
     "① 수입일보 (+ 결제 줄로 선/후, ② 카드 객실 찾기)"),
    ("card", "카드 승인내역", "객실관리 → 카드결제 → 승인 내역조회 → 엑셀전환", True, "② 카드 승인내역"),
    ("credit", "대외후불 리스트", "영업보고서 → 입금관련 → 대외 후불 리스트", True, "③ 대외후불"),
    ("inhouse", "IN HOUSE LIST", "영업보고서 → 객실/예약 → Clerk → IN HOUSE LIST", True, "④ 조식명단 · ⑤ 조식 인원수"),
    ("rooms", "객실조회 (상태 전체)", "객실관리 → 조회 및 출력 → 객실조회", True,
     "② 카드 객실 채우기 · ⑤ 오늘 도착 손님 조식"),
    ("prev_report", "어제 수입일보 (선택)", "자동: 결과 폴더의 어제 일보 / 처음 한 번만 수작업 일보를 '일보MMDD'로", False,
     "① 누계·숙박수·선/후 이어받기"),
]


def daily_named(stem, cfg=None):
    """'수입1102', '인하 1103'처럼 오늘 마감 이름 규칙(약칭 + MMDD)을 따른 파일 이름인가"""
    names = ((cfg or load_config()).get("daily") or {}).get("names", {})
    t = norm(stem)
    return any(re.match(rf"^{re.escape(norm(p))}\D*\d{{4}}(?!\d)", t) for p in names.values() if p)


def daily_plan(folder, business_date, cfg=None, overrides=None):
    """오늘 마감 파일 계획: {키: 경로 또는 None} + 거래내역 목록. overrides = 화면에서 직접 고른 파일(폴더보다 우선)"""
    cfg = cfg or load_config()
    f = find_daily_files(folder, business_date, cfg)
    bdate = _parse_date(business_date)
    plan = {k: f.get(k) for k, *_ in DAILY_SLOTS}
    if f.get("reserve") and not plan.get("rooms"):
        plan["rooms"] = f["reserve"]
    plan["prev_report"] = find_prev_report(default_out_dir(None, cfg), bdate - dt.timedelta(days=1), cfg)
    plan.update({k: v for k, v in (overrides or {}).items() if v and k != "folios"})
    named = {v for v in plan.values() if v}
    folios = list((overrides or {}).get("folios") or [])
    if not folios:  # 객실별 거래내역(정산한 객실만, 선택): 이름과 상관없이 내용으로 찾는다
        # 공유폴더에는 지난 파일이 쌓이므로 최근 며칠 안에 저장한 파일만 열어 본다 (2026-11-03, 네트워크에서 느려지지 않게)
        recent = time.time() - float((cfg.get("daily") or {}).get("folio_days", 3)) * 86400
        try:
            for name in sorted(os.listdir(folder)):
                p = os.path.join(folder, name)
                if p in named or name.startswith("~$") or not name.lower().endswith(_TABLE_EXT):
                    continue
                if daily_named(os.path.splitext(name)[0], cfg) or os.path.getmtime(p) < recent:
                    continue  # 다른 날짜의 수입·카드·인하… 파일이거나 오래된 파일
                if read_folio(p):
                    folios.append(p)
        except OSError:
            pass
    plan["folios"] = folios
    return plan


def daily_jobs(folder, business_date, cfg=None, overrides=None):
    """오늘 마감 작업 목록 [(종류, 파일, 짝 파일들)]과 안내 문구. 파일끼리 연결해 빈칸을 채운다:
    수입일보 ← 어제 일보(누계·숙박수·선/후) + 거래내역(정산 객실) / 카드 ← 객실조회(Acct#→객실) + 수입일보 ERP 카드 결제(금액→객실)
    조식 ← IN HOUSE LIST + 객실조회(오늘 도착·이후 체크인) — 객실조회가 없으면 4번 조식명단만"""
    cfg = cfg or load_config()
    plan = daily_plan(folder, business_date, cfg, overrides)
    mmdd = _parse_date(business_date).strftime("%m%d")
    names = (cfg.get("daily") or {}).get("names", {})
    jobs, notes = [], []
    if plan["income"]:
        jobs.append(("income", plan["income"], {"folios": plan["folios"] or None, "prev_report": plan["prev_report"]}))
        if plan["folios"]:
            notes.append(f"객실별 거래내역 {len(plan['folios'])}개로 정산한 객실의 누계를 계산합니다: "
                         + ", ".join(os.path.basename(p) for p in plan["folios"]))
    if plan["card"]:
        jobs.append(("card", plan["card"], {"room_source": plan["rooms"], "income_source": plan["income"]}))
    if plan["credit"]:
        jobs.append(("credit", plan["credit"], None))
    if plan["inhouse"]:
        jobs.append(("bf_count", plan["inhouse"], plan["rooms"]) if plan["rooms"] else ("breakfast", plan["inhouse"], None))
    for key, label, *_ in DAILY_SLOTS[:4]:
        if not plan[key]:
            notes.append(f"'{names.get(key, key)}{mmdd}' 파일이 없어 {label}은(는) 건너뜁니다.")
    if plan["card"] and not plan["rooms"]:
        notes.append(f"'{names.get('rooms')}{mmdd}'(객실조회) 파일이 없어 카드 객실 칸은 수입일보 결제 금액으로만 찾습니다.")
    if plan["inhouse"] and not plan["rooms"]:
        notes.append(f"'{names.get('rooms')}{mmdd}' 파일이 없어 조식은 IN HOUSE LIST만으로 조식명단을 만듭니다(도착 예정 손님 빠짐).")
    if plan["income"] and not plan["prev_report"]:
        notes.append("어제 수입일보가 없어 누계·숙박수는 도착일부터 셉니다 (처음 한 번 수작업 일보를 '일보MMDD'로 넣어 주세요).")
    found = {k: v for k, v in plan.items() if k != "folios" and v}
    return jobs, notes, found


def default_business_date(cfg=None, now=None):
    """나이트 근무는 자정을 넘기므로, 기준 시각(기본 12시) 이전이면 전날을 영업일자로 본다."""
    cfg = cfg or {}
    now = now or dt.datetime.now()
    cut = cfg.get("business_day_rollover_hour", 12)
    d = now.date() - dt.timedelta(days=1) if now.hour < cut else now.date()
    return d.strftime("%Y-%m-%d")


def today_str():
    return dt.date.today().strftime("%Y-%m-%d")
