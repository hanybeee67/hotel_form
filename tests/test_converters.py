# -*- coding: utf-8 -*-
"""변환 로직 테스트 — 모든 데이터는 가상 데이터(실제 투숙객·카드 정보 아님)."""
import json
import os
import sys

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, Side

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import converters as cv  # noqa: E402

BDATE = "2026-10-22"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture
def cfg():
    with open(os.path.join(ROOT, "config.json"), encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture(autouse=True)
def no_templates(monkeypatch, tmp_path):
    """기본은 양식 파일이 없는 상태로 테스트 (저장소의 templates/ 영향 받지 않게)."""
    monkeypatch.setattr(cv, "app_dir", lambda: str(tmp_path / "app"))
    monkeypatch.setattr(cv, "bundled_dir", lambda: str(tmp_path / "app"))
    return tmp_path / "app"


def _xlsx(path, rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(path)
    return str(path)


def _values(path):
    ws = load_workbook(path).active
    return ws, [[c.value for c in row] for row in ws.iter_rows()]


INCOME = [
    ["매출구분", "투숙객명", "도착일", "출발일", "객실타입", "거래처", "객실#", "회원번호", "RATE", "DC", "입력자",
     "입력일시", "객실수", "NET", "봉사료", "부가세", "예수금", "매출", "입금"],
    ["ROOM", "가상손님B", "2026-10-22", "2026-10-23", "D/T", "가상여행", "0305", "", 105000, 6, "", "", 1, 90000, 0,
     9000, 0, 99000, 0],
    ["ROOM", "가상손님A", "2026-10-22", "2026-10-23", "F/T", "가상회사", "0302", "", 125000, 0, "", "", 1, 80000, 0,
     8000, 0, 88000, 0],
    ["ROOM", "가상단체", "2026-10-22", "2026-10-23", "D/T", "가상회사", "1030", "", 0, 0, "", "", 1, 0, 0,
     0, 0, 50000, 0],
]
CARD = [
    ["영업일자", "순번", "빌번호[업장]", "Acct#[객실]", "구분", "카드번호", "할부기간", "부가세", "승인금액", "실정산금액",
     "사업자번호", "응답", "발급사", "매입사", "가맹점번호", "승인일시", "승인번호", "거래번호"],
    [BDATE, 1, 0, "A001", "카드승인", "1111-xxxx-xxxx-xxxx", 0, 9000, 99000, 99000, "0000000000", "성공", "가상발급",
     "가상카드", "123", "", "11111111", "1SN"],
    [BDATE, 2, 0, "", "카드승인", "", 0, 9000, 99000, 0, "0000000000", "실패", "", "", "", "", "", ""],
    [BDATE, 3, 0, "A002", "현금영수증", "", 0, 1000, 11000, 11000, "0000000000", "성공", "", "", "", "", "222", ""],
    [None, 3, None, None, None, None, None, None, 209000, None, None, None, None, None, None, None, None, None],
]
CREDIT = [
    ["지점명", "업장명", "빌번호", "객실번호", "입금종류", "입금일자", "고객명", "도착일", "출발일", "거래처", "입금액", "카드번호"],
    ["가상호텔", "ROOM", 1, "0609", "대외후불", BDATE, "가상손님C", BDATE, "2026-10-23", "가상여행", 150000, ""],
    ["가상호텔", "ROOM", 1, "0313", "대외후불", BDATE, "가상손님D", BDATE, "2026-10-23", "가상예약", 80402, ""],
    ["계", 2, None, None, None, None, None, None, None, None, 230402, None],
]
BREAKFAST = [
    ["순번", "예약자", "연락처", "비고", "고객명", "객실번호", "인원수[성인]", "인원수[소인]"],
    [1, "가", "", "2인 조식 포함", "가상손님E", "0503", 2, 0],
    [2, "나", "", "룸온리", "가상손님F", "0601", 1, 0],
    [3, "다", "", "조포", "가상손님G", "0412", 1, 1],
    [4, "라", "", "1인 조포", "가상악단", "0601", 1, 0],
    [5, "라", "", "1인 조포", "가상악단", "0602", 1, 0],
    [6, "라", "", "1인 조포", "가상악단", "0702", 1, 0],
    [7, "마", "", "1인 조포", "가상그룹-Kim", "0703", 1, 0],
    [8, "마", "", "2인 조포", "가상그룹-Lee", "0704", 2, 0],
    [9, "바", "", "조포", "가상단체계정", "1287", 1, 0],
]


def test_income(cfg, tmp_path):
    src = _xlsx(tmp_path / "수입.xlsx", INCOME)
    res = cv.convert_file("income", src, str(tmp_path / "out"), BDATE, cfg)
    ws, v = _values(res["output"])
    assert v[0][0] == "수 입 일 보 (10/22)"
    assert [r[0] for r in v[2:4]] == ["0302", "0305"]  # 객실 순 정렬, 앞자리 0 유지
    assert v[2][7] == 88000  # 매출 = 부가세 포함 금액 (NET 아님)
    assert v[2][8] == "=H3"  # 누계 = 매출 (자기 자신을 가리키면 순환 참조)
    assert v[4][6] == "계" and v[4][7] == "=SUM(H3:H4)" and v[4][8] == "=SUM(I3:I4)"
    assert ws["F3"].fill.fgColor.rgb.endswith("FFF2B3")  # 선/후 노란 칸
    dvs = {str(dv.sqref): dv for dv in ws.data_validations.dataValidation}
    assert dvs["F3:F4"].formula1 == '"선,후"' and dvs["G3:G4"].type == "whole"  # 직접 입력 칸 도움
    assert ws.page_setup.blackAndWhite  # 노란 칸은 인쇄되지 않게
    assert ws["B3"].font.name == "굴림" and ws["B3"].font.sz == 18 and ws["A1"].font.sz == 28
    assert os.path.basename(res["output"]) == "수입일보 2026-10-22.xlsx"
    assert not any(r[0] == "1030" for r in v)  # 객실 범위(301~713) 밖 제외
    assert any("범위" in w and "1030" in w for w in res["warnings"])
    # 표 전체에 검은 선 (바깥 굵게, 안쪽 가늘게), A4 너비 맞춤
    for ref in ("A2", "E3", "I5", "B4"):
        b = ws[ref].border
        assert all(getattr(b, s).style for s in ("left", "right", "top", "bottom")), ref
        assert b.left.color.rgb.endswith("000000")
    assert ws["A3"].border.left.style == "medium" and ws["I3"].border.right.style == "medium"
    assert ws["E5"].border.bottom.style == "medium" and ws["E4"].border.bottom.style == "thin"
    ps = ws.page_setup
    assert int(ps.paperSize) == 9 and ps.orientation == "landscape" and ps.fitToWidth == 1 and ps.fitToHeight == 0
    assert ws.print_options.horizontalCentered and ws.print_title_rows == "$1:$2"


def test_card(cfg, tmp_path):
    src = _xlsx(tmp_path / "카드.xlsx", CARD)
    res = cv.convert_file("card", src, str(tmp_path / "out"), BDATE, cfg)
    _, v = _values(res["output"])
    assert len(v) == 5  # 제목, 머리글, 성공 1건, 현금영수증, 계
    assert v[2][1] == "1111-XXXX-XXXX-XXXX"
    assert v[2][3] == "가상카드"  # 발급사가 아니라 매입사
    assert v[3][5] == "현금영수증" and v[3][7] == 11000
    assert v[4][6] == "계" and v[4][7] == "=SUM(H3:H3)+H4"
    ws = load_workbook(res["output"]).active
    assert ws["A3"].fill.fgColor.rgb.endswith("FFF2B3")  # 객실은 Acct#를 보고 직접 입력하는 노란 칸


def test_credit(cfg, tmp_path):
    src = _xlsx(tmp_path / "후불.xlsx", CREDIT)
    res = cv.convert_file("credit", src, str(tmp_path / "out"), BDATE, cfg)
    _, v = _values(res["output"])
    assert v[0][0] == "대외후불리스트(10/22)"
    assert [r[0] for r in v[2:4]] == ["0313", "0609"]
    assert v[4][4] == "계" and v[4][5] == "=SUM(F3:F4)"
    assert not any("합계" in w and "다릅니다" in w for w in res["warnings"])  # 원본 합계와 일치
    assert os.path.basename(res["output"]) == "10월22일 대외후불리스트.xlsx"


def test_breakfast(cfg, tmp_path):
    src = _xlsx(tmp_path / "inhouse.xlsx", BREAKFAST)
    res = cv.convert_file("breakfast", src, str(tmp_path / "out"), BDATE, cfg)
    ws, v = _values(res["output"])
    rows = [r for r in v[2:] if r[0] and r[0].isdigit()]
    # 룸온리 제외, 인원 = N인 또는 성인+소인. '단체명-이름'처럼 이름이 있으면 각자 한 줄
    assert [(r[0], r[1], r[2]) for r in rows] == [("0412", "가상손님G", 2), ("0503", "가상손님E", 2),
                                                 ("0703", "가상그룹-Kim", 1), ("0704", "가상그룹-Lee", 2)]
    # 고객명이 똑같은 단체는 맨 아래 한 줄 (객실 수·인원 합), 301~713 밖(1287)은 제외
    groups = [r[1:3] for r in v[2:] if r[0] == "단체"]
    assert groups == [["가상악단 (3실)", 3]]
    assert [r[0] for r in v[2:7]] == ["0412", "0503", "0703", "0704", "단체"]
    assert not any(r[0] == "1287" for r in v)
    assert ws["B3"].font.name == "맑은 고딕" and ws["B3"].font.sz == 18


def test_wrong_file_message(cfg, tmp_path):
    src = _xlsx(tmp_path / "아무거나.xlsx", [["가", "나"], [1, 2]])
    with pytest.raises(cv.ConversionError) as e:
        cv.convert_file("credit", src, str(tmp_path / "out"), BDATE, cfg)
    assert "필수 항목" in str(e.value)


def test_detect_type(cfg, tmp_path):
    assert cv.detect_type(_xlsx(tmp_path / "a.xlsx", CARD), cfg) == "card"
    assert cv.detect_type(_xlsx(tmp_path / "b.xlsx", CREDIT), cfg) == "credit"
    assert cv.detect_type(_xlsx(tmp_path / "c.xlsx", BREAKFAST), cfg) == "breakfast"
    assert cv.detect_type(_xlsx(tmp_path / "d.xlsx", INCOME), cfg) == "income"


def test_template_keeps_hotel_format(cfg, tmp_path, no_templates):
    """가상 호텔 양식을 넣으면 양식의 폰트·행 높이·열 너비·병합·인쇄 방향이 그대로 유지된다."""
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "대외후불리스트(10/01)"
    ws["A1"].font = Font(name="굴림", size=16)
    ws.merge_cells("A1:F1")
    for i, h in enumerate(["객실", "고객명", "도착일", "출발일", "거래처", "금액"], 1):
        ws.cell(2, i, h).border = box
        ws.cell(3, i).font = Font(name="굴림", size=10)
        ws.cell(3, i).border = box
        ws.cell(4, i).border = box
    ws["E4"] = "계"
    ws.column_dimensions["B"].width = 9.5
    ws.row_dimensions[3].height = 21.75
    ws.page_setup.orientation = "landscape"
    os.makedirs(no_templates / "templates")
    wb.save(no_templates / "templates" / "대외후불리스트.xlsx")

    src = _xlsx(tmp_path / "후불.xlsx", CREDIT)
    res = cv.convert_file("credit", src, str(tmp_path / "out"), BDATE, cfg)
    assert not any("양식 파일" in w for w in res["warnings"])
    ws, v = _values(res["output"])
    assert v[0][0] == "대외후불리스트(10/22)" and ws["A1"].font.sz == 16
    assert ws["B3"].font.sz == 10 and ws.row_dimensions[4].height == 21.75
    assert ws.column_dimensions["B"].width == 9.5
    assert ws.page_setup.orientation == "landscape"
    assert v[4][4] == "계" and v[4][5] == "=SUM(F3:F4)"


@pytest.mark.parametrize("kind,data,title_size,body_size,font", [
    ("income", INCOME, 28, 18, "굴림"),
    ("card", CARD, 26, 16, "굴림"),
    ("credit", CREDIT, 28, 16, "굴림"),
    ("breakfast", BREAKFAST, 26, 18, "맑은 고딕"),
])
def test_repo_templates(cfg, tmp_path, monkeypatch, kind, data, title_size, body_size, font):
    """저장소의 호텔 빈 양식(templates/)으로 변환하면 호텔 양식의 글꼴·크기·열 너비가 그대로 나온다."""
    monkeypatch.setattr(cv, "app_dir", lambda: ROOT)
    tpl = load_workbook(os.path.join(ROOT, cfg["converters"][kind]["template"])).active
    src = _xlsx(tmp_path / f"{kind}.xlsx", data)
    res = cv.convert_file(kind, src, str(tmp_path / "out"), BDATE, cfg)
    assert not any("양식 파일" in w for w in res["warnings"])
    ws, v = _values(res["output"])
    assert ws["A1"].font.name == font and ws["A1"].font.sz == title_size
    assert ws["B3"].font.name == font and ws["B3"].font.sz == body_size and ws["B3"].font.b
    for col in "ABCD":
        assert ws.column_dimensions[col].width == tpl.column_dimensions[col].width
    assert ws.row_dimensions[3].height == tpl.row_dimensions[3].height
    assert "10/22" in v[0][0]
    if kind == "breakfast":
        firsts = [r[0] for r in v]
        assert "외부식사" in firsts and any(str(x).startswith("조식 명단에") for x in firsts)
        assert ws["E3"].border.right.style == "medium"  # 병합된 Sign 칸의 오른쪽 열도 테두리 유지
        assert ws.auto_filter.ref.startswith("A2:E")
