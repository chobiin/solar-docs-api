#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_engine.py — Word 템플릿 기반 태양광 계약서류 자동입력 엔진
Word 파일 템플릿의 표(Table) 셀 + 단락(Paragraph)을 직접 찾아 채웁니다.

지원 서류:
  1. 공급계약신고서
  2. 이용신청서
  3. 전력공급계약신고서
  5. 수금용결제계좌신고서
  6. 전기설비이용계약서
  7. 준법서약서
"""

import copy
import re
from datetime import datetime
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree

TEMPLATE_DIR = Path(__file__).parent / "templates"

MY = {
    "상호명":           "한화솔루션 주식회사",
    "대표자명":         "박승덕, 남정운, 김동관",
    "대표자명_단독":    "박승덕",
    "주소":             "서울특별시 중구 청계천로 86, 23층 24층(장교동)",
    "사업자번호":       "725-85-01217",
    "법인번호":         "110111-0360935",
    "전화번호":         "02-1600-3400",
    "전기사업자번호":   "서울 2010-가-00001",
    "전기신사업자번호": "2024-서울-0001",
    "업무담당부서":     "한국사업부 채널영업팀",
}


# ════════════════════════════════════════════════════════
#  헬퍼
# ════════════════════════════════════════════════════════

def _cell_text(cell) -> str:
    return cell.text.strip()


def _set_cell_text(cell, text: str, bold=False, font_size=9, keep_format=True,
                   force_no_bold=False):
    if not cell.paragraphs:
        return
    para = cell.paragraphs[0]
    existing_format = None
    if para.runs:
        r = para.runs[0]
        existing_format = {
            'bold':      r.bold,
            'italic':    r.italic,
            'font_name': r.font.name,
            'font_size': r.font.size,
            'color':     r.font.color.rgb if r.font.color and r.font.color.type else None,
        }
    for run in para.runs:
        run._element.getparent().remove(run._element)
    new_run = para.add_run(str(text) if text else "")
    if keep_format and existing_format:
        new_run.bold = False if force_no_bold else existing_format['bold']
        if existing_format['italic'] is not None:
            new_run.italic = existing_format['italic']
        if existing_format['font_name']:
            new_run.font.name = existing_format['font_name']
        if existing_format['font_size']:
            new_run.font.size = existing_format['font_size']
        if existing_format['color']:
            new_run.font.color.rgb = existing_format['color']
    else:
        new_run.bold = bold
        new_run.font.size = Pt(font_size)
    rPr = new_run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), '맑은 고딕')


def _set_label_cell_font(cell, font_size_pt=7.5):
    """
    카테고리 라벨 셀 bold 제거 + 크기 축소.
    run이 없는 병합셀(XML 직접 접근)도 처리.
    """
    for para in cell.paragraphs:
        if para.runs:
            for run in para.runs:
                run.bold = False
                run.font.size = Pt(font_size_pt)
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.insert(0, rFonts)
                rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        else:
            # run이 없는 셀: pPr > rPr 직접 설정
            pPr = para._element.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                para._element.insert(0, pPr)
            rPr_def = pPr.find(qn('w:rPr'))
            if rPr_def is None:
                rPr_def = OxmlElement('w:rPr')
                pPr.append(rPr_def)
            for tag in (qn('w:b'), qn('w:bCs')):
                el = rPr_def.find(tag)
                if el is not None:
                    rPr_def.remove(el)
            for tag, attr in ((qn('w:sz'), qn('w:val')), (qn('w:szCs'), qn('w:val'))):
                el = rPr_def.find(tag)
                if el is None:
                    el = OxmlElement(tag)
                    rPr_def.append(el)
                el.set(attr, str(int(font_size_pt * 2)))
            rFonts = rPr_def.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr_def.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), '맑은 고딕')


def _find_cell_after_label(table, label: str, col_offset: int = 1):
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            if label in _cell_text(cell):
                target_ci = ci + col_offset
                if target_ci < len(row.cells):
                    return row.cells[target_ci]
    return None


def _find_cell_by_label(table, label: str):
    for row in table.rows:
        for cell in row.cells:
            if label in _cell_text(cell):
                return cell
    return None


def _find_row_by_label(table, label: str):
    for row in table.rows:
        for cell in row.cells:
            if label in _cell_text(cell):
                return row
    return None


def _get_value_cell(row, label_col: int = 0, value_col: int = None):
    cells = row.cells
    if value_col is not None:
        return cells[value_col] if value_col < len(cells) else None
    for i in range(label_col + 1, len(cells)):
        if not cells[i].text.strip():
            return cells[i]
    return cells[-1] if cells else None


def _fill_signature_paragraphs(doc, addr: str, rep: str,
                                addr_labels=("주소", "주 소"),
                                rep_labels=("대표자", "대 표 자"),
                                skip_keywords=("서울특별시", "박승덕", "한화솔루션")):
    """
    단락(paragraph) 기반 서명란에서 주소/대표자를 채웁니다.
    Word에서 서명란이 표가 아닌 단락으로 구성된 경우 사용.
    라벨 run 바로 다음 빈 run에 값을 삽입합니다.
    """
    for para in doc.paragraphs:
        txt = para.text
        if any(kw in txt for kw in skip_keywords):
            continue
        runs = para.runs
        for idx, run in enumerate(runs):
            rt = run.text.strip()
            # 주소 라벨
            if any(lbl in rt for lbl in addr_labels):
                if len(rt) > max(len(l) for l in addr_labels) + 2:
                    continue
                for j in range(idx + 1, len(runs)):
                    if not runs[j].text.strip():
                        runs[j].text = addr
                        break
                else:
                    para.add_run(addr)
            # 대표자 라벨
            elif any(lbl in rt for lbl in rep_labels):
                if len(rt) > max(len(l) for l in rep_labels) + 2:
                    continue
                for j in range(idx + 1, len(runs)):
                    if not runs[j].text.strip():
                        runs[j].text = rep
                        break
                else:
                    para.add_run(rep)


def _fill_date(doc, year, month, day):
    today = datetime.now()
    y = str(year) if year else str(today.year)
    m = str(month) if month else str(today.month)
    d = str(day) if day else str(today.day)
    for para in doc.paragraphs:
        if '년' in para.text and '월' in para.text and '일' in para.text:
            for run in para.runs:
                run.text = run.text.replace(
                    '20   년   월   일',
                    f'20{y[-2:]}년  {m}월  {d}일'
                )
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if '년' in para.text and '월' in para.text and '일' in para.text:
                        for run in para.runs:
                            t = run.text
                            t = re.sub(r'20\s*년\s*월\s*일',
                                       f'20{y[-2:]}년  {m}월  {d}일', t)
                            run.text = t


# ════════════════════════════════════════════════════════
#  서류 1: 공급계약신고서
# ════════════════════════════════════════════════════════

def fill_doc1_supply_contract(d: dict) -> Document:
    tpl = TEMPLATE_DIR / "1_공급계약신고서.docx"
    doc = Document(str(tpl))

    SKIP_KEYWORDS = ["한화솔루션", "박승덕", "725-85-01217", "110111-0360935",
                     "서울 2010", "02-1600"]

    def _is_fixed_row(row_text):
        return any(kw in row_text for kw in SKIP_KEYWORDS)

    def _fill_next_empty(cells, start_idx, value):
        for j in range(start_idx + 1, len(cells)):
            if not cells[j].text.strip():
                _set_cell_text(cells[j], value)
                return True
        return False

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = " ".join(c.text.strip() for c in cells)
            if _is_fixed_row(row_text):
                continue
            for i, cell in enumerate(cells):
                ct = cell.text.strip()
                if ct in ("상호", "상 호") and i + 1 < len(cells):
                    _fill_next_empty(cells, i, d.get("상호명", ""))
                elif ("대표자명" in ct or "대 표 자 명" in ct) and i + 1 < len(cells):
                    _fill_next_empty(cells, i, d.get("대표자명", ""))
                    for j in range(i + 1, len(cells)):
                        ctj = cells[j].text.strip()
                        if "전화번호" in ctj or "전 화 번 호" in ctj:
                            _fill_next_empty(cells, j, d.get("연락처", ""))
                            break
                elif ("대표자" in ct or "대 표 자" in ct) \
                        and "대표자명" not in ct and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        _set_cell_text(cells[i + 1], d.get("대표자명", ""))
                elif ct in ("주소", "주 소") \
                        and "서울특별시" not in row_text and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        _set_cell_text(cells[i + 1], d.get("사업장주소", ""))
                elif "법인" in ct and "등록번호" in ct and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        _set_cell_text(cells[i + 1], d.get("법인등록번호", ""))
                elif "사업자등록번호" in ct and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        _set_cell_text(cells[i + 1], d.get("사업자등록번호", ""))
                elif "전기사업자등록번호" in ct and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        _set_cell_text(cells[i + 1], d.get("전기사업자등록번호", ""))
                elif ("발전기" in ct and "주소" in ct) and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        addr = d.get("발전기주소") or d.get("사업장주소", "")
                        _set_cell_text(cells[i + 1], addr)
                elif ("설비용량" in ct or ("태양광" in ct and "kW" in ct)) \
                        and i + 1 < len(cells):
                    if not cells[i + 1].text.strip():
                        _set_cell_text(cells[i + 1], str(d.get("설비용량", "")))

    _fill_signature_paragraphs(
        doc,
        addr=d.get("사업장주소", ""),
        rep=d.get("대표자명", ""),
        skip_keywords=("서울특별시", "박승덕", "한화솔루션", "02-1600"),
    )
    _fill_date(doc, d.get("계약연도"), d.get("계약월"), d.get("계약일"))
    return doc


# ════════════════════════════════════════════════════════
#  서류 2: 이용신청서
# ════════════════════════════════════════════════════════

def _set_label_cell_font_alias(cell, font_size_pt=7.5):
    _set_label_cell_font(cell, font_size_pt)


def fill_doc2_application(d: dict) -> Document:
    tpl = TEMPLATE_DIR / "2_이용신청서.docx"
    doc = Document(str(tpl))

    LABEL_KEYWORDS = [
        "재생에너지발전사업자", "재생에 너지 발전사업자",
        "전기사용자", "전 기 사 용 자",
        "재생에너지전기공급사업자", "재생에너지 전기공급사업자",
        "발전사업자", "공급사업자",
    ]

    field_map = {
        "①사 업 자 명":   d.get("상호명", ""),
        "①사업자명":      d.get("상호명", ""),
        "사 업 자 명":    d.get("상호명", ""),
        "②대 표 자 명":  d.get("대표자명", ""),
        "②대표자명":     d.get("대표자명", ""),
        "대 표 자 명":   d.get("대표자명", ""),
        "③이 용 장 소":  d.get("발전기주소") or d.get("사업장주소", ""),
        "③이용장소":     d.get("발전기주소") or d.get("사업장주소", ""),
        "이 용 장 소":   d.get("발전기주소") or d.get("사업장주소", ""),
        "⑥전 화 번 호":  d.get("연락처", ""),
        "⑥전화번호":     d.get("연락처", ""),
        "⑧전자우편주소": d.get("이메일", ""),
    }

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = " ".join(c.text for c in cells)

            # 모든 셀에서 카테고리 라벨 확인 → bold 제거 + 폰트 축소
            for ci, cc in enumerate(cells):
                for kw in LABEL_KEYWORDS:
                    if kw in cc.text:
                        _set_label_cell_font(cc, font_size_pt=7.5)
                        break

            if "한화솔루션" in row_text or MY["전화번호"] in row_text:
                continue

            for i, cell in enumerate(cells):
                ct = cell.text.strip()

                # 하단 서명란: 주소
                if ("주소" in ct or "주 소" in ct) \
                        and "서울" not in row_text \
                        and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업장주소", ""))
                    continue

                # 하단 서명란: 대표자
                if ("대표자" in ct or "대 표 자" in ct) \
                        and "박승덕" not in row_text \
                        and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("대표자명", ""))
                    continue

                # 일반 필드 매핑
                for label, value in field_map.items():
                    if label in ct:
                        if i + 1 < len(cells):
                            target = cells[i + 1]
                            if not target.text.strip() or target.text.strip() in ("　", " "):
                                _set_cell_text(target, value)
                        break

    # 하단 서명란이 단락(paragraph) 구조인 경우 추가 처리
    _fill_signature_paragraphs(
        doc,
        addr=d.get("사업장주소", ""),
        rep=d.get("대표자명", ""),
        skip_keywords=("서울특별시", "박승덕", "한화솔루션", "02-1600"),
    )
    _fill_date(doc, d.get("계약연도"), d.get("계약월"), d.get("계약일"))
    return doc


# ════════════════════════════════════════════════════════
#  서류 3: 전력공급계약신고서
# ════════════════════════════════════════════════════════

def fill_doc3_power_contract(d: dict) -> Document:
    tpl = TEMPLATE_DIR / "3_전력공급계약신고서.docx"
    doc = Document(str(tpl))

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = " ".join(c.text for c in cells)
            if "한화솔루션" in row_text or "110111-0360935" in row_text:
                continue
            for i, cell in enumerate(cells):
                ct = cell.text.strip()
                if ct in ("상호", "상 호") and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("상호명", ""))
                elif ("대표자명" in ct or "대 표 자 명" in ct) and "박승덕" not in row_text:
                    if i + 1 < len(cells) and not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("대표자명", ""))
                elif ("전화번호" in ct or "전 화 번 호" in ct) and "1600" not in row_text:
                    if i + 1 < len(cells) and not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("연락처", ""))
                elif ct in ("주소", "주 소") and "서울특별시" not in row_text:
                    if i + 1 < len(cells) and not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업장주소", ""))
                elif "주소" in ct and "서울특별시" not in row_text and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업장주소", ""))
                elif ("대표자" in ct or "대 표 자" in ct) \
                        and "박승덕" not in row_text and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("대표자명", ""))
                elif "법인" in ct and "등록번호" in ct and "110111" not in row_text:
                    if i + 1 < len(cells) and not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("법인등록번호", ""))
                elif "사업자등록번호" in ct and "725-85" not in row_text:
                    if i + 1 < len(cells) and not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업자등록번호", ""))
                elif ("전기신사업자" in ct or "전기사업자등록번호" in ct) \
                        and "서울 2010" not in row_text:
                    if i + 1 < len(cells) and not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("전기사업자등록번호", ""))

    for para in doc.paragraphs:
        if "발전기" in para.text and "주소" in para.text:
            for run in para.runs:
                if not run.text.strip():
                    run.text = d.get("발전기주소") or d.get("사업장주소", "")

    _fill_signature_paragraphs(
        doc,
        addr=d.get("사업장주소", ""),
        rep=d.get("대표자명", ""),
        skip_keywords=("서울특별시", "박승덕", "한화솔루션", "02-1600"),
    )
    _fill_date(doc, d.get("계약연도"), d.get("계약월"), d.get("계약일"))
    return doc


# ════════════════════════════════════════════════════════
#  서류 5: 수금용결제계좌신고서
# ════════════════════════════════════════════════════════

def fill_doc5_bank_account(d: dict) -> Document:
    tpl = TEMPLATE_DIR / "5_수금용결제계좌신고서.docx"
    doc = Document(str(tpl))

    for para in doc.paragraphs:
        t = para.text
        if "위본인" in t or "주 소" in t:
            for run in para.runs:
                if "주 소" in run.text or "주소" in run.text:
                    idx = para.runs.index(run)
                    if idx + 1 < len(para.runs):
                        para.runs[idx+1].text = d.get("사업장주소", "")
        if "상 호" in t or "상호" in t:
            for run in para.runs:
                run.text = run.text.replace(
                    "상	호 :", f"상	호 : {d.get('상호명','')}"
                ).replace(
                    "상 호 :", f"상 호 : {d.get('상호명','')}"
                )
        if "대\t표\t자" in t or "대표자" in t or "대	표	자" in t:
            for run in para.runs:
                if not any(x in run.text for x in ["대", "표", "자", "인"]):
                    run.text = d.get("대표자명", "")

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = " ".join(c.text for c in cells)
            for i, cell in enumerate(cells):
                ct = cell.text.strip()
                if ("상호" in ct or "상 호" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("상호명", ""))
                elif ("대표자" in ct or "대 표 자" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("대표자명", ""))
                elif ("주소" in ct or "주 소" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업장주소", ""))
            if "은행" in row_text:
                for i, cell in enumerate(cells):
                    if "은행" in cell.text and i + 1 < len(cells):
                        if not cells[i+1].text.strip():
                            _set_cell_text(cells[i+1], d.get("은행명", ""))
            if "예금주" in row_text:
                for i, cell in enumerate(cells):
                    if "예금주" in cell.text and i + 1 < len(cells):
                        if not cells[i+1].text.strip():
                            _set_cell_text(cells[i+1], d.get("예금주명", ""))
            if "계좌번호" in row_text:
                for i, cell in enumerate(cells):
                    if "계좌번호" in cell.text and i + 1 < len(cells):
                        if not cells[i+1].text.strip():
                            _set_cell_text(cells[i+1], d.get("계좌번호", ""))

    _fill_date(doc, d.get("계약연도"), d.get("계약월"), d.get("계약일"))
    return doc


# ════════════════════════════════════════════════════════
#  서류 6: 전기설비이용계약서
# ════════════════════════════════════════════════════════

def fill_doc6_facility_contract(d: dict) -> Document:
    tpl = TEMPLATE_DIR / "6_전기설비이용계약서.docx"
    doc = Document(str(tpl))

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            row_text = " ".join(c.text for c in cells)
            if "한국전력거래소" in row_text or "KPX" in row_text:
                continue
            if "한화솔루션" in row_text:
                continue
            if "이용장소" in row_text or "이용 장소" in row_text:
                continue
            for i, cell in enumerate(cells):
                ct = cell.text.strip()
                if ct in ("상호", "상 호", "사업자명") and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("상호명", ""))
                elif ("대표자" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("대표자명", ""))
                elif ("주소" in ct) and "서울" not in row_text and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업장주소", ""))
                elif "사업자번호" in ct and "725" not in row_text and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업자등록번호", ""))

    _fill_date(doc, d.get("계약연도"), d.get("계약월"), d.get("계약일"))
    return doc


# ════════════════════════════════════════════════════════
#  서류 7: 준법서약서
# ════════════════════════════════════════════════════════

def fill_doc7_compliance(d: dict) -> Document:
    tpl = TEMPLATE_DIR / "7_준법서약서.docx"
    doc = Document(str(tpl))

    company = d.get("상호명", "")
    rep = d.get("대표자명", "")
    addr = d.get("사업장주소", "")

    for para in doc.paragraphs:
        if "협력업체" not in para.text:
            continue
        for run in para.runs:
            if "협력업체" in run.text:
                run.text = run.text.replace("협력업체", company)

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            for i, cell in enumerate(cells):
                ct = cell.text.strip()
                if "협력업체" in ct:
                    _set_cell_text(cell, ct.replace("협력업체", company))
                    continue
                if ("회사" in ct and "법인" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], company)
                elif "사업자등록번호" in ct and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], d.get("사업자등록번호", ""))
                elif ("주소" in ct or "주 소" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], addr)
                elif ("상호" in ct or "상 호" in ct) \
                        and "회사" not in ct and "법인" not in ct \
                        and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], company)
                elif ("대표자" in ct or "대 표 자" in ct) and i + 1 < len(cells):
                    if not cells[i+1].text.strip():
                        _set_cell_text(cells[i+1], rep)

    _fill_date(doc, d.get("계약연도"), d.get("계약월"), d.get("계약일"))
    return doc


# ════════════════════════════════════════════════════════
#  통합 생성 함수
# ════════════════════════════════════════════════════════

GENERATORS = {
    "1_공급계약신고서":       fill_doc1_supply_contract,
    "2_이용신청서":           fill_doc2_application,
    "3_전력공급계약신고서":   fill_doc3_power_contract,
    "5_수금용결제계좌신고서": fill_doc5_bank_account,
    "6_전기설비이용계약서":   fill_doc6_facility_contract,
    "7_준법서약서":           fill_doc7_compliance,
}

TEMPLATE_FILENAMES = {
    "1_공급계약신고서":       "1_공급계약신고서.docx",
    "2_이용신청서":           "2_이용신청서.docx",
    "3_전력공급계약신고서":   "3_전력공급계약신고서.docx",
    "5_수금용결제계좌신고서": "5_수금용결제계좌신고서.docx",
    "6_전기설비이용계약서":   "6_전기설비이용계약서.docx",
    "7_준법서약서":           "7_준법서약서.docx",
}


def generate_all_docs(d: dict, output_dir: str) -> tuple[list, list]:
    import os
    os.makedirs(output_dir, exist_ok=True)
    data = _normalize_data(d)
    success = []
    errors = []
    for name, fn in GENERATORS.items():
        try:
            doc = fn(data)
            safe_company = re.sub(r'[\\/:*?"<>|]', '_', data.get("상호명", "발전소"))
            filename = f"{name}_{safe_company}.docx"
            filepath = os.path.join(output_dir, filename)
            doc.save(filepath)
            success.append(filepath)
        except FileNotFoundError as e:
            errors.append(f"{name}: 템플릿 파일 없음 - {e}")
        except Exception as e:
            errors.append(f"{name}: {e}")
    return success, errors


def _normalize_data(d: dict) -> dict:
    data = dict(d)
    today = datetime.now()
    if not data.get("계약연도"):
        data["계약연도"] = str(today.year)
    if not data.get("계약월"):
        data["계약월"] = str(today.month)
    if not data.get("계약일"):
        data["계약일"] = str(today.day)
    if not data.get("발전기주소"):
        data["발전기주소"] = data.get("사업장주소", "")
    biz = data.get("사업자등록번호", "")
    biz_clean = re.sub(r"[^0-9]", "", biz)
    if len(biz_clean) == 10 and "-" not in biz:
        data["사업자등록번호"] = f"{biz_clean[:3]}-{biz_clean[3:5]}-{biz_clean[5:]}"
    return data


def check_templates() -> dict:
    result = {}
    for name, fname in TEMPLATE_FILENAMES.items():
        path = TEMPLATE_DIR / fname
        result[name] = path.exists()
    return result
