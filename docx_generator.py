import io
from datetime import datetime
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls


FONT_NAME = "맑은 고딕"


def _set_cell_font(run, size_pt=9.5, bold=False, color=None):
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)


def _set_cell_text(cell, text, size_pt=9.5, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT,
                   valign=WD_ALIGN_VERTICAL.CENTER, space_before=1, space_after=1, line_spacing=1.15):
    cell.vertical_alignment = valign
    cell.text = ""
    lines = str(text).split("\n")
    for idx, line in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before if idx == 0 else 0.5)
        p.paragraph_format.space_after = Pt(space_after if idx == len(lines) - 1 else 0.5)
        p.paragraph_format.line_spacing = line_spacing
        run = p.add_run(line)
        _set_cell_font(run, size_pt=size_pt, bold=bold)


def _set_cell_shading(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def _set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_table_borders(table, color="595959", sz="4", outer_sz="8"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{outer_sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def _set_row_height(row, mm_val, rule="atLeast"):
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(mm_val * 56.7)))
    trHeight.set(qn("w:hRule"), rule)
    trPr.append(trHeight)


def _set_col_widths(table, widths_mm):
    table.autofit = False
    tblGrid = table._tbl.tblGrid
    if tblGrid is not None:
        for child in list(tblGrid):
            tblGrid.remove(child)
        for w in widths_mm:
            gridCol = OxmlElement("w:gridCol")
            gridCol.set(qn("w:w"), str(int(w * 56.7)))
            tblGrid.append(gridCol)
    for row in table.rows:
        for idx, w in enumerate(widths_mm):
            if idx < len(row.cells):
                row.cells[idx].width = Mm(w)


def _format_korean_date(date_str):
    """YYYY-MM-DD -> YYYY년 M월 D일"""
    if not date_str:
        return "2026년   월   일"
    s = str(date_str).strip()
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        return f"{dt.year}년 {dt.month}월 {dt.day}일"
    except ValueError:
        return s


def _format_dot_date(date_str):
    """YYYY-MM-DD -> YYYY.MM.DD."""
    if not date_str:
        return "    .   .   ."
    s = str(date_str).strip()
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        return f"{dt.year}.{dt.month:02d}.{dt.day:02d}."
    except ValueError:
        return s


def _format_won(amount):
    try:
        return f"{int(round(float(amount))):,}원"
    except (ValueError, TypeError):
        return "0원"


def generate_subcontract_approval_docx(payload):
    """
    붙임1. 하도급 계약 승인신청서 (.docx)를 메모리 버퍼(BytesIO)로 생성합니다.
    """
    doc = Document()

    # A4 페이지 설정 (210mm x 297mm, 여백 15mm 좌우 / 12mm 상하 -> 1페이지 내 깔끔하게 수용)
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(12)
    section.bottom_margin = Mm(11)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)

    # 기본 스타일 폰트 지정
    normal_style = doc.styles["Normal"]
    normal_style.font.name = FONT_NAME
    normal_style.font.size = Pt(9.5)
    rPr = normal_style.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    for attr in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]:
        rFonts.set(qn(attr), FONT_NAME)
    rPr.insert(0, rFonts)

    # 값 추출
    approval_type = payload.get("approval_type", "하도급")
    chk_sub = "√" if approval_type == "하도급" else " "
    chk_resub = "√" if approval_type == "재하도급" else " "

    prime_company = payload.get("prime_company", "주식회사 ○○○")
    prime_rep = payload.get("prime_rep", "○○○")
    prime_biz_no = payload.get("prime_biz_no", "")

    sub_company = payload.get("sub_company", "")
    sub_rep = payload.get("sub_rep", "")
    sub_biz_no = payload.get("sub_biz_no", "")

    project_name = payload.get("project_name", "")
    contract_no = payload.get("contract_no", "")
    contract_amount = float(payload.get("contract_amount", 0) or 0)
    contract_date = _format_korean_date(payload.get("contract_date", ""))
    contract_start = _format_dot_date(payload.get("contract_start", ""))
    contract_end = _format_dot_date(payload.get("contract_end", ""))

    sub_contract_name = payload.get("sub_contract_name", "")
    sub_amount = float(payload.get("sub_amount", 0) or 0)
    ratio = (sub_amount / contract_amount * 100.0) if contract_amount > 0 else 0.0
    sub_plan_date = _format_korean_date(payload.get("sub_plan_date", ""))
    sub_start = _format_dot_date(payload.get("sub_start", ""))
    sub_end = _format_dot_date(payload.get("sub_end", ""))
    sub_reason = payload.get("sub_reason", "")

    apply_date = _format_korean_date(payload.get("apply_date", ""))
    applicant_name = payload.get("applicant_name") or prime_rep or "○○○"
    ordering_agency = payload.get("ordering_agency", "발주기관장")

    # ==========================================================
    # 상단 제목 및 안내 영역 (테두리 없는 레이아웃 테이블)
    # ==========================================================
    top_p = doc.add_paragraph()
    top_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    top_p.paragraph_format.space_before = Pt(0)
    top_p.paragraph_format.space_after = Pt(2)
    r_top = top_p.add_run("■ 소프트웨어 진흥법 시행규칙 [별지 제9호서식]")
    _set_cell_font(r_top, size_pt=8.5, color=(80, 80, 80))

    title_table = doc.add_table(rows=2, cols=3)
    title_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_col_widths(title_table, [52, 54, 74])

    # Row 0: 소프트웨어사업 | [√] 하도급 / [ ] 재하도급 | 계약 승인신청서
    c0, c1, c2 = title_table.rows[0].cells
    _set_cell_text(c0, "소프트웨어사업", size_pt=15, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _set_cell_text(c1, f"[{chk_sub}] 하 도 급\n[{chk_resub}] 재하도급", size_pt=12.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(c2, "계약 승인신청서", size_pt=15, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Row 1: 안내문구 + (앞쪽)
    m_notice = title_table.cell(1, 0).merge(title_table.cell(1, 1))
    _set_cell_text(
        m_notice,
        "※ 소프트웨어산업정보종합관리체계(www.swit.or.kr)에서도 신청할 수 있습니다.\n"
        "※ [  ]에는 해당되는 곳에 √표를 합니다.",
        size_pt=8,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _set_cell_text(title_table.cell(1, 2), "(앞쪽)", size_pt=8.5, align=WD_ALIGN_PARAGRAPH.RIGHT, valign=WD_ALIGN_VERTICAL.BOTTOM)

    for row in title_table.rows:
        for cell in row.cells:
            _set_cell_margins(cell, top=30, bottom=30, left=40, right=40)

    # ==========================================================
    # 본문 서식 메인 테이블 (6열 x 18행, 총 폭 180mm)
    # 열 폭: [22, 26, 36, 26, 35, 35] = 180mm
    # ==========================================================
    col_widths = [22, 26, 36, 26, 35, 35]
    tbl = doc.add_table(rows=18, cols=6)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(tbl, color="555555", sz="4", outer_sz="8")
    _set_col_widths(tbl, col_widths)

    LABEL_BG = "F2F4F7"
    SUB_BG = "F8F9FB"

    # --- Row 0: 접수번호 | 접수일자 | 처리기간 10일 ---
    r0_0 = tbl.cell(0, 0).merge(tbl.cell(0, 1))
    _set_cell_text(r0_0, "접수번호", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_shading(r0_0, SUB_BG)
    r0_1 = tbl.cell(0, 2).merge(tbl.cell(0, 3))
    _set_cell_text(r0_1, "접수일자", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_shading(r0_1, SUB_BG)
    r0_2 = tbl.cell(0, 4).merge(tbl.cell(0, 5))
    _set_cell_text(r0_2, "처리기간   10일", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_shading(r0_2, SUB_BG)
    _set_row_height(tbl.rows[0], 9)

    # --- Rows 1~6: 하도급 거래 계약 당사자 ---
    party_hdr = tbl.cell(1, 0).merge(tbl.cell(6, 0))
    _set_cell_text(party_hdr, "하도급\n거래\n계약\n당사자", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(party_hdr, LABEL_BG)

    # 수급인 (Rows 1~2)
    c_prime_lbl = tbl.cell(1, 1).merge(tbl.cell(2, 1))
    _set_cell_text(c_prime_lbl, "수급인\n(공동수급체\n구성원)", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(c_prime_lbl, SUB_BG)

    c_prime_name = tbl.cell(1, 2).merge(tbl.cell(1, 3))
    _set_cell_text(c_prime_name, f"상호 : {prime_company}", size_pt=9)
    c_prime_rep = tbl.cell(1, 4).merge(tbl.cell(1, 5))
    _set_cell_text(c_prime_rep, f"대표자 : {prime_rep}", size_pt=9)

    c_prime_biz = tbl.cell(2, 2).merge(tbl.cell(2, 5))
    _set_cell_text(c_prime_biz, f"사업자등록번호 : {prime_biz_no}", size_pt=9)
    _set_row_height(tbl.rows[1], 8.5)
    _set_row_height(tbl.rows[2], 8.5)

    # 하수급인 (Rows 3~4)
    c_sub_lbl = tbl.cell(3, 1).merge(tbl.cell(4, 1))
    _set_cell_text(c_sub_lbl, "하수급인", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(c_sub_lbl, SUB_BG)

    c_sub_name = tbl.cell(3, 2).merge(tbl.cell(3, 3))
    _set_cell_text(c_sub_name, f"상호 : {sub_company}", size_pt=9, bold=True)
    c_sub_rep = tbl.cell(3, 4).merge(tbl.cell(3, 5))
    _set_cell_text(c_sub_rep, f"대표자 : {sub_rep}", size_pt=9)

    c_sub_biz = tbl.cell(4, 2).merge(tbl.cell(4, 5))
    _set_cell_text(c_sub_biz, f"사업자등록번호 : {sub_biz_no}", size_pt=9)
    _set_row_height(tbl.rows[3], 8.5)
    _set_row_height(tbl.rows[4], 8.5)

    # 재하수급인 (Rows 5~6)
    c_resub_lbl = tbl.cell(5, 1).merge(tbl.cell(6, 1))
    _set_cell_text(c_resub_lbl, "재하수급인", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(c_resub_lbl, SUB_BG)

    c_resub_name = tbl.cell(5, 2).merge(tbl.cell(5, 3))
    _set_cell_text(c_resub_name, "상호 : 해당 없음", size_pt=9)
    c_resub_rep = tbl.cell(5, 4).merge(tbl.cell(5, 5))
    _set_cell_text(c_resub_rep, "대표자 : ", size_pt=9)

    c_resub_biz = tbl.cell(6, 2).merge(tbl.cell(6, 3))
    _set_cell_text(c_resub_biz, "사업자등록번호 : ", size_pt=9)
    c_resub_addr = tbl.cell(6, 4).merge(tbl.cell(6, 5))
    _set_cell_text(c_resub_addr, "소재지 : ", size_pt=9)
    _set_row_height(tbl.rows[5], 8)
    _set_row_height(tbl.rows[6], 8)

    # --- Rows 7~8: 사업 내용 ---
    biz_hdr = tbl.cell(7, 0).merge(tbl.cell(8, 0))
    _set_cell_text(biz_hdr, "사업\n내용", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(biz_hdr, LABEL_BG)

    _set_cell_text(tbl.cell(7, 1), "계약명", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(7, 1), SUB_BG)
    _set_cell_text(tbl.cell(7, 2), project_name, size_pt=9)

    _set_cell_text(tbl.cell(7, 3), "계약번호", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(7, 3), SUB_BG)
    _set_cell_text(tbl.cell(7, 4), contract_no, size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell_text(tbl.cell(7, 5), f"계약금액(A)\n{_format_won(contract_amount)}", size_pt=8.8, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _set_row_height(tbl.rows[7], 11)

    _set_cell_text(tbl.cell(8, 1), "계약일", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(8, 1), SUB_BG)
    _set_cell_text(tbl.cell(8, 2), contract_date, size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell_text(tbl.cell(8, 3), "계약기간", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(8, 3), SUB_BG)
    c_biz_period = tbl.cell(8, 4).merge(tbl.cell(8, 5))
    _set_cell_text(c_biz_period, f"{contract_start}부터 {contract_end}까지", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_row_height(tbl.rows[8], 9.5)

    # --- Rows 9~11: 하도급 내용 ---
    sub_hdr = tbl.cell(9, 0).merge(tbl.cell(11, 0))
    _set_cell_text(sub_hdr, "하도급\n내용", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(sub_hdr, LABEL_BG)

    _set_cell_text(tbl.cell(9, 1), "계약명", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(9, 1), SUB_BG)
    c_sub_cname = tbl.cell(9, 2).merge(tbl.cell(9, 3))
    _set_cell_text(c_sub_cname, sub_contract_name, size_pt=9)

    _set_cell_text(tbl.cell(9, 4), "하도급액(B)", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(9, 4), SUB_BG)
    _set_cell_text(
        tbl.cell(9, 5),
        f"{_format_won(sub_amount)}\n([비율 B/A] {ratio:.2f}%)",
        size_pt=9,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
    )
    _set_row_height(tbl.rows[9], 12)

    _set_cell_text(tbl.cell(10, 1), "계약 예정일", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(10, 1), SUB_BG)
    _set_cell_text(tbl.cell(10, 2), sub_plan_date, size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell_text(tbl.cell(10, 3), "계약기간", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(10, 3), SUB_BG)
    c_sub_period = tbl.cell(10, 4).merge(tbl.cell(10, 5))
    _set_cell_text(c_sub_period, f"{sub_start}부터 {sub_end}까지", size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_row_height(tbl.rows[10], 9.5)

    _set_cell_text(tbl.cell(11, 1), "하도급 내역\n및 사유", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(11, 1), SUB_BG)
    c_sub_reason = tbl.cell(11, 2).merge(tbl.cell(11, 5))
    _set_cell_text(c_sub_reason, sub_reason, size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT, valign=WD_ALIGN_VERTICAL.TOP, line_spacing=1.25)
    _set_row_height(tbl.rows[11], 24)

    # --- Rows 12~14: 재하도급 내용 ---
    resub_hdr = tbl.cell(12, 0).merge(tbl.cell(14, 0))
    _set_cell_text(resub_hdr, "재하도급\n내용", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(resub_hdr, LABEL_BG)

    _set_cell_text(tbl.cell(12, 1), "계약명", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(12, 1), SUB_BG)
    c_resub_cname = tbl.cell(12, 2).merge(tbl.cell(12, 3))
    _set_cell_text(c_resub_cname, "", size_pt=9)
    _set_cell_text(tbl.cell(12, 4), "재하도급액(C)", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(12, 4), SUB_BG)
    _set_cell_text(tbl.cell(12, 5), "([비율 C/B]    %)", size_pt=8.5, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _set_row_height(tbl.rows[12], 8.5)

    _set_cell_text(tbl.cell(13, 1), "계약예정일", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(13, 1), SUB_BG)
    _set_cell_text(tbl.cell(13, 2), "      년    월    일", size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(tbl.cell(13, 3), "계약기간", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(13, 3), SUB_BG)
    c_resub_period = tbl.cell(13, 4).merge(tbl.cell(13, 5))
    _set_cell_text(c_resub_period, "    .   .   .부터     .   .   .까지", size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_row_height(tbl.rows[13], 8.5)

    _set_cell_text(tbl.cell(14, 1), "재하도급\n내역 및 사유", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(14, 1), SUB_BG)
    c_resub_reason = tbl.cell(14, 2).merge(tbl.cell(14, 5))
    _set_cell_text(c_resub_reason, "해당 없음", size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_row_height(tbl.rows[14], 9.5)

    # --- Row 15: 법적 근거 및 신청인 서명란 ---
    sign_cell = tbl.cell(15, 0).merge(tbl.cell(15, 5))
    sign_cell.text = ""
    sign_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_law = sign_cell.paragraphs[0]
    p_law.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_law.paragraph_format.space_before = Pt(6)
    p_law.paragraph_format.space_after = Pt(8)
    r_law = p_law.add_run("  「소프트웨어 진흥법」 제51조제5항 및 같은 법 시행규칙 제14조제1항에 따라 위와 같이 신청합니다.")
    _set_cell_font(r_law, size_pt=9.5)

    p_date = sign_cell.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.paragraph_format.space_after = Pt(6)
    r_date = p_date.add_run(f"{apply_date}    ")
    _set_cell_font(r_date, size_pt=9.5)

    p_app = sign_cell.add_paragraph()
    p_app.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_app.paragraph_format.space_after = Pt(10)
    r_app = p_app.add_run(f"신청인                  {applicant_name}   (서명 또는 인)    ")
    _set_cell_font(r_app, size_pt=10, bold=True)

    p_to = sign_cell.add_paragraph()
    p_to.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_to.paragraph_format.space_after = Pt(4)
    r_to1 = p_to.add_run(f"  {ordering_agency} ")
    _set_cell_font(r_to1, size_pt=12.5, bold=True)
    r_to2 = p_to.add_run("귀하")
    _set_cell_font(r_to2, size_pt=10.5, bold=True)
    _set_row_height(tbl.rows[15], 38)

    # --- Row 16: 첨부서류 | 수수료 없음 ---
    _set_cell_text(tbl.cell(16, 0), "첨부서류", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(16, 0), LABEL_BG)
    c_attach = tbl.cell(16, 1).merge(tbl.cell(16, 4))
    _set_cell_text(
        c_attach,
        "1. 하도급 계약서안 사본 1부\n"
        "2. 하도급 사업수행계획서(하도급 금액 산출내역서 및 사업추진 일정표 포함) 1부\n"
        "3. 하도급 계약 적정성 판단 자기평가표 1부",
        size_pt=8.5,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.25,
    )
    _set_cell_text(tbl.cell(16, 5), "수수료\n없  음", size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_shading(tbl.cell(16, 5), SUB_BG)
    _set_row_height(tbl.rows[16], 18)

    # --- Row 17: 하단 비고 (용지규격) ---
    c_foot = tbl.cell(17, 0).merge(tbl.cell(17, 5))
    _set_cell_text(c_foot, "210mm×297mm[백상지(80g/㎡)]", size_pt=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _set_row_height(tbl.rows[17], 6)

    # 모든 셀 패딩 일괄 적용
    for row in tbl.rows:
        for cell in row.cells:
            _set_cell_margins(cell, top=45, bottom=45, left=80, right=80)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
