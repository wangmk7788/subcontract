import re
from collections import defaultdict
import openpyxl


def _to_number(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("원", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _clean_str(val):
    if val is None:
        return ""
    return str(val).strip()


def parse_project_excel(file_stream):
    """
    사업대금 계산식_예시.xlsx 형태의 엑셀 파일을 분석하여
    사업 기본정보(사업명, 총 계약금액, 주사업자 계약금액)와
    하도급 업체별 금액 및 투입인력 내역을 추출합니다.
    """
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    sheet_names = wb.sheetnames

    # 시트 찾기
    cost_sheet = wb["계약금액정보"] if "계약금액정보" in sheet_names else wb[sheet_names[0]]
    staff_sheet = wb["투입인력 총괄표"] if "투입인력 총괄표" in sheet_names else (
        wb[sheet_names[1]] if len(sheet_names) > 1 else None
    )

    # 1. 사업명 추출 (상단 타이틀 셀 탐색)
    raw_title = ""
    for r in range(1, 4):
        for c in range(1, 6):
            v = _clean_str(cost_sheet.cell(r, c).value)
            if v and len(v) > 3:
                raw_title = v
                break
        if raw_title:
            break

    project_name = re.sub(r"\s*(선금신청|기성신청|잔금신청|대금계산).*$", "", raw_title).strip()
    if not project_name:
        project_name = raw_title or "공공 SW 구축 사업"

    # 2. 총 계약금액 및 주사업자 계약금액 추출
    total_contract_amount = 0.0
    prime_contract_amount = 0.0

    for r in range(1, cost_sheet.max_row + 1):
        b_val = _clean_str(cost_sheet.cell(r, 2).value)
        if b_val == "합계" and r < 25:
            c_val = _to_number(cost_sheet.cell(r, 3).value)
            d_val = _to_number(cost_sheet.cell(r, 4).value)
            if c_val > 0:
                total_contract_amount = c_val
            if d_val > 0:
                prime_contract_amount = d_val
            break

    # 하단 요약표 보조 확인
    if total_contract_amount == 0:
        for r in range(25, cost_sheet.max_row + 1):
            b_val = _clean_str(cost_sheet.cell(r, 2).value)
            if b_val == "계약금액":
                total_contract_amount = _to_number(cost_sheet.cell(r, 4).value) or _to_number(cost_sheet.cell(r, 3).value)
            elif b_val == "주사업자" and prime_contract_amount == 0:
                prime_contract_amount = _to_number(cost_sheet.cell(r, 3).value)

    # 3. 계약금액정보 시트에서 [하도급 선금 지급] 테이블 파싱
    subcontractors_cost = {}
    sub_header_row = None
    for r in range(1, cost_sheet.max_row + 1):
        b_val = _clean_str(cost_sheet.cell(r, 2).value)
        if "업체명" in b_val and "하도" in b_val:
            sub_header_row = r
            break

    if sub_header_row:
        for r in range(sub_header_row + 1, cost_sheet.max_row + 1):
            name = _clean_str(cost_sheet.cell(r, 2).value)
            if not name or name == "합계" or "업체명" in name:
                break
            amount = _to_number(cost_sheet.cell(r, 3).value)
            advance = _to_number(cost_sheet.cell(r, 4).value)
            balance = _to_number(cost_sheet.cell(r, 5).value)
            subcontractors_cost[name] = {
                "name": name,
                "amount": int(round(amount)),
                "advance": int(round(advance)),
                "balance": int(round(balance)),
            }

    # 4. 투입인력 총괄표 시트에서 하도급 업체별 업무/MM/인력 파싱
    staff_by_company = defaultdict(list)
    if staff_sheet is not None:
        header_row = None
        col_map = {}
        for r in range(1, min(staff_sheet.max_row + 1, 25)):
            row_texts = [_clean_str(staff_sheet.cell(r, c).value) for c in range(1, staff_sheet.max_column + 1)]
            if "구분" in row_texts and "소속회사" in row_texts:
                header_row = r
                for c_idx, txt in enumerate(row_texts, start=1):
                    if txt:
                        col_map[txt] = c_idx
                break

        if header_row:
            c_type = col_map.get("구분", 2)
            c_comp = col_map.get("소속회사", 3)
            c_task = col_map.get("담당업무", 4)
            c_detail = col_map.get("업무상세", 5)
            c_person = col_map.get("서명", col_map.get("성명", 6))
            c_rank = col_map.get("직급", 7)
            c_prop_mm = col_map.get("제안MM", 8)
            c_chg_mm = col_map.get("변경MM", 9)
            c_role = col_map.get("직무", 10)
            c_total = col_map.get("합계", 17)

            for r in range(header_row + 1, staff_sheet.max_row + 1):
                gubun = _clean_str(staff_sheet.cell(r, c_type).value)
                comp = _clean_str(staff_sheet.cell(r, c_comp).value)
                if "하도급" in gubun and comp:
                    mm = _to_number(staff_sheet.cell(r, c_chg_mm).value) or _to_number(staff_sheet.cell(r, c_prop_mm).value)
                    staff_by_company[comp].append({
                        "task": _clean_str(staff_sheet.cell(r, c_task).value),
                        "detail": _clean_str(staff_sheet.cell(r, c_detail).value),
                        "person": _clean_str(staff_sheet.cell(r, c_person).value),
                        "rank": _clean_str(staff_sheet.cell(r, c_rank).value),
                        "mm": round(mm, 2),
                        "role": _clean_str(staff_sheet.cell(r, c_role).value),
                        "amount": int(round(_to_number(staff_sheet.cell(r, c_total).value))),
                    })

    # 5. 데이터 통합
    all_names = list(subcontractors_cost.keys())
    for cname in staff_by_company.keys():
        if cname not in subcontractors_cost:
            all_names.append(cname)

    subcontractors = []
    for idx, name in enumerate(all_names, start=1):
        cost_info = subcontractors_cost.get(name, {})
        staff_list = staff_by_company.get(name, [])

        amount = cost_info.get("amount", 0)
        if amount == 0 and staff_list:
            amount = sum(s["amount"] for s in staff_list)

        total_mm = round(sum(s["mm"] for s in staff_list), 2)
        headcount = len(staff_list)

        # 업무 요약 텍스트 생성
        task_parts = []
        seen_tasks = set()
        for s in staff_list:
            t_str = f"{s['task']}({s['detail']})" if s["detail"] and s["detail"] != s["task"] else s["task"]
            if t_str and t_str not in seen_tasks:
                seen_tasks.add(t_str)
                task_parts.append(t_str)

        task_summary = ", ".join(task_parts) if task_parts else "전문 기술지원 및 시스템 구축"
        roles = sorted(set(s["role"] for s in staff_list if s["role"]))
        role_text = ", ".join(roles) if roles else "전문인력"

        default_sub_contract_name = f"{project_name} ({task_summary} 용역)"
        if staff_list:
            default_reason = (
                f"1. 하도급 내역: {task_summary} 수행 (투입인력 {headcount}명, 총 {total_mm:g} M/M / 직무: {role_text})\n"
                f"2. 하도급 사유: 해당 분야({task_summary})의 전문 기술력 확보 및 안정적 사업 수행을 위한 전문업체 위탁"
            )
        else:
            default_reason = (
                f"1. 하도급 내역: {project_name} 일부 과업 수행\n"
                f"2. 하도급 사유: 전문 기술력 확보 및 일정 준수를 위한 전문업체 위탁"
            )

        ratio_total = round((amount / total_contract_amount) * 100, 2) if total_contract_amount > 0 else 0.0
        ratio_prime = round((amount / prime_contract_amount) * 100, 2) if prime_contract_amount > 0 else 0.0

        subcontractors.append({
            "id": idx,
            "name": name,
            "rep_name": "",
            "biz_no": "",
            "amount": amount,
            "advance": cost_info.get("advance", 0),
            "balance": cost_info.get("balance", 0),
            "ratio_total": ratio_total,
            "ratio_prime": ratio_prime,
            "headcount": headcount,
            "total_mm": total_mm,
            "task_summary": task_summary,
            "sub_contract_name": default_sub_contract_name,
            "reason": default_reason,
            "staff": staff_list,
        })

    return {
        "project_name": project_name,
        "raw_title": raw_title,
        "total_contract_amount": int(round(total_contract_amount)),
        "prime_contract_amount": int(round(prime_contract_amount)),
        "subcontractor_total_amount": sum(s["amount"] for s in subcontractors),
        "subcontractors": subcontractors,
    }
