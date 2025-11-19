import os
import json
import openpyxl
from datetime import date
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import FormulaRule
from .utils import visual_len, _to_int, _to_float
from .utils import _fmt_dt, get_box_count_from_items

header_fill = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
center = Alignment(horizontal='center', vertical='center', wrap_text=False)
thin = Side(style="thin", color="000000")
thick = Side(style="thick", color="000000")

# ─────────────────────────────────────────────────────────
# Rich Text(한 셀 안에 서로 다른 스타일) 지원 여부 체크
# ─────────────────────────────────────────────────────────
try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    RICH_TEXT_AVAILABLE = True
except ImportError:
    RICH_TEXT_AVAILABLE = False


def create_orders_sheet():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "주문내역"
    # C열(총 상품결제금액)과 D열(수취인 이름) 사이에 '체크' 열 추가
    headers = [
        '플랫폼',           # A
        '주문일시',         # B
        '총 상품결제금액',   # C
        '체크',             # D (신규)
        '수취인 이름',      # E
        '상품명 + 옵션명',  # F
        '수량',             # G
        '수취인 전화번호',  # H
        '등록옵션명',       # I
        '배송메세지',       # J
    ]
    ws.append(headers)
    for c in ws[1]:
        c.fill = header_fill
    return wb, ws


def apply_border_block(ws, start_row, end_row, start_col=1, end_col=10):
    for r in range(start_row, end_row+1):
        for c in range(start_col, end_col+1):
            ws.cell(row=r, column=c).border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_thick_bottom(ws, block_start, block_end, start_col=1, end_col=10):
    for c in range(start_col, end_col+1):
        cell = ws.cell(row=block_end, column=c)
        cell.border = Border(
            left=cell.border.left or thin,
            right=cell.border.right or thin,
            top=cell.border.top or thin,
            bottom=thick
        )
    # 굵은 테두리 시작 기준 컬럼도 수취인 이름(E열=5번)로 변경
    top_left = ws.cell(row=block_start, column=5)
    top_left.border = Border(
        left=top_left.border.left or thin,
        right=top_left.border.right or thin,
        top=top_left.border.top or thin,
        bottom=thick
    )


def merge_receiver_name(ws, start_row, end_row):
    # 수취인 이름이 이제 5열(E)이므로 5번 컬럼 기준으로 병합
    if end_row > start_row:
        ws.merge_cells(start_row=start_row, start_column=5, end_row=end_row, end_column=5)
        ws.cell(row=start_row, column=5).alignment = Alignment(horizontal='center', vertical='center')


def finalize_orders_sheet(ws):
    ws.sheet_view.zoomScale = 75
    min_widths = {
        '플랫폼': 8,
        '주문일시': 16,
        '총 상품결제금액': 14,
        '체크': 6,
        '수취인 이름': 20,
        '상품명 + 옵션명': 70,
        '수량': 10,
        '수취인 전화번호': 16,
        '등록옵션명': 46,
        '배송메세지': 50
    }
    headers = [cell.value for cell in ws[1]]
    for col in ws.columns:
        col_idx = col[0].column
        col_letter = get_column_letter(col_idx)
        header = headers[col_idx - 1] if col_idx - 1 < len(headers) else ''

        max_len = visual_len(header)
        for cell in col:
            vlen = visual_len(cell.value)
            if vlen > max_len:
                max_len = vlen

            # 🔹 상품명 + 옵션명 / 배송메세지 둘 다 긴 경우 줄바꿈 허용
            if header in ('상품명 + 옵션명', '배송메세지') and vlen > 50:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)

            if header == '등록옵션명':
                cell.number_format = '@'

        auto_width = int(max_len * 0.5)
        if header == '등록옵션명':
            auto_width = int(max_len * 0.5) + 4
        target_width = max(auto_width, min_widths.get(header, 12))
        ws.column_dimensions[col_letter].width = target_width

    # 상품명+옵션명 열 인덱스: 6열(F)
    # 배송메세지 열 인덱스: 10열(J)
    for r in range(2, ws.max_row + 1):
        prod_cell = ws.cell(row=r, column=6)
        memo_cell = ws.cell(row=r, column=10)

        pclen = visual_len(prod_cell.value)
        mlen = visual_len(memo_cell.value)

        # 두 컬럼 중 더 긴 쪽 기준으로 높이 결정
        base_len = max(pclen, mlen)

        rd = ws.row_dimensions[r]

        # 이미 다른 데서 높이를 지정한 행(예: 부모행 height=65)은 건드리지 않는다
        if rd.height is not None:
            continue

        if base_len > 40:
            rd.height = 34
        else:
            rd.height = 24

    # 🔹 체크 열(D열에 값이 있고, F열이 '+ '로 시작하지 않는 = 부모행만 색상 변경)
    last_row = ws.max_row
    if last_row >= 2:
        fill_checked = PatternFill(
            start_color="FFE6FFCC",
            end_color="FFE6FFCC",
            fill_type="solid"
        )

        # AND(
        #   LEN($D2)>0,          → 체크 열에 뭔가 들어있고
        #   LEFT($F2,2)<>" + "   → F열이 "+ " 로 시작하지 않음 = 자식행이 아님
        # )
        rule = FormulaRule(
            formula=['AND(LEN($D2)>0, LEFT($F2,2)<>" + ")'],
            fill=fill_checked
        )

        # A2 ~ J{마지막 행}까지 적용 → 실제로는 부모행만 색이 들어감
        ws.conditional_formatting.add(f"A2:J{last_row}", rule)


def append_coupang_block(ws, coupang_orders):
    current_row = ws.max_row + 1
    for od in coupang_orders:
        block_start = current_row
        ordered_at = _fmt_dt(od.get('orderedAt') or od.get('orderDate', ''))

        total_price = 0.0
        for item in od.get('orderItems', []):
            price = _to_float(item.get('orderPrice', item.get('price', 0)))
            qty = _to_int(item.get('shippingCount', 1), 1)
            total_price += price * qty
        total_price_str = f"{int(total_price):,}원"

        receiver_name = (
            (od.get('shippingAddress') or {}).get('name', '') or
            (od.get('receiver') or {}).get('name', '')
        )

        item_names = []
        total_qty = 0
        for item in od.get('orderItems', []):
            name = item.get('sellerProductName') or item.get('vendorItemName') or item.get('productName') or ""
            option = item.get('sellerProductItemName') or item.get('vendorItemName') or ""
            qty = _to_int(item.get('shippingCount', 1), 1)
            total_qty += qty
            if name and option and option != name:
                item_names.append(f"{name} / {option}")
            else:
                item_names.append(name or option)
        product_info = " / ".join([x for x in item_names if x])
        total_qty = total_qty or 1

        phone = (
            (od.get('shippingAddress') or {}).get('safeNumber') or
            (od.get('receiver') or {}).get('safeNumber') or
            (od.get('receiver') or {}).get('phone') or
            (od.get('receiver') or {}).get('receiverPhone') or ''
        )

        option_names = []
        for item in od.get('orderItems', []):
            option_name = item.get('sellerProductItemName') or item.get('vendorItemName') or ""
            if option_name:
                option_names.append(str(option_name))
        option_name_str = ", ".join(option_names)

        # 🔹 쿠팡 배송메세지: parcelPrintMessage
        coupang_memo = od.get('parcelPrintMessage', '') or ''

        # A:플랫폼, B:주문일시, C:총금액, D:체크(빈칸), E:수취인, F:상품+옵션, G:수량, H:전화, I:등록옵션명, J:배송메세지
        ws.append([
            "쿠팡",
            ordered_at,
            total_price_str,
            "",                # 체크 열(사용자가 나중에 수동으로 ☑ 등 입력)
            receiver_name,
            product_info,
            total_qty,
            phone,
            option_name_str,
            coupang_memo,
        ])
        current_row += 1

        # 테두리/굵은 라인 범위 1~10열로 확장
        apply_border_block(ws, block_start, current_row - 1, 1, 10)
        merge_receiver_name(ws, block_start, current_row - 1)
        apply_thick_bottom(ws, block_start, current_row - 1, 1, 10)


def append_godo_sets(ws, grouped_orders):
    """
    고도몰 주문을 엑셀 주문내역 시트에 추가.
    - 부모행(본상품)의 '상품명 + 옵션명' 셀(6열)에:
        상품명
        orderoptionInfo
      이렇게 줄바꿈해서 표시.
    - 상품명은 볼드,
      orderoptionInfo는 회색+기울임(가능하면).
    """
    current_row = ws.max_row + 1
    for grp in grouped_orders:
        block_start = current_row
        first_parent = True

        for s in grp["sets"]:
            p = s["parent"]
            goodsCd  = (p.get('goodsCd') or '').strip()
            goodsNm  = (p.get('goodsNm') or p.get('goodsNmStandard') or '').strip()
            opt_text = (p.get('optionTextInfo') or '').strip()
            qty      = _to_int(p.get('goodsCnt', 1), 1)
            price    = _to_float(p.get('goodsPrice', 0.0), 0.0)

            # ▶ 상품명 + 옵션명(부모셀) 구성 로직
            product_name = goodsNm or goodsCd

            # 1) 먼저 orderoptionInfo / orderOptionInfo 에 사람이 읽기 좋게 들어있는지 확인
            option_info = (p.get('orderoptionInfo') or p.get('orderOptionInfo') or '').strip()

            # 2) 없으면 optionInfo(JSON 문자열)를 파싱해서 "옵션명: 값" 형태로 뽑기
            if not option_info:
                raw_opt = (p.get('optionInfo') or '').strip()
                if raw_opt:
                    try:
                        opt_list = json.loads(raw_opt)  # [[옵션명, 옵션값, ...], [...], ...]
                        parts = []
                        for opt in opt_list:
                            if isinstance(opt, (list, tuple)) and len(opt) >= 2:
                                name = str(opt[0]).strip()
                                val = str(opt[1]).strip()
                                if name and val:
                                    parts.append(f"{name}: {val}")
                        option_info = "\n".join(parts)
                    except Exception:
                        option_info = ""

            if option_info:
                product_info_parent = f"{product_name}\n{option_info}"
            else:
                product_info_parent = product_name

            reg_option_value = goodsCd

            set_total = price * (qty or 1)
            for add in s["children"]:
                add_qty   = _to_int(add.get('goodsCnt', 1), 1)
                add_price = _to_float(add.get('goodsPrice', 0.0), 0.0)
                set_total += add_price * add_qty
            total_price_str = f"{int(set_total):,}원"

            order_memo = grp.get("orderMemo", "") or grp.get("orderInfo", {}).get("orderMemo", "")

            # A:플랫폼, B:주문일시, C:총금액, D:체크, E:수취인, F:상품+옵션, G:수량, H:전화, I:등록옵션명, J:배송메세지
            ws.append([
                "고도몰",
                grp["orderedAt"] if first_parent else "",
                total_price_str,
                "",   # 체크 열(사용자 수동 입력용)
                grp["receiver"]["name"] if first_parent else "",
                product_info_parent,
                (qty or 1),
                grp["receiver"]["phone"] if first_parent else "",
                reg_option_value,
                order_memo if first_parent else ""
            ])
            current_row += 1
            first_parent = False

            # 부모 셀 스타일링 (상품명+옵션명: 6열)
            prow = current_row - 1
            pcell = ws.cell(row=prow, column=6)

            if option_info and RICH_TEXT_AVAILABLE:
                pcell.value = CellRichText(
                    TextBlock(
                        text=product_name,
                        font=InlineFont(
                            b=True
                        )
                    ),
                    TextBlock(
                        text="\n" + option_info,
                        font=InlineFont(
                            i=True,
                            color="00666666"
                        )
                    ),
                )
            else:
                pcell.value = product_info_parent
                pcell.font = Font(bold=True)

            pcell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            pcell.fill = PatternFill(start_color="FFF7F7F7", end_color="FFF7F7F7", fill_type="solid")

            ws.row_dimensions[prow].height = 65

            # 자식(추가옵션) 행
            for add in s["children"]:
                add_name = (add.get('goodsNm') or add.get('goodsNmStandard') or '').strip()
                add_qty  = _to_int(add.get('goodsCnt', 1), 1)
                # A~J 열 구조에 맞춰서 한 칸씩 밀어줌
                ws.append(["", "", "", "", "", f"+ {add_name}", add_qty, "", "", ""])
                current_row += 1
                crow = current_row - 1
                ccell = ws.cell(row=crow, column=6)
                ccell.font = Font(italic=True, color="00666666")
                ccell.alignment = Alignment(horizontal='left', vertical='center', indent=1)

        apply_border_block(ws, block_start, current_row - 1, 1, 10)
        merge_receiver_name(ws, block_start, current_row - 1)
        apply_thick_bottom(ws, block_start, current_row - 1, 1, 10)


def create_waybill_workbook(coupang_orders):
    """
    대한통운 송장등록용 엑셀 워크북 생성.
    - 시트명: '판매 주문수집'
    - 열 구조: 기존 단일 파일 버전의 first_col1 과 동일
    - coupang_orders: coupang.normalize_coupang_orders(...) 결과 리스트
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "판매 주문수집"
    ws.sheet_view.zoomScale = 75
    header = [
        '예약구분', '집하예정일', '받는분성명', '받는분전화번호', '받는분기타연락처',
        '받는분우편번호', '받는분주소(전체, 분할)', '운송장번호', '고객주문번호',
        '품목명', '박스수량', '박스타입', '기본운임', '배송메세지1',
        '배송메세지2', '품목명', '운임구분'
    ]
    ws.append(header)

    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align

    today_str = date.today().strftime('%Y%m%d')

    for od in coupang_orders:
        name = od.get("name", "")
        phone = od.get("phone", "")
        addr1 = od.get("addr1", "")
        addr2 = od.get("addr2", "")
        zipcode = od.get("zipcode", "")
        address = f"{addr1} {addr2}".strip()

        box_cnt = get_box_count_from_items(od.get("items", []))

        platform_name = "쿠팡"

        row = [
            "일반",
            today_str,
            name,
            phone,
            "",
            zipcode,
            address,
            "",
            "",
            "",
            box_cnt,
            "",
            "",
            "",
            platform_name,
            "",
            ""
        ]
        ws.append(row)

    center_align = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_align

    for column_cells in ws.columns:
        max_len = 0
        col_letter = column_cells[0].column_letter
        for cell in column_cells:
            val = str(cell.value) if cell.value is not None else ""
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max_len * 1.3 + 2

    return wb, ws
