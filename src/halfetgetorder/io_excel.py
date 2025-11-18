import os
import json
import openpyxl
from datetime import date
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.text import InlineFont

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
    headers = ['플랫폼', '주문일시', '총 상품결제금액', '수취인 이름', '상품명 + 옵션명', '수량', '수취인 전화번호', '등록옵션명', '배송메세지']
    ws.append(headers)
    for c in ws[1]:
        c.fill = header_fill
    return wb, ws


def apply_border_block(ws, start_row, end_row, start_col=1, end_col=8):
    for r in range(start_row, end_row+1):
        for c in range(start_col, end_col+1):
            ws.cell(row=r, column=c).border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_thick_bottom(ws, block_start, block_end, start_col=1, end_col=8):
    for c in range(start_col, end_col+1):
        cell = ws.cell(row=block_end, column=c)
        cell.border = Border(
            left=cell.border.left or thin,
            right=cell.border.right or thin,
            top=cell.border.top or thin,
            bottom=thick
        )
    top_left = ws.cell(row=block_start, column=4)
    top_left.border = Border(
        left=top_left.border.left or thin,
        right=top_left.border.right or thin,
        top=top_left.border.top or thin,
        bottom=thick
    )


def merge_receiver_name(ws, start_row, end_row):
    if end_row > start_row:
        ws.merge_cells(start_row=start_row, start_column=4, end_row=end_row, end_column=4)
        ws.cell(row=start_row, column=4).alignment = Alignment(horizontal='center', vertical='center')


def finalize_orders_sheet(ws):
    ws.sheet_view.zoomScale = 75
    min_widths = {
        '플랫폼': 8,
        '주문일시': 16,
        '총 상품결제금액': 14,
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
            if header == '상품명 + 옵션명' and vlen > 50:
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

    for r in range(2, ws.max_row + 1):
        prod_cell = ws.cell(row=r, column=5)
        pclen = visual_len(prod_cell.value)

        rd = ws.row_dimensions[r]

        # 🔴 이미 다른 데서 높이를 지정한 행(예: 부모행 height=55)은 건드리지 않는다
        if rd.height is not None:
            continue

        # ⬇️ 높이가 아직 없는 행만 기본 규칙 적용
        if pclen > 40:
            rd.height = 34
        else:
            rd.height = 24


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

        ws.append([
            "쿠팡", ordered_at, total_price_str, receiver_name,
            product_info, total_qty, phone, option_name_str
        ])
        current_row += 1

        apply_border_block(ws, block_start, current_row - 1, 1, 8)
        merge_receiver_name(ws, block_start, current_row - 1)
        apply_thick_bottom(ws, block_start, current_row - 1, 1, 8)


def append_godo_sets(ws, grouped_orders):
    """
    고도몰 주문을 엑셀 주문내역 시트에 추가.
    - 부모행(본상품)의 '상품명 + 옵션명' 셀(5열)에:
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

            # ▶ 상품명 + 옵션명(부모셀) 구성 로직 수정
            #   - 1줄: 상품명(goodsNm, 없으면 goodsCd)
            #   - 2줄: orderoptionInfo (없으면 optionInfo 사용 시도)
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
                            # 최소한 0: 옵션명, 1: 옵션값
                            if isinstance(opt, (list, tuple)) and len(opt) >= 2:
                                name = str(opt[0]).strip()
                                val = str(opt[1]).strip()
                                if name and val:
                                    # "(필수선택) 제품등급: S급 외관 / S급 배터리"
                                    parts.append(f"{name}: {val}")
                        option_info = "\n".join(parts)
                    except Exception:
                        # JSON 파싱 실패하면 그냥 원본을 버리거나(raw_opt) 그대로 쓰고 싶으면 여기 선택
                        option_info = ""

            # 3) 최종 부모 셀 텍스트 구성
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

            ws.append([
                "고도몰",
                grp["orderedAt"] if first_parent else "",
                total_price_str,
                grp["receiver"]["name"] if first_parent else "",
                product_info_parent,
                (qty or 1),
                grp["receiver"]["phone"] if first_parent else "",
                reg_option_value,
                order_memo if first_parent else ""   # ← 새로 추가된 열
            ])
            current_row += 1
            first_parent = False

            # 부모 셀 스타일링
            prow = current_row - 1
            pcell = ws.cell(row=prow, column=5)

            # RichText 지원되면: 상품명/옵션에 서로 다른 스타일 적용
            if option_info and RICH_TEXT_AVAILABLE:
                pcell.value = CellRichText(
                    TextBlock(
                        text=product_name,
                        font=InlineFont(
                            b=True   # bold=True가 아니라 b=True
                        )
                    ),
                    TextBlock(
                        text="\n" + option_info,
                        font=InlineFont(
                            i=True,               # italic
                            color="00666666"      # 회색 계열
                        )
                    ),
                )
            else:
                # 지원 안 되면: 전체 문자열 그대로, 폰트만 볼드
                # (ws.append 에 이미 product_info_parent 가 들어가 있음)
                pcell.value = product_info_parent
                pcell.font = Font(bold=True)

            pcell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            pcell.fill = PatternFill(start_color="FFF7F7F7", end_color="FFF7F7F7", fill_type="solid")

            ws.row_dimensions[prow].height = 65

            # 자식(추가옵션) 행
            for add in s["children"]:
                add_name = (add.get('goodsNm') or add.get('goodsNmStandard') or '').strip()
                add_qty  = _to_int(add.get('goodsCnt', 1), 1)
                ws.append(["", "", "", "", f"+ {add_name}", add_qty, "", ""])
                current_row += 1
                crow = current_row - 1
                ccell = ws.cell(row=crow, column=5)
                ccell.font = Font(italic=True, color="00666666")
                ccell.alignment = Alignment(horizontal='left', vertical='center', indent=1)

        apply_border_block(ws, block_start, current_row - 1, 1, 9)
        merge_receiver_name(ws, block_start, current_row - 1)
        apply_thick_bottom(ws, block_start, current_row - 1, 1, 9)


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
    # 대한통운 양식 헤더 (참고용 파일 first_col1 그대로)
    header = [
        '예약구분', '집하예정일', '받는분성명', '받는분전화번호', '받는분기타연락처',
        '받는분우편번호', '받는분주소(전체, 분할)', '운송장번호', '고객주문번호',
        '품목명', '박스수량', '박스타입', '기본운임', '배송메세지1',
        '배송메세지2', '품목명', '운임구분'
    ]
    ws.append(header)

    # 헤더 스타일 (기존 주문수집 시트와 동일 톤)
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align

    # 데이터 행 작성 (쿠팡 주문만 사용)
    today_str = date.today().strftime('%Y%m%d')

    for od in coupang_orders:
        name = od.get("name", "")
        phone = od.get("phone", "")
        addr1 = od.get("addr1", "")
        addr2 = od.get("addr2", "")
        zipcode = od.get("zipcode", "")
        address = f"{addr1} {addr2}".strip()

        # 박스수량: 기존 로직 (3대까지 1박스, 이후 4대마다 1박스 추가)
        box_cnt = get_box_count_from_items(od.get("items", []))

        # 배송메세지2 에 플랫폼명(쿠팡) 기입 (참고용 스크립트와 동일 로직)
        platform_name = "쿠팡"

        row = [
            "일반",            # 예약구분
            today_str,        # 집하예정일 (YYYYMMDD)
            name,             # 받는분성명
            phone,            # 받는분전화번호
            "",               # 받는분기타연락처
            zipcode,          # 받는분우편번호
            address,          # 받는분주소(전체, 분할)
            "",               # 운송장번호
            "",               # 고객주문번호
            "",               # 품목명
            box_cnt,          # 박스수량
            "",               # 박스타입
            "",               # 기본운임
            "",               # 배송메세지1
            platform_name,    # 배송메세지2 (플랫폼명)
            "",               # 품목명 (두 번째)
            ""                # 운임구분
        ]
        ws.append(row)

    # 전체 가운데 정렬 + 열 너비 자동 조정 (참고용 파일 스타일 그대로)
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

