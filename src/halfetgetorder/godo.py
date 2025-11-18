import requests, xmltodict, json
from datetime import date, timedelta
from .config import PARTNER_KEY, GODO_KEY
from .utils import _as_list, _to_int, _to_float


def fetch_add_goods_map(refresh=False):
    # 지금은 안 쓰므로 빈 dict
    return {}


def fetch_orders(created_from=None, created_to=None):
    if created_from is None:
        created_from = str(date.today() - timedelta(days=7))
    if created_to is None:
        created_to = str(date.today())

    url = (
        "https://openhub.godo.co.kr/godomall5/order/Order_Search.php"
        f"?partner_key={PARTNER_KEY}&key={GODO_KEY}"
        f"&startDate={created_from}&endDate={created_to}"
        "&dateType=order&orderStatus=g1"
    )
    r = requests.post(url, timeout=30)

    ctype = (r.headers.get('Content-Type') or '').lower()
    if 'euc-kr' in ctype or 'cp949' in ctype:
        r.encoding = 'cp949'
    elif not r.encoding:
        r.encoding = 'utf-8'

    return xmltodict.parse(r.text)


def _extract_option_info(raw):
    """
    optionInfo 문자열을 파싱해
    [[옵션명, 옵션값, ...], [...]] 형태에서
    '옵션명: 옵션값' 형태로 추출한 문자열을 반환.
    """
    if not raw:
        return ""

    try:
        opt_list = json.loads(raw)  # 문자열을 JSON으로 변환
    except:
        return ""

    parts = []
    for opt in opt_list:
        # 최소 0번(옵션명), 1번(값)은 존재해야 의미 있음
        if len(opt) >= 2:
            name = str(opt[0]).strip()
            val = str(opt[1]).strip()
            if name and val:
                parts.append(f"{name}: {val}")

    return "\n".join(parts)


def group_sets(godo_json):
    """
    고도몰 주문 JSON을 세트 구조로 묶고,
    각 parent 상품의 optionInfo를 '(필수선택) 등급: S급' 형태로 정리하여
    io_excel에서 사용 가능하도록 데이터 구조 정리.
    """
    root = godo_json.get('data', {}) if isinstance(godo_json, dict) else {}
    ret = root.get('return', {}) or {}
    orders = _as_list(ret.get('order_data'))

    results = []

    for od in orders:
        info = od.get("orderInfoData") or {}

        # 수령인 이름
        name = (info.get("receiverName") or "").strip()

        # 안심번호 처리
        safe_fl = str(info.get("receiverUseSafeNumberFl") or "").strip().lower() == "y"
        safe_no = (info.get("receiverSafeNumber") or "").strip()
        phone = safe_no if (safe_fl and safe_no) else (
            (info.get("receiverPhone") or info.get("receiverCellPhone") or "").strip()
        )

        # 주문일시
        ordered_at = (od.get("orderDate") or "").strip()

        # 배송메세지
        order_memo = (info.get("orderMemo") or "").strip()

        # 본상품, 추가옵션
        parents = _as_list(od.get("orderGoodsData"))
        adds = _as_list(od.get("addGoodsData") or od.get("orderAddGoodsData"))

        # goodsNo → parent index 매핑
        idx_by_goodsno = {}
        for i, p in enumerate(parents):
            gno = str(p.get("goodsNo") or "").strip()
            if gno:
                idx_by_goodsno[gno] = i

        # 세트 구조 기본 생성
        group = [{"parent": p, "children": []} for p in parents]

        # addGoods 를 parent 밑에 붙이기
        for add in adds:
            pno = str(add.get("parentGoodsNo") or "").strip()
            if pno and pno in idx_by_goodsno:
                group[idx_by_goodsno[pno]]["children"].append(add)

        # ★ 여기서 부모 상품 optionInfo 정리까지 포함시켜도 되지만
        # io_excel에서 읽을 때 parent에서 직접 읽도록 두는 편이 좋음.

        results.append({
            "orderedAt": ordered_at,
            "receiver": {"name": name, "phone": phone},
            "sets": group,
            "orderMemo": order_memo,
        })

    return results
