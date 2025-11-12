
import requests, xmltodict, json
from datetime import date, timedelta
from .config import PARTNER_KEY, GODO_KEY
from .utils import _as_list, _to_int, _to_float

def fetch_add_goods_map(refresh=False):
    return {}

def fetch_orders(created_from=None, created_to=None):
    if created_from is None: created_from = str(date.today() - timedelta(days=7))
    if created_to   is None: created_to   = str(date.today())
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

def group_sets(godo_json):
    root = godo_json.get('data', {}) if isinstance(godo_json, dict) else {}
    ret  = root.get('return', {}) or {}
    orders = _as_list(ret.get('order_data'))

    results = []
    for od in orders:
        info = od.get('orderInfoData') or {}
        name = (info.get('receiverName') or '').strip()
        safe_fl = str(info.get('receiverUseSafeNumberFl') or '').strip().lower() == 'y'
        safe_no = (info.get('receiverSafeNumber') or '').strip()
        phone = safe_no if (safe_fl and safe_no) else ( (info.get('receiverPhone') or info.get('receiverCellPhone') or '').strip() )
        ordered_at = (od.get('orderDate') or '').strip()

        parents = _as_list(od.get('orderGoodsData'))
        adds    = _as_list(od.get('addGoodsData') or od.get('orderAddGoodsData'))

        idx_by_goodsno = {}
        for i, p in enumerate(parents):
            gno = str(p.get('goodsNo') or '').strip()
            if gno:
                idx_by_goodsno[gno] = i

        group = [{"parent": p, "children": []} for p in parents]
        for add in adds:
            pno = str(add.get('parentGoodsNo') or '').strip()
            if pno and pno in idx_by_goodsno:
                group[idx_by_goodsno[pno]]["children"].append(add)
            else:
                pass

        results.append({
            "orderedAt": ordered_at,
            "receiver": {"name": name, "phone": phone},
            "sets": group
        })
    return results
