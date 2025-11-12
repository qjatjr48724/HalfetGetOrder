
import hmac, hashlib, urllib.parse, urllib.request, ssl, json, time
from datetime import date, datetime, timedelta
from .config import CP_ACCESS, CP_SECRET

CONTENT_TYPE = "application/json;charset=UTF-8"
METHOD = "GET"
VENDOR_ID = "A01093941"

def fetch_orders(created_from=None, created_to=None):
    if created_from is None: created_from = str(date.today() - timedelta(days=7))
    if created_to   is None: created_to   = str(date.today())

    datetime_signed = time.strftime('%y%m%d') + 'T' + time.strftime('%H%M%S') + 'Z'
    cp_path = f"/v2/providers/openapi/apis/api/v4/vendors/{VENDOR_ID}/ordersheets"
    cp_query = urllib.parse.urlencode({
        "createdAtFrom": created_from,
        "createdAtTo": created_to,
        "status": "INSTRUCT"
    })
    message = datetime_signed + METHOD + cp_path + cp_query
    signature = hmac.new(CP_SECRET.encode('utf-8'), message.encode('utf-8'), hashlib.sha256).hexdigest()
    authorization = (
        f"CEA algorithm=HmacSHA256, access-key={CP_ACCESS}, signed-date={datetime_signed}, signature={signature}"
    )
    cp_url = f"https://api-gateway.coupang.com{cp_path}?{cp_query}"
    req = urllib.request.Request(cp_url)
    req.add_header("Content-type", CONTENT_TYPE)
    req.add_header("Authorization", authorization)
    req.get_method = lambda: METHOD

    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    try:
        resp = urllib.request.urlopen(req, context=ctx)
        return resp.read().decode(resp.headers.get_content_charset() or "utf-8")
    except Exception as e:
        print("쿠팡 API 호출 오류:", e)
        return ""

def normalize_coupang_orders(coupang_body):
    try:
        data = json.loads(coupang_body)
    except Exception:
        return []
    orders = data.get('data') or data.get('content') or []
    norm = []
    for od in orders:
        ship = od.get('shippingAddress') or {}
        recv = od.get('receiver') or {}
        orderer = od.get('orderer', {}) or {}
        name = ship.get('name') or recv.get('name') or orderer.get('name') or ""
        phone = ship.get('safeNumber') or recv.get('safeNumber') or ship.get('phone') or ship.get('phoneNo') or recv.get('receiverPhone') or orderer.get('phone') or ""
        addr1 = ship.get('address1') or recv.get('addr1') or ""
        addr2 = ship.get('address2') or recv.get('addr2') or ""
        zipcode = ship.get('zipcode') or recv.get('zipCode') or ""
        items_raw = od.get('orderItems', []) or []
        items = [{"quantity": int(str(it.get('shippingCount') or it.get('quantity') or 1))} for it in items_raw]
        norm.append({
            "channel": "coupang",
            "name": name, "phone": phone, "addr1": addr1, "addr2": addr2,
            "zipcode": zipcode, "items": items, "raw": od
        })
    return norm
