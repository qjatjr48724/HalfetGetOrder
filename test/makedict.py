# 1. 라이브러리 불러오기
import os
import time
import hmac, hashlib
import urllib.parse
import urllib.request
import ssl
import requests
import openpyxl
import json
import unicodedata
import math
import xmltodict
from datetime import date, datetime, timedelta
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────
# 공통 사용 변수
# ─────────────────────────────────────────────────────────
content_type = "application/json;charset=UTF-8"

os.environ['TZ'] = 'GMT+0'
datetime = time.strftime('%y%m%d') + 'T' + time.strftime('%H%M%S') + 'Z'
method = "GET"

# 날짜범위 - 시작날짜(오늘날짜 - 7일) / 종료날짜(오늘날짜)
created_from = str(date.today() - timedelta(days=7))
created_to   = str(date.today())
print(created_from, created_to)

# ─────────────────────────────────────────────────────────
# API 키/기본 설정
# ─────────────────────────────────────────────────────────
# 쿠팡 API KEY
cp_accesskey = "7b6058d9-7745-4cf8-9881-cf4a469c0512"
cp_secretkey = "a4e91ae56e91a47b696712dd29008a405e8d4c25"

# 고도몰 API KEY
partner_key = "MCVFRCVENSVCNiUyOCUxMSVEQ28="
godo_key = "JTBBJUNFJTQwZCUxMiVCOCUzQSVBRCVDQW0lRjMlODclRjUlQUElQjIlQ0QlRUYlMDMlQkUlQzklREIlOTh6bU0lMTElNUIlREElM0I0JTFCJUZCcCUxMyVFRCU0MCU3RTUlMUYlRjk="


# ─────────────────────────────────────────────────────────
# 고도몰 API 연동 (XML)
# ─────────────────────────────────────────────────────────
godo_url = (
    "https://openhub.godo.co.kr/godomall5/goods/Goods_Add_Search.php?"
    f"partner_key={partner_key}&key={godo_key}"
)

godo_req = requests.post(godo_url, timeout=30)
ctype = (godo_req.headers.get('Content-Type') or '').lower()
if 'euc-kr' in ctype or 'cp949' in ctype:
    godo_req.encoding = 'cp949'
elif not godo_req.encoding:
    godo_req.encoding = 'utf-8'

data = godo_req.text
json_data = xmltodict.parse(data)

# ─────────────────────────────────────────────────────────
# 모든 추가상품 → 딕셔너리로 정규화 (addGoodsNo 전부!)
# ─────────────────────────────────────────────────────────

def _as_list(x):
    if x is None:
        return []
    return x if isinstance(x, list) else [x]

def _txt(d, k, default=""):
    v = d.get(k)
    if v is None:
        return default
    return v.strip() if isinstance(v, str) else str(v)

def _price(d):
    # 가격 후보 키들 중 먼저 발견되는 값 사용
    for k in ("goodsPrice", "price", "salePrice", "fixedPrice", "consumerPrice"):
        if d.get(k) not in (None, ""):
            try:
                return float(str(d[k]).strip())
            except:
                pass
    return 0.0

# 루트 위치 잡기
root = json_data.get("data", {}) if isinstance(json_data, dict) else {}
ret  = root.get("return", {}) or {}

# 추가상품 리스트가 들어올 수 있는 후보 키들(상점/버전마다 다를 수 있음)
candidates = None
for k in (
    "add_goods_data", "addGoodsData",       # 추가상품 전용으로 자주 쓰임
    "goods_data", "goodsData",
    "goodsList", "goods_list", "rows", "list"
):
    if ret.get(k) is not None:
        candidates = ret.get(k)
        break

items = _as_list(candidates)

# 일부 응답은 한 단계 더 안쪽에 들어 있음 (예: {'addGoodsData': {'item': [...]}})
if items and isinstance(items[0], dict) and any(subk in items[0] for subk in ("goods", "item", "row")):
    tmp = []
    for it in items:
        for subk in ("goods", "item", "row"):
            if it.get(subk) is not None:
                tmp.extend(_as_list(it[subk]))
    if tmp:
        items = tmp

# 👉 최종 딕셔너리: addGoodsNo 를 키로, 모든 추가상품 하나하나 담기
add_goods_map = {}
for it in items:
    add_no = _txt(it, "addGoodsNo") or _txt(it, "add_goods_no") or _txt(it, "goodsNo")
    name   = _txt(it, "goodsNm") or _txt(it, "goodsNmStandard") or _txt(it, "goods_name")
    price  = _price(it)

    # addGoodsNo, goodsNm 둘 다 있어야 기록
    if add_no and name:
        add_goods_map[add_no] = {
            "goodsNm": name,
            "goodsPrice": price,
        }

print("추가상품 총 개수:", len(add_goods_map))
# 샘플 3개만 확인
print(list(add_goods_map.items())[:3])

# (옵션) 파일로 저장하고 싶다면:
# import json, time, os
# save_dir = r"C:/Users/UserK/Desktop/[작업물]/[판매] 주문수집 프로그램"
# os.makedirs(save_dir, exist_ok=True)
# ts = time.strftime("%Y%m%d_%H%M%S")
# out_path = os.path.join(save_dir, f"godo_add_goods_map_{ts}.json")
# with open(out_path, "w", encoding="utf-8") as f:
#     json.dump(add_goods_map, f, ensure_ascii=False, indent=2)
# print("저장 경로:", out_path)


# ─────────────────────────────────────────────────────────
# 추가상품 딕셔너리(JSON 파일로 저장)
# ─────────────────────────────────────────────────────────
import json
import time

# 저장 폴더 (필요시 수정 가능)
save_dir = r"C:/Users/UserK/Desktop/HalfetGetOrder/HalfetGetOrder"
os.makedirs(save_dir, exist_ok=True)

# 파일명: 예) godo_add_goods_map_20251029_153015.json
timestamp = time.strftime("%Y%m%d_%H%M%S")
save_path = os.path.join(save_dir, f"godo_add_goods.json")

# JSON 저장 (한글 깨짐 방지 + 예쁘게 포맷)
with open(save_path, "w", encoding="utf-8") as f:
    json.dump(add_goods_map, f, ensure_ascii=False, indent=2)

print(f"✅ 추가상품 딕셔너리 저장 완료: {save_path}")
