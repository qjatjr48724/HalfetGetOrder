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

save_dir = r"C:/Users/UserK/Desktop/HalfetGetOrder/test"
os.makedirs(save_dir, exist_ok=True)

timestamp = time.strftime('%Y%m%d_%H%M%S')  # datetime 이름 충돌 피하기
json_path = os.path.join(save_dir, f"godo_orders.json")

with open(json_path, "w", encoding="utf-8") as f:
    json.dump(json_data, f, ensure_ascii=False, indent=2)

print(f"✅ 고도몰 json_data 저장 완료: {json_path}")

# (옵션) 원본 XML도 함께 저장하고 싶다면 아래 주석 해제
# xml_path = os.path.join(save_dir, f"godo_orders_{timestamp}.xml")
# with open(xml_path, "w", encoding=(godo_req.encoding or "utf-8")) as f:
#     f.write(data)
# print(f"✅ 고도몰 원본 XML 저장 완료: {xml_path}")