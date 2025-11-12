
import os, json
from datetime import date
from .config import DATA_DIR
from . import godo, coupang
from .io_excel import (
    create_orders_sheet, finalize_orders_sheet,
    append_coupang_block, append_godo_sets
)

def main():
    cp_body = coupang.fetch_orders()
    godo_json = godo.fetch_orders()
    grouped = godo.group_sets(godo_json)

    wb2, ws2 = create_orders_sheet()

    try:
        resp_json = json.loads(cp_body) if cp_body else {}
        orders = resp_json.get('data') or resp_json.get('content', []) or []
        append_coupang_block(ws2, orders)
    except Exception as e:
        print("⚠️ 쿠팡 JSON 파싱 오류:", e)

    append_godo_sets(ws2, grouped)
    finalize_orders_sheet(ws2)

    today = date.today().strftime("%Y%m%d")
    order_xlsx = os.path.join(DATA_DIR, f"주문수집_{today}.xlsx")
    wb2.save(order_xlsx)
    print(f"✅ 엑셀 저장 완료: {order_xlsx}")
