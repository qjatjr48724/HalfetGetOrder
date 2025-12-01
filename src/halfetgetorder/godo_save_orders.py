import os
import json
from datetime import datetime

from . import godo


def save_godo_orders_json(output_dir: str | None = None) -> str:
    """
    고도몰 주문을 조회해서 JSON 파일로 저장하는 유틸 함수.

    - godo.fetch_orders() 를 호출해서 주문 데이터를 가져오고
    - 프로젝트 최상위 폴더 기준 test/test_data 아래에
      godo_orders_YYYYMMDD_HHMMSS.json 형식으로 저장한다.
    - 저장된 파일 경로를 리턴.
    """

    # 1) 기본 저장 경로 계산
    if output_dir is None:
        # 이 파일 위치: .../src/halfetgetorder/godo_save_orders.py
        base_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.abspath(os.path.join(base_dir, "..", ".."))
        output_dir = os.path.join(project_root, "test", "test_data")

    os.makedirs(output_dir, exist_ok=True)

    # 2) 고도몰 주문 가져오기
    #    godo.fetch_orders() 는 날짜를 내부에서 처리하도록 그대로 사용
    godo_json = godo.fetch_orders()

    # 3) 파일명 만들기
    now = datetime.now()
    filename = f"godo_orders_{now:%Y%m%d_%H%M%S}.json"
    filepath = os.path.join(output_dir, filename)

    # 4) JSON 파일로 저장
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(godo_json, f, ensure_ascii=False, indent=2)

    print(f"✅ 고도몰 주문 JSON 저장 완료: {filepath}")
    return filepath


if __name__ == "__main__":
    # 모듈을 직접 실행했을 때 동작
    save_godo_orders_json()
