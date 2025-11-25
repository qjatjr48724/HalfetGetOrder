# src/halfetgetorder/build_godo_add_goods_all.py

import os
import json
import requests
import xmltodict
from datetime import datetime

from .keys import partner_key, godo_key

BASE_URL = "https://openhub.godo.co.kr/godomall5/goods/Goods_Add_Search.php"


def _find_add_goods_items(node):
    """
    XML -> dict로 파싱된 전체 구조 안에서
    'addGoodsNo' 라는 키를 가진 dict들을 전부 찾아서 리스트로 반환.
    경로(data.return.addGoodsData) 를 몰라도 값만 보고 찾게 하기 위함.
    """
    results = []

    if isinstance(node, dict):
        # 이 dict 자체가 addGoodsNo 를 가진 '한 줄'일 수 있음
        if "addGoodsNo" in node:
            results.append(node)

        # 자식들도 재귀 탐색
        for v in node.values():
            results.extend(_find_add_goods_items(v))

    elif isinstance(node, list):
        for item in node:
            results.extend(_find_add_goods_items(item))

    return results


def fetch_add_goods_page(page: int = 1, size: int = 100, session: requests.Session | None = None):
    """
    Goods_Add_Search.php 한 페이지 호출해서 '추가상품' 리스트를 반환.
    XML 전체 구조 안에서 addGoodsNo 를 가진 dict들을 전부 찾아서 리스트로 만든다.
    """
    if session is None:
        session = requests.Session()

    params = {
        "partner_key": partner_key,
        "key": godo_key,
        "page": page,
        "size": size,
    }

    # 1차 시도: GET 방식 (검색 API는 대부분 GET을 많이 씀)
    resp = session.get(BASE_URL, params=params, timeout=30)

    # 인코딩 설정
    ctype = (resp.headers.get("Content-Type") or "").lower()
    if "euc-kr" in ctype or "cp949" in ctype:
        resp.encoding = "cp949"
    elif not resp.encoding:
        resp.encoding = "utf-8"

    text = resp.text.strip()

    if not text.startswith("<"):
        print("⚠️ XML 형식이 아닌 응답 (앞 300자):")
        print(text[:300])
        raise RuntimeError("고도몰에서 XML이 아닌 응답을 반환했습니다. (인증/파라미터/요청 방식 확인 필요)")

    # XML → dict
    data = xmltodict.parse(text)

    # 한 번만 구조를 찍어보고 싶으면 아래 주석을 잠깐 풀어서 확인해도 된다.
    # print(json.dumps(data, ensure_ascii=False, indent=2)[:1000])

    # 일반적인 openhub 형태면 최상위에 'data' 태그가 있고,
    # 그 아래에 header / return 이 있는 구조지만, 혹시 몰라서
    # header 체크는 '있으면' 하는 정도로만 사용.
    root = data.get("data") or data  # data 태그가 없으면 전체 dict 사용
    header = (root.get("header") or {}) if isinstance(root, dict) else {}
    code = header.get("code") or header.get("result")  # 환경에 따라 code 대신 result 일 수도 있음

    if code and code != "000" and str(code) != "1":
        msg = header.get("msg") or header.get("message") or "알 수 없는 오류"
        raise RuntimeError(f"고도몰 API 오류: code={code}, msg={msg}")

    # 🔍 핵심: 전체 구조에서 addGoodsNo 가 있는 dict들을 싹 모은다.
    items = _find_add_goods_items(data)

    # 디버그용 로그
    print(f"    → 이 페이지에서 발견한 추가상품 후보 개수: {len(items)}")

    return items


def fetch_all_add_goods(size: int = 200) -> list:
    """
    페이지를 돌며 모든 추가상품 리스트를 전부 모아 반환.
    """
    all_items: list = []
    page = 1

    with requests.Session() as session:
        while True:
            print(f"[INFO] 추가상품 페이지 조회: page={page}, size={size}")
            items = fetch_add_goods_page(page=page, size=size, session=session)

            if not items:
                print("[INFO] 더 이상 가져올 추가상품이 없습니다. 종료.")
                break

            all_items.extend(items)
            print(f"[INFO] 이번 페이지 {len(items)}건, 누적 {len(all_items)}건")

            # size보다 적게 오면 마지막 페이지라고 보고 종료
            if len(items) < size:
                print("[INFO] 마지막 페이지로 판단.")
                break

            page += 1

    return all_items


def build_add_goods_map(items: list) -> dict:
    """
    API에서 가져온 추가상품 리스트를

    {
      "1000000015": { "name": "고급 노트북 가방 구매", "summary": "" },
      "1000000096": { "name": "리브레오피스 설치",   "summary": "" },
      ...
    }

    이런 형태의 딕셔너리로 변환.
    """
    result: dict[str, dict] = {}

    for it in items:
        add_no = str(it.get("addGoodsNo") or "").strip()
        name = (it.get("goodsNm") or it.get("goodsNmStandard") or "").strip()

        if not add_no or not name:
            continue

        # 같은 번호가 여러 번 나와도 최초 한 번만 등록
        if add_no not in result:
            result[add_no] = {
                "name": name,
                "summary": ""  # 여기다 나중에 직접 B값(요약이름) 채우면 됨
            }

    return result


def save_add_goods_map(mapping: dict, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    print(f"✅ godo_add_goods_all.json 저장 완료: {output_path} (총 {len(mapping)}개)")


def main():
    # 이 스크립트 파일 기준으로 프로젝트 루트 경로 계산
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(current_dir, "..", ".."))

    output_path = os.path.join(project_root, "godo_add_goods_all.json")

    all_items = fetch_all_add_goods(size=200)
    mapping = build_add_goods_map(all_items)
    save_add_goods_map(mapping, output_path)


if __name__ == "__main__":
    main()
