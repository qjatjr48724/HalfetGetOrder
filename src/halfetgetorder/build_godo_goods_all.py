# src/halfetgetorder/build_godo_goods_all.py

import os
import json
import requests
import xmltodict

from .keys import partner_key, godo_key

# 고도몰 본상품 검색 API
BASE_URL = "https://openhub.godo.co.kr/godomall5/goods/Goods_Search.php"


def _find_goods_items(node):
    """
    XML -> dict 구조 전체에서 'goodsNo' 를 가진 dict들을 전부 찾아서 리스트로 반환.
    (응답 구조를 정확히 몰라도 goodsNo 기준으로 상품 레코드만 모을 수 있게)
    """
    results = []

    if isinstance(node, dict):
        if "goodsNo" in node:
            results.append(node)

        for v in node.values():
            results.extend(_find_goods_items(v))

    elif isinstance(node, list):
        for item in node:
            results.extend(_find_goods_items(item))

    return results


def fetch_goods_page(page: int = 1, size: int = 100, session: requests.Session | None = None):
    """
    Goods_Search.php 한 페이지 호출해서 '상품' 리스트를 반환.
    XML 전체 구조 안에서 goodsNo 를 가진 dict들을 전부 찾아서 리스트로 만든다.
    """
    if session is None:
        session = requests.Session()

    params = {
        "partner_key": partner_key,
        "key": godo_key,
        "page": page,
        "size": size,
        # 필요하면 여기서 필터 추가 가능:
        # "scmNo": ...,
        # "goodsNo": ...,
        # "goodsNm": ...,
        # "brandCd": ...,
    }

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
        raise RuntimeError("고도몰 상품 검색 API 응답이 XML 형식이 아닙니다. (인증/파라미터 확인 필요)")

    data = xmltodict.parse(text)

    # 일반적인 openhub 패턴: <data><header>...</header><return>...</return></data>
    root = data.get("data") or data
    header = (root.get("header") or {}) if isinstance(root, dict) else {}
    code = header.get("code") or header.get("result")

    # code가 존재한다면 000 또는 1 이 정상일 확률이 높음
    if code and code not in ("000", "1"):
        msg = header.get("msg") or header.get("message") or "알 수 없는 오류"
        raise RuntimeError(f"고도몰 상품 검색 API 오류: code={code}, msg={msg}")

    # 전체 구조에서 goodsNo 를 가진 dict들만 추출
    items = _find_goods_items(data)
    print(f"    → 이 페이지에서 발견한 상품 후보 개수: {len(items)}")

    return items


def fetch_all_goods(size: int = 200) -> list:
    """
    페이지를 돌며 모든 상품 리스트를 전부 모아 반환.
    """
    all_items: list = []
    page = 1

    with requests.Session() as session:
        while True:
            print(f"[INFO] 상품 페이지 조회: page={page}, size={size}")
            items = fetch_goods_page(page=page, size=size, session=session)

            if not items:
                print("[INFO] 더 이상 가져올 상품이 없습니다. 종료.")
                break

            all_items.extend(items)
            print(f"[INFO] 이번 페이지 {len(items)}건, 누적 {len(all_items)}건")

            # size보다 적게 오면 마지막 페이지라고 보고 종료
            if len(items) < size:
                print("[INFO] 마지막 페이지로 판단.")
                break

            page += 1

    return all_items


def build_goods_map(items: list) -> dict:
    """
    API에서 가져온 상품 리스트를

    {
      "1000001824": {  <-- goodsNo
         ... goods_search 응답 내 이 상품에 대한 모든 필드 ...
      },
      "1000002110": {
         ... 전체 필드 ...
      },
      ...
    }

    이런 형태의 딕셔너리로 변환.

    ✨ 응답 파라미터를 하나도 버리지 않고,
       goodsNo 기준으로만 1차 정리하는 게 목적.
    """
    result: dict[str, dict] = {}

    for it in items:
        goods_no = str(it.get("goodsNo") or "").strip()
        if not goods_no:
            continue

        # 이미 등록된 goodsNo이면 스킵 (중복 방지)
        if goods_no in result:
            continue

        # 여기서 it 전체를 그대로 저장 → 응답 모든 필드 유지
        result[goods_no] = it

    return result


def save_goods_map(mapping: dict, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    print(f"✅ godo_goods_all.json 저장 완료: {output_path} (총 {len(mapping)}개)")


def main():
    # 이 스크립트 파일 기준으로 프로젝트 루트 경로 계산
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(current_dir, "..", ".."))

    output_path = os.path.join(project_root, "godo_goods_all.json")

    all_items = fetch_all_goods(size=200)
    mapping = build_goods_map(all_items)
    save_goods_map(mapping, output_path)


if __name__ == "__main__":
    main()
