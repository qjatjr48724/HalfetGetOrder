# HalfetGetOrder
하프전자 주문수집

고도몰 + 쿠팡 주문을 한 번에 수집해서

- **주문수집 엑셀**
- **대한통운 송장등록 엑셀**

두 가지 파일을 자동으로 생성해주는 Windows용 도구입니다.


## 기능
- 고도몰, 쿠팡 주문 API 연동
- 대한통운 송장등록용 엑셀 자동 생성
- 엑셀 기본 UI 개선

## 파일 트리구조
### 초기 파일 트리
```
HalfetGetOrder
├─ godo_add_goods_cache.json
├─ HalfetGetOrder
│  ├─ godo_add_goods.json
│  ├─ HalfetGetOrder.py
│  └─ keys.py
├─ README.md
├─ requirements.txt
└─ test
   ├─ godo_orders.json
   ├─ Goods_Add_Search.py
   ├─ HalfetGetOrder (1).py
   └─ makedict.py

```

### 수정 파일 트리
```
HalfetGetOrder
├─ data
│  ├─ ~$주문수집_20251117.xlsx
│  ├─ 대한통운 송장등록_20251117.xlsx
│  ├─ 주문수집_20251114.xlsx
│  └─ 주문수집_20251117.xlsx
├─ godo_add_goods_cache.json
├─ README.md
├─ requirements.txt
├─ src
│  └─ halfetgetorder
│     ├─ app.py
│     ├─ config.py
│     ├─ coupang.py
│     ├─ godo.py
│     ├─ io_excel.py
│     ├─ keys.py
│     ├─ resources
│     │  └─ godo_add_goods.json
│     ├─ utils.py
│     ├─ __init__.py
│     └─ __main__.py
└─ test
   ├─ godo_orders.json
   ├─ Goods_Add_Search.py
   ├─ HalfetGetOrder (1).py
   ├─ HalfetGetOrder.py
   └─ makedict.py

```

## 설치 및 실행

### 1. 파이썬으로 실행 (개발용)
cd HalfetGetOrder
.\venv\Scripts\activate
cd src
python -m halfetgetorder

### 2. EXE 파일 사용(배포용)


## 파일 배포
```
1. pyinstaller --onefile --name HalfetGetOrder --icon=icon/app.ico entry.py
2. .\build.bat
```