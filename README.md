# HalfetGetOrder
하프전자 주문수집

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


# pyinstaller --onefile --name HalfetGetOrder --icon=icon/app.ico entry.py