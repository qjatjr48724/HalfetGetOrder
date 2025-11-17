
import unicodedata, math
from datetime import datetime

def first_non_empty(*vals):
    for v in vals:
        if v is not None and str(v).strip() != "":
            return v
    return ""

def _to_int(x, default=0):
    try:
        return int(float(str(x).strip()))
    except Exception:
        return default

def _to_float(x, default=0.0):
    try:
        return float(str(x).strip())
    except Exception:
        return default

def _as_list(x):
    if x is None:
        return []
    return x if isinstance(x, list) else [x]

def _fmt_dt(s):
    if not s:
        return ""
    s = str(s).strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y.%m.%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y.%m.%d %H:%M")
        except Exception:
            pass
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").strftime("%Y.%m.%d 00:00")
    except Exception:
        return s

def visual_len(s):
    if s is None:
        return 0
    total = 0
    for ch in str(s):
        ea = unicodedata.east_asian_width(ch)
        total += 2 if ea in ('W','F','A') else 1
    return total

def get_box_count_from_items(items):
    if not items:
        return 1
    try:
        total_qty = sum(int(i.get('quantity', 1)) for i in items)
        if total_qty <= 3:
            return 1
        return 1 + math.ceil((total_qty - 3) / 4)
    except Exception:
        return 1
