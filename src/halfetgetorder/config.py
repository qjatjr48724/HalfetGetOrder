
import os, sys
from pathlib import Path
from dotenv import load_dotenv

def _project_root() -> Path:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]

def _load_env():
    root = _project_root()
    env = root / ".env"
    if env.exists():
        load_dotenv(env)
    else:
        load_dotenv()
_load_env()

PARTNER_KEY = os.getenv("PARTNER_KEY", "")
GODO_KEY    = os.getenv("GODO_KEY", "")
CP_ACCESS   = os.getenv("CP_ACCESSKEY", "")
CP_SECRET   = os.getenv("CP_SECRETKEY", "")

try:
    if not (PARTNER_KEY and GODO_KEY and CP_ACCESS and CP_SECRET):
        from .keys import partner_key as _pk, godo_key as _gk, cp_accesskey as _ak, cp_secretkey as _sk
        PARTNER_KEY = PARTNER_KEY or _pk
        GODO_KEY    = GODO_KEY or _gk
        CP_ACCESS   = CP_ACCESS   or _ak
        CP_SECRET   = CP_SECRET   or _sk
except Exception:
    pass

def resource_path(rel: str) -> str:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base = Path(sys._MEIPASS) / "halfetgetorder" / "resources"
    else:
        base = Path(__file__).resolve().parent / "resources"
    return str(base / rel)

def app_data_dir() -> str:
    home = Path.home()
    desktop = home / "Desktop"
    # 바탕화면에 이름 변경하려면 밑줄 수정하기
    data = desktop / "하프전자 주문수집기"
    data.mkdir(parents=True, exist_ok=True)
    return str(data)

DATA_DIR = app_data_dir()
