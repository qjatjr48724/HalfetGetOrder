
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


# config.py 맨 아래쪽에 추가

def _mask_key(v: str, front: int = 4, back: int = 4) -> str:
    """키값을 화면에 보여줄 때 앞/뒤 일부만 보이게 마스킹."""
    v = str(v or "").strip()
    if not v:
        return "(미설정)"
    if len(v) <= front + back:
        return v[0:1] + "..."  # 너무 짧으면 대충만
    return f"{v[:front]}...{v[-back:]}"


def _load_env_file(path: Path) -> dict:
    """단순 .env 파서 (key=value). 중복 키는 나중 줄이 이김."""
    data: dict[str, str] = {}
    if not path.exists():
        return data
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        k, v = line.split("=", 1)
        data[k.strip()] = v.strip()
    return data


def _save_env_file(path: Path, env: dict):
    """env dict를 .env 형식으로 저장."""
    lines = []
    for k, v in env.items():
        lines.append(f"{k}={v}")
    text = "\n".join(lines) + "\n"
    path.write_text(text, encoding="utf-8")


def configure_api_keys_interactive():
    """
    콘솔에서 쿠팡/고도몰 API 키를 확인·변경하는 인터랙티브 설정 함수.

    - 현재 설정된 키를 일부만 보여주고
    - 변경 의사를 물어본 뒤
    - 새로 입력받은 값을 .env 에 저장하고
    - config 모듈의 전역 변수도 동시에 갱신한다.
    """
    global PARTNER_KEY, GODO_KEY, CP_ACCESS, CP_SECRET

    root = _project_root()
    env_path = root / ".env"

    print("────────────────────────────────────────────")
    print(" [API 키 설정] 쿠팡 / 고도몰 키 확인 및 변경")
    print("────────────────────────────────────────────")
    print(f" .env 위치: {env_path}")
    print()
    print(f"  고도몰 PARTNER_KEY : {_mask_key(PARTNER_KEY)}")
    print(f"  고도몰 GODO_KEY    : {_mask_key(GODO_KEY)}")
    print(f"  쿠팡 ACCESSKEY     : {_mask_key(CP_ACCESS)}")
    print(f"  쿠팡 SECRETKEY     : {_mask_key(CP_SECRET)}")
    print()

    ans = input(" ▶ API 키를 변경하시겠습니까? (Y/N, 엔터=아니오): ").strip().lower()
    if ans not in ("y", "yes"):
        print(" - 기존 설정을 그대로 사용합니다.")
        print()
        return

    print()
    print("※ 새 값을 입력하지 않고 그냥 엔터만 누르면 기존 값을 유지합니다.")
    print()

    new_partner = input("  고도몰 PARTNER_KEY : ").strip()
    new_godo    = input("  고도몰 GODO_KEY    : ").strip()
    new_access  = input("  쿠팡 ACCESSKEY     : ").strip()
    new_secret  = input("  쿠팡 SECRETKEY     : ").strip()

    if new_partner:
        PARTNER_KEY = new_partner
    if new_godo:
        GODO_KEY = new_godo
    if new_access:
        CP_ACCESS = new_access
    if new_secret:
        CP_SECRET = new_secret

    # 기존 .env 내용 읽어서 덮어쓰기
    env = _load_env_file(env_path)
    env["PARTNER_KEY"] = PARTNER_KEY
    env["GODO_KEY"] = GODO_KEY
    env["CP_ACCESSKEY"] = CP_ACCESS
    env["CP_SECRETKEY"] = CP_SECRET

    try:
        _save_env_file(env_path, env)
        print()
        print(" ✅ .env 파일에 API 키를 저장했습니다.")
        print("    다음 실행부터 변경된 값으로 사용됩니다.")
    except Exception as e:
        print()
        print(" ⚠️ .env 파일 저장 중 오류가 발생했습니다:", e)

    print()

