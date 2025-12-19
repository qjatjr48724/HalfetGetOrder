# update_keys.py
# 쿠팡 / 고도몰 API 키를 입력받아서 keys.py 와 .env 를 동시에 갱신하는 도구

from pathlib import Path
import getpass

ROOT = Path(__file__).resolve().parent
KEYS_PATH = ROOT / "keys.py"
ENV_PATH = ROOT / ".env"


def mask(v: str, front: int = 4, back: int = 4) -> str:
    """키값 일부만 보이도록 마스킹."""
    v = (v or "").strip()
    if not v:
        return "(미설정)"
    if len(v) <= front + back:
        return v[0:1] + "..."
    return f"{v[:front]}...{v[-back:]}"


def load_current_keys() -> dict:
    """
    keys.py 에서 현재 키 값 읽기.
    keys.py 는 아래 형태라고 가정:
        cp_accesskey = "..."
        cp_secretkey = "..."
        # 고도몰 API KEY
        partner_key = "..."
        godo_key = "..."
    """
    current = {
        "cp_accesskey": "",
        "cp_secretkey": "",
        "partner_key": "",
        "godo_key": "",
    }

    if not KEYS_PATH.exists():
        return current

    ns: dict = {}
    code = KEYS_PATH.read_text(encoding="utf-8")
    exec(code, ns, ns)

    for k in current.keys():
        v = ns.get(k, "")
        if isinstance(v, str):
            current[k] = v
    return current


def load_env_dict(path: Path) -> dict:
    """간단한 .env 파서 (key=value). 다른 설정 라인은 유지."""
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


def save_env_dict(path: Path, env: dict):
    lines = [f"{k}={v}" for k, v in env.items()]
    text = "\n".join(lines) + "\n"
    path.write_text(text, encoding="utf-8")


def main():
    print("────────────────────────────────────────────")
    print("  하프전자 주문수집기 - API 키 설정 도구")
    print("────────────────────────────────────────────")
    print(f"  keys.py : {KEYS_PATH}")
    print(f"  .env    : {ENV_PATH}")
    print()

    cur = load_current_keys()

    print("현재 설정된 키 (일부만 표시):")
    print(f"  쿠팡 ACCESSKEY : {mask(cur['cp_accesskey'])}")
    print(f"  쿠팡 SECRETKEY : {mask(cur['cp_secretkey'])}")
    print(f"  고도몰 PARTNER : {mask(cur['partner_key'])}")
    print(f"  고도몰 GODO    : {mask(cur['godo_key'])}")
    print()

    ans = input("▶ 키를 변경하시겠습니까? (Y/N, 엔터=아니오): ").strip().lower()
    if ans not in ("y", "yes"):
        print("기존 키를 유지하고 종료합니다.")
        return

    print()
    print("※ 새 값을 입력하지 않고 엔터만 치면 기존 값을 유지합니다.")
    print("※ 터미널에 키가 그대로 찍히지 않도록 getpass 로 입력받습니다.")
    print()

    new_cp_access = getpass.getpass("  쿠팡 ACCESSKEY : ").strip()
    new_cp_secret = getpass.getpass("  쿠팡 SECRETKEY : ").strip()
    new_partner   = getpass.getpass("  고도몰 PARTNER_KEY : ").strip()
    new_godo      = getpass.getpass("  고도몰 GODO_KEY    : ").strip()

    # 입력이 비어 있으면 기존 값 유지
    cp_access = new_cp_access or cur["cp_accesskey"]
    cp_secret = new_cp_secret or cur["cp_secretkey"]
    partner   = new_partner   or cur["partner_key"]
    godo      = new_godo      or cur["godo_key"]

    # 1) keys.py 갱신 (→ 빌드된 exe 안에도 이 값이 들어가게 됨)
    keys_content = f'''cp_accesskey = "{cp_access}"
cp_secretkey = "{cp_secret}"

# 고도몰 API KEY
partner_key = "{partner}"
godo_key = "{godo}"
'''
    KEYS_PATH.write_text(keys_content, encoding="utf-8")
    print()
    print("✅ keys.py 를 새 키값으로 저장했습니다.")

    # 2) .env 갱신 (로컬에서 디버깅/테스트용)
    env = load_env_dict(ENV_PATH)
    env["CP_ACCESSKEY"] = cp_access
    env["CP_SECRETKEY"] = cp_secret
    env["PARTNER_KEY"] = partner
    env["GODO_KEY"] = godo
    save_env_dict(ENV_PATH, env)
    print("✅ .env 파일에도 키값을 저장했습니다.")
    print()
    print("이제 build.bat 에서 빌드를 실행하면, 새 키가 포함된 exe 가 만들어집니다.")


if __name__ == "__main__":
    main()
