"""앱 데이터(견적서 표본 · 최근 다운로드 · 품목 카탈로그) 영구 저장소.

Streamlit Cloud 의 컨테이너 파일시스템은 재배포·재시작 때마다 초기화된다.
그래서 앱에서 저장한 표본과 히스토리가 새 버전을 배포할 때 사라졌다.

여기서는 GitHub 레포의 **별도 브랜치**(기본 `app-data`)에 JSON 으로 저장한다.
- 앱이 배포되는 브랜치(main)와 분리 → 데이터를 저장해도 재배포가 트리거되지 않음
- 컨테이너가 새로 떠도 브랜치에서 다시 읽어오므로 데이터가 유지됨

GITHUB_TOKEN 이 없으면(로컬 개발 등) 로컬 파일만 사용한다.
"""

from __future__ import annotations

import base64
import json
from pathlib import Path

import streamlit as st

# 로컬 미러 — 컨테이너가 살아있는 동안의 캐시 겸 개발 환경 저장소
_LOCAL_DIR_NAME = "_store"


def _conf() -> dict:
    """GitHub 접속 정보. 토큰이 없으면 token=None."""
    try:
        return {
            "token": st.secrets.get("GITHUB_TOKEN"),
            "owner": st.secrets.get("GITHUB_OWNER", "softment0je"),
            "repo": st.secrets.get("GITHUB_REPO", "QR_quote"),
            "code_branch": st.secrets.get("GITHUB_BRANCH", "main"),
            "branch": st.secrets.get("GITHUB_DATA_BRANCH", "app-data"),
        }
    except Exception:  # secrets.toml 자체가 없는 경우
        return {"token": None, "owner": "", "repo": "",
                "code_branch": "main", "branch": "app-data"}


def is_remote_enabled() -> bool:
    return bool(_conf()["token"])


def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _ensure_branch(cfg: dict) -> None:
    """데이터 브랜치가 없으면 코드 브랜치에서 만들어 둔다."""
    import requests
    base = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}"
    h = _headers(cfg["token"])
    r = requests.get(f"{base}/git/ref/heads/{cfg['branch']}", headers=h, timeout=10)
    if r.status_code == 200:
        return
    r = requests.get(f"{base}/git/ref/heads/{cfg['code_branch']}",
                     headers=h, timeout=10)
    if r.status_code != 200:
        return
    sha = r.json().get("object", {}).get("sha")
    if not sha:
        return
    requests.post(f"{base}/git/refs", headers=h, timeout=10,
                  json={"ref": f"refs/heads/{cfg['branch']}", "sha": sha})


@st.cache_data(ttl=120, show_spinner=False)
def _remote_get(rel_path: str, owner: str, repo: str,
                branch: str, token: str) -> dict | None:
    """데이터 브랜치에서 JSON 을 읽어온다. 없으면 None."""
    try:
        import requests
        url = (f"https://api.github.com/repos/{owner}/{repo}"
               f"/contents/{rel_path}")
        r = requests.get(url, headers=_headers(token),
                         params={"ref": branch}, timeout=10)
        if r.status_code != 200:
            return None
        content = r.json().get("content", "")
        raw = base64.b64decode(content).decode("utf-8")
        return json.loads(raw)
    except Exception:  # noqa: BLE001 — 네트워크/파싱 실패는 로컬로 폴백
        return None


def _local_path(project_root: Path, rel_path: str) -> Path:
    return project_root / "output" / _LOCAL_DIR_NAME / rel_path


def load(project_root: Path, rel_path: str, default: dict,
         repo_fallback: Path | None = None) -> dict:
    """저장된 JSON 을 읽는다.

    순서: 데이터 브랜치 → 로컬 미러 → repo_fallback(레포에 커밋된 원본) → default
    """
    cfg = _conf()
    if cfg["token"]:
        data = _remote_get(rel_path, cfg["owner"], cfg["repo"],
                           cfg["branch"], cfg["token"])
        if data is not None:
            return data
    local = _local_path(project_root, rel_path)
    if local.exists():
        try:
            return json.loads(local.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass
    if repo_fallback is not None and repo_fallback.exists():
        try:
            return json.loads(repo_fallback.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass
    return default


def save(project_root: Path, rel_path: str, data: dict,
         message: str) -> tuple[bool, str]:
    """JSON 을 로컬 미러 + 데이터 브랜치에 저장. (성공여부, 메시지) 반환."""
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")

    local = _local_path(project_root, rel_path)
    local.parent.mkdir(parents=True, exist_ok=True)
    try:
        local.write_bytes(payload)
    except OSError as e:
        return False, f"로컬 저장 실패: {e}"

    cfg = _conf()
    if not cfg["token"]:
        return False, "GITHUB_TOKEN 미설정 — 이 컨테이너에만 저장됨"

    try:
        import requests
    except ImportError:
        return False, "requests 모듈 없음"

    try:
        _ensure_branch(cfg)
        url = (f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}"
               f"/contents/{rel_path}")
        h = _headers(cfg["token"])
        r = requests.get(url, headers=h, params={"ref": cfg["branch"]},
                         timeout=10)
        sha = r.json().get("sha") if r.status_code == 200 else None
        body = {
            "message": message,
            "content": base64.b64encode(payload).decode("ascii"),
            "branch": cfg["branch"],
        }
        if sha:
            body["sha"] = sha
        r = requests.put(url, headers=h, json=body, timeout=15)
        _remote_get.clear()
        if r.status_code in (200, 201):
            return True, "GitHub 영구 저장 완료"
        return False, f"PUT {r.status_code}: {r.text[:120]}"
    except Exception as e:  # noqa: BLE001
        return False, f"네트워크 오류: {e}"
