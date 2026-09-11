"""견적서 자동 생성 웹 인터페이스 (Streamlit).

`python -m src.cli web` 로 실행하면 브라우저가 자동으로 열립니다.

페이지 구성:
  📋 견적서 작성       — 폼 입력 → DOCX/PDF 다운로드
  📦 품목 관리     — 상품 추가/수정/삭제
  ⚙ 설정              — 브랜드 정보, 양식 라벨/문구 편집
"""
from __future__ import annotations

import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path

# Streamlit이 이 파일을 단독 스크립트로 실행하므로,
# 상위 폴더(src의 부모)를 path에 추가해 `src.xxx` 형태로 import 가능하게 함
_PARENT = Path(__file__).resolve().parent.parent
if str(_PARENT) not in sys.path:
    sys.path.insert(0, str(_PARENT))

import pandas as pd
import streamlit as st

from src.labels import DocumentLabels, load_labels
from src.loader import load_brand
from src.models import (
    BankAccount,
    Brand,
    BrandColors,
    BrandingAssets,
    CompanyInfo,
    ContactPerson,
    Counterparty,
    LineItem,
    QuoteDocument,
    SignatureInfo,
    Totals,
)
from src.membership_models import (
    MembershipCategory,
    MembershipLineItem,
    MembershipParty,
    MembershipQuoteDocument,
    MembershipScenario,
    MembershipSection,
    MembershipSubItem,
    category_subtotal,
    scenario_grand_total_by_period,
    scenario_vat_and_total,
    section_subtotals_by_period,
)
from src.membership_renderer import render_membership_docx
from src.pdf_converter import convert_docx_to_pdf, find_soffice
from src.renderer import render_docx


def _detect_project_root() -> Path:
    """프로젝트 루트 자동 감지 (CLI/로컬/클라우드 환경 모두 호환)."""
    # 1) CLI가 환경변수로 명시했으면 우선
    if env := os.environ.get("CONTRACT_SYSTEM_ROOT"):
        return Path(env).resolve()
    # 2) 현재 파일이 <root>/src/webapp.py 형태인지 확인
    here = Path(__file__).resolve().parent
    if here.name == "src" and (here.parent / "brands").exists():
        return here.parent
    # 3) 작업 디렉터리에 brands/ 가 있으면 거기
    cwd = Path.cwd().resolve()
    if (cwd / "brands").exists():
        return cwd
    # 4) 최종 fallback
    return here.parent


PROJECT_ROOT = _detect_project_root()


# ═════════════════════════════════════════════════════════════
# 자동저장 (새로고침/이탈 보호)
# ═════════════════════════════════════════════════════════════

AUTOSAVE_DIR = PROJECT_ROOT / "output" / "_autosave"
QR_AUTOSAVE_PATH = AUTOSAVE_DIR / "qr_quote.json"
MC_AUTOSAVE_PATH = AUTOSAVE_DIR / "membership_quote.json"

QR_HISTORY_DIR = PROJECT_ROOT / "output" / "_history" / "qr"
QR_TEMPLATE_DIR = PROJECT_ROOT / "output" / "_templates" / "qr"
MC_HISTORY_DIR = PROJECT_ROOT / "output" / "_history" / "mc"
MC_TEMPLATE_DIR = PROJECT_ROOT / "output" / "_templates" / "mc"
HISTORY_LIMIT = 10

# 저장/표시 시각은 한국 표준시(KST, UTC+9) 기준 — Streamlit Cloud 가 UTC 라
# datetime.now() 결과를 그대로 쓰면 9시간 어긋남
from datetime import timezone, timedelta as _timedelta
KST = timezone(_timedelta(hours=9))


def _now_kst():
    from datetime import datetime
    return datetime.now(KST)

# 자동 저장/복원할 위젯 키들
QR_FORM_KEYS = [
    "issuer_name", "issuer_phone", "issuer_title", "issuer_email",
    "cp_name", "cp_reg", "cp_address", "cp_contact_name",
    "cp_contact_title", "cp_email", "subject", "notes",
    "total_discount_pct",
]


def _read_json_safe(path: Path) -> dict | None:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _write_json_safe(path: Path, payload: dict) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, default=str),
                        encoding="utf-8")
    except OSError:
        pass


def _qr_autosave_load_once() -> None:
    """페이지 진입 시 1회만 호출 — 디스크의 autosave를 session_state에 채워넣음."""
    if st.session_state.get("_qr_autosave_loaded"):
        return
    st.session_state["_qr_autosave_loaded"] = True
    payload = _read_json_safe(QR_AUTOSAVE_PATH)
    if not payload:
        return
    items = payload.get("items")
    if items:
        try:
            # 이전 버전 호환: '💰 할인' → '💰 할인행'
            for it in items:
                if it.get("분류") == "💰 할인":
                    it["분류"] = "💰 할인행"
            st.session_state["items_df"] = pd.DataFrame(items)
        except (ValueError, KeyError):
            pass
    for k in QR_FORM_KEYS:
        v = payload.get(k)
        if v not in (None, "") and k not in st.session_state:
            st.session_state[k] = v


def _qr_autosave_write() -> None:
    df = st.session_state.get("items_df")
    payload = {k: st.session_state.get(k) for k in QR_FORM_KEYS}
    if df is not None and not df.empty:
        payload["items"] = df.to_dict(orient="records")
    has_any = (df is not None and not df.empty) or any(payload.get(k) for k in QR_FORM_KEYS)
    if has_any:
        _write_json_safe(QR_AUTOSAVE_PATH, payload)
    else:
        # 비어있으면 autosave 파일 제거
        try:
            QR_AUTOSAVE_PATH.unlink(missing_ok=True)
        except OSError:
            pass


def _qr_autosave_clear() -> None:
    try:
        QR_AUTOSAVE_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def _mc_autosave_load_once() -> None:
    if st.session_state.get("_mc_autosave_loaded"):
        return
    st.session_state["_mc_autosave_loaded"] = True
    payload = _read_json_safe(MC_AUTOSAVE_PATH)
    if not payload:
        return
    if "mc_doc" not in st.session_state and payload.get("mc_doc"):
        st.session_state["mc_doc"] = payload["mc_doc"]
    if "mc_items_df" not in st.session_state and payload.get("mc_items"):
        try:
            items = payload["mc_items"]
            for it in items:
                if it.get("분류") == "💰 할인":
                    it["분류"] = "💰 할인행"
            st.session_state["mc_items_df"] = pd.DataFrame(items)
        except (ValueError, KeyError):
            pass
    for k in ("mc_issuer_name", "mc_issuer_title",
              "mc_issuer_phone", "mc_issuer_email"):
        v = payload.get(k)
        if v not in (None, "") and k not in st.session_state:
            st.session_state[k] = v


def _mc_autosave_write() -> None:
    doc = st.session_state.get("mc_doc")
    df = st.session_state.get("mc_items_df")
    payload = {
        "mc_doc": doc,
        "mc_items": (df.to_dict(orient="records")
                     if df is not None and not df.empty else None),
        "mc_issuer_name": st.session_state.get("mc_issuer_name"),
        "mc_issuer_title": st.session_state.get("mc_issuer_title"),
        "mc_issuer_phone": st.session_state.get("mc_issuer_phone"),
        "mc_issuer_email": st.session_state.get("mc_issuer_email"),
    }
    has_any = bool(doc) or bool(payload.get("mc_items")) or any(
        payload.get(k) for k in ("mc_issuer_name", "mc_issuer_title",
                                  "mc_issuer_phone", "mc_issuer_email")
    )
    if has_any:
        _write_json_safe(MC_AUTOSAVE_PATH, payload)
    else:
        try:
            MC_AUTOSAVE_PATH.unlink(missing_ok=True)
        except OSError:
            pass


def _mc_autosave_clear() -> None:
    try:
        MC_AUTOSAVE_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def _qr_snapshot_payload() -> dict:
    """현재 견적서 입력 상태 전체를 직렬화 가능한 dict로."""
    df = st.session_state.get("items_df")
    payload = {k: st.session_state.get(k) for k in QR_FORM_KEYS}
    if df is not None and not df.empty:
        payload["items"] = df.to_dict(orient="records")
    return payload


def _qr_apply_snapshot(payload: dict) -> None:
    """저장된 견적서 payload를 session_state에 복원."""
    items = payload.get("items")
    if items:
        try:
            st.session_state["items_df"] = pd.DataFrame(items)
        except (ValueError, KeyError):
            pass
    for k in QR_FORM_KEYS:
        v = payload.get(k)
        if v not in (None, ""):
            st.session_state[k] = v


def _qr_save_history(snapshot: dict, document_id: str) -> None:
    """견적서 생성/미리보기 직후 히스토리에 저장.

    같은 document_id 인 기존 파일은 모두 제거 → DOCX/PDF 또는 미리보기/생성에서
    중복으로 쌓이지 않고 항상 1개만 유지 (최신 입력 반영).
    """
    QR_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    from datetime import datetime
    safe_id = "".join(c for c in document_id if c.isalnum() or c in "-_")[:40]

    # 같은 document_id 기존 파일 모두 제거 (중복 방지)
    for old in QR_HISTORY_DIR.glob(f"*_{safe_id}.json"):
        try:
            old.unlink(missing_ok=True)
        except OSError:
            pass

    ts = _now_kst().strftime("%Y%m%d_%H%M%S")
    payload = {
        **snapshot,
        "_document_id": document_id,
        "_saved_at": _now_kst().isoformat(timespec="seconds"),
        "_subject": st.session_state.get("subject", ""),
        "_cp_name": st.session_state.get("cp_name", ""),
    }
    path = QR_HISTORY_DIR / f"{ts}_{safe_id}.json"
    _write_json_safe(path, payload)
    # 오래된 히스토리 정리 (HISTORY_LIMIT 초과분 삭제)
    files = sorted(QR_HISTORY_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    for old in files[HISTORY_LIMIT:]:
        try:
            old.unlink()
        except OSError:
            pass


def _qr_list_history() -> list[tuple[Path, dict]]:
    if not QR_HISTORY_DIR.exists():
        return []
    files = sorted(QR_HISTORY_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)[:HISTORY_LIMIT]
    result = []
    for p in files:
        data = _read_json_safe(p)
        if data:
            result.append((p, data))
    return result


def _qr_save_template(name: str, snapshot: dict) -> bool:
    name = (name or "").strip()
    if not name:
        return False
    QR_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    safe = "".join(c for c in name if c.isalnum() or c in " -_가-힣")[:60].strip()
    if not safe:
        return False
    from datetime import datetime
    payload = {
        **snapshot,
        "_template_name": name,
        "_saved_at": _now_kst().isoformat(timespec="seconds"),
    }
    path = QR_TEMPLATE_DIR / f"{safe}.json"
    _write_json_safe(path, payload)
    return True


def _qr_list_templates() -> list[tuple[Path, dict]]:
    if not QR_TEMPLATE_DIR.exists():
        return []
    files = sorted(QR_TEMPLATE_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    result = []
    for p in files:
        data = _read_json_safe(p)
        if data:
            result.append((p, data))
    return result


def _qr_delete_template(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


# ── 멤버십 견적서 — 최근 다운로드 / 표본 헬퍼 ──

def _mc_snapshot_payload() -> dict:
    """현재 멤버십 견적서 입력 상태 전체."""
    doc = st.session_state.get("mc_doc")
    df = st.session_state.get("mc_items_df")
    return {
        "mc_doc": doc,
        "mc_items": (df.to_dict(orient="records")
                     if df is not None and not df.empty else None),
        "mc_issuer_name": st.session_state.get("mc_issuer_name"),
        "mc_issuer_title": st.session_state.get("mc_issuer_title"),
        "mc_issuer_phone": st.session_state.get("mc_issuer_phone"),
        "mc_issuer_email": st.session_state.get("mc_issuer_email"),
    }


def _mc_apply_snapshot(payload: dict) -> None:
    if payload.get("mc_doc"):
        st.session_state["mc_doc"] = payload["mc_doc"]
    if payload.get("mc_items"):
        try:
            items = payload["mc_items"]
            for it in items:
                if it.get("분류") == "💰 할인":
                    it["분류"] = "💰 할인행"
            st.session_state["mc_items_df"] = pd.DataFrame(items)
        except (ValueError, KeyError):
            pass
    for k in ("mc_issuer_name", "mc_issuer_title",
              "mc_issuer_phone", "mc_issuer_email"):
        v = payload.get(k)
        if v not in (None, ""):
            st.session_state[k] = v


def _mc_save_history(snapshot: dict, document_id: str) -> None:
    """멤버십 견적서 생성/미리보기 시 히스토리에 저장 (같은 document_id 면 1개로 유지)."""
    MC_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    from datetime import datetime
    safe_id = "".join(c for c in document_id if c.isalnum() or c in "-_")[:40]
    for old in MC_HISTORY_DIR.glob(f"*_{safe_id}.json"):
        try:
            old.unlink(missing_ok=True)
        except OSError:
            pass
    ts = _now_kst().strftime("%Y%m%d_%H%M%S")
    doc = (snapshot.get("mc_doc") or {})
    cp_name = (doc.get("counterparty") or {}).get("name") or ""
    title = doc.get("title") or "멤버십 클라우드 견적서"
    payload = {
        **snapshot,
        "_document_id": document_id,
        "_saved_at": _now_kst().isoformat(timespec="seconds"),
        "_subject": title,
        "_cp_name": cp_name,
    }
    path = MC_HISTORY_DIR / f"{ts}_{safe_id}.json"
    _write_json_safe(path, payload)
    files = sorted(MC_HISTORY_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    for old in files[HISTORY_LIMIT:]:
        try:
            old.unlink()
        except OSError:
            pass


def _mc_list_history() -> list[tuple[Path, dict]]:
    if not MC_HISTORY_DIR.exists():
        return []
    files = sorted(MC_HISTORY_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)[:HISTORY_LIMIT]
    result = []
    for p in files:
        data = _read_json_safe(p)
        if data:
            result.append((p, data))
    return result


def _mc_save_template(name: str, snapshot: dict) -> bool:
    name = (name or "").strip()
    if not name:
        return False
    MC_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    safe = "".join(c for c in name if c.isalnum() or c in " -_가-힣")[:60].strip()
    if not safe:
        return False
    from datetime import datetime
    payload = {
        **snapshot,
        "_template_name": name,
        "_saved_at": _now_kst().isoformat(timespec="seconds"),
    }
    path = MC_TEMPLATE_DIR / f"{safe}.json"
    _write_json_safe(path, payload)
    return True


def _mc_list_templates() -> list[tuple[Path, dict]]:
    if not MC_TEMPLATE_DIR.exists():
        return []
    files = sorted(MC_TEMPLATE_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    result = []
    for p in files:
        data = _read_json_safe(p)
        if data:
            result.append((p, data))
    return result


def _mc_delete_template(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _render_mc_recent_panel() -> None:
    """📂 멤버십 최근 다운로드 견적서."""
    history = _mc_list_history()
    label = (f"📂 최근 다운로드 견적서 ({len(history)})"
             if history else "📂 최근 다운로드 견적서")
    with st.expander(label, expanded=False):
        st.caption(
            "견적서 **생성** 또는 **미리보기** 를 누르면 자동으로 **최근 10개**가 보관됩니다. "
            "같은 견적서를 DOCX·PDF 둘 다 받아도 한 건으로만 남고, "
            "다시 불러와 이어서 편집할 수 있어요."
        )
        if not history:
            st.caption("아직 생성/미리보기한 견적서가 없습니다.")
            return
        delete_path = None
        for i, (path, data) in enumerate(history):
            saved = (data.get("_saved_at") or "")[:16].replace("T", " ")
            subj = data.get("_subject") or "(건명 없음)"
            cp = data.get("_cp_name") or "(수신처 없음)"
            with st.container(border=True):
                rcols = st.columns([6, 1.2, 1.2])
                with rcols[0]:
                    st.markdown(
                        f"🕐 {saved}<br>"
                        f"<span style='color:#555'>**{cp}**  ·  {subj}</span>",
                        unsafe_allow_html=True,
                    )
                with rcols[1]:
                    if st.button("📥 불러오기",
                                 key=f"mc_hist_load_{i}",
                                 use_container_width=True):
                        _mc_apply_snapshot(data)
                        st.success("✅ 견적서를 불러왔습니다.")
                        st.rerun()
                with rcols[2]:
                    if st.button("🗑 삭제",
                                 key=f"mc_hist_del_{i}",
                                 use_container_width=True):
                        delete_path = path
        if delete_path is not None:
            try:
                delete_path.unlink(missing_ok=True)
            except OSError:
                pass
            st.rerun()


def _render_mc_template_panel() -> None:
    """📋 멤버십 견적서 표본."""
    templates = _mc_list_templates()
    label = (f"📋 견적서 표본 ({len(templates)})"
             if templates else "📋 견적서 표본")
    with st.expander(label, expanded=False):
        st.caption(
            "**표본**은 자주 쓰는 멤버십 견적서 형태를 수기 저장해 두는 공간입니다. "
            "정석 케이스를 미리 만들어 두고, 새 견적서 작성할 때 끌어와서 빠르게 시작하세요."
        )
        save_col1, save_col2 = st.columns([3, 2])
        with save_col1:
            tpl_name = st.text_input(
                "표본 이름",
                placeholder="비우면 견적서 '제목'을 자동 사용",
                key="mc_tpl_name", label_visibility="collapsed",
            )
        with save_col2:
            if st.button("💾 현재 입력을 표본으로 저장",
                         use_container_width=True, key="mc_tpl_save"):
                # 이름 비어있으면 현재 견적서 제목 자동 사용
                mc_doc = st.session_state.get("mc_doc") or {}
                effective_name = (tpl_name or "").strip() or (
                    mc_doc.get("title") or ""
                ).strip()
                if _mc_save_template(effective_name, _mc_snapshot_payload()):
                    st.success(f"✅ 표본 저장: {effective_name}")
                    st.rerun()
                else:
                    st.warning(
                        "표본 이름이 비어있고 견적서 '제목'도 비어있어 저장할 수 없습니다. "
                        "둘 중 하나는 입력해 주세요."
                    )
        st.divider()
        if not templates:
            st.caption("저장된 표본이 없습니다.")
            return
        delete_path = None
        for i, (path, data) in enumerate(templates):
            name = data.get("_template_name") or path.stem
            saved = (data.get("_saved_at") or "")[:10]
            with st.container(border=True):
                rcols = st.columns([6, 1.2, 1.2])
                with rcols[0]:
                    st.markdown(
                        f"📋 **{name}**<br>"
                        f"<span style='color:#777;font-size:0.85rem'>저장일 {saved}</span>",
                        unsafe_allow_html=True,
                    )
                with rcols[1]:
                    if st.button("📥 불러오기",
                                 key=f"mc_tpl_load_{i}",
                                 use_container_width=True):
                        _mc_apply_snapshot(data)
                        st.success(f"✅ 표본 '{name}' 을 불러왔습니다.")
                        st.rerun()
                with rcols[2]:
                    if st.button("🗑 삭제",
                                 key=f"mc_tpl_del_{i}",
                                 use_container_width=True):
                        delete_path = path
        if delete_path is not None:
            _mc_delete_template(delete_path)
            st.rerun()


def _inject_beforeunload(active: bool) -> None:
    """작성 중 내용이 있으면 페이지 이탈 시 브라우저 경고를 띄움."""
    from streamlit.components.v1 import html
    flag = "1" if active else "0"
    html(f"""
<script>
(function() {{
  try {{
    var w = window.parent;
    if (!w._qrUnloadHandler) {{
      w._qrUnloadHandler = function(e) {{
        if (w._qrUnloadActive) {{
          e.preventDefault();
          e.returnValue = '작성 중인 내용이 있어요. 정말 이 페이지를 떠나시겠어요?';
          return e.returnValue;
        }}
      }};
      w.addEventListener('beforeunload', w._qrUnloadHandler);
    }}
    w._qrUnloadActive = ({flag} === 1);
  }} catch (err) {{ /* iframe cross-origin 등 무시 */ }}
}})();
</script>
""", height=0)


def _render_qr_recent_panel() -> None:
    """📂 최근 다운로드 견적서 — 생성/미리보기한 최근 10개를 불러와 이어서 편집."""
    history = _qr_list_history()
    label = (f"📂 최근 다운로드 견적서 ({len(history)})"
             if history else "📂 최근 다운로드 견적서")
    with st.expander(label, expanded=False):
        st.caption(
            "견적서 **생성** 또는 **미리보기** 를 누르면 자동으로 **최근 10개**가 보관됩니다. "
            "같은 견적서를 DOCX·PDF 둘 다 받아도 한 건으로만 남고, "
            "오타나 누락이 있을 때 다시 불러와 이어서 편집할 수 있어요."
        )
        if not history:
            st.caption("아직 생성/미리보기한 견적서가 없습니다.")
            return

        delete_path = None
        for i, (path, data) in enumerate(history):
            saved = (data.get("_saved_at") or "")[:16].replace("T", " ")
            subj = data.get("_subject") or "(건명 없음)"
            cp = data.get("_cp_name") or "(수신처 없음)"
            with st.container(border=True):
                rcols = st.columns([6, 1.2, 1.2])
                with rcols[0]:
                    st.markdown(
                        f"🕐 {saved}<br>"
                        f"<span style='color:#555'>**{cp}**  ·  {subj}</span>",
                        unsafe_allow_html=True,
                    )
                with rcols[1]:
                    if st.button("📥 불러오기",
                                 key=f"qr_hist_load_{i}",
                                 use_container_width=True):
                        _qr_apply_snapshot(data)
                        st.success("✅ 견적서를 불러왔습니다.")
                        st.rerun()
                with rcols[2]:
                    if st.button("🗑 삭제",
                                 key=f"qr_hist_del_{i}",
                                 use_container_width=True,
                                 help="이 견적서 히스토리만 삭제 (DOCX/PDF 파일은 그대로)"):
                        delete_path = path
        if delete_path is not None:
            try:
                delete_path.unlink(missing_ok=True)
            except OSError:
                pass
            st.rerun()


def _render_qr_template_panel() -> None:
    """📋 견적서 표본 — 정석/반복용 케이스를 수기 저장해 빠르게 끌어와 시작."""
    templates = _qr_list_templates()
    label = f"📋 견적서 표본 ({len(templates)})" if templates else "📋 견적서 표본"
    with st.expander(label, expanded=False):
        st.caption(
            "**표본**은 자주 쓰는 견적서 형태를 수기 저장해 두는 공간입니다. "
            "정석 케이스를 미리 만들어 두고, 새 견적서 작성할 때 끌어와서 빠르게 시작하세요."
        )
        save_col1, save_col2 = st.columns([3, 2])
        with save_col1:
            tpl_name = st.text_input(
                "표본 이름",
                placeholder="비우면 견적서 '건명'을 자동 사용",
                key="qr_tpl_name", label_visibility="collapsed",
            )
        with save_col2:
            if st.button("💾 현재 입력을 표본으로 저장",
                         use_container_width=True, key="qr_tpl_save"):
                # 이름 비어있으면 견적서 건명을 자동 사용
                effective_name = (tpl_name or "").strip() or (
                    st.session_state.get("subject") or ""
                ).strip()
                if _qr_save_template(effective_name, _qr_snapshot_payload()):
                    st.success(f"✅ 표본 저장: {effective_name}")
                    st.rerun()
                else:
                    st.warning(
                        "표본 이름이 비어있고 견적서 '건명'도 비어있어 저장할 수 없습니다. "
                        "둘 중 하나는 입력해 주세요."
                    )
        st.divider()
        if not templates:
            st.caption("저장된 표본이 없습니다.")
            return

        delete_path = None
        for i, (path, data) in enumerate(templates):
            name = data.get("_template_name") or path.stem
            saved = (data.get("_saved_at") or "")[:10]
            with st.container(border=True):
                rcols = st.columns([6, 1.2, 1.2])
                with rcols[0]:
                    st.markdown(
                        f"📋 **{name}**<br>"
                        f"<span style='color:#777;font-size:0.85rem'>저장일 {saved}</span>",
                        unsafe_allow_html=True,
                    )
                with rcols[1]:
                    if st.button("📥 불러오기",
                                 key=f"qr_tpl_load_{i}",
                                 use_container_width=True):
                        _qr_apply_snapshot(data)
                        st.success(f"✅ 표본 '{name}' 을 불러왔습니다.")
                        st.rerun()
                with rcols[2]:
                    if st.button("🗑 삭제",
                                 key=f"qr_tpl_del_{i}",
                                 use_container_width=True):
                        delete_path = path
        if delete_path is not None:
            _qr_delete_template(delete_path)
            st.rerun()


# ═════════════════════════════════════════════════════════════
# 데이터 로더 / 저장 헬퍼
# ═════════════════════════════════════════════════════════════

# 카탈로그 종류별 경로
CATALOG_FILES = {
    "qr": "products.json",                  # 일반 QR오더 견적기
    "outdoor": "qr_outdoor_products.json",  # [야외형] QR오더 견적기
}


@st.cache_data
def _load_products_for(catalog_kind: str = "qr") -> list[dict]:
    fname = CATALOG_FILES.get(catalog_kind, "products.json")
    path = PROJECT_ROOT / "catalog" / fname
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8")).get("products", [])


@st.cache_data
def _load_products() -> list[dict]:
    """기존 호환 — 일반 QR 카탈로그."""
    return _load_products_for("qr")


def _push_to_github(file_path: str, content_bytes: bytes,
                    message: str) -> tuple[bool, str]:
    """GitHub repo 의 파일을 자동 commit + push.
    Streamlit secrets 에 GITHUB_TOKEN 이 없으면 (False, 사유) 반환.
    """
    import base64
    try:
        import requests
    except ImportError:
        return False, "requests 모듈 없음"
    try:
        token = st.secrets.get("GITHUB_TOKEN")
        owner = st.secrets.get("GITHUB_OWNER", "jieunpark322")
        repo = st.secrets.get("GITHUB_REPO", "QR_quote")
        branch = st.secrets.get("GITHUB_BRANCH", "main")
    except Exception:
        return False, "secrets 접근 실패"
    if not token:
        return False, "GITHUB_TOKEN 미설정 (수동 백업 모드)"
    api = f"https://api.github.com/repos/{owner}/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    try:
        resp = requests.get(api, headers=headers, params={"ref": branch}, timeout=10)
        sha = resp.json().get("sha") if resp.status_code == 200 else None
        payload = {
            "message": message,
            "content": base64.b64encode(content_bytes).decode("ascii"),
            "branch": branch,
        }
        if sha:
            payload["sha"] = sha
        resp = requests.put(api, headers=headers, json=payload, timeout=15)
        if resp.status_code in (200, 201):
            return True, "GitHub 영구 저장 완료"
        return False, f"PUT {resp.status_code}: {resp.text[:120]}"
    except Exception as e:  # noqa: BLE001
        return False, f"네트워크 오류: {e}"


def _save_products_for(catalog_kind: str, products: list[dict]) -> None:
    fname = CATALOG_FILES.get(catalog_kind, "products.json")
    path = PROJECT_ROOT / "catalog" / fname
    path.parent.mkdir(parents=True, exist_ok=True)
    payload_bytes = json.dumps(
        {"products": products}, ensure_ascii=False, indent=2
    ).encode("utf-8")
    path.write_bytes(payload_bytes)
    _load_products_for.clear()
    _load_products.clear()
    # GitHub 자동 commit (secrets 에 토큰 있으면)
    ok, msg = _push_to_github(
        f"catalog/{fname}", payload_bytes,
        f"품목 관리({catalog_kind}): {len(products)}개 저장"
    )
    st.session_state[f"_github_sync_{catalog_kind}"] = (ok, msg)


def _save_products(products: list[dict]) -> None:
    """기존 호환 — 일반 QR 카탈로그 저장."""
    _save_products_for("qr", products)


@st.cache_data
def _list_brands() -> list[str]:
    brands_dir = PROJECT_ROOT / "brands"
    if not brands_dir.exists():
        return ["softment"]
    return sorted([d.name for d in brands_dir.iterdir()
                   if d.is_dir() and (d / "brand.json").exists()])


def _save_brand(brand_id: str, brand_dict: dict) -> None:
    path = PROJECT_ROOT / "brands" / brand_id / "brand.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(brand_dict, ensure_ascii=False, indent=2),
                    encoding="utf-8")


def _save_labels(labels_dict: dict) -> None:
    path = PROJECT_ROOT / "config" / "labels.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(labels_dict, ensure_ascii=False, indent=2),
                    encoding="utf-8")


# ═════════════════════════════════════════════════════════════
# 페이지 1: 견적서 작성
# ═════════════════════════════════════════════════════════════

# 편집 가능 컬럼 (공급가는 계산 결과라 별도)
ITEM_COLUMNS = ["분류", "항목", "설명", "단가", "기간(횟수)", "수량",
                "할인율(%)", "할인금액", "비고"]
# 화면 표시 순서 (공급가 포함)
DISPLAY_COLUMNS = ["분류", "항목", "설명", "단가", "기간(횟수)", "수량",
                   "할인율(%)", "할인금액", "공급가", "비고"]

# 분류 선택지
# 금액을 어떻게 산정하는 행인지 — 세 값이 같은 층위가 되도록 정액/정률/할인으로 구성
ITEM_KIND_NORMAL = "📋 정액(원)"
ITEM_KIND_DISCOUNT = "💰 할인(차감)"
ITEM_KIND_DEFERRED = "💸 정률(%)"
ITEM_KINDS = [ITEM_KIND_NORMAL, ITEM_KIND_DISCOUNT, ITEM_KIND_DEFERRED]
# 구버전 라벨 → 현재 라벨 (기존 세션·저장본 호환)
ITEM_KIND_ALIASES = {
    "📋 품목": ITEM_KIND_NORMAL,
    "💰 할인행": ITEM_KIND_DISCOUNT,
    "💸 후불(QR결제%)": ITEM_KIND_DEFERRED,
    "💸 정률(결제액 %)": ITEM_KIND_DEFERRED,
}


def _normalize_kind_column(series):
    """'분류' 컬럼의 구버전 라벨을 현재 라벨로 맞춘다."""
    return (series.fillna(ITEM_KIND_NORMAL)
                  .replace("", ITEM_KIND_NORMAL)
                  .replace(ITEM_KIND_ALIASES))
# 품목 관리 — 청구 방식 라벨 (구버전 '후불(QR결제%)' 도 계속 인식)
BILLING_LABEL_FIXED = "일시납"
BILLING_LABEL_RATE = "정률(결제액 %)"
BILLING_LABEL_RATE_LEGACY = "후불(QR결제%)"


def _is_rate_billing(label: str) -> bool:
    """'정률(결제액 %)' 또는 구버전 '후불(QR결제%)' 라벨이면 True."""
    return str(label).strip() in (BILLING_LABEL_RATE, BILLING_LABEL_RATE_LEGACY)


def _empty_items_df() -> pd.DataFrame:
    return pd.DataFrame(columns=ITEM_COLUMNS).astype({
        "분류": "string",
        "항목": "string",
        "설명": "string",
        "단가": "Int64",
        "기간(횟수)": "Int64",
        "수량": "Int64",
        "할인율(%)": "Int64",
        "할인금액": "Int64",
        "비고": "string",
    })


def _ensure_items_state():
    if "items_df" not in st.session_state:
        st.session_state.items_df = _empty_items_df()


def _add_catalog_row(product: dict):
    # billing_type 메타로 분류 자동 셋팅 (deferred_percent → 정률)
    btype = (product.get("billing_type") or "").lower()
    if btype == "deferred_percent":
        kind = ITEM_KIND_DEFERRED
        period_v = None
        qty_v = None
        notes_v = ""
    else:
        kind = ITEM_KIND_NORMAL
        period_v = 1
        qty_v = 1
        notes_v = ""
    new_row = {
        "분류": kind,
        "항목": product["name"],
        "설명": product.get("description", ""),
        "단가": int(product.get("unit_price", 0)),
        "기간(횟수)": period_v,
        "수량": qty_v,
        "할인율(%)": None,
        "할인금액": None,
        "비고": notes_v,
    }
    df = st.session_state.items_df
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True
    )


def _add_blank_row():
    df = st.session_state.items_df
    blank = {c: None for c in ITEM_COLUMNS}
    blank["분류"] = ITEM_KIND_NORMAL
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([blank])],
        ignore_index=True,
    )


def _add_discount_row():
    """할인 행 추가 — 분류 '💰 할인'. 단가(고정금액) 또는 할인율(%)(일괄) 둘 다 사용 가능."""
    df = st.session_state.items_df
    new_row = {
        "분류": ITEM_KIND_DISCOUNT,
        "항목": "할인",
        "설명": "단가(고정금액) 또는 할인율(%)(일괄) 입력",
        "단가": None,
        "기간(횟수)": None,
        "수량": None,
        "할인율(%)": None,
        "할인금액": None,
        "비고": "협상가",
    }
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True,
    )


def _reset_items():
    st.session_state.items_df = _empty_items_df()
    for k in QR_FORM_KEYS:
        if k in st.session_state:
            del st.session_state[k]
    _qr_autosave_clear()


def _row_amount_normal(row) -> int | None:
    """일반 품목 행 한 줄의 공급가 (항목별 할인 반영). 단가 없으면 None."""
    qty = row.get("수량")
    period = row.get("기간(횟수)")
    price = row.get("단가")
    if not (pd.notna(price) and price):
        return None
    q = qty if pd.notna(qty) and qty else 1
    p = period if pd.notna(period) and period else 1
    try:
        gross = int(q) * int(p) * int(price)
    except (TypeError, ValueError):
        return None
    disc_amt = row.get("할인금액")
    disc_rate = row.get("할인율(%)")
    discount = 0
    if pd.notna(disc_amt) and disc_amt:
        try:
            discount = int(disc_amt)
        except (TypeError, ValueError):
            discount = 0
    elif pd.notna(disc_rate) and disc_rate:
        try:
            discount = int(round(gross * float(disc_rate) / 100))
        except (TypeError, ValueError):
            discount = 0
    return gross - discount


def _normal_items_sum(df) -> int:
    """일반(품목) 분류 행들의 amount 합 — 할인 행의 일괄 할인 기준값."""
    if df is None:
        return 0
    total = 0
    for _, row in df.iterrows():
        if row.get("분류") == ITEM_KIND_DISCOUNT:
            continue
        a = _row_amount_normal(row)
        if a is not None:
            total += a
    return total


def _row_amount(row, df=None):
    """행의 공급가 계산.

    - 분류='📋 품목': (수량×기간×단가) − 항목별 할인
    - 분류='💰 할인행': 다른 품목 합 또는 고정 금액 차감
    - 분류='💸 정률(결제액 %)': 합계에서 제외 — 별도 정률 항목 안내로 표시
    """
    kind = row.get("분류")
    if kind == ITEM_KIND_DEFERRED:
        return None    # 정률 행은 즉시 합계에서 제외
    if kind != ITEM_KIND_DISCOUNT:
        return _row_amount_normal(row)

    deduct = 0
    # 1) 단가 × 수량 × 기간 (고정 금액 차감)
    qty = row.get("수량")
    period = row.get("기간(횟수)")
    price = row.get("단가")
    if pd.notna(price) and price:
        q = qty if pd.notna(qty) and qty else 1
        p = period if pd.notna(period) and period else 1
        try:
            deduct += int(q) * int(p) * abs(int(price))
        except (TypeError, ValueError):
            pass
    # 2) 할인금액 (별도 컬럼)
    disc_amt = row.get("할인금액")
    if pd.notna(disc_amt) and disc_amt:
        try:
            deduct += abs(int(disc_amt))
        except (TypeError, ValueError):
            pass
    # 3) 할인율(%) → 다른 품목 합의 N% 만큼 일괄 차감
    disc_rate = row.get("할인율(%)")
    if pd.notna(disc_rate) and disc_rate:
        try:
            base = _normal_items_sum(df) if df is not None else 0
            deduct += int(round(base * float(disc_rate) / 100))
        except (TypeError, ValueError):
            pass
    if deduct == 0:
        return None
    return -deduct


CATALOG_LABELS = {
    "qr": "QR오더 견적기",
    "outdoor": "[야외형] QR오더 견적기",
}


def render_quote_page(catalog_kind: str = "qr"):
    # 메뉴 전환 시 items_df 를 catalog_kind 별로 분리 보관/복원
    active_kind = st.session_state.get("_active_catalog_kind", "qr")
    if active_kind != catalog_kind:
        # 현재 화면의 items_df 를 이전 종류 슬롯에 보관
        if "items_df" in st.session_state:
            st.session_state[f"_items_df_slot_{active_kind}"] = st.session_state["items_df"]
        # 새 종류 슬롯 불러오기 (없으면 빈 표)
        st.session_state["items_df"] = st.session_state.get(
            f"_items_df_slot_{catalog_kind}", _empty_items_df()
        )
        st.session_state["_active_catalog_kind"] = catalog_kind

    _qr_autosave_load_once()
    _ensure_items_state()
    labels = load_labels(PROJECT_ROOT)
    products = _load_products_for(catalog_kind)
    brand_ids = _list_brands()
    soffice_available = find_soffice() is not None

    # ─── 사이드바: 발행 정보 + 카탈로그 ───
    with st.sidebar:
        st.divider()
        st.header("⚙ 발행 정보")
        brand_id = st.selectbox("브랜드", brand_ids,
                                index=0 if brand_ids else None,
                                disabled=len(brand_ids) <= 1)
        issued_date = st.date_input("발행일", value=date.today())
        valid_days = st.number_input(
            "유효기간 (일)",
            value=labels.quote.default_validity_days,
            min_value=1, max_value=365,
        )
        valid_until = issued_date + timedelta(days=int(valid_days))
        st.caption(f"📅 유효기간: **{valid_until.isoformat()}** 까지")

        if not soffice_available:
            st.divider()
            st.warning("⚠ LibreOffice 미감지 — PDF 생성이 불가합니다.")

    # ─── 메인 ───
    kind_label = CATALOG_LABELS.get(catalog_kind, "QR오더 견적기")
    st.title(f"📋 {kind_label}")

    # ─── 최근 다운로드 견적서 (자동 기록) ───
    _render_qr_recent_panel()
    # ─── 견적서 표본 (수기 저장된 정석 케이스) ───
    _render_qr_template_panel()

    st.subheader("0. 발행 담당자 정보")
    # Tab 키가 좌→우→다음 줄 순으로 이동하도록 행 단위로 columns 를 생성
    ic_r1_l, ic_r1_r = st.columns(2)
    with ic_r1_l:
        issuer_name = st.text_input(
            "담당자명", key="issuer_name",
            placeholder="예: 박지은",
        )
    with ic_r1_r:
        issuer_title = st.text_input(
            "직책", key="issuer_title",
            placeholder="예: QR사업부 매니저",
        )
    ic_r2_l, ic_r2_r = st.columns(2)
    with ic_r2_l:
        issuer_phone = st.text_input(
            "연락처", key="issuer_phone",
            placeholder="예: 010-0000-0000",
        )
    with ic_r2_r:
        issuer_email = st.text_input(
            "이메일", key="issuer_email",
            placeholder="예: name@softment.co.kr",
        )

    st.subheader("1. 수신처 정보")
    # Tab 이 좌→우→다음 줄 순으로 이동하도록 행 단위 columns
    cp_r1_l, cp_r1_r = st.columns(2)
    with cp_r1_l:
        cp_name = st.text_input("회사명 *", key="cp_name",
                                placeholder="예: 주식회사 ○○")
    with cp_r1_r:
        cp_contact_name = st.text_input("담당자", key="cp_contact_name",
                                        placeholder="예: 김담당")
    cp_r2_l, cp_r2_r = st.columns(2)
    with cp_r2_l:
        cp_reg = st.text_input("사업자등록번호", key="cp_reg",
                               placeholder="000-00-00000")
    with cp_r2_r:
        cp_contact_title = st.text_input("직책", key="cp_contact_title",
                                         placeholder="예: 구매팀장")
    cp_r3_l, cp_r3_r = st.columns(2)
    with cp_r3_l:
        cp_address = st.text_input("주소", key="cp_address",
                                   placeholder="시/도 ○○구 ○○로 ...")
    with cp_r3_r:
        cp_email = st.text_input("Email", key="cp_email",
                                 placeholder="buyer@example.com")

    st.subheader("2. 건명")
    subject = st.text_input("건명 *", key="subject",
                            placeholder="예: 주문 접수 / QR오더 솔루션 도입 견적",
                            label_visibility="collapsed")

    st.subheader("3. 품목 내역")

    # 품목 selectbox on_change 콜백 — 선택만 해도 자동 행 추가 후 selectbox 리셋
    def _on_catalog_pick():
        idx = st.session_state.get("catalog_pick")
        if idx is None:
            return
        try:
            i = int(idx)
            if 0 <= i < len(products):
                _add_catalog_row(products[i])
        except (TypeError, ValueError):
            pass
        # 다음 선택을 위해 셀렉트박스 비움
        st.session_state["catalog_pick"] = None

    pick_col, blank_col, disc_col, del_col, reset_col = st.columns(
        [5.2, 1.2, 1.3, 1.3, 1.2]
    )
    with pick_col:
        if products:
            options = list(range(len(products)))
            st.selectbox(
                "품목 추가",
                options=options,
                index=None,
                format_func=lambda i: (
                    f"{products[i]['name']} · ₩{int(products[i].get('unit_price', 0)):,}"
                ),
                placeholder="품목 선택 → 자동으로 표에 추가됩니다 (검색 가능)...",
                key="catalog_pick",
                on_change=_on_catalog_pick,
                label_visibility="collapsed",
            )
        else:
            st.info("품목이 비어있습니다. '품목 관리' 페이지에서 상품을 추가하세요.")
    with blank_col:
        if st.button("+ 빈 행", use_container_width=True):
            _add_blank_row()
            st.rerun()
    with disc_col:
        if st.button("💰 + 할인 행", use_container_width=True,
                     help="음수 단가 행이 자동 추가됩니다. 협상 금액으로 단가 조정 가능."):
            _add_discount_row()
            st.rerun()
    with del_col:
        del_clicked = st.button(
            "🗑 선택 삭제", use_container_width=True,
            help="첫 컬럼 '🗑' 체크박스로 선택한 행만 삭제합니다.",
            key="qr_del_selected",
        )
    with reset_col:
        if st.button("전체 초기화", use_container_width=True,
                     key="qr_reset_request"):
            st.session_state["_qr_reset_confirm"] = True

    # ── 전체 초기화 확인 다이얼로그 ──
    if st.session_state.get("_qr_reset_confirm"):
        with st.container(border=True):
            st.markdown(
                "<div style='background:#FDECEA;border-left:4px solid #C0392B;"
                "padding:10px 14px;border-radius:6px'>"
                "<strong style='color:#C0392B'>⚠ 전체 초기화 확인</strong><br>"
                "<span style='font-size:0.9rem'>"
                "현재 작성 중인 모든 입력(품목·수신처·담당자·기타 안내)이 사라집니다. "
                "정말 초기화할까요?</span></div>",
                unsafe_allow_html=True,
            )
            cc1, cc2, _ = st.columns([1.2, 1.2, 4])
            with cc1:
                if st.button("초기화 진행", type="primary",
                             use_container_width=True, key="qr_reset_yes"):
                    _reset_items()
                    st.session_state["_qr_reset_confirm"] = False
                    st.rerun()
            with cc2:
                if st.button("취소", use_container_width=True,
                             key="qr_reset_no"):
                    st.session_state["_qr_reset_confirm"] = False
                    st.rerun()

    if st.session_state.items_df.empty:
        st.info("아직 품목이 없습니다. 위 드롭다운에서 품목을 추가하거나 '+ 빈 행' 을 누르세요.")
        edited_df = st.session_state.items_df
    else:
        # 행 순서 변경 — 카드 + ☰ 아이콘 + ⬆⬇ 버튼 (한 클릭에 행 이동)
        df_for_order = st.session_state.items_df.reset_index(drop=True)
        if len(df_for_order) > 1:
            with st.expander(f"🔃 행 순서 변경 ({len(df_for_order)}건)", expanded=False):
                move_action: tuple[str, int] | None = None
                for i in range(len(df_for_order)):
                    name = df_for_order.at[i, "항목"] or "(이름 없음)"
                    price = df_for_order.at[i, "단가"]
                    label = f"**[{i + 1}]** {name}"
                    if pd.notna(price) and price:
                        try:
                            label += f"  ·  ₩{int(price):,}"
                        except (TypeError, ValueError):
                            pass
                    with st.container(border=True):
                        rcols = st.columns([0.5, 8, 1, 1])
                        with rcols[0]:
                            st.markdown(
                                "<div style='font-size:1.3rem;color:#888;"
                                "text-align:center;padding-top:2px;'>☰</div>",
                                unsafe_allow_html=True,
                            )
                        with rcols[1]:
                            st.write(label)
                        with rcols[2]:
                            if st.button("⬆", key=f"row_up_{i}",
                                         disabled=(i == 0),
                                         use_container_width=True):
                                move_action = ("up", i)
                        with rcols[3]:
                            if st.button("⬇", key=f"row_dn_{i}",
                                         disabled=(i == len(df_for_order) - 1),
                                         use_container_width=True):
                                move_action = ("dn", i)
                if move_action is not None:
                    direction, idx = move_action
                    j = idx - 1 if direction == "up" else idx + 1
                    df = st.session_state.items_df.reset_index(drop=True)
                    df.iloc[[idx, j]] = df.iloc[[j, idx]].values
                    st.session_state.items_df = df
                    st.rerun()

        # 공급가 컬럼을 계산해서 디스플레이용 DataFrame 생성
        display_df = st.session_state.items_df.copy()
        # 분류 값이 비었으면 기본 '품목'으로 채움
        display_df["분류"] = _normalize_kind_column(display_df["분류"])
        display_df["공급가"] = display_df.apply(
            lambda r: _row_amount(r, df=display_df), axis=1
        ).astype("Int64")
        # 후불(QR결제%) 행: 단가가 0/빈 값이면 화면에서도 빈 칸 (₩0 표기 방지)
        deferred_mask = display_df["분류"].astype(str) == ITEM_KIND_DEFERRED
        if deferred_mask.any():
            zero_price = display_df["단가"].fillna(0) == 0
            display_df.loc[deferred_mask & zero_price, "단가"] = pd.NA
        display_df = display_df[DISPLAY_COLUMNS]
        # 첫 컬럼에 선택용 체크박스 — '🗑 선택 삭제' 버튼이 활용
        display_df.insert(0, "🗑", False)

        edited_df = st.data_editor(
            display_df,
            column_config={
                "🗑": st.column_config.CheckboxColumn(
                    "🗑", width="small", default=False,
                    help="삭제할 행을 체크 후 우측 상단 '🗑 선택 삭제' 클릭",
                ),
                "분류": st.column_config.SelectboxColumn(
                    "금액 방식", options=ITEM_KINDS, required=True,
                    width="small",
                    help=("이 행의 금액을 어떻게 정할지 고릅니다. "
                          "'📋 정액(원)': 단가 × 수량 × 기간. "
                          "'💸 정률(%)': 결제액의 N% — 합계에서 빠지고 "
                          "견적서에는 단가·공급가가 '-' 로 표시됩니다. "
                          "'💰 할인(차감)': 입력한 만큼 합계에서 차감."),
                ),
                "항목": st.column_config.TextColumn("항목", required=True, width="medium"),
                "설명": st.column_config.TextColumn("설명", width="large"),
                "단가": st.column_config.NumberColumn(
                    "단가", step=1000, format="₩%,d", width="small",
                    help=("일반 품목: 단가(원). '💰 할인행': 자동 차감. "
                          "'💸 정률(결제액 %)': 단가 칸 값이 % 로 해석됩니다 "
                          "(예: 5 → QR결제액의 5%). 2.5% 처럼 소수 요율은 "
                          "단가를 비우고 설명 칸에 'QR오더 결제액의 2.5%' 로 "
                          "적으면 견적서에 그대로 출력됩니다."),
                ),
                "기간(횟수)": st.column_config.NumberColumn(
                    "기간(횟수)", min_value=0, step=1, width="small",
                ),
                "수량": st.column_config.NumberColumn(
                    "수량", min_value=0, step=1, width="small",
                ),
                "할인율(%)": st.column_config.NumberColumn(
                    "할인율(%)", min_value=0, max_value=100, step=1, format="%d%%",
                    width="small",
                    help=("일반 품목 행: 그 행에만 적용되는 할인율 (할인금액이 있으면 그쪽 우선). "
                          "💰 할인 행: 일반 품목 합계의 N% 만큼 일괄 차감."),
                ),
                "할인금액": st.column_config.NumberColumn(
                    "할인금액", min_value=0, step=1000, format="₩%,d", width="small",
                    help="항목별 할인 금액 (양수).",
                ),
                "공급가": st.column_config.NumberColumn(
                    "공급가", disabled=True, format="₩%,d", width="small",
                    help="(수량 × 기간 × 단가) − 항목별 할인",
                ),
                "비고": st.column_config.TextColumn("비고", width="medium"),
            },
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            key="items_editor",
        )
        # ── '🗑 선택 삭제' 처리: 체크된 행만 items_df 에서 제거 ──
        if del_clicked:
            try:
                mask_del = edited_df["🗑"].fillna(False).astype(bool).tolist()
                if any(mask_del):
                    new_df = edited_df.loc[[not m for m in mask_del]][ITEM_COLUMNS].reset_index(drop=True)
                    st.session_state.items_df = new_df
                    st.toast(f"🗑 {sum(mask_del)}개 행 삭제", icon="🗑")
                    st.rerun()
                else:
                    st.toast("삭제할 행을 첫 컬럼 '🗑' 체크박스로 선택하세요",
                             icon="ℹ️")
            except KeyError:
                pass
        # 편집 가능 컬럼만 세션에 저장 ('🗑' 선택 컬럼은 제외)
        edited_core = edited_df[ITEM_COLUMNS]
        # 공급가가 입력 변경에 따라 즉시 갱신되도록 — 변경 감지 시 rerun
        old = st.session_state.items_df.reset_index(drop=True)
        new = edited_core.reset_index(drop=True)
        changed = (
            len(old) != len(new) or
            not all((old[c].astype(object).fillna("__").tolist()
                     == new[c].astype(object).fillna("__").tolist())
                    for c in ITEM_COLUMNS)
        )
        st.session_state.items_df = edited_core
        if changed:
            st.rerun()

    if not edited_df.empty:
        amounts = edited_df.apply(
            lambda r: _row_amount(r, df=edited_df), axis=1
        )
        subtotal = int(pd.to_numeric(amounts, errors="coerce").fillna(0).sum())

        st.divider()

        vat_rate = labels.quote.vat_rate
        vat = int(round(subtotal * vat_rate))
        total = subtotal + vat

        m1, m2, m3 = st.columns(3)
        m1.metric("공급가액", f"₩{subtotal:,}")
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", f"₩{vat:,}")
        m3.metric("합계 금액", f"₩{total:,}")

        # ── 후청구(후불) 안내 박스 — 분류='💸 후불(QR결제%)' 행 모음 ──
        deferred_mask = edited_df["분류"].astype(str) == ITEM_KIND_DEFERRED
        deferred_rows = edited_df[deferred_mask]
        if not deferred_rows.empty:
            lines = []
            for _, r in deferred_rows.iterrows():
                name = (r.get("항목") or "솔루션 사용료") or ""
                rate = r.get("단가")
                # 2.5 같은 소수 요율도 그대로 표시 (2.5% → '2.5%', 3.0 → '3%')
                rate_txt = (f"{float(rate):g}%"
                            if pd.notna(rate) and rate else None)
                desc = (r.get("설명") or "").strip()
                if rate_txt:
                    body = (f"결제액의 <strong>{rate_txt}</strong>"
                            + (f" — {desc}" if desc else ""))
                else:
                    # 단가 칸을 비우고 설명에 요율 문구를 쓴 경우
                    body = desc or "설명 칸에 요율 문구를 입력할 수 있어요"
                lines.append(f"<li><strong>{name}</strong> · {body}</li>")
            st.markdown(
                f"""
<div style="background:#FFF7E6; border-left:4px solid #D97706;
            border-radius:6px; padding:10px 14px; margin:8px 0 4px;">
  <div style="color:#92400E; font-weight:700; font-size:0.95rem;">
    💸 정률 항목 · {len(deferred_rows)}건 (위 합계와 별도로 결제액 기준 정산)
  </div>
  <ul style="margin:6px 0 0 18px; padding:0; color:#7C2D12; font-size:0.88rem;">
    {''.join(lines)}
  </ul>
</div>
                """,
                unsafe_allow_html=True,
            )

    st.subheader("4. 기타 안내")
    # 자동 안내 문구를 textarea 기본값으로 미리 채움 — 사용자가 직접 수정 가능
    # (notes 가 비어있을 때만 자동 채움. 표본/히스토리 불러오면 그 값이 들어가 자동 채움 X)
    if not st.session_state.get("notes"):
        auto_lines = _compute_auto_notice_lines(brand_id, valid_until, labels)
        if auto_lines:
            st.session_state["notes"] = "\n".join(auto_lines)
    notes = st.text_area(
        "기타 안내", key="notes",
        placeholder=(
            "유효기간·입금 계좌·결제 조건 등 안내문구를 줄별로 입력하세요. "
            "(처음에는 자동으로 채워지며, 자유롭게 수정 가능합니다.)"
        ),
        label_visibility="collapsed", height=120,
        help="여기에 입력된 내용이 견적서 PDF 의 '기타 안내' 영역에 그대로 표시됩니다.",
    )

    st.divider()
    common_args = dict(
        brand_id=brand_id, issued_date=issued_date, valid_until=valid_until,
        counterparty_data=dict(
            name=cp_name.strip(),
            registration_number=cp_reg.strip() or None,
            address=cp_address.strip() or None,
            contact_name=cp_contact_name.strip() or None,
            contact_title=cp_contact_title.strip() or None,
            email=cp_email.strip() or None,
        ),
        issuer_contact_data=dict(
            name=issuer_name.strip(),
            title=issuer_title.strip() or None,
            phone=issuer_phone.strip() or None,
            email=issuer_email.strip() or None,
        ),
        subject=subject.strip(),
        items_df=edited_df,
        notes=notes.strip() or None,
        soffice_available=soffice_available,
    )

    btn_prev, btn_gen = st.columns([1, 1])
    with btn_prev:
        preview_clicked = st.button(
            "👁 미리보기", use_container_width=True,
            disabled=not soffice_available,
            help=("작성한 내용을 PDF로 렌더링해서 이 페이지 안에서 바로 확인합니다. "
                  "다운로드는 아래 '견적서 생성' 버튼을 누르세요.")
            if soffice_available else "PDF 변환기(LibreOffice) 미감지 — 미리보기 불가",
        )
    with btn_gen:
        generate_clicked = st.button(
            "📝 견적서 생성 (다운로드)", type="primary", use_container_width=True,
        )

    if preview_clicked:
        _preview_quote(**common_args)
    if generate_clicked:
        _generate_quote(**common_args)

    # 작성 내용 자동저장 + 페이지 이탈 시 브라우저 경고
    _qr_autosave_write()
    df = st.session_state.get("items_df")
    has_data = (df is not None and not df.empty) or any(
        st.session_state.get(k) for k in QR_FORM_KEYS
    )
    _inject_beforeunload(has_data)


def _compute_auto_notice_lines(brand_id: str, valid_until: date,
                               labels: DocumentLabels) -> list[str]:
    """'기타 안내' 영역에 자동으로 들어가는 고정/계산 문구 목록을 반환."""
    ql = labels.quote
    lines: list[str] = []
    if valid_until and ql.auto_notices.validity_template:
        valid_str = valid_until.strftime("%Y년 %m월 %d일")
        lines.append(ql.auto_notices.validity_template.format(valid_until=valid_str))
    try:
        brand = load_brand(PROJECT_ROOT, brand_id)
    except Exception:
        brand = None
    if brand and brand.bank_account and ql.auto_notices.bank_account_template:
        ba = brand.bank_account
        lines.append(ql.auto_notices.bank_account_template.format(
            bank=ba.bank,
            account_number=ba.account_number,
            account_holder=ba.account_holder or "",
        ))
    for static in ql.static_notices:
        if static:
            lines.append(static)
    return lines


def _build_quote_artifacts(*, brand_id, issued_date, valid_until,
                           counterparty_data, issuer_contact_data,
                           subject, items_df, notes,
                           soffice_available,
                           status_label="문서 생성 중..."):
    """입력값을 검증·렌더링하여 (document_id, docx_bytes, pdf_bytes) 를 반환.
    실패 시 None 반환 (사용자에게 에러는 이미 표시됨)."""
    if not counterparty_data["name"]:
        st.error("❌ 회사명은 필수입니다.")
        return None
    if not subject:
        st.error("❌ 건명은 필수입니다.")
        return None

    # 할인 행의 일괄 할인율(%) 계산 기준값 — 일반 품목 행들의 amount 합
    normal_sum = _normal_items_sum(items_df)

    items: list[LineItem] = []
    for _, row in items_df.iterrows():
        name = row.get("항목")
        name = name.strip() if isinstance(name, str) else None
        if not name:
            continue
        qty = row.get("수량")
        period = row.get("기간(횟수)")
        unit_price_raw = row.get("단가")
        kind = row.get("분류")
        is_discount_row = kind == ITEM_KIND_DISCOUNT
        is_deferred_row = kind == ITEM_KIND_DEFERRED
        desc_val = row.get("설명")
        notes_val = row.get("비고")
        disc_rate = row.get("할인율(%)")
        disc_amt = row.get("할인금액")

        if is_deferred_row:
            # 후불(QR결제 %): 합계 미포함, PDF 에서 단가/공급가 '-' 로 표시
            # 자동 prefix 추가하지 않음 — 사용자가 설명/비고에 직접 작성
            items.append(LineItem(
                name=name,
                description=desc_val if isinstance(desc_val, str) and desc_val else None,
                qty=None, period=None,
                unit_price=0,
                billing_type="deferred_percent",
                notes=notes_val if isinstance(notes_val, str) and notes_val else None,
            ))
            continue

        if is_discount_row:
            # 분류=할인: (단가×수량×기간 + 할인금액 + 일반품목합×할인율%) 만큼 음수 단가로 환산
            deduct = 0.0
            if pd.notna(unit_price_raw) and unit_price_raw:
                q_v = float(qty) if pd.notna(qty) and qty else 1.0
                p_v = float(period) if pd.notna(period) and period else 1.0
                deduct += q_v * p_v * abs(float(unit_price_raw))
            if pd.notna(disc_amt) and disc_amt:
                deduct += abs(float(disc_amt))
            if pd.notna(disc_rate) and disc_rate:
                deduct += normal_sum * float(disc_rate) / 100.0
            if deduct <= 0:
                continue  # 차감 금액이 0이면 의미 없는 행 → skip
            items.append(LineItem(
                name=name,
                description=desc_val if isinstance(desc_val, str) and desc_val else None,
                qty=None, period=None,
                unit_price=-deduct,
                notes=notes_val if isinstance(notes_val, str) and notes_val else None,
            ))
        else:
            # 일반 품목 행
            unit_price = (float(unit_price_raw)
                          if pd.notna(unit_price_raw) and unit_price_raw else 0.0)
            items.append(LineItem(
                name=name,
                description=desc_val if isinstance(desc_val, str) and desc_val else None,
                qty=float(qty) if pd.notna(qty) and qty else None,
                period=float(period) if pd.notna(period) and period else None,
                unit_price=unit_price,
                discount_rate=(float(disc_rate) / 100
                               if pd.notna(disc_rate) and disc_rate else None),
                discount_amount=(float(disc_amt)
                                 if pd.notna(disc_amt) and disc_amt else None),
                notes=notes_val if isinstance(notes_val, str) and notes_val else None,
            ))

    if not items:
        st.error("❌ 최소 1개 이상의 품목을 입력해주세요.")
        return None

    cp_name_short = "".join(c for c in counterparty_data["name"] if c.isalnum())[:8]
    document_id = f"Q-{issued_date.strftime('%Y%m%d')}-{cp_name_short or '고객'}"

    document = QuoteDocument(
        document_id=document_id, document_type="quote",
        brand_id=brand_id,
        issued_date=issued_date, valid_until=valid_until,
        counterparty=Counterparty(**counterparty_data),
        subject=subject, line_items=items,
        totals=Totals(), clauses=[], notes=notes,
    )

    try:
        brand = load_brand(PROJECT_ROOT, brand_id)
    except FileNotFoundError as e:
        st.error(f"❌ 브랜드 로드 실패: {e}")
        return None

    # 발행 담당자 정보 — 폼에서 입력한 값으로 일회용 오버라이드
    issuer_name = (issuer_contact_data.get("name") or "").strip()
    if issuer_name:
        brand = brand.model_copy(update={
            "contact_person": ContactPerson(
                name=issuer_name,
                title=issuer_contact_data.get("title"),
                phone=issuer_contact_data.get("phone"),
                email=issuer_contact_data.get("email"),
            )
        })
    else:
        # 담당자명 비워두면 견적서에서 담당자/연락처/이메일 줄 자체를 숨김
        brand = brand.model_copy(update={"contact_person": None})

    output_dir = PROJECT_ROOT / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    docx_path = output_dir / f"{document_id}.docx"

    with st.status(status_label, expanded=True) as status:
        st.write("📝 DOCX 렌더링 중...")
        try:
            render_docx(brand, document, PROJECT_ROOT, docx_path)
        except Exception as e:
            status.update(label="❌ DOCX 생성 실패", state="error")
            st.exception(e)
            return None
        st.write(f"   ✓ {docx_path.name}")

        pdf_bytes = None
        if soffice_available:
            st.write("📑 PDF 변환 중...")
            try:
                pdf_path = convert_docx_to_pdf(docx_path, output_dir)
                pdf_bytes = pdf_path.read_bytes()
                st.write(f"   ✓ {pdf_path.name}")
            except Exception as e:
                st.warning(f"PDF 변환 실패 (DOCX만 사용 가능): {e}")

        status.update(label="✅ 완료", state="complete")

    docx_bytes = docx_path.read_bytes()
    return document_id, docx_bytes, pdf_bytes


def _generate_quote(**kwargs):
    result = _build_quote_artifacts(status_label="문서 생성 중...", **kwargs)
    if result is None:
        return
    document_id, docx_bytes, pdf_bytes = result

    # 최근 다운로드 견적서 히스토리에 저장 (같은 document_id 면 1개로 유지)
    _qr_save_history(_qr_snapshot_payload(), document_id)

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📝 DOCX 다운로드", data=docx_bytes,
            file_name=f"{document_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl2:
        if pdf_bytes:
            st.download_button(
                "📑 PDF 다운로드", data=pdf_bytes,
                file_name=f"{document_id}.pdf",
                mime="application/pdf", use_container_width=True,
            )
        else:
            st.button("📑 PDF 사용 불가", disabled=True, use_container_width=True)


def _preview_quote(**kwargs):
    result = _build_quote_artifacts(status_label="미리보기 생성 중...", **kwargs)
    if result is None:
        return
    document_id, docx_bytes, pdf_bytes = result

    if not pdf_bytes:
        st.error("❌ PDF 변환기(LibreOffice)를 사용할 수 없어 미리보기를 표시할 수 없습니다.")
        return

    # 미리보기에서 다운로드한 경우도 히스토리에 남도록 저장
    # (같은 document_id 면 중복 없이 1개로 갱신)
    _qr_save_history(_qr_snapshot_payload(), document_id)

    st.success(f"미리보기: **{document_id}.pdf**")

    # PDF를 페이지별 이미지로 렌더링 (브라우저 PDF 차단 회피)
    rendered = False
    try:
        import pypdfium2 as pdfium  # noqa: WPS433  (지연 임포트로 실패 시 graceful fallback)
        pdf = pdfium.PdfDocument(pdf_bytes)
        try:
            n_pages = len(pdf)
            for i in range(n_pages):
                page = pdf[i]
                bitmap = page.render(scale=2)  # 2배 해상도 (선명)
                img = bitmap.to_pil()
                st.image(img, caption=f"페이지 {i + 1} / {n_pages}",
                         use_container_width=True)
            rendered = True
        finally:
            pdf.close()
    except Exception as e:
        st.warning(f"이미지 미리보기를 사용할 수 없습니다 ({e}). 아래 다운로드로 확인하세요.")

    # DOCX + PDF 다운로드 (택1)
    dl1, dl2 = st.columns(2)
    with dl1:
        if docx_bytes:
            st.download_button(
                "📝 DOCX 다운로드", data=docx_bytes,
                file_name=f"{document_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_docx_preview_{document_id}",
            )
        else:
            st.button("📝 DOCX 사용 불가", disabled=True,
                      use_container_width=True,
                      key=f"dl_docx_preview_disabled_{document_id}")
    with dl2:
        st.download_button(
            "📑 PDF 다운로드", data=pdf_bytes,
            file_name=f"{document_id}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"dl_pdf_preview_{document_id}",
        )


# ═════════════════════════════════════════════════════════════
# 페이지 2: 품목 관리
# ═════════════════════════════════════════════════════════════

def render_catalog_page():
    st.title("📦 품목 관리")
    st.caption(
        "견적서 작성 화면의 '품목 빠른 추가' 드롭다운에 사용되는 상품 목록입니다. "
        "탭으로 견적서 종류별로 관리할 수 있어요."
    )

    # ── 영구 저장 상태 확인 ──
    try:
        token_set = bool(st.secrets.get("GITHUB_TOKEN"))
    except Exception:
        token_set = False

    if token_set:
        pass  # 영구 저장 활성화 상태 — 별도 안내 표시 안 함
    else:
        with st.expander(
            "⚠ **영구 저장 설정이 필요해요 (3분 소요 · 한 번만)** — 클릭해서 펼치기",
            expanded=True,
        ):
            st.markdown("""
            지금은 **저장한 품목이 시간이 지나면 사라지는 상태**예요.
            아래 3단계만 따라하면 이후엔 누가 사용해도 영구 보존됩니다.

            ---

            ### ① GitHub 토큰 발급 (2분)

            아래 링크를 **새 탭에서 열어** 진행하세요 👇

            🔗 [**GitHub 토큰 발급 페이지 열기**](https://github.com/settings/tokens/new?scopes=repo&description=QR_quote+%EC%9E%90%EB%8F%99%EC%A0%80%EC%9E%A5)

            1. 페이지가 열리면 → **Expiration** 을 `No expiration` 으로 선택
            2. 페이지 맨 아래 초록 버튼 **Generate token** 클릭
            3. 화면에 `ghp_xxxxxxxxxxxxx...` 같은 긴 문자열이 표시돼요
            4. **복사 버튼(📋)** 을 눌러 토큰을 복사 (한 번만 보여요!)

            ---

            ### ② Streamlit Cloud Secrets 등록 (1분)

            🔗 [**Streamlit Cloud 열기**](https://share.streamlit.io)

            1. 본인의 앱 `qrquote` 클릭
            2. 우측 위 점 3개(`⋮`) → **Settings** 클릭
            3. 왼쪽 메뉴에서 **Secrets** 클릭
            4. 큰 텍스트 박스에 **아래 4줄 그대로 복사·붙여넣기**:
            """)
            st.code("""GITHUB_TOKEN = "여기에_복사한_토큰_붙여넣기"
GITHUB_OWNER = "jieunpark322"
GITHUB_REPO = "QR_quote"
GITHUB_BRANCH = "main\"""", language="toml")
            st.markdown("""
            5. 첫 줄의 `여기에_복사한_토큰_붙여넣기` 자리에 ①에서 복사한 토큰을 따옴표 안에 붙여넣기
            6. 오른쪽 아래 **Save** 클릭 → 앱이 1~2분 안에 다시 시작

            ---

            ### ③ 동작 확인

            앱이 다시 켜지면 이 페이지 위쪽에 🌐 **자동 영구 저장 활성화** 라는 초록 박스가 보여요.
            그러면 끝! 이제 품목을 저장할 때마다 GitHub 에 자동 commit 됩니다.

            ---

            > 💡 **위 절차가 어렵게 느껴지면** — 아래 **📥 다운로드 / 📤 업로드** 로 임시 백업도 가능합니다.
            > 다만 매번 수동이라 자동 설정을 한 번 해두는 게 편해요.
            """)
        st.write("")

        for kind, label, fname in [
            ("qr", "📋 QR오더 품목", "products.json"),
            ("outdoor", "🌳 [야외형] QR오더 품목", "qr_outdoor_products.json"),
            ("membership", "🏢 멤버십 품목", "membership_products.json"),
        ]:
            with st.container(border=True):
                c1, c2, c3 = st.columns([3, 2, 2])
                with c1:
                    st.markdown(f"**{label}**<br>"
                                f"<span style='color:#6B7280;font-size:0.82rem'>"
                                f"catalog/{fname}</span>",
                                unsafe_allow_html=True)
                with c2:
                    # 현재 카탈로그 JSON 다운로드
                    if kind == "membership":
                        cur_products = _load_membership_products()
                    else:
                        cur_products = _load_products_for(kind)
                    json_bytes = json.dumps(
                        {"products": cur_products},
                        ensure_ascii=False, indent=2,
                    ).encode("utf-8")
                    st.download_button(
                        "📥 다운로드",
                        data=json_bytes,
                        file_name=fname,
                        mime="application/json",
                        use_container_width=True,
                        key=f"dl_{kind}",
                    )
                with c3:
                    uploaded = st.file_uploader(
                        "📤 업로드(JSON)", type=["json"],
                        key=f"up_{kind}",
                        label_visibility="collapsed",
                    )
                    if uploaded is not None:
                        try:
                            payload = json.loads(uploaded.read().decode("utf-8"))
                            new_products = payload.get("products", [])
                            if not isinstance(new_products, list):
                                st.error("올바른 형식이 아닙니다.")
                            else:
                                if kind == "membership":
                                    _save_membership_products(new_products)
                                else:
                                    _save_products_for(kind, new_products)
                                st.success(f"✅ {len(new_products)}개 상품 복원 완료")
                                st.session_state.pop(f"up_{kind}", None)
                                st.rerun()
                        except (json.JSONDecodeError, UnicodeDecodeError) as e:
                            st.error(f"파일 읽기 실패: {e}")

    tab_qr, tab_out, tab_mc = st.tabs([
        "📋 QR오더 품목",
        "🌳 [야외형] QR오더 품목",
        "🏢 멤버십 품목",
    ])
    with tab_qr:
        _render_qr_catalog_editor(catalog_kind="qr")
    with tab_out:
        _render_qr_catalog_editor(catalog_kind="outdoor")
    with tab_mc:
        _render_membership_catalog_editor()


def _safe_str(v) -> str:
    """pandas Series 값을 NaN-safe 하게 문자열로 변환 (NaN/None → '')."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return v if isinstance(v, str) else str(v)


def _qr_catalog_signature(products: list[dict]) -> str:
    """카탈로그 상태 비교용 시그니처 (변경 감지)."""
    return json.dumps(products, ensure_ascii=False, sort_keys=True, default=str)


def _render_qr_catalog_editor(catalog_kind: str = "qr"):
    """QR/야외형 견적서용 평면 카탈로그 편집기 — catalog_kind 별로 분기."""
    products = _load_products_for(catalog_kind)
    sig_key = f"_qr_catalog_original_sig_{catalog_kind}"
    # 비교용 원본 시그니처를 첫 진입 시 저장
    if sig_key not in st.session_state:
        st.session_state[sig_key] = _qr_catalog_signature(products)

    if not products:
        df = pd.DataFrame([{
            "code": "",
            "name": "(여기를 클릭해서 수정)",
            "description": "",
            "unit_price": 0,
            "currency": "KRW",
            "billing_type": "일시납",
        }])
    else:
        df = pd.DataFrame(products)
        for col, default in [("description", ""), ("currency", "KRW"),
                              ("code", ""), ("billing_type", "")]:
            if col not in df.columns:
                df[col] = default
        # 내부 enum 값을 사람이 보기 좋은 라벨로
        df["billing_type"] = df["billing_type"].apply(
            lambda v: BILLING_LABEL_RATE if str(v).lower() == "deferred_percent"
            else BILLING_LABEL_FIXED
        )
        # 정률 항목 단가가 0/빈 값이면 화면에서 빈 칸 (₩0 표기 방지)
        if "unit_price" in df.columns:
            deferred = df["billing_type"].apply(_is_rate_billing)
            if deferred.any():
                zero_or_na = df["unit_price"].fillna(0) == 0
                df.loc[deferred & zero_or_na, "unit_price"] = pd.NA

    # ── 상단 액션 바: 캡션(좌) + 저장 버튼(우) ──
    cap_col, save_col = st.columns([5, 1.3])
    with cap_col:
        st.caption(
            "💡 설명/청구기준은 셀 안에서 Alt+Enter로 줄바꿈 입력 가능합니다. "
            "행 앞 **☰** 아이콘 영역을 잡아 위·아래로 드래그하면 순서를 바꿀 수 있어요."
        )
    with save_col:
        save_label = "💾 야외형 품목 저장" if catalog_kind == "outdoor" else "💾 QR 품목 저장"
        save_clicked = st.button(
            save_label, type="primary",
            use_container_width=True, key=f"save_qr_catalog_{catalog_kind}",
        )

    # ── 저장 직후 화면 상단 토스트만 표시 (안내 row 제거) ──
    saved_marker = st.session_state.get(f"_catalog_saved_at_{catalog_kind}")
    if saved_marker:
        ts, count, page_name = saved_marker
        # 화면 중간 상단에 5초간 떠 있는 fixed toast (HTML/CSS 만으로 자동 fade)
        from streamlit.components.v1 import html as _html
        _html(f"""
<style>
@keyframes qr_toast_in {{
  0% {{ opacity: 0; transform: translate(-50%, -20px); }}
  10% {{ opacity: 1; transform: translate(-50%, 0); }}
  85% {{ opacity: 1; transform: translate(-50%, 0); }}
  100% {{ opacity: 0; transform: translate(-50%, -20px); }}
}}
</style>
<script>
(function() {{
  try {{
    var doc = window.parent.document;
    var prev = doc.getElementById('qr-save-toast');
    if (prev) prev.remove();
    var div = doc.createElement('div');
    div.id = 'qr-save-toast';
    div.innerHTML = '✅ <strong>{count}개 상품 저장 완료</strong>';
    div.style.cssText = [
      'position: fixed', 'top: 70px', 'left: 50%',
      'transform: translateX(-50%)',
      'background: #10B981', 'color: white',
      'padding: 14px 28px', 'border-radius: 999px',
      'font-size: 0.95rem', 'font-weight: 600',
      'box-shadow: 0 10px 25px rgba(16,185,129,0.35)',
      'z-index: 9999', 'animation: qr_toast_in 5s ease forwards',
      'pointer-events: none',
    ].join(';');
    doc.body.appendChild(div);
    setTimeout(function() {{ if (div.parentNode) div.remove(); }}, 5200);
  }} catch (e) {{}}
}})();
</script>
""", height=0)
        # 한 번 표시 후 마커 제거 — 페이지 재방문/rerun 시 토스트가 다시 뜨지 않도록
        st.session_state.pop(f"_catalog_saved_at_{catalog_kind}", None)

    # ── 드래그앤드롭 순서 변경 ──
    if len(df) > 1:
        with st.expander("🔃 행 순서 변경 (드래그앤드롭)", expanded=False):
            st.caption("왼쪽 **☰** 아이콘 영역을 잡아 위·아래로 드래그하세요.")
            try:
                from streamlit_sortables import sort_items
                cat_labels = [
                    f"☰  [{i + 1}]  {df.iloc[i]['name'] or '(이름 없음)'}"
                    + (f"  ·  ₩{int(df.iloc[i]['unit_price']):,}"
                       if pd.notna(df.iloc[i].get('unit_price'))
                       and df.iloc[i].get('unit_price') else "")
                    for i in range(len(df))
                ]
                sorted_labels = sort_items(
                    cat_labels, direction="vertical",
                    key=f"qr_cat_sort_dnd_{catalog_kind}",
                )
                if sorted_labels and sorted_labels != cat_labels:
                    try:
                        new_order = [
                            int(s.split("[", 1)[1].split("]", 1)[0]) - 1
                            for s in sorted_labels
                        ]
                        df = df.iloc[new_order].reset_index(drop=True)
                    except (ValueError, IndexError):
                        pass
            except ImportError:
                st.warning("드래그앤드롭 컴포넌트가 설치되지 않았습니다.")

    edited = st.data_editor(
        df[["code", "name", "description", "unit_price",
             "currency", "billing_type"]],
        column_config={
            "code": st.column_config.TextColumn(
                "코드 (선택사항)", help="고유 식별자 (영문/숫자/하이픈). 비워두면 자동 생성.",
                width="small",
            ),
            "name": st.column_config.TextColumn(
                "품목명", help="견적서에 표시되는 이름", required=True,
                width="medium",
            ),
            "description": st.column_config.TextColumn(
                "설명/청구기준",
                help=("Alt+Enter 로 줄바꿈 입력. "
                      "여러 줄 입력하면 표·PDF 모두 줄바꿈 그대로 반영됩니다."),
                width="large",
            ),
            "unit_price": st.column_config.NumberColumn(
                "단가", min_value=0, step=1000, format="%,d",
                width="small",
                help=("일시납은 원 단위 / 정률(결제액 %) 은 % 값 (예: 5 → 5%) 입력."),
            ),
            "currency": st.column_config.SelectboxColumn(
                "통화", options=["KRW", "USD", "EUR", "JPY"],
                width="small",
                help="견적서 PDF 에서 단가 옆에 자동으로 붙는 통화 기호 (₩ / $ / € / ¥).",
            ),
            "billing_type": st.column_config.SelectboxColumn(
                "청구 방식",
                options=[BILLING_LABEL_FIXED, BILLING_LABEL_RATE],
                width="small",
                help=("'일시납' 은 일반 단가. "
                      "'정률(결제액 %)' 은 견적서에서 자동으로 '💸 정률' 분류로 추가되고, "
                      "단가는 % 로 해석되어 합계와 별도로 안내 영역에 표시."),
            ),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=720,
        key=f"catalog_editor_qr_{catalog_kind}",
    )

    # ── 저장 처리 ──
    if save_clicked:
        new_products = []
        for idx, row in edited.iterrows():
            name = _safe_str(row.get("name")).strip()
            if not name:
                continue
            code = _safe_str(row.get("code")).strip()
            if not code:
                base = "".join(c for c in name if c.isalnum() or c in "_-")[:20] or "ITEM"
                code = f"{base}-{idx + 1}"
            up_v = row.get("unit_price")
            btype_label = _safe_str(row.get("billing_type")).strip()
            item: dict = {
                "code": code,
                "name": name,
                "description": _safe_str(row.get("description")).strip(),
                "unit_price": (int(up_v)
                               if pd.notna(up_v) and up_v not in ("", None)
                               else 0),
                "currency": _safe_str(row.get("currency")).strip() or "KRW",
            }
            if _is_rate_billing(btype_label):
                item["billing_type"] = "deferred_percent"
            new_products.append(item)
        with st.spinner("💾 저장 중..."):
            _save_products_for(catalog_kind, new_products)
            st.session_state[sig_key] = _qr_catalog_signature(new_products)
        page_name = CATALOG_LABELS.get(catalog_kind, "QR오더 견적기")
        # 풍선 애니메이션 + 토스트 + 세션 마커 (다음 rerun 에서 큰 박스로 안내)
        st.balloons()
        st.toast(f"✅ {len(new_products)}개 상품 저장 완료", icon="✅")
        st.session_state[f"_catalog_saved_at_{catalog_kind}"] = (
            _now_kst().strftime("%H:%M:%S"), len(new_products), page_name
        )
        st.rerun()

    # (저장 완료 안내는 위쪽 — 저장 버튼 직후로 이동됨)

    # ── 미저장 변경 감지 → 페이지 이탈 시 브라우저 경고 ──
    current_snapshot = [
        {
            "code": _safe_str(r.get("code")).strip(),
            "name": _safe_str(r.get("name")).strip(),
            "description": _safe_str(r.get("description")).strip(),
            "unit_price": (int(r.get("unit_price"))
                           if pd.notna(r.get("unit_price"))
                           and r.get("unit_price") not in ("", None)
                           else 0),
            "currency": _safe_str(r.get("currency")).strip() or "KRW",
            "billing_type": ("deferred_percent"
                             if _is_rate_billing(_safe_str(r.get("billing_type")))
                             else "fixed"),
        }
        for _, r in edited.iterrows()
        if _safe_str(r.get("name")).strip()
    ]
    dirty = _qr_catalog_signature(current_snapshot) != st.session_state.get(sig_key, "")
    if dirty:
        st.caption("⚠ **미저장 변경사항이 있습니다.** 우측 상단 '💾 저장' 버튼을 눌러주세요.")
        # 사용자가 다시 편집했으므로 이전 저장 완료 메시지 제거
        st.session_state.pop(f"_catalog_saved_at_{catalog_kind}", None)
    _inject_beforeunload(dirty)


def _mc_catalog_signature(products: list[dict]) -> str:
    return json.dumps(products, ensure_ascii=False, sort_keys=True, default=str)


def _render_membership_catalog_editor():
    """멤버십 견적서용 계층 카탈로그 편집기 (구분/분류 컬럼 포함)."""
    products = _load_membership_products()
    if "_mc_catalog_original_sig" not in st.session_state:
        st.session_state["_mc_catalog_original_sig"] = _mc_catalog_signature(products)

    if not products:
        df = pd.DataFrame([{
            "code": "",
            "section": "멤버십 클라우드",
            "subcategory": "초기구축비",
            "name": "(여기를 클릭해서 수정)",
            "billing_period": "1회성",
            "unit_price": 0,
            "unit_price_text": "",
            "default_amount_text": "",
            "notes": "",
        }])
        original_extra = {}
    else:
        df = pd.DataFrame(products)
        # 사용자 표에서 다루지 않을 필드(sub_items, name_detail 등)는 보존
        original_extra = {p["code"]: {k: v for k, v in p.items()
                                       if k not in {"code", "section", "subcategory",
                                                    "name", "billing_period", "unit_price",
                                                    "unit_price_text", "default_amount_text",
                                                    "notes"}}
                          for p in products if p.get("code")}
        # 누락 가능한 컬럼 보완
        for col, default in [
            ("section", ""), ("subcategory", ""), ("billing_period", ""),
            ("unit_price", None), ("unit_price_text", ""),
            ("default_amount_text", ""), ("notes", ""), ("code", ""),
        ]:
            if col not in df.columns:
                df[col] = default

    # 기존 카탈로그에 등록된 구분들 + 기본값
    existing_sections = sorted({p.get("section", "") for p in products if p.get("section")})
    section_options = existing_sections or ["멤버십 클라우드", "오더 솔루션"]

    # ── 상단 액션 바: 캡션(좌) + 저장 버튼(우) ──
    cap_col, save_col = st.columns([5, 1.3])
    with cap_col:
        st.caption(
            "구분(예: 멤버십 클라우드)과 분류(초기구축비/사용료/옵션)로 묶인 상품 목록입니다. "
            "행 앞 **☰** 아이콘 영역을 잡아 위·아래로 드래그하면 순서를 바꿀 수 있어요."
        )
    with save_col:
        save_clicked = st.button(
            "💾 멤버십 품목 저장", type="primary",
            use_container_width=True, key="save_mc_catalog",
        )

    # ── 드래그앤드롭 순서 변경 ──
    if len(df) > 1:
        with st.expander("🔃 행 순서 변경 (드래그앤드롭)", expanded=False):
            st.caption("왼쪽 **☰** 아이콘 영역을 잡아 위·아래로 드래그하세요.")
            try:
                from streamlit_sortables import sort_items
                cat_labels = [
                    f"☰  [{i + 1}]  [{df.iloc[i]['section'] or '?'}/{df.iloc[i]['subcategory'] or '?'}]  "
                    f"{df.iloc[i]['name'] or '(이름 없음)'}"
                    for i in range(len(df))
                ]
                sorted_labels = sort_items(
                    cat_labels, direction="vertical",
                    key="mc_cat_sort_dnd",
                )
                if sorted_labels and sorted_labels != cat_labels:
                    try:
                        new_order = [
                            int(s.split("[", 1)[1].split("]", 1)[0]) - 1
                            for s in sorted_labels
                        ]
                        df = df.iloc[new_order].reset_index(drop=True)
                    except (ValueError, IndexError):
                        pass
            except ImportError:
                st.warning("드래그앤드롭 컴포넌트가 설치되지 않았습니다.")

    edited = st.data_editor(
        df[["code", "section", "subcategory", "name", "billing_period",
            "unit_price", "unit_price_text", "default_amount_text", "notes"]],
        column_config={
            "code": st.column_config.TextColumn(
                "코드 (선택사항)",
                help="고유 식별자 (예: MC-INIT-SERVER). 비워두면 자동 생성됩니다.",
            ),
            "section": st.column_config.SelectboxColumn(
                "구분", help="대분류 (시나리오 내 그룹)",
                options=section_options + ["기타"],
            ),
            "subcategory": st.column_config.SelectboxColumn(
                "분류", help="중분류",
                options=_DEFAULT_SUBCATEGORIES + ["기타"],
            ),
            "name": st.column_config.TextColumn(
                "상세 구분", help="견적서에 표시되는 항목명", required=True,
            ),
            "billing_period": st.column_config.SelectboxColumn(
                "기간", options=["", "1회성", "매월", "발생시", "발생월", "1개당"],
                help="청구 주기",
            ),
            "unit_price": st.column_config.NumberColumn(
                "단가 (숫자)", step=100000, format="₩%,d",
                help="숫자 단가. 할인은 음수로 (예: -1000000). 텍스트는 옆 칸 사용.",
            ),
            "unit_price_text": st.column_config.TextColumn(
                "단가(텍스트)",
                help="예: '투입기간 X SW개발자 임금', '무상 제공', '7.9원 / 건'",
            ),
            "default_amount_text": st.column_config.TextColumn(
                "기본 금액 텍스트",
                help="비우면 자동 계산. '후청구', '협의 금액', '무상제공' 등 텍스트 입력 가능",
            ),
            "notes": st.column_config.TextColumn(
                "비고", width="large",
                help="Alt+Enter 로 줄바꿈 입력 가능. PDF에도 줄바꿈 반영.",
            ),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=720,
        key="catalog_editor_mc",
    )

    if save_clicked:
        new_products = []
        for idx, row in edited.iterrows():
            name = _safe_str(row.get("name")).strip()
            if not name:
                continue
            code = _safe_str(row.get("code")).strip()
            if not code:
                base = "".join(c for c in name if c.isalnum() or c in "_-")[:20] or "MC"
                code = f"{base}-{idx + 1}"
            item: dict = {
                "code": code,
                "section": _safe_str(row.get("section")).strip(),
                "subcategory": _safe_str(row.get("subcategory")).strip(),
                "name": name,
            }
            bp = row.get("billing_period")
            if isinstance(bp, str) and bp.strip():
                item["billing_period"] = bp.strip()
            up = row.get("unit_price")
            if pd.notna(up) and up not in ("", None):
                try:
                    item["unit_price"] = float(up)
                except (TypeError, ValueError):
                    item["unit_price"] = None
            else:
                item["unit_price"] = None
            upt = row.get("unit_price_text")
            if isinstance(upt, str) and upt.strip():
                item["unit_price_text"] = upt.strip()
            amt_text = row.get("default_amount_text")
            if isinstance(amt_text, str) and amt_text.strip():
                item["default_amount_text"] = amt_text.strip()
            notes = row.get("notes")
            if isinstance(notes, str) and notes.strip():
                item["notes"] = notes.strip()
            extra = original_extra.get(code, {})
            item.update(extra)
            new_products.append(item)
        _save_membership_products(new_products)
        st.session_state["_mc_catalog_original_sig"] = _mc_catalog_signature(new_products)
        st.success(
            f"✅ {len(new_products)}개 멤버십 상품 저장 완료. "
            "'멤버십 견적서 작성' 페이지로 가면 즉시 반영됩니다."
        )

    # ── 미저장 변경 감지 → 페이지 이탈 시 브라우저 경고 ──
    snapshot = []
    for idx, r in edited.iterrows():
        name = _safe_str(r.get("name")).strip()
        if not name:
            continue
        up_v = r.get("unit_price")
        snapshot.append({
            "code": _safe_str(r.get("code")).strip(),
            "section": _safe_str(r.get("section")).strip(),
            "subcategory": _safe_str(r.get("subcategory")).strip(),
            "name": name,
            "billing_period": _safe_str(r.get("billing_period")).strip(),
            "unit_price": (float(up_v)
                           if pd.notna(up_v) and up_v not in ("", None)
                           else None),
            "unit_price_text": _safe_str(r.get("unit_price_text")).strip(),
            "default_amount_text": _safe_str(r.get("default_amount_text")).strip(),
            "notes": _safe_str(r.get("notes")).strip(),
        })
    dirty = _mc_catalog_signature(snapshot) != st.session_state.get(
        "_mc_catalog_original_sig", ""
    )
    if dirty:
        st.caption("⚠ **미저장 변경사항이 있습니다.** 우측 상단 '💾 저장' 버튼을 눌러주세요.")
    _inject_beforeunload(dirty)


# ═════════════════════════════════════════════════════════════
# 페이지 3: 설정 (브랜드 정보 + 양식 라벨)
# ═════════════════════════════════════════════════════════════

def render_settings_page():
    st.title("⚙ 설정")
    tab_brand, tab_labels = st.tabs(["🏢 브랜드 정보", "📐 양식 라벨/문구"])

    with tab_brand:
        _render_brand_settings()
    with tab_labels:
        _render_label_settings()


def _render_brand_settings():
    brand_ids = _list_brands()
    if not brand_ids:
        st.warning("등록된 브랜드가 없습니다.")
        return

    target_brand_id = st.selectbox(
        "편집할 브랜드", brand_ids,
        disabled=len(brand_ids) <= 1,
    )
    try:
        brand = load_brand(PROJECT_ROOT, target_brand_id)
    except FileNotFoundError as e:
        st.error(f"브랜드 파일을 찾을 수 없습니다: {e}")
        return

    st.subheader("회사 기본 정보")
    c1, c2 = st.columns(2)
    with c1:
        name_ko = st.text_input("회사명 (한글)", brand.company.name_ko)
        registration_number = st.text_input("사업자등록번호", brand.company.registration_number)
        ceo = st.text_input("대표자", brand.company.ceo)
    with c2:
        name_en = st.text_input("회사명 (영문, 선택)", brand.company.name_en or "")
        phone = st.text_input("대표 전화 (선택)", brand.company.phone or "")
        email = st.text_input("대표 이메일 (선택)", brand.company.email or "")
    address = st.text_input("주소", brand.company.address or "")

    st.subheader("담당자 정보")
    cp_c1, cp_c2 = st.columns(2)
    cp = brand.contact_person
    with cp_c1:
        cp_name = st.text_input("담당자 이름", (cp.name if cp else ""))
        cp_title = st.text_input("직책", (cp.title if cp else "") or "")
    with cp_c2:
        cp_phone = st.text_input("담당자 전화", (cp.phone if cp else "") or "")
        cp_email = st.text_input("담당자 이메일", (cp.email if cp else "") or "")

    st.subheader("서명자 (견적서/계약서 서명란)")
    sig_c1, sig_c2 = st.columns(2)
    with sig_c1:
        signer_name = st.text_input("서명자명", brand.signature.signer_name)
    with sig_c2:
        signer_title = st.text_input("서명자 직책", brand.signature.signer_title)

    st.subheader("입금 계좌 (선택)")
    ba = brand.bank_account
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        bank = st.text_input("은행", (ba.bank if ba else ""))
    with bc2:
        account_number = st.text_input("계좌번호", (ba.account_number if ba else ""))
    with bc3:
        account_holder = st.text_input("예금주", (ba.account_holder if ba else "") or "")

    st.subheader("색상 (CI)")
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        primary = st.color_picker("Primary (주 색상)", brand.branding.colors.primary)
    with cc2:
        accent = st.color_picker("Accent (강조)", brand.branding.colors.accent)
    with cc3:
        text_color = st.color_picker("Text", brand.branding.colors.text)
    font_family = st.text_input("폰트 패밀리", brand.branding.font_family)

    st.divider()
    if st.button("💾 브랜드 정보 저장", type="primary", use_container_width=True):
        new_dict = {
            "brand_id": brand.brand_id,
            "company": {
                "name_ko": name_ko,
                "name_en": name_en or None,
                "registration_number": registration_number,
                "ceo": ceo,
                "address": address or None,
                "phone": phone or None,
                "email": email or None,
            },
            "branding": {
                "logo_path": brand.branding.logo_path,
                "colors": {
                    "primary": primary,
                    "accent": accent,
                    "text": text_color,
                },
                "font_family": font_family,
            },
            "signature": {
                "signer_name": signer_name,
                "signer_title": signer_title,
            },
            "contact_person": ({
                "name": cp_name,
                "title": cp_title or None,
                "phone": cp_phone or None,
                "email": cp_email or None,
            } if cp_name else None),
            "bank_account": ({
                "bank": bank,
                "account_number": account_number,
                "account_holder": account_holder or None,
            } if bank and account_number else None),
            "footer_text": brand.footer_text or "",
        }
        _save_brand(brand.brand_id, new_dict)
        st.success("✅ 브랜드 정보 저장 완료")


def _render_label_settings():
    labels = load_labels(PROJECT_ROOT)
    ql = labels.quote

    st.subheader("기본 값")
    c1, c2 = st.columns(2)
    with c1:
        vat_rate_pct = st.number_input(
            "VAT율 (%)", value=float(ql.vat_rate * 100),
            min_value=0.0, max_value=100.0, step=0.5,
            help="예: 10 = 10%",
        )
    with c2:
        default_validity_days = st.number_input(
            "유효기간 기본 일수", value=ql.default_validity_days,
            min_value=1, max_value=365,
        )

    st.subheader("자동 안내 문구")
    st.caption("`{valid_until}`, `{bank}`, `{account_number}`, `{account_holder}` 같은 자리표시자가 자동으로 채워집니다. "
               "**빈 칸으로 두면** 해당 자동 안내가 표시되지 않습니다.")
    validity_template = st.text_input(
        "유효기간 자동 문구",
        value=ql.auto_notices.validity_template,
        placeholder="본 견적의 유효기간은 {valid_until}까지입니다.",
    )
    bank_template = st.text_input(
        "입금 계좌 자동 문구",
        value=ql.auto_notices.bank_account_template,
        placeholder="입금 계좌: {bank} {account_number} (예금주: {account_holder})",
    )

    st.subheader("고정 안내 문구")
    st.caption("모든 견적서의 '기타 안내' 섹션에 항상 표시되는 문구. 한 줄에 하나씩 입력.")
    static_notices_text = st.text_area(
        "고정 안내",
        value="\n".join(ql.static_notices),
        height=100,
        label_visibility="collapsed",
    )
    static_notices = [
        line.strip() for line in static_notices_text.splitlines() if line.strip()
    ]

    st.subheader("표시 라벨")
    l1, l2, l3 = st.columns(3)
    with l1:
        subtotal_label = st.text_input("공급가액 라벨", ql.labels.subtotal)
        counterparty_section = st.text_input("수신처 섹션", ql.labels.counterparty_section)
    with l2:
        vat_label = st.text_input("부가세 라벨", ql.labels.vat)
        subject_prefix = st.text_input("건명 prefix", ql.labels.subject_prefix)
    with l3:
        total_label = st.text_input("합계 라벨", ql.labels.total)
        etc_notice_section = st.text_input("기타 안내 섹션", ql.labels.etc_notice_section)

    vat_separate = st.text_input(
        "VAT 별도 표시 (비우면 미표시)",
        value=ql.labels.vat_separate_notice,
        placeholder="* VAT 별도",
    )

    st.subheader("표 컬럼 헤더")
    th = ql.table_headers
    th1, th2, th3, th4 = st.columns(4)
    with th1:
        h_name = st.text_input("품목 컬럼명", th.name)
        h_unit_price = st.text_input("단가 컬럼명", th.unit_price)
    with th2:
        h_description = st.text_input("설명 컬럼명", th.description)
        h_amount = st.text_input("금액 컬럼명", th.amount)
    with th3:
        h_qty = st.text_input("수량 컬럼명", th.qty)
        h_notes = st.text_input("비고 컬럼명", th.notes)
    with th4:
        h_period = st.text_input("기간 컬럼명", th.period)

    st.subheader("문서 제목")
    t1, t2 = st.columns(2)
    with t1:
        quote_title = st.text_input("견적서 제목", ql.title)
    with t2:
        contract_title = st.text_input("계약서 제목", labels.contract.title)

    st.divider()
    if st.button("💾 양식 설정 저장", type="primary", use_container_width=True):
        new_labels = {
            "_doc": "양식의 라벨·문구·기본값을 관리하는 파일입니다.",
            "quote": {
                "title": quote_title,
                "vat_rate": round(vat_rate_pct / 100, 4),
                "default_validity_days": int(default_validity_days),
                "labels": {
                    "counterparty_section": counterparty_section,
                    "subject_prefix": subject_prefix,
                    "etc_notice_section": etc_notice_section,
                    "vat_separate_notice": vat_separate,
                    "subtotal": subtotal_label,
                    "vat": vat_label,
                    "total": total_label,
                },
                "table_headers": {
                    "name": h_name,
                    "description": h_description,
                    "qty": h_qty,
                    "period": h_period,
                    "unit_price": h_unit_price,
                    "amount": h_amount,
                    "notes": h_notes,
                },
                "auto_notices": {
                    "validity_template": validity_template,
                    "bank_account_template": bank_template,
                },
                "static_notices": static_notices,
            },
            "contract": {
                "title": contract_title,
                "preamble_template": labels.contract.preamble_template,
                "labels": labels.contract.labels.model_dump(),
            },
        }
        _save_labels(new_labels)
        st.success("✅ 양식 설정 저장 완료. 새로 생성하는 견적서/계약서에 반영됩니다.")


# ═════════════════════════════════════════════════════════════
# 페이지 4: 멤버십 클라우드 견적서 작성
# ═════════════════════════════════════════════════════════════

# 분류 미리보기 옵션 (사용자가 직접 입력도 가능)
# '할인' 은 음수 단가로 등록 → 총비용에서 자동 차감됨
_DEFAULT_SUBCATEGORIES = ["초기구축비", "사용료", "옵션", "할인"]


@st.cache_data
def _load_membership_products() -> list[dict]:
    path = PROJECT_ROOT / "catalog" / "membership_products.json"
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8")).get("products", [])


def _save_membership_products(products: list[dict]) -> None:
    path = PROJECT_ROOT / "catalog" / "membership_products.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    payload_bytes = json.dumps(
        {"products": products}, ensure_ascii=False, indent=2
    ).encode("utf-8")
    path.write_bytes(payload_bytes)
    _load_membership_products.clear()
    # GitHub 자동 commit
    ok, msg = _push_to_github(
        "catalog/membership_products.json", payload_bytes,
        f"품목 관리(membership): {len(products)}개 저장",
    )
    st.session_state["_github_sync_membership"] = (ok, msg)


def _load_membership_sample() -> dict:
    """샘플 데이터 (사용자가 '샘플 불러오기' 버튼을 눌렀을 때 사용)."""
    path = PROJECT_ROOT / "data" / "membership_quotes" / "sample.json"
    if path.exists():
        sample = json.loads(path.read_text(encoding="utf-8"))
        sample.setdefault("issued_date", date.today().isoformat())
        return sample
    return _empty_membership_state()


def _empty_membership_state() -> dict:
    """빈 초기 상태 — QR 견적서와 동일하게 placeholder만 보이는 상태로 시작.
    supplier(회사 정보)는 brand.json에서 자동 로드되므로 None."""
    return {
        "document_id": "",
        "title": "멤버십 클라우드 견적서",
        "issued_date": date.today().isoformat(),
        "counterparty": {
            "label": "제휴사", "name": "",
            "address": None, "ceo": None, "contact": None,
        },
        "supplier": None,
        "scenarios": [],
        "remarks": [
            "견적유효 : 견적일로부터 15일",
            "결제조건 : 현금결제 (귀사 결제조건)",
        ],
    }


def _ensure_membership_state():
    if "mc_doc" not in st.session_state:
        st.session_state.mc_doc = _empty_membership_state()
    if "mc_items_df" not in st.session_state:
        st.session_state.mc_items_df = _mc_empty_items_df()


# ── 멤버십 평면 품목 표 — QR 견적서와 동일 패턴 ──

_MC_DEFAULT_SECTIONS = ["멤버십 클라우드", "오더 솔루션"]
_MC_DEFAULT_CATEGORIES = ["초기구축비", "사용료", "옵션"]
_MC_DEFAULT_PERIODS = ["", "1회성", "매월", "발생시", "발생월", "1개당"]

MC_ITEM_COLUMNS = [
    "분류", "구분", "카테고리", "상세 구분", "기간", "단가",
    "할인율(%)", "할인금액", "비고",
]
MC_DISPLAY_COLUMNS = [
    "분류", "구분", "카테고리", "상세 구분", "기간", "단가",
    "할인율(%)", "할인금액", "공급가", "비고",
]


def _mc_empty_items_df() -> pd.DataFrame:
    return pd.DataFrame(columns=MC_ITEM_COLUMNS).astype({
        "분류": "string", "구분": "string", "카테고리": "string",
        "상세 구분": "string", "기간": "string",
        "단가": "Int64", "할인율(%)": "Int64", "할인금액": "Int64",
        "비고": "string",
    })


def _mc_normal_items_sum(df) -> int:
    if df is None or df.empty:
        return 0
    total = 0
    for _, row in df.iterrows():
        if row.get("분류") == ITEM_KIND_DISCOUNT:
            continue
        price = row.get("단가")
        if not (pd.notna(price) and price):
            continue
        try:
            gross = int(price)
        except (TypeError, ValueError):
            continue
        # 항목별 할인 적용 (할인금액 우선)
        d_amt = row.get("할인금액")
        d_rate = row.get("할인율(%)")
        discount = 0
        if pd.notna(d_amt) and d_amt:
            try:
                discount = int(d_amt)
            except (TypeError, ValueError):
                pass
        elif pd.notna(d_rate) and d_rate:
            try:
                discount = int(round(gross * float(d_rate) / 100))
            except (TypeError, ValueError):
                pass
        total += gross - discount
    return total


def _mc_row_amount(row, df=None):
    """멤버십 평면 표 한 행의 공급가 계산. QR _row_amount 와 동일 로직."""
    if row.get("분류") != ITEM_KIND_DISCOUNT:
        # 일반 품목
        price = row.get("단가")
        if not (pd.notna(price) and price):
            return None
        try:
            gross = int(price)
        except (TypeError, ValueError):
            return None
        d_amt = row.get("할인금액")
        d_rate = row.get("할인율(%)")
        discount = 0
        if pd.notna(d_amt) and d_amt:
            try:
                discount = int(d_amt)
            except (TypeError, ValueError):
                pass
        elif pd.notna(d_rate) and d_rate:
            try:
                discount = int(round(gross * float(d_rate) / 100))
            except (TypeError, ValueError):
                pass
        return gross - discount

    # 할인 행 — 단가/할인금액 = 고정 차감, 할인율(%) = 일반 품목 합 × 비율
    deduct = 0
    price = row.get("단가")
    if pd.notna(price) and price:
        try:
            deduct += abs(int(price))
        except (TypeError, ValueError):
            pass
    d_amt = row.get("할인금액")
    if pd.notna(d_amt) and d_amt:
        try:
            deduct += abs(int(d_amt))
        except (TypeError, ValueError):
            pass
    d_rate = row.get("할인율(%)")
    if pd.notna(d_rate) and d_rate:
        try:
            base = _mc_normal_items_sum(df) if df is not None else 0
            deduct += int(round(base * float(d_rate) / 100))
        except (TypeError, ValueError):
            pass
    if deduct == 0:
        return None
    return -deduct


def _mc_add_catalog_row(product: dict):
    df = st.session_state.mc_items_df
    new_row = {
        "분류": ITEM_KIND_NORMAL,
        "구분": product.get("section", "") or "",
        "카테고리": product.get("subcategory", "") or "",
        "상세 구분": product.get("name", "") or "",
        "기간": product.get("billing_period", "") or "",
        "단가": (int(product.get("unit_price"))
                 if product.get("unit_price") not in (None, "") else None),
        "할인율(%)": None, "할인금액": None,
        "비고": product.get("notes", "") or "",
    }
    # 텍스트 단가/기본금액 텍스트는 비고에 함께 표시
    extras = []
    if product.get("unit_price_text"):
        extras.append(f"단가: {product['unit_price_text']}")
    if product.get("default_amount_text"):
        extras.append(f"금액: {product['default_amount_text']}")
    if extras:
        new_row["비고"] = " · ".join(extras + ([new_row["비고"]] if new_row["비고"] else []))
    st.session_state.mc_items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True,
    )


def _mc_add_blank_row():
    df = st.session_state.mc_items_df
    blank = {c: None for c in MC_ITEM_COLUMNS}
    blank["분류"] = ITEM_KIND_NORMAL
    st.session_state.mc_items_df = pd.concat(
        [df, pd.DataFrame([blank])], ignore_index=True,
    )


def _mc_add_discount_row():
    df = st.session_state.mc_items_df
    new_row = {
        "분류": ITEM_KIND_DISCOUNT,
        "구분": "", "카테고리": "",
        "상세 구분": "할인",
        "기간": "1회성",
        "단가": None,
        "할인율(%)": None, "할인금액": None,
        "비고": "협상가",
    }
    st.session_state.mc_items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True,
    )


def _mc_reset_items():
    st.session_state.mc_items_df = _mc_empty_items_df()


def _mc_items_df_to_scenario(df: pd.DataFrame, scenario_name: str = "기본") -> dict:
    """평면 df → 단일 시나리오 (구분별 섹션 · 카테고리별 카테고리 · 항목들)."""
    sections: dict[str, dict] = {}
    section_order: list[str] = []
    for _, row in df.iterrows():
        name = (row.get("상세 구분") or "").strip() if isinstance(row.get("상세 구분"), str) else ""
        if not name:
            continue
        is_disc = row.get("분류") == ITEM_KIND_DISCOUNT
        sec_name = (row.get("구분") or "").strip() if isinstance(row.get("구분"), str) else ""
        cat_name = (row.get("카테고리") or "").strip() if isinstance(row.get("카테고리"), str) else ""
        if is_disc:
            sec_name = sec_name or "할인"
            cat_name = cat_name or "할인"
        else:
            sec_name = sec_name or "기타"
            cat_name = cat_name or "기타"
        sec_key = sec_name
        if sec_key not in sections:
            sections[sec_key] = {"name": sec_name, "categories_map": {}, "categories_order": []}
            section_order.append(sec_key)
        sec = sections[sec_key]
        if cat_name not in sec["categories_map"]:
            sec["categories_map"][cat_name] = {"name": cat_name, "items": []}
            sec["categories_order"].append(cat_name)
        item: dict = {"name": name}
        bp = row.get("기간")
        if isinstance(bp, str) and bp.strip():
            item["billing_period"] = bp.strip()
        price = row.get("단가")
        if pd.notna(price) and price not in ("", None):
            try:
                v = float(price)
                if is_disc:
                    v = -abs(v)
                item["unit_price"] = v
            except (TypeError, ValueError):
                pass
        d_rate = row.get("할인율(%)")
        if pd.notna(d_rate) and d_rate and not is_disc:
            try:
                item["discount_rate"] = float(d_rate) / 100
            except (TypeError, ValueError):
                pass
        d_amt = row.get("할인금액")
        if pd.notna(d_amt) and d_amt and not is_disc:
            try:
                item["discount_amount"] = float(d_amt)
            except (TypeError, ValueError):
                pass
        # 할인 행이고 할인율(%) 만 있는 경우: 일반 품목 합의 N% 를 음수 단가로 환산
        if is_disc and item.get("unit_price") is None:
            base_sum = _mc_normal_items_sum(df)
            d_rate_val = row.get("할인율(%)")
            d_amt_val = row.get("할인금액")
            deduct = 0.0
            if pd.notna(d_amt_val) and d_amt_val:
                deduct += abs(float(d_amt_val))
            if pd.notna(d_rate_val) and d_rate_val:
                deduct += base_sum * float(d_rate_val) / 100.0
            if deduct > 0:
                item["unit_price"] = -deduct
        notes = row.get("비고")
        if isinstance(notes, str) and notes.strip():
            item["notes"] = notes.strip()
        sec["categories_map"][cat_name]["items"].append(item)

    out_sections = []
    for sec_key in section_order:
        sec = sections[sec_key]
        cats = []
        for cat_name in sec["categories_order"]:
            cat = sec["categories_map"][cat_name]
            cats.append({
                "name": cat["name"],
                "items": cat["items"],
                "show_subtotal": cat["name"] != "옵션",
            })
        out_sections.append({
            "name": sec["name"],
            "categories": cats,
            "show_section_total": True,
        })
    return {
        "name": scenario_name,
        "sections": out_sections,
        "show_grand_total": True,
    }


def _mc_items_to_df(items: list[dict]) -> pd.DataFrame:
    """[{name, billing_period, unit_price, ...}, ...] → DataFrame."""
    if not items:
        return pd.DataFrame(columns=[
            "분류", "상세 구분", "기간", "단가", "단가(텍스트)",
            "할인율(%)", "할인금액", "금액 텍스트", "비고",
        ])
    rows = []
    for it in items:
        dr = it.get("discount_rate")
        # 분류='할인' 행은 표 표시 시 단가를 양수로 노출 (자동 차감 처리는 저장 단계에서)
        up = it.get("unit_price")
        sub = (it.get("_subcategory") or "").strip()
        if sub == "할인" and up is not None:
            try:
                up = abs(float(up))
            except (TypeError, ValueError):
                pass
        rows.append({
            "분류": sub,
            "상세 구분": it.get("name", ""),
            "기간": it.get("billing_period", ""),
            "단가": up,
            "단가(텍스트)": it.get("unit_price_text", ""),
            "할인율(%)": int(dr * 100) if dr else None,
            "할인금액": it.get("discount_amount"),
            "금액 텍스트": it.get("amount_text", ""),
            "비고": it.get("notes", ""),
        })
    return pd.DataFrame(rows)


def _df_to_section_categories(df: pd.DataFrame) -> list[dict]:
    """편집된 DataFrame → categories 리스트 (분류별 그룹).
    설정된 분류가 없는 행은 '기타' 분류로 묶음.
    """
    if df.empty:
        return []
    # 분류별 그룹 유지 순서
    seen_order = []
    groups: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        name = (row.get("상세 구분") or "").strip() if isinstance(row.get("상세 구분"), str) else ""
        if not name:
            continue
        sub = (row.get("분류") or "").strip() if isinstance(row.get("분류"), str) else ""
        if not sub:
            sub = "기타"
        if sub not in groups:
            groups[sub] = []
            seen_order.append(sub)
        item: dict = {"name": name}
        billing = row.get("기간")
        if isinstance(billing, str) and billing.strip():
            item["billing_period"] = billing.strip()
        up = row.get("단가")
        if pd.notna(up) and up not in ("", None):
            try:
                up_val = float(up)
                # 분류='할인' 이면 단가 양수 입력도 음수로 강제 (자동 차감)
                if sub == "할인":
                    up_val = -abs(up_val)
                item["unit_price"] = up_val
            except (TypeError, ValueError):
                pass
        upt = row.get("단가(텍스트)")
        if isinstance(upt, str) and upt.strip():
            item["unit_price_text"] = upt.strip()
        dr = row.get("할인율(%)")
        if pd.notna(dr) and dr not in ("", None):
            try:
                drv = float(dr)
                if drv > 0:
                    item["discount_rate"] = drv / 100
            except (TypeError, ValueError):
                pass
        da = row.get("할인금액")
        if pd.notna(da) and da not in ("", None):
            try:
                dav = float(da)
                if dav > 0:
                    item["discount_amount"] = dav
            except (TypeError, ValueError):
                pass
        amt_text = row.get("금액 텍스트")
        if isinstance(amt_text, str) and amt_text.strip():
            item["amount_text"] = amt_text.strip()
        notes = row.get("비고")
        if isinstance(notes, str) and notes.strip():
            item["notes"] = notes.strip()
        groups[sub].append(item)

    cats = []
    for sub in seen_order:
        cats.append({
            "name": sub,
            "items": groups[sub],
            "show_subtotal": sub != "옵션",  # 옵션은 합계 미표시 기본
        })
    return cats


def _section_to_flat_items(section: dict) -> list[dict]:
    """section의 categories→items 를 평면화. _subcategory 필드 추가."""
    flat = []
    for cat in section.get("categories", []):
        for it in cat.get("items", []):
            flat.append({**it, "_subcategory": cat.get("name", "")})
    return flat


def _add_blank_row_to_section(section: dict,
                              default_subcategory: str = "초기구축비") -> None:
    """빈 항목 한 줄 추가. 사용자가 수정해서 진짜 항목으로 만듦."""
    section.setdefault("categories", [])
    target_cat = None
    for cat in section["categories"]:
        if cat.get("name") == default_subcategory:
            target_cat = cat
            break
    if target_cat is None:
        target_cat = {
            "name": default_subcategory,
            "items": [],
            "show_subtotal": default_subcategory != "옵션",
        }
        section["categories"].append(target_cat)
    target_cat.setdefault("items", []).append({
        "name": "(새 항목 — 클릭해서 수정)",
    })


def _add_discount_row_to_section(section: dict) -> None:
    """'할인' 분류에 할인 행 추가. 단가는 양수 입력 시 자동 차감."""
    section.setdefault("categories", [])
    target_cat = None
    for cat in section["categories"]:
        if cat.get("name") == "할인":
            target_cat = cat
            break
    if target_cat is None:
        target_cat = {
            "name": "할인",
            "items": [],
            "show_subtotal": True,
        }
        section["categories"].append(target_cat)
    target_cat.setdefault("items", []).append({
        "name": "협상 할인 (이름 수정 가능)",
        "billing_period": "1회성",
        "unit_price": 1000000,
        "notes": "협상에 따라 단가 조정",
    })


def _build_mc_document(state: dict) -> MembershipQuoteDocument:
    """session_state.mc_doc 를 Pydantic 객체로 변환."""
    return MembershipQuoteDocument.model_validate(state)


def render_membership_quote_page():
    _mc_autosave_load_once()
    _ensure_membership_state()
    state = st.session_state.mc_doc
    products = _load_membership_products()
    soffice_available = find_soffice() is not None

    # 사이드바: 발행 정보
    with st.sidebar:
        st.divider()
        st.header("⚙ 발행 정보")
        try:
            cur_issued = date.fromisoformat(state.get("issued_date", date.today().isoformat()))
        except (TypeError, ValueError):
            cur_issued = date.today()
        new_issued = st.date_input("발행일", value=cur_issued, key="mc_issued_date")
        state["issued_date"] = new_issued.isoformat()
        if not soffice_available:
            st.warning("⚠ LibreOffice 미감지 — PDF 변환 건너뜀")

    st.title("🏢 멤버십 클라우드 견적서 작성")

    # ─── 최근 다운로드 견적서 / 견적서 표본 ───
    _render_mc_recent_panel()
    _render_mc_template_panel()

    # ─── 0. 발행 담당자 정보 ───
    try:
        _brand_for_default = load_brand(PROJECT_ROOT, "softment")
        _default_cp = _brand_for_default.contact_person
    except FileNotFoundError:
        _default_cp = None

    st.subheader("0. 발행 담당자 정보")
    st.caption("이 견적서를 발행하는 우리 쪽 담당자. 회사(우리)의 '담당자/연락처/이메일' 칸에 표시됩니다.")
    # Tab 좌→우→다음 줄 순으로 이동하도록 행 단위 columns
    ic_r1_l, ic_r1_r = st.columns(2)
    with ic_r1_l:
        issuer_name = st.text_input(
            "담당자명", key="mc_issuer_name",
            value=(_default_cp.name if _default_cp else ""),
            placeholder="예: 박지은",
        )
    with ic_r1_r:
        issuer_title = st.text_input(
            "직책", key="mc_issuer_title",
            value=(_default_cp.title if _default_cp and _default_cp.title else ""),
            placeholder="예: QR사업부 매니저",
        )
    ic_r2_l, ic_r2_r = st.columns(2)
    with ic_r2_l:
        issuer_phone = st.text_input(
            "연락처", key="mc_issuer_phone",
            value=(_default_cp.phone if _default_cp and _default_cp.phone else ""),
            placeholder="예: 010-0000-0000",
        )
    with ic_r2_r:
        issuer_email = st.text_input(
            "이메일", key="mc_issuer_email",
            value=(_default_cp.email if _default_cp and _default_cp.email else ""),
            placeholder="예: name@softment.co.kr",
        )

    # ─── 1. 수신처 정보 ───
    st.subheader("1. 수신처 정보")
    st.caption("회사(우리) 기본 정보는 '⚙ 설정' 의 브랜드 정보에서 자동으로 가져옵니다.")

    cp = state.setdefault("counterparty", {"label": "제휴사", "name": ""})
    # supplier 는 PDF 렌더 시 brand.json 에서 자동 채워지므로 폼에서 제거
    state["supplier"] = None

    # Tab 좌→우→다음 줄 순으로
    pc_r1_l, pc_r1_r = st.columns(2)
    with pc_r1_l:
        cp["name"] = st.text_input(
            "회사명 *", value=cp.get("name") or "", key="mc_cp_name",
            placeholder="예: 주식회사 ○○",
        )
    with pc_r1_r:
        cp["address"] = st.text_input(
            "주소", value=cp.get("address") or "", key="mc_cp_addr",
            placeholder="예: 서울특별시 ○○구 ○○로 ○○",
        )
    pc_r2_l, pc_r2_r = st.columns(2)
    with pc_r2_l:
        cp["ceo"] = st.text_input(
            "대표이사", value=cp.get("ceo") or "", key="mc_cp_ceo",
            placeholder="예: 홍길동",
        )
    with pc_r2_r:
        cp["contact"] = st.text_input(
            "담당자", value=cp.get("contact") or "", key="mc_cp_contact",
            placeholder="예: 김담당 (구매팀장)",
        )

    # ─── 2. 건명 ───
    st.subheader("2. 건명")
    state["title"] = st.text_input(
        "건명",
        value=state.get("title", "") or "",
        key="mc_title",
        placeholder="예: 멤버십 클라우드 견적서",
        label_visibility="collapsed",
    )

    # ─── 3. 품목 내역 (평면 표 — QR 견적서와 동일 UX) ───
    st.subheader("3. 품목 내역")
    st.caption(
        "💡 품목 선택 또는 '+ 빈 행' 으로 추가. 할인은 **'+ 할인행'** 으로 추가. "
        "할인 행의 **할인율(%)** 은 일반 품목 합의 N% 만큼 일괄 차감됩니다."
    )

    # 액션 바
    products_options = list(range(len(products)))
    pick_col, add_col, blank_col, disc_col, reset_col = st.columns(
        [5, 1.2, 1.2, 1.2, 1.2]
    )
    with pick_col:
        if products:
            picked_idx = st.selectbox(
                "품목에서 추가",
                options=products_options,
                index=None,
                format_func=lambda i: (
                    f"[{products[i].get('section', '')}/{products[i].get('subcategory', '')}] "
                    f"{products[i].get('name', '')}"
                ),
                placeholder="품목 선택 (검색 가능)...",
                key="mc_catalog_pick",
                label_visibility="collapsed",
            )
        else:
            st.info("품목이 비어있습니다. '품목 관리' 페이지에서 추가하세요.")
            picked_idx = None
    with add_col:
        if st.button("+ 추가", use_container_width=True,
                     disabled=picked_idx is None,
                     help="선택한 품목을 표에 추가",
                     key="mc_add_cat"):
            _mc_add_catalog_row(products[picked_idx])
            st.rerun()
    with blank_col:
        if st.button("+ 빈 행", use_container_width=True, key="mc_add_blank"):
            _mc_add_blank_row()
            st.rerun()
    with disc_col:
        if st.button("💰 + 할인행", use_container_width=True,
                     help="단가(고정) 또는 할인율(%)(일괄) 입력 가능",
                     key="mc_add_disc"):
            _mc_add_discount_row()
            st.rerun()
    with reset_col:
        if st.button("전체 초기화", use_container_width=True,
                     key="mc_items_reset_request"):
            st.session_state["_mc_items_reset_confirm"] = True

    # 전체 초기화 확인 다이얼로그
    if st.session_state.get("_mc_items_reset_confirm"):
        with st.container(border=True):
            st.markdown(
                "<div style='background:#FDECEA;border-left:4px solid #C0392B;"
                "padding:10px 14px;border-radius:6px'>"
                "<strong style='color:#C0392B'>⚠ 전체 초기화 확인</strong><br>"
                "<span style='font-size:0.9rem'>"
                "현재 작성 중인 모든 입력(제휴사·회사·품목 내역)이 사라집니다. "
                "정말 초기화할까요?</span></div>",
                unsafe_allow_html=True,
            )
            cc1, cc2, _ = st.columns([1.2, 1.2, 4])
            with cc1:
                if st.button("초기화 진행", type="primary",
                             use_container_width=True, key="mc_items_reset_yes"):
                    st.session_state.mc_doc = _empty_membership_state()
                    st.session_state.mc_items_df = _mc_empty_items_df()
                    for k in ("mc_issuer_name", "mc_issuer_title",
                              "mc_issuer_phone", "mc_issuer_email"):
                        if k in st.session_state:
                            del st.session_state[k]
                    _mc_autosave_clear()
                    st.session_state["_mc_items_reset_confirm"] = False
                    st.rerun()
            with cc2:
                if st.button("취소", use_container_width=True,
                             key="mc_items_reset_no"):
                    st.session_state["_mc_items_reset_confirm"] = False
                    st.rerun()

    # 본문 표
    if st.session_state.mc_items_df.empty:
        st.info("아직 품목이 없습니다. 품목에서 추가하거나 '+ 빈 행' 을 누르세요.")
        mc_edited_df = st.session_state.mc_items_df
    else:
        # 행 순서 변경
        df_for_order = st.session_state.mc_items_df.reset_index(drop=True)
        if len(df_for_order) > 1:
            with st.expander(f"🔃 행 순서 변경 ({len(df_for_order)}건)", expanded=False):
                move_action = None
                for i in range(len(df_for_order)):
                    name = df_for_order.at[i, "상세 구분"] or "(이름 없음)"
                    price = df_for_order.at[i, "단가"]
                    lbl = f"**[{i + 1}]** {name}"
                    if pd.notna(price) and price:
                        try:
                            lbl += f"  ·  ₩{int(price):,}"
                        except (TypeError, ValueError):
                            pass
                    with st.container(border=True):
                        rc = st.columns([0.5, 8, 1, 1])
                        with rc[0]:
                            st.markdown(
                                "<div style='font-size:1.3rem;color:#888;"
                                "text-align:center;padding-top:2px;'>☰</div>",
                                unsafe_allow_html=True,
                            )
                        with rc[1]:
                            st.write(lbl)
                        with rc[2]:
                            if st.button("⬆", key=f"mc_row_up_{i}",
                                         disabled=(i == 0),
                                         use_container_width=True):
                                move_action = ("up", i)
                        with rc[3]:
                            if st.button("⬇", key=f"mc_row_dn_{i}",
                                         disabled=(i == len(df_for_order) - 1),
                                         use_container_width=True):
                                move_action = ("dn", i)
                if move_action is not None:
                    direction, idx = move_action
                    j = idx - 1 if direction == "up" else idx + 1
                    df = st.session_state.mc_items_df.reset_index(drop=True)
                    df.iloc[[idx, j]] = df.iloc[[j, idx]].values
                    st.session_state.mc_items_df = df
                    st.rerun()

        display_df = st.session_state.mc_items_df.copy()
        display_df["분류"] = _normalize_kind_column(display_df["분류"])
        display_df["공급가"] = display_df.apply(
            lambda r: _mc_row_amount(r, df=display_df), axis=1
        ).astype("Int64")
        display_df = display_df[MC_DISPLAY_COLUMNS]

        mc_edited_df = st.data_editor(
            display_df,
            column_config={
                "분류": st.column_config.SelectboxColumn(
                    "금액 방식", options=ITEM_KINDS, required=True,
                    width="small",
                ),
                "구분": st.column_config.SelectboxColumn(
                    "구분", options=[""] + _MC_DEFAULT_SECTIONS + ["기타"],
                    width="small",
                ),
                "카테고리": st.column_config.SelectboxColumn(
                    "카테고리", options=[""] + _MC_DEFAULT_CATEGORIES + ["기타"],
                    width="small",
                ),
                "상세 구분": st.column_config.TextColumn("상세 구분", required=True, width="medium"),
                "기간": st.column_config.SelectboxColumn(
                    "기간", options=_MC_DEFAULT_PERIODS, width="small",
                ),
                "단가": st.column_config.NumberColumn(
                    "단가", step=100000, format="₩%,d", width="small",
                    help="단가(양수). '💰 할인행' 분류는 자동 차감.",
                ),
                "할인율(%)": st.column_config.NumberColumn(
                    "할인율(%)", min_value=0, max_value=100, step=1, format="%d%%",
                    width="small",
                ),
                "할인금액": st.column_config.NumberColumn(
                    "할인금액", min_value=0, step=1000, format="₩%,d", width="small",
                ),
                "공급가": st.column_config.NumberColumn(
                    "공급가", disabled=True, format="₩%,d", width="small",
                    help="단가 − 항목별 할인 (할인행은 일괄 차감)",
                ),
                "비고": st.column_config.TextColumn("비고", width="medium"),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="mc_items_editor",
        )
        edited_core = mc_edited_df[MC_ITEM_COLUMNS]
        old = st.session_state.mc_items_df.reset_index(drop=True)
        new = edited_core.reset_index(drop=True)
        changed = (
            len(old) != len(new) or
            not all((old[c].astype(object).fillna("__").tolist()
                     == new[c].astype(object).fillna("__").tolist())
                    for c in MC_ITEM_COLUMNS)
        )
        st.session_state.mc_items_df = edited_core
        if changed:
            st.rerun()

    # 합계
    if not (isinstance(mc_edited_df, pd.DataFrame) and mc_edited_df.empty):
        amounts = mc_edited_df.apply(
            lambda r: _mc_row_amount(r, df=mc_edited_df), axis=1
        )
        subtotal = int(pd.to_numeric(amounts, errors="coerce").fillna(0).sum())
        try:
            vat_rate = load_labels(PROJECT_ROOT).quote.vat_rate
        except Exception:  # noqa: BLE001
            vat_rate = 0.10
        vat = int(round(subtotal * vat_rate))
        total = subtotal + vat
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("공급가액", f"₩{subtotal:,}")
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", f"₩{vat:,}")
        m3.metric("합계 금액", f"₩{total:,}")

    # ─── 4. 기타 안내 ───
    st.subheader("4. 기타 안내")
    remarks = state.setdefault("remarks", [])
    remarks_text = st.text_area(
        "한 줄에 하나씩",
        value="\n".join(remarks),
        height=80,
        key="mc_remarks",
        label_visibility="collapsed",
    )
    state["remarks"] = [ln.strip() for ln in remarks_text.splitlines() if ln.strip()]

    # 평면 표 → 시나리오 1개로 변환하여 state 에 반영
    state["scenarios"] = [_mc_items_df_to_scenario(
        st.session_state.mc_items_df,
        scenario_name=(state.get("title") or "").strip() or "기본"
    )]

    # ─── 5. 생성 / 미리보기 ───
    st.divider()
    can_generate = bool(cp.get("name")) and not st.session_state.mc_items_df.empty
    if not can_generate:
        st.info("제휴사 회사명과 품목 최소 1건을 입력해주세요.")

    issuer_contact = dict(
        name=issuer_name.strip(),
        title=issuer_title.strip() or None,
        phone=issuer_phone.strip() or None,
        email=issuer_email.strip() or None,
    )

    btn_preview, btn_generate = st.columns(2)
    with btn_preview:
        preview_disabled = (not can_generate) or (not soffice_available)
        preview_clicked = st.button(
            "👁 미리보기", use_container_width=True,
            disabled=preview_disabled,
            help=(
                "현재 입력값으로 PDF 미리보기를 표시합니다."
                if soffice_available
                else "PDF 변환기(LibreOffice) 미감지 — 미리보기 불가"
            ),
            key="mc_btn_preview",
        )
    with btn_generate:
        generate_clicked = st.button(
            "📝 멤버십 견적서 생성", type="primary",
            use_container_width=True, disabled=not can_generate,
            key="mc_btn_generate",
        )

    if preview_clicked:
        _preview_membership_quote(state, soffice_available, issuer_contact)
    if generate_clicked:
        _generate_membership_quote(state, soffice_available, issuer_contact)

    # 자동저장 + 페이지 이탈 시 브라우저 경고
    _mc_autosave_write()
    doc_has_data = bool(state.get("scenarios")) or bool(state.get("counterparty", {}).get("name"))
    issuer_has_data = any(
        st.session_state.get(k)
        for k in ("mc_issuer_name", "mc_issuer_title", "mc_issuer_phone", "mc_issuer_email")
    )
    _inject_beforeunload(doc_has_data or issuer_has_data)


def _render_scenario_editor(s_idx: int, scenario: dict, products: list[dict]) -> None:
    """한 시나리오 탭 안의 편집기."""
    sc1, sc2 = st.columns([5, 1])
    with sc1:
        scenario["name"] = st.text_input(
            "시나리오 이름", value=scenario.get("name", ""),
            key=f"sc_name_{s_idx}",
            placeholder="예: 앱&POS 연동 (멤버십+오더)",
        )
    with sc2:
        if st.button("❌ 시나리오 삭제", key=f"sc_del_{s_idx}",
                     use_container_width=True):
            st.session_state.mc_doc["scenarios"].pop(s_idx)
            st.rerun()

    scenario["subject"] = st.text_input(
        "항목 부제 (선택)", value=scenario.get("subject") or "",
        key=f"sc_subject_{s_idx}",
        placeholder="예: 멤버십 클라우드 솔루션 + 오더 솔루션",
    )

    # 구분(섹션) 목록
    sections = scenario.setdefault("sections", [])

    sec_btn1, sec_btn2 = st.columns([1, 5])
    with sec_btn1:
        if st.button(f"+ 구분 추가", key=f"sec_add_{s_idx}",
                     use_container_width=True):
            sections.append({
                "name": "새 구분",
                "categories": [],
                "show_section_total": True,
            })
            st.rerun()

    if not sections:
        st.info("이 시나리오에 구분이 없습니다. '+ 구분 추가' 를 눌러 추가하세요.")
        return

    for sec_idx, section in enumerate(sections):
        with st.expander(f"📂 {section.get('name', '(이름 없음)')}",
                         expanded=True):
            _render_section_editor(s_idx, sec_idx, section, products)

    # 시나리오 합계 — 공급가액 / 부가세 / 합계 금액 (샘플 로드 시에도 항상 정확 반영)
    try:
        sc_obj = MembershipScenario.model_validate(scenario)
    except Exception as e:  # noqa: BLE001
        st.warning(f"⚠ 시나리오 데이터 검증 실패: {e}")
        return
    try:
        vat_rate = load_labels(PROJECT_ROOT).quote.vat_rate
    except Exception:  # noqa: BLE001
        vat_rate = 0.10

    # 전체 일괄 할인율 입력 (문서 단위; 모든 시나리오에 공통 적용)
    st.divider()
    st.caption("📊 **이 시나리오의 합계** (모든 기간 합산, 할인 차감 반영)")
    td_col, _ = st.columns([1, 3])
    with td_col:
        td_pct = st.number_input(
            "전체 일괄 할인율 (%)",
            min_value=0, max_value=100, step=1,
            value=int((st.session_state.mc_doc.get("total_discount_rate") or 0) * 100),
            key=f"mc_total_discount_pct_{s_idx}",
            help="공급가액 합계에 일괄 적용. 항목별 할인은 별도 행에서 입력합니다.",
        )
    st.session_state.mc_doc["total_discount_rate"] = (td_pct / 100) if td_pct > 0 else None
    td_rate = td_pct / 100 if td_pct > 0 else 0
    items_sum = sum(scenario_vat_and_total(sc_obj, vat_rate, None)[0:1])  # 공급가액 (할인 전)
    items_sum = scenario_vat_and_total(sc_obj, vat_rate, None)[0]
    subtotal, vat, total = scenario_vat_and_total(sc_obj, vat_rate, td_rate)
    td_value = items_sum - subtotal

    def _money_label(v: float) -> str:
        val = int(round(v))
        return f"-₩{abs(val):,}" if val < 0 else f"₩{val:,}"

    if td_pct > 0:
        m0, m1, m2, m3 = st.columns(4)
        m0.metric("일괄 할인", f"-₩{int(round(td_value)):,}",
                  delta=f"{td_pct}%", delta_color="inverse")
        m1.metric("공급가액", _money_label(subtotal))
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", _money_label(vat))
        m3.metric("합계 금액", _money_label(total))
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("공급가액", _money_label(subtotal))
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", _money_label(vat))
        m3.metric("합계 금액", _money_label(total))


def _render_section_editor(s_idx: int, sec_idx: int, section: dict,
                           products: list[dict]) -> None:
    sec1, sec2 = st.columns([5, 1])
    with sec1:
        section["name"] = st.text_input(
            "구분 이름", value=section.get("name", ""),
            key=f"sec_name_{s_idx}_{sec_idx}",
            placeholder="예: 멤버십 클라우드",
        )
    with sec2:
        if st.button("❌ 구분 삭제", key=f"sec_del_{s_idx}_{sec_idx}",
                     use_container_width=True):
            st.session_state.mc_doc["scenarios"][s_idx]["sections"].pop(sec_idx)
            st.rerun()

    # 이 구분의 항목들을 평면 표로 (분류는 컬럼)
    flat_items = _section_to_flat_items(section)
    df = _mc_items_to_df(flat_items)

    # 품목 빠른 추가 (드롭다운) + 빈 행 + 할인 행
    section_name = section.get("name", "")
    matching = [p for p in products if p.get("section") == section_name]
    quick_col1, quick_col2, quick_col3, quick_col4 = st.columns([4, 1.3, 1.0, 1.3])
    with quick_col1:
        if matching:
            options = list(range(len(matching)))
            picked = st.selectbox(
                "품목 추가",
                options=options,
                index=None,
                format_func=lambda i: (
                    f"[{matching[i].get('subcategory', '-')}] {matching[i]['name']}"
                ),
                placeholder=f"'{section_name}' 품목 선택...",
                key=f"sec_pick_{s_idx}_{sec_idx}",
                label_visibility="collapsed",
            )
        else:
            st.caption(
                f"'{section_name}' 와 매칭되는 품목이 없습니다. "
                "오른쪽 '+ 빈 행' 으로 직접 추가할 수 있어요."
            )
            picked = None
    with quick_col2:
        if st.button("+ 품목 추가", key=f"sec_addrow_{s_idx}_{sec_idx}",
                     use_container_width=True,
                     help="좌측 드롭다운에서 항목을 선택한 뒤 누르세요"):
            if picked is None or not matching:
                st.warning("⚠ 좌측 드롭다운에서 품목을 먼저 선택해주세요.")
            else:
                p = matching[picked]
                new_row = {
                    "분류": p.get("subcategory", "기타"),
                    "상세 구분": p["name"],
                    "기간": p.get("billing_period", ""),
                    "단가": p.get("unit_price"),
                    "단가(텍스트)": p.get("unit_price_text", ""),
                    "할인율(%)": None,
                    "할인금액": None,
                    "금액 텍스트": p.get("default_amount_text", ""),
                    "비고": p.get("notes", ""),
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                section["categories"] = _df_to_section_categories(df)
                st.rerun()
    with quick_col3:
        if st.button("+ 빈 행", key=f"sec_addblank_{s_idx}_{sec_idx}",
                     use_container_width=True,
                     help="빈 행을 추가하고 직접 입력합니다"):
            _add_blank_row_to_section(section)
            st.rerun()
    with quick_col4:
        if st.button("💰 + 할인", key=f"sec_adddiscount_{s_idx}_{sec_idx}",
                     use_container_width=True,
                     help="'할인' 분류에 음수 단가 행이 추가됩니다. 협상가로 조정 가능."):
            _add_discount_row_to_section(section)
            st.rerun()

    # 표 편집
    edited = st.data_editor(
        df,
        column_config={
            "분류": st.column_config.SelectboxColumn(
                "분류", options=_DEFAULT_SUBCATEGORIES + ["기타"],
                help="초기구축비/사용료/옵션 등",
            ),
            "상세 구분": st.column_config.TextColumn("상세 구분", required=True),
            "기간": st.column_config.TextColumn("기간",
                help="1회성 / 매월 / 발생시 / 발생월 / 1개당"),
            "단가": st.column_config.NumberColumn(
                "단가 (숫자)", format="₩%,d",
                help="단가 (양수). 분류='할인'은 자동 차감됩니다.",
            ),
            "단가(텍스트)": st.column_config.TextColumn(
                "단가(텍스트)",
                help="숫자로 표현 불가한 단가 (예: '투입기간 X SW개발자 임금')",
            ),
            "할인율(%)": st.column_config.NumberColumn(
                "할인율(%)", min_value=0, max_value=100, step=1, format="%d%%",
                help="항목별 할인율 (0~100). 할인금액 입력 시 할인금액이 우선.",
            ),
            "할인금액": st.column_config.NumberColumn(
                "할인금액", min_value=0, step=1000, format="₩%,d",
                help="항목별 할인 금액 (양수).",
            ),
            "금액 텍스트": st.column_config.TextColumn(
                "금액 텍스트",
                help="비워두면 자동 계산. '후청구', '협의 금액' 등 텍스트도 입력 가능.",
            ),
            "비고": st.column_config.TextColumn("비고"),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"sec_editor_{s_idx}_{sec_idx}",
    )

    # 편집 결과를 section에 반영
    section["categories"] = _df_to_section_categories(edited)

    # 소계 미리보기
    try:
        sec_obj = MembershipSection.model_validate(section)
        by_period = section_subtotals_by_period(sec_obj)
        if by_period:
            badge_cols = st.columns(len(by_period) + 1)
            badge_cols[0].caption("**예상 합계**")
            for col, (period, amt) in zip(badge_cols[1:], by_period.items()):
                col.metric(period, f"₩{int(amt):,}")
    except Exception:
        pass


def _build_membership_artifacts(state: dict, soffice_available: bool,
                                issuer_contact: dict | None = None,
                                status_label: str = "멤버십 견적서 생성 중..."):
    """state를 검증·렌더링하여 (document_id, docx_bytes, pdf_bytes) 반환.
    실패 시 None 반환 (에러는 이미 화면에 표시됨)."""
    # 문서번호 비어있으면 자동 생성 (QR 견적서와 동일 패턴)
    if not (state.get("document_id") or "").strip():
        try:
            issued = date.fromisoformat(state.get("issued_date", date.today().isoformat()))
        except (TypeError, ValueError):
            issued = date.today()
        cp_name = ((state.get("counterparty") or {}).get("name") or "고객").strip()
        cp_short = "".join(c for c in cp_name if c.isalnum())[:8] or "고객"
        state["document_id"] = f"MC-{issued.strftime('%Y%m%d')}-{cp_short}"
    try:
        document = MembershipQuoteDocument.model_validate(state)
    except Exception as e:
        st.error(f"❌ 데이터 검증 실패: {e}")
        return None

    try:
        brand = load_brand(PROJECT_ROOT, document.brand_id)
    except FileNotFoundError as e:
        st.error(f"❌ 브랜드 로드 실패: {e}")
        return None

    # 발행 담당자 정보 — 폼에서 입력한 값으로 brand.contact_person 일회용 override
    if issuer_contact and (issuer_contact.get("name") or "").strip():
        brand = brand.model_copy(update={
            "contact_person": ContactPerson(
                name=issuer_contact["name"].strip(),
                title=issuer_contact.get("title"),
                phone=issuer_contact.get("phone"),
                email=issuer_contact.get("email"),
            )
        })

    output_dir = PROJECT_ROOT / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    docx_path = output_dir / f"{document.document_id}.docx"

    with st.status(status_label, expanded=True) as status:
        st.write("📝 DOCX 생성 중...")
        try:
            render_membership_docx(brand, document, PROJECT_ROOT, docx_path)
        except Exception as e:
            status.update(label="❌ DOCX 생성 실패", state="error")
            st.exception(e)
            return None
        st.write(f"   ✓ {docx_path.name}")

        pdf_bytes = None
        if soffice_available:
            st.write("📑 PDF 변환 중...")
            try:
                pdf_path = convert_docx_to_pdf(docx_path, output_dir)
                pdf_bytes = pdf_path.read_bytes()
                st.write(f"   ✓ {pdf_path.name}")
            except Exception as e:
                st.warning(f"PDF 변환 실패 (DOCX만 다운로드 가능): {e}")
        status.update(label="✅ 생성 완료", state="complete")

    docx_bytes = docx_path.read_bytes()
    # 최근 다운로드 견적서 히스토리에 저장 (같은 document_id 면 1개로 유지)
    _mc_save_history(_mc_snapshot_payload(), document.document_id)
    return document.document_id, docx_bytes, pdf_bytes


def _generate_membership_quote(state: dict, soffice_available: bool,
                               issuer_contact: dict | None = None):
    result = _build_membership_artifacts(
        state, soffice_available, issuer_contact,
        status_label="멤버십 견적서 생성 중...",
    )
    if result is None:
        return
    document_id, docx_bytes, pdf_bytes = result

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📝 DOCX 다운로드",
            data=docx_bytes,
            file_name=f"{document_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl2:
        if pdf_bytes:
            st.download_button(
                "📑 PDF 다운로드",
                data=pdf_bytes,
                file_name=f"{document_id}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.button("📑 PDF 사용 불가", disabled=True, use_container_width=True)


def _preview_membership_quote(state: dict, soffice_available: bool,
                              issuer_contact: dict | None = None):
    result = _build_membership_artifacts(
        state, soffice_available, issuer_contact,
        status_label="미리보기 생성 중...",
    )
    if result is None:
        return
    document_id, docx_bytes, pdf_bytes = result

    if not pdf_bytes:
        st.error("❌ PDF 변환기(LibreOffice)를 사용할 수 없어 미리보기를 표시할 수 없습니다.")
        return

    st.success(f"미리보기: **{document_id}.pdf**")

    rendered = False
    try:
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(pdf_bytes)
        try:
            n_pages = len(pdf)
            for i in range(n_pages):
                page = pdf[i]
                bitmap = page.render(scale=2)
                img = bitmap.to_pil()
                st.image(img, caption=f"페이지 {i + 1} / {n_pages}",
                         use_container_width=True)
            rendered = True
        finally:
            pdf.close()
    except Exception as e:
        st.warning(f"이미지 미리보기를 사용할 수 없습니다 ({e}). 아래 다운로드로 확인하세요.")

    # DOCX + PDF 다운로드 (택1)
    dl1, dl2 = st.columns(2)
    with dl1:
        if docx_bytes:
            st.download_button(
                "📝 DOCX 다운로드", data=docx_bytes,
                file_name=f"{document_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_docx_mc_preview_{document_id}",
            )
        else:
            st.button("📝 DOCX 사용 불가", disabled=True,
                      use_container_width=True,
                      key=f"dl_docx_mc_preview_disabled_{document_id}")
    with dl2:
        st.download_button(
            "📑 PDF 다운로드", data=pdf_bytes,
            file_name=f"{document_id}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"dl_pdf_mc_preview_{document_id}",
        )


# ═════════════════════════════════════════════════════════════
# 메인 라우터
# ═════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="견적서 자동 생성", page_icon="📋", layout="wide")

    QUOTE_PAGES = [
        "📋 QR오더 견적기",
        "🌳 [야외형] QR오더 견적기",
        "🏢 멤버십 견적기",
    ]
    SETTING_PAGES = [
        "📦 품목 관리",
        "⚙ 기본 설정",
    ]
    ALL_PAGES = QUOTE_PAGES + SETTING_PAGES

    # 첫 진입 시: URL ?page=... 쿼리 파라미터로 마지막 페이지 복원 (새로고침 보존)
    if "menu_quote" not in st.session_state and "menu_setting" not in st.session_state:
        saved = st.query_params.get("page") or QUOTE_PAGES[0]
        if saved in QUOTE_PAGES:
            st.session_state["menu_quote"] = saved
            st.session_state["menu_setting"] = None
        elif saved in SETTING_PAGES:
            st.session_state["menu_quote"] = None
            st.session_state["menu_setting"] = saved
        else:
            st.session_state["menu_quote"] = QUOTE_PAGES[0]
            st.session_state["menu_setting"] = None

    # 한 그룹 클릭 시 다른 그룹 자동 선택 해제 + URL 갱신
    def _on_quote_change():
        st.session_state["menu_setting"] = None
        try:
            st.query_params["page"] = st.session_state.get("menu_quote") or QUOTE_PAGES[0]
        except Exception:
            pass

    def _on_setting_change():
        st.session_state["menu_quote"] = None
        try:
            st.query_params["page"] = st.session_state.get("menu_setting") or SETTING_PAGES[0]
        except Exception:
            pass

    with st.sidebar:
        st.markdown("### 📑 견적기")
        st.radio(
            "견적기", QUOTE_PAGES, key="menu_quote",
            label_visibility="collapsed",
            on_change=_on_quote_change,
        )
        st.markdown("---")
        st.markdown("### ⚙ 설정")
        st.radio(
            "설정", SETTING_PAGES, key="menu_setting",
            label_visibility="collapsed",
            on_change=_on_setting_change,
        )

    page = (st.session_state.get("menu_quote")
            or st.session_state.get("menu_setting")
            or QUOTE_PAGES[0])
    # URL 쿼리에도 현재 페이지 동기화 (새로고침 시 복원용)
    try:
        if st.query_params.get("page") != page:
            st.query_params["page"] = page
    except Exception:
        pass

    # ── 페이지 전환 감지: 미해결 다이얼로그 + 임시 상태 정리 + 새로고침 ──
    prev_page = st.session_state.get("_last_active_page")
    if prev_page is not None and prev_page != page:
        # 1) 확인 다이얼로그/저장 마커 등 임시 플래그 모두 제거
        transient_keys = [
            # 전체 초기화 확인
            "_qr_reset_confirm", "_mc_reset_confirm", "_mc_items_reset_confirm",
            # 저장 완료 토스트 마커
            "_catalog_saved_at_qr", "_catalog_saved_at_outdoor",
            "_catalog_saved_at_membership",
            # GitHub sync 결과
            "_github_sync_qr", "_github_sync_outdoor", "_github_sync_membership",
            # 자동저장 load_once 플래그 — 다음 진입 시 디스크에서 재로드
            "_qr_autosave_loaded", "_mc_autosave_loaded",
        ]
        for k in transient_keys:
            st.session_state.pop(k, None)
        # 2) 메모리상 작성 데이터도 비움 (자동저장 파일은 디스크에 그대로 보존됨)
        for k in ("items_df", "mc_items_df", "mc_doc",
                  "catalog_pick"):
            st.session_state.pop(k, None)
        # 견적서 폼 위젯 값들도 비움 — 다음 진입 시 자동저장에서 복원
        for k in list(st.session_state.keys()):
            if k.startswith("qr_hist_") or k.startswith("qr_tpl_") \
                    or k.startswith("mc_hist_") or k.startswith("mc_tpl_"):
                st.session_state.pop(k, None)
        st.session_state["_last_active_page"] = page
        st.rerun()
    st.session_state["_last_active_page"] = page

    if page == "📋 QR오더 견적기":
        render_quote_page(catalog_kind="qr")
    elif page == "🌳 [야외형] QR오더 견적기":
        render_quote_page(catalog_kind="outdoor")
    elif page == "🏢 멤버십 견적기":
        render_membership_quote_page()
    elif page == "📦 품목 관리":
        render_catalog_page()
    elif page == "⚙ 기본 설정":
        render_settings_page()


if __name__ == "__main__":
    main()
