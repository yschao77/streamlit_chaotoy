"""Call sku_submit Apps Script Web App (submitPendingBatch / getFormMeta)."""

from __future__ import annotations

import json
import urllib.error
import urllib.request
from typing import Any

import streamlit as st

_DEFAULT_TIMEOUT = 60


def _secret(name: str, default: str = "") -> str:
    """Read top-level Streamlit secret. Prefer bracket access (same as textkey)."""
    try:
        val = st.secrets[name]
    except Exception:
        return default
    if val is None:
        return default
    return str(val).strip()


def sku_submit_webapp_url() -> str:
    return _secret("SKU_SUBMIT_WEBAPP_URL")


def sku_submit_batch_token() -> str:
    return _secret("SKU_SUBMIT_BATCH_TOKEN")


def sku_submit_configured() -> bool:
    return bool(sku_submit_webapp_url())


def call_sku_submit(action: str, **payload: Any) -> dict:
    """POST JSON to Web App doPost. Never log token or full secrets."""
    url = sku_submit_webapp_url()
    if not url:
        return {
            "ok": False,
            "message": "未設定 SKU_SUBMIT_WEBAPP_URL（.streamlit/secrets.toml 頂層，勿放在 [textkey] 內）",
        }
    body: dict[str, Any] = {"action": action}
    body.update(payload)
    token = sku_submit_batch_token()
    if token:
        body["token"] = token
    raw = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=raw,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    text = ""
    status = 0
    try:
        with urllib.request.urlopen(req, timeout=_DEFAULT_TIMEOUT) as res:
            text = res.read().decode("utf-8", errors="replace").strip()
            status = getattr(res, "status", 200)
    except urllib.error.HTTPError as e:
        try:
            text = e.read().decode("utf-8", errors="replace").strip()
        except Exception:
            text = ""
        status = e.code
    except urllib.error.URLError:
        return {"ok": False, "message": "無法連線 SKU 送審 Web App（網路或網址錯誤）"}

    if not text:
        return {
            "ok": False,
            "message": (
                f"Web App 空回應（HTTP {status}）。"
                "若部署為「僅自己」，外部無法呼叫，請改為「任何人」並設 BATCH_API_TOKEN。"
            ),
        }
    try:
        data = json.loads(text)
    except ValueError:
        looks_html = text.lstrip().lower().startswith("<!DOCTYPE html") or "<html" in text[:200].lower()
        if status in (401, 403) or looks_html:
            return {
                "ok": False,
                "message": (
                    f"Web App 要 Google 登入（HTTP {status}）。"
                    "部署對象須為「任何人」（ANYONE_ANONYMOUS），"
                    "並在 Script Properties 設 BATCH_API_TOKEN 與 secrets 的 SKU_SUBMIT_BATCH_TOKEN 相同。"
                ),
            }
        snippet = text[:120].replace("\n", " ")
        return {
            "ok": False,
            "message": f"Web App 回傳非 JSON（HTTP {status}）：{snippet}…",
        }
    if not isinstance(data, dict):
        return {"ok": False, "message": "Web App 回傳格式錯誤"}
    return data


def fetch_sku_form_meta() -> dict:
    cached = st.session_state.get("sku_submit_form_meta")
    if isinstance(cached, dict) and cached.get("ok"):
        return cached
    data = call_sku_submit("getFormMeta")
    if data.get("ok") or data.get("categories"):
        data = dict(data)
        data["ok"] = True
        st.session_state["sku_submit_form_meta"] = data
    return data


def submit_pending_batch(rows: list[dict]) -> dict:
    return call_sku_submit("submitPendingBatch", rows=rows)


def suggest_from_name(name: str) -> dict:
    return call_sku_submit("suggestFromName", name=name)


def category_code_options(meta: dict | None, level: str) -> list[str]:
    """level: l2 | l3 | l4"""
    cats = (meta or {}).get("categories") or {}
    items = cats.get(level) or []
    codes: list[str] = []
    for it in items:
        if isinstance(it, dict) and it.get("code"):
            codes.append(str(it["code"]))
        elif isinstance(it, str) and it.strip():
            codes.append(it.strip())
    if level == "l4" and not codes:
        return ["NEW", "SEC"]
    return codes
