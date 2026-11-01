from __future__ import annotations

import json
import threading
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from personal_budget_plaid_excel.domain.models import SyncItemState
from personal_budget_plaid_excel.infrastructure.keychain import MacOSKeychain
from personal_budget_plaid_excel.infrastructure.plaid_rest import PlaidRestClient
from personal_budget_plaid_excel.infrastructure.telemetry import Telemetry
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook


STAGES = {"link_opened", "institution_selected", "share_data_completed", "callback_received", "token_persisted", "workbook_saved", "callback_failed"}


class LinkLifecycle:
    def __init__(self, telemetry: Telemetry) -> None:
        self.telemetry = telemetry
        self.stage = "starting"
        self._lock = threading.Lock()

    def record(self, stage: str, **details: Any) -> None:
        if stage not in STAGES:
            raise ValueError(f"Unsupported Link stage: {stage}")
        with self._lock:
            self.stage = stage
        self.telemetry.event("plaid_link_stage", stage=stage, **details)

    def snapshot(self) -> dict[str, str]:
        with self._lock:
            return {"stage": self.stage}


def run_link_flow(workbook_path: Path, plaid: PlaidRestClient, *, port: int = 8765, connections: int = 1) -> dict[str, Any]:
    if connections < 1:
        raise ValueError("connections must be at least 1")
    link_token = plaid.create_link_token()
    completed = threading.Event()
    result: dict[str, Any] = {"connections": []}
    saved_count = 0
    saved_count_lock = threading.Lock()
    lifecycle = LinkLifecycle(Telemetry(workbook_path.parent / "logs" / "personal-budget-plaid-excel.jsonl"))

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802
            if self.path == "/connection-status":
                body = json.dumps(lifecycle.snapshot()).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            body = _link_html(link_token).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_POST(self) -> None:  # noqa: N802
            if self.path == "/status":
                payload = json.loads(self.rfile.read(int(self.headers.get("Content-Length", "0"))).decode("utf-8"))
                stage = payload.get("stage")
                if stage in {"link_opened", "institution_selected", "share_data_completed"}:
                    lifecycle.record(stage)
                    self.send_response(204)
                    self.end_headers()
                    return
                self.send_error(400)
                return
            if self.path != "/exchange":
                self.send_error(404)
                return
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            try:
                lifecycle.record("callback_received")
                exchange = plaid.exchange_public_token(payload["public_token"])
                item_id = exchange["item_id"]
                access_token_key = f"plaid:{item_id}"
                MacOSKeychain().set_secret(access_token_key, exchange["access_token"])
                lifecycle.record("token_persisted")
                workbook = BudgetWorkbook(workbook_path)
                workbook.ensure_structure()
                workbook.upsert_sync_state(
                    SyncItemState(
                        item_id=item_id,
                        institution_name=payload.get("institution", {}).get("name", ""),
                        cursor="",
                        access_token_key=access_token_key,
                        updated_at=datetime.now(timezone.utc),
                    )
                )
                workbook.save_with_backup(backup=True)
                lifecycle.record("workbook_saved")
                with saved_count_lock:
                    nonlocal saved_count
                    saved_count += 1
                    remaining = connections - saved_count
                    result["connections"].append({"item_id": item_id, "access_token_key": access_token_key})
                    if remaining <= 0:
                        completed.set()
            except Exception as exc:
                lifecycle.record("callback_failed", error_type=type(exc).__name__)
                self.send_error(500, "Local save failed; check the connection status on this page.")
                return
            body = json.dumps({"ok": True, "remaining": max(remaining, 0)}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, fmt: str, *args: Any) -> None:
            return

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        url = f"http://127.0.0.1:{port}/"
        webbrowser.open(url)
        if not completed.wait(timeout=600):
            raise TimeoutError("Timed out waiting for Plaid Link completion")
        result["connected"] = len(result["connections"])
        return result
    finally:
        server.shutdown()
        thread.join(timeout=5)


def _link_html(link_token: str) -> str:
    token_json = json.dumps(link_token)
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Connect Plaid</title>
  <script src="https://cdn.plaid.com/link/v2/stable/link-initialize.js"></script>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 32px; }}
    button {{ font-size: 16px; padding: 10px 14px; }}
    #status {{ margin-top: 16px; color: #374151; }}
  </style>
</head>
<body>
  <h1>Connect financial accounts</h1>
  <button id="connect">Open Plaid Link</button>
  <p id="status">Waiting to start.</p>
  <script>
    const status = document.getElementById("status");
    const stage = (name) => fetch("/status", {{method: "POST", headers: {{"Content-Type": "application/json"}}, body: JSON.stringify({{stage: name}})}}).catch(() => {{}});
    const labels = {{link_opened: "Plaid Link opened.", institution_selected: "Institution selected.", share_data_completed: "Share Data completed; waiting for local callback.", callback_received: "Local callback received.", token_persisted: "Connection saved securely.", workbook_saved: "Workbook connection record saved.", callback_failed: "Local save failed; keep this page open and review diagnostics."}};
    setInterval(async () => {{
      try {{ const response = await fetch("/connection-status"); const state = await response.json(); if (labels[state.stage]) status.textContent = labels[state.stage]; }} catch (_) {{}}
    }}, 750);
    stage("link_opened");
    const handler = Plaid.create({{
      token: {token_json},
      onSuccess: async (public_token, metadata) => {{
        stage("share_data_completed");
        status.textContent = "Saving connection locally...";
        const response = await fetch("/exchange", {{
          method: "POST",
          headers: {{ "Content-Type": "application/json" }},
          body: JSON.stringify({{ public_token, institution: metadata.institution }})
        }});
        if (!response.ok) throw new Error(await response.text());
        const saved = await response.json();
        status.textContent = saved.remaining ? `Connection saved. ${{saved.remaining}} more connection(s) needed; choose Open Plaid Link.` : "Connection saved. You can close this window.";
      }},
      onExit: (err) => {{
        status.textContent = err ? err.display_message || err.error_message : "Plaid Link closed.";
      }},
      onEvent: (eventName) => {{ if (eventName === "SELECT_INSTITUTION") stage("institution_selected"); }}
    }});
    document.getElementById("connect").addEventListener("click", () => handler.open());
    handler.open();
  </script>
</body>
</html>
"""
