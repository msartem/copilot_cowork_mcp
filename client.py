"""
client.py — Cowork agent HTTP/SSE client.

Handles authentication, routing, and the subscribe/messages protocol
for communicating with the Microsoft Cowork (Copilot Studio) agent.

Architecture:
  - First message: POST /v1/subscribe (sends message + opens SSE stream)
  - The SSE stream stays open in a background thread for the whole session
  - Follow-up messages: POST /v1/messages (202 ack, response on existing SSE)
  - If SSE disconnects, auto-reconnect via GET /v1/subscribe?conversationId=...
  - The `fr` (finish reason) event signals end of ONE response, not end of stream
"""

import json
import os
import uuid
import base64
import urllib.parse
import threading
import queue
import requests

SHOW_THINKING = os.environ.get("COWORK_SHOW_THINKING", "").lower() in ("1", "true", "yes")


ROUTING_HOST = "cowork.us-ia888.gateway.prod.island.powerapps.com"
FALLBACK_RUNTIME = "mcsaetherruntime.cus-ia302.gateway.prod.island.powerapps.com"


def _decode_jwt(token: str) -> dict:
    parts = token.split(".")
    payload = parts[1] + "=" * (4 - len(parts[1]) % 4)
    return json.loads(base64.urlsafe_b64decode(payload))


def _headers(token: str, conversation_id: str | None = None) -> dict:
    h = {
        "Authorization": f"Bearer {token}",
        "x-ms-weave-auth": f"Bearer {token}",
        "Content-Type": "application/json",
        "Origin": "https://m365.cloud.microsoft",
    }
    if conversation_id:
        h["x-conversation-id"] = conversation_id
    return h


def discover_runtime(token: str) -> str:
    """GET /v1/routing → regional runtime hostname."""
    url = f"https://{ROUTING_HOST}/v1/routing"
    hdrs = _headers(token)
    hdrs["x-ms-user-pdl"] = "NAM"
    try:
        resp = requests.get(url, headers=hdrs, timeout=10)
        resp.raise_for_status()
        endpoint = resp.json().get("endpoint", "")
        host = endpoint.replace("https://", "").replace("http://", "").rstrip("/")
        return host or FALLBACK_RUNTIME
    except Exception:
        return FALLBACK_RUNTIME


class CoworkSession:
    """Stateful conversation session with the Cowork agent.

    Keeps a single SSE stream open for the lifetime of the conversation.
    All responses arrive on that stream; messages are sent via separate
    HTTP POSTs.
    """

    def __init__(self, token: str, runtime_host: str | None = None):
        self.token = token.strip()
        if self.token.startswith("Bearer "):
            self.token = self.token[7:]

        claims = _decode_jwt(self.token)
        self.tenant_id = claims["tid"]
        self.user_oid = claims["oid"]
        self.user_name = claims.get("name", "Unknown")

        self.runtime = runtime_host or discover_runtime(self.token)
        self.conversation_id = f"{self.tenant_id}:{self.user_oid}:{uuid.uuid4()}"
        self.session_id: str | None = None
        self.last_event_id: str | None = None
        self._turn = 0

        # Background SSE stream state
        self._response_queue: queue.Queue[str | None] = queue.Queue()
        self._stream_thread: threading.Thread | None = None
        self._stream_alive = threading.Event()

        # Pending action approval (set by `ta` event, consumed by approve_action)
        self.pending_approval: dict | None = None

    def send(self, message: str, file_paths: list[str] | None = None) -> str:
        """Send a message and return the full text response.

        Args:
            message: Text message to send.
            file_paths: Optional list of local file paths to upload to the
                        Cowork workspace before sending the message.
        """
        if file_paths:
            for path in file_paths:
                try:
                    self.upload_file(path)
                except Exception:
                    # Failed upload may taint the conversation — start fresh
                    self.reset()
                    raise

        self._turn += 1

        if self._turn == 1:
            return self._subscribe_with_message(message)
        else:
            return self._followup(message)

    def reset(self):
        """Start a new conversation (closes existing SSE stream)."""
        self._close_stream()
        self.conversation_id = f"{self.tenant_id}:{self.user_oid}:{uuid.uuid4()}"
        self.session_id = None
        self.last_event_id = None
        self._turn = 0
        self.pending_approval = None

    def approve_action(self) -> str:
        """Approve the pending action and return the result.

        Call this after send() returns action details (pending_approval is set).
        Posts approval to the runtime, then waits for the result on the
        existing SSE stream.
        """
        if not self.pending_approval:
            return "No pending action to approve."

        approval = self.pending_approval
        self.pending_approval = None

        try:
            url = f"https://{self.runtime}/v1/tool-approval"
            resp = requests.post(
                url,
                headers=_headers(self.token, self.conversation_id),
                json=approval,
                timeout=30,
            )
            if resp.status_code != 200:
                return f"Approval failed ({resp.status_code}): {resp.text[:200]}"
        except Exception as e:
            return f"Approval error: {e}"

        # Wait for the result on the existing SSE stream
        return self._wait_for_response()

    def upload_file(self, file_path: str) -> dict:
        """Upload a local file to the Cowork workspace input directory.

        The file becomes available at /mnt/workspace/input/<filename> inside
        the Cowork container, where the agent can read and attach it.

        Args:
            file_path: Absolute path to the file on disk.

        Returns:
            API response dict with file_id, workspace_path, etc.

        Raises:
            ValueError: If the file exceeds the size limit.
            requests.HTTPError: If the upload fails.
        """
        import mimetypes as mt

        MAX_FILE_SIZE = 1024 * 1024  # 1 MB

        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE:
            raise ValueError(
                f"File too large ({file_size / 1024:.0f} KB). "
                f"Maximum upload size is {MAX_FILE_SIZE // 1024} KB. "
                f"Convert to JPEG and resize to stay under the limit."
            )

        url = f"https://{self.runtime}/v1/conversations/{self.conversation_id}/files"
        hdrs = {
            "Authorization": f"Bearer {self.token}",
            "x-ms-weave-auth": f"Bearer {self.token}",
            "Origin": "https://m365.cloud.microsoft",
            "x-conversation-id": self.conversation_id,
        }
        mime_type, _ = mt.guess_type(file_path)
        filename = os.path.basename(file_path)
        with open(file_path, "rb") as f:
            files = {"file": (filename, f, mime_type or "application/octet-stream")}
            resp = requests.post(url, headers=hdrs, files=files, timeout=60)
        resp.raise_for_status()
        return resp.json()

    # ── First message: POST /v1/subscribe ────────────────────────────────

    def _subscribe_with_message(self, text: str) -> str:
        url = f"https://{self.runtime}/v1/subscribe"
        hdrs = _headers(self.token, self.conversation_id)
        hdrs["Accept"] = "text/event-stream"
        hdrs["Cache-Control"] = "no-cache"

        resp = requests.post(
            url, headers=hdrs, json=self._msg_body(text),
            stream=True, timeout=180,
        )
        if resp.status_code != 200:
            return f"Error {resp.status_code}: {resp.text[:500]}"

        # Start background SSE reader thread — it stays alive for the session
        self._stream_thread = threading.Thread(
            target=self._sse_reader_loop, args=(resp,), daemon=True,
        )
        self._stream_thread.start()
        self._stream_alive.set()

        # Wait for the first response to complete
        return self._wait_for_response()

    # ── Follow-up messages ───────────────────────────────────────────────

    def _followup(self, text: str) -> str:
        # If the SSE stream died, reconnect first
        if not self._stream_alive.is_set():
            self._reconnect_sse()

        # Send message via POST /v1/messages
        url = f"https://{self.runtime}/v1/messages"
        resp = requests.post(
            url, headers=_headers(self.token, self.conversation_id),
            json=self._msg_body(text), timeout=30,
        )
        if resp.status_code != 202:
            return f"Error sending message: {resp.status_code}: {resp.text[:500]}"

        # Wait for response on the existing SSE stream
        return self._wait_for_response()

    # ── SSE reconnect ────────────────────────────────────────────────────

    def _reconnect_sse(self):
        qs = urllib.parse.urlencode({
            "conversationId": self.conversation_id,
            "_nonce": f"{int(uuid.uuid1().time)}-{uuid.uuid4().hex[:8]}",
        })
        url = f"https://{self.runtime}/v1/subscribe?{qs}"
        hdrs = _headers(self.token, self.conversation_id)
        hdrs["Accept"] = "text/event-stream"
        hdrs["Cache-Control"] = "no-cache"
        if self.last_event_id:
            hdrs["Last-Event-ID"] = self.last_event_id

        resp = requests.get(url, headers=hdrs, stream=True, timeout=180)
        if resp.status_code == 200:
            self._stream_thread = threading.Thread(
                target=self._sse_reader_loop, args=(resp,), daemon=True,
            )
            self._stream_thread.start()
            self._stream_alive.set()

    # ── Background SSE reader ────────────────────────────────────────────

    def _sse_reader_loop(self, resp):
        """Continuously read SSE events. Pushes completed responses to queue."""
        buffer = ""
        chunks: list[str] = []

        try:
            for line in resp.iter_lines(decode_unicode=True):
                if line is None:
                    continue
                if line == "":
                    if buffer.strip():
                        finished = self._handle_event(buffer, chunks)
                        if finished:
                            # fr event = end of one response, deliver it
                            self._response_queue.put("".join(chunks))
                            chunks = []
                    buffer = ""
                else:
                    buffer += line + "\n"
        except Exception:
            pass
        finally:
            # Stream died — deliver any partial response
            if chunks:
                self._response_queue.put("".join(chunks))
            self._stream_alive.clear()

    def _handle_event(self, raw: str, chunks: list[str]) -> bool:
        """Parse one SSE event. Returns True when the response is complete.

        Handles both query responses (dx → fr) and action flows (th → ta).
        """
        event_type = None
        event_data = ""
        event_id = None

        for line in raw.strip().split("\n"):
            if line.startswith("event:"):
                event_type = line[6:].strip()
            elif line.startswith("data:"):
                event_data = line[5:].strip()
            elif line.startswith("id:"):
                event_id = line[3:].strip()

        if event_id:
            self.last_event_id = event_id
        if not event_data:
            return False

        try:
            data = json.loads(event_data)
        except json.JSONDecodeError:
            return False

        if event_type == "session":
            self.session_id = data.get("sid")

        elif event_type == "dx":
            # Text chunk (query responses)
            chunks.append(data.get("t", ""))

        elif event_type == "th":
            # Thinking chunk (action reasoning) — only include if configured
            if SHOW_THINKING:
                chunks.append(data.get("c", ""))

        elif event_type == "ta":
            # Tool approval — store for user confirmation via approve_action()
            tool_full = data.get("tn", "")
            params = data.get("params", {})
            approval_id = data.get("aid", "")

            # mcp__m365_teams__PostMessage → server=m365_teams, tool=PostMessage
            parts = tool_full.split("__")
            server_name = parts[1] if len(parts) >= 3 else ""
            tool_name = parts[-1] if "__" in tool_full else tool_full

            self.pending_approval = {
                "always_allow": False,
                "approval_id": approval_id,
                "approved": True,
                "conversation_id": self.conversation_id,
                "edited_input": params,
                "scope": None,
                "server_name": server_name,
                "session_id": self.conversation_id,
                "tool_name": tool_name,
            }

            # Return action details so the caller can present them to the user
            chunks.clear()
            action_lines = [f"Action: {tool_name}"]
            for k, v in params.items():
                action_lines.append(f"  {k}: {v}")
            chunks.append("\n".join(action_lines))
            return True

        elif event_type == "ts":
            # Tool start — log which tool is being invoked
            tool = data.get("tn", "")
            if tool and "ToolSearch" not in tool:
                short = tool.rsplit("__", 1)[-1] if "__" in tool else tool
                chunks.append(f"\n[Calling: {short}]\n")

        elif event_type == "tx":
            # Tool execution result
            if not data.get("ok", True):
                chunks.append(f"\n[Tool error: {data.get('tn', 'unknown')}]\n")

        elif event_type == "error":
            # Backend error — surface the message and finish
            err_msg = data.get("err", "Unknown error")
            code = data.get("code", "")
            chunks.clear()
            chunks.append(f"Cowork error: {err_msg}" + (f" [{code}]" if code else ""))
            return True

        elif event_type == "rl":
            # Run-level end — if status is fail, finish the response
            if data.get("st") == "fail":
                if not chunks:
                    err_msg = data.get("err", "Request failed")
                    chunks.append(f"Cowork error: {err_msg}")
                return True

        elif event_type == "fr":
            return True  # signals end of response

        return False

    # ── Wait for response ────────────────────────────────────────────────

    def _wait_for_response(self, timeout: float = 180) -> str:
        try:
            result = self._response_queue.get(timeout=timeout)
            return result if result else "(empty response)"
        except queue.Empty:
            # Session is likely dead — reset so next call starts fresh
            self.reset()
            return "(timeout waiting for Cowork response — session reset, please retry)"

    # ── Cleanup ──────────────────────────────────────────────────────────

    def _close_stream(self):
        self._stream_alive.clear()
        # Drain any pending items
        while not self._response_queue.empty():
            try:
                self._response_queue.get_nowait()
            except queue.Empty:
                break

    def _msg_body(self, text: str) -> dict:
        return {
            "content": [{"text": text, "type": "text"}],
            "conversationId": self.conversation_id,
            "messageId": str(uuid.uuid4()),
            "role": "user",
        }
