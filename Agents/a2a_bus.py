"""
In-memory Agent-to-Agent message bus with optional JSONL audit log.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from threading import Lock
from typing import Callable

from a2a_messages import AgentMessage, MessageType


# Default audit directory next to this module.
_DEFAULT_LOG_DIR = Path(__file__).resolve().parent / "a2a_logs"


class MessageBus:
    """
    Process-local publish/subscribe bus.

    - Keeps full ordered history for Streamlit / CLI inspection.
    - Optionally appends each message as one JSON line under a2a_logs/.
    """

    def __init__(
        self,
        *,
        audit: bool = True,
        log_dir: str | Path | None = None,
        run_id: str | None = None,
    ) -> None:
        self._history: list[AgentMessage] = []
        self._lock = Lock()
        self._subscribers: list[Callable[[AgentMessage], None]] = []
        self._audit = audit
        self._log_dir = Path(log_dir) if log_dir else _DEFAULT_LOG_DIR
        self._run_id = run_id
        self._log_path: Path | None = None
        if self._audit:
            self._log_dir.mkdir(parents=True, exist_ok=True)
            name = f"a2a_{self._run_id or 'session'}.jsonl"
            self._log_path = self._log_dir / name

    @property
    def history(self) -> list[AgentMessage]:
        with self._lock:
            return list(self._history)

    @property
    def log_path(self) -> str | None:
        return str(self._log_path) if self._log_path else None

    def clear(self) -> None:
        with self._lock:
            self._history.clear()

    def subscribe(self, callback: Callable[[AgentMessage], None]) -> None:
        with self._lock:
            self._subscribers.append(callback)

    def publish(self, message: AgentMessage) -> AgentMessage:
        with self._lock:
            self._history.append(message)
            subscribers = list(self._subscribers)
            self._append_audit(message)

        for callback in subscribers:
            try:
                callback(message)
            except Exception as exc:  # noqa: BLE001 — UI callbacks must not kill the bus
                print(f"[MessageBus] subscriber error: {exc}")

        return message

    def latest(
        self,
        msg_type: MessageType | str | None = None,
        *,
        sender: str | None = None,
        recipient: str | None = None,
        correlation_id: str | None = None,
    ) -> AgentMessage | None:
        """Return the most recent message matching optional filters."""
        if isinstance(msg_type, str):
            msg_type = MessageType(msg_type)

        with self._lock:
            for message in reversed(self._history):
                if msg_type is not None and message.type != msg_type:
                    continue
                if sender is not None and message.sender != sender:
                    continue
                if recipient is not None and message.recipient != recipient:
                    continue
                if correlation_id is not None and message.correlation_id != correlation_id:
                    continue
                return message
        return None

    def filter(
        self,
        msg_type: MessageType | str | None = None,
        *,
        correlation_id: str | None = None,
    ) -> list[AgentMessage]:
        if isinstance(msg_type, str):
            msg_type = MessageType(msg_type)

        with self._lock:
            results = []
            for message in self._history:
                if msg_type is not None and message.type != msg_type:
                    continue
                if correlation_id is not None and message.correlation_id != correlation_id:
                    continue
                results.append(message)
            return results

    def _append_audit(self, message: AgentMessage) -> None:
        if not self._audit or self._log_path is None:
            return
        try:
            with open(self._log_path, "a", encoding="utf-8") as handle:
                handle.write(json.dumps(message.to_dict(), default=str) + os.linesep)
        except OSError as exc:
            print(f"[MessageBus] audit write failed: {exc}")
