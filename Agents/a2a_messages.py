"""
A2A message schema for local Agent-to-Agent communication.

Message types:
  TASK_START  — orchestrator assigns work to an agent
  HANDOFF     — Agent_1 completes master list; ready for Agent_2 QA/QC
  QA_RESULT   — Agent_2 finishes pricebook validation / attribute revise
  ERROR       — failure that should stop the pipeline
  STATUS      — progress / informational update
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from enum import Enum
from typing import Any
from uuid import uuid4


class MessageType(str, Enum):
    TASK_START = "TASK_START"
    HANDOFF = "HANDOFF"
    QA_RESULT = "QA_RESULT"
    ERROR = "ERROR"
    STATUS = "STATUS"


AGENT_1 = "Agent_1"
AGENT_2 = "Agent_2"
ORCHESTRATOR = "Orchestrator"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


@dataclass
class AgentMessage:
    """Structured envelope for local A2A messages."""

    type: MessageType
    sender: str
    recipient: str
    payload: dict[str, Any] = field(default_factory=dict)
    correlation_id: str = field(default_factory=lambda: str(uuid4()))
    id: str = field(default_factory=lambda: str(uuid4()))
    timestamp: str = field(default_factory=_utc_now_iso)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["type"] = self.type.value if isinstance(self.type, MessageType) else str(self.type)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> AgentMessage:
        msg_type = data["type"]
        if isinstance(msg_type, str):
            msg_type = MessageType(msg_type)
        return cls(
            type=msg_type,
            sender=data["sender"],
            recipient=data["recipient"],
            payload=dict(data.get("payload") or {}),
            correlation_id=data.get("correlation_id") or str(uuid4()),
            id=data.get("id") or str(uuid4()),
            timestamp=data.get("timestamp") or _utc_now_iso(),
        )

    def summary(self) -> str:
        """One-line summary suitable for UI timelines."""
        keys = list(self.payload.keys())[:5]
        preview = ", ".join(keys) if keys else "(empty)"
        return (
            f"{self.sender} -> {self.recipient} | {self.type.value} | "
            f"payload keys: {preview}"
        )


def make_message(
    msg_type: MessageType | str,
    sender: str,
    recipient: str,
    payload: dict[str, Any] | None = None,
    *,
    correlation_id: str | None = None,
) -> AgentMessage:
    """Factory for AgentMessage with optional correlation id reuse."""
    if isinstance(msg_type, str):
        msg_type = MessageType(msg_type)
    kwargs: dict[str, Any] = {
        "type": msg_type,
        "sender": sender,
        "recipient": recipient,
        "payload": payload or {},
    }
    if correlation_id:
        kwargs["correlation_id"] = correlation_id
    return AgentMessage(**kwargs)
