"""
Streamlit UI for the local A2A pricebook QA/QC pipeline.

Run from the project root (this folder):
  py -3.12 -m streamlit run app.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parent
AGENTS_DIR = PROJECT_ROOT / "Agents"
UPLOAD_DIR = PROJECT_ROOT / "uploads"

if str(AGENTS_DIR) not in sys.path:
    sys.path.insert(0, str(AGENTS_DIR))

from a2a_bus import MessageBus
from a2a_orchestrator import (
    AGENT_CARDS,
    PROJECT_ROOT as ORCH_ROOT,
    run_a2a_pipeline,
    run_agent_1_only,
    run_agent_2_only,
)

DEFAULT_MASTER = str(PROJECT_ROOT / "HK Master Price List - July 2026 OS_CLEANED.xlsm")
DEFAULT_OUTPUT = str(PROJECT_ROOT / "active_price_list.xlsx")


def _init_state() -> None:
    if "last_result" not in st.session_state:
        st.session_state.last_result = None
    if "message_log" not in st.session_state:
        st.session_state.message_log = []
    if "uploaded_pricebook_path" not in st.session_state:
        st.session_state.uploaded_pricebook_path = None


def _save_uploaded_pricebook(uploaded_file) -> str:
    """Persist an uploaded pricebook under project uploads/ and return its path."""
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    dest = UPLOAD_DIR / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return str(dest.resolve())


def _file_download_button(label: str, path: str | None, key: str) -> None:
    if not path:
        return
    p = Path(path)
    if not p.is_file():
        st.caption(f"Not found: {path}")
        return
    data = p.read_bytes()
    st.download_button(
        label=label,
        data=data,
        file_name=p.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )


def _render_messages(messages: list[dict]) -> None:
    if not messages:
        st.info("No A2A messages yet. Run a pipeline to populate the timeline.")
        return
    for msg in messages:
        msg_type = msg.get("type", "?")
        header = f"**{msg_type}** · {msg.get('sender')} → {msg.get('recipient')}"
        with st.expander(header, expanded=msg_type in ("HANDOFF", "QA_RESULT", "ERROR")):
            st.caption(msg.get("timestamp", ""))
            st.json(msg.get("payload") or {})


def main() -> None:
    st.set_page_config(
        page_title="Pricebook A2A QA/QC",
        page_icon="📋",
        layout="wide",
    )
    _init_state()

    st.title("Pricebook A2A QA/QC")
    st.caption(
        "Local Agent-to-Agent pipeline: Agent_1 cleans the master list, "
        "then hands off to Agent_2 for pricebook validation and attribute revise."
    )

    with st.sidebar:
        st.header("Agent cards")
        for name, card in AGENT_CARDS.items():
            with st.expander(name, expanded=False):
                st.write(card.get("description", ""))
                st.markdown("**Skills**")
                st.write(", ".join(card.get("skills") or []))
                st.markdown("**Inputs**")
                st.write(", ".join(card.get("inputs") or []))
                st.markdown("**Outputs**")
                st.write(", ".join(card.get("outputs") or []))
        st.divider()
        st.caption(f"Project root: `{ORCH_ROOT}`")

    col_a, col_b = st.columns(2)
    with col_a:
        master_path = st.text_input("Master pricelist (Agent_1)", value=DEFAULT_MASTER)
        output_path = st.text_input("Active list output (handoff artifact)", value=DEFAULT_OUTPUT)
    with col_b:
        uploaded_pricebook = st.file_uploader(
            "Pricebook (Agent_2)",
            type=["xlsx", "xlsm", "xls"],
            help="Upload the pricebook Excel file for Agent_2 QA/QC.",
        )
        if uploaded_pricebook is not None:
            pricebook_path = _save_uploaded_pricebook(uploaded_pricebook)
            st.session_state.uploaded_pricebook_path = pricebook_path
            st.caption(f"Uploaded: `{pricebook_path}`")
        else:
            pricebook_path = st.session_state.uploaded_pricebook_path
        use_llm = st.checkbox("Use LLM agents (Azure OpenAI)", value=True)

    btn1, btn2, btn3 = st.columns(3)
    run_full = btn1.button("Run full A2A pipeline", type="primary", use_container_width=True)
    run_a1 = btn2.button("Run Agent_1 only", use_container_width=True)
    run_a2 = btn3.button("Run Agent_2 only", use_container_width=True)

    if run_full or run_a1 or run_a2:
        needs_pricebook = run_full or run_a2
        if needs_pricebook and not pricebook_path:
            st.error("Upload a pricebook Excel file before running Agent_2 or the full pipeline.")
        else:
            bus = MessageBus(audit=True)
            progress = st.empty()
            try:
                if run_full:
                    progress.info("Running Agent_1 → HANDOFF → Agent_2…")
                    result = run_a2a_pipeline(
                        master_path,
                        pricebook_path,
                        output_path=output_path,
                        use_llm=use_llm,
                        bus=bus,
                    )
                elif run_a1:
                    progress.info("Running Agent_1 only…")
                    result = run_agent_1_only(
                        master_path,
                        output_path,
                        use_llm=use_llm,
                        bus=bus,
                    )
                else:
                    progress.info("Running Agent_2 only…")
                    result = run_agent_2_only(
                        pricebook_path,
                        output_path,
                        use_llm=use_llm,
                        bus=bus,
                    )
                st.session_state.last_result = result
                st.session_state.message_log = result.get("messages") or []
                if result.get("status") == "completed":
                    progress.success(f"Completed · correlation `{result.get('correlation_id')}`")
                else:
                    progress.error(
                        f"Failed · {result.get('error') or result.get('status')} "
                        f"· `{result.get('correlation_id')}`"
                    )
            except Exception as exc:  # noqa: BLE001
                progress.error(f"Pipeline exception: {exc}")
                st.exception(exc)

    result = st.session_state.last_result
    if result:
        st.subheader("Results")
        status = result.get("status")
        st.metric("Pipeline status", status or "unknown")

        handoff = result.get("handoff") or {}
        qa = result.get("qa_result") or {}

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Active products", handoff.get("active_count", "—"))
        c2.metric("Discontinued removed", handoff.get("discontinued_removed", "—"))
        c3.metric("New products", handoff.get("new_count", "—"))
        c4.metric("Updated fields", handoff.get("updated_count", "—"))

        dl1, dl2 = st.columns(2)
        with dl1:
            st.markdown("**Agent_1 artifact**")
            _file_download_button(
                "Download active_price_list.xlsx",
                handoff.get("artifact_path") or output_path,
                key="dl_active",
            )
        with dl2:
            st.markdown("**Agent_2 corrected pricebook**")
            _file_download_button(
                "Download corrected pricebook",
                qa.get("corrected_path"),
                key="dl_corrected",
            )

        if handoff.get("summary"):
            with st.expander("Agent_1 summary", expanded=True):
                st.text(handoff["summary"])

        step1 = (qa.get("step1_summary") or "").strip()
        step2 = (qa.get("step2_summary") or "").strip()
        if step1 or step2:
            # LLM mode returns one combined reply in both fields — show it once.
            if step1 and step2 and step1 == step2:
                with st.expander("Agent_2 summary", expanded=True):
                    st.text(step1)
            elif step1 and step2:
                with st.expander("Agent_2 summary", expanded=True):
                    st.markdown("**Step 1 — validate_pricebook**")
                    st.text(step1)
                    st.markdown("**Step 2 — revise_pricebook_attributes**")
                    st.text(step2)
            else:
                with st.expander("Agent_2 summary", expanded=True):
                    st.text(step1 or step2)

        if result.get("bus_log_path"):
            st.caption(f"Audit log: `{result['bus_log_path']}`")

    st.subheader("A2A message timeline")
    _render_messages(st.session_state.message_log)

    if st.session_state.message_log:
        st.download_button(
            "Download message log (JSON)",
            data=json.dumps(st.session_state.message_log, indent=2, default=str),
            file_name="a2a_messages.json",
            mime="application/json",
            key="dl_messages",
        )


if __name__ == "__main__":
    main()
