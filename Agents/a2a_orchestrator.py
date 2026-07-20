"""
A2A orchestrator — Agent_1 (master clean) → HANDOFF → Agent_2 (pricebook QA/QC).

Usage:
  python a2a_orchestrator.py --master "HK Master....xlsm" --pricebook "AD_July2026 (en_HK).xlsx"
  python a2a_orchestrator.py --master ... --pricebook ... --llm
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from uuid import uuid4

from dotenv import load_dotenv

from a2a_bus import MessageBus
from a2a_messages import (
    AGENT_1,
    AGENT_2,
    ORCHESTRATOR,
    MessageType,
    make_message,
)

# Project root = parent of Agents/ (Excel files + .env + prompt/ live here).
AGENTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = AGENTS_DIR.parent

# Agent capability cards (discovery metadata only — no HTTP).
AGENT_CARDS = {
    AGENT_1: {
        "name": AGENT_1,
        "description": "Clean master pricelist: merge sheets, drop discontinued, export active list.",
        "skills": [
            "load_price_list",
            "remove_discontinued_models",
            "list_updated_products",
            "list_new_products",
            "export_active_list",
        ],
        "inputs": ["master_excel_path"],
        "outputs": ["active_price_list.xlsx", "HANDOFF"],
    },
    AGENT_2: {
        "name": AGENT_2,
        "description": "Pricebook QA/QC: validate HardwareStandard vs master, revise ATTR/DHI.",
        "skills": [
            "validate_pricebook",
            "revise_pricebook_attributes",
        ],
        "inputs": ["pricebook_excel_path", "active_price_list.xlsx"],
        "outputs": ["corrected_pricebook", "QA_RESULT"],
    },
}


def _ensure_project_env() -> None:
    """Load .env from project root and prefer that as cwd for relative Excel paths."""
    if str(AGENTS_DIR) not in sys.path:
        sys.path.insert(0, str(AGENTS_DIR))
    env_path = PROJECT_ROOT / ".env"
    load_dotenv(env_path)
    load_dotenv()  # also allow cwd override
    # Prompt path in agent_2 is relative: prompt/prompt_v4.txt
    # Excel defaults also live at project root.
    os.chdir(PROJECT_ROOT)


def run_a2a_pipeline(
    master_path: str,
    pricebook_path: str,
    *,
    output_path: str = "active_price_list.xlsx",
    use_llm: bool = False,
    bus: MessageBus | None = None,
) -> dict:
    """
    Execute the full A2A pipeline and return a result dict for CLI / Streamlit.

    On Agent_1 failure or missing artifact, publishes ERROR and does not call Agent_2.
    """
    _ensure_project_env()

    # Import agents after env/cwd are ready (they read Azure creds at import).
    import agent_1
    import agent_2

    correlation_id = str(uuid4())
    bus = bus or MessageBus(audit=True, run_id=correlation_id[:8])

    master_path = str(Path(master_path).resolve())
    pricebook_path = str(Path(pricebook_path).resolve())
    out = Path(output_path)
    if not out.is_absolute():
        output_path = str((PROJECT_ROOT / out).resolve())
    else:
        output_path = str(out.resolve())

    result: dict = {
        "correlation_id": correlation_id,
        "status": "running",
        "handoff": None,
        "qa_result": None,
        "messages": [],
        "bus_log_path": bus.log_path,
        "agent_cards": AGENT_CARDS,
    }

    # --- TASK_START → Agent_1 ---
    bus.publish(
        make_message(
            MessageType.TASK_START,
            ORCHESTRATOR,
            AGENT_1,
            {
                "master_path": master_path,
                "output_path": output_path,
                "use_llm": use_llm,
            },
            correlation_id=correlation_id,
        )
    )
    bus.publish(
        make_message(
            MessageType.STATUS,
            ORCHESTRATOR,
            AGENT_1,
            {"message": "Starting Agent_1 master pricelist clean"},
            correlation_id=correlation_id,
        )
    )

    try:
        handoff = agent_1.run_agent_1(
            master_path=master_path,
            output_path=output_path,
            use_llm=use_llm,
        )
    except Exception as exc:  # noqa: BLE001
        bus.publish(
            make_message(
                MessageType.ERROR,
                AGENT_1,
                ORCHESTRATOR,
                {"error": str(exc), "stage": "agent_1"},
                correlation_id=correlation_id,
            )
        )
        result["status"] = "failed"
        result["error"] = str(exc)
        result["messages"] = [m.to_dict() for m in bus.history]
        return result

    result["handoff"] = handoff
    bus.publish(
        make_message(
            MessageType.HANDOFF,
            AGENT_1,
            AGENT_2,
            handoff,
            correlation_id=correlation_id,
        )
    )

    artifact = handoff.get("artifact_path")
    if handoff.get("status") != "ready_for_qa" or not artifact or not os.path.isfile(artifact):
        error = handoff.get("error") or f"Agent_1 handoff artifact missing: {artifact}"
        bus.publish(
            make_message(
                MessageType.ERROR,
                ORCHESTRATOR,
                AGENT_2,
                {"error": error, "stage": "handoff_validation"},
                correlation_id=correlation_id,
            )
        )
        result["status"] = "failed"
        result["error"] = error
        result["messages"] = [m.to_dict() for m in bus.history]
        return result

    # --- TASK_START → Agent_2 ---
    bus.publish(
        make_message(
            MessageType.TASK_START,
            ORCHESTRATOR,
            AGENT_2,
            {
                "pricebook_path": pricebook_path,
                "master_path": artifact,
                "use_llm": use_llm,
            },
            correlation_id=correlation_id,
        )
    )
    bus.publish(
        make_message(
            MessageType.STATUS,
            ORCHESTRATOR,
            AGENT_2,
            {"message": "Starting Agent_2 pricebook QA/QC with Agent_1 artifact"},
            correlation_id=correlation_id,
        )
    )

    try:
        qa_result = agent_2.run_agent_2(
            pricebook_path=pricebook_path,
            master_path=artifact,
            use_llm=use_llm,
        )
    except Exception as exc:  # noqa: BLE001
        bus.publish(
            make_message(
                MessageType.ERROR,
                AGENT_2,
                ORCHESTRATOR,
                {"error": str(exc), "stage": "agent_2"},
                correlation_id=correlation_id,
            )
        )
        result["status"] = "failed"
        result["error"] = str(exc)
        result["messages"] = [m.to_dict() for m in bus.history]
        return result

    result["qa_result"] = qa_result
    bus.publish(
        make_message(
            MessageType.QA_RESULT,
            AGENT_2,
            ORCHESTRATOR,
            qa_result,
            correlation_id=correlation_id,
        )
    )

    if qa_result.get("status") != "completed":
        result["status"] = "failed"
        result["error"] = qa_result.get("error") or "Agent_2 failed"
    else:
        result["status"] = "completed"

    result["messages"] = [m.to_dict() for m in bus.history]
    return result


def run_agent_1_only(
    master_path: str,
    output_path: str = "active_price_list.xlsx",
    *,
    use_llm: bool = False,
    bus: MessageBus | None = None,
) -> dict:
    """Run only Agent_1 and publish HANDOFF (no Agent_2)."""
    _ensure_project_env()
    import agent_1

    correlation_id = str(uuid4())
    bus = bus or MessageBus(audit=True, run_id=correlation_id[:8])
    master_path = str(Path(master_path).resolve())
    if not Path(output_path).is_absolute():
        output_path = str((PROJECT_ROOT / output_path).resolve())

    bus.publish(
        make_message(
            MessageType.TASK_START,
            ORCHESTRATOR,
            AGENT_1,
            {"master_path": master_path, "output_path": output_path, "use_llm": use_llm},
            correlation_id=correlation_id,
        )
    )
    try:
        handoff = agent_1.run_agent_1(master_path, output_path, use_llm=use_llm)
        bus.publish(
            make_message(
                MessageType.HANDOFF,
                AGENT_1,
                ORCHESTRATOR,
                handoff,
                correlation_id=correlation_id,
            )
        )
        status = "completed" if handoff.get("status") == "ready_for_qa" else "failed"
        return {
            "correlation_id": correlation_id,
            "status": status,
            "handoff": handoff,
            "qa_result": None,
            "messages": [m.to_dict() for m in bus.history],
            "bus_log_path": bus.log_path,
        }
    except Exception as exc:  # noqa: BLE001
        bus.publish(
            make_message(
                MessageType.ERROR,
                AGENT_1,
                ORCHESTRATOR,
                {"error": str(exc)},
                correlation_id=correlation_id,
            )
        )
        return {
            "correlation_id": correlation_id,
            "status": "failed",
            "error": str(exc),
            "handoff": None,
            "qa_result": None,
            "messages": [m.to_dict() for m in bus.history],
            "bus_log_path": bus.log_path,
        }


def run_agent_2_only(
    pricebook_path: str,
    master_path: str,
    *,
    use_llm: bool = False,
    bus: MessageBus | None = None,
) -> dict:
    """Run only Agent_2 against an existing master artifact."""
    _ensure_project_env()
    import agent_2

    correlation_id = str(uuid4())
    bus = bus or MessageBus(audit=True, run_id=correlation_id[:8])
    pricebook_path = str(Path(pricebook_path).resolve())
    master_path = str(Path(master_path).resolve())

    bus.publish(
        make_message(
            MessageType.TASK_START,
            ORCHESTRATOR,
            AGENT_2,
            {
                "pricebook_path": pricebook_path,
                "master_path": master_path,
                "use_llm": use_llm,
            },
            correlation_id=correlation_id,
        )
    )
    try:
        qa_result = agent_2.run_agent_2(
            pricebook_path=pricebook_path,
            master_path=master_path,
            use_llm=use_llm,
        )
        bus.publish(
            make_message(
                MessageType.QA_RESULT,
                AGENT_2,
                ORCHESTRATOR,
                qa_result,
                correlation_id=correlation_id,
            )
        )
        return {
            "correlation_id": correlation_id,
            "status": qa_result.get("status", "failed"),
            "handoff": None,
            "qa_result": qa_result,
            "messages": [m.to_dict() for m in bus.history],
            "bus_log_path": bus.log_path,
        }
    except Exception as exc:  # noqa: BLE001
        bus.publish(
            make_message(
                MessageType.ERROR,
                AGENT_2,
                ORCHESTRATOR,
                {"error": str(exc)},
                correlation_id=correlation_id,
            )
        )
        return {
            "correlation_id": correlation_id,
            "status": "failed",
            "error": str(exc),
            "handoff": None,
            "qa_result": None,
            "messages": [m.to_dict() for m in bus.history],
            "bus_log_path": bus.log_path,
        }


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Local A2A pricebook QA/QC orchestrator")
    parser.add_argument(
        "--master",
        default=str(PROJECT_ROOT / "HK Master Price List - July 2026 OS_CLEANED.xlsm"),
        help="Master pricelist Excel path (Agent_1 input)",
    )
    parser.add_argument(
        "--pricebook",
        default=str(PROJECT_ROOT / "AD_July2026 (en_HK).xlsx"),
        help="Pricebook Excel path (Agent_2 input)",
    )
    parser.add_argument(
        "--output",
        default=str(PROJECT_ROOT / "active_price_list.xlsx"),
        help="Active price list export path (Agent_1 → Agent_2 handoff artifact)",
    )
    parser.add_argument(
        "--llm",
        action="store_true",
        help="Use Azure OpenAI LangChain agents (default: direct tool mode)",
    )
    parser.add_argument(
        "--agent1-only",
        action="store_true",
        help="Run Agent_1 only",
    )
    parser.add_argument(
        "--agent2-only",
        action="store_true",
        help="Run Agent_2 only (requires existing --output master artifact)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    if args.agent1_only:
        result = run_agent_1_only(args.master, args.output, use_llm=args.llm)
    elif args.agent2_only:
        result = run_agent_2_only(args.pricebook, args.output, use_llm=args.llm)
    else:
        result = run_a2a_pipeline(
            args.master,
            args.pricebook,
            output_path=args.output,
            use_llm=args.llm,
        )

    print(f"\nA2A status: {result.get('status')}")
    print(f"Correlation: {result.get('correlation_id')}")
    if result.get("bus_log_path"):
        print(f"Audit log: {result['bus_log_path']}")
    if result.get("error"):
        print(f"Error: {result['error']}")
        return 1

    handoff = result.get("handoff") or {}
    if handoff:
        print(
            f"Agent_1 handoff: active={handoff.get('active_count')} "
            f"artifact={handoff.get('artifact_path')}"
        )
    qa = result.get("qa_result") or {}
    if qa:
        print(f"Agent_2 corrected: {qa.get('corrected_path')}")
        if qa.get("step1_summary"):
            print("\n--- Step 1 ---\n", qa["step1_summary"][:2000])
        if qa.get("step2_summary"):
            print("\n--- Step 2 ---\n", qa["step2_summary"][:2000])

    print("\nMessage timeline:")
    for msg in result.get("messages") or []:
        print(
            f"  [{msg.get('type')}] {msg.get('sender')} → {msg.get('recipient')}"
        )
    return 0 if result.get("status") == "completed" else 1


if __name__ == "__main__":
    # Allow `python a2a_orchestrator.py` from Agents/ without installing a package.
    if str(AGENTS_DIR) not in sys.path:
        sys.path.insert(0, str(AGENTS_DIR))
    raise SystemExit(main())
