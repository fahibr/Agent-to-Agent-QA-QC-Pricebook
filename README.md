# Pricebook QA/QC — Agent-to-Agent (A2A) Pipeline

Automates master pricelist cleaning and pricebook QA/QC using two cooperating agents and a local A2A message bus. A Streamlit UI orchestrates the full run and shows handoff messages, summaries, and downloadable Excel outputs.

**Requires:** Python **3.12.13**

---

## Workflow overview

```text
┌─────────────────┐     HANDOFF          ┌─────────────────┐
│    Agent_1      │  active_price_list   │    Agent_2      │
│ Master clean    │ ───────────────────► │ Pricebook QA/QC │
│ + export        │   .xlsx + counts     │ Validate + ATTR │
└─────────────────┘                      └─────────────────┘
        ▲                                         │
        │              Message bus                ▼
        └──────── Orchestrator / Streamlit ───────┘
```

| Step | Agent | What it does | Main output |
|------|--------|----------------|-------------|
| 1 | **Agent_1** | Loads the cleaned master pricelist, merges sheets, flags new/updated/discontinued, removes discontinued products | `active_price_list.xlsx` + **HANDOFF** message |
| 2 | **Agent_2** | Validates pricebook `HardwareStandard` against the active master; corrects mismatches (yellow); revises HardwareType / ATTYPE / DHI | `… - corrected.xlsx` + **QA_RESULT** message |

If Agent_1 fails or the handoff artifact is missing, Agent_2 is **not** started.

---

## What each agent does

### Agent_1 — Master pricelist

**Input:** Master Excel (e.g. `HK Master Price List - July 2026 OS_CLEANED.xlsm`)

1. Merge all brand sheets into one table  
2. Detect status from cell colors (discontinued / updated / new)  
3. Remove discontinued products  
4. Export active products to `active_price_list.xlsx`

**Tools:** `load_price_list`, `remove_discontinued_models`, `list_updated_products`, `list_new_products`, `find_model`, `export_active_list`

### Agent_2 — Pricebook QA/QC

**Inputs:** Uploaded pricebook Excel + Agent_1’s `active_price_list.xlsx`

1. **`validate_pricebook`** — Match Brand + Model + Finish to master; correct Product Category Code/Category, Description, Unit, Price; yellow-highlight corrected cells; leave “not in master” unchanged  
2. **`revise_pricebook_attributes`** — Detect locale from filename; fill HardwareType blanks with `NULL`; sync ATTYPE1–4 and DHI from HardwareType by locale + category  

Original pricebook files are never overwritten. Corrections go to `* - corrected.xlsx`.

---

## Project layout

```text
Combine_AI_Agents/
├── README.md                 ← this file
├── app.py                    ← Streamlit UI
├── requirements.txt
├── .env                      ← Azure OpenAI credentials (not committed)
├── prompt/
│   └── prompt_v3.txt         ← Agent_2 LLM instruction prompt
├── Agents/
│   ├── a2a_orchestrator.py   ← CLI + pipeline runner
│   ├── a2a_messages.py       ← message schema
│   ├── a2a_bus.py            ← in-memory bus + JSONL audit
│   ├── agent_1.py
│   ├── agent_2.py
│   └── a2a_logs/             ← audit logs per run
├── uploads/                  ← pricebooks uploaded via Streamlit
└── *.xlsx / *.xlsm           ← master / pricebook workbooks
```

---

## Setup

1. Use Python 3.12.13:

```bash
py -3.12 --version
```

2. Install dependencies from this folder:

```bash
py -3.12 -m pip install -r requirements.txt
```

3. Configure Azure OpenAI in `.env` (same folder as this README):

```env
AZURE_OPENAI_API_KEY=...
AZURE_OPENAI_ENDPOINT=https://....openai.azure.com/
DEPLOYMENT_NAME=...          # Agent_1
DEPLOYMENT_NAME_2=...        # Agent_2
API_VERSION=2025-01-01-preview
```

---

## How to run

### Streamlit UI (recommended)

From this project root (`Combine_AI_Agents/`):

```bash
py -3.12 -m streamlit run app.py
```

In the UI:

1. Confirm **Master pricelist** path (Agent_1)  
2. **Upload** the pricebook Excel (Agent_2)  
3. Leave **Use LLM agents** checked if you want Azure orchestration (default)  
4. Click **Run full A2A pipeline**  

You can also run **Agent_1 only** or **Agent_2 only** (Agent_2 needs an existing `active_price_list.xlsx`).

Results include Agent_1 counts, one Agent_2 summary, the A2A message timeline, and download buttons for:

- `active_price_list.xlsx`  
- `* - corrected.xlsx`

### CLI orchestrator

From `Agents/`:

```bash
# Full pipeline (Agent_1 → HANDOFF → Agent_2)
py -3.12 a2a_orchestrator.py --master "HK Master Price List - July 2026 OS_CLEANED.xlsm" --pricebook "AD_July2026 (en_HK).xlsx"

# Agent_1 only
py -3.12 a2a_orchestrator.py --agent1-only

# Agent_2 only (uses existing active_price_list.xlsx)
py -3.12 a2a_orchestrator.py --agent2-only --pricebook "AD_July2026 (en_HK).xlsx" --output "active_price_list.xlsx"

# Force LLM agents (Azure OpenAI)
py -3.12 a2a_orchestrator.py --llm --master "..." --pricebook "..."
```

Without `--llm`, tools run in **direct mode** (no LLM) — useful when Azure hits rate limits (429).

### Standalone agents

```bash
cd Agents
py -3.12 agent_1.py
py -3.12 agent_2.py          # direct Step 1 + Step 2
py -3.12 agent_2.py --agent  # LLM + prompt_v3.txt
```

---

## A2A messages

Local (in-process) bus — not Google A2A HTTP.

| Type | Meaning |
|------|---------|
| `TASK_START` | Orchestrator assigns work to an agent |
| `HANDOFF` | Agent_1 finished; artifact ready for Agent_2 |
| `QA_RESULT` | Agent_2 finished validation / attribute revise |
| `STATUS` | Progress update |
| `ERROR` | Failure; pipeline stops |

Audit logs are written under `Agents/a2a_logs/`.

**HANDOFF payload (example):**

```json
{
  "artifact_path": ".../active_price_list.xlsx",
  "active_count": 2841,
  "discontinued_removed": 35,
  "new_count": 547,
  "updated_count": 117,
  "status": "ready_for_qa"
}
```

---

## Match / correction rules (Agent_2)

- **Product key:** Brand Code + Model + Finish (different finish = different product)  
- **Compared fields:** Product Category Code, Product Category, Description, Unit, Price  
- **Not in master:** leave row unchanged, no yellow highlight  
- **Mismatch:** write master value, yellow highlight that cell only  
- **Attributes:** locale from filename (e.g. `en_HK`); map HardwareType → HardwareStandard ATTYPE1–4 and DHI  

---

## Notes

- Place master and pricebook workbooks in this project folder (or use absolute paths / Streamlit upload).  
- Corrected pricebooks keep all worksheets; validation helper columns are stripped.  
- LLM mode uses Azure deployments from `.env`; direct mode still performs the Excel work without calling the model for orchestration.
=======
# Agent-to-Agent-QA-QC-Pricebook
Pricebook QA/QC automation with local Agent-to-Agent (A2A) handoff: Agent_1 cleans the master list, Agent_2 validates and corrects the pricebook (Streamlit + Azure OpenAI).

