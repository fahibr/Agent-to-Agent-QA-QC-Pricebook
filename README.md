# Pricebook QA/QC — Agent-to-Agent (A2A) Pipeline

Automates master pricelist cleaning and pricebook QA/QC with two cooperating LangChain agents, a local A2A message bus, and a Streamlit UI.

**Requires:** Python **3.12.13**

---

## Workflow

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

| Step | Agent | Action | Output |
|------|--------|--------|--------|
| 1 | **Agent_1** | Merge master sheets, flag new/updated/discontinued, remove discontinued, export active list | `active_price_list.xlsx` + **HANDOFF** |
| 2 | **Agent_2** | Validate pricebook vs active master; correct mismatches (yellow); revise ATTR/DHI | `* - corrected.xlsx` + **QA_RESULT** |

If Agent_1 fails or the handoff artifact is missing, Agent_2 does **not** run.

---

## Agents

### Agent_1 — Master pricelist

**Input:** Master Excel (e.g. `HK Master Price List - July 2026 OS_CLEANED.xlsm`)  
**Prompt:** `prompt/prompt_agent1.txt`

1. `load_price_list` — merge sheets; Status from colors (red = discontinued, yellow = updated fields, green = new)  
2. `remove_discontinued_models` — keep Active (+ New) only  
3. `export_active_list` — write `active_price_list.xlsx`  

**HANDOFF metrics** (discontinued count is taken *before* removal):

- `active_count`, `discontinued_removed`, `new_count`, `updated_count`  
- `artifact_path`, `status=ready_for_qa`

### Agent_2 — Pricebook QA/QC

**Inputs:** Uploaded pricebook + Agent_1’s `active_price_list.xlsx`  
**Prompt:** `prompt/prompt_v4.txt` (Agent_2 section; Agent_1 tools are not called)

1. **`validate_pricebook`** — match Brand + Model + Finish; correct Product Category Code/Category, Description, Unit, Price; yellow-highlight only corrected cells; “not in master” unchanged (count only)  
2. **`revise_pricebook_attributes`** — locale from filename; HardwareType blanks → `NULL`; sync ATTYPE1–4 and DHI from HardwareType  

Original files are never overwritten. Corrections go to `* - corrected.xlsx`.

---

## Project layout

```text
Combine_AI_Agents/
├── README.md
├── app.py                 ← Streamlit UI (run from here)
├── requirements.txt
├── .env                   ← Azure OpenAI (not committed)
├── prompt/
│   ├── prompt_agent1.txt  ← Agent_1 LLM prompt
│   ├── prompt_v1.txt … v3.txt   ← earlier Agent_2 prompts
│   └── prompt_v4.txt      ← enhanced A2A prompt (Agent_1 + Agent_2)
├── Agents/
│   ├── a2a_orchestrator.py
│   ├── a2a_messages.py
│   ├── a2a_bus.py
│   ├── agent_1.py
│   ├── agent_2.py
│   └── a2a_logs/
├── uploads/               ← pricebooks uploaded in Streamlit
└── *.xlsx / *.xlsm
```

---

## Setup

```bash
py -3.12 --version
py -3.12 -m pip install -r requirements.txt
```

Create `.env` in this folder:

```env
AZURE_OPENAI_API_KEY=...
AZURE_OPENAI_ENDPOINT=https://....openai.azure.com/
DEPLOYMENT_NAME=...       # Agent_1
DEPLOYMENT_NAME_2=...     # Agent_2
API_VERSION=2025-01-01-preview
```

---

## How to run

### Streamlit UI (recommended)

From this project root:

```bash
py -3.12 -m streamlit run app.py
```

1. Set **Master pricelist** path (Agent_1)  
2. **Upload** the pricebook Excel (Agent_2)  
3. **Use LLM agents** is on by default (Azure OpenAI)  
4. Click **Run full A2A pipeline** (or Agent_1 / Agent_2 only)

**Results panel**

- Metrics: Active, Discontinued removed, New, Updated fields  
- One **Agent_1 summary** and one **Agent_2 summary** (not duplicated)  
- Downloads: `active_price_list.xlsx`, `* - corrected.xlsx`  
- A2A message timeline + audit log path  

### CLI orchestrator

From `Agents/`:

```bash
# Full pipeline
py -3.12 a2a_orchestrator.py --master "HK Master Price List - July 2026 OS_CLEANED.xlsm" --pricebook "AD_July2026 (en_HK).xlsx"

# Agent_1 only
py -3.12 a2a_orchestrator.py --agent1-only

# Agent_2 only (needs existing active_price_list.xlsx)
py -3.12 a2a_orchestrator.py --agent2-only --pricebook "AD_July2026 (en_HK).xlsx" --output "active_price_list.xlsx"

# LLM mode
py -3.12 a2a_orchestrator.py --llm --master "..." --pricebook "..."
```

Without `--llm`, tools run in **direct mode** (no Azure calls) — useful on 429 rate limits.

### Standalone agents

```bash
cd Agents
py -3.12 agent_1.py              # LLM + prompt_agent1.txt
py -3.12 agent_2.py              # direct validate + revise
py -3.12 agent_2.py --agent      # LLM + prompt_v4.txt
```

---

## Prompts

| File | Used by | Role |
|------|---------|------|
| `prompt/prompt_agent1.txt` | Agent_1 | Master clean → export → HANDOFF |
| `prompt/prompt_v4.txt` | Agent_2 (+ A2A docs) | Full pipeline doc; runtime scoped to Agent_2 tools |
| `prompt/prompt_v3.txt` | Legacy | Agent_2-only (validate + revise) |
| `prompt/prompt_v1.txt`, `v2.txt` | Legacy | Earlier Agent_2 drafts |

---

## A2A messages

In-process bus (not Google A2A HTTP). Audit logs: `Agents/a2a_logs/`.

| Type | Meaning |
|------|---------|
| `TASK_START` | Work assigned to an agent |
| `HANDOFF` | Agent_1 done; artifact ready for Agent_2 |
| `QA_RESULT` | Agent_2 done |
| `STATUS` | Progress |
| `ERROR` | Failure; pipeline stops |

**HANDOFF example:**

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

## Agent_2 match rules

- **Key:** Brand Code + Model + Finish (different finish = different product)  
- **Compared:** Product Category Code, Product Category, Description, Unit, Price  
- **Not in master:** unchanged, no highlight (count only)  
- **Mismatch:** master value written, yellow highlight on that cell only  
- **Attributes:** locale from filename (e.g. `en_HK`); HardwareType → ATTYPE1–4 / DHI  

---

## Notes

- Workbooks can live in this folder, or use absolute paths / Streamlit upload for the pricebook.  
- Corrected pricebooks keep all worksheets; no extra validation columns.  
- LLM mode uses Azure deployments from `.env`; direct mode still runs the Excel tools without LLM orchestration.
