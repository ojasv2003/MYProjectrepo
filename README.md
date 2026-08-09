# India Logistics Data Agent

Python service that automatically collects **Indian multi-modal logistics hub (MMLH / MMLP)** data from public websites and fills a **30-column** master CSV template using the **Cursor Python SDK** (`CURSOR_API_KEY`).

**Stack:** Python 3.11+, FastAPI, `cursor-sdk`, requests + BeautifulSoup, pydantic, python-dotenv. No paid search APIs. No Gemini / Google API key required.

## Architecture

```text
                  ┌─────────────────────┐
                  │   FastAPI / CLI     │
                  │  main.py / run_cli  │
                  └──────────┬──────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │    MasterAgent      │
                  │  agents/master_     │
                  │      agent.py       │
                  └──────────┬──────────┘
           ┌─────────────────┼─────────────────┐
           ▼                 ▼                 ▼
 ┌─────────────────┐ ┌──────────────┐ ┌─────────────────┐
 │  SearchAgent    │ │ FetcherAgent │ │  PipelineAgent  │
 │ DuckDuckGo HTML │ │ clean HTML   │ │ Cursor SDK →    │
 │ (no LLM quota)  │ │ keep tables  │ │ append CSV      │
 └─────────────────┘ └──────────────┘ └────────┬────────┘
                                               ▼
                              output/India_MMLH_Master_Collected_Data.csv
```

| Agent | File | Role |
| --- | --- | --- |
| Search | `agents/search_agent.py` | Given a topic, find candidate source URLs via DuckDuckGo HTML (free, no API key, does **not** use Cursor LLM quota). |
| Fetcher | `agents/fetcher_agent.py` | Download a URL, strip scripts/nav/footer/ads, preserve tables, return clean text. Handles timeouts and non-HTML. |
| Pipeline | `agents/pipeline_agent.py` | Send cleaned text to Cursor (`Agent.prompt`) with the template JSON schema in the prompt. Append rows; skip duplicates (`Hub_Name` + `Location`). |
| Master | `agents/master_agent.py` | Orchestrate per URL: fetch → pipeline. Return status `ok` / `fetch_failed` / `extract_failed` / `no_data` plus `rows_added`. |

Schema source of truth: `template_schema.py` (`TEMPLATE_COLUMNS` + `LogisticsHubRecord` with field descriptions that steer extraction).

## Setup (Windows PowerShell)

```powershell
# 1) Clone / open the repo, then create a virtual environment
cd path\to\this\repo
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2) Install dependencies
pip install -r requirements.txt

# 3) Configure Cursor API key
Copy-Item .env.example .env
notepad .env   # set CURSOR_API_KEY=... from https://cursor.com/dashboard/api

# 4) Start the API
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

**Cursor Cloud Agents:** add a Secret named `CURSOR_API_KEY` (Dashboard → API Keys → create key, then Cloud Agents → Secrets). The pipeline reads only `CURSOR_API_KEY` — it does **not** use `GEMINI_API_KEY`.

Open interactive docs at `http://127.0.0.1:8000/docs`.

### CLI smoke tests

```powershell
python run_cli.py --search "MMLP Nagpur logistics park capacity"
python run_cli.py https://nhlml.org/multi-modal-logistics-park
python run_cli.py --pretty https://nicdc.in/projects/4-projects-nearing-completion/integrated-multi-modal-logistics-hub-nangal-chaudhary
```

## API Endpoints

| Method | Path | Body | Description |
| --- | --- | --- | --- |
| `POST` | `/run` | `{"urls": ["https://..."]}` | Fetch + extract + save for each URL |
| `POST` | `/search` | `{"query": "...", "max_urls": 8}` | DuckDuckGo search, then run the pipeline |
| `GET` | `/results` | — | All collected rows as JSON |
| `GET` | `/download` | — | Download `India_MMLH_Master_Collected_Data.csv` |
| `GET` | `/template` | — | The 30 column names (in order) |
| `GET` | `/health` | — | Liveness + Cursor model name |

### Example requests

```powershell
# Health
Invoke-RestMethod http://127.0.0.1:8000/health

# Run on known URLs
Invoke-RestMethod -Method Post -Uri http://127.0.0.1:8000/run `
  -ContentType "application/json" `
  -Body '{"urls":["https://nhlml.org/multi-modal-logistics-park"]}'

# Search then collect
Invoke-RestMethod -Method Post -Uri http://127.0.0.1:8000/search `
  -ContentType "application/json" `
  -Body '{"query":"MMLP Nagpur logistics park capacity","max_urls":5}'

# Results / download
Invoke-RestMethod http://127.0.0.1:8000/results
Invoke-WebRequest http://127.0.0.1:8000/download -OutFile India_MMLH_Master_Collected_Data.csv
```

## Output

Rows are appended to:

```text
output/India_MMLH_Master_Collected_Data.csv
```

Duplicates with the same **Hub_Name + Location** (case-insensitive) are skipped. One failing URL never aborts the rest of the batch; each URL gets its own status in the report.

## Environment variables

| Variable | Required | Default | Purpose |
| --- | --- | --- | --- |
| `CURSOR_API_KEY` | Yes (for extract) | — | Cursor user/service API key from [cursor.com/dashboard/api](https://cursor.com/dashboard/api) |
| `CURSOR_MODEL` | No | `composer-2.5` | Model id passed to `Agent.prompt` (must be available on your account) |
| `LOG_LEVEL` | No | `INFO` | Standard logging level |
| `HOST` / `PORT` | No | `0.0.0.0` / `8000` | Used when launching via `python main.py` |

## Project layout

```text
agents/
  search_agent.py
  fetcher_agent.py
  pipeline_agent.py
  master_agent.py
template_schema.py
main.py
run_cli.py
requirements.txt
.env.example
output/
```

## Notes

- Search uses DuckDuckGo HTML/Lite endpoints only — it never consumes Cursor LLM quota.
- Fetcher strips scripts, nav, footer, and common ad/cookie chrome, but keeps table content as `col | col` rows so the model can read capacity/investment tables.
- Pipeline calls Cursor `Agent.prompt` with the `ExtractionResult` / `LogisticsHubRecord` JSON schema embedded in the prompt, then validates the reply with pydantic.
- Extraction runs in an isolated temp workspace with sandboxing enabled so the agent does not touch the project files.
- All stages use the standard `logging` module.
