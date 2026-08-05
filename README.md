# India Logistics Data Agent

Python service that automatically collects **Indian multi-modal logistics hub (MMLH / MMLP)** data from public websites and fills a **30-column** master CSV template using Google Gemini structured output.

**Stack:** Python 3.11+, FastAPI, `google-genai`, requests + BeautifulSoup, pydantic, python-dotenv. No paid search APIs.

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
 │ DuckDuckGo HTML │ │ clean HTML   │ │ Gemini schema → │
 │ (no Gemini)     │ │ keep tables  │ │ append CSV      │
 └─────────────────┘ └──────────────┘ └────────┬────────┘
                                               ▼
                              output/India_MMLH_Master_Collected_Data.csv
```

| Agent | File | Role |
| --- | --- | --- |
| Search | `agents/search_agent.py` | Given a topic, find candidate source URLs via DuckDuckGo HTML (free, no API key, does **not** use Gemini). |
| Fetcher | `agents/fetcher_agent.py` | Download a URL, strip scripts/nav/footer/ads, preserve tables, return clean text. Handles timeouts and non-HTML. |
| Pipeline | `agents/pipeline_agent.py` | Send cleaned text to Gemini using the template as a structured-output schema. Append rows; skip duplicates (`Hub_Name` + `Location`). |
| Master | `agents/master_agent.py` | Orchestrate per URL: fetch → pipeline. Return status `ok` / `fetch_failed` / `extract_failed` / `no_data` plus `rows_added`. |

Schema source of truth: `template_schema.py` (`TEMPLATE_COLUMNS` + `LogisticsHubRecord` with field descriptions for Gemini).

## Setup (Windows PowerShell)

```powershell
# 1) Clone / open the repo, then create a virtual environment
cd path\to\this\repo
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2) Install dependencies
pip install -r requirements.txt

# 3) Configure Gemini
Copy-Item .env.example .env
notepad .env   # set GEMINI_API_KEY=...

# 4) Start the API
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

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
| `GET` | `/health` | — | Liveness + Gemini model name |

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
| `GEMINI_API_KEY` | Yes (for extract) | — | Google AI Studio / Gemini API key |
| `GEMINI_MODEL` | No | `gemini-2.5-flash` | Model used for structured extraction |
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

- Search uses DuckDuckGo HTML/Lite endpoints only — it never consumes Gemini quota.
- Fetcher strips scripts, nav, footer, and common ad/cookie chrome, but keeps table content as `col | col` rows so Gemini can read capacity/investment tables.
- Pipeline asks Gemini for JSON shaped by `ExtractionResult` / `LogisticsHubRecord` (`response_json_schema`).
- All stages use the standard `logging` module.
