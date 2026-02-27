# Yahoo Chart MCP Agent (n8n + Ollama + Python)

This project provides a Python MCP tool that does the full workflow:

1. Open Yahoo Finance quote page (`https://au.finance.yahoo.com/quote/HUB.AX/` or `{CODE}.AX`)
2. Capture the chart image
3. Insert the image into a Microsoft Word (`.docx`) file
4. Convert the Word file to PDF
5. Email the PDF with Apple Mail

The MCP tool is intended for use from an n8n AI Agent with Ollama.

## Prerequisites

- macOS (required for Apple Mail automation and `docx2pdf` via Word)
- Microsoft Word installed
- Apple Mail configured with a sending account
- Python 3.10+
- n8n running
- Ollama running locally (`http://localhost:11434`)

## Install and Initial Setup

```bash
python3 -m venv .venv 
source .venv/bin/activate
pip install -e . 
playwright install chromium
```
this:
makes an isolated python virtual environment
activates venv so uses .venv's python, not system python
installs packages in editable mode so local code changes are picked up
installs chromium browser package for playwright screenshots

## Run (Two Terminals)

Use two separate terminal windows/tabs.

### Terminal 1: Start MCP Server

```bash
cd ~/Desktop/asx-chart-n8n-integrated-tool
source .venv/bin/activate 
pip install -e . 
set -a; source .env; set +a 
asx-mcp-server 
```
this:
goes to project folder
activates this project's virtual environment
ensures editable install and dependencies are present
loads .env settings from shell into program (uses -a auto-export)
starts MCP server for n8n to connect to

By default this starts an SSE MCP server at:

- `http://127.0.0.1:8001/sse`

### Terminal 2: Start n8n

```bash
npx n8n #start the n8n editor and runtime
```

n8n editor opens at:

- `http://localhost:5678`

## MCP Tools Available

The MCP server currently exposes exactly these tools:

1. `create_asx_report`

Arguments for `create_asx_report`:

- `asx_code` (string): e.g. `BHP`, `CBA`, `CSL`; URL becomes `{ASX_CODE}.AX` on Yahoo Finance (default `HUB`)
- `recipient` (string): default `test@gmail.com`
- `output_dir` (string): default `output`
- `email_subject` (string): optional
- `email_body` (string): optional
- `send_email` (boolean): default `true`

Note: functions in `src/asx_mcp/pipeline.py` are internal helper functions for this one MCP tool; they are not separate MCP tools.

## n8n Wiring (AI Agent + Ollama + MCP)

Create a workflow with these nodes:

1. `Chat Trigger` (or Scheduled timing for execution)
2. `AI Agent`
3. `Ollama Chat Model` (connect to AI Agent)
4. `MCP Client Tool` (connect to AI Agent tools)

### MCP Client Tool Settings

- `SSE Endpoint`: `http://127.0.0.1:8001/sse`

### Suggested AI Agent System Instruction

```text
You are a Yahoo Finance chart reporting agent.
When the user asks for a stock chart report, call tool create_asx_report.
Always default recipient to test@gmail.com unless user provides a different email.
If user gives a ticker code, pass it as asx_code (without .AX).
If user does not provide a ticker, use HUB.
Do not invent output_dir values like /path/to/local/directory/.
```

### Example User Message in n8n Chat

```text
Create today's Yahoo Finance chart report for HUB and email it to test@gmail.com
```

## Local CLI Test (without MCP)

Run this in a separate terminal tab/window while your virtualenv is active.

```bash
asx-report-cli --asx-code HUB --recipient test@gmail.com
```

Dry run (build files only, skip email):

```bash
asx-report-cli --asx-code HUB --no-email
```

## Notes

- Yahoo Finance DOM can change; if chart capture fails, update selectors in `src/asx_mcp/pipeline.py`.
- `docx2pdf` requires Microsoft Word to be installed.
- Mail sending may trigger macOS automation permission prompts on first run.
- Word file access prompts are a one-time macOS privacy grant; click `Allow` once for the project folder.
- macOS Automation/Files & Folders permissions cannot be auto-approved by code; they must be allowed once in System Settings.
- `ASX_WATCH_WINDOW_MS` (default `20000`) controls how long capture waits for a rendered chart before failing.
- `ASX_WATCH_POLL_MS` (default `120`) controls how often capture checks whether the chart has rendered.
- `MCP_DEDUPE_SECONDS` (default `90`) suppresses duplicate tool calls with the same inputs so n8n retries do not send twice.
- `MCP_ERROR_DEDUPE_SECONDS` (default `25`) suppresses immediate re-runs after a recent failure on the same request.
- `ASX_PDF_ENGINE` (default `auto`) tries headless LibreOffice first (if installed), then falls back to `docx2pdf`.
- `ASX_LIBREOFFICE_TIMEOUT_SECONDS` (default `120`) controls how long headless LibreOffice conversion can run.
- Capture uses a single page load with no reloads and no cookie-button clicks; it screenshots only after the chart is rendered.
- If a site serves placeholder content to headless browsers, set `ASX_HEADLESS=false` in `.env` and retry.

## More notes for myself

cli.py
- NOTE CLI.PY ISN'T USED WHEN CALLED IN N8N VIA MCP
    - cli.py is command-line wrapper around the pipeline (in pipeline.py), it parses arguments, calls run_asx_report function
    - ONLY USED WHEN RUNNING LOCAL COMMAND asx-report-cli for local testing, or an n8n Execute Command style workflow (NOT as an MCP tool in an AI Agent workflow)

init.py
- does basically NOTHING EXCEPT if someone wants to import all my code from the src/asx_mcp folder as a package using "from asx_mcp import *", it imports the right modules needed without breaking things (hopefully).

pipeline.py
- does the whole end-to-end implementation
- defines configuration for Yahoo chart capture and validation
- normalise asx company name into a variable (securities/ticker code) to put into url (e.g. HUB)
- reads env variables
- validates webpage host/path and checks screenshot is valid
- opens chromium browser using playwright, waits for chart to show, captures in screenshot
- create docx containing title, timestamp, URL, and screenshot of chart
- converts docx to pdf (using libreoffice or docx2pdf as fallback)
- sends email with apple mail via osascript, attaching the PDF
- puts files in writable output directory
- runs whole pipeline in order, returns a result dictionary(with image_path, docx_path, pdf_path, final_url, etc.)

server.py
- creates MCP Server object using FastMcp import
- registers MCP tool create_asx_report
- when tool called, runs the actual workflow from pipeline.py in a worker thread 
(similar to a child process, but shares same memory space)
- has a deduping system using ticker + recipient + send flag, stops duplicate runs
- reads server/env settings (host, port, sse paths, transport) and starts MCP in main()

i.e. pipeline.py does the work, server.py exposes it safely to n8n and stops duplicate executions.

FOR NOW, only final mcp tool is exposed to reduce agent mistakes, like wrong tool order, missing args, and more duplicate sends. 
Can still easily make all MCP tools available seperately by adding more @mcp.tool() functions in server.py
Each wrapper can call internal functions from pipeline.py as defined/needed.


pyproject.toml
- the manifest file on packages, config, etc.
- tells Python how to build, install, and run the app
- important bit is it defines the command-line entrypoints; like asx-mcp-server running asx_mcp.server:main
and asx-report-cli running asx_mcp.cli:main.

n8n via MCP workflow
- n8n mcp client talks to the asx-mcp-server (this is defined as mcp server entrypoint in pyproject.toml), which runs server.py file.
- server.py exposes the MCP tool and calls the pipeline directly in pipeline.py

# note if downloading off git repo
.env wasn't uploaded, only .env.example, so need to run cd ~/Desktop/asx-chart-n8n-integrated-tool and cp .env.example .env after and make sure set -a; source .env; set +a has been run before running program.
also need to rerun Install and Initial Setup as need to make .venv as well