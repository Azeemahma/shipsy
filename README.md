# Shipsy Automation

This repo automates:
- Company revenue lookup + tiering
- Contact enrichment (LinkedIn, title, email)

## Files
- agent_workflow.json : Agentic workflow description
- agent_executor.py   : Python executor script
- requirements.txt    : Dependencies
- .env.example        : API keys configuration
- automation_output.xlsx : Output file (auto-filled)

## Usage
```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
export WORKFLOW_JSON=agent_workflow.json
python agent_executor.py
```
# shipsy
