# CompanyInfo UA on Render

This repository contains a Render-ready backend for generating a company report by EDRPOU.

## Safe default

Test requests should not spend money.

The service now defaults to mock mode:

- `POST /api/generate-report` uses `live_mode: false` by default
- `POST /api/generate-report-file` uses `live_mode: false` by default
- in mock mode, OpenAI API is not called
- the service still generates a real DOCX file for testing

Paid live mode is only enabled when both conditions are true:

1. request body contains `"live_mode": true`
2. Render environment contains `ALLOW_LIVE_OPENAI=true`

## Main endpoints

- `GET /health`
- `GET /api/agent-context`
- `GET /downloads/docx-template`
- `POST /api/render-map`
- `POST /api/generate-report`
- `POST /api/generate-report-file`

## Render environment variables

- `OPENAI_API_KEY`
- `OPENAI_MODEL` default: `gpt-5.4-mini`
- `OPENAI_WEB_SEARCH_TOOL_TYPE` default: `web_search`
- `OPENAI_REASONING_EFFORT` default: `medium`
- `ALLOW_LIVE_OPENAI` default: `false`

## Free test request

```json
{
  "edrpou": "12345678"
}
```

## Paid live request

```json
{
  "edrpou": "12345678",
  "live_mode": true
}
```

## Local run

```bash
pip install -r requirements.txt
uvicorn app:app --reload
```
