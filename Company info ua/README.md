# CompanyInfo UA on Render

Цей репозиторій тепер містить Render-ready backend для серверного формування довідки по ЄДРПОУ.

В основі все ще лежать наявні артефакти:

- DOCX-шаблон і головну інструкцію для формування довідки по ЄДРПОУ
- SVG/JSON-пакет і Python-рендерер карти активів бізнес-групи

Доданий Python API-сервіс, який:

- приймає код ЄДРПОУ
- викликає OpenAI API на сервері через `OPENAI_API_KEY`
- використовує веб-пошук моделі для збору підтверджених фактів
- рендерить карту бізнес-активів
- формує готовий `.docx` на базі `CompanyInfo_UA_report_template.docx`
- віддає інструкції та метадані агента через HTTP
- віддає DOCX-шаблон
- рендерить карту бізнес-активів через існуючий `render_business_map.py`

## Що це вирішує

Це прибирає залежність від локальної машини та тарифу зовнішнього користувача в частині генерації моделю OpenAI та дозволяє:

- централізовано зберігати інструкції й шаблони
- використовувати один серверний URL для всіх зовнішніх користувачів
- перенести рендер карти на сервер
- перенести сам виклик моделі на ваш серверний API-ключ

## Локальний запуск

```bash
pip install -r requirements.txt
uvicorn app:app --reload
```

## API

- `GET /health` — healthcheck
- `GET /api/agent-context` — інструкції, шаблон і параметри карти
- `GET /downloads/docx-template` — завантаження DOCX-шаблону
- `POST /api/render-map` — рендер карти
- `POST /api/generate-report` — повний запуск агента, повертає `docx_base64`
- `POST /api/generate-report-file` — повний запуск агента, повертає готовий `.docx` файлом

## Змінні середовища Render

- `OPENAI_API_KEY` — обов'язково
- `OPENAI_MODEL` — необов'язково, за замовчуванням `gpt-5.4-mini`
- `OPENAI_WEB_SEARCH_TOOL_TYPE` — необов'язково, за замовчуванням `web_search`
- `OPENAI_REASONING_EFFORT` — необов'язково, за замовчуванням `medium`

## Версія Python

У репозиторії додано `.python-version` = `3.13`.
Це важливо для Render, тому що станом на липень 2026 його дефолтна Python-версія для нових сервісів — `3.14.3`, а окремі залежності на `3.14` можуть вимагати локальну збірку замість готових wheel-пакетів.

Приклад `POST /api/render-map`:

```json
{
  "data": {
    "group_name": "Назва групи",
    "show_title": false,
    "show_legend": false,
    "assets": [
      {
        "name": "Завод",
        "type": "production",
        "region_id": "UA12",
        "lat": 48.0,
        "lon": 33.0,
        "city": "Кривий Ріг"
      }
    ]
  },
  "include_png": true
}
```

У відповіді повертаються `svg_base64` і, якщо ввімкнено, `png_base64`.

Приклад `POST /api/generate-report-file`:

```json
{
  "edrpou": "12345678",
  "company_name_hint": "Необов'язково",
  "additional_instructions": "Необов'язково"
}
```

## Деплой на Render

1. Завантажити репозиторій у GitHub.
2. У Render створити `Web Service`.
3. Вказати репозиторій.
4. Render підхопить `render.yaml`.
5. У `Environment` додати `OPENAI_API_KEY`.
5. Після деплою перевірити:

```text
https://<your-service>.onrender.com/health
https://<your-service>.onrender.com/api/agent-context
```
