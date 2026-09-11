# LEONI Stock Reconciliation — FastAPI backend

Backend extracted from `Réconciliation 2(1).ipynb` for a web application.

## What it does

Accepts:
- SI CSV (required)
- SF CSV (required)
- FMALS032 Movement 1 CSV (required)
- FMALS032 Movement 2 CSV (optional)

Returns:
- reconciliation preview
- KPI summary
- Excel workbook with `Réconcil`, `TCD1`, `TCD2`

## Important

The notebook remains the business-logic reference. The `SIGNES` dictionary is preserved exactly. FMALS032 `BookType` is normalized (`01 -> 1`, `02 -> 2`, etc.) before creating `Mouv`, which is important for the second movement file.

Uploaded files are processed in memory by the API and are not intentionally persisted to disk.

## Local run

```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Health check: `GET /api/health`

Reconciliation endpoint: `POST /api/reconcile` using multipart fields:
`si_file`, `sf_file`, `movement_file_1`, optional `movement_file_2`.

## Frontend integration

The frontend should POST `FormData` to `/api/reconcile`. The response contains `summary`, `preview`, and `excel_base64`. Decode `excel_base64` into a Blob and trigger the Excel download in the browser.
