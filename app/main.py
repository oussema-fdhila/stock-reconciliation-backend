import base64
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware

from .reconciliation import reconcile

app = FastAPI(title="LEONI Stock Reconciliation API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/health")
def health():
    return {"status": "ok"}

@app.post("/api/reconcile")
async def api_reconcile(
    si_file: UploadFile = File(...),
    sf_file: UploadFile = File(...),
    movement_file_1: UploadFile = File(...),
    movement_file_2: UploadFile | None = File(None),
):
    try:
        si = await si_file.read()
        sf = await sf_file.read()
        m1 = await movement_file_1.read()
        m2 = await movement_file_2.read() if movement_file_2 else None
        excel, summary, preview = reconcile(si, sf, m1, m2)
        return {
            "success": True,
            "summary": summary,
            "preview": preview,
            "excel_base64": base64.b64encode(excel).decode("ascii"),
            "filename": "Reconciliation_Stock_Automatique.xlsx",
        }
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
