import base64

from fastapi import (
    FastAPI,
    File,
    UploadFile,
    HTTPException,
    Form,
)

from fastapi.middleware.cors import CORSMiddleware

from .reconciliation import reconcile


# ============================================================
# APPLICATION FASTAPI
# ============================================================

app = FastAPI(
    title="LEONI Stock Reconciliation API",
    version="1.0.0"
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,

    allow_origins=["*"],

    allow_credentials=False,

    allow_methods=["*"],

    allow_headers=["*"],
)


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/api/health")
def health():

    return {
        "status": "ok"
    }


# ============================================================
# RECONCILIATION
# ============================================================

@app.post("/api/reconcile")
async def api_reconcile(

    # --------------------------------------------------------
    # PLANT
    # --------------------------------------------------------

    plant: str = Form(...),

    # --------------------------------------------------------
    # FICHIER SI
    # --------------------------------------------------------

    si_file: UploadFile = File(...),

    # --------------------------------------------------------
    # FICHIER SF
    # --------------------------------------------------------

    sf_file: UploadFile = File(...),

    # --------------------------------------------------------
    # FMALS032 #1
    # --------------------------------------------------------

    movement_file_1: UploadFile = File(...),

    # --------------------------------------------------------
    # FMALS032 #2 OPTIONNEL
    # --------------------------------------------------------

    movement_file_2: UploadFile | None = File(None),
):

    try:

        # ====================================================
        # VALIDATION PLANT
        # ====================================================

        plant = str(
            plant
        ).strip()

        allowed_plants = {
            "21",
            "47",
            "48",
            "66",
            "77"
        }

        if plant not in allowed_plants:

            raise HTTPException(
                status_code=400,
                detail=(
                    f"Plant '{plant}' invalide. "
                    "Plants disponibles : "
                    "21, 47, 48, 66, 77."
                )
            )

        # ====================================================
        # LECTURE DES FICHIERS
        # ====================================================

        si = await si_file.read()

        sf = await sf_file.read()

        m1 = await movement_file_1.read()

        m2 = (
            await movement_file_2.read()
            if movement_file_2
            else None
        )

        # ====================================================
        # RÉCONCILIATION
        # ====================================================

        excel, summary, preview = reconcile(

            si_bytes=si,

            sf_bytes=sf,

            movement1_bytes=m1,

            movement2_bytes=m2,

            plant=plant,
        )

        # ====================================================
        # RÉPONSE
        # ====================================================

        return {

            "success": True,

            "plant": plant,

            "summary": summary,

            "preview": preview,

            "excel_base64":
                base64.b64encode(
                    excel
                ).decode("ascii"),

            "filename":
                "Reconciliation_Stock_Automatique.xlsx",
        }

    except HTTPException:

        raise

    except Exception as exc:

        raise HTTPException(
            status_code=400,
            detail=str(exc)
        ) from exc
