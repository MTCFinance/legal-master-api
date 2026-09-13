import os
import re
import secrets
from pathlib import Path
from typing import Optional, List
from uuid import uuid4

from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.responses import FileResponse
from fastapi.security import APIKeyHeader
from pydantic import BaseModel, Field
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


app = FastAPI(
    title="Legal Master API",
    version="1.3.0",
    description="API për Hartues Aktesh Juridike - Kosovë / ARBK",
    servers=[
        {
            "url": "https://legal-master-api.onrender.com",
            "description": "Production server"
        }
    ]
)


# =========================================================
# CONFIG
# =========================================================

BASE_DIR = Path("/tmp/legal-master")
BASE_DIR.mkdir(parents=True, exist_ok=True)

DOWNLOAD_TOKENS = {}


# =========================================================
# API KEY SECURITY
# =========================================================

api_key_header = APIKeyHeader(
    name="X-API-Key",
    auto_error=False
)


def verify_api_key(
    api_key: Optional[str] = Depends(api_key_header)
):
    expected_key = os.getenv("LEGAL_MASTER_API_KEY")

    if not expected_key:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="API key is not configured on the server"
        )

    if api_key is None or not secrets.compare_digest(
        api_key,
        expected_key
    ):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or missing API key"
        )

    return True


# =========================================================
# MODELS
# =========================================================

class HealthResponse(BaseModel):
    status: str
    service: str


class BusinessRequest(BaseModel):
    company_name: str
    legal_form: str = "SHPK"
    address: Optional[str] = None
    capital: Optional[float] = None
    employees: Optional[int] = None


class DocumentRequest(BaseModel):
    document_type: str
    title: str
    body: str
    filename: Optional[str] = None


class ARBKPerson(BaseModel):
    full_name: str
    role: Optional[str] = None


class ARBKBusinessData(BaseModel):
    business_name: str
    trade_name: Optional[str] = None
    nui: str

    legal_form: Optional[str] = None
    business_status: Optional[str] = None
    registration_date: Optional[str] = None

    address: Optional[str] = None
    municipality: Optional[str] = None

    primary_activity: Optional[str] = None
    other_activities: List[str] = Field(default_factory=list)

    owners: List[ARBKPerson] = Field(default_factory=list)
    directors: List[ARBKPerson] = Field(default_factory=list)

    email: Optional[str] = None
    phone: Optional[str] = None

    source_url: Optional[str] = None
    verification_date: Optional[str] = None


# =========================================================
# HELPERS
# =========================================================

def clean_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_\- ]", "", name)
    name = name.strip().replace(" ", "_")

    if not name:
        name = "dokument_juridik"

    return name[:100]


def create_word_document(
    data: DocumentRequest,
    output_path: Path
):
    document = Document()

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title_run = title.add_run(data.title.upper())
    title_run.bold = True
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(14)

    document.add_paragraph()

    blocks = data.body.split("\n")

    for block in blocks:
        text = block.strip()

        if not text:
            continue

        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if text.lower().startswith("neni "):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run(text)
            run.bold = True
        else:
            run = paragraph.add_run(text)

        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    document.save(output_path)


# =========================================================
# ROOT
# =========================================================

@app.get(
    "/",
    operation_id="root"
)
def root():
    return {
        "service": "Legal Master API",
        "status": "running",
        "version": "1.3.0"
    }


# =========================================================
# HEALTH
# =========================================================

@app.get(
    "/health",
    response_model=HealthResponse,
    operation_id="health",
    dependencies=[Depends(verify_api_key)]
)
def health():
    return {
        "status": "ok",
        "service": "legal-master-api"
    }


# =========================================================
# BUSINESS VALIDATION
# =========================================================

@app.post(
    "/business/validate",
    operation_id="validate_business",
    dependencies=[Depends(verify_api_key)]
)
def validate_business(data: BusinessRequest):
    issues = []

    if not data.company_name.strip():
        issues.append("Mungon emri i biznesit")

    if data.capital is not None and data.capital < 0:
        issues.append("Kapitali nuk mund të jetë negativ")

    if data.employees is not None and data.employees < 0:
        issues.append(
            "Numri i punëtorëve nuk mund të jetë negativ"
        )

    return {
        "valid": len(issues) == 0,
        "issues": issues,
        "data": data
    }


# =========================================================
# ARBK PUBLIC DATA VERIFICATION
# =========================================================

@app.post(
    "/arbk/verify-business",
    operation_id="verify_arbk_business",
    dependencies=[Depends(verify_api_key)]
)
def verify_arbk_business(data: ARBKBusinessData):

    critical_issues = []
    warnings = []

    if not data.business_name.strip():
        critical_issues.append(
            "Mungon emri i saktë i biznesit"
        )

    if not data.nui.strip():
        critical_issues.append(
            "Mungon NUI"
        )

    if data.business_status is None:
        warnings.append(
            "Statusi i biznesit nuk është dhënë"
        )

    if data.address is None:
        warnings.append(
            "Adresa e biznesit nuk është dhënë"
        )

    if data.legal_form is None:
        warnings.append(
            "Forma juridike nuk është dhënë"
        )

    if not data.directors:
        warnings.append(
            "Nuk është dhënë drejtori/përfaqësuesi"
        )

    if not data.owners:
        warnings.append(
            "Nuk janë dhënë pronarët/anëtarët"
        )

    if data.email is None:
        warnings.append(
            "Emaili nuk është publik ose nuk është dhënë"
        )

    if data.phone is None:
        warnings.append(
            "Telefoni nuk është publik ose nuk është dhënë"
        )

    normalized = {
        "business_name": data.business_name.strip(),
        "trade_name": data.trade_name,
        "nui": data.nui.strip(),
        "legal_form": data.legal_form,
        "business_status": data.business_status,
        "registration_date": data.registration_date,
        "address": data.address,
        "municipality": data.municipality,
        "primary_activity": data.primary_activity,
        "other_activities": data.other_activities,
        "owners": [
            person.model_dump()
            for person in data.owners
        ],
        "directors": [
            person.model_dump()
            for person in data.directors
        ],
        "email": data.email,
        "phone": data.phone,
        "source_url": data.source_url,
        "verification_date": data.verification_date
    }

    return {
        "valid": len(critical_issues) == 0,
        "status": (
            "VERIFIED_DATA"
            if len(critical_issues) == 0
            else "INCOMPLETE_DATA"
        ),
        "critical_issues": critical_issues,
        "warnings": warnings,
        "business": normalized,
        "note": (
            "Përdoren vetëm të dhënat publike të dhëna "
            "nga burimi ARBK. Të dhënat që nuk publikohen "
            "nuk duhet të supozohen ose shpiken."
        )
    }


# =========================================================
# WORD DOCUMENT GENERATION
# =========================================================

@app.post(
    "/documents/generate",
    operation_id="generate_document",
    dependencies=[Depends(verify_api_key)]
)
def generate_document(data: DocumentRequest):

    if not data.title.strip():
        raise HTTPException(
            status_code=400,
            detail="Titulli i dokumentit mungon"
        )

    if not data.body.strip():
        raise HTTPException(
            status_code=400,
            detail="Përmbajtja e dokumentit mungon"
        )

    if data.filename:
        base_filename = clean_filename(data.filename)
    else:
        base_filename = clean_filename(
            f"{data.document_type}_{uuid4().hex[:8]}"
        )

    filename = f"{base_filename}.docx"

    internal_name = f"{uuid4().hex}_{filename}"
    output_path = BASE_DIR / internal_name

    create_word_document(
        data=data,
        output_path=output_path
    )

    download_token = secrets.token_urlsafe(32)

    DOWNLOAD_TOKENS[download_token] = {
        "path": str(output_path),
        "filename": filename
    }

    download_url = (
        "https://legal-master-api.onrender.com"
        f"/files/{download_token}"
    )

    return {
        "success": True,
        "document_type": data.document_type,
        "filename": filename,
        "download_url": download_url
    }


# =========================================================
# TEMPORARY DOWNLOAD
# =========================================================

@app.get(
    "/files/{token}",
    operation_id="download_document"
)
def download_document(token: str):

    file_data = DOWNLOAD_TOKENS.get(token)

    if not file_data:
        raise HTTPException(
            status_code=404,
            detail="Dokumenti nuk ekziston ose linku ka skaduar"
        )

    file_path = Path(file_data["path"])

    if not file_path.exists():
        raise HTTPException(
            status_code=404,
            detail="Dokumenti nuk gjendet më në server"
        )

    return FileResponse(
        path=file_path,
        filename=file_data["filename"],
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )
