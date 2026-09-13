import os
import secrets
from typing import Optional

from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.security import APIKeyHeader
from pydantic import BaseModel


app = FastAPI(
    title="Legal Master API",
    version="1.1.0",
    description="API për Hartues Aktesh Juridike - Kosovë / ARBK",
    servers=[
        {
            "url": "https://legal-master-api.onrender.com",
            "description": "Production server"
        }
    ]
)


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


# =========================================================
# PUBLIC ROOT
# =========================================================

@app.get(
    "/",
    operation_id="root"
)
def root():
    return {
        "service": "Legal Master API",
        "status": "running"
    }


# =========================================================
# PROTECTED HEALTH CHECK
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
