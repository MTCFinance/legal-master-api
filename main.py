from fastapi import FastAPI
from pydantic import BaseModel
from typing import Optional

app = FastAPI(
    title="Legal Master API",
    version="1.0.0",
    description="API për Hartues Aktesh Juridike - Kosovë / ARBK",
    servers=[
        {
            "url": "https://legal-master-api.onrender.com",
            "description": "Production server"
        }
    ]
)

class HealthResponse(BaseModel):
    status: str
    service: str


class BusinessRequest(BaseModel):
    company_name: str
    legal_form: str = "SHPK"
    address: Optional[str] = None
    capital: Optional[float] = None
    employees: Optional[int] = None


@app.get("/", operation_id="root")
def root():
    return {
        "service": "Legal Master API",
        "status": "running"
    }


@app.get(
    "/health",
    response_model=HealthResponse,
    operation_id="health"
)
def health():
    return {
        "status": "ok",
        "service": "legal-master-api"
    }


@app.post(
    "/business/validate",
    operation_id="validate_business"
)
def validate_business(data: BusinessRequest):
    issues = []

    if not data.company_name.strip():
        issues.append("Mungon emri i biznesit")

    if data.capital is not None and data.capital < 0:
        issues.append("Kapitali nuk mund të jetë negativ")

    if data.employees is not None and data.employees < 0:
        issues.append("Numri i punëtorëve nuk mund të jetë negativ")

    return {
        "valid": len(issues) == 0,
        "issues": issues,
        "data": data
    }
