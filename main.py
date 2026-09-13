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
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError


app = FastAPI(
    title="Legal Master API",
    version="1.4.0",
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

ARBK_SEARCH_URL = "https://arbk.rks-gov.net/TableSearch"


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


def clean_text(value: Optional[str]) -> Optional[str]:
    if not value:
        return None

    value = re.sub(r"\s+", " ", value).strip()

    if not value:
        return None

    return value


def find_first(patterns, text):
    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE | re.MULTILINE
        )
        if match:
            return clean_text(match.group(1))

    return None


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
# ARBK BROWSER HELPERS
# =========================================================

async def fill_arbk_nui(page, nui: str):
    """
    Provon disa mënyra për të identifikuar fushën NUI.
    Kjo e bën kodin më rezistent ndaj ndryshimeve të vogla në HTML.
    """

    selectors = [
        'input[placeholder*="Numri unik"]',
        'input[placeholder*="numri i biznesit"]',
        'input[name*="nui" i]',
        'input[id*="nui" i]',
        'input[name*="business" i]',
        'input[id*="business" i]'
    ]

    for selector in selectors:
        locator = page.locator(selector)

        if await locator.count() > 0:
            try:
                await locator.first.fill(nui)
                return True
            except Exception:
                pass

    # Fallback: kërko input-et e dukshme.
    inputs = page.locator("input")

    count = await inputs.count()

    for i in range(count):
        item = inputs.nth(i)

        try:
            if not await item.is_visible():
                continue

            placeholder = (
                await item.get_attribute("placeholder") or ""
            ).lower()

            name = (
                await item.get_attribute("name") or ""
            ).lower()

            item_id = (
                await item.get_attribute("id") or ""
            ).lower()

            combined = f"{placeholder} {name} {item_id}"

            if (
                "unik" in combined
                or "nui" in combined
                or "biznes" in combined
            ):
                await item.fill(nui)
                return True

        except Exception:
            continue

    return False


async def click_arbk_search(page):
    search_candidates = [
        page.get_by_role(
            "button",
            name=re.compile("KËRKO|KERKO", re.IGNORECASE)
        ),
        page.locator(
            'button:has-text("KËRKO")'
        ),
        page.locator(
            'button:has-text("KERKO")'
        ),
        page.locator(
            'input[type="submit"]'
        )
    ]

    for locator in search_candidates:
        try:
            if await locator.count() > 0:
                await locator.first.click()
                return True
        except Exception:
            continue

    return False


def extract_public_arbk_data(
    text: str,
    nui: str
):
    """
    Nxjerr vetëm atë që gjendet në tekstin publik të faqes.
    Çdo fushë që nuk identifikohet mbetet None.
    """

    normalized = re.sub(r"[ \t]+", " ", text)

    # Status
    business_status = find_first(
        [
            r"\bStatusi\b\s*:?\s*(Aktiv|Pasiv)",
            r"\b(Aktiv|Pasiv)\b"
        ],
        normalized
    )

    # Forma juridike
    legal_form = find_first(
        [
            (
                r"(Shoqëri\s+me\s+përgjegjësi\s+"
                r"të\s+kufizuar)"
            ),
            r"\b(SH\.?P\.?K\.?)\b",
            r"(Shoqëri\s+Aksionare)",
            r"\b(SH\.?A\.?)\b",
            r"(Biznes\s+Individual)"
        ],
        normalized
    )

    # Email
    email_match = re.search(
        r"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b",
        text,
        re.IGNORECASE
    )

    email = (
        email_match.group(0)
        if email_match
        else None
    )

    # Telefon
    phone = find_first(
        [
            (
                r"(?:Telefon|Telefoni|Tel\.?)\s*:?\s*"
                r"(\+?\d[\d\s\-\/]{6,20})"
            )
        ],
        normalized
    )

    # Komuna
    municipality = find_first(
        [
            r"(?:Komuna)\s*:?\s*([A-Za-zÇËçë\s\-]+)",
        ],
        normalized
    )

    # Data e regjistrimit
    registration_date = find_first(
        [
            (
                r"(?:Themeluar|Regjistruar|"
                r"Data e regjistrimit)\s*:?\s*"
                r"(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{4})"
            )
        ],
        normalized
    )

    # Adresa
    address = find_first(
        [
            (
                r"(?:Adresa|Selia)\s*:?\s*"
                r"(.{3,120}?)(?:\n|Komuna|Telefon|Email|$)"
            )
        ],
        text
    )

    # Emri i biznesit:
    # Fillimisht kërkojmë rreshtat pranë NUI.
    business_name = None

    lines = [
        clean_text(line)
        for line in text.splitlines()
        if clean_text(line)
    ]

    for index, line in enumerate(lines):
        if nui in line:
            nearby = lines[
                max(0, index - 4):
                min(len(lines), index + 5)
            ]

            for candidate in nearby:
                if (
                    candidate
                    and candidate != nui
                    and len(candidate) <= 150
                    and not candidate.lower().startswith(
                        (
                            "rezultatet",
                            "pronari",
                            "përfaqësues",
                            "komuna",
                            "status"
                        )
                    )
                ):
                    if (
                        "sh.p.k" in candidate.lower()
                        or "shpk" in candidate.lower()
                        or "sh.a" in candidate.lower()
                        or "biznes" not in candidate.lower()
                    ):
                        business_name = candidate
                        break

        if business_name:
            break

    # Pronaret / perfaqesuesit - best effort.
    owners = []
    directors = []

    owner_section = re.search(
        (
            r"PRONAR[ËE]T?(.*?)(?:"
            r"P[ËE]RFAQ[ËE]SUES|BORDI|$)"
        ),
        text,
        re.IGNORECASE | re.DOTALL
    )

    if owner_section:
        owner_text = owner_section.group(1)

        owner_lines = [
            clean_text(x)
            for x in owner_text.splitlines()
            if clean_text(x)
        ]

        for item in owner_lines[:10]:
            if (
                len(item) >= 3
                and not re.fullmatch(r"\d+", item)
            ):
                owners.append(
                    {
                        "full_name": item,
                        "role": "Pronar"
                    }
                )

    representative_section = re.search(
        (
            r"P[ËE]RFAQ[ËE]SUES(?:IT)?(.*?)(?:"
            r"BORDI|PRONAR[ËE]T|$)"
        ),
        text,
        re.IGNORECASE | re.DOTALL
    )

    if representative_section:
        rep_text = representative_section.group(1)

        rep_lines = [
            clean_text(x)
            for x in rep_text.splitlines()
            if clean_text(x)
        ]

        for item in rep_lines[:10]:
            if (
                len(item) >= 3
                and not re.fullmatch(r"\d+", item)
            ):
                directors.append(
                    {
                        "full_name": item,
                        "role": "Përfaqësues"
                    }
                )

    return {
        "business_name": business_name,
        "trade_name": None,
        "nui": nui,
        "legal_form": legal_form,
        "business_status": business_status,
        "registration_date": registration_date,
        "address": address,
        "municipality": municipality,
        "primary_activity": None,
        "other_activities": [],
        "owners": owners,
        "directors": directors,
        "email": email,
        "phone": phone
    }


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
        "version": "1.4.0"
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
        issues.append(
            "Kapitali nuk mund të jetë negativ"
        )

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
# ARBK INDIVIDUAL PUBLIC LOOKUP
# =========================================================

@app.get(
    "/arbk/lookup/{nui}",
    operation_id="lookup_arbk_business",
    dependencies=[Depends(verify_api_key)]
)
async def lookup_arbk_business(nui: str):

    nui = re.sub(r"\D", "", nui)

    if not nui:
        raise HTTPException(
            status_code=400,
            detail="NUI nuk është valid"
        )

    async with async_playwright() as playwright:

        browser = None

        try:
            browser = await playwright.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage"
                ]
            )

            page = await browser.new_page(
                viewport={
                    "width": 1440,
                    "height": 1000
                }
            )

            await page.goto(
                ARBK_SEARCH_URL,
                wait_until="domcontentloaded",
                timeout=60000
            )

            await page.wait_for_timeout(2500)

            filled = await fill_arbk_nui(
                page,
                nui
            )

            if not filled:
                raise HTTPException(
                    status_code=502,
                    detail=(
                        "Nuk u identifikua fusha NUI "
                        "në faqen publike të ARBK-së"
                    )
                )

            searched = await click_arbk_search(page)

            if not searched:
                raise HTTPException(
                    status_code=502,
                    detail=(
                        "Nuk u identifikua butoni KËRKO "
                        "në faqen publike të ARBK-së"
                    )
                )

            try:
                await page.wait_for_load_state(
                    "networkidle",
                    timeout=20000
                )
            except PlaywrightTimeoutError:
                pass

            await page.wait_for_timeout(3000)

            body_text = await page.locator(
                "body"
            ).inner_text()

            if nui not in body_text:
                return {
                    "found": False,
                    "nui": nui,
                    "source": ARBK_SEARCH_URL,
                    "message": (
                        "Nuk u gjet subjekt publik "
                        "me këtë NUI në rezultatin e ARBK-së."
                    )
                }

            extracted = extract_public_arbk_data(
                body_text,
                nui
            )

            return {
                "found": True,
                "source": ARBK_SEARCH_URL,
                "data": extracted,
                "raw_public_text": body_text[:12000],
                "note": (
                    "Të dhënat janë lexuar nga rezultati "
                    "publik për këtë NUI. Fushat që nuk "
                    "u identifikuan publikisht mbeten null "
                    "ose lista boshe. Nuk janë shpikur të dhëna."
                )
            }

        except HTTPException:
            raise

        except PlaywrightTimeoutError:
            raise HTTPException(
                status_code=504,
                detail=(
                    "Faqja e ARBK-së nuk u përgjigj "
                    "brenda afatit të lejuar"
                )
            )

        except Exception as exc:
            raise HTTPException(
                status_code=502,
                detail=(
                    "Gabim gjatë leximit të faqes publike "
                    f"të ARBK-së: {str(exc)}"
                )
            )

        finally:
            if browser:
                await browser.close()


# =========================================================
# MANUAL ARBK DATA VERIFICATION
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

    return {
        "valid": len(critical_issues) == 0,
        "critical_issues": critical_issues,
        "warnings": warnings,
        "business": data.model_dump()
    }


# =========================================================
# DOCUMENT GENERATION
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
        base_filename = clean_filename(
            data.filename
        )
    else:
        base_filename = clean_filename(
            f"{data.document_type}_{uuid4().hex[:8]}"
        )

    filename = f"{base_filename}.docx"

    internal_name = (
        f"{uuid4().hex}_{filename}"
    )

    output_path = (
        BASE_DIR / internal_name
    )

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
# DOWNLOAD
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
            detail=(
                "Dokumenti nuk ekziston ose "
                "linku ka skaduar"
            )
        )

    file_path = Path(
        file_data["path"]
    )

    if not file_path.exists():
        raise HTTPException(
            status_code=404,
            detail=(
                "Dokumenti nuk gjendet më në server"
            )
        )

    return FileResponse(
        path=file_path,
        filename=file_data["filename"],
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )
