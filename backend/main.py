import os
import io
import json
import httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient

load_dotenv()

app = FastAPI(title="Contract Extractor API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Config ──────────────────────────────────────────────────────────────────
AZURE_STORAGE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
BLOB_CONTAINER                  = os.getenv("BLOB_CONTAINER", "contracts")

AZURE_OPENAI_ENDPOINT   = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_KEY        = os.getenv("AZURE_OPENAI_KEY")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4")

# ─── Blob Storage helpers ─────────────────────────────────────────────────────

def get_blob_client():
    return BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)

def list_blobs() -> list[dict]:
    client = get_blob_client()
    container = client.get_container_client(BLOB_CONTAINER)
    files = []
    for blob in container.list_blobs():
        name = blob.name.lower()
        if name.endswith(".pdf") or name.endswith(".docx") or name.endswith(".txt"):
            files.append({
                "id": blob.name,          # use blob name as ID
                "name": blob.name.split("/")[-1],
                "path": blob.name,
                "size": blob.size,
            })
    return files

def download_blob_text(blob_name: str) -> str:
    client = get_blob_client()
    blob = client.get_blob_client(container=BLOB_CONTAINER, blob=blob_name)
    content = blob.download_blob().readall()

    name_lower = blob_name.lower()

    if name_lower.endswith(".txt"):
        return content.decode("utf-8", errors="ignore")

    elif name_lower.endswith(".docx"):
        try:
            import docx
            doc = docx.Document(io.BytesIO(content))
            return "\n".join([p.text for p in doc.paragraphs])
        except Exception as e:
            return f"[Could not parse DOCX: {e}]"

    elif name_lower.endswith(".pdf"):
        try:
            import pdfplumber
            text_parts = []
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages[:10]:
                    text_parts.append(page.extract_text() or "")
            return "\n".join(text_parts)
        except Exception as e:
            return f"[Could not parse PDF: {e}]"

    return "[Unsupported file type]"

# ─── Azure OpenAI extraction ──────────────────────────────────────────────────

EXTRACTION_PROMPT = """You are a contract analysis assistant. Extract the following information from the contract text provided.

Return ONLY a valid JSON object with these exact keys:
{
  "vendor_name": "the name of the vendor/supplier/contractor",
  "contract_value": "the total contract value or cost as a string (include currency symbol)",
  "contract_value_numeric": 0.0,
  "currency": "USD/CAD/EUR/etc",
  "notes": "any important caveats (e.g. recurring, per-year, not found)"
}

If a field cannot be found, use null. For contract_value_numeric, extract the numeric value only (no symbols).
Do not include any explanation outside the JSON.

Contract text:
"""

async def extract_contract_info(text: str) -> dict:
    if not text.strip():
        return {"vendor_name": None, "contract_value": None, "contract_value_numeric": None, "currency": None, "notes": "Empty document"}

    truncated = text[:12000]
    url = f"{AZURE_OPENAI_ENDPOINT.rstrip('/')}/openai/deployments/{AZURE_OPENAI_DEPLOYMENT}/chat/completions?api-version=2024-02-01"
    payload = {
        "messages": [
            {"role": "system", "content": "You extract structured information from contracts. Always respond with valid JSON only."},
            {"role": "user", "content": EXTRACTION_PROMPT + truncated}
        ],
        "temperature": 0,
        "max_tokens": 500,
    }

    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            url,
            headers={"api-key": AZURE_OPENAI_KEY, "Content-Type": "application/json"},
            json=payload,
        )
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"].strip()

    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip("` \n")

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"vendor_name": None, "contract_value": None, "contract_value_numeric": None, "currency": None, "notes": f"Parse error: {raw[:100]}"}

# ─── Routes ───────────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.get("/contracts/list")
async def list_contracts():
    try:
        files = list_blobs()
        return {"files": files, "total": len(files)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

class ExtractRequest(BaseModel):
    file_ids: list[str]

@app.post("/contracts/extract")
async def extract_contracts(req: ExtractRequest):
    results = []
    for blob_name in req.file_ids:
        try:
            text = download_blob_text(blob_name)
            extracted = await extract_contract_info(text)
            results.append({
                "file_id": blob_name,
                "file_name": blob_name.split("/")[-1],
                "path": blob_name,
                **extracted,
            })
        except Exception as e:
            results.append({
                "file_id": blob_name,
                "file_name": blob_name.split("/")[-1],
                "error": str(e),
            })
    return {"results": results}

@app.post("/contracts/export-excel")
async def export_excel(req: ExtractRequest):
    extraction = await extract_contracts(req)
    results = extraction["results"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contract Analysis"

    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", start_color="1B3A5C")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="C0CCDA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill     = PatternFill("solid", start_color="EEF3F8")
    currency_font = Font(name="Arial", size=10, bold=True, color="1B5E20")
    error_fill   = PatternFill("solid", start_color="FFEBEE")
    error_font   = Font(name="Arial", size=10, color="C62828")

    ws.merge_cells("A1:G1")
    ws["A1"].value = "Contract Vendor & Cost Analysis"
    ws["A1"].font  = Font(name="Arial", bold=True, size=14, color="1B3A5C")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:G2")
    ws["A2"].value = f"Generated via Contract Extractor  •  {len(results)} contracts processed"
    ws["A2"].font  = Font(name="Arial", size=9, color="666666", italic=True)
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    headers    = ["#", "File Name", "Vendor Name", "Contract Value", "Currency", "Notes", "Blob Path"]
    col_widths = [5, 35, 30, 18, 10, 35, 40]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[4].height = 28

    for i, row in enumerate(results, 1):
        excel_row = i + 4
        is_alt    = i % 2 == 0
        has_error = "error" in row

        row_data = [
            i,
            row.get("file_name", ""),
            row.get("vendor_name") or ("ERROR" if has_error else "Not found"),
            row.get("contract_value") or ("—" if not has_error else row.get("error", "Error")),
            row.get("currency") or "—",
            row.get("notes") or "—",
            row.get("path", ""),
        ]

        for col, value in enumerate(row_data, 1):
            cell = ws.cell(row=excel_row, column=col, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(col in [2, 6, 7]))

            if has_error:
                cell.fill = error_fill
                cell.font = error_font
            else:
                if is_alt:
                    cell.fill = alt_fill
                cell.font = currency_font if col == 4 else Font(name="Arial", size=10)

        ws.row_dimensions[excel_row].height = 22

    summary_row = len(results) + 6
    ws.merge_cells(f"A{summary_row}:C{summary_row}")
    ws[f"A{summary_row}"].value = f"Total contracts analyzed: {len(results)}"
    ws[f"A{summary_row}"].font  = Font(name="Arial", bold=True, size=10, color="1B3A5C")
    ws[f"A{summary_row}"].alignment = Alignment(horizontal="left")

    ws.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=contract_analysis.xlsx"},
    )
