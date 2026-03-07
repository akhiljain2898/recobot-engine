"""
RecoBot — Tally Ledger Reconciliation Engine
Main FastAPI application
"""

import os
import hmac
import hashlib
import json
import secrets
import threading
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from parser_tally import parse_tally_xml
from engine import run_reconciliation
from excel_generator import generate_excel

app = FastAPI(title="RecoBot Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# In-memory session store
# ─────────────────────────────────────────────
sessions: dict = {}
sessions_lock = threading.Lock()
TTL_SECONDS = 3600  # 60 minutes


def _session_gc():
    while True:
        time.sleep(60)
        now = time.time()
        with sessions_lock:
            expired = [t for t, s in list(sessions.items()) if now - s["created_at"] > TTL_SECONDS]
            for t in expired:
                del sessions[t]


threading.Thread(target=_session_gc, daemon=True).start()
# ─────────────────────────────────────────────
# File validation
# ─────────────────────────────────────────────
MAX_FILE_SIZE_MB = 8
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024


def _validate_upload(file: UploadFile, file_bytes: bytes, label: str):
    """
    Validate file before passing to parser.
    Returns a JSONResponse error if invalid, None if OK.
    Checks: empty file, size > 8 MB, wrong extension.
    """
    filename = file.filename or ""

    # Check 1 — empty file
    if len(file_bytes) == 0:
        return JSONResponse({
            "status": "error",
            "error_code": "EMPTY_FILE",
            "message": (
                f"File {label} appears to be empty. "
                "Please re-export your Tally XML and try again."
            ),
        })

    # Check 2 — file too large
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        size_mb = round(len(file_bytes) / (1024 * 1024), 1)
        return JSONResponse({
            "status": "error",
            "error_code": "FILE_TOO_LARGE",
            "message": (
                f"File {label} is {size_mb} MB which exceeds the 8 MB limit. "
                "Please export a shorter date range or a specific ledger "
                "instead of the full Day Book."
            ),
        })

    # Check 3 — wrong file extension
    if not filename.lower().endswith(".xml"):
        ext = Path(filename).suffix or "unknown"
        return JSONResponse({
            "status": "error",
            "error_code": "WRONG_FILE_TYPE",
            "message": (
                f"File {label} is a {ext} file. "
                "Only Tally XML exports (.xml) are accepted. "
                "Visit the Instructions tab for a step-by-step guide on "
                "how to export the correct file from Tally."
            ),
        })

    return None




# ─────────────────────────────────────────────
# /analyse
# ─────────────────────────────────────────────
@app.post("/analyse")
async def analyse(
    party_a_name: str = Form(...),
    party_a_role: str = Form(...),
    party_b_name: str = Form(...),
    party_b_role: str = Form(...),
    file_a: UploadFile = File(...),
    file_b: UploadFile = File(...),
):
    try:
        bytes_a = await file_a.read()
        bytes_b = await file_b.read()
    except Exception:
        raise HTTPException(400, "Could not read uploaded files.")

    # Validate both files before any processing
    err = _validate_upload(file_a, bytes_a, "A")
    if err:
        return err
    err = _validate_upload(file_b, bytes_b, "B")
    if err:
        return err

    try:
        parsed_a = parse_tally_xml(bytes_a, label="A")
        parsed_b = parse_tally_xml(bytes_b, label="B")
    except ValueError as e:
        return JSONResponse({
            "status": "error",
            "error_code": "INVALID_XML",
            "message": str(e),
        })

    def date_range(entries):
        dates = [e["date"] for e in entries if e.get("date")]
        return (min(dates), max(dates)) if dates else (None, None)

    fa_from, fa_to = date_range(parsed_a["entries"])
    fb_from, fb_to = date_range(parsed_b["entries"])

    if fa_from is None or fb_from is None:
        return JSONResponse({
            "status": "error",
            "error_code": "NO_DATES",
            "message": "Could not determine date range from one or both files.",
        })

    overlap_from = max(fa_from, fb_from)
    overlap_to   = min(fa_to,   fb_to)

    if overlap_from > overlap_to:
        return JSONResponse({
            "status": "error",
            "error_code": "NO_OVERLAP",
            "message": "Files do not cover any common period. Please check your exports.",
        })

    session_token = secrets.token_hex(16)

    with sessions_lock:
        sessions[session_token] = {
            "created_at":   time.time(),
            "party_a_name": party_a_name,
            "party_a_role": party_a_role,
            "party_b_name": party_b_name,
            "party_b_role": party_b_role,
            "parsed_a":     parsed_a,
            "parsed_b":     parsed_b,
            "period": {
                "file_a":   {"from": fa_from,       "to": fa_to},
                "file_b":   {"from": fb_from,       "to": fb_to},
                "overlap":  {"from": overlap_from,  "to": overlap_to},
            },
            "period_used":  {"from": overlap_from, "to": overlap_to},
            "results":      None,
            "paid":         True,
            "downloaded":   False,
        }

    return {
        "session_token": session_token,
        "status": "success",
        "period": {
            "file_a":  {"from": fa_from,      "to": fa_to},
            "file_b":  {"from": fb_from,      "to": fb_to},
            "overlap": {"from": overlap_from, "to": overlap_to},
        },
        "voucher_types":    parsed_a["voucher_types"] + parsed_b["voucher_types"],
        "total_entries_a":  len(parsed_a["entries"]),
        "total_entries_b":  len(parsed_b["entries"]),
    }


# ─────────────────────────────────────────────
# /reconcile
# ─────────────────────────────────────────────
class Classification(BaseModel):
    name: str
    classification: str


class ReconcileRequest(BaseModel):
    session_token: str
    period_override: Optional[dict] = None
    classifications: list[Classification]


@app.post("/reconcile")
async def reconcile(req: ReconcileRequest):
    with sessions_lock:
        session = sessions.get(req.session_token)

    if not session:
        raise HTTPException(404, "Session not found or expired.")

    if time.time() - session["created_at"] > TTL_SECONDS:
        raise HTTPException(410, "Session expired. Please re-upload your files.")

    classifications_map = {c.name: c.classification for c in req.classifications}

    period = session["period"]["overlap"]
    if req.period_override:
        period = req.period_override

    results = run_reconciliation(
        parsed_a=session["parsed_a"],
        parsed_b=session["parsed_b"],
        classifications=classifications_map,
        period_from=period["from"],
        period_to=period["to"],
    )

    with sessions_lock:
        sessions[req.session_token]["results"]     = results
        sessions[req.session_token]["period_used"] = period

    inv = results["invoices"]
    pay = results["payments"]
    total_mismatches = (
        inv["a_not_in_b_count"] + inv["b_not_in_a_count"] +
        pay["a_not_in_b_count"] + pay["b_not_in_a_count"]
    )

    # Only send counts/values to frontend — strip the large detail lists
    inv_preview = {k: v for k, v in inv.items() if not k.startswith("_")}
    pay_preview = {k: v for k, v in pay.items() if not k.startswith("_")}

    return {
        "session_token": req.session_token,
        "status": "success",
        "results_preview": {
            "period_used":      period,
            "invoices":         inv_preview,
            "payments":         pay_preview,
            "zero_mismatches":  total_mismatches == 0,
            "payment_required": total_mismatches > 0,
            "amount":           0.0 if total_mismatches == 0 else 12.0,
            "currency":         "INR",
        },
    }


# ─────────────────────────────────────────────
# /confirm-payment  (Razorpay webhook)
# ─────────────────────────────────────────────
RAZORPAY_WEBHOOK_SECRET = os.getenv("RAZORPAY_WEBHOOK_SECRET", "")


@app.post("/confirm-payment")
async def confirm_payment(request: Request):
    body = await request.body()
    received_sig = request.headers.get("x-razorpay-signature", "")

    if RAZORPAY_WEBHOOK_SECRET:
        expected = hmac.new(
            RAZORPAY_WEBHOOK_SECRET.encode(),
            body,
            hashlib.sha256,
        ).hexdigest()
        if not hmac.compare_digest(expected, received_sig):
            raise HTTPException(400, "Invalid webhook signature.")

    payload = json.loads(body)
    order_id = (
        payload
        .get("payload", {})
        .get("payment", {})
        .get("entity", {})
        .get("order_id")
    )

    if not order_id:
        raise HTTPException(400, "Missing order_id in payload.")

    with sessions_lock:
        if order_id in sessions:
            sessions[order_id]["paid"] = True

    return {"status": "ok"}


# ─────────────────────────────────────────────
# /download
# ─────────────────────────────────────────────
@app.get("/download")
async def download(token: str):
    with sessions_lock:
        session = sessions.get(token)

    if not session:
        raise HTTPException(404, "Session not found or expired.")

    if time.time() - session["created_at"] > TTL_SECONDS:
        raise HTTPException(410, "Session expired. Please re-upload and pay again.")

    if session["results"] is None:
        raise HTTPException(400, "Reconciliation not yet run for this session.")

    # Zero-mismatch sessions don't need payment
    r   = session["results"]
    inv = r["invoices"]
    pay = r["payments"]
    zero = (
        inv["a_not_in_b_count"] + inv["b_not_in_a_count"] +
        pay["a_not_in_b_count"] + pay["b_not_in_a_count"]
    ) == 0

    if not zero and not session["paid"]:
        raise HTTPException(402, "Payment required before download.")

    excel_bytes = generate_excel(
        results=session["results"],
        party_a_name=session["party_a_name"],
        party_b_name=session["party_b_name"],
        period=session["period_used"],
    )

    with sessions_lock:
        sessions[token]["downloaded"] = True

    date_str  = datetime.now().strftime("%Y%m%d")
    filename  = (
        f"{session['party_a_name']}_vs_{session['party_b_name']}"
        f"_Reconciliation_{date_str}.xlsx"
    ).replace(" ", "_")

    return StreamingResponse(
        BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ─────────────────────────────────────────────
# Serve frontend HTML
# ─────────────────────────────────────────────
@app.get("/")
def serve_frontend():
    html_path = Path(__file__).parent / "index.html"
    if not html_path.exists():
        raise HTTPException(404, "Frontend not found.")
    return FileResponse(str(html_path), media_type="text/html")


# ─────────────────────────────────────────────
# Health
# ─────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}
