"""
easemybot — Tally Ledger Reconciliation Engine
Main FastAPI application
"""

import os
import hmac
import hashlib
import logging
import json
import secrets
import threading
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

import asyncio
import razorpay

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from parser_tally import parse_tally_xml
from engine import run_reconciliation
from excel_generator import generate_excel

# ── Audit logging ────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%SZ",
)
audit = logging.getLogger("easemyreco.audit")

app = FastAPI(title="easemybot Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://www.easemyreco.com", "https://easemyreco.com"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# In-memory session store
# ─────────────────────────────────────────────
sessions: dict = {}
sessions_lock = threading.Lock()
TTL_SECONDS   = 3600  # 60 minutes — matches Privacy Policy promise
MAX_SESSIONS  = 50    # RAM guard: 50 sessions × ~6MB = ~300MB, safe on Railway Hobby (512MB)

# ─────────────────────────────────────────────
# Razorpay client — keys come from Railway env vars, never hardcoded
# ─────────────────────────────────────────────
RAZORPAY_KEY_ID         = os.getenv("RAZORPAY_KEY_ID", "")
RAZORPAY_KEY_SECRET     = os.getenv("RAZORPAY_KEY_SECRET", "")
RAZORPAY_WEBHOOK_SECRET = os.getenv("RAZORPAY_WEBHOOK_SECRET", "").strip()

# Client is initialised once at startup — reused for every request
rzp_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET))


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

    # RAM guard — reject if server is already holding too many sessions
    # Prevents OOM crash on Railway Hobby when traffic spikes
    with sessions_lock:
        if len(sessions) >= MAX_SESSIONS:
            return JSONResponse({
                "status":  "error",
                "error_code": "SERVER_BUSY",
                "message": "Server is busy. Please try again in a few minutes.",
            }, status_code=503)

    try:
        # asyncio.to_thread moves CPU-heavy parsing off the event loop
        # so Razorpay webhooks and health checks aren't blocked while files parse
        parsed_a = await asyncio.to_thread(parse_tally_xml, bytes_a, "A")
        parsed_b = await asyncio.to_thread(parse_tally_xml, bytes_b, "B")
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
            "paid":         False,   # set to True only after Razorpay signature verification
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

    # asyncio.to_thread keeps the event loop free during CPU-heavy reconciliation
    results = await asyncio.to_thread(
        run_reconciliation,
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
            "amount":           0.0 if total_mismatches == 0 else 20.0,
            "currency":         "INR",
        },
    }


# ─────────────────────────────────────────────
# /create-order  — frontend calls this when user clicks "Pay ₹20"
# We create a Razorpay Order server-side first (required by Razorpay)
# The order_id is then passed to the frontend checkout widget
# ─────────────────────────────────────────────
@app.post("/create-order")
async def create_order(request: Request):
    body = await request.json()
    session_token = body.get("session_token")

    with sessions_lock:
        session = sessions.get(session_token)

    if not session:
        raise HTTPException(404, "Session not found or expired.")

    if time.time() - session["created_at"] > TTL_SECONDS:
        raise HTTPException(410, "Session expired. Please re-upload your files.")

    # Guard: if already paid (e.g. user double-clicked), don't create a duplicate order
    if session.get("paid"):
        return {"status": "already_paid"}

    if not RAZORPAY_KEY_ID or not RAZORPAY_KEY_SECRET:
        audit.error("RAZORPAY_KEYS_MISSING — set RAZORPAY_KEY_ID and RAZORPAY_KEY_SECRET in Railway env vars")
        raise HTTPException(500, "Payment system not configured. Contact support.")

    try:
        # Amount in paise — ₹20 = 2000 paise
        order = rzp_client.order.create({
            "amount":   2000,
            "currency": "INR",
            # receipt ties this order back to our session for webhook lookup
            # max 40 chars — take first 40 of the 32-char hex token (safe)
            "receipt":  session_token[:40],
            "notes": {
                # storing session_token in notes so the webhook can look up the session
                # even if order_id matching fails
                "session_token": session_token,
                "party_a":       session.get("party_a_name", ""),
                "party_b":       session.get("party_b_name", ""),
            },
        })
    except Exception as e:
        audit.error("RAZORPAY_ORDER_CREATE_FAILED | session=%s | error=%s", session_token, str(e))
        raise HTTPException(500, "Could not create payment order. Please try again.")

    audit.info("ORDER_CREATED | razorpay_order_id=%s | session=%s", order["id"], session_token)

    return {
        "status":   "ok",
        "order_id": order["id"],
        "amount":   2000,
        "currency": "INR",
        # key_id (NOT key_secret) is safe to send to frontend — it's the public identifier
        "key_id":   RAZORPAY_KEY_ID,
    }


# ─────────────────────────────────────────────
# /verify-payment  — frontend calls this after Razorpay checkout succeeds
# We verify Razorpay's signature to confirm the payment is genuine
# Only after this check does the session get marked paid
# ─────────────────────────────────────────────
class VerifyPaymentRequest(BaseModel):
    session_token:       str
    razorpay_order_id:   str
    razorpay_payment_id: str
    razorpay_signature:  str


@app.post("/verify-payment")
async def verify_payment(req: VerifyPaymentRequest):
    with sessions_lock:
        session = sessions.get(req.session_token)

    if not session:
        raise HTTPException(404, "Session not found or expired.")

    if time.time() - session["created_at"] > TTL_SECONDS:
        raise HTTPException(410, "Session expired. Contact support if amount was deducted.")

    # Razorpay signature = HMAC-SHA256(order_id + "|" + payment_id, key_secret)
    # This proves the payment response came from Razorpay, not a tampered client request
    try:
        rzp_client.utility.verify_payment_signature({
            "razorpay_order_id":   req.razorpay_order_id,
            "razorpay_payment_id": req.razorpay_payment_id,
            "razorpay_signature":  req.razorpay_signature,
        })
    except razorpay.errors.SignatureVerificationError:
        audit.warning(
            "SIGNATURE_MISMATCH | session=%s | payment_id=%s",
            req.session_token, req.razorpay_payment_id,
        )
        raise HTTPException(400, "Payment verification failed. Contact support if amount was deducted.")

    # Signature valid — safe to unlock download
    with sessions_lock:
        sessions[req.session_token]["paid"]                = True
        sessions[req.session_token]["razorpay_payment_id"] = req.razorpay_payment_id
        sessions[req.session_token]["razorpay_order_id"]   = req.razorpay_order_id

    audit.info(
        "PAYMENT_VERIFIED | razorpay_payment_id=%s | session=%s",
        req.razorpay_payment_id, req.session_token,
    )

    return {"status": "paid"}


# ─────────────────────────────────────────────
# /confirm-payment  — Razorpay webhook (safety net)
# Razorpay calls this directly on payment.captured events
# Acts as backup if the frontend /verify-payment call dropped
# ─────────────────────────────────────────────
@app.post("/confirm-payment")
async def confirm_payment(request: Request):
    body         = await request.body()
    received_sig = request.headers.get("x-razorpay-signature", "")

    # Hard-fail if webhook secret is not configured — never silently skip auth
    if not RAZORPAY_WEBHOOK_SECRET:
        audit.error("RAZORPAY_WEBHOOK_SECRET not set — rejecting webhook to prevent auth bypass")
        raise HTTPException(500, "Webhook secret not configured.")

    # Verify the request is genuinely from Razorpay
    expected = hmac.new(
        RAZORPAY_WEBHOOK_SECRET.encode(),
        body,
        hashlib.sha256,
    ).hexdigest()
    if not hmac.compare_digest(expected, received_sig):
        audit.warning("WEBHOOK_SIGNATURE_MISMATCH — possible spoofed request")
        raise HTTPException(400, "Invalid webhook signature.")

    payload    = json.loads(body)
    event      = payload.get("event", "")
    entity     = payload.get("payload", {}).get("payment", {}).get("entity", {})
    notes      = entity.get("notes", {})
    payment_id = entity.get("id", "unknown")

    # Only act on captured/authorized events — ignore refunds, failures, etc.
    if event not in ("payment.captured", "payment.authorized"):
        return {"status": "ignored", "event": event}

    # Primary lookup: session_token stored in order notes at creation time
    session_token = notes.get("session_token", "")

    if not session_token:
        audit.warning("WEBHOOK_NO_SESSION_TOKEN | payment_id=%s", payment_id)
        raise HTTPException(400, "Could not identify session from webhook payload.")

    with sessions_lock:
        if session_token in sessions:
            sessions[session_token]["paid"]                = True
            sessions[session_token]["razorpay_payment_id"] = payment_id

    audit.info("WEBHOOK_PAYMENT_OK | razorpay_payment_id=%s | session=%s", payment_id, session_token)
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

    # asyncio.to_thread keeps event loop free during Excel generation (CPU + openpyxl)
    excel_bytes = await asyncio.to_thread(
        generate_excel,
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

    audit.info(
        "DOWNLOAD_SUCCESS | razorpay_payment_id=%s | session=%s | file=%s",
        session.get("razorpay_payment_id", "zero_mismatch_free"),
        token,
        filename,
    )

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

@app.get("/privacy-policy")
def serve_privacy_policy():
    path = Path(__file__).parent / "privacy-policy.html"
    if not path.exists():
        raise HTTPException(404, "Page not found.")
    return FileResponse(str(path), media_type="text/html")

@app.get("/terms")
def serve_terms():
    path = Path(__file__).parent / "terms.html"
    if not path.exists():
        raise HTTPException(404, "Page not found.")
    return FileResponse(str(path), media_type="text/html")

@app.get("/refund-policy")
def serve_refund_policy():
    path = Path(__file__).parent / "refund-policy.html"
    if not path.exists():
        raise HTTPException(404, "Page not found.")
    return FileResponse(str(path), media_type="text/html")


# ─────────────────────────────────────────────
# Health
# ─────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}
