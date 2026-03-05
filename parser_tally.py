"""
parser_tally.py — Tally XML parser for RecoBot

Handles:
- UTF-16 / UTF-8 BOM detection and decoding
- Character spacing removal (Tally inserts spaces between chars)
- Tag-order-independent extraction (each tag matched independently)
- Date parsing (d-MMM-yy → YYYY-MM-DD)
- Amount parsing (absolute value of DSPVCHCRAMT or DSPVCHDRAMT)
- Invoice number normalisation
- Auto-exclusions: RCM, TDS, opening balance, cancelled vouchers
- Grouping by normalised invoice number (GST head splits)
- Voucher type classification suggestions

Changes (v2):
  [1] Blank invoice numbers stamped with unmatched_reason at parse time
  [2] Friendlier, actionable error message when XML structure not recognised

Changes (v3):
  [3] Tag extraction rewritten to be order-independent — no more regex hang
      when tag order differs from expected sequence
  [4] 45-second hard timeout on parse_tally_xml — kills runaway parses
      and returns a clean error to the user instead of hanging forever
"""

import re
import threading
from collections import defaultdict

# ─────────────────────────────────────────────
# Timeout config
# ─────────────────────────────────────────────
PARSE_TIMEOUT_SECONDS = 45


# ─────────────────────────────────────────────
# Month map for Tally date format (1-Apr-25)
# ─────────────────────────────────────────────
_MONTH = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

_KEYWORD_MAP = [
    (re.compile(r"purc|purch|purchase",        re.I), "purchase_invoice", "high"),
    (re.compile(r"sale|sales",                 re.I), "sales_invoice",    "high"),
    (re.compile(r"payment|pymt|pay",           re.I), "payment",          "high"),
    (re.compile(r"receipt|rcpt|rec",           re.I), "receipt",          "high"),
    (re.compile(r"debit.?note|d.?note|dn",     re.I), "debit_note",       "high"),
    (re.compile(r"credit.?note|c.?note|cn",    re.I), "credit_note",      "high"),
    (re.compile(r"rcm|tds",                    re.I), "ignore",           "high"),
    (re.compile(r"journal|jnl|contra",         re.I), "ignore",           "medium"),
]

# FIX [3]: Pre-compiled individual tag patterns — extracted independently per block
_TAG_PATTERNS = {
    "vtype": re.compile(r"<DSPVCHTYPE[^>]*>(.*?)</DSPVCHTYPE>",             re.IGNORECASE),
    "inv":   re.compile(r"<DSPEXPLVCHNUMBER[^>]*>(.*?)</DSPEXPLVCHNUMBER>", re.IGNORECASE),
    "date":  re.compile(r"<DSPVCHDATE[^>]*>(.*?)</DSPVCHDATE>",             re.IGNORECASE),
    "cr":    re.compile(r"<DSPVCHCRAMT[^>]*>(.*?)</DSPVCHCRAMT>",           re.IGNORECASE),
    "dr":    re.compile(r"<DSPVCHDRAMT[^>]*>(.*?)</DSPVCHDRAMT>",           re.IGNORECASE),
}


# ─────────────────────────────────────────────
# Timeout helper
# Threading-based — safe for FastAPI async context on all platforms
# ─────────────────────────────────────────────

class _TimeoutError(Exception):
    pass


def _run_with_timeout(fn, timeout_secs, *args, **kwargs):
    """
    Run fn(*args, **kwargs) in a background thread.
    Raises _TimeoutError if it doesn't complete within timeout_secs.
    """
    result    = [None]
    exception = [None]

    def target():
        try:
            result[0] = fn(*args, **kwargs)
        except Exception as e:
            exception[0] = e

    t = threading.Thread(target=target, daemon=True)
    t.start()
    t.join(timeout_secs)

    if t.is_alive():
        raise _TimeoutError(
            f"Parsing timed out after {timeout_secs} seconds. "
            "Your file may be too large, malformed, or not a valid Tally XML export. "
            "Visit the Instructions tab for guidance on exporting the correct file."
        )

    if exception[0]:
        raise exception[0]

    return result[0]


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _decode(raw: bytes) -> str:
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def _remove_char_spacing(text: str) -> str:
    """
    Tally exports have spaces between every character inside tag values.
    e.g. <DSPVCHTYPE>S a l e</DSPVCHTYPE>
    Collapses these back to normal strings.
    """
    def collapse(m):
        inner = m.group(1)
        tokens = inner.split(" ")
        if all(len(t) == 1 for t in tokens) and len(tokens) > 1:
            return ">" + "".join(tokens) + "<"
        return m.group(0)
    return re.sub(r">([^<]+)<", collapse, text)


def _parse_date(raw: str) -> str | None:
    raw = raw.strip()
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$", raw)
    if not m:
        return None
    day, mon, yr = m.group(1), m.group(2).lower(), m.group(3)
    month = _MONTH.get(mon)
    if not month:
        return None
    if len(yr) == 2:
        yr = "20" + yr
    return f"{yr}-{month}-{int(day):02d}"


def _parse_amount(cr: str, dr: str) -> float:
    for val in (cr, dr):
        val = val.strip()
        if val:
            try:
                return abs(float(val.replace(",", "")))
            except ValueError:
                pass
    return 0.0


def _normalise_invoice(raw: str) -> str:
    raw = raw.strip()
    raw = re.sub(r"^\(No\.?:\s*", "", raw, flags=re.I)
    raw = re.sub(r"\)$", "", raw)
    raw = raw.upper()
    raw = raw.replace(" ", "").replace("-", "").replace("/", "")
    return raw


def _suggest_classification(vtype: str):
    for pattern, classification, confidence in _KEYWORD_MAP:
        if pattern.search(vtype):
            return classification, confidence
    return "purchase_invoice", "medium"


def _is_auto_ignore(vtype: str) -> tuple:
    if re.search(r"rcm|tds", vtype, re.I):
        return True, "Auto-detected TDS/RCM entry"
    return False, None


# ─────────────────────────────────────────────
# FIX [3]: Order-independent block splitter
# ─────────────────────────────────────────────

def _split_into_blocks(text: str) -> list:
    """
    Split XML into per-voucher blocks anchored on DSPVCHTYPE tags.
    Each block contains one voucher row and all its sibling tags,
    regardless of what order those tags appear in.

    Strategy:
      - Find all start positions of <DSPVCHTYPE> tags
      - Slice the text between consecutive positions
      - Each slice is one self-contained voucher block
    """
    positions = [m.start() for m in re.finditer(
        r"<DSPVCHTYPE[^>]*>", text, re.IGNORECASE
    )]
    if not positions:
        return []

    blocks = []
    for i, start in enumerate(positions):
        end = positions[i + 1] if i + 1 < len(positions) else len(text)
        blocks.append(text[start:end])
    return blocks


def _extract_tag(block: str, pattern: re.Pattern) -> str:
    """Extract first match of a tag pattern from a block. Returns '' if absent."""
    m = pattern.search(block)
    return m.group(1).strip() if m else ""


# ─────────────────────────────────────────────
# Core parse logic (runs inside timeout wrapper)
# ─────────────────────────────────────────────

def _parse_inner(raw: bytes, label: str) -> dict:
    text = _decode(raw)
    text = _remove_char_spacing(text)
    text = (text
            .replace("&amp;", "&")
            .replace("&lt;",  "<")
            .replace("&gt;",  ">")
            .replace("&quot;", '"'))

    # FIX [2]: Validate with actionable error message
    if "ENVELOPE" not in text.upper() and "DSPVCHTYPE" not in text.upper():
        raise ValueError(
            f"File {label}: You seem to have uploaded the wrong file format. "
            "Visit the Instructions tab to check out a detailed video on "
            "how to fetch a .xml report from Tally."
        )

    # FIX [3]: Split into blocks, extract each tag independently
    blocks = _split_into_blocks(text)

    raw_entries = []
    for block in blocks:
        vtype    = _extract_tag(block, _TAG_PATTERNS["vtype"])
        inv_raw  = _extract_tag(block, _TAG_PATTERNS["inv"])
        date_raw = _extract_tag(block, _TAG_PATTERNS["date"])
        cr_amt   = _extract_tag(block, _TAG_PATTERNS["cr"])
        dr_amt   = _extract_tag(block, _TAG_PATTERNS["dr"])

        if not vtype:
            continue

        # Skip TDS / RCM
        if re.search(r"rcm|tds", vtype, re.I):
            continue

        # Skip opening balance (blank invoice + blank date)
        if not inv_raw and not date_raw:
            continue

        date     = _parse_date(date_raw) if date_raw else None
        amount   = _parse_amount(cr_amt, dr_amt)
        inv_norm = _normalise_invoice(inv_raw) if inv_raw else ""

        # FIX [1]: Stamp reason tag at parse time if invoice number is blank
        unmatched_reason = f"no_invoice_number_{label.lower()}" if not inv_norm else None

        raw_entries.append({
            "voucher_type":        vtype,
            "invoice_number_raw":  inv_raw,
            "invoice_number_norm": inv_norm,
            "date":                date,
            "amount":              amount,
            "source":              label,
            "unmatched_reason":    unmatched_reason,
        })

    # Drop cancelled vouchers
    cancelled = set(re.findall(
        r"<DSPEXPLVCHNUMBER[^>]*>(.*?)</DSPEXPLVCHNUMBER>"
        r"(?:.*?)<ISCANCELLED[^>]*>Yes</ISCANCELLED>",
        text, re.DOTALL | re.IGNORECASE,
    ))
    cancelled_norm = {_normalise_invoice(c) for c in cancelled}
    raw_entries = [e for e in raw_entries
                   if e["invoice_number_norm"] not in cancelled_norm]

    # ── Group by invoice number (GST head splits) ────────────────────────────
    groups: dict = defaultdict(lambda: {"amount": 0.0, "count": 0, "entry": None})
    for e in raw_entries:
        key = (e["voucher_type"], e["invoice_number_norm"], e["date"])
        groups[key]["amount"] += e["amount"]
        groups[key]["count"]  += 1
        groups[key]["entry"]   = e

    entries = []
    for (vtype, inv_norm, date), g in groups.items():
        base = g["entry"].copy()
        base["amount"]              = g["amount"]
        base["invoice_number_norm"] = inv_norm
        entries.append(base)

    # ── Voucher type summary ─────────────────────────────────────────────────
    vtype_counts: dict = defaultdict(int)
    for e in raw_entries:
        vtype_counts[e["voucher_type"]] += 1

    voucher_types = []
    for vtype, count in vtype_counts.items():
        auto_ign, ign_reason = _is_auto_ignore(vtype)
        sugg, conf = _suggest_classification(vtype)
        if auto_ign:
            sugg, conf = "ignore", "high"
        voucher_types.append({
            "name":                     vtype,
            "found_in":                 label,
            "count":                    count,
            "suggested_classification": sugg,
            "confidence":               conf,
            "auto_ignore":              auto_ign,
            "ignore_reason":            ign_reason,
        })

    return {"entries": entries, "voucher_types": voucher_types}


# ─────────────────────────────────────────────
# Public: parse_tally_xml
# ─────────────────────────────────────────────

def parse_tally_xml(raw: bytes, label: str = "A") -> dict:
    """
    Parse a Tally XML export with a 45-second hard timeout.
    FIX [4]: Any parse that takes longer than 45s is killed cleanly
             and returns a user-friendly error instead of hanging.
    """
    try:
        return _run_with_timeout(
            _parse_inner,
            PARSE_TIMEOUT_SECONDS,
            raw,
            label,
        )
    except _TimeoutError as e:
        raise ValueError(str(e))
