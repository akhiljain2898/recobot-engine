"""
parser_tally.py — Tally XML parser for RecoBot

Handles:
- UTF-16 / UTF-8 BOM detection and decoding
- Character spacing removal (Tally inserts spaces between chars)
- 6-tag block extraction via regex
- Date parsing (d-MMM-yy → YYYY-MM-DD)
- Amount parsing (absolute value of DSPVCHCRAMT or DSPVCHDRAMT)
- Invoice number normalisation
- Auto-exclusions: RCM, TDS, opening balance, cancelled vouchers
- Grouping by normalised invoice number (GST head splits)
- Voucher type classification suggestions
"""

import re
from collections import defaultdict

# ─────────────────────────────────────────────
# Month map for Tally date format (1-Apr-25)
# ─────────────────────────────────────────────
_MONTH = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

# Voucher type keyword → suggested classification + confidence
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


def _decode(raw: bytes) -> str:
    """Detect BOM and decode bytes to string."""
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def _remove_char_spacing(text: str) -> str:
    """
    Tally exports have spaces between every character inside tag values.
    e.g.  <DSPVCHTYPE>P u r c</DSPVCHTYPE>
    This regex collapses those spaced-out strings inside XML tags.
    Strategy: replace 'X Y Z' patterns inside > ... < with 'XYZ'.
    """
    def collapse(m):
        inner = m.group(1)
        # If every token is 1 char separated by single spaces → collapse
        tokens = inner.split(" ")
        if all(len(t) == 1 for t in tokens) and len(tokens) > 1:
            return ">" + "".join(tokens) + "<"
        return m.group(0)

    return re.sub(r">([^<]+)<", collapse, text)


def _parse_date(raw: str) -> str | None:
    """Convert '1-Apr-25' or '01-Apr-2025' → '2025-04-01'."""
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
    """Return absolute value of whichever field is populated."""
    for val in (cr, dr):
        val = val.strip()
        if val:
            try:
                return abs(float(val.replace(",", "")))
            except ValueError:
                pass
    return 0.0


def _normalise_invoice(raw: str) -> str:
    """
    Strip (No.: prefix and ) suffix, then:
    uppercase → remove spaces → remove hyphens → remove forward slashes
    """
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
    return "purchase_invoice", "medium"  # fallback — ask user


def _is_auto_ignore(vtype: str) -> tuple:
    """Returns (bool, reason)."""
    if re.search(r"rcm|tds", vtype, re.I):
        return True, "Auto-detected TDS/RCM entry"
    return False, None


def parse_tally_xml(raw: bytes, label: str = "A") -> dict:
    """
    Parse a Tally XML export.

    Returns:
        {
          "entries":       [ { date, voucher_type, invoice_number_raw,
                               invoice_number_norm, amount, source } ],
          "voucher_types": [ { name, found_in, count, suggested_classification,
                               confidence, auto_ignore, ignore_reason } ]
        }
    """
    text = _decode(raw)
    text = _remove_char_spacing(text)

    # Decode XML entities
    text = (text
            .replace("&amp;", "&")
            .replace("&lt;",  "<")
            .replace("&gt;",  ">")
            .replace("&quot;", '"'))

    # Validate it looks like a Tally export
    if "ENVELOPE" not in text.upper() and "DSPVCHTYPE" not in text:
        raise ValueError(
            f"File {label} does not appear to be a valid Tally XML export. "
            "Could not find ENVELOPE root tag after cleaning."
        )

    # ── Extract 6-tag repeating blocks ──────────────────────────────────────
    # We look for the 6 confirmed Tally tags in sequence, flexible whitespace
    pattern = re.compile(
        r"<DSPVCHTYPE[^>]*>(.*?)</DSPVCHTYPE>"
        r".*?<DSPEXPLVCHNUMBER[^>]*>(.*?)</DSPEXPLVCHNUMBER>"
        r".*?<DSPVCHDATE[^>]*>(.*?)</DSPVCHDATE>"
        r".*?<DSPVCHLEDACCOUNT[^>]*>(.*?)</DSPVCHLEDACCOUNT>"
        r".*?<DSPVCHCRAMT[^>]*>(.*?)</DSPVCHCRAMT>"
        r".*?<DSPVCHDRAMT[^>]*>(.*?)</DSPVCHDRAMT>",
        re.DOTALL | re.IGNORECASE,
    )

    raw_entries = []
    for m in pattern.finditer(text):
        vtype    = m.group(1).strip()
        inv_raw  = m.group(2).strip()
        date_raw = m.group(3).strip()
        _ledger  = m.group(4).strip()
        cr_amt   = m.group(5).strip()
        dr_amt   = m.group(6).strip()

        # ── Auto-exclusions ─────────────────────────────────────────────────
        # Skip cancelled
        # (ISCANCELLED check done separately below if present)

        # Skip TDS / RCM
        if re.search(r"rcm|tds", vtype, re.I):
            continue

        # Skip opening balance (blank invoice + blank date)
        if not inv_raw and not date_raw:
            continue

        date   = _parse_date(date_raw) if date_raw else None
        amount = _parse_amount(cr_amt, dr_amt)
        inv_norm = _normalise_invoice(inv_raw) if inv_raw else ""

        raw_entries.append({
            "voucher_type":       vtype,
            "invoice_number_raw": inv_raw,
            "invoice_number_norm": inv_norm,
            "date":               date,
            "amount":             amount,
            "source":             label,
        })

    # Also catch ISCANCELLED=Yes and drop those entries
    # We re-scan the full text for cancelled voucher numbers and exclude them
    cancelled = set(re.findall(
        r"<DSPEXPLVCHNUMBER[^>]*>(.*?)</DSPEXPLVCHNUMBER>"
        r"(?:.*?)<ISCANCELLED[^>]*>Yes</ISCANCELLED>",
        text, re.DOTALL | re.IGNORECASE,
    ))
    cancelled_norm = {_normalise_invoice(c) for c in cancelled}
    raw_entries = [e for e in raw_entries if e["invoice_number_norm"] not in cancelled_norm]

    # ── Group by invoice number (GST head splits) ────────────────────────────
    # Key = (voucher_type, invoice_number_norm, date)
    groups: dict = defaultdict(lambda: {"amount": 0.0, "count": 0, "entry": None})
    for e in raw_entries:
        key = (e["voucher_type"], e["invoice_number_norm"], e["date"])
        groups[key]["amount"] += e["amount"]
        groups[key]["count"]  += 1
        groups[key]["entry"]   = e  # keep last for metadata

    entries = []
    for (vtype, inv_norm, date), g in groups.items():
        base = g["entry"].copy()
        base["amount"]             = g["amount"]
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
            "name":                    vtype,
            "found_in":                label,
            "count":                   count,
            "suggested_classification": sugg,
            "confidence":              conf,
            "auto_ignore":             auto_ign,
            "ignore_reason":           ign_reason,
        })

    return {"entries": entries, "voucher_types": voucher_types}
