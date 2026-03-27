"""
engine.py — easemybot Reconciliation Engine

Engine 1: Invoice matching (Purchase↔Sales, Debit↔Credit Notes)
  - Pass 1: Exact invoice number + exact amount
  - Pass 2: Exact invoice number + ±1% amount tolerance
  - Pass 3: Fuzzy invoice number + exact amount (common prefixes / substrings)
  - D/C Note matching: date + amount (±1%), no invoice number

Engine 2: Payment matching (Payment↔Receipt)
  - Exact date + exact amount
  - Duplicate handling

All entries outside the overlap period are excluded before matching.

Changes (v2):
  [1] Pass 1 — fixed tuple-vs-id bug causing one-to-many matching
  [2] Period parsing — parsed once, validated once upfront
  [3] Missing keys — .get() guards on voucher_type and invoice_number_norm
  [4] Amount coercion — _amount() safely coerces strings to float
  [5] Fuzzy O(N²) — hard cap; skipped if cartesian product > 5,000,000
  [6] Unmatched credit notes — reported symmetrically alongside debit notes
  [7] Blank invoice entries — sidelined before matching, reported separately
  [8] Reason tags — every unmatched entry carries a human-readable reason
"""

from collections import defaultdict
from datetime import date as Date


# ─────────────────────────────────────────────
# Human-readable reason labels
# ─────────────────────────────────────────────
REASON_LABELS = {
    "no_invoice_number_a":  "No invoice number found in A",
    "no_invoice_number_b":  "No invoice number found in B",
    "no_match_found_a":     "Invoice found in A — no match in B",
    "no_match_found_b":     "Invoice found in B — no match in A",
    "fuzzy_skipped":        "Match attempt incomplete — dataset too large",
}


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _to_date(s: str | None) -> Date | None:
    if not s:
        return None
    try:
        return Date.fromisoformat(s)
    except ValueError:
        return None


def _within_tolerance(a: float, b: float, pct: float = 1.0) -> bool:
    if a == 0 and b == 0:
        return True
    larger = max(abs(a), abs(b))
    return abs(a - b) / larger * 100 <= pct


def _amount(e) -> float:
    # FIX [4]: safely coerce strings or bad types
    v = e.get("amount", 0.0)
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _filter_by_period(entries: list, from_d: Date, to_d: Date) -> list:
    # FIX [2]: accepts pre-parsed date objects
    result = []
    for e in entries:
        d = _to_date(e.get("date"))
        if d and from_d <= d <= to_d:
            result.append(e)
    return result


def _group(entries: list) -> dict:
    """Group entries by normalised invoice number → list of entries."""
    g = defaultdict(list)
    for e in entries:
        # FIX [3]: guard missing invoice_number_norm
        key = e.get("invoice_number_norm") or ""
        g[key].append(e)
    return g


def _stamp_reason(entry: dict, reason_key: str) -> dict:
    """Return entry with unmatched_reason set to human-readable label."""
    e = entry.copy()
    e["unmatched_reason"] = REASON_LABELS.get(reason_key, reason_key)
    return e


# ─────────────────────────────────────────────
# Engine 1 — Invoice matching
# ─────────────────────────────────────────────

def _engine1(inv_a: list, inv_b: list, dn_a: list, dn_b: list,
             cn_a: list, cn_b: list) -> dict:

    matched_exact     = []
    matched_variation = []
    a_not_in_b        = []
    b_not_in_a        = []

    # FIX [7]: Sidelined blank invoice entries — pulled out before matching
    no_inv_a = [e for e in inv_a if not (e.get("invoice_number_norm") or "").strip()]
    no_inv_b = [e for e in inv_b if not (e.get("invoice_number_norm") or "").strip()]
    inv_a    = [e for e in inv_a if (e.get("invoice_number_norm") or "").strip()]
    inv_b    = [e for e in inv_b if (e.get("invoice_number_norm") or "").strip()]

    # Stamp reason tags on sidelined entries
    no_inv_a = [_stamp_reason(e, "no_invoice_number_a") for e in no_inv_a]
    no_inv_b = [_stamp_reason(e, "no_invoice_number_b") for e in no_inv_b]

    # ── Three-pass invoice matching ──────────────────────────────────────────
    group_a = _group(inv_a)
    group_b = _group(inv_b)

    def _match_passes(g_a, g_b):
        m_exact   = []
        m_var     = []
        matched_a = set()
        matched_b = set()

        # Pass 1 — exact invoice number + exact amount
        # FIX [1]: check id() directly, not tuples
        for inv_num, entries_a in g_a.items():
            if inv_num in g_b:
                for ea in entries_a:
                    if id(ea) in matched_a:
                        continue
                    for eb in g_b[inv_num]:
                        if id(eb) in matched_b:
                            continue
                        if abs(_amount(ea) - _amount(eb)) < 0.01:
                            m_exact.append({"entry_a": ea, "entry_b": eb, "difference": 0.0})
                            matched_a.add(id(ea))
                            matched_b.add(id(eb))
                            break

        # Pass 2 — exact invoice number + ±1% amount
        for inv_num, entries_a in g_a.items():
            if inv_num in g_b:
                for ea in entries_a:
                    if id(ea) in matched_a:
                        continue
                    for eb in g_b[inv_num]:
                        if id(eb) in matched_b:
                            continue
                        if _within_tolerance(_amount(ea), _amount(eb)):
                            diff = round(_amount(ea) - _amount(eb), 2)
                            m_var.append({"entry_a": ea, "entry_b": eb, "difference": diff})
                            matched_a.add(id(ea))
                            matched_b.add(id(eb))
                            break

        # Pass 3 — fuzzy: substring match on invoice number
        unmatched_a = [e for es in g_a.values() for e in es if id(e) not in matched_a]
        unmatched_b = [e for es in g_b.values() for e in es if id(e) not in matched_b]

        # FIX [5]: hard cap on fuzzy to prevent O(N²) timeout
        FUZZY_LIMIT = 5_000_000
        if len(unmatched_a) * len(unmatched_b) <= FUZZY_LIMIT:
            matched_a_fuzzy = set()
            matched_b_fuzzy = set()
            for ea in unmatched_a:
                if id(ea) in matched_a_fuzzy:
                    continue
                for eb in unmatched_b:
                    if id(eb) in matched_b_fuzzy:
                        continue
                    na = ea.get("invoice_number_norm") or ""
                    nb = eb.get("invoice_number_norm") or ""
                    if na and nb and (na in nb or nb in na):
                        if _within_tolerance(_amount(ea), _amount(eb)):
                            diff = round(_amount(ea) - _amount(eb), 2)
                            row = {"entry_a": ea, "entry_b": eb, "difference": diff}
                            if abs(diff) < 0.01:
                                m_exact.append(row)
                            else:
                                m_var.append(row)
                            matched_a_fuzzy.add(id(ea))
                            matched_b_fuzzy.add(id(eb))
                            break

            unmatched_a = [e for e in unmatched_a if id(e) not in matched_a_fuzzy]
            unmatched_b = [e for e in unmatched_b if id(e) not in matched_b_fuzzy]
        else:
            # FIX [8]: stamp fuzzy_skipped reason on entries that couldn't be checked
            unmatched_a = [_stamp_reason(e, "fuzzy_skipped") for e in unmatched_a]
            unmatched_b = [_stamp_reason(e, "fuzzy_skipped") for e in unmatched_b]
            return m_exact, m_var, unmatched_a, unmatched_b

        # FIX [8]: stamp no_match_found reason on remaining unmatched entries
        unmatched_a = [_stamp_reason(e, "no_match_found_a") for e in unmatched_a]
        unmatched_b = [_stamp_reason(e, "no_match_found_b") for e in unmatched_b]

        return m_exact, m_var, unmatched_a, unmatched_b

    me, mv, ua, ub = _match_passes(group_a, group_b)
    matched_exact.extend(me)
    matched_variation.extend(mv)
    a_not_in_b.extend(ua)
    b_not_in_a.extend(ub)

    # ── D/C Note matching: date + amount ±1%, NO invoice number ──────────────
    dn_matched_exact     = []
    dn_matched_variation = []
    dn_a_not_in_b        = []
    dn_b_not_in_a        = []
    cn_a_not_in_b        = []   # FIX [6]
    cn_b_not_in_a        = []   # FIX [6]

    sides = [
        (dn_a, cn_b, dn_a_not_in_b, cn_b_not_in_a, "no_match_found_a", "no_match_found_b"),
        (dn_b, cn_a, dn_b_not_in_a, cn_a_not_in_b, "no_match_found_b", "no_match_found_a"),
    ]

    for source_list, target_list, unmatched_src, unmatched_tgt, src_reason, tgt_reason in sides:
        used_target: set = set()
        for ea in source_list:
            found = False
            for eb in target_list:
                if id(eb) in used_target:
                    continue
                if ea.get("date") == eb.get("date"):
                    amt_a, amt_b = _amount(ea), _amount(eb)
                    if abs(amt_a - amt_b) < 0.01:
                        dn_matched_exact.append({"entry_a": ea, "entry_b": eb, "difference": 0.0})
                        used_target.add(id(eb))
                        found = True
                        break
                    elif _within_tolerance(amt_a, amt_b):
                        diff = round(amt_a - amt_b, 2)
                        dn_matched_variation.append({"entry_a": ea, "entry_b": eb, "difference": diff})
                        used_target.add(id(eb))
                        found = True
                        break
            if not found:
                unmatched_src.append(_stamp_reason(ea, src_reason))

        # FIX [6]: collect unmatched credit notes
        for eb in target_list:
            if id(eb) not in used_target:
                unmatched_tgt.append(_stamp_reason(eb, tgt_reason))

    # FIX [6]: include credit notes in totals
    total_a = len(inv_a) + len(no_inv_a) + len(dn_a) + len(cn_a)
    total_b = len(inv_b) + len(no_inv_b) + len(dn_b) + len(cn_b)

    all_exact     = matched_exact + dn_matched_exact
    all_variation = matched_variation + dn_matched_variation

    # FIX [7]: exclude no-invoice entries from a_not_in_b / b_not_in_a
    # They are reported separately in _no_invoice_number
    all_ua = a_not_in_b + dn_a_not_in_b + cn_a_not_in_b
    all_ub = b_not_in_a + dn_b_not_in_a + cn_b_not_in_a

    return {
        "total_a":             total_a,
        "total_b":             total_b,
        "matched_exact":       len(all_exact),
        "matched_variation":   len(all_variation),
        "a_not_in_b_count":    len(all_ua),
        "a_not_in_b_value":    round(sum(_amount(e) for e in all_ua), 2),
        "b_not_in_a_count":    len(all_ub),
        "b_not_in_a_value":    round(sum(_amount(e) for e in all_ub), 2),
        "no_invoice_count_a":  len(no_inv_a),
        "no_invoice_value_a":  round(sum(_amount(e) for e in no_inv_a), 2),
        "no_invoice_count_b":  len(no_inv_b),
        "no_invoice_value_b":  round(sum(_amount(e) for e in no_inv_b), 2),
        # Detailed rows for Excel
        "_matched_exact":      all_exact,
        "_matched_variation":  all_variation,
        "_a_not_in_b":         all_ua,
        "_b_not_in_a":         all_ub,
        "_no_invoice_number":  no_inv_a + no_inv_b,   # Sheet 7
        # Granular D/C note breakdown
        "_dn_a_not_in_b":      dn_a_not_in_b,
        "_dn_b_not_in_a":      dn_b_not_in_a,
        "_cn_a_not_in_b":      cn_a_not_in_b,
        "_cn_b_not_in_a":      cn_b_not_in_a,
    }


# ─────────────────────────────────────────────
# Engine 2 — Payment matching
# ─────────────────────────────────────────────

def _engine2(pay_a: list, rec_b: list, rec_a: list, pay_b: list) -> dict:
    matched    = []
    unmatched_a: list = []
    unmatched_b: list = []

    def _cross_match(source: list, target: list, unmatched_source: list):
        target_pool: dict = defaultdict(list)
        for e in target:
            key = (e.get("date"), round(_amount(e), 2))
            target_pool[key].append(e)

        used_ids: set = set()
        for ea in source:
            key = (ea.get("date"), round(_amount(ea), 2))
            avail = [e for e in target_pool.get(key, []) if id(e) not in used_ids]
            if avail:
                eb = avail[0]
                matched.append({"entry_a": ea, "entry_b": eb})
                used_ids.add(id(eb))
            else:
                unmatched_source.append(ea)

    _cross_match(pay_a, rec_b, unmatched_a)
    _cross_match(rec_a, pay_b, unmatched_b)  # FIX: was unmatched_a — rec_a mismatches are B's missing payments, not A's

    all_used_ids = {id(m["entry_b"]) for m in matched}
    for eb in list(rec_b) + list(pay_b):
        if id(eb) not in all_used_ids:
            unmatched_b.append(eb)

    return {
        "total_a":          len(pay_a) + len(rec_a),
        "total_b":          len(pay_b) + len(rec_b),
        "matched":          len(matched),
        "a_not_in_b_count": len(unmatched_a),
        "a_not_in_b_value": round(sum(_amount(e) for e in unmatched_a), 2),
        "b_not_in_a_count": len(unmatched_b),
        "b_not_in_a_value": round(sum(_amount(e) for e in unmatched_b), 2),
        "_matched":    matched,
        "_a_not_in_b": unmatched_a,
        "_b_not_in_a": unmatched_b,
    }


# ─────────────────────────────────────────────
# Public: run_reconciliation
# ─────────────────────────────────────────────

def run_reconciliation(
    parsed_a: dict,
    parsed_b: dict,
    classifications: dict,
    period_from: str,
    period_to: str,
) -> dict:

    # FIX [2]: parse and validate period dates once, upfront
    from_d = _to_date(period_from)
    to_d   = _to_date(period_to)
    if not from_d or not to_d:
        raise ValueError(f"Invalid period_from/period_to: '{period_from}' / '{period_to}'")
    if from_d > to_d:
        raise ValueError(f"period_from ({period_from}) cannot be after period_to ({period_to})")

    def _classify(entries: list) -> dict:
        buckets: dict = defaultdict(list)
        for e in entries:
            # FIX [3]: guard missing voucher_type
            cls = classifications.get(e.get("voucher_type"), "ignore")
            if cls != "ignore":
                buckets[cls].append(e)
        return buckets

    ea = _filter_by_period(parsed_a["entries"], from_d, to_d)
    eb = _filter_by_period(parsed_b["entries"], from_d, to_d)

    ba = _classify(ea)
    bb = _classify(eb)

    inv_a = ba.get("purchase_invoice", []) + ba.get("sales_invoice", [])
    inv_b = bb.get("purchase_invoice", []) + bb.get("sales_invoice", [])
    dn_a  = ba.get("debit_note",  [])
    dn_b  = bb.get("debit_note",  [])
    cn_a  = ba.get("credit_note", [])
    cn_b  = bb.get("credit_note", [])

    invoice_results = _engine1(inv_a, inv_b, dn_a, dn_b, cn_a, cn_b)

    pay_a = ba.get("payment", [])
    rec_a = ba.get("receipt", [])
    pay_b = bb.get("payment", [])
    rec_b = bb.get("receipt", [])

    payment_results = _engine2(pay_a, rec_b, rec_a, pay_b)

    return {
        "invoices":    invoice_results,
        "payments":    payment_results,
        "period_from": period_from,
        "period_to":   period_to,
    }
