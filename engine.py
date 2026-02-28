"""
engine.py — RecoBot Reconciliation Engine

Engine 1: Invoice matching (Purchase↔Sales, Debit↔Credit Notes)
  - Pass 1: Exact invoice number + exact amount
  - Pass 2: Exact invoice number + ±1% amount tolerance
  - Pass 3: Fuzzy invoice number + exact amount (common prefixes / substrings)
  - D/C Note matching: date + amount (±1%), no invoice number

Engine 2: Payment matching (Payment↔Receipt)
  - Exact date + exact amount
  - Duplicate handling

All entries outside the overlap period are excluded before matching.
"""

from collections import defaultdict
from datetime import date as Date


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


def _within_period(entry, from_str: str, to_str: str) -> bool:
    d = _to_date(entry.get("date"))
    if d is None:
        return False
    return _to_date(from_str) <= d <= _to_date(to_str)


def _within_tolerance(a: float, b: float, pct: float = 1.0) -> bool:
    if a == 0 and b == 0:
        return True
    larger = max(abs(a), abs(b))
    return abs(a - b) / larger * 100 <= pct


def _filter_by_period(entries: list, from_str: str, to_str: str) -> list:
    return [e for e in entries if _within_period(e, from_str, to_str)]


def _group(entries: list) -> dict:
    """Group entries by normalised invoice number → list of entries."""
    g = defaultdict(list)
    for e in entries:
        g[e["invoice_number_norm"]].append(e)
    return g


def _amount(e) -> float:
    return e.get("amount", 0.0)


# ─────────────────────────────────────────────
# Engine 1 — Invoice matching
# ─────────────────────────────────────────────

def _engine1(inv_a: list, inv_b: list, dn_a: list, dn_b: list,
             cn_a: list, cn_b: list) -> dict:
    """
    Matches:
      - Purchase invoices (A) ↔ Sales invoices (B)   [and vice-versa]
      - Debit notes ↔ Credit notes (date + amount)
    """

    matched_exact     = []  # invoice number + amount identical
    matched_variation = []  # invoice number match + amount within 1%
    a_not_in_b        = []  # in A, no match in B
    b_not_in_a        = []  # in B, no match in A

    # ── Three-pass invoice matching ──────────────────────────────────────────
    group_a = _group(inv_a)  # norm_invoice → [entries]
    group_b = _group(inv_b)

    used_a: set = set()
    used_b: set = set()

    def _match_passes(g_a, g_b):
        """Run 3 passes, return (exact, variation, unmatched_a_keys, unmatched_b_keys)."""
        m_exact = []
        m_var   = []
        matched_a = set()
        matched_b = set()

        # Pass 1 — exact invoice number + exact amount
        for inv_num, entries_a in g_a.items():
            if inv_num in g_b:
                for ea in entries_a:
                    for eb in g_b[inv_num]:
                        if (ea["invoice_number_norm"], id(ea)) not in matched_a and \
                           (eb["invoice_number_norm"], id(eb)) not in matched_b:
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

        # Pass 3 — fuzzy: check if one invoice number is contained in another
        unmatched_a = [e for es in g_a.values() for e in es if id(e) not in matched_a]
        unmatched_b = [e for es in g_b.values() for e in es if id(e) not in matched_b]

        matched_a_fuzzy = set()
        matched_b_fuzzy = set()
        for ea in unmatched_a:
            for eb in unmatched_b:
                na, nb = ea["invoice_number_norm"], eb["invoice_number_norm"]
                if na and nb and (na in nb or nb in na):
                    if _within_tolerance(_amount(ea), _amount(eb)):
                        diff = round(_amount(ea) - _amount(eb), 2)
                        row = {"entry_a": ea, "entry_b": eb, "difference": diff}
                        if diff == 0.0:
                            m_exact.append(row)
                        else:
                            m_var.append(row)
                        matched_a_fuzzy.add(id(ea))
                        matched_b_fuzzy.add(id(eb))
                        break

        final_unmatched_a = [e for e in unmatched_a if id(e) not in matched_a_fuzzy]
        final_unmatched_b = [e for e in unmatched_b if id(e) not in matched_b_fuzzy]

        return m_exact, m_var, final_unmatched_a, final_unmatched_b

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

    used_dn_a: set = set()
    used_dn_b: set = set()

    # Debit note A ↔ Credit note B  (and swap)
    sides = [(dn_a, cn_b, dn_a_not_in_b), (dn_b, cn_a, dn_b_not_in_a)]

    for source_list, target_list, not_matched_list in sides:
        used_target: set = set()
        for ea in source_list:
            found = False
            for eb in target_list:
                if id(eb) in used_target:
                    continue
                if ea.get("date") == eb.get("date"):
                    if abs(_amount(ea) - _amount(eb)) < 0.01:
                        dn_matched_exact.append({"entry_a": ea, "entry_b": eb, "difference": 0.0})
                        used_target.add(id(eb))
                        found = True
                        break
                    elif _within_tolerance(_amount(ea), _amount(eb)):
                        diff = round(_amount(ea) - _amount(eb), 2)
                        dn_matched_variation.append({"entry_a": ea, "entry_b": eb, "difference": diff})
                        used_target.add(id(eb))
                        found = True
                        break
            if not found:
                not_matched_list.append(ea)

    total_a = len(inv_a) + len(dn_a)
    total_b = len(inv_b) + len(dn_b)
    all_exact     = matched_exact + dn_matched_exact
    all_variation = matched_variation + dn_matched_variation
    all_ua        = a_not_in_b + dn_a_not_in_b
    all_ub        = b_not_in_a + dn_b_not_in_a

    return {
        "total_a":          total_a,
        "total_b":          total_b,
        "matched_exact":    len(all_exact),
        "matched_variation": len(all_variation),
        "a_not_in_b_count": len(all_ua),
        "a_not_in_b_value": round(sum(_amount(e) for e in all_ua), 2),
        "b_not_in_a_count": len(all_ub),
        "b_not_in_a_value": round(sum(_amount(e) for e in all_ub), 2),
        # Detailed rows for Excel
        "_matched_exact":     all_exact,
        "_matched_variation": all_variation,
        "_a_not_in_b":        all_ua,
        "_b_not_in_a":        all_ub,
    }


# ─────────────────────────────────────────────
# Engine 2 — Payment matching
# ─────────────────────────────────────────────

def _engine2(pay_a: list, rec_b: list, rec_a: list, pay_b: list) -> dict:
    """
    Cross-type matching:
      Payment A ↔ Receipt B
      Receipt A ↔ Payment B
    Exact date + exact amount. No tolerance.
    Handles duplicates: same date+amount appears N times → match one-to-one.
    """
    matched       = []
    a_not_in_b    = []
    b_not_in_a    = []

    def _cross_match(source: list, target: list,
                     unmatched_source: list, unmatched_target: list):
        used_target = defaultdict(int)  # (date, amount) → used count
        target_pool = defaultdict(list)  # (date, amount) → [entries]
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

        # Find unmatched in target
        all_used_ids = {id(m["entry_b"]) for m in matched}
        for eb in target:
            if id(eb) not in all_used_ids:
                unmatched_target.append(eb)

    unmatched_a: list = []
    unmatched_b: list = []
    _cross_match(pay_a, rec_b, unmatched_a, unmatched_b)
    _cross_match(rec_a, pay_b, unmatched_a, unmatched_b)

    return {
        "total_a":          len(pay_a) + len(rec_a),
        "total_b":          len(pay_b) + len(rec_b),
        "matched":          len(matched),
        "a_not_in_b_count": len(unmatched_a),
        "a_not_in_b_value": round(sum(_amount(e) for e in unmatched_a), 2),
        "b_not_in_a_count": len(unmatched_b),
        "b_not_in_a_value": round(sum(_amount(e) for e in unmatched_b), 2),
        # Detailed rows for Excel
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
    classifications: dict,        # voucher_type_name → classification string
    period_from: str,
    period_to: str,
) -> dict:
    """
    Orchestrate Engine 1 + Engine 2.

    classifications values:
      purchase_invoice | sales_invoice | payment | receipt |
      debit_note | credit_note | ignore
    """

    # Map each entry to its classification based on its voucher_type
    def _classify(entries: list, source_label: str) -> dict:
        """Return { classification: [entries] }."""
        buckets: dict = defaultdict(list)
        for e in entries:
            cls = classifications.get(e["voucher_type"], "ignore")
            if cls != "ignore":
                buckets[cls].append(e)
        return buckets

    # Filter to period window
    ea = _filter_by_period(parsed_a["entries"], period_from, period_to)
    eb = _filter_by_period(parsed_b["entries"], period_from, period_to)

    ba = _classify(ea, "A")
    bb = _classify(eb, "B")

    # Invoice engine:
    # Party A's purchase invoices ↔ Party B's sales invoices (and vice-versa)
    inv_a = ba.get("purchase_invoice", []) + bb.get("sales_invoice", [])
    inv_b = bb.get("purchase_invoice", []) + ba.get("sales_invoice", [])

    dn_a = ba.get("debit_note",  [])
    dn_b = bb.get("debit_note",  [])
    cn_a = ba.get("credit_note", [])
    cn_b = bb.get("credit_note", [])

    invoice_results = _engine1(inv_a, inv_b, dn_a, dn_b, cn_a, cn_b)

    # Payment engine:
    pay_a = ba.get("payment", [])
    rec_a = ba.get("receipt", [])
    pay_b = bb.get("payment", [])
    rec_b = bb.get("receipt", [])

    payment_results = _engine2(pay_a, rec_b, rec_a, pay_b)

    return {
        "invoices": invoice_results,
        "payments": payment_results,
        "period_from": period_from,
        "period_to":   period_to,
    }
