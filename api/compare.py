"""
POST /api/compare
Body: { "ours": [...], "bank": [...], "bank_pre_range_balance": <float|null> }

Matching philosophy for BANK reconciliation - deliberately simpler than a
supplier reconciliation: banks don't have invoice numbers, and any
reference/cheque number a bank prints is usually irrelevant on the OTHER
side's ledger, so ID matching is not used AT ALL. Every row is matched
purely by amount, exact to the cent (no $1-style tolerance - a bank
statement is a source of truth to the cent, so amount tolerance stays at
essentially zero, just enough to absorb float rounding). Date is used
only to disambiguate when several rows share the same amount, and to
flag a real but harmless timing difference (a payment posted on our side
a day or two before/after the bank shows it) as a date_mismatch rather
than silently treating differing dates as identical or, worse, refusing
to match at all.
"""

from http.server import BaseHTTPRequestHandler
import json
import datetime as dt

BUILD_TAG = "2026-08-27-day-total-matching"

AMOUNT_TOLERANCE = 0.01  # exact to the cent - only float rounding is absorbed, nothing more


def amounts_close(a, b):
    return abs((a or 0) - (b or 0)) <= AMOUNT_TOLERANCE


def _parse_iso_date(s):
    if not s:
        return None
    try:
        return dt.date.fromisoformat(s)
    except ValueError:
        return None


def clean(rows):
    """Keep only real transaction rows with a genuine amount. Every row
    gets 'amt' (whichever of debit/credit is non-zero - a bank row is
    always one or the other, never both) for amount-only matching."""
    out = []
    for r in rows:
        if r.get("row_type") != "transaction":
            continue
        debit = r.get("debit") or 0.0
        credit = r.get("credit") or 0.0
        if debit <= 0.01 and credit <= 0.01:
            continue
        out.append({
            "date": r.get("date") or "",
            "id": r.get("id") or "",
            "description": r.get("description") or "",
            "debit": debit,
            "credit": credit,
            "balance": r.get("balance"),
            "amt": debit if debit > 0.01 else credit,
        })
    return out


def match_by_amount(ours, bank, ours_pool, bank_pool):
    """The ENTIRE matching algorithm for this project. Every candidate
    pair with a matching amount (to the cent) is a candidate, globally
    ranked by how close their dates are - the closest-date pairing always
    wins a shared amount before a farther one, so two rows for the exact
    same amount on different real dates don't get paired arbitrarily by
    list order. A pair with identical dates is a clean match; a pair
    whose dates differ is still a match (the amount agreeing is treated
    as strong enough evidence on its own) but reported as a date_mismatch
    so the timing difference is visible rather than silently hidden."""
    candidates = []
    for i in ours_pool:
        o_date = _parse_iso_date(ours[i]["date"])
        for j in bank_pool:
            if not amounts_close(ours[i]["amt"], bank[j]["amt"]):
                continue
            b_date = _parse_iso_date(bank[j]["date"])
            diff_days = abs((o_date - b_date).days) if (o_date and b_date) else 10**9
            candidates.append((diff_days, i, j))
    candidates.sort(key=lambda c: c[0])

    used_o, used_b, pairs = set(), set(), []
    for _, i, j in candidates:
        if i in used_o or j in used_b:
            continue
        pairs.append((i, j))
        used_o.add(i)
        used_b.add(j)

    exact_pairs = [(i, j) for i, j in pairs if ours[i]["date"] == bank[j]["date"]]
    date_mm_pairs = [(i, j) for i, j in pairs if ours[i]["date"] != bank[j]["date"]]
    leftover_ours = [i for i in ours_pool if i not in used_o]
    leftover_bank = [j for j in bank_pool if j not in used_b]
    return exact_pairs, date_mm_pairs, leftover_ours, leftover_bank


def row_out(issue, o=None, b=None):
    return {
        "issue": issue,
        "date": (o or b)["date"],
        "id": (o or {}).get("id") or (b or {}).get("id") or "",
        "description": (o or b)["description"],
        "our_date": o["date"] if o else None,
        "bank_date": b["date"] if b else None,
        "our_id": o["id"] if o else None,
        "bank_id": b["id"] if b else None,
        "our_debit": o["debit"] if o else None,
        "our_credit": o["credit"] if o else None,
        "bank_debit": b["debit"] if b else None,
        "bank_credit": b["credit"] if b else None,
    }


def matched_out(o, b):
    return {
        "date": o["date"],
        "id": o["id"] or b["id"],
        "our_id": o["id"],
        "bank_id": b["id"],
        "description": o["description"],
        "our_debit": o["debit"],
        "our_credit": o["credit"],
        "bank_debit": b["debit"],
        "bank_credit": b["credit"],
    }


def out_of_our_range(row, our_range):
    if our_range is None or not row["date"]:
        return False
    return row["date"] < our_range[0] or row["date"] > our_range[1]


def match_by_day_total(ours, bank, ours_pool, bank_pool):
    """Last resort for whatever's still unmatched after every individual
    line-item check above. Some accounts split the same day's activity
    into a DIFFERENT NUMBER of lines on each side - seen for real: our
    ledger records daily settlement batches, the bank statement lists
    each card network (VISA, Mastercard, etc.) separately, so individual
    amounts almost never line up even though nothing is actually missing
    (verified on a real account: July 1st was 7 lines on our side and 5
    on the bank's, but both totaled exactly $38,700.96). Groups whatever
    is still unmatched by date, and if a WHOLE DAY's net total agrees on
    both sides, treats every row that day as reconciled as a group
    instead of reporting dozens of individually-unmatched rows.

    This can only ever REDUCE how many rows get reported as missing - it
    never touches a row already paired by an earlier, stricter step, and
    it changes nothing about the totals themselves: every row already
    counts in our_total_debit/credit and bank_total_debit/credit
    regardless of match status, both before and after this step. It only
    changes which ISSUE CATEGORY a leftover row lands in."""
    ours_by_date, bank_by_date = {}, {}
    for i in ours_pool:
        ours_by_date.setdefault(ours[i]["date"], []).append(i)
    for j in bank_pool:
        bank_by_date.setdefault(bank[j]["date"], []).append(j)

    day_matched_ours, day_matched_bank = [], []
    for date in set(ours_by_date) & set(bank_by_date):
        o_indices, b_indices = ours_by_date[date], bank_by_date[date]
        o_net = sum(ours[i]["credit"] - ours[i]["debit"] for i in o_indices)
        b_net = sum(bank[j]["debit"] - bank[j]["credit"] for j in b_indices)
        if abs(o_net) > 0.01 and amounts_close(o_net, b_net):
            day_matched_ours.extend(o_indices)
            day_matched_bank.extend(b_indices)

    day_matched_ours_set, day_matched_bank_set = set(day_matched_ours), set(day_matched_bank)
    leftover_ours = [i for i in ours_pool if i not in day_matched_ours_set]
    leftover_bank = [j for j in bank_pool if j not in day_matched_bank_set]
    return day_matched_ours, day_matched_bank, leftover_ours, leftover_bank


def compare(ours_raw, bank_raw, bank_pre_range_balance=None):
    ours = clean(ours_raw)
    bank = clean(bank_raw)

    o_pool = list(range(len(ours)))
    b_pool = list(range(len(bank)))

    exact, date_mm, o_pool, b_pool = match_by_amount(ours, bank, o_pool, b_pool)
    day_matched_ours, day_matched_bank, o_pool, b_pool = match_by_day_total(ours, bank, o_pool, b_pool)

    matched_rows = [matched_out(ours[i], bank[j]) for i, j in exact]
    issues = []
    for i, j in date_mm:
        issues.append(row_out("date_mismatch", ours[i], bank[j]))
    for i in day_matched_ours:
        issues.append(row_out("day_total_match", ours[i], None))
    for j in day_matched_bank:
        issues.append(row_out("day_total_match", None, bank[j]))
    for i in o_pool:
        issues.append(row_out("missing_in_bank", ours[i], None))
    for j in b_pool:
        issues.append(row_out("missing_in_ours", None, bank[j]))

    our_dates = [r["date"] for r in ours if r["date"]]
    our_range = (min(our_dates), max(our_dates)) if our_dates else None

    our_total_debit = round(sum(r["debit"] for r in ours), 2)
    our_total_credit = round(sum(r["credit"] for r in ours), 2)
    our_net_debit_credit = round(our_total_credit - our_total_debit, 2)

    # Any bank row paired with one of ours - clean match or date_mismatch
    # alike - is by definition relevant to this reconciliation and counts
    # in the totals regardless of its own date (e.g. the bank posting a
    # transfer a day later than our own books). Only a row nothing was
    # ever paired with is still subject to the date-range exclusion.
    matched_bank_indices = set(j for _, j in exact) | set(j for _, j in date_mm)
    bank_in_range = [
        bank[j] for j in range(len(bank))
        if j in matched_bank_indices or not out_of_our_range(bank[j], our_range)
    ]
    bank_total_debit = round(sum(r["debit"] for r in bank_in_range), 2)
    bank_total_credit = round(sum(r["credit"] for r in bank_in_range), 2)
    # NOT credit-minus-debit, unlike our own side. Confirmed empirically
    # against real matched pairs: our_debit always equals the bank's
    # credit for the SAME transaction, and our_credit always equals the
    # bank's debit (e.g. a card-sale settlement lands in OUR debit column
    # but the bank's own statement shows it as a positive/credit entry;
    # a bank fee is OUR credit but the bank's own negative/debit entry).
    # The two sides are mirrored for the same real-world transaction, the
    # same way a supplier's AP ledger mirrors ours - so combining them
    # with the SAME "credit minus debit" formula on both sides would
    # produce close to DOUBLE the real gap for a well-matched dataset
    # (verified: summing credit-debit across matched pairs gave +83,518.96
    # on our side and -83,518.96 on the bank side - exact opposites, for
    # transactions that are by definition NOT in dispute). Flipping this
    # side's formula to debit-minus-credit re-aligns it to the same
    # orientation as our_net, so a fully-matched pair correctly
    # contributes zero to the difference instead of double-counting it.
    bank_net_debit_credit = round(bank_total_debit - bank_total_credit, 2)

    # Cross-check Net against each file's own printed Balance column,
    # same reasoning as the supplier tool: a debit/credit sum can silently
    # miss a genuine extraction gap that a balance-column check would
    # catch immediately, so prefer the balance-derived figure whenever it
    # agrees with the debit/credit total (within $1) - and fall back
    # quietly to debit/credit otherwise rather than trusting a
    # balance-column figure that disagrees.
    our_opening_rows = [r for r in ours_raw if r.get("row_type") == "opening_balance" and r.get("balance") is not None]
    our_closing_rows = [r for r in ours_raw if r.get("row_type") == "closing_balance" and r.get("balance") is not None]
    our_net_balance = None
    if len(our_opening_rows) == 1 and len(our_closing_rows) == 1:
        our_net_balance = round(our_opening_rows[0]["balance"] - our_closing_rows[0]["balance"], 2)

    our_net = our_net_balance if (our_net_balance is not None and amounts_close(our_net_balance, our_net_debit_credit)) else our_net_debit_credit
    our_net_source = "balance" if our_net == our_net_balance and our_net_balance is not None else "debit_credit"

    bank_net_balance = None
    if bank_pre_range_balance is not None:
        in_range_dated = [r for r in bank_in_range if r["date"]]
        if in_range_dated:
            end_row = sorted(enumerate(in_range_dated), key=lambda ir: ir[1]["date"])[-1][1]
            if end_row.get("balance") is not None:
                # Flipped to match bank_net_debit_credit's new orientation
                # above (pre-range minus end, not end minus pre-range) -
                # the bank's own balance still rises with a positive/credit
                # entry exactly as printed, but bank_net itself is now
                # reported debit-minus-credit, so the balance-derived
                # figure needs the same flip to compare apples to apples.
                bank_net_balance = round(bank_pre_range_balance - end_row["balance"], 2)

    bank_net = bank_net_balance if (bank_net_balance is not None and amounts_close(bank_net_balance, bank_net_debit_credit)) else bank_net_debit_credit
    bank_net_source = "balance" if bank_net == bank_net_balance and bank_net_balance is not None else "debit_credit"

    net_difference = round(our_net - bank_net, 2)

    summary = {
        "build_tag": BUILD_TAG,
        "our_transactions": len(ours),
        "bank_transactions": len(bank),
        "matched": len(exact),
        "date_mismatch": len(date_mm),
        "day_total_match": len(day_matched_ours) + len(day_matched_bank),
        "missing_in_bank": len(o_pool),
        "missing_in_ours": len(b_pool),
        "our_date_range": list(our_range) if our_range else None,
        "our_total_debit": our_total_debit,
        "our_total_credit": our_total_credit,
        "our_net": our_net,
        "our_net_source": our_net_source,
        "our_net_debit_credit": our_net_debit_credit,
        "bank_total_debit": bank_total_debit,
        "bank_total_credit": bank_total_credit,
        "bank_net": bank_net,
        "bank_net_source": bank_net_source,
        "bank_net_debit_credit": bank_net_debit_credit,
        "net_difference": net_difference,
    }
    return {"summary": summary, "matched": matched_rows, "issues": issues}


class handler(BaseHTTPRequestHandler):
    def _send(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length) or b"{}")
            ours = data.get("ours", [])
            bank = data.get("bank", [])
            pre_range_balance = data.get("bank_pre_range_balance")
            self._send(200, compare(ours, bank, pre_range_balance))
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Compare failed: {e}"})
