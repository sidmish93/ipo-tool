"""Post-SHP block/bulk adjustment.

A quarterly SHP is as-of the quarter-end date (June SHP = through 30 June).
Disclosed NSE/BSE block and bulk trades from the next calendar day through
today are applied in share counts: sells subtract, buys add. Adjusted % is
adjusted shares / shares outstanding.

Bulk and block feeds often describe the same client-day, so block quantity
is counted once and only leftover bulk is added (same rule as blocks_tracker).
Same-day self-trades (one client on both sides) are ignored.
"""

from __future__ import annotations

import calendar
import csv
import datetime
import io
import re
from collections import defaultdict

_NON_ALNUM = re.compile(r"[^A-Z0-9]+")
_TOKEN_ALIASES = {
    "LIMITED": "LTD",
    "PRIVATE": "PVT",
    "COMPANY": "CO",
    "CORPORATION": "CORP",
    "AND": "&",
    "INVESTMENTS": "INVESTMENT",
    "SECURITIES": "SEC",
}
_DROP = {"THE"}
_SUFFIXES = ("LTD", "PVT LTD", "PVT", "CORP", "INC", "PLC")

_MONTH_END = {
    "jan": (1, 31), "january": (1, 31),
    "feb": (2, 28), "february": (2, 28),
    "mar": (3, 31), "march": (3, 31),
    "apr": (4, 30), "april": (4, 30),
    "may": (5, 31),
    "jun": (6, 30), "june": (6, 30),
    "jul": (7, 31), "july": (7, 31),
    "aug": (8, 31), "august": (8, 31),
    "sep": (9, 30), "sept": (9, 30), "september": (9, 30),
    "oct": (10, 31), "october": (10, 31),
    "nov": (11, 30), "november": (11, 30),
    "dec": (12, 31), "december": (12, 31),
}

_FETCH_TEXT = """async (u) => {
    try {
      const r = await fetch(u);
      return {ok: r.ok, status: r.status, text: await r.text()};
    } catch (e) { return {ok: false, status: 0, text: ''+e}; }
}"""

NSE_DEALS = ("https://www.nseindia.com/api/historicalOR/bulk-block-short-deals"
             "?optionType={option}&from={start}&to={end}&csv=true")
BSE_DEALS = ("https://api.bseindia.com/BseIndiaAPI/api/BulkDealData_ng/w"
             "?DealType={deal_type}&sc_code={scrip}&FDate={start}&TDate={end}")


def canonical(text):
    if not text:
        return ""
    cleaned = _NON_ALNUM.sub(" ", str(text).upper()).strip()
    tokens = [_TOKEN_ALIASES.get(t, t) for t in cleaned.split()]
    tokens = [t for t in tokens if t not in _DROP]
    return " ".join(tokens)


def canonical_party(name):
    text = canonical(name)
    changed = True
    while changed:
        changed = False
        for suf in _SUFFIXES:
            if text.endswith(" " + suf):
                text = text[: -(len(suf) + 1)].strip()
                changed = True
    return text


def quarter_end(qname):
    """Latest day covered by a SHP labelled 'June 2026' / '30-Jun-2026'."""
    if not qname:
        return None
    qname = str(qname).strip()
    m = re.search(r"(\d{1,2})[-/ ]([A-Za-z]{3,})[-/ ](\d{2,4})", qname)
    if m:
        mon = _MONTH_END.get(m.group(2).lower())
        yr = int(m.group(3))
        if yr < 100:
            yr += 2000
        if mon:
            last = calendar.monthrange(yr, mon[0])[1]
            day = min(int(m.group(1)), last)
            return datetime.date(yr, mon[0], day)
    m = re.search(r"([A-Za-z]{3,})\s+(\d{4})", qname)
    if m:
        mon = _MONTH_END.get(m.group(1).lower())
        if mon:
            yr = int(m.group(2))
            return datetime.date(yr, mon[0], calendar.monthrange(yr, mon[0])[1])
    return None


def adjust_window(qname, today=None):
    """(from_date, to_date) = day after quarter-end through today."""
    today = today or datetime.date.today()
    end = quarter_end(qname)
    if not end:
        return None, None
    start = end + datetime.timedelta(days=1)
    if start > today:
        return start, today
    return start, today


def _num(value):
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return 0.0


def implied_outstanding(holders):
    """Median shares / (pct/100) from named holders with a real stake."""
    estimates = []
    for h in holders or []:
        shares = h.get("shares")
        pct = h.get("pct")
        if shares and pct and pct >= 0.05:
            estimates.append(float(shares) * 100.0 / float(pct))
    if not estimates:
        return None
    estimates.sort()
    return round(estimates[len(estimates) // 2])


def names_match(holder, client):
    a, b = canonical_party(holder), canonical_party(client)
    if not a or not b:
        return False
    if a == b:
        return True
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    if shorter in longer and (len(shorter) >= 10 or len(shorter.split()) >= 2):
        return True
    return False


def _best_holder(holders, client):
    exact, loose = [], []
    ck = canonical_party(client)
    for h in holders:
        hk = canonical_party(h.get("name"))
        if not hk or not ck:
            continue
        if hk == ck:
            exact.append(h)
        elif names_match(h.get("name"), client):
            loose.append(h)
    if exact:
        return exact[0]
    if len(loose) == 1:
        return loose[0]
    if len(loose) > 1:
        scored = []
        for h in loose:
            hk = set(canonical_party(h.get("name")).split())
            ck_set = set(ck.split())
            scored.append((len(hk - ck_set) + len(ck_set - hk), h))
        scored.sort(key=lambda x: x[0])
        if scored[0][0] < scored[1][0]:
            return scored[0][1]
    return None


def _dedupe_client_day(legs):
    """One quantity per (exchange, client, side, date): block + leftover bulk."""
    grouped = defaultdict(lambda: {"BULK": 0.0, "BLOCK": 0.0, "name": ""})
    for leg in legs:
        key = (leg["exchange"], canonical_party(leg["client"]),
               leg["side"], leg["date"])
        bucket = grouped[key]
        bucket[leg["deal_type"]] += leg["qty"]
        if len(leg["client"]) > len(bucket["name"]):
            bucket["name"] = leg["client"]

    sides = defaultdict(set)
    for (ex, client, side, day), _ in grouped.items():
        sides[(ex, client, day)].add(side)
    self_traders = {k for k, s in sides.items() if len(s) > 1}

    out = []
    for (ex, client, side, day), bucket in grouped.items():
        if (ex, client, day) in self_traders:
            continue
        block, bulk = bucket["BLOCK"], bucket["BULK"]
        qty = block + max(0.0, bulk - block)
        if qty <= 0:
            continue
        out.append({
            "exchange": ex,
            "client": bucket["name"] or client,
            "side": side,
            "date": day,
            "qty": qty,
        })
    return out


def _merge_exchanges(legs):
    """Collapse the same client/side/date printed on NSE and BSE.

    Near-equal quantities are one deal copied on both feeds (count once).
    Different quantities are two real prints (sum).
    """
    grouped = defaultdict(list)
    for leg in legs:
        grouped[(canonical_party(leg["client"]), leg["side"],
                 leg["date"])].append(leg)
    out = []
    for _, items in grouped.items():
        if len(items) == 1:
            out.append(items[0])
            continue
        by_ex = defaultdict(float)
        name = ""
        for it in items:
            by_ex[it["exchange"]] += it["qty"]
            if len(it.get("client") or "") > len(name):
                name = it["client"]
        qtys = [q for q in by_ex.values() if q > 0]
        if (len(qtys) == 2
                and abs(qtys[0] - qtys[1]) / max(qtys) <= 0.02):
            qty = max(qtys)
        else:
            qty = sum(qtys)
        first = items[0]
        out.append({
            "exchange": "+".join(sorted(by_ex)),
            "client": name or first["client"],
            "side": first["side"],
            "date": first["date"],
            "qty": qty,
        })
    return out


def net_positions(legs):
    """client_key -> {name, bought, sold} after feed/self-trade cleanup."""
    nets = {}
    for leg in _merge_exchanges(_dedupe_client_day(legs)):
        key = canonical_party(leg["client"])
        rec = nets.setdefault(key, {"name": leg["client"],
                                    "bought": 0.0, "sold": 0.0})
        if len(leg["client"]) > len(rec["name"]):
            rec["name"] = leg["client"]
        if leg["side"] == "BUY":
            rec["bought"] += leg["qty"]
        else:
            rec["sold"] += leg["qty"]
    return nets


def apply_adjustments(holders, nets, outstanding, mcap):
    """Mutate holder dicts; return extra rows for unmatched post-SHP buyers."""
    holders = holders or []
    used = set()
    extras = []
    for key, rec in (nets or {}).items():
        match = _best_holder(holders, rec["name"])
        if match is None:
            net = rec["bought"] - rec["sold"]
            if net != 0:
                extras.append(_new_tape_row(rec, outstanding, mcap))
            continue
        used.add(id(match))
        match["bought"] = match.get("bought", 0) + rec["bought"]
        match["sold"] = match.get("sold", 0) + rec["sold"]

    for h in holders + extras:
        _finish_holder(h, outstanding, mcap)
    return extras


def _new_tape_row(rec, outstanding, mcap):
    """Holder not named in the latest SHP (often under 1%) but on the tape."""
    net = rec["bought"] - rec["sold"]
    kind = "Post-SHP block/bulk buy" if net > 0 else "Post-SHP block/bulk sell"
    return {
        "name": rec["name"],
        "category": kind,
        "pct": 0.0,
        "shares": 0.0,
        "holding_cr": 0.0,
        "bought": rec["bought"],
        "sold": rec["sold"],
        "adj_shares": max(0.0, net),
        "adj_pct": None,
        "adj_holding_cr": None,
    }


def _finish_holder(h, outstanding, mcap):
    shares = _num(h.get("shares"))
    bought = _num(h.get("bought"))
    sold = _num(h.get("sold"))
    adj = shares + bought - sold
    if adj < 0:
        adj = 0.0
    h["bought"] = bought
    h["sold"] = sold
    h["adj_shares"] = adj
    if h.get("holding_cr") is None:
        h["holding_cr"] = _rupee_holding(h.get("pct"), mcap)
    # No post-SHP trade: keep BSE % and rupee holding as-is. Recomputing
    # % from implied outstanding was shifting the Rs cr column.
    if bought == 0 and sold == 0:
        h["adj_shares"] = shares
        h["adj_pct"] = h.get("pct")
        h["adj_holding_cr"] = h.get("holding_cr")
        return
    if outstanding:
        h["adj_pct"] = round(adj * 100.0 / outstanding, 4)
    else:
        h["adj_pct"] = None
    h["adj_holding_cr"] = _rupee_holding(h.get("adj_pct"), mcap)


def _rupee_holding(pct, mcap):
    if pct is None or mcap is None:
        return None
    try:
        return round(float(pct) * float(mcap) / 100.0, 2)
    except (TypeError, ValueError):
        return None


def _page_text(page, url):
    if page is None:
        return None
    try:
        return page.evaluate(_FETCH_TEXT, url)
    except Exception:
        return None


def _nse_row(row):
    return {(k or "").strip(): (v or "").strip() for k, v in (row or {}).items()}


def _nse_date(raw):
    raw = (raw or "").strip().strip('"')
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def parse_nse_csv(text, start, end, kind):
    """Parse NSE bulk/block CSV. Headers come quoted with trailing spaces."""
    if not text or text.lstrip().startswith("{") or text.lstrip().startswith("<"):
        return []
    text = text.lstrip("\ufeff")
    try:
        raw_rows = list(csv.DictReader(io.StringIO(text)))
    except Exception:
        return []
    legs = []
    for raw in raw_rows:
        row = _nse_row(raw)
        day = _nse_date(row.get("Date"))
        if day is None or not (start <= day <= end):
            continue
        side = "BUY" if (row.get("Buy / Sell") or "").upper().startswith("B") else "SELL"
        qty = _num(row.get("Quantity Traded"))
        if qty <= 0:
            continue
        legs.append({
            "exchange": "NSE",
            "deal_type": kind,
            "date": day,
            "client": row.get("Client Name") or "",
            "side": side,
            "qty": qty,
        })
    return legs


def fetch_nse_legs(npage, symbol, start, end):
    if not symbol or not start or not end:
        return []
    legs = []
    for option, kind in (("block_deals", "BLOCK"), ("bulk_deals", "BULK")):
        url = NSE_DEALS.format(
            option=option,
            start=start.strftime("%d-%m-%Y"),
            end=end.strftime("%d-%m-%Y"),
        ) + "&symbol=" + symbol
        res = _page_text(npage, url)
        text = (res or {}).get("text") or ""
        legs.extend(parse_nse_csv(text, start, end, kind))
    return legs


def fetch_bse_legs(bpage, scripcode, start, end):
    if not scripcode or not start or not end:
        return []
    import json
    legs = []
    for code, kind in ((2, "BLOCK"), (1, "BULK")):
        url = BSE_DEALS.format(
            deal_type=code,
            scrip=scripcode,
            start=start.strftime("%d/%m/%Y"),
            end=end.strftime("%d/%m/%Y"),
        )
        res = _page_text(bpage, url)
        text = (res or {}).get("text") or ""
        payload = None
        try:
            payload = json.loads(text) if text else None
        except Exception:
            payload = None
        rows = []
        if isinstance(payload, list):
            rows = payload
        elif isinstance(payload, dict):
            rows = payload.get("Table") or []
        for row in rows:
            raw = str(row.get("DEAL_DATE") or "")
            try:
                day = datetime.datetime.fromisoformat(raw[:10]).date()
            except ValueError:
                continue
            if not start <= day <= end:
                continue
            side = "BUY" if str(row.get("TRANSACTION_TYPE") or "").upper().startswith("P") else "SELL"
            qty = _num(row.get("QUANTITY"))
            if qty <= 0:
                continue
            legs.append({
                "exchange": "BSE",
                "deal_type": kind,
                "date": day,
                "client": str(row.get("CLIENT_NAME") or "").strip(),
                "side": side,
                "qty": qty,
            })
    return legs


def fetch_and_adjust(bpage, npage, nse_symbol, scripcode, qname, holders, mcap):
    """Apply post-SHP deals to holders. Returns (holders, extras, meta)."""
    start, end = adjust_window(qname)
    outstanding = implied_outstanding(holders)
    meta = {
        "quarter_end": start - datetime.timedelta(days=1) if start else None,
        "adjust_from": start,
        "adjust_to": end,
        "shares_outstanding": outstanding,
        "deal_legs": 0,
        "matched_clients": 0,
    }
    if not start or not end or start > end:
        for h in holders:
            _finish_holder(h, outstanding, mcap)
        return holders, [], meta

    legs = []
    try:
        legs.extend(fetch_nse_legs(npage, nse_symbol, start, end))
    except Exception:
        pass
    try:
        legs.extend(fetch_bse_legs(bpage, scripcode, start, end))
    except Exception:
        pass
    meta["deal_legs"] = len(legs)
    nets = net_positions(legs)
    extras = apply_adjustments(holders, nets, outstanding, mcap)
    meta["matched_clients"] = sum(
        1 for h in holders if _num(h.get("bought")) or _num(h.get("sold"))
    )
    if outstanding and not meta["shares_outstanding"]:
        meta["shares_outstanding"] = outstanding
    else:
        meta["shares_outstanding"] = implied_outstanding(holders) or outstanding
    return holders, extras, meta
