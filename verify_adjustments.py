"""Live reconciliation of SHP + post-quarter block/bulk adjustments.

Uses the same browser session as the app. For each company:
  - run the pipeline lookup
  - re-download NSE block/bulk CSV independently
  - check share arithmetic, no-trade identity, and that every
    material NSE client was applied to a holder or extra row
"""

from __future__ import annotations

import datetime
import sys

from playwright.sync_api import sync_playwright

from deals import (
    canonical_party,
    fetch_bse_legs,
    fetch_nse_legs,
    names_match,
    net_positions,
)
from pipeline import Pipeline, _open_session


COMPANIES = [
    "Aster DM",
    "Gland Pharma",
    "Pine Labs",
    "Urban Company",
    "Clean Max",
    "Ola Electric",
    "Waaree Energies",
    "Premier Energies",
    "FirstCry",
    "Honasa",
    "Go Digit",
    "Afcons",
    "Bajaj Housing",
    "Vishal Mega Mart",
    "Sagility",
    "BlackBuck",
    "HDFC Bank",
    "Bharti Airtel",
    "Adani Enterprises",
    "Hindustan Unilever",
    "Maruti Suzuki",
    "Bajaj Finance",
    "Avenue Supermarts",
    "HDFC AMC",
    "DLF",
    "Apollo Hospitals",
    "HAL",
]


def _holders(result):
    return list(result.get("promoters") or []) + list(result.get("public") or [])


def _find(holders, name):
    hits = [h for h in holders if names_match(h.get("name"), name)
            or canonical_party(h.get("name")) == canonical_party(name)]
    if hits:
        return hits
    needle = canonical_party(name)
    if not needle:
        return []
    out = []
    for h in holders:
        hk = canonical_party(h.get("name"))
        if not hk:
            continue
        if needle in hk.split() or (len(needle) >= 10 and needle in hk):
            out.append(h)
    return out


def check_company(ppl, bpage, npage, query, failures):
    print(f"\n======== {query} ========")
    result = ppl._lookup_on_pages(bpage, npage, query)
    if not result.get("found"):
        print("  NOT FOUND")
        failures.append(f"{query}: BSE lookup failed")
        return result

    holders = _holders(result)
    nse = result.get("nse_ticker")
    start = datetime.date.fromisoformat(result["adjust_from"]) if result.get("adjust_from") else None
    end = datetime.date.fromisoformat(result["adjust_to"]) if result.get("adjust_to") else None
    outst = result.get("shares_outstanding")
    mcap = result.get("mcap")
    print(f"  {result['name']}  NSE={nse}  BSE={result['scripcode']}  "
          f"qtr={result.get('quarter')} as-of={result.get('quarter_end')}")
    outst_s = f"{outst:,.0f}" if outst else "—"
    print(f"  window {result.get('adjust_from')} → {result.get('adjust_to')}  "
          f"outst={outst_s}  mcap={mcap}  holders={len(holders)}")

    # Independent NSE + BSE re-fetch (same window the tool uses)
    raw_legs = []
    if start and end:
        if nse:
            raw_legs.extend(fetch_nse_legs(npage, nse, start, end))
        raw_legs.extend(fetch_bse_legs(bpage, result.get("scripcode"), start, end))
    nets = net_positions(raw_legs)
    nse_n = sum(1 for x in raw_legs if x.get("exchange") == "NSE")
    bse_n = sum(1 for x in raw_legs if x.get("exchange") == "BSE")
    print(f"  independent legs NSE={nse_n} BSE={bse_n}  net clients={len(nets)}")

    # Arithmetic on every holder
    for h in holders:
        shp = float(h.get("shares") or 0)
        bought = float(h.get("bought") or 0)
        sold = float(h.get("sold") or 0)
        adj = float(h.get("adj_shares") if h.get("adj_shares") is not None else shp)
        expect = max(0.0, shp + bought - sold)
        if abs(adj - expect) > 1:
            failures.append(
                f"{query}: {h.get('name')} adj {adj} != {expect} "
                f"(shp {shp} +{bought} -{sold})"
            )
        if bought == 0 and sold == 0:
            if h.get("adj_pct") != h.get("pct"):
                failures.append(
                    f"{query}: {h.get('name')} no-trade but adj% "
                    f"{h.get('adj_pct')} != shp% {h.get('pct')}"
                )
            if h.get("adj_holding_cr") != h.get("holding_cr"):
                failures.append(
                    f"{query}: {h.get('name')} no-trade but holding "
                    f"{h.get('holding_cr')} vs adj {h.get('adj_holding_cr')}"
                )
        if mcap is not None and h.get("pct") is not None and h.get("holding_cr") is not None:
            expect_cr = round(float(h["pct"]) * float(mcap) / 100.0, 2)
            if abs(float(h["holding_cr"]) - expect_cr) > 0.05:
                failures.append(
                    f"{query}: {h.get('name')} holding_cr {h['holding_cr']} "
                    f"!= pct*mcap {expect_cr}"
                )

    # Every material NSE net must land on a holder or extra
    unmatched_sells = []
    for rec in nets.values():
        if rec["sold"] < 10000 and rec["bought"] < 10000:
            continue
        hits = _find(holders, rec["name"])
        if hits:
            h = hits[0]
            # sold/bought on the matched row should cover this client
            if rec["sold"] > 0 and float(h.get("sold") or 0) + 1 < rec["sold"]:
                # might be split across two similar names; sum hits
                total_sold = sum(float(x.get("sold") or 0) for x in hits)
                if total_sold + 1 < rec["sold"]:
                    unmatched_sells.append(
                        f"{rec['name']} tape sold {rec['sold']:,.0f} "
                        f"but holder sold {total_sold:,.0f}"
                    )
            continue
        extra = [h for h in holders
                 if (h.get("category") or "").startswith("Post-SHP")
                 and names_match(h.get("name"), rec["name"])]
        if extra:
            continue
        unmatched_sells.append(f"{rec['name']} bought {rec['bought']:,.0f} sold {rec['sold']:,.0f}")

    if unmatched_sells:
        for u in unmatched_sells:
            failures.append(f"{query}: tape client not applied: {u}")
            print(f"  MISS {u}")

    moved = [h for h in holders if h.get("bought") or h.get("sold")]
    print(f"  moved rows={len(moved)}")
    for h in moved[:20]:
        print(f"    {h.get('name')[:48]:48} shp={float(h.get('shares') or 0):12,.0f}  "
              f"b={float(h.get('bought') or 0):10,.0f}  s={float(h.get('sold') or 0):10,.0f}  "
              f"adj={float(h.get('adj_shares') or 0):12,.0f}  {h.get('adj_pct')}%")
    return result


def expect_zero(result, needle, failures, label):
    hits = _find(_holders(result), needle)
    if not hits:
        failures.append(f"{label}: no holder matching {needle}")
        return
    for h in hits:
        adj = float(h.get("adj_shares") or 0)
        sold = float(h.get("sold") or 0)
        shp = float(h.get("shares") or 0)
        if adj > 1:
            failures.append(
                f"{label}: {h.get('name')} expected full exit, adj={adj:,.0f} "
                f"(shp {shp:,.0f} sold {sold:,.0f})"
            )
        else:
            print(f"  OK full exit {h.get('name')} sold {sold:,.0f}")


def expect_sold_at_least(result, needle, min_sold, failures, label):
    hits = _find(_holders(result), needle)
    if not hits:
        failures.append(f"{label}: no holder matching {needle}")
        return
    sold = sum(float(h.get("sold") or 0) for h in hits)
    if sold + 1 < min_sold:
        failures.append(
            f"{label}: {needle} sold {sold:,.0f} < expected {min_sold:,.0f}"
        )
    else:
        print(f"  OK {hits[0].get('name')} sold {sold:,.0f} (>= {min_sold:,.0f})")


def expect_bought_at_least(result, needle, min_bought, failures, label):
    hits = _find(_holders(result), needle)
    if not hits:
        failures.append(f"{label}: no holder matching {needle}")
        return
    bought = sum(float(h.get("bought") or 0) for h in hits)
    if bought + 1 < min_bought:
        failures.append(
            f"{label}: {needle} bought {bought:,.0f} < expected {min_bought:,.0f}"
        )
    else:
        print(f"  OK {hits[0].get('name')} bought {bought:,.0f} (>= {min_bought:,.0f})")


def expect_resolved(result, query, must_include, failures):
    if not result.get("found"):
        failures.append(f"{query}: BSE lookup failed")
        return
    name = (result.get("name") or "")
    low = name.lower()
    if any(bad in low for bad in ("etf", "t+0")):
        failures.append(f"{query} resolved to {name}")
        return
    if not any(tok.lower() in low for tok in must_include):
        failures.append(f"{query} resolved to {name}")
        return
    print(f"  OK {query} → {name}")


def main():
    failures = []
    results = {}
    with sync_playwright() as p:
        browser, ctx, label = _open_session(p)
        print("browser", label)
        bpage = ctx.new_page()
        bpage.goto("https://www.bseindia.com/", wait_until="domcontentloaded")
        bpage.wait_for_timeout(1200)
        npage = ctx.new_page()
        try:
            npage.goto("https://www.nseindia.com/", wait_until="domcontentloaded")
            npage.wait_for_timeout(1800)
        except Exception as e:
            print("NSE goto failed", e)
        ppl = Pipeline()
        for name in COMPANIES:
            try:
                results[name] = check_company(ppl, bpage, npage, name, failures)
            except Exception as e:
                failures.append(f"{name}: exception {e}")
                print(f"  EXCEPTION {e}")
        ctx.close()
        browser.close()

    print("\n======== known-deal assertions ========")
    if results.get("Aster DM", {}).get("found"):
        expect_resolved(results["Aster DM"], "Aster DM",
                        ["aster"], failures)
        expect_sold_at_least(results["Aster DM"], "Centella Mauritius",
                             58100000, failures, "Aster Centella qty")

    if results.get("Gland Pharma", {}).get("found"):
        expect_sold_at_least(results["Gland Pharma"], "Fosun",
                             7500000, failures, "Gland Fosun qty")

    if results.get("Pine Labs", {}).get("found"):
        expect_sold_at_least(results["Pine Labs"], "Alpha Wave",
                             1, failures, "Pine Labs Alpha Wave")

    if results.get("Urban Company", {}).get("found"):
        expect_sold_at_least(results["Urban Company"], "Accel India",
                             20000000, failures, "Urban Accel qty")
        expect_bought_at_least(results["Urban Company"], "SBI Mutual Fund",
                               31500000, failures, "Urban SBI MF qty")

    if results.get("HDFC AMC", {}).get("found"):
        expect_resolved(results["HDFC AMC"], "HDFC AMC",
                        ["asset management", "amc"], failures)

    print("\n======== SUMMARY ========")
    if failures:
        print(f"FAILURES ({len(failures)})")
        for f in failures:
            print(" -", f)
        return 1
    print("ALL CHECKS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
