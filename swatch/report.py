# -*- coding: utf-8 -*-
"""Write the dashboard data (docs/data.js), events.csv and the Excel workbook."""
import datetime as dt
import json
import os
import shutil

import pandas as pd

from . import analyze, geo, lists
from . import sanctions_geo
from . import lexicon as L

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def _clean(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


def event_records(ev):
    out = []
    for e in ev:
        r = e["risk"]
        out.append({
            "i": e["id"], "d": e["date"], "ca": e["category"], "it": e["item"], "cg": e.get("cargo"),
            "rg": e.get("regimes"), "au": e.get("authorities"), "ag": e.get("agency"),
            "c": e.get("seizure_country"), "p": e.get("seizure_place"), "o": e.get("origin"), "op": e.get("origin_place"),
            "de": e.get("destination"), "dp": e.get("destination_place"), "t": e.get("transit"),
            "kg": e.get("qty_kg"), "bb": e.get("qty_barrels"), "u": e.get("qty_units"), "ut": e.get("unit_type"), "v": (e.get("value_usd") if (e.get("value_usd") or 0) <= 1e9 else None),
            "ve": e.get("vessel"), "imo": e.get("imo"), "fl": e.get("flag"), "tr": e.get("transport") if e.get("transport") != "Unknown" else None,
            "in": e.get("indicators"), "sc": r["score"], "lv": r["level"], "rs": r["reasons"],
            "ti": (e.get("title_en") or e.get("title") or "")[:200], "ur": e["urls"][0] if e.get("urls") else None,
            "sl": e.get("listed") or None, "ac": e.get("actor"), "lc": e.get("location_conf"), "ou": (e.get("outlets") or "")[:70], "ar": e.get("arrests"), "n": e.get("n"), "mo": e.get("mo_summary"),
        })
    return out


def geo_for(ev):
    """Coordinates for the countries and places that appear in the events (keeps data.js small)."""
    full = sanctions_geo.geo_payload()
    countries, places = {}, {}
    for e in ev:
        for k in ("seizure_country", "origin", "destination"):
            c = e.get(k)
            if c and c in full["c"]:
                countries[c] = full["c"][c]
        for c in (e.get("transit") or "").split("|"):
            if c and c in full["c"]:
                countries[c] = full["c"][c]
        for k in ("seizure_place", "origin_place", "destination_place"):
            p = e.get(k)
            if p and p in full["p"]:
                places[p] = full["p"][p]
    return {"c": countries, "p": places}


def export_table(ev):
    rows = []
    for e in ev:
        r = e["risk"]
        rows.append({
            "Date": e["date"], "What was seized": e["item"], "Category": e["category"], "Cargo aboard": e.get("cargo"),
            "Sanctions regime(s)": (e.get("regimes") or "").replace("|", "; "), "Authority": (e.get("authorities") or "").replace("|", "; "),
            "Seized by (agency)": e.get("agency"), "Seized by (country)": e.get("actor"), "Seized in (country)": e.get("seizure_country"), "Seized at (place)": e.get("seizure_place"),
            "Origin": e.get("origin"), "Origin place": e.get("origin_place"), "Transit": (e.get("transit") or "").replace("|", "; "),
            "Destination": e.get("destination"), "Destination place": e.get("destination_place"), "Route": e.get("route"),
            "Vessel": e.get("vessel"), "IMO": e.get("imo"), "Flag": e.get("flag"), "Container numbers": e.get("container_numbers"),
            "Quantity (kg)": e.get("qty_kg"), "Quantity (barrels)": e.get("qty_barrels"), "Quantity (units)": e.get("qty_units"), "Unit": e.get("unit_type"),
            "Stated value (USD, approx.)": e.get("value_usd"), "Arrests": e.get("arrests"),
            "Evasion score": r["score"], "Level": r["level"], "Why": "; ".join(r["reasons"]), "Indicators": "; ".join(r["indicator_labels"]),
            "Reports": e.get("n"), "Outlets": e.get("outlets"), "Headline": e.get("title_en") or e.get("title"), "Link": e["urls"][0] if e.get("urls") else None,
            "Verified?": "",
        })
    return pd.DataFrame(rows)


def write_excel(path, res, table):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        table.to_excel(xw, sheet_name="Events", index=False)
        pd.DataFrame(res["corridors"]).to_excel(xw, sheet_name="Corridors", index=False)
        pd.DataFrame(res["hotspots"]).to_excel(xw, sheet_name="Hotspots", index=False)
        pd.DataFrame(res["vessels"]).to_excel(xw, sheet_name="Vessels", index=False)
        pd.DataFrame([{k: v for k, v in a.items() if k != "urls"} | {"links": " ".join(a.get("urls", []))} for a in res["alerts"]]).to_excel(xw, sheet_name="Alerts", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col[:60] if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(48, max(10, width + 2))


def write_all(res, out_dir, cfg, demo=False):
    os.makedirs(out_dir, exist_ok=True)
    ev = res["events"]
    _idx = lists.annotate(ev, log=lambda m: print(m, flush=True))
    lists.write_catalog(_idx, out_dir, log=lambda m: print(m, flush=True))
    recs = event_records(ev)
    table = export_table(ev)
    table.to_csv(os.path.join(out_dir, "events.csv"), index=False, encoding="utf-8-sig")
    write_excel(os.path.join(out_dir, "sanctions_watch.xlsx"), res, table)
    data = {
        "asof": res["asof"].isoformat(), "generated": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "first": res["first_date"],
        "demo": bool(demo), "geo": geo_for(ev), "events": recs, "model": res["model"],
        "alerts": [{k: v for k, v in a.items()} for a in res["alerts"][:40]],
        "corridors": res["corridors"][:80], "vessels": [v for v in res["vessels"] if v["events"] >= 1][:80], "hotspots": res["hotspots"][:60],
        "indicators": analyze.LABEL,
    }
    open(os.path.join(out_dir, "data.js"), "w", encoding="utf-8").write(
        "window.SWD=" + json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=lambda x: None) + ";")
    # the page and its map data sit next to data.js so GitHub Pages serves them from /docs
    for f in ("index.html", "world.js", "registry.js", "goods.js", "pccp-logo-banner.jpg", "byte.js", "byte-config.js"):
        src = os.path.join(ROOT, "web", f)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(out_dir, f))
    hb = os.path.join(ROOT, "data", "Handbook_Strategic_Trade_Controls.pdf")
    if os.path.exists(hb):
        shutil.copy2(hb, os.path.join(out_dir, "Handbook_Strategic_Trade_Controls.pdf"))
    for ext in ("xlsx", "csv", "json"):
        xl = os.path.join(ROOT, "data", "Global_Sanctions_Source_Registry." + ext)
        if os.path.exists(xl):
            shutil.copy2(xl, os.path.join(out_dir, "Global_Sanctions_Source_Registry." + ext))
    return {"events": len(ev), "alerts": len(res["alerts"])}


def write_empty(out_dir, cfg, demo=False):
    """First run with nothing found yet: still publish the page so the GitHub Pages link works."""
    os.makedirs(out_dir, exist_ok=True)
    data = {"asof": dt.date.today().isoformat(), "generated": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
            "first": None, "demo": bool(demo), "geo": {"c": {}, "p": {}}, "events": [], "model": analyze.risk_model_description(cfg),
            "alerts": [], "corridors": [], "vessels": [], "hotspots": [], "indicators": analyze.LABEL}
    open(os.path.join(out_dir, "data.js"), "w", encoding="utf-8").write("window.SWD=" + json.dumps(data, ensure_ascii=False) + ";")
    for f in ("index.html", "world.js", "registry.js", "goods.js", "pccp-logo-banner.jpg", "byte.js", "byte-config.js"):
        src = os.path.join(ROOT, "web", f)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(out_dir, f))
    return {"events": 0, "alerts": 0}
