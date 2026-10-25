#!/usr/bin/env python3
"""
tradestat.py - download and label TradeStat (Dept. of Commerce, India) Excel reports.

Setup (once):
    pip install playwright openpyxl
    playwright install chromium

Commands:
    python tradestat.py inspect                 # dump the query form's fields to form_fields.json
    python tradestat.py assist  [-o data/raw]       # you submit queries by hand; every download is saved + labelled
    python tradestat.py auto    [-o data/raw]       # loop over FLOWS x YEARS using the form config below
    python tradestat.py label   folder [-o data/raw] [--copy]  # label xlsx files you already have (moved unless --copy)

Files land in ./data/raw/import and ./data/raw/export (by the flow read from each file);
data/raw/manifest.csv logs every file. Downloads pass through data/raw/_incoming first.

Labelled name:  tradestat_<flow>_<basis>_<period>.xlsx
    e.g.        tradestat_import_commodity-wise_FY2023-24_FY2024-25.xlsx
Flow, basis and years are read from the file itself (title cell + header row),
not from the original file name, so mislabelled downloads still come out right.
"""
import argparse, csv, filecmp, json, re, shutil, sys, time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]   # project root, so paths do not depend on the working directory
DEFAULT_OUT = str(ROOT / "data" / "raw")

EIDB = "https://tradestat.commerce.gov.in/eidb/"
START_URL = EIDB + "commodity_wise_import"   # the bare /eidb/ path returns 404

# ---------------------------------------------------------------- auto mode config
# Taken from the live commodity-wise forms (Oct 2026). The import and export forms
# name their fields differently, so each flow carries its own selectors.
FLOWS = {
    "import": dict(url=EIDB + "commodity_wise_import",
                   year="select[name='Eidb_YearCwi']",
                   hs_level="select[name='Eidb_ComLevelCwi']",
                   unit="select[name='Eidb_ReportCwi']"),
    "export": dict(url=EIDB + "commodity_wise_export",
                   year="select[name='EidbYearCwe']",
                   hs_level="select[name='EidbComLevelCwe']",
                   unit="select[name='Eidb_ReportCwe']"),
}
SUBMIT = "form button[type='submit']:visible"   # the form has a second, hidden submit button
DOWNLOAD = "button.buttons-excel"                # DataTables Excel button on the results page
YEARS = ["2023-2024", "2024-2025", "2025-2026"]  # dropdown labels; each report also includes the year before
HS_LEVEL = "8 digit Level"
UNIT = "US $ Million"

# ---------------------------------------------------------------- labelling
def _slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")

def read_meta(path) -> dict[str, Any]:
    """Pull flow, basis, years, HS digits, row count and totals out of a TradeStat xlsx.
    gap_y1/gap_y2 = reported total minus the sum of the commodity rows (should be ~0)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    title = str(next(rows)[0] or "")
    sub = str(next(rows)[0] or "")
    # title looks like: TradeStat->Eidb->Import->Commodity-wise
    parts = [p.strip() for p in title.split("->")]
    flow = _slug(parts[2]) if len(parts) > 2 else "unknown"
    basis = _slug(parts[3]) if len(parts) > 3 else "unknown"
    header, years, hs_len, n, total, sums = None, [], None, 0, [None, None], [0.0, 0.0]
    for r in rows:
        if header is None:
            if r and any(isinstance(c, str) and re.fullmatch(r"\d{4}\s*-\s*\d{4}", c.strip()) for c in r if c):
                header = r
                years = [c.strip().replace(" ", "") for c in r
                         if isinstance(c, str) and re.fullmatch(r"\d{4}\s*-\s*\d{4}", c.strip())]
            continue
        code = str(r[1]) if len(r) > 1 and r[1] is not None else ""
        if code.isdigit():
            n += 1
            hs_len = hs_len or len(code)
            for i, v in enumerate((r[3], r[5]) if len(r) > 5 else ()):
                if isinstance(v, (int, float)):
                    sums[i] += v
        elif code.lower().startswith("india's total"):
            total = [r[3], r[5]] if len(r) > 5 else total  # year-1 and year-2 value columns
    wb.close()
    m = re.search(r"Report Generated on:\s*([\d/]+)", sub)
    unit = "usd-million" if "Million" in sub else ("usd-billion" if "Billion" in sub else "unknown")
    gaps = [round(t - s, 2) if isinstance(t, (int, float)) else None for t, s in zip(total, sums)]
    return dict(flow=flow, basis=basis, years=years, hs_digits=hs_len, rows=n,
                total_y1=total[0], total_y2=total[1], gap_y1=gaps[0], gap_y2=gaps[1],
                generated=m.group(1) if m else "", unit=unit)

def check_totals(meta):
    """Warnings if the total row is missing or the commodity rows do not add up to it."""
    warns = []
    for y in ("y1", "y2"):
        t, g = meta[f"total_{y}"], meta[f"gap_{y}"]
        if g is None:
            warns.append(f"no India's Total value for {y} (truncated download?)")
        elif abs(g) > 0.001 * abs(t):  # rounding of ~10k rows gives ~1 unit; missing rows give far more
            warns.append(f"rows sum to {t - g:.2f} but total is {t:.2f} for {y} (gap {g:.2f})")
    return warns

def label_name(meta):
    yrs = "_".join("FY" + y.split("-")[0] + "-" + y.split("-")[1][-2:] for y in
                   (re.sub(r"\s", "", y) for y in meta["years"])) or "unknown-period"
    return f"tradestat_{meta['flow']}_{meta['basis']}_{yrs}.xlsx"

def label_file(src, base, manifest, move=False):
    """Label src and place it in <base>/<flow>/ (e.g. data/import, data/export)."""
    src, base = Path(src), Path(base)
    try:
        meta = read_meta(src)
    except Exception as e:
        print(f"  ! could not read {src.name}: {e}")
        return None
    outdir = base / (meta["flow"] if meta["flow"] in ("import", "export") else "other")
    outdir.mkdir(parents=True, exist_ok=True)
    dest = outdir / label_name(meta)
    if dest.exists() and dest.resolve() != src.resolve() and filecmp.cmp(src, dest, shallow=False):
        if move:
            src.unlink()
        manifest.append(dict(original=src.name, labelled=dest.name, **meta))
        print(f"  {src.name}  ==  {dest.name} (identical, skipped)")
        return dest
    k = 2
    while dest.exists() and dest.resolve() != src.resolve():
        dest = outdir / f"{label_name(meta)[:-5]}_v{k}.xlsx"; k += 1
    if dest.resolve() != src.resolve():
        (shutil.move if move else shutil.copy2)(src, dest)
    manifest.append(dict(original=src.name, labelled=dest.name, **meta))
    print(f"  {src.name}  ->  {dest.name}  ({meta['rows']} lines, hs{meta['hs_digits']})")
    for w in check_totals(meta):
        print(f"  ! {dest.name}: {w}")
    return dest

def write_manifest(manifest, outdir):
    """Merge new rows into manifest.csv, keeping one row per labelled file."""
    if not manifest:
        return
    p = Path(outdir) / "manifest.csv"
    rows = {}
    if p.exists():
        with open(p, newline="") as f:
            rows = {r["labelled"]: r for r in csv.DictReader(f)}
    for r in manifest:
        old = rows.get(r["labelled"])
        if old and r["original"] == r["labelled"]:  # re-labelling a labelled file: keep the real original name
            r = dict(r, original=old["original"])
        rows[r["labelled"]] = r
    fields = list(dict.fromkeys(k for r in rows.values() for k in r))
    with open(p, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows.values())
    print(f"manifest: {p}")

def cmd_label(a):
    out = Path(a.out)
    man = []
    for f in sorted(Path(a.folder).glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        label_file(f, out, man, move=not a.copy)
    write_manifest(man, out)

# ---------------------------------------------------------------- browser modes
def _pw():
    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except ImportError:
        sys.exit("pip install playwright && playwright install chromium")

def cmd_inspect(a):
    with _pw()() as p:
        b = p.chromium.launch(headless=False)
        pg = b.new_page()
        pg.goto(a.url or START_URL)
        input("Navigate to the exact query page you want to automate (e.g. Import > Commodity-wise), then press Enter... ")
        fields = pg.evaluate("""() => [...document.querySelectorAll('select,input,button,a[href*=xls],a[href*=download]')].map(e => ({
            tag: e.tagName, type: e.type||'', name: e.name||'', id: e.id||'', text: (e.innerText||e.value||'').trim().slice(0,60),
            href: e.href||'', options: e.tagName==='SELECT' ? [...e.options].map(o=>[o.value,o.text]) : undefined}))""")
        (ROOT / "data" / "form_fields.json").write_text(json.dumps({"url": pg.url, "fields": fields}, indent=2))
        print(f"saved data/form_fields.json ({len(fields)} controls) for {pg.url}")
        b.close()

def _save_download(dl, base, man):
    inc = Path(base) / "_incoming"
    inc.mkdir(parents=True, exist_ok=True)
    tmp = inc / (dl.suggested_filename or f"download_{int(time.time())}.xlsx")
    dl.save_as(tmp)
    print(f"downloaded {tmp.name}")
    if tmp.suffix.lower() == ".xlsx":
        label_file(tmp, base, man, move=True)
    else:
        print("  (not .xlsx - left unlabelled)")

def cmd_assist(a):
    raw, man = Path(a.out), []
    with _pw()() as p:
        b = p.chromium.launch(headless=False)
        ctx = b.new_context(accept_downloads=True)
        pg = ctx.new_page()
        ctx.on("page", lambda newp: newp.on("download", lambda d: _save_download(d, raw, man)))
        pg.on("download", lambda d: _save_download(d, raw, man))
        pg.goto(a.url or START_URL)
        print("Run your queries in the browser and click the Excel download each time.")
        print("Every download is saved and labelled automatically. Close the browser window (or Ctrl+C) when done.")
        try:
            pg.wait_for_event("close", timeout=0)
        except KeyboardInterrupt:
            print("\ninterrupted")
        except Exception:
            pass
        finally:
            write_manifest(man, raw)  # always log what was downloaded, even on Ctrl+C

def cmd_auto(a):
    raw, man, failed = Path(a.out), [], []
    try:
        with _pw()() as p:
            b = p.chromium.launch(headless=False)
            ctx = b.new_context(accept_downloads=True)
            pg = ctx.new_page()
            for flow, form in FLOWS.items():
                for y in YEARS:
                    try:
                        pg.goto(form["url"])
                        pg.select_option(form["year"], label=y)
                        pg.select_option(form["hs_level"], label=HS_LEVEL)
                        pg.select_option(form["unit"], label=UNIT)
                        pg.click(SUBMIT)
                        pg.wait_for_load_state("networkidle", timeout=180_000)  # 8-digit reports are slow
                        with pg.expect_download(timeout=120_000) as d:
                            pg.click(DOWNLOAD)
                        _save_download(d.value, raw, man)
                    except Exception as e:  # one bad query should not abort the whole run
                        print(f"  ! {flow} {y} failed: {e}")
                        failed.append((flow, y))
                    time.sleep(2)  # be polite to the server
            b.close()
    finally:
        write_manifest(man, raw)  # keep a record of whatever did download
        if failed:
            print("failed: " + ", ".join(f"{f} {y}" for f, y in failed))

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    for name, fn in [("inspect", cmd_inspect), ("assist", cmd_assist), ("auto", cmd_auto)]:
        s = sp.add_parser(name); s.add_argument("-o", "--out", default=DEFAULT_OUT); s.add_argument("--url"); s.set_defaults(fn=fn)
    s = sp.add_parser("label"); s.add_argument("folder"); s.add_argument("-o", "--out", default=DEFAULT_OUT); s.add_argument("--copy", action="store_true", help="keep the originals instead of moving them"); s.set_defaults(fn=cmd_label)
    a = ap.parse_args(); a.fn(a)