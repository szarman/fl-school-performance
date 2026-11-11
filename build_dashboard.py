"""Build the FL School Performance dashboard from the FLDOE School Grades file.

Usage:  py -3.11 -I build_dashboard.py [data/SchoolGrades26.xlsx] [--artifact PATH]
Output: FL School Dashboard.html (self-contained, opens offline)
        docs/index.html (same page, served by GitHub Pages)
        --artifact PATH also writes the body-only version for claude.ai Artifacts

Method
- Expected performance = OLS fit of "Percent of Total Possible Points" on
  "Percent of Economically Disadvantaged Students", fit separately per school
  type (elementary / middle / high / combination) because each type has its own
  components and grading scale.
- Points above expected = actual - expected (percentage points).
- Alternative/ESE center schools and schools with no points or no ED value
  (grade "I" incomplete, etc.) are dropped.
"""
import json
import statistics as st
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
_args = sys.argv[1:]
ARTIFACT_OUT = None
if "--artifact" in _args:
    i = _args.index("--artifact")
    ARTIFACT_OUT = Path(_args[i + 1])
    del _args[i:i + 2]
SRC = Path(_args[0]) if _args else HERE / "data" / "SchoolGrades26.xlsx"
OUT = HERE / "FL School Dashboard.html"
DOCS_OUT = HERE / "docs" / "index.html"

ED = "Percent of Economically Disadvantaged Students"
PP = "Percent of Total Possible Points"
COMPONENTS = [
    ("English Language Arts Achievement", "ELA achievement"),
    ("English Language Arts Learning Gains", "ELA learning gains"),
    ("English Language Arts Learning Gains of the Lowest 25%", "ELA gains, lowest 25%"),
    ("Mathematics Achievement", "Math achievement"),
    ("Mathematics Learning Gains", "Math learning gains"),
    ("Mathematics Learning Gains of the Lowest 25%", "Math gains, lowest 25%"),
    ("Science Achievement", "Science achievement"),
    ("Social Studies Achievement", "Social studies achievement"),
    ("Middle School Acceleration", "Middle school acceleration"),
    ("Graduation Rate 2024-25", "Graduation rate"),
    ("College and Career Acceleration 2024-25", "College & career acceleration"),
]
TYPES = {"01": "Elementary", "02": "Middle", "03": "High", "04": "Combination"}
# Prior-year files for the momentum view. Same layout; College & Career
# Acceleration is reported a year late, so its header carries the cohort year.
PRIOR = {2024: HERE / "data" / "SchoolGrades24.xlsx", 2025: HERE / "data" / "SchoolGrades25.xlsx"}
MOM = ["Mathematics Learning Gains of the Lowest 25%",
       "English Language Arts Learning Gains of the Lowest 25%",
       "College and Career Acceleration"]  # prefix; full header ends in the cohort year


def num(v):
    return v if isinstance(v, (int, float)) else None


def read_sheet(path):
    rows = list(openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
    i = next(i for i, r in enumerate(rows) if r[0] == "District Number")
    hdr = [str(c).strip() if c is not None else "" for c in rows[i]]
    return rows, [dict(zip(hdr, r)) for r in rows[i + 1:] if r[0]], hdr


def school_key(d):
    try:
        return int(d["District Number"]), int(d["School Number"])
    except (TypeError, ValueError):
        return str(d["District Number"]).strip(), str(d["School Number"]).strip()


def load_prior(path):
    """{school key: [math L25, ELA L25, C&C accel, % of points]} plus the C&C cohort label."""
    _, recs, hdr = read_sheet(path)
    cca = next(h for h in hdr if h.startswith(MOM[2]))
    vals = {school_key(d): [num(d[MOM[0]]), num(d[MOM[1]]), num(d[cca]), num(d[PP])] for d in recs}
    return vals, cca.rsplit(" ", 1)[-1]


def main():
    rows, recs, hdr = read_sheet(SRC)
    title = next((r[2] for r in rows[:3] if r[2]), "School Grades")
    prior, cohorts = {}, {}
    for y, p in PRIOR.items():
        if p.exists():
            prior[y], cohorts[y] = load_prior(p)
        else:
            print(f"  (no {p.name}: momentum view will be empty for {y})")
            prior[y], cohorts[y] = {}, ""

    keep, dropped = [], 0
    for d in recs:
        ed, pp = num(d[ED]), num(d[PP])
        if ed is None or pp is None or d["Alternative/ESE Center School"] != "N":
            dropped += 1
            continue
        keep.append(d)

    fits = {}
    for t in TYPES:
        xs = [d[ED] for d in keep if d["School Type"] == t]
        ys = [d[PP] for d in keep if d["School Type"] == t]
        b, a = st.linear_regression(xs, ys)
        fits[t] = {"a": round(a, 4), "b": round(b, 5), "r": round(st.correlation(xs, ys), 3), "n": len(xs)}

    out = []
    for d in keep:
        t = d["School Type"]
        exp = fits[t]["a"] + fits[t]["b"] * d[ED]
        out.append({
            "dn": d["District Name"].title(),
            "sn": d["School Number"],
            "s": d["School Name"],
            "t": t,
            "ed": round(d[ED], 1),
            "pp": d[PP],
            "ex": round(exp, 1),
            "res": round(d[PP] - exp, 1),
            "g": [d["Grade 2026"] or "", d["Grade 2025"] or "", d["Grade 2024"] or ""],
            "ch": d["Charter School"] == "YES",
            "t1": d["Title I"] == "YES",
            "c": [num(d[k]) for k, _ in COMPONENTS],
            "tp": num(d["Total Points Earned"]),
            "tc": num(d["Total Components"]),
            # [2024, 2025] x [math L25, ELA L25, C&C accel, % points]; null if not in that year's file
            "h": [prior[y].get(school_key(d)) for y in (2024, 2025)],
        })

    payload = {
        "title": title,
        "source": SRC.name,
        "dropped": dropped,
        "types": TYPES,
        "fits": fits,
        "components": [lbl for _, lbl in COMPONENTS],
        "cohorts": [cohorts[2024], cohorts[2025], "2024-25"],
        "rows": out,
    }
    page = TEMPLATE.replace("__DATA__", json.dumps(payload, separators=(",", ":")))
    # Local file: full document. --artifact PATH: body-only fragment for claude.ai
    # Artifacts (the publisher adds the doctype/head/body skeleton).
    full = DOC_OPEN + page.replace("__BODY_OPEN__", "</head>\n<body>\n") + DOC_CLOSE
    OUT.write_text(full, encoding="utf-8")
    DOCS_OUT.parent.mkdir(exist_ok=True)
    DOCS_OUT.write_text(full, encoding="utf-8")  # GitHub Pages (served from /docs)
    if ARTIFACT_OUT:
        ARTIFACT_OUT.write_text(page.replace("__BODY_OPEN__", ""), encoding="utf-8")
        print(f"wrote {ARTIFACT_OUT}")

    hp = [r for r in out if r["ed"] >= 75]
    beat = [r for r in hp if r["res"] >= 10]
    print(f"{title}: {len(recs)} schools, kept {len(out)}, dropped {dropped}")
    for t, f in fits.items():
        print(f"  {TYPES[t]:<12} n={f['n']:<5} r={f['r']:<7} expected = {f['a']:.1f} {f['b']:+.3f} x ED%")
    print(f"  high-poverty (>=75% ED): {len(hp)}; >=10 pts above expected: {len(beat)}")
    for k, y in enumerate((2024, 2025)):
        print(f"  matched in {y} file: {sum(1 for r in out if r['h'][k])} of {len(out)} (C&C cohort {cohorts[y]})")
    print(f"wrote {OUT} ({OUT.stat().st_size / 1024:.0f} KB)")


DOC_OPEN = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
"""
DOC_CLOSE = "</body>\n</html>\n"

TEMPLATE = r"""<title>FL School Performance</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb; --ink: #0b0b0b; --ink-2: #52514e; --muted: #898781;
  --grid: #e1e0d9; --axis: #c3c2b7; --border: rgba(11,11,11,0.10);
  --hi: #2a78d6; --other: #eb6834; --dot: #a9a8a0;
  --gA: #0d366b; --gB: #1c5cab; --gC: #2a78d6; --gD: #5598e7; --gF: #86b6ef;
  --wash: rgba(42,120,214,0.08);
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --page: #0d0d0d; --surface: #1a1a19; --ink: #f0efec; --ink-2: #c3c2b7; --muted: #898781;
    --grid: #2c2c2a; --axis: #383835; --border: rgba(255,255,255,0.10);
    --hi: #3987e5; --other: #d95926; --dot: #5f5e59;
    --gA: #b7d3f6; --gB: #86b6ef; --gC: #5598e7; --gD: #2a78d6; --gF: #184f95;
    --wash: rgba(57,135,229,0.12);
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0d0d0d; --surface: #1a1a19; --ink: #f0efec; --ink-2: #c3c2b7; --muted: #898781;
  --grid: #2c2c2a; --axis: #383835; --border: rgba(255,255,255,0.10);
  --hi: #3987e5; --other: #d95926; --dot: #5f5e59;
  --gA: #b7d3f6; --gB: #86b6ef; --gC: #5598e7; --gD: #2a78d6; --gF: #184f95;
  --wash: rgba(57,135,229,0.12);
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--page); color: var(--ink);
  font: 14px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif; }
main { max-width: 1180px; margin: 0 auto; padding-block: 24px 48px;
  padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); }
#csvBox { width: 100%; height: 140px; margin-top: 10px; font: 12px/1.4 ui-monospace, Consolas, monospace; color: var(--ink);
  background: var(--page); border: 1px solid var(--axis); border-radius: 6px; padding: 8px; }
button:focus-visible, select:focus-visible, input:focus-visible { outline: 2px solid var(--hi); outline-offset: 2px; }
h1 { font-size: 22px; margin: 0 0 4px; font-weight: 650; }
h2 { font-size: 15px; margin: 0 0 2px; font-weight: 600; }
.sub { color: var(--ink-2); margin: 0 0 18px; }
.card { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px; margin-bottom: 16px; }
.card .sub { margin-bottom: 12px; font-size: 13px; }
.filters { display: flex; flex-wrap: wrap; gap: 12px 18px; align-items: end; }
.filters label { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-2); }
select, input[type=search] { font: inherit; color: var(--ink); background: var(--surface); border: 1px solid var(--axis);
  border-radius: 6px; padding: 5px 8px; min-width: 140px; }
input[type=range] { width: 160px; accent-color: var(--hi); }
.rv { color: var(--ink); font-weight: 600; font-variant-numeric: tabular-nums; }
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; margin-bottom: 16px; }
.tile { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; }
.tile .l { color: var(--ink-2); font-size: 12px; }
.tile .v { font-size: 26px; font-weight: 650; margin-top: 2px; }
.tile .d { color: var(--muted); font-size: 12px; }
.two { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
@media (max-width: 860px) { .two { grid-template-columns: 1fr; } }
.two .card { margin-bottom: 0; }
.legend { display: flex; flex-wrap: wrap; gap: 14px; font-size: 12px; color: var(--ink-2); margin-bottom: 10px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; }
.sw { width: 10px; height: 10px; border-radius: 3px; display: inline-block; }
.swdot { width: 9px; height: 9px; border-radius: 50%; display: inline-block; }
#scatter { position: relative; }
#scatter svg { display: block; }
svg text { fill: var(--muted); font-size: 11px; }
.gl { stroke: var(--grid); stroke-width: 1; }
.ax { stroke: var(--axis); stroke-width: 1; }
.pt { fill: var(--dot); }
.pt-hi { fill: var(--hi); stroke: var(--surface); stroke-width: 2; }
.fit { stroke: var(--ink-2); stroke-width: 2; fill: none; }
.fit-band { stroke: var(--hi); stroke-width: 1; fill: none; opacity: .6; }
.thr { stroke: var(--muted); stroke-width: 1; }
.hov { fill: none; stroke: var(--ink); stroke-width: 2; }
.lbl-ink { fill: var(--ink-2); font-size: 12px; stroke: var(--surface); stroke-width: 4px; paint-order: stroke; stroke-linejoin: round; }
.tip { position: fixed; pointer-events: none; z-index: 10; background: var(--surface); color: var(--ink);
  border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 12px; line-height: 1.4;
  box-shadow: 0 4px 16px rgba(0,0,0,.12); max-width: 280px; display: none; }
.tip b { font-weight: 600; }
.tip .k { color: var(--ink-2); }
.brow { display: grid; grid-template-columns: 92px 1fr; gap: 10px; align-items: center; margin: 6px 0; }
.brow .bl { font-size: 12px; color: var(--ink-2); text-align: right; font-variant-numeric: tabular-nums; }
.brow .bl small { display: block; color: var(--muted); font-size: 11px; }
.stack { display: flex; gap: 2px; height: 24px; }
.seg { height: 100%; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 600; min-width: 0; }
.seg:first-child { border-radius: 4px 0 0 4px; } .seg:last-child { border-radius: 0 4px 4px 0; }
.seg:only-child { border-radius: 4px; }
.crow { display: grid; grid-template-columns: 150px 1fr; gap: 10px; align-items: center; margin: 8px 0; }
.crow .cl { font-size: 12px; color: var(--ink-2); text-align: right; }
.cbars { display: flex; flex-direction: column; gap: 2px; }
.cbar { display: flex; align-items: center; gap: 6px; height: 10px; }
.cbar i { display: block; height: 10px; border-radius: 0 4px 4px 0; }
.cbar em { font-style: normal; font-size: 11px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.tbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: space-between; margin-bottom: 10px; }
.tbar .tl { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
button { font: inherit; color: var(--ink); background: var(--surface); border: 1px solid var(--axis); border-radius: 6px; padding: 5px 12px; cursor: pointer; }
button:hover { background: var(--wash); }
.tw { max-height: 560px; overflow: auto; border: 1px solid var(--border); border-radius: 8px; }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th, td { padding: 6px 10px; text-align: left; border-bottom: 1px solid var(--grid); white-space: nowrap; }
th { position: sticky; top: 0; background: var(--surface); font-weight: 600; font-size: 12px; color: var(--ink-2); cursor: pointer; user-select: none; }
th.num, td.num { text-align: right; font-variant-numeric: tabular-nums; }
th .ar { color: var(--muted); }
tbody tr:hover { background: var(--wash); }
td.school { white-space: normal; min-width: 220px; }
.pos { color: var(--ink); font-weight: 600; }
.notes { color: var(--ink-2); font-size: 13px; }
.notes li { margin: 4px 0; }
.count { color: var(--ink-2); font-size: 12px; }
.mtiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 10px; margin-bottom: 14px; }
.mtiles .tile { padding: 10px 12px; }
.mtiles .tile .v { font-size: 20px; }
.mgrid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; margin-bottom: 14px; }
@media (max-width: 860px) { .mgrid { grid-template-columns: 1fr; } }
.mp h3 { font-size: 13px; font-weight: 600; margin: 0; }
.mp .ms { color: var(--ink-2); font-size: 12px; margin: 2px 0 6px; min-height: 34px; }
.mp svg { display: block; }
.mline { fill: none; stroke-width: 2; stroke-linejoin: round; }
.mdot { stroke: var(--surface); stroke-width: 2; }
.mhit { fill: transparent; cursor: default; }
th.imp, td.imp { white-space: normal; min-width: 170px; }
</style>
__BODY_OPEN__<main>
  <h1>Florida school performance vs. poverty</h1>
  <p class="sub" id="subtitle"></p>

  <div class="card">
    <div class="filters">
      <label>School type <select id="fType"></select></label>
      <label>District <select id="fDist"></select></label>
      <label>Charter <select id="fCh"><option value="ALL">All</option><option value="Y">Charter only</option><option value="N">District-run only</option></select></label>
      <label>Title I <select id="fT1"><option value="ALL">All</option><option value="Y">Title I only</option><option value="N">Not Title I</option></select></label>
      <label>High poverty = at least <span><span class="rv" id="povV"></span>% econ. disadvantaged</span>
        <input type="range" id="fPov" min="50" max="100" step="5" value="75"></label>
      <label>Beating the odds = at least <span><span class="rv" id="resV"></span> pts above expected</span>
        <input type="range" id="fRes" min="0" max="25" step="1" value="10"></label>
    </div>
  </div>

  <div class="tiles" id="tiles"></div>

  <div class="card">
    <h2>Poverty vs. share of possible points earned</h2>
    <p class="sub" id="scSub"></p>
    <div class="legend" id="scLegend"></div>
    <div id="scatter"></div>
  </div>

  <div class="two" style="margin-bottom:16px">
    <div class="card">
      <h2>2026 grade mix by poverty level</h2>
      <p class="sub">Share of schools at each grade, by % economically disadvantaged. 100% is its own row (see notes).</p>
      <div class="legend" id="gLegend"></div>
      <div id="gbars"></div>
    </div>
    <div class="card">
      <h2 id="cTitle"></h2>
      <p class="sub" id="cSub"></p>
      <select id="cMode" aria-label="Groups to compare" style="margin-bottom:12px;max-width:100%">
        <option value="beat">Beating the odds vs. other high-poverty</option>
        <option value="t1a">Title I A schools vs. other A schools</option>
      </select>
      <div class="legend" id="cLegend"></div>
      <div id="cbars"></div>
    </div>
  </div>

  <div class="card">
    <h2>Momentum toward an A: Title I vs. other schools</h2>
    <p class="sub" id="mSub"></p>
    <div class="mtiles" id="mTiles"></div>
    <div class="legend" id="mLegend"></div>
    <div class="mgrid" id="mPanels"></div>
    <div class="tbar">
      <div class="tl">
        <h2 style="margin:0;font-size:14px">Title I schools</h2>
        <select id="mMode">
          <option value="near">Not yet A, closest first</option>
          <option value="a">Earned an A in 2026</option>
          <option value="all">All Title I</option>
        </select>
        <span class="count" id="mCount"></span>
      </div>
    </div>
    <p class="count" style="margin:0 0 10px"><b>Points needed</b> = component points short of an A; one point is one percentage point on any one component.
      <b>Where to improve</b> = the three components furthest below the average Title I school of the same type that earned an A, with the school's score and its gap. L25 = lowest 25% of students; C&C = college & career. Δ = change 2024 → 2026, in points.</p>
    <div class="tw" style="max-height:420px"><table><thead id="mHead"></thead><tbody id="mBody"></tbody></table></div>
  </div>

  <div class="card">
    <div class="tbar">
      <div class="tl">
        <h2 style="margin:0">Schools</h2>
        <select id="tMode">
          <option value="beat">Beating the odds</option>
          <option value="hp">All high-poverty</option>
          <option value="all">All schools in filter</option>
        </select>
        <input type="search" id="tSearch" placeholder="Search school or district">
        <span class="count" id="tCount"></span>
      </div>
      <button id="csv">Copy table as CSV</button>
    </div>
    <textarea id="csvBox" readonly hidden aria-label="CSV of the table"></textarea>
    <div class="tw"><table><thead id="thead"></thead><tbody id="tbody"></tbody></table></div>
  </div>

  <div class="card notes">
    <h2 style="color:var(--ink)">Method and caveats</h2>
    <ul id="notes"></ul>
  </div>
</main>
<div class="tip" id="tip"></div>

<script>
const D = __DATA__;
const $ = id => document.getElementById(id);
const TYPES = D.types;
const GRADES = ["A","B","C","D","F"];
const S = { type: "01", dist: "ALL", ch: "ALL", t1: "ALL", pov: 75, res: 10, mode: "beat", q: "", sort: "res", dir: -1, mmode: "near", msort: "need", mdir: 1, cmode: "beat" };
let cur = [];

const fmt1 = v => v == null ? "–" : (Math.round(v * 10) / 10).toFixed(1);
const sgn = v => (v > 0 ? "+" : v < 0 ? "−" : "") + Math.abs(v).toFixed(1);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const cssv = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();

// ---------- tooltip ----------
const tip = $("tip");
function showTip(html, x, y) {
  tip.innerHTML = html; tip.style.display = "block";
  const r = tip.getBoundingClientRect();
  let left = x + 14, top = y + 14;
  if (left + r.width > innerWidth - 8) left = x - r.width - 14;
  if (top + r.height > innerHeight - 8) top = y - r.height - 14;
  tip.style.left = Math.max(8, left) + "px"; tip.style.top = Math.max(8, top) + "px";
}
const hideTip = () => { tip.style.display = "none"; };

// ---------- filters ----------
function initFilters() {
  $("subtitle").textContent = `${D.title} · Florida Department of Education · ${D.rows.length.toLocaleString()} graded schools`;
  $("fType").innerHTML = Object.entries(TYPES).map(([k, v]) => `<option value="${k}">${v}</option>`).join("") + `<option value="ALL">All types</option>`;
  const dists = [...new Set(D.rows.map(r => r.dn))].sort();
  $("fDist").innerHTML = `<option value="ALL">All districts</option>` + dists.map(d => `<option>${esc(d)}</option>`).join("");
  const bind = (id, key, ev = "change", cast = v => v) => $(id).addEventListener(ev, e => { S[key] = cast(e.target.value); render(); });
  bind("fType", "type"); bind("fDist", "dist"); bind("fCh", "ch"); bind("fT1", "t1");
  bind("fPov", "pov", "input", Number); bind("fRes", "res", "input", Number);
  bind("tMode", "mode"); bind("tSearch", "q", "input", v => v.trim().toLowerCase());
  $("mMode").addEventListener("change", e => { S.mmode = e.target.value; drawMomTable(); });
  $("cMode").addEventListener("change", e => { S.cmode = e.target.value; drawComponents(); });
  $("csv").addEventListener("click", downloadCsv);
  addEventListener("resize", () => { drawScatter(); drawMomentum(); });
}

function filtered() {
  return D.rows.filter(r =>
    (S.type === "ALL" || r.t === S.type) &&
    (S.dist === "ALL" || r.dn === S.dist) &&
    (S.ch === "ALL" || r.ch === (S.ch === "Y")) &&
    (S.t1 === "ALL" || r.t1 === (S.t1 === "Y")));
}
const isHP = r => r.ed >= S.pov;
const isBeat = r => isHP(r) && r.res >= S.res;

function pearson(rows) {
  const n = rows.length; if (n < 3) return null;
  let sx = 0, sy = 0; rows.forEach(r => { sx += r.ed; sy += r.pp; });
  const mx = sx / n, my = sy / n; let num = 0, dx = 0, dy = 0;
  rows.forEach(r => { const a = r.ed - mx, b = r.pp - my; num += a * b; dx += a * a; dy += b * b; });
  return dx && dy ? num / Math.sqrt(dx * dy) : null;
}

// ---------- tiles ----------
function drawTiles() {
  const hp = cur.filter(isHP), bt = cur.filter(isBeat), r = pearson(cur);
  const tiles = [
    ["Schools shown", cur.length.toLocaleString(), S.type === "ALL" ? "All school types" : TYPES[S.type] + " schools"],
    ["Poverty–performance correlation", r == null ? "–" : r.toFixed(2), "Pearson r; −1 = poverty fully predicts lower scores"],
    ["High-poverty schools", hp.length.toLocaleString(), `≥ ${S.pov}% economically disadvantaged`],
    ["Beating the odds", bt.length.toLocaleString(), hp.length ? `${Math.round(100 * bt.length / hp.length)}% of high-poverty schools` : "No high-poverty schools in filter"],
  ];
  $("tiles").innerHTML = tiles.map(([l, v, d]) => `<div class="tile"><div class="l">${l}</div><div class="v">${v}</div><div class="d">${d}</div></div>`).join("");
}

// ---------- scatter ----------
let scPts = [];
function drawScatter() {
  const el = $("scatter");
  const W = el.clientWidth || 800, H = Math.round(Math.min(540, Math.max(320, W * 0.52)));
  const m = { l: 46, r: 14, t: 18, b: 42 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b;
  const X = v => m.l + v / 100 * iw, Y = v => m.t + ih - v / 100 * ih;
  let s = `<svg width="${W}" height="${H}" role="img" aria-label="Scatter plot of poverty versus points earned">`;
  for (let v = 0; v <= 100; v += 20) {
    s += `<line class="gl" x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}"/>`;
    s += `<text x="${m.l - 8}" y="${Y(v) + 4}" text-anchor="end">${v}%</text>`;
    s += `<text x="${X(v)}" y="${H - m.b + 18}" text-anchor="middle">${v}%</text>`;
  }
  s += `<line class="ax" x1="${m.l}" x2="${W - m.r}" y1="${Y(0)}" y2="${Y(0)}"/>`;
  s += `<text x="${m.l + iw / 2}" y="${H - 6}" text-anchor="middle">% economically disadvantaged students</text>`;
  s += `<text transform="translate(12 ${m.t + ih / 2}) rotate(-90)" text-anchor="middle">% of possible points earned</text>`;
  // poverty threshold
  s += `<line class="thr" x1="${X(S.pov)}" x2="${X(S.pov)}" y1="${m.t}" y2="${Y(0)}"/>`;
  s += `<text x="${X(S.pov) - 6}" y="${m.t + 10}" text-anchor="end">high poverty →</text>`;
  // expected line (single type only)
  if (S.type !== "ALL") {
    const f = D.fits[S.type];
    s += `<line class="fit-band" x1="${X(S.pov)}" y1="${Y(f.a + f.b * S.pov + S.res)}" x2="${X(100)}" y2="${Y(f.a + f.b * 100 + S.res)}"/>`;
    s += `<line class="fit" x1="${X(0)}" y1="${Y(f.a)}" x2="${X(100)}" y2="${Y(f.a + f.b * 100)}"/>`;
    s += `<text class="lbl-ink" x="${X(3)}" y="${Y(f.a) - 8}">Expected for poverty level</text>`;
  }
  scPts = cur.map(r => ({ r, x: X(r.ed), y: Y(r.pp), hi: isBeat(r) }));
  scPts.filter(p => !p.hi).forEach(p => { s += `<circle class="pt" cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="3" opacity=".55"/>`; });
  scPts.filter(p => p.hi).forEach(p => { s += `<circle class="pt-hi" cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4.5"/>`; });
  s += `<circle class="hov" id="hov" r="7" style="display:none"/></svg>`;
  el.innerHTML = s;
  const svg = el.querySelector("svg"), hov = el.querySelector("#hov");
  svg.addEventListener("mousemove", e => {
    const b = svg.getBoundingClientRect(), mx = e.clientX - b.left, my = e.clientY - b.top;
    let best = null, bd = 144;
    for (const p of scPts) { const d = (p.x - mx) ** 2 + (p.y - my) ** 2; if (d < bd || (d === bd && p.hi)) { bd = d; best = p; } }
    if (!best) { hov.style.display = "none"; hideTip(); return; }
    hov.setAttribute("cx", best.x); hov.setAttribute("cy", best.y); hov.style.display = "";
    showTip(schoolTip(best.r), e.clientX, e.clientY);
  });
  svg.addEventListener("mouseleave", () => { hov.style.display = "none"; hideTip(); });
  const sub = S.type === "ALL"
    ? "Each dot is a school. Pick a single school type to see the expected-performance line (each type is scored on different components)."
    : `Each dot is a school. The gray line is what ${S.type === "01" ? "an" : "a"} ${TYPES[S.type].toLowerCase()} school typically earns at that poverty level (r = ${D.fits[S.type].r}); the thin blue line marks ${S.res} points above it.`;
  $("scSub").textContent = sub;
  $("scLegend").innerHTML = `<span><i class="swdot" style="background:var(--hi)"></i>Beating the odds (${scPts.filter(p => p.hi).length})</span><span><i class="swdot" style="background:var(--dot)"></i>Other schools</span>`;
}
function schoolTip(r) {
  return `<b>${esc(r.s)}</b><br><span class="k">${esc(r.dn)} · ${TYPES[r.t]}${r.ch ? " · Charter" : ""}${r.t1 ? " · Title I" : ""}</span><br>
    Econ. disadvantaged: <b>${fmt1(r.ed)}%</b><br>Points earned: <b>${r.pp}%</b> (expected ${fmt1(r.ex)}%)<br>
    vs. expected: <b>${sgn(r.res)} pts</b><br>Grades 2026 / 25 / 24: <b>${r.g.map(g => g || "–").join(" / ")}</b>`;
}

// ---------- grade mix by poverty band ----------
const BANDS = [[0, 20, "0–19%"], [20, 40, "20–39%"], [40, 60, "40–59%"], [60, 80, "60–79%"], [80, 99.95, "80–99%"], [99.95, 101, "100%"]];
function lum(hex) {
  const c = hex.replace("#", "").match(/../g).map(h => { const v = parseInt(h, 16) / 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; });
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
}
function labelInk(hex) { const L = lum(hex); return (1.05 / (L + 0.05)) >= ((L + 0.05) / 0.0553) ? "#ffffff" : "#0b0b0b"; }
function drawGrades() {
  $("gLegend").innerHTML = GRADES.map(g => `<span><i class="sw" style="background:var(--g${g})"></i>${g}</span>`).join("");
  const host = $("gbars"); const W = host.clientWidth - 102;
  host.innerHTML = BANDS.map(([lo, hi, lbl]) => {
    const rs = cur.filter(r => r.ed >= lo && r.ed < hi && GRADES.includes(r.g[0]));
    if (!rs.length) return `<div class="brow"><div class="bl">${lbl}<small>0 schools</small></div><div class="count">No schools</div></div>`;
    const segs = GRADES.map(g => {
      const n = rs.filter(r => r.g[0] === g).length; if (!n) return "";
      const pct = 100 * n / rs.length, col = cssv("--g" + g), fits = W * pct / 100 >= 34;
      return `<div class="seg" style="flex:${n} 1 0;background:${col};color:${labelInk(col)}" data-tip="${lbl} economically disadvantaged<br><b>${g}: ${Math.round(pct)}%</b> (${n} of ${rs.length} schools)">${fits ? Math.round(pct) + "%" : ""}</div>`;
    }).join("");
    return `<div class="brow"><div class="bl">${lbl}<small>${rs.length} schools</small></div><div class="stack">${segs}</div></div>`;
  }).join("");
  host.querySelectorAll(".seg").forEach(el => {
    el.addEventListener("mousemove", e => showTip(el.dataset.tip, e.clientX, e.clientY));
    el.addEventListener("mouseleave", hideTip);
  });
}

// ---------- component comparison ----------
function drawComponents() {
  let bt, ot, nA, nB;
  if (S.cmode === "t1a") {
    const as = momBase().filter(r => r.g[0] === "A");
    bt = as.filter(r => r.t1); ot = as.filter(r => !r.t1); nA = "Title I, graded A"; nB = "Not Title I, graded A";
    $("cTitle").textContent = "How Title I A schools compare with other A schools";
    $("cSub").textContent = "Average component score (% of students), schools graded A in 2026. Ignores the Title I and poverty settings.";
  } else {
    const hp = cur.filter(isHP);
    bt = hp.filter(isBeat); ot = hp.filter(r => !isBeat(r)); nA = "Beating the odds"; nB = "Other high-poverty";
    $("cTitle").textContent = "Where the beating-the-odds schools pull ahead";
    $("cSub").textContent = "Average component score (% of students), high-poverty schools only.";
  }
  $("cLegend").innerHTML = `<span><i class="sw" style="background:var(--hi)"></i>${nA} (${bt.length})</span><span><i class="sw" style="background:var(--other)"></i>${nB} (${ot.length})</span>`;
  const avg = (rs, i) => { const v = rs.map(r => r.c[i]).filter(x => x != null); return v.length >= 5 ? { m: v.reduce((a, b) => a + b, 0) / v.length, n: v.length } : null; };
  const rows = D.components.map((lbl, i) => ({ lbl, a: avg(bt, i), b: avg(ot, i) })).filter(x => x.a && x.b);
  if (!rows.length) { $("cbars").innerHTML = `<p class="count">Not enough schools in one of the two groups for this filter (need at least 5 per component).</p>`; return; }
  rows.sort((x, y) => (y.a.m - y.b.m) - (x.a.m - x.b.m));
  $("cbars").innerHTML = rows.map(x => {
    const gap = x.a.m - x.b.m;
    const t = `<b>${x.lbl}</b><br>${nA}: ${x.a.m.toFixed(1)}% (${x.a.n} schools)<br>${nB}: ${x.b.m.toFixed(1)}% (${x.b.n} schools)<br>Gap: ${sgn(gap)} pts`;
    return `<div class="crow" data-tip="${esc(t)}"><div class="cl">${x.lbl}<br><span style="color:var(--muted)">gap ${sgn(gap)}</span></div><div class="cbars">
      <div class="cbar"><i style="width:calc(${x.a.m}% - 40px);background:var(--hi)"></i><em>${Math.round(x.a.m)}</em></div>
      <div class="cbar"><i style="width:calc(${x.b.m}% - 40px);background:var(--other)"></i><em>${Math.round(x.b.m)}</em></div></div></div>`;
  }).join("");
  $("cbars").querySelectorAll(".crow").forEach(el => {
    el.addEventListener("mousemove", e => showTip(el.dataset.tip, e.clientX, e.clientY));
    el.addEventListener("mouseleave", hideTip);
  });
}

// ---------- momentum toward an A ----------
// Year index 0/1/2 = 2024/2025/2026. Measure j: 0 math L25, 1 ELA L25, 2 C&C acceleration, 3 % of points.
const YRS = ["2024", "2025", "2026"];
const AMIN = { "01": 62, "02": 64, "03": 65, "04": 65 };
const C26 = [5, 2, 10];
const MEAS = ["Math gains, lowest 25%", "ELA gains, lowest 25%", "College & career acceleration"];
const mv = (r, j, y) => y === 2 ? (j === 3 ? r.pp : r.c[C26[j]]) : (r.h[y] ? r.h[y][j] : null);
const gy = (r, y) => r.g[2 - y];
const toA = r => Math.max(0, AMIN[r.t] - r.pp);
// FLDOE rounds % of points half up, so an A needs (cutoff - 0.5)% of the possible points.
const need = r => r.tp == null || !r.tc ? null : Math.max(0, Math.ceil((AMIN[r.t] - 0.5) * r.tc) - r.tp);
// Average component scores of Title I schools graded A in 2026, per school type: the peer benchmark.
const REFA = {};
Object.keys(D.types).forEach(t => {
  const rs = D.rows.filter(r => r.t === t && r.t1 && r.g[0] === "A");
  REFA[t] = D.components.map((_, i) => { const v = rs.map(r => r.c[i]).filter(x => x != null); return v.length >= 5 ? v.reduce((a, b) => a + b, 0) / v.length : null; });
});
const SHORT = ["ELA ach.", "ELA gains", "ELA gains L25", "Math ach.", "Math gains", "Math gains L25", "Science", "Social studies", "MS accel.", "Grad. rate", "C&C accel."];
const improve = r => D.components.map((_, i) => r.c[i] != null && REFA[r.t][i] != null ? { lbl: SHORT[i], v: r.c[i], g: REFA[r.t][i] - r.c[i] } : null)
  .filter(x => x && x.g >= 0.5).sort((a, b) => b.g - a.g).slice(0, 3);
const dlt = (r, j) => { const a = mv(r, j, 0), b = mv(r, j, 2); return a == null || b == null ? null : b - a; };
const MSER = [
  { k: "t1", name: "Title I", col: "--hi" },
  { k: "nt", name: "Not Title I", col: "--other" },
  { k: "a", name: "A schools (statewide)", col: "--ink-2", dash: true },
];
function momBase() {
  return D.rows.filter(r => (S.type === "ALL" || r.t === S.type) && (S.dist === "ALL" || r.dn === S.dist) && (S.ch === "ALL" || r.ch === (S.ch === "Y")));
}
function drawMomentum() {
  const base = momBase(), t1 = base.filter(r => r.t1), nt = base.filter(r => !r.t1);
  const ref = D.rows.filter(r => S.type === "ALL" || r.t === S.type);
  const typ = S.type === "ALL" ? "schools" : TYPES[S.type].toLowerCase() + " schools";
  $("mSub").textContent = `Same schools each year (only those with a score in all three years), using 2026 Title I status. The dashed line is the average of every ${typ.replace(/s$/, "")} statewide that earned an A that year: the bar to clear. Follows the type, district and charter filters above.`;

  // tiles
  const aN = y => t1.filter(r => gy(r, y) === "A").length;
  const gN = y => t1.filter(r => GRADES.includes(gy(r, y))).length;
  const pc = y => gN(y) ? Math.round(100 * aN(y) / gN(y)) + "%" : "–";
  const rose = t1.filter(r => gy(r, 2) === "A" && GRADES.includes(gy(r, 0)) && gy(r, 0) !== "A").length;
  const near = t1.filter(r => GRADES.includes(gy(r, 2)) && gy(r, 2) !== "A" && toA(r) <= 5).length;
  const tiles = [
    [`Title I ${typ} with an A`, `${aN(0)} → ${aN(1)} → ${aN(2)}`, "2024 → 2025 → 2026"],
    ["Share of graded Title I with an A", `${pc(0)} → ${pc(1)} → ${pc(2)}`, `${t1.length.toLocaleString()} Title I ${typ} in filter`],
    ["Rose to an A since 2024", rose.toLocaleString(), "B or lower in 2024, A in 2026"],
    ["Within 5 points of an A", near.toLocaleString(), "Not an A in 2026; ≤ 5 pts below the A cutoff"],
  ];
  $("mTiles").innerHTML = tiles.map(([l, v, d]) => `<div class="tile"><div class="l">${l}</div><div class="v">${v}</div><div class="d">${d}</div></div>`).join("");
  $("mLegend").innerHTML = MSER.map(s => `<span><i class="sw" style="background:var(${s.col})${s.dash ? ";height:2px;border-radius:0" : ""}"></i>${s.name}</span>`).join("");

  // series: balanced panel per group & measure; A reference = that year's A schools
  const mean = vs => vs.length >= 5 ? { m: vs.reduce((a, b) => a + b, 0) / vs.length, n: vs.length } : null;
  const panels = MEAS.map((lbl, j) => {
    const bal = rs => rs.filter(r => [0, 1, 2].every(y => mv(r, j, y) != null));
    const b1 = bal(t1), b2 = bal(nt);
    const ser = {
      t1: [0, 1, 2].map(y => mean(b1.map(r => mv(r, j, y)))),
      nt: [0, 1, 2].map(y => mean(b2.map(r => mv(r, j, y)))),
      a: [0, 1, 2].map(y => mean(ref.filter(r => gy(r, y) === "A" && mv(r, j, y) != null).map(r => mv(r, j, y)))),
    };
    return { lbl, j, ser };
  });
  const all = panels.flatMap(p => Object.values(p.ser).flat()).filter(Boolean).map(p => p.m);
  const lo = all.length ? Math.max(0, Math.floor((Math.min(...all) - 4) / 10) * 10) : 0;
  const hi = all.length ? Math.min(100, Math.ceil((Math.max(...all) + 4) / 10) * 10) : 100;

  const host = $("mPanels");
  host.innerHTML = panels.map((p, i) => `<div class="mp" id="mp${i}"></div>`).join("");
  panels.forEach((p, i) => {
    const el = $("mp" + i), s1 = p.ser.t1, sa = p.ser.a;
    const has = Object.values(p.ser).some(v => v.some(Boolean));
    let note;
    if (!has) note = "Not scored for this school type, or too few schools in the filter.";
    else {
      const parts = [];
      if (s1[0] && s1[2]) parts.push(`Title I ${sgn(s1[2].m - s1[0].m)} pts since 2024`);
      if (s1[2] && sa[2]) { const g = sa[2].m - s1[2].m; parts.push(`${Math.abs(g).toFixed(1)} pts ${g >= 0 ? "behind" : "ahead of"} A schools in 2026`); }
      note = parts.join(" · ") || "Too few Title I schools in the filter.";
    }
    const cc = p.j === 2 ? `<p class="count" style="margin:4px 0 0">Reported a year late: graduates of ${D.cohorts.join(", ")}.</p>` : "";
    el.innerHTML = `<h3>${p.lbl}</h3><p class="ms">${note}</p>`;
    if (!has) return;
    const W = el.clientWidth || 320, H = 210, m = { l: 34, r: 40, t: 10, b: 24 };
    const iw = W - m.l - m.r, ih = H - m.t - m.b;
    const X = y => m.l + iw * (0.08 + 0.84 * y / 2), Y = v => m.t + ih - (v - lo) / (hi - lo) * ih;
    let s = `<svg width="${W}" height="${H}" role="img" aria-label="${esc(p.lbl)} by year">`;
    for (let v = lo; v <= hi; v += 10) {
      s += `<line class="${v === lo ? "ax" : "gl"}" x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}"/><text x="${m.l - 6}" y="${Y(v) + 4}" text-anchor="end">${v}</text>`;
    }
    YRS.forEach((yl, y) => { s += `<text x="${X(y)}" y="${H - 6}" text-anchor="middle">${yl}</text>`; });
    const ends = [];
    MSER.forEach(se => {
      const pts = p.ser[se.k].map((v, y) => v && { x: X(y), y: Y(v.m), v, yr: y }).filter(Boolean);
      if (!pts.length) return;
      s += `<polyline class="mline" points="${pts.map(q => q.x.toFixed(1) + "," + q.y.toFixed(1)).join(" ")}" style="stroke:var(${se.col})"${se.dash ? ' stroke-dasharray="5 4"' : ""}/>`;
      pts.forEach(q => {
        s += `<circle class="mdot" cx="${q.x}" cy="${q.y}" r="4" style="fill:var(${se.col})"/>`;
        const t = `<b>${se.name}</b> · ${YRS[q.yr]}<br>${esc(p.lbl)}: <b>${q.v.m.toFixed(1)}%</b><br><span class="k">${q.v.n} schools${se.k === "a" ? " graded A that year" : ", same schools all 3 years"}</span>`;
        s += `<circle class="mhit" cx="${q.x}" cy="${q.y}" r="10" data-tip="${esc(t)}"/>`;
      });
      const last = pts[pts.length - 1];
      if (last.yr === 2) ends.push({ y: last.y, txt: Math.round(last.v.m) });
    });
    ends.sort((a, b) => a.y - b.y);
    for (let k = 1; k < ends.length; k++) if (ends[k].y - ends[k - 1].y < 12) ends[k].y = ends[k - 1].y + 12;
    ends.forEach(e => { s += `<text class="lbl-ink" x="${X(2) + 9}" y="${e.y + 4}">${e.txt}</text>`; });
    el.innerHTML += s + "</svg>" + cc;
    el.querySelectorAll(".mhit").forEach(c => {
      c.addEventListener("mousemove", e => showTip(c.dataset.tip, e.clientX, e.clientY));
      c.addEventListener("mouseleave", hideTip);
    });
  });
  drawMomTable(t1);
}

const MCOLS = [
  ["s", "School", 0], ["dn", "District", 0], ["m0", "2024", 0], ["m1", "2025", 0], ["m2", "2026", 0],
  ["pp", "Points % 2026", 1], ["need", "Points needed", 1], ["imp", "Where to improve (gap to Title I A avg.)", 0],
  ["d0", "Δ Math L25", 1], ["d1", "Δ ELA L25", 1], ["d2", "Δ C&C accel.", 1],
];
const mval = (r, k) => k === "need" ? need(r) : k === "imp" ? (improve(r)[0] || { g: null }).g : k[0] === "m" && k.length === 2 ? gy(r, +k[1]) : k[0] === "d" && k.length === 2 ? dlt(r, +k[1]) : r[k];
let momT1 = [];
function drawMomTable(t1) {
  if (t1) momT1 = t1;
  let rs = S.mmode === "near" ? momT1.filter(r => GRADES.includes(gy(r, 2)) && gy(r, 2) !== "A")
    : S.mmode === "a" ? momT1.filter(r => gy(r, 2) === "A") : momT1.slice();
  rs.sort((a, b) => {
    const x = mval(a, S.msort), y = mval(b, S.msort);
    if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
    return (x > y ? 1 : x < y ? -1 : 0) * S.mdir;
  });
  $("mHead").innerHTML = "<tr>" + MCOLS.map(([k, l, n]) => `<th class="${n ? "num" : k === "imp" ? "imp" : ""}" data-k="${k}">${l} <span class="ar">${S.msort === k ? (S.mdir < 0 ? "▼" : "▲") : ""}</span></th>`).join("") + "</tr>";
  $("mHead").querySelectorAll("th").forEach(th => th.addEventListener("click", () => {
    const k = th.dataset.k; if (S.msort === k) S.mdir *= -1; else { S.msort = k; S.mdir = ["s", "dn", "need"].includes(k) ? 1 : -1; } drawMomTable();
  }));
  $("mCount").textContent = `${rs.length.toLocaleString()} schools`;
  const dc = v => v == null ? "–" : sgn(v);
  $("mBody").innerHTML = rs.slice(0, 1500).map(r => `<tr>
    <td class="school">${esc(r.s)}</td><td>${esc(r.dn)}</td><td>${gy(r, 0) || "–"}</td><td>${gy(r, 1) || "–"}</td><td>${gy(r, 2) || "–"}</td>
    <td class="num">${r.pp}</td><td class="num">${need(r) ? need(r) : "–"}</td>
    <td class="imp">${improve(r).map(x => `${x.lbl} <b>${Math.round(x.v)}</b> (−${Math.round(x.g)})`).join("<br>") || "–"}</td>
    <td class="num">${dc(dlt(r, 0))}</td><td class="num">${dc(dlt(r, 1))}</td><td class="num">${dc(dlt(r, 2))}</td></tr>`).join("") +
    (rs.length ? "" : `<tr><td colspan="11" class="count">No Title I schools match.</td></tr>`) +
    (rs.length > 1500 ? `<tr><td colspan="11" class="count">Showing first 1,500.</td></tr>` : "");
}

// ---------- table ----------
const COLS = [
  ["s", "School", 0], ["dn", "District", 0], ["t", "Type", 0], ["ed", "Econ. disadv. %", 1], ["pp", "Points %", 1],
  ["ex", "Expected %", 1], ["res", "vs. expected", 1], ["g0", "2026", 0], ["g1", "2025", 0], ["g2", "2024", 0], ["ch", "Charter", 0], ["t1", "Title I", 0],
];
const val = (r, k) => k[0] === "g" && k.length === 2 ? r.g[+k[1]] : k === "t" ? TYPES[r.t] : r[k];
function tableRows() {
  let rs = S.mode === "beat" ? cur.filter(isBeat) : S.mode === "hp" ? cur.filter(isHP) : cur.slice();
  if (S.q) rs = rs.filter(r => r.s.toLowerCase().includes(S.q) || r.dn.toLowerCase().includes(S.q));
  return rs.sort((a, b) => { const x = val(a, S.sort), y = val(b, S.sort); return (x > y ? 1 : x < y ? -1 : 0) * S.dir; });
}
function drawTable() {
  $("thead").innerHTML = "<tr>" + COLS.map(([k, l, n]) => `<th class="${n ? "num" : ""}" data-k="${k}">${l} <span class="ar">${S.sort === k ? (S.dir < 0 ? "▼" : "▲") : ""}</span></th>`).join("") + "</tr>";
  $("thead").querySelectorAll("th").forEach(th => th.addEventListener("click", () => {
    const k = th.dataset.k; if (S.sort === k) S.dir *= -1; else { S.sort = k; S.dir = ["s", "dn", "t"].includes(k) ? 1 : -1; } drawTable();
  }));
  const rs = tableRows();
  $("tCount").textContent = `${rs.length.toLocaleString()} schools`;
  $("tbody").innerHTML = rs.slice(0, 1500).map(r => `<tr>
    <td class="school">${esc(r.s)}</td><td>${esc(r.dn)}</td><td>${TYPES[r.t]}</td>
    <td class="num">${fmt1(r.ed)}</td><td class="num">${r.pp}</td><td class="num">${fmt1(r.ex)}</td>
    <td class="num ${r.res >= S.res ? "pos" : ""}">${sgn(r.res)}</td>
    <td>${r.g[0] || "–"}</td><td>${r.g[1] || "–"}</td><td>${r.g[2] || "–"}</td>
    <td>${r.ch ? "Yes" : "No"}</td><td>${r.t1 ? "Yes" : "No"}</td></tr>`).join("") +
    (rs.length > 1500 ? `<tr><td colspan="12" class="count">Showing first 1,500. Copy the CSV for all rows.</td></tr>` : "");
}
function downloadCsv() {
  const rs = tableRows();
  const head = ["District", "School Number", "School", "Type", "Econ Disadv %", "Points %", "Expected %", "vs Expected", "Grade 2026", "Grade 2025", "Grade 2024", "Charter", "Title I", ...D.components, "Total points 2026", "Components 2026", "Points needed for an A",
    ...[0, 1].flatMap(y => [`Points % ${YRS[y]}`, `Math gains lowest 25% ${YRS[y]}`, `ELA gains lowest 25% ${YRS[y]}`, `College & career acceleration ${YRS[y]} (${D.cohorts[y]})`])];
  const q = v => v == null ? "" : /[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : v;
  const lines = [head.join(",")].concat(rs.map(r => [r.dn, r.sn, r.s, TYPES[r.t], r.ed, r.pp, r.ex, r.res, ...r.g, r.ch ? "Yes" : "No", r.t1 ? "Yes" : "No", ...r.c, r.tp, r.tc, need(r),
    ...[0, 1].flatMap(y => [mv(r, 3, y), mv(r, 0, y), mv(r, 1, y), mv(r, 2, y)])].map(q).join(",")));
  const text = lines.join("\r\n"), btn = $("csv"), box = $("csvBox");
  const done = msg => { btn.textContent = msg; setTimeout(() => { btn.textContent = "Copy table as CSV"; }, 2500); };
  const fallback = () => { box.value = text; box.hidden = false; box.focus(); box.select(); done("Select all and copy below"); };
  box.hidden = true;
  try {
    navigator.clipboard.writeText(text).then(() => done(`Copied ${rs.length} rows`), fallback);
  } catch (e) { fallback(); }
}

// ---------- notes ----------
function drawNotes() {
  const f = D.fits, n100 = D.rows.filter(r => r.ed >= 99.95).length;
  $("notes").innerHTML = [
    `<b>Source:</b> FLDOE “${esc(D.title)}” (${esc(D.source)}). ${D.dropped} schools left out: alternative/ESE centers and schools with no points or no poverty figure (mostly grade “I”, incomplete).`,
    `<b>Expected performance</b> is a straight-line fit of “% of total possible points” on “% economically disadvantaged”, fit separately for each school type: ` +
      Object.entries(f).map(([k, v]) => `${TYPES[k]} ${v.a.toFixed(0)} ${v.b < 0 ? "−" : "+"} ${Math.abs(v.b).toFixed(2)}×poverty (r ${v.r}, n ${v.n})`).join("; ") + ".",
    `<b>Beating the odds</b> = at or above the high-poverty threshold and at least the chosen number of points above expected. Both thresholds are adjustable; neither is an FLDOE definition.`,
    `<b>The 100% stripe:</b> ${n100.toLocaleString()} schools report exactly 100% economically disadvantaged. Most are likely on the federal Community Eligibility Provision, where a whole school is counted as disadvantaged. Treat 100% as “high poverty, exact level unknown”, not as literally every student.`,
    `<b>Grade scales differ by type</b> (elementary A ≥ 62%, middle A ≥ 64%, high/combination A ≥ 65%), which is why the expected line is fit per type. “All types” mixes them.`,
    `<b>Beating the odds is one year.</b> A school can beat the odds once by chance. The 2025 and 2024 grades in the table show whether it is a pattern.`,
    `<b>Momentum</b> uses the 2024 and 2025 School Grades files, matched on district and school number. The trend starts in 2024: 2023 was the baseline year of the new state test, with no learning gains. Title I lines average the same schools in all three years, grouped by 2026 Title I status. The A line is every school of that type graded A that year, so its schools change year to year. <b>Points needed</b> is how many component points a school is short of an A. FLDOE rounds the % of points to a whole number, so a high school needs 645 of 1,000 points (64.5%) for the 65% cutoff (elementary 62, middle 64, high and combination 65). <b>Where to improve</b> compares each component with the average Title I school of the same type graded A in 2026 and lists the three largest shortfalls; it shows where a school trails its successful peers, not what will work.`,
    `<b>College & career acceleration runs a year behind</b>: the 2024, 2025 and 2026 grades use graduates of ${D.cohorts.join(", ")}. It is scored only for high and combination schools.`,
    `<b>Correlation, not cause.</b> This file has charter and Title I status only. Programs (magnet, IB/AP, tutoring, staffing) and per-pupil spending need other sources: the FLDOE School Public Accountability Reports and the federal Civil Rights Data Collection.`,
  ].map(t => `<li>${t}</li>`).join("");
}

function render() {
  $("povV").textContent = S.pov; $("resV").textContent = S.res;
  cur = filtered();
  drawTiles(); drawScatter(); drawGrades(); drawComponents(); drawMomentum(); drawTable();
}
initFilters(); drawNotes(); render();
</script>
"""

if __name__ == "__main__":
    main()
