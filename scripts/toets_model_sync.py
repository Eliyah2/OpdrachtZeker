"""Toets dat de Python-mirror (xlsx-model) en het dashboard (JS) niet uit elkaar lopen.

Strategie: extraheer de rekenlogica uit BEIDE bronnen en vergelijk de uitkomsten
op een breed scenario-palet (base, presets + 40 deterministische steekproeven).
Geen derde implementatie: de twee bestaande implementaties worden letterlijk
tegen elkaar gedraaid. Draait in <2 sec; bedoeld als pre-commit-hook.
"""
import ast
import copy
import json
import random
import subprocess
import sys
import tempfile
from pathlib import Path

GEN = Path("scripts/financieel_model_bouwen.py")
DASH = Path("dashboard/index.html")
TOL = 1e-6

# ---------- 1. Python-mirror extraheren uit het generatiescript ----------
src = GEN.read_text(encoding="utf-8")
tree = ast.parse(src)
wanted = []
for node in tree.body:
    if isinstance(node, ast.FunctionDef) and node.name in ("rnd", "run"):
        wanted.append(node)
    elif isinstance(node, ast.Assign):
        names = [t.id for t in node.targets if isinstance(t, ast.Name)]
        if any(n in ("A", "NIEUWE_FREE") for n in names):
            wanted.append(node)
ns = {"math": __import__("math")}
exec(compile(ast.Module(body=wanted, type_ignores=[]), str(GEN), "exec"), ns)
run_py, A0, NEWFREE = ns["run"], ns["A"], ns["NIEUWE_FREE"]
assert NEWFREE == [50, 70, 90, 110, 140, 170, 200, 240, 280, 320, 370, 420]

# ---------- 2. JS-model extraheren uit het dashboard ----------
html = DASH.read_text(encoding="utf-8")
js = html[html.index("const NEWFREE"): html.index("const nf = new Intl")]

MAP = [  # (JS-sleutel, Aannames-sleutel)
    ("arpu", "arpu"), ("proPrijs", "pro_prijs"), ("conversie", "conversie_pro"),
    ("churn", "churn"), ("cac", "cac"), ("salesNew", "sales_new"),
    ("salaris", "salaris"), ("salesHire", "sales_hire"),
    ("mkt1", "mkt1"), ("mkt2", "mkt2"), ("mkt3", "mkt3"),
    ("betaEinde", "beta_einde"), ("betaBoost", "beta_boost"),
]

# ---------- 3. Scenario-palet ----------
base = {"startkapitaal": 75000, "arpu": 230, "proPrijs": 12, "conversie": 0.04, "churn": 0.03,
        "cac": 600, "salesNew": 2, "salaris": 5000, "salesHire": 4500,
        "mkt1": 1000, "mkt2": 2000, "mkt3": 3000, "betaEinde": 3, "betaBoost": 1.5}
presets = [
    {**base, "churn": 0.06, "cac": 1000, "arpu": 190},
    {**base, "churn": 0.02, "cac": 450, "arpu": 260},
]
rng = random.Random(42)
draws = []
for _ in range(40):
    draws.append({
        "startkapitaal": 75000,
        "arpu": rng.randint(100, 500), "proPrijs": rng.randint(0, 30),
        "conversie": round(rng.uniform(0, 0.15), 3), "churn": round(rng.uniform(0, 0.10), 3),
        "cac": rng.randint(200, 2000), "salesNew": rng.randint(0, 10),
        "salaris": rng.randint(2000, 15000), "salesHire": rng.randint(0, 10000),
        "mkt1": rng.randint(0, 5000), "mkt2": rng.randint(0, 8000), "mkt3": rng.randint(0, 15000),
        "betaEinde": rng.randint(0, 12), "betaBoost": round(rng.uniform(1, 3), 2),
    })
scenarios = [base] + presets + draws

# ---------- 4. Python-uitkomsten ----------
def py_outcome(p):
    A = copy.deepcopy(A0)
    for js_key, a_key in MAP:
        A[a_key] = (p[js_key], "", "")
    ns["A"] = A
    rows, be, minKap = run_py(p["churn"], p["cac"], p["arpu"], p["salesNew"])
    return {
        "be": be, "minKap": float(minKap), "kasZ12": float(rows[-1][7]),
        "omzet12": float(rows[-1][3]),
        "rows": [[o[1], o[2], o[3], o[4], o[5], o[7]] for o in rows],
    }

# ---------- 5. JS-uitkomsten via Node ----------
harness = js + """
const scenarios = __SCENARIOS__;
const out = scenarios.map(p => {
  const r = runModel(p);
  return { be: r.be, minKap: r.minKap, kasZ12: r.kasZ12, omzet12: r.omzet12,
           rows: r.rows.map(x => [x.act, x.pro, Math.round(x.omzet), Math.round(x.ktot),
                                  Math.round(x.res), Math.round(x.kasZ)]) };
});
console.log(JSON.stringify(out));
"""
harness = harness.replace("__SCENARIOS__", json.dumps(scenarios))
with tempfile.TemporaryDirectory() as td:
    mjs = Path(td) / "model.mjs"
    mjs.write_text(harness, encoding="utf-8")
    res = subprocess.run(["node", str(mjs)], capture_output=True, text=True, timeout=60)
    if res.returncode != 0:
        print("FAIL - Node kon het dashboard-model niet draaien:\n" + res.stderr)
        sys.exit(1)
    js_out = json.loads(res.stdout)

# ---------- 6. Vergelijken ----------
fails = 0
for i, (p, pj, pjs) in enumerate(zip(scenarios, [py_outcome(s) for s in scenarios], js_out)):
    tag = f"scenario {i} ({'base' if i == 0 else 'preset' if i < 3 else 'steekproef'})"
    for key in ("be", "minKap", "kasZ12", "omzet12"):
        a, b = pj[key], pjs[key]
        if (a is None) != (b is None) or (a is not None and abs(float(a) - float(b)) > TOL):
            print(f"FAIL - {tag}: {key} python={a} js={b}")
            fails += 1
    if pj["rows"] != pjs["rows"]:
        for m, (ra, rb) in enumerate(zip(pj["rows"], pjs["rows"]), start=1):
            if ra != rb:
                print(f"FAIL - {tag}: maand {m} python={ra} js={rb}")
        fails += 1

if fails:
    print(f"\nModel-sync-toets MISLUKT ({fails} afwijkingen) - xlsx en dashboard lopen uit elkaar!")
    sys.exit(1)
print(f"PASS - Model-sync: {len(scenarios)} scenario's identiek in xlsx-mirror en dashboard (be, minKap, kas, omzet, alle 12 maanden)")
