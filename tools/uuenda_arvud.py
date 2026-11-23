#!/usr/bin/env python3
"""Loeb toidulehe (X:\\Statistika\\Toidulehed_2026 UUS.xlsx) ja kirjutab osakonnad.json.
Ainult arvud, nimesid ei võeta. Kasutus:
  python uuenda_arvud.py "X:\\Statistika\\Toidulehed_2026 UUS.xlsx" [--push]
Vajab: pip install openpyxl
"""
import sys, json, subprocess, datetime, pathlib, re
import openpyxl

src = sys.argv[1] if len(sys.argv) > 1 else r"X:\Statistika\Toidulehed_2026 UUS.xlsx"
out = pathlib.Path(__file__).resolve().parent.parent / "osakonnad.json"
OLD = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
UNITS = ["1A","2A-1","2A-2","2A-3","2B","2D","3A","3B","3C","3D","4A-1","4A-2","4A-3"]

wb = openpyxl.load_workbook(src, data_only=True, read_only=False)
n = lambda ws, a: int(ws[a].value or 0) if isinstance(ws[a].value, (int, float)) else 0
deps, sookla, dates, warn = [], 0, {}, []
for u in UNITS:
    ws = wb[u]
    g = lambda r, c: n(ws, f"{c}{r}")          # ridu 6=Üldtoit 7=Pehme 8=Püreeritud; veerud B tavaline, C DB, D PV, E DB+PV
    pur = sum(g(8, c) for c in "BCDE")
    db = sum(g(r, c) for r in (6, 7) for c in "CE")
    pv = sum(g(r, c) for r in (6, 7, 8) for c in "DE") + g(8, "C") * 0
    pv = sum(g(r, c) for r in (6, 7, 8) for c in "DE")
    osk, sk = g(20, "B"), g(21, "B")
    if sum(g(r, c) for r in (6, 7, 8) for c in "BCDE") != osk:
        warn.append(f"{u}: tabeli summa ei võrdu 'Osakonnas sööb' ({osk})")
    deps.append({"n": u.replace("-", "/"), "t": osk, "p": pur, "m": pv, "d": db, "h": g(25, "B")})
    sookla += sk
    v = ws["E2"].value
    dates[u.replace("-", "/")] = v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else None
deps.append({"n": "Söökla", "t": sookla, "p": 0, "m": 0, "d": 0})
puh = OLD.get("puhvet", 70)
puh_w = OLD.get("puhvet_w", 50)   # puhvet laupäeval ja pühapäeval
deps.append({"n": "Puhvet", "t": puh, "w": puh_w, "p": 0, "m": 0, "d": 0})
res = {"updated": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), "puhvet": puh, "puhvet_w": puh_w, "dates": dates, "deps": deps}
out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
print("Kirjutatud", out, "· sööjaid kokku", sum(d["t"] for d in deps))
for w in warn: print("HOIATUS:", w)
if "--push" in sys.argv:
    d = out.parent
    subprocess.run(["git", "-C", str(d), "add", "osakonnad.json"], check=True)
    if subprocess.run(["git", "-C", str(d), "diff", "--cached", "--quiet"]).returncode:
        subprocess.run(["git", "-C", str(d), "commit", "-m", "Arvud uuendatud " + res["updated"]], check=True)
        subprocess.run(["git", "-C", str(d), "push"], check=True)
