#!/usr/bin/env python3
"""Loeb toidulehe (X:\\Statistika\\Toidulehed_2026 UUS.xlsx) ja kirjutab osakonnad.json.
Ainult arvud, nimesid ei võeta. Kasutus:
  python uuenda_arvud.py "X:\\Statistika\\Toidulehed_2026 UUS.xlsx" [--push]
Vajab: pip install openpyxl
"""
import sys, json, subprocess, datetime, pathlib, re
import openpyxl

src = sys.argv[1] if len(sys.argv) > 1 else r"X:\Statistika\Toidulehed_2026 UUS.xlsx"
out = pathlib.Path(__file__).resolve().parent.parent / "kokad" / "osakonnad.json"
OLD = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
UNITS = ["1A","2A-1","2A-2","2A-3","2B","2D","3A","3B","3C","3D","4A-1","4A-2","4A-3"]

# Nimede maskeerimine: suure algustähega sõnarühm (nimi, perekonnanimi, initsiaal) = üks inimene.
# Lubatud suure algustähega sõnad (lause algus, tavasõnad) jäävad; KOGU-SUURTÄHTEDES sõnad jäävad. Privaatsus ees: kahtluse korral maskeeritakse.
ALLOW = {"Nimed", "Nimi", "Üks", "Kaks", "Kolm", "Neli", "Viis", "Sööb", "Lisa", "Palun", "Ei", "Soovib", "Kuni", "Iga", "Jah", "Eri", "Tavaline", "Peenestatud", "Püreeritud", "Valgurikas", "Piimavaba", "Suhkruvaba", "Insuliin", "Diabeedi", "Diabeetik"}
_CAP = r"(?:[A-ZÕÄÖÜŠŽ][a-zõäöüšž]+\.?|[A-ZÕÄÖÜŠŽ]\.)"   # Nimi/Perenimi või initsiaal "M."; KOGU-SUURTÄHTEDES sõnad (nt LISA) ei ole nimi
_NAME_RE = re.compile(r"(?<![A-Za-zÕÄÖÜŠŽõäöüšž])" + _CAP + r"(?:[ ]*" + _CAP + r")*")
def mask_names(text):
    if text is None or not str(text).strip():
        return ""
    def repl(m):
        words = re.findall(_CAP, m.group(0))
        keep = []
        while words and words[0].rstrip(".") in ALLOW:
            keep.append(words.pop(0).rstrip("."))
        out = " ".join(keep)
        if words:
            out = (out + " " if out else "") + "\x00"      # nimi eemaldatakse (jäetakse välja), arv on sildis
        return out
    t = _NAME_RE.sub(repl, str(text))
    t = t.replace("\x00", "")
    t = re.sub(r"\bNimed\b(\s*\([^)]*\))?\s*:?", "", t)           # "Nimed (toidulehelt):" jms silt
    t = re.sub(r"\(\s*[,/;\-\s]*\)", "", t)                         # tühjad sulud
    t = re.sub(r"\(\s*[,/;\-]*\s*", "(", t); t = re.sub(r"\s*[,/;\-]*\s*\)", ")", t)
    t = re.sub(r"([,/;])(?:\s*[,/;])+", r"\1", t)                   # korduvad eraldajad
    t = re.sub(r"\.\s+(?:[,/;\-]\s*)+", ". ", t)                    # ". , -" pärast nimede eemaldamist
    segs = [re.sub(r"^[\s,/;\-]+|[\s,/;\-]+$", "", s) for s in t.split("|")]
    t = " | ".join(s for s in segs if s)
    return re.sub(r"\s+", " ", t).replace(" .", ".").strip()[:400]

# LISAMÄRKED (read 12-17) ja LISAPALAD (28-30): arv + kõrvalolev tekst (veerud C-G), nimed maskeeritud
ROW_LABELS = {12: "Gluteenivaba", 13: "Puriinivaene", 14: "Vedel / sond", 15: "Täielik piimavaba", 16: "Taimetoit", 17: "Muu eridieet",
              30: "Valgurikas lisapala (VR)"}   # võileib (28) ja tops (29) on eraldi sakis Tops/Võileib, märkustesse ei panda
def collect_notes(ws):
    out = []
    for r, lab in ROW_LABELS.items():
        b = ws[f"B{r}"].value
        cnt = int(b) if isinstance(b, (int, float)) else 0
        txt = " | ".join(mask_names(ws[f"{c}{r}"].value) for c in "CDEFG"
                         if isinstance(ws[f"{c}{r}"].value, str) and ws[f"{c}{r}"].value.strip())
        if cnt or txt:
            out.append([f"{lab} ({cnt})" if cnt else lab, txt])
    # tekstuuritabeli "märkus" veerg (G6-G8): AINULT toidu juhised (mida anda), ilma sildita (Aleksandr 26.11.2026).
    # Dieedi kirjeldusi (PV/DB, laktoos, diabeet, ...vaba), söötmist, sondi ja "Sööklas" (C21) ei näidata.
    for cell in ("G6", "G7", "G8"):
        for seg in clean_note(ws[cell].value).split(" | "):
            if seg and _FOOD.search(seg) and not _DIET.search(seg):
                out.append([seg, "", 1])        # 3. väli 1 = uus märkus (äpis ajutiselt punane)
    return out

_FOOD = re.compile(r"lisand|asemel|\banda\b|antakse|keefir|jogu?urt|joogurt|kohuke|\bmuna|mahl|kissell|\bpiim(?!avaba)|puder|pudru|supp|tatar|leib|sai\b|võileib", re.I)
_DIET = re.compile(r"\[(?:DB|PV)|piimavaba|laktoos|diabee|\w+vaba\b|sööda|gastrostoom|sond", re.I)

_TAG = re.compile(r"\[?\(?\b(?:DB\s*\+\s*PV|DB|PV)\b\)?\]?\s*:?", re.I)
def clean_note(v):
    """Märkuse tekst ilma nimede, initsiaalide ja tühjade DB/PV siltideta."""
    if not isinstance(v, str) or not v.strip():
        return ""
    segs = []
    for s in mask_names(v).split("|"):
        s = re.sub(r"(?<![A-Za-zÕÄÖÜŠŽõäöüšž])[A-ZÕÄÖÜŠŽ](?![A-Za-zÕÄÖÜŠŽõäöüšž+])", "", s)   # üksikud initsiaalid
        s = re.sub(r"\bja\b(?=\s*[,|]|\s*$)", "", s)                                            # "ja" ilma nimeta
        s = re.sub(r"\(?\bDIAB\b\)?\s*:?", "", s)                                               # "DIAB" = sama mis [DB]
        s = re.sub(r"\(\s*\)", "", s)
        rest = _TAG.sub("", s)
        rest = re.sub(r"[\s,/;\-:()]+", " ", rest).strip()
        if not rest:
            continue
        s = re.sub(r"\s*,\s*(?:[,\-]\s*)*", ", ", s)
        s = re.sub(r"^\s*,\s*", "", s); s = re.sub(r"\]\s*,\s*", "] ", s)
        s = re.sub(r"^\s*\(([^()]*)\)\s*$", r"\1", s)                                          # terve tekst sulgudes
        s = re.sub(r"^[\s,/;\-]+|[\s,/;\-]+$", "", s)
        s = re.sub(r"\[\s*", "[", s); s = re.sub(r"\s+", " ", s).strip(" ,")
        segs.append(s)
    return " | ".join(segs)

import time
LOGFILE = pathlib.Path(__file__).resolve().parent.parent / "arvud.log"   # *.log on .gitignore'is, jääb ainult sellesse arvutisse
def log(msg):
    try:
        with open(LOGFILE, "a", encoding="utf-8") as f:
            f.write(datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S") + " " + msg + "\n")
    except Exception:
        pass

wb = None
for attempt in range(1, 4):       # toiduleht võib olla salvestamisel pooleli või X-ketas hetkeks maas
    try:
        wb = openpyxl.load_workbook(src, data_only=True, read_only=False)
        break
    except Exception as e:
        log(f"toidulehte ei saanud lugeda (katse {attempt}/3): {type(e).__name__}: {e}")
        if attempt < 3:
            time.sleep(10)
if wb is None:
    print("Toidulehte ei saanud lugeda, jätan vana osakonnad.json")
    sys.exit(1)
n = lambda ws, a: int(ws[a].value or 0) if isinstance(ws[a].value, (int, float)) else 0
deps, sookla, dates, warn = [], 0, {}, []
for u in UNITS:
    ws = wb[u]
    g = lambda r, c: n(ws, f"{c}{r}")          # ridu 6=Üldtoit 7=Pehme 8=Püreeritud; veerud B tavaline, C DB, D PV, E DB+PV
    pur = sum(g(8, c) for c in "BCDE")
    db = sum(g(r, c) for r in (6, 7) for c in "CE")
    pv = sum(g(r, c) for r in (6, 7, 8) for c in "DE") + g(8, "C") * 0
    pv = sum(g(r, c) for r in (6, 7, 8) for c in "DE")
    mv = sum(g(r, c) for r in (6, 7) for c in "DE")   # piimavaba, mitte-peenestatud
    both = sum(g(r, "E") for r in (6, 7))              # suhkruvaba JA piimavaba (veerg E), mitte-peenestatud; on nii db kui mv sees
    osk, sk = g(20, "B"), g(21, "B")
    if sum(g(r, c) for r in (6, 7, 8) for c in "BCDE") != osk:
        warn.append(f"{u}: tabeli summa ei võrdu 'Osakonnas sööb' ({osk})")
    deps.append({"n": u.replace("-", "/"), "t": osk + g(25, "B"), "p": pur, "m": pv, "mv": mv, "d": db, "b": both, "h": g(25, "B"),
                 # toidulehe LISAMÄRKED (ridu 12-17) ja LISAPALAD (ridu 28-30): ainult arvud, nimesid ei võeta
                 "gv": g(12, "B"), "pn": g(13, "B"), "ved": g(14, "B"), "tpv": g(15, "B"), "tt": g(16, "B"), "me": g(17, "B"),
                 "v": g(28, "B"), "jp": g(29, "B"), "vr": g(30, "B"),
                 "mt": mask_names(ws["B31"].value), "ml": collect_notes(ws)})   # rida 31 "Eri / märkused": tekst, kus nimed on asendatud sõnaga "inimene" (üks nimi = üks inimene)
    sookla += sk
    v = ws["E2"].value
    dates[u.replace("-", "/")] = v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else None
deps.append({"n": "Söökla", "t": sookla, "p": 0, "m": 0, "d": 0})
puh = OLD.get("puhvet", 70)
puh_w = OLD.get("puhvet_w", 46)   # puhvet laupäeval ja pühapäeval
deps.append({"n": "Puhvet", "t": puh, "w": puh_w, "p": 0, "m": 0, "d": 0})
res = {"updated": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), "puhvet": puh, "puhvet_w": puh_w, "dates": dates, "deps": deps}
same = False
if out.exists():
    try:
        o = dict(OLD); o.pop("updated", None); n = dict(res); n.pop("updated", None)
        same = (o == n)
    except Exception:
        same = False
if same:
    print("Arvud ei muutunud, faili ei kirjutata (uuendamise aeg jääb endiseks)")
    sys.exit(0)
out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
print("Kirjutatud", out, "· sööjaid kokku", sum(d["t"] for d in deps))
for w in warn: print("HOIATUS:", w)
log("osakonnad.json kirjutatud, sööjaid kokku " + str(sum(d["t"] for d in deps)))
if "--push" in sys.argv:
    d = str(out.parent.parent)
    git = lambda *a: subprocess.run(["git", "-C", d, *a], capture_output=True, text=True)
    git("add", "kokad/osakonnad.json")
    if git("diff", "--cached", "--quiet").returncode:
        c = git("commit", "-m", "Arvud uuendatud " + res["updated"])
        if c.returncode:
            log("commit ebaõnnestus: " + (c.stderr or c.stdout).strip()[:200]); sys.exit(1)
        ok = False
        for attempt in range(1, 4):   # push võib ebaõnnestuda, kui menüü avaldab samal ajal: tõmba ja proovi uuesti
            p = git("push")
            if p.returncode == 0:
                ok = True; log(f"push õnnestus (katse {attempt})"); break
            log(f"push ebaõnnestus (katse {attempt}/3): " + (p.stderr or p.stdout).strip()[:200].replace("\n", " "))
            r = git("pull", "--rebase", "--autostash")
            if r.returncode:
                git("rebase", "--abort")
                log("pull --rebase ebaõnnestus (konflikt?), katkestasin rebase'i: " + (r.stderr or r.stdout).strip()[:200].replace("\n", " "))
                break
            time.sleep(5)
        if not ok:
            print("PUSH EBAÕNNESTUS, vaata arvud.log"); sys.exit(1)
