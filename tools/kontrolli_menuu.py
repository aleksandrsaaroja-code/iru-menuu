#!/usr/bin/env python3
"""Kirjavigade kontroll kogu menüüle (5 nädalat x 7 päeva). Ei peata avaldamist.
Kirjutab kirjavead.txt repo juurde. Esimene rida: "Kirjavigu: N"."""
import json, re, sys, pathlib, datetime, collections

root = pathlib.Path(__file__).resolve().parent.parent
FIX = "--paranda" in sys.argv
sys.argv = [a for a in sys.argv if a != "--paranda"]
menu = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else root / "menu.json"
known_file = pathlib.Path(__file__).resolve().parent / "tuntud-sonad.txt"
out = root / "kirjavead.txt"

# Sama toit kirjutatud erinevalt: kui mõlemad vormid esinevad, hoiatame.
PAIRS = [("VANILJEKASTE", "VANILLIKASTE"), ("KAKAOKISSELL", "KAKAO-KISSELL"),
         ("MAISIMANNA", "MAISIMANNAPUDER")]
LETTERS = r"[A-ZÄÖÜÕŠŽ]+"

known = set()
if known_file.exists():
    known = {l.strip().upper() for l in known_file.read_text(encoding="utf-8").splitlines() if l.strip()}

data = json.loads(menu.read_text(encoding="utf-8"))
issues, names = [], collections.Counter()
for w, days in data["weeks"].items():
    for d in days:
        for s in d["s"]:
            meal = s["t"].split()[0].title()
            for n, q in s["i"]:
                if re.match(r"^LEIB, SAI", n):
                    continue
                where = f"Nädal {w}, {d['d']}, {meal}"
                text = n.split(": ", 1)[1] if re.match(r"^(DIABEET|PEENESTATUD): ", n) else n
                if n != n.strip(): issues.append((where, n, "tühik alguses või lõpus"))
                if "  " in n: issues.append((where, n, "topelttühik"))
                if re.search(r"\s,", n): issues.append((where, n, "tühik enne koma"))
                if re.search(r",\S", n): issues.append((where, n, "koma järel puudub tühik"))
                if n != n.upper(): issues.append((where, n, "väiketähed (menüü on suurtähtedega)"))
                if re.search(r"\bJA\b", n): issues.append((where, n, "„ja“ asemel sidekriips (nt MUNA-SINGIKASTE)"))
                for x in re.findall(LETTERS, text.upper()):
                    names[x] += 1
                    if known and x not in known:
                        issues.append((where, n, f"tundmatu sõna „{x}“ (kirjaviga? kui on õige, lisa faili tools/tuntud-sonad.txt)"))
                names[text.strip().upper()] += 0
all_text = " ".join(n for days in data["weeks"].values() for d in days for s in d["s"] for n, q in s["i"]).upper()
for a, b in PAIRS:
    if re.search(rf"(?<![A-ZÄÖÜÕŠŽ-]){re.escape(a)}(?![A-ZÄÖÜÕŠŽ-])", all_text) and re.search(rf"(?<![A-ZÄÖÜÕŠŽ-]){re.escape(b)}(?![A-ZÄÖÜÕŠŽ-])", all_text):
        issues.append(("kogu menüü", f"{a} / {b}", "sama toit kirjutatud kahel viisil, vali üks"))


# --- Automaatne parandus (ainult --paranda): tühikud, koma ja tuntud kirjavead failist tools/parandused.txt ---
if FIX:
    pf = pathlib.Path(__file__).resolve().parent / "parandused.txt"
    amap = {}
    if pf.exists():
        for l in pf.read_text(encoding="utf-8").splitlines():
            if l.strip() and not l.startswith("#") and "=" in l:
                a, b = l.split("=", 1); amap[a.strip().upper()] = b.strip().upper()
    def clean(t):
        t = re.sub(r"\s+", " ", t).strip()
        t = re.sub(r"\s+,", ",", t)
        t = re.sub(r",(?=\S)", ", ", t)
        for a, b in amap.items():
            t = re.sub(rf"(?<![A-ZÄÖÜÕŠŽ-]){re.escape(a)}(?![A-ZÄÖÜÕŠŽ-])", b, t)
        return t
    changed = 0
    for days in data["weeks"].values():
        for d in days:
            for s in d["s"]:
                for it in s["i"]:
                    new = clean(it[0])
                    if new != it[0]:
                        it[0] = new; changed += 1
    if changed:
        menu.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"Parandatud ridu: {changed}")
    # kontrolli uuesti parandatud andmetega
    issues, names = [], collections.Counter()
    for w, days in data["weeks"].items():
        for d in days:
            for s in d["s"]:
                meal = s["t"].split()[0].title()
                for n, q in s["i"]:
                    if re.match(r"^LEIB, SAI", n): continue
                    text = n.split(": ", 1)[1] if re.match(r"^(DIABEET|PEENESTATUD): ", n) else n
                    for x in re.findall(LETTERS, text.upper()):
                        if known and x not in known:
                            issues.append((f"Nädal {w}, {d['d']}, {meal}", n, f"tundmatu sõna „{x}“ (kirjaviga? kui on õige, lisa faili tools/tuntud-sonad.txt)"))
    all_text = " ".join(n for days in data["weeks"].values() for d in days for s in d["s"] for n, q in s["i"]).upper()
    for a, b in PAIRS:
        if re.search(rf"(?<![A-ZÄÖÜÕŠŽ-]){re.escape(a)}(?![A-ZÄÖÜÕŠŽ-])", all_text) and re.search(rf"(?<![A-ZÄÖÜÕŠŽ-]){re.escape(b)}(?![A-ZÄÖÜÕŠŽ-])", all_text):
            issues.append(("kogu menüü", f"{a} / {b}", "sama toit kirjutatud kahel viisil, vali üks"))
    seen, uniq = set(), []
    for it in issues:
        if it not in seen: seen.add(it); uniq.append(it)
    lines = [f"Kirjavigu: {len(uniq)}", f"Kontrollitud: {datetime.datetime.now():%d.%m.%Y %H:%M} (kõik 5 nädalat, kõik päevad); automaatselt parandatud ridu: {changed}", ""]
    lines += [f"{w} | {n} | {why}" for w, n, why in uniq] or ["Kõik korras."]
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(lines[0]); sys.exit(0)
seen, uniq = set(), []
for it in issues:
    if it not in seen:
        seen.add(it); uniq.append(it)
lines = [f"Kirjavigu: {len(uniq)}", f"Kontrollitud: {datetime.datetime.now():%d.%m.%Y %H:%M} (kõik 5 nädalat, kõik päevad)", ""]
lines += [f"{w} | {n} | {why}" for w, n, why in uniq] or ["Kõik korras."]
out.write_text("\n".join(lines) + "\n", encoding="utf-8")
print(lines[0])
