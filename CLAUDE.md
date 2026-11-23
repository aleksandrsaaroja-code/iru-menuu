# Iru menüü (äpp) – töökorraldus

Kasutaja: Aleksandr Saaroja, Iru Hooldekodu toitlustusjuht. Räägi eesti keeles, lühidalt ja asjalikult, alusta tulemusega.

## Mälu (loe enne töö alustamist)
- Koguste arvutamise reeglid: `tools/REEGLID.md`. Loe need enne, kui muudad `index.html` arvutust. Kasutaja on neid mitu korda öelnud, ära küsi uuesti.
- Kaks äppi: juur (`index.html`) = vana menüüäpp, ÄRA MUUDA; `kokad/` = uus äpp kokkadele (kogused).
- Arvud tulevad toidulehelt: `tools/uuenda_arvud.py` → `kokad/osakonnad.json`. Menüü: `menu.json`.

## Töö reeglid
- **Kui arvad, et kasutaja eksib või ütlus on kahemõtteline, küsi üle** enne tegemist. Ära arva.
- Kui kasutaja ütleb reegli, lisa see kohe faili `tools/REEGLID.md`.
- Fraas „peenestatud rida ei ole vaja“ = peenestatud ei näidata eraldi, läheb üldkogusega kokku. See EI tähenda, et peenestatud ei söö.
- Muutuse järel kontrolli kõiki 5 nädalat × 7 päeva (ei tohi olla NaN ega kogust, mille üksuste summa ei klapi).
- Isikuandmeid (elanike nimesid) äppi ega repo ei pane, ainult arvud.
- Kirju ei saada ilma loata, tee mustand. Pull requesti ega `main`-i muudatust ei tee ilma selge loata.
