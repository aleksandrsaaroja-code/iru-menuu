# Koguste arvutamise reeglid (Aleksandr)

Allikas: toidulehed (X:\Statistika\Toidulehed_2026 UUS.xlsx) → osakonnad.json. Ainult arvud, nimesid ei hoita.

## Kliendid
- Üldtoit = kõik kliendid (osakonnad + söökla), peenestatud ja suhkruvaba kaasa arvatud.
- Puhvet lisandub ainult lõunal: E–R 70, L–P 46 (kliendid 304 + puhvet = 374 / 350).
- Päevahoid (4A/2, 4) on E–R klientide hulgas; L ja P lahutatakse. Seega E–R kliente = osakonnad + päevahoid + söökla (toidulehe statistikaga 309), L–P 304.
- Hommik, õhtu, ööoode: sama klientide arv kui lõunal ilma puhvetita.
- Salat: kõik kliendid + puhvet; kliendile 50 g, puhvetile 100 g.

## Peenestatud eraldi rida EI OLE (läheb kokku)
- Pudrud, helbed, moos, keedis, merevaik, kohv, tee, piim, keefir, mahl, morss, kakao
- Kisellid, pasteedid, või
- Kastmed, hapukoor, ketšup – ainult kui portsjon on 20, 30 või 50 g
- Kakao-kohupiim, banaani-kohupiim
- Salat
- Saiake: peenestatud ei saa (tükikaup)

## Peenestatud eraldi rida ON
- Supp (üldkogusest lahutada peenestatud kogus ja püreerida)
- Kartul, liha, makaronid, vahud, kohupiimavorm/-kook, kastmed ≥ 100 g jne

## Tükikaup, peenestatud juures kaal g
- Tükikaubal on PORTSJON kaalus (nt kalapala 4 tk = 100 g). Põhikogus näidatakse kg-des (nt 30,8 kg), tükkide arv (1232 tk) on ainult INFOKS.
- Lõhevorm 150 g; kalapala, sardell, pihv, kanapihv, toorvorst, viiner 100 g (kaste mitte)

## Piimavaba rida (osakonniti) ON
- Pudrud, piim, keefir, jogurt, kohupiim, kakao, kohuke, hapukoor, frikadellikaste, karamellkissell
## Piimavaba rida EI OLE
- Juust, juustupasteet (laktoosivaba), vahukoor, või
- Hirsipuder (esmaspäeval) on tehtud ilma piimata, seega piimavaba rida pole vaja

## Vorm
- Üksuse rida: sulgudes tavatoidu arv (kliendid MIINUS peenestatud) ja selle kg, kõrval "+ peenestatud (arv) kg". Nt 1A (24) 3,6 kg + peenestatud (4) 0,6 kg.
- Piimavaba sama põhimõte: sulgudes tavatoit = kliendid − peenestatud − piimavaba; kõrval "+ piimavaba (arv) kg". Piimavaba = ainult mitte-peenestatud (peenestatud piimavabu eraldi rida pole, nad on peenestatud hulgas). Arv tuleb väljalt `mv` (osakonnad.json).
- Portsjoni kogust (nt kalapala 4 tk → 1 tk) saab äpis päeva kaupa muuta (avatud rea lahter „Portsjon“); kaal EI muutu (portsjon jääb nt 100 g, tk on ainult info). Muudatus kehtib ainult sellel päeval ja selles telefonis, „Lähtesta“ toob menüü väärtuse tagasi.
- KÕIGIL tükkides (tk) toitudel (banaan, kotlet, pannkook jne) on lahter „Portsjoni kaal (g)“: selle muutmine arvutab kaalu (kg) vastavalt tootele (nt 80 g → 24,64 kg); tk lahter on ainult info. Päeva kaupa, selles telefonis.
- Ühine salvestus (Google Apps Script, tools/google-skript.gs): PIN-iga (Seaded avatud) tehtud muudatused (osakondade arvud ja portsjonid/kaalud) näevad kõik kokad; kokad ei saa kirjutada. Käsitsi arvud kehtivad, kuni toidulehe andmed päriselt muutuvad. Sünkroon SYNC_URL on kokad/index.html alguses.
