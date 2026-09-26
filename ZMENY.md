# PDA Suite 3J — zmeny

## 3.2.16 — 2026-09-26

**Drobné opravy po druhej kontrole kódu:**
- klik na hlavičku Osobný stav už panel ani vnútorne nezbalí a hlavička nemá kurzor ruky;
- neaktívna dlaždica Components / BOM (alebo Výkres) sa pod myšou nenadvihne.

## 3.2.15 — 2026-09-26

**Opravy po kontrole kódu** (vzhľad sa nemení):
- výška ľavého zoznamu má rezervu 8 px pod sebou (väčšiu než prah zmeny výšky), takže detail nikdy nepretečie ani o pár px a nedá sa posunúť pod hlavičku;
- keď je stránka posunutá, výška zoznamu sa ráta bez posunu (predtým mohla spadnúť na náhradných 460 px);
- s pevne pripnutým ľavým stĺpcom (rozpracovaný modul) sa zoznam pri posúvaní nezmenšuje;
- bočný panel rozpoznáva hlavnú stránku len podľa toho, či je naozaj zobrazená — do appky pri tom nezasahuje (predtým volal funkciu, ktorá pri štarte appky mohla ovplyvniť prvú otvorenú stránku).

## 3.2.14 — 2026-09-26

**Načítavanie** — namiesto čakacích značiek S v každej karte je jedna, 3× väčšia, v strede obrazovky (rovnaký vzhľad ako v appke). Karty sú počas načítania stále jemne zastreté. Okná (napr. BOM) a rozbaľovacie zoznamy si svoj indikátor nechávajú.

**Stavové tlačidlá** — obrázky v pozadí ešte výraznejšie: sila 1,0 (predtým 0,65, pôvodne 0,5), zmiešanie „soft-light". Ďalší stupeň by bolo výraznejšie zmiešanie „overlay".

## 3.2.13 — 2026-09-26

**Bočný panel PDA aj na hlavnej stránke** — panel (Výkresy, CHIPS, Majster, Materiál, TOOLSHOP, Ela, Flexus) je teraz aj na úvodnej obrazovke s pracoviskami, nielen v detaile pracoviska. Leží nad obsahom (nezaberá miesto) a dá sa schovať pásikom HF SLOVAKIA. Začína v jednej výške s kartou Vyhľadať zákazku, pri posúvaní stránky stojí a končí nad pätičkou (text v pätičke je vidieť). Na prihlásení, Reportoch, Správach, Rozvrhu a v Admine panel nie je.

**Horné menu už nezakrýva tlačidlá Stretnutia / Prestávka / Čakanie** — obsah detailu bol o ~20 px vyšší ako obrazovka (výška ľavého zoznamu nepočítala s okrajmi panela pod ním), stránka sa preto dala posunúť a horná časť zaliezla pod hlavičku. Zoznam teraz končí presne nad pätičkou a posúvanie kolieskom v zozname už nepohne celou stránkou.

**SAP ČASY** — hotová časť koláča je zelená (aj bodka „Hotovo" v legende).

**Vyhľadať zákazku** — pole je 450 px široká pilulka so zaoblenými rohmi a tenkým tmavomodrým rámom dookola; ikonky zrušiť / hľadať sú na jej konci.

**Hlavná stránka bez tlačidla *Hľadať výrobný príkaz*** — tlačidlo v hlavičke karty Pracoviská sa nezobrazuje (zákazky sa hľadajú v karte Vyhľadať zákazku); v appke ostáva, je len skryté.

**Panely sa nedajú zbaliť** — Osobný stav (detail aj hlavná stránka) a panel pracoviska v detaile už nemajú šípku na zbalenie; obsah je vždy zobrazený.

**Nadpis pracoviska v detaile** — namiesto „Arbeitsplatz: 5388 - PORTALKA FG 3010 CNC" len „5388 - PORTALKA FG 3010 CNC", o 50 % väčšie písmo (24 px) a tučne.

**Výkres a Components / BOM** — pri prechode myšou sa nadvihnú ako ostatné tlačidlá.

**Stavové tlačidlá** — obrázky v pozadí (Stretnutia, Prestávka, Čakanie, Výroba, Údržba…) sú o 30 % výraznejšie.

## 3.2.12 — 2026-09-26

**Okná appky — jeden vzhľad pre všetky** (BOM, potvrdenia, záznamy, správy, zmena používateľa, hlásenia chýb…). Spodná lišta s tlačidlami bola priehľadná a tlačidlá *Zavrieť* / *Načítať všetko* sa strácali na pozadí. Teraz má každé okno plnú bielu plochu, tenký tmavomodrý rám, tmavomodrú hlavičku s bielym nadpisom (ako Dokumentácia) a svetlú spodnú lištu so zreteľnými tlačidlami (biele s tmavomodrým rámom, pod myšou sa vyplnia). Hlavná akcia je plná tmavomodrá, súhlas zelený, odmietnutie červené; hlásenia chyby / varovania majú farebnú ikonu a pruh pod hlavičkou. Rozbaľovacie zoznamy majú tenký tmavomodrý obrys. Rovnako zladené aj vlastné okná skriptu (Výkresy, Popis operácie, HF menu, Nastavenia).

**Popis operácie** — v okne sa text delí na riadky pri každej čiarke (bez medzier navyše; desatinná čiarka ako 0,8 ostáva). Okno je v strede obrazovky a veľké podľa textu; dlhý text má vpravo posuvník.

**SAP ČASY** — priestorové koláče s výškou (naklonený prstenec so stenou, leskom a tieňom). Farby sedia s legendou: Hotovo modrá, Zostáva sivomodrá. Pri prechode myšou sa karta nadvihne ako tlačidlá a koláč sa zdvihne nad svoj tieň. Koláč appky ostáva pod ním (neviditeľný), údaje sa nemenia.

**Rámy kariet** — aj vnútorné karty (SAP časy, paralelné procesy, dlaždice Výkres a BOM, tlačidlá panela PDA, Graf / Tabuľka, stavové tlačidlá) majú jemný tmavomodrý rám. Oprava 3.2.11: tri karty úvodnej obrazovky (Osobný stav, Vyhľadať zákazku, Pracoviská) mali stále svetlý rám.

**Pätička** — bez slova „Jaro": *PDA App Extension · verzia 3J (…) · HF Slovakia*.

## 3.2.11 — 2026-09-26

**Rám kariet tmavomodrý** (#13315c, ako hlavička Dokumentácia) — všetky samostatné karty sú jasne ohraničené aj na rušnom pozadí; vnútorné položky ostávajú bez neho.

**Zákazka a materiál** — obrázok v pozadí o 25 % jemnejší (biely závoj), text sa číta lepšie.

**Vyhľadať zákazku** — pole už nie je cez celú obrazovku, má šírku na 20 znakov.

## 3.2.10 — 2026-09-26

**Oprava: nadpis POPIS OPERÁCIE** sa znovu zobrazuje. Od 3.2.3 ho zakrýval pás hlavičky — nadpis leží v inej vrstve než „Celý text ›" a plná farba pásu ho prekryla (kým bol pás priehľadný, text presvital).

## 3.2.9 — 2026-09-26

**Zákazka a materiál: nové pozadie karty** — namiesto technickej kresby fotografia obrobku vpravo dole na bielom (JPEG 1100×619 / 26 kB, `karta-zakazka.jpg`, pôvodný obrázok `karta-zakazka.png`). Obrázok vypĺňa kartu a obrobok ostáva v pravom dolnom rohu; text sa na svetlom pozadí dobre číta.

## 3.2.8 — 2026-09-26

**Pracoviská vo farbe oblasti** — každé pracovisko v rozbalenej oblasti má tenký rám (1 px) vo farbe svojej oblasti; pod myšou je rám plnou farbou s jemným nádychom a tieňom.

## 3.2.7 — 2026-09-26

**Rámy kariet na úvodnej obrazovke** — *Osobný stav*, *Vyhľadať zákazku* a *Pracoviská* majú rovnaký rám, plochu a tieň ako ostatné karty. Hlavne počas načítavania boli bez rámu a pod „čakacím" závojom aplikácie vyzerali ako vyblednuté obdĺžniky. Závoj je jemnejší a má zaoblenie karty. *Osobný stav* je odteraz samostatná karta (predtým priehľadný obal).

## 3.2.6 — 2026-09-26

**Dokumentácia podľa vzoru** — tmavomodrá hlavička s ikonou dokumentu; *Výkres* a *Components / BOM* ako dve veľké dlaždice s ikonou v kruhu, podnadpisom a šípkou (Výkres svetlomodrý, BOM biely); Výkres ukazuje ďalej číslo výkresu a revíziu. Zvislá deliaca čiara oddeľuje prepínač *Stroj* a *Operation Complete*. Len vzhľad — obe tlačidlá robia to isté ako predtým.

**Prihlasovacia stránka — poistka** — tlačidlo Azure a odkaz „meno a heslo" sa hľadajú aj podľa textu a toho, že patria stránke Login, nielen podľa názvu prvku, takže vzhľad z 3.2.5 sa uplatní aj vtedy, keď má prvok na stránke iný názov.

## 3.2.5 — 2026-09-26

**Prihlasovacia stránka** — *Prihlásenie cez Azure* je väčšie modré tlačidlo s logom Microsoft, tieňom a nadvihnutím pod myšou. *Prihlásenie pomocou mena a hesla* je namiesto bieleho textu, ktorý sa na svetlom pozadí strácal, biele tlačidlo s tmavým textom. Len vzhľad — prihlasovanie funguje rovnako.

## 3.2.4 — 2026-09-26

**Hrubší farebný rám oblastí** na úvodnej obrazovke — 3 px namiesto 1,5 px; pruh vľavo ostáva 8 px.

## 3.2.3 — 2026-09-26

**Popis operácie: plne biela karta** — rovnako ako Zákazka a materiál, bez priehľadnosti. Hlavička ostala rovnaká; pod myšou má karta jemný modrý nádych.

## 3.2.2 — 2026-09-26

**Zákazka a materiál: biela karta so strojárskou grafikou** — len táto karta je plne biela (nie priehľadná) a vpravo dole má jemnú technickú kresbu: prírubu s otvormi na roztečnej kružnici, osi a kóty ako na výkrese. Hlavička ostala rovnaká ako na ostatných kartách.

## 3.2.1 — 2026-09-26

**Výraznejšie farby oblastí** — pruh vľavo 8 px (predtým 5), celý rám karty jemne vo farbe oblasti a tieň s jej nádychom; pod myšou a pri rozbalenej oblasti je rám ešte sýtejší.

## 3.2.0 — 2026-09-26

**Úvodná obrazovka: oblasti pracovísk vo farbe** (podľa vzoru) — každá oblasť má vlastnú farbu: pruh vľavo, dlaždicu ikony aj ikonu. Assembly modrá, Welding oranžová, Machining tmavomodrá, Quality Control zelená. Oblasť, ktorá sa objaví navyše, dostane automaticky ďalšiu farbu (fialová, ružová, zlatá, …) — dve nové oblasti nikdy nemajú rovnakú. Nové ikony v jednotnom štýle: kľúč, plameň, ozubené koleso, lupa, pre inú oblasť krabica. Rozbalená oblasť má jemný nádych svojej farby a rám v nej.

## 3.1.7 — 2026-09-26

**Úvodná obrazovka: Pracoviská stále rozbalené** — šípka na zbalenie panela *Pracoviská* je skrytá a obsah je vždy zobrazený. Len vzhľad (CSS), nič sa neklika; ostatné panely sa nemenia.

## 3.1.6 — 2026-09-26

**Hlavičky cez celú šírku karty** — pás hlavičky pri Dokumentácii, Zákazke a materiáli, SAP časoch a Stave operácie teraz siaha od okraja po okraj karty, začína hneď hore a má zaoblené oba horné rohy (ako Popis operácie). Prebíjali ho staršie pravidlá s iným odsadením — v Dokumentácii pravidlo s ID (6 px vľavo), v SAP časoch a Zákazke iné odsadenie karty. Paralelné procesy dostali rovnakú hlavičku (predtým ostali v starom malom písme, lebo ich produkcia štýluje selektorom s dvoma ID).

**Menej priehľadné karty** — priehľadnosť znížená o 30 % (karty 24 % → 17 %, vnútorné okienka 10 % → 7 %, horný pruh, pätka a ostatné panely úmerne).

## 3.1.5 — 2026-09-26

**Zaoblené hlavičky** — pás hlavičky má zaoblené oba horné rohy rovnako ako karta (predtým mal vľavo hore ostrý roh, lebo sa spoliehal na orezanie kartou, ktoré nie všade fungovalo). Platí pre všetky lišty: Stav operácie, Dokumentácia, Zákazka a materiál, SAP časy, Popis operácie, Paralelné procesy, Prehľad zdrojov, PDA, Pracovný zoznam.

## 3.1.4 — 2026-09-26

**Nové pozadie — pozadie-5** (skúška). Obrobky v strede a vpravo, sklenené pásy, vľavo hore len značka HF, bez textov. Vložené ako JPEG 1672×941 / 105 kB (`pozadie-5.jpg`, pôvodný návrh `pozadie-5.png`). Návrat k predošlému pozadiu: `archiv/pda-suite-3j_3.1.3.user.js` (pozadie-3j).

## 3.1.3 — 2026-09-25

**Vlastný repozitár** — verzia 3J má odteraz samostatný repozitár [JaroTvarozek/PDA-3J](https://github.com/JaroTvarozek/PDA-3J). Adresa aktualizácií ukazuje naň, takže Tampermonkey novú verziu stiahne sám. Všetky staršie verzie sú v `archiv/`. Vzhľad ani funkcie sa nemenili.

## 3.1.2 — 2026-09-25

**Pracovný zoznam až po spodok** — posuvný zoznam zákaziek siaha až nad pätku, pod ním
ostáva len box *Prehľad zdrojov*. Produkcia mala pevných 460 px; výška sa teraz počíta
len z vecí, ktoré od výšky zoznamu nezávisia (horný okraj zoznamu, box pod ním, pätka),
takže neposkakuje. Najnižšia výška 260 px, pri chybe merania sa vráti pevných 460 px.

## 3.1.1 — 2026-09-25

**Hlavičky ako podfarbený pás** (ako v návrhu) — hlavička každej karty je pás cez celú
šírku karty, jemne modrejší než karta a od obsahu oddelený tenkou bielou čiarou.
Platí pre Stav operácie, Dokumentácia, Zákazka a materiál, SAP časy, Popis operácie,
Paralelné procesy, Prehľad zdrojov, PDA a Pracovný zoznam. Pás je o odtieň tmavší
a pod ním je viditeľná tenká modrosivá čiara.

**Pracovný zoznam** — vybraný záznam je znovu tmavomodrý (nové farby plôch ho prebili
a ostával svetlý).

## 3.1.0 — 2026-09-25

**Pozadie**
- z obrázka pozadia odstránené texty „SLOVAKIA", „Better parts. A cleaner tomorrow."
  a „People / Technology / Process / Results" — ostala len značka **HF**
- plochy dopočítané z okolitých pixelov (bez viditeľných záplat), pozadie jemne dofarbené
  do modra ako v návrhu; zdroj: `pozadie-3j.jpg` (50 kB, vložené v skripte)

**Karty**
- namiesto bielych plôch jemne sivomodré polopriehľadné „sklo" — pozadie cez ne presvitá
- vnútorné okienka (jednotlivé SAP časy, paralelné procesy, pilulky v zozname,
  tlačidlá panela PDA) tiež jemne sfarbené, nie biele
- horný pruh a pätka priehľadné

**Hlavičky** (Stav operácie, Dokumentácia, Zákazka a materiál, SAP časy, Popis operácie,
Paralelné procesy, Pracovný zoznam, PDA) — všetky jednotne ako v návrhu: 16 px,
polotučne, veľkými písmenami, rovnaké písmo

**Zákazka a materiál** — popisok vľavo, hodnota vpravo, medzi riadkami tenká čiara,
písmo Segoe UI; veľký názov materiálu oddelený čiarou

**SAP časy** — o niečo širšie (pravý stĺpec mriežky 1,15 : 0,85)

## 3.0.0 — 2026-09-25

Vznik verzie 3J z produkčného buildu **2.3.0** (Daniel Gabriš).
Všetky funkcie zachované; zmeny sú len vo vzhľade (modul **Dizajn 3J (Jaro)**,
dá sa vypnúť v ⚙ → vráti vzhľad Production 2.3.0).

**SAP ČASY**
- každý graf (Setup / Machine / Labor) vo vlastnej orámovanej karte
- nadpis grafu normálnym písmom a čas hneď pod ním, nad koláčom
- väčší koláč (rastie so šírkou okna, 110–180 px), v strede percento a „hotovo"
- pod koláčom legenda *Hotovo X % / Zostáva Y %*
- karty sa rozťahujú vedľa seba podľa šírky; sekcia ako polopriehľadná karta

**PARALELNÉ PROCESY**
- každá položka ako biela karta namiesto zelenej plochy
- pred textom ikona — zelená dlaždica so šípkami
- stav tučne, meno pracovníka pod ním; veľký čas; *Zastaviť* ako červené tlačidlo

**Celá stránka**
- mimo kariet priehľadná — veľké panely (Osobný stav, pracovisko) nemajú vlastnú
  plochu, medzi kartami presvitá pozadie HF Slovakia

Dáta a čísla sa nemenia — percentá sa počítajú z časov, ktoré zobrazuje aplikácia.
