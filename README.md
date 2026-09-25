# PDA Suite 3J — Jarova verzia

Samostatná verzia PDA Suite, v ktorej **Jaroslav Tvarožek** upravuje dizajn.
Vychádza z produkčného buildu **2.3.0** (Daniel Gabriš, repozitár
[PDA-D_J-NEW](https://github.com/JaroTvarozek/PDA-D_J-NEW), priečinok `production/`) —
**všetky funkcie, volania a dáta ostávajú rovnaké**, mení sa len vzhľad.

| | |
|---|---|
| Skript | `pda-suite-3j.user.js` |
| Názov v Tampermonkey | **PDA Suite 3J (Jaro · HF Slovakia)** |
| Výpis v konzole (F12) | `[PDA 3J]` |
| **Inštalácia / aktualizácia** | **[▶ Inštalovať PDA Suite 3J](https://github.com/JaroTvarozek/PDA-3J/raw/refs/heads/main/pda-suite-3j.user.js)** |
| Zmeny po verziách | [`ZMENY.md`](ZMENY.md) |
| Staršie verzie | [`archiv/`](archiv/) — každá vydaná verzia ako samostatný súbor |

Verzie začínajú na **3.0.0** („3J"). Tampermonkey porovnáva len čísla, preto je
„J" v názve skriptu, nie v čísle verzie. Skript sa aktualizuje sám z tohto repozitára.

## Inštalácia a prepínanie

1. Otvor [inštalačnú adresu](https://github.com/JaroTvarozek/PDA-3J/raw/refs/heads/main/pda-suite-3j.user.js) → **Install / Reinstall**
2. V Tampermonkey Dashboard nechaj **zapnutý iba jeden** zo skriptov PDA Suite:

| Skript | pri používaní 3J |
|---|---|
| PDA Suite (HF Slovakia) 1.25.4 | vypnutý |
| PDA Suite NEW Design 2.x | vypnutý |
| PDA Suite Production 2.3.0 | vypnutý |
| **PDA Suite 3J** | **zapnutý** |

3. V PDA **Ctrl + F5**

Na návrat k Production stačí prepnúť prepínače opačne — nič sa nemaže.
Na návrat k staršej 3J stačí nainštalovať súbor z `archiv/`.

⚠️ **Nastavenia sa neprenášajú** — Tampermonkey ich drží pre každý skript zvlášť.
Prenos: v doterajšom skripte ⚙ → *Súbor s nastaveniami* → **Stiahnuť**, v 3J → **Nahrať**.
Bez toho 3J nepozná adresu Excelu a nenájde výkres.

## Čo je v repozitári

| | |
|---|---|
| `pda-suite-3j.user.js` | aktuálna verzia skriptu |
| `archiv/` | všetky vydané verzie `pda-suite-3j_<verzia>.user.js` |
| `ZMENY.md` | čo sa v ktorej verzii zmenilo |
| `pozadie-3j.jpg` | pozadie (bez textov, len značka HF) — v skripte je vložené, tu je len na úpravy |

Úpravy vzhľadu sú v module **Dizajn 3J (Jaro)** (sekcia 3.20 v skripte) — dá sa vypnúť
v ⚙ a vtedy má 3J vzhľad Production 2.3.0.

## Postup pri novej verzii

1. upraviť `pda-suite-3j.user.js`, zvýšiť `@version`
2. `node --check pda-suite-3j.user.js`
3. skopírovať do `archiv/pda-suite-3j_<verzia>.user.js`, zapísať do `ZMENY.md`
4. commit + push — Tampermonkey si novú verziu stiahne sám (alebo Reinstall cez inštalačnú adresu)
