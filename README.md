# HF PDA scripts

Vylepšenia pre aplikáciu **PDA** bežiacu na `https://hf.simplifier.cloud/appDirect/PDA/`.

Skripty nemenia aplikáciu na serveri — bežia až v prehliadači a dopĺňajú do nej
UI: slovenské popisky, farebné stavové tlačidlá, vyhľadávanie zákaziek naprieč
pracoviskami, načítanie výkresov, obsluhu skenera a čítačky kariet.

---

## Inštalácia — stačí jeden skript

**[▶ Inštalovať PDA Suite](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/pda-suite.user.js)**

`pda-suite.user.js` obsahuje **všetkých 8 modulov v jednom súbore**. Nainštaluješ
ho raz a jednotlivé moduly si potom zapínaš a vypínaš v nastaveniach — nemusíš
nič odinštalovávať ani doinštalovávať.

### Predpoklady (raz za počítač)

1. **Tampermonkey** —
   [Chrome Web Store](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
2. **Povoliť používateľské skripty** — `chrome://extensions` → Tampermonkey →
   *Podrobnosti* → zapnúť **„Allow user scripts"**.
   ⚠️ Bez tohto kroku sa skripty tvária, že bežia, ale nerobia nič.

### Overenie

Otvor PDA a daj **Ctrl + F5**. Vpravo dole sa objaví **⚙ ozubené koliesko** —
to je panel nastavení.

---

## Nastavenia

Panel otvoríš **ozubeným kolieskom vpravo dole**, alebo cez ikonu Tampermonkey →
*Nastavenia PDA Suite*.

| Modul | Predvolene | Poznámka |
|---|---|---|
| Vylepšená hlavička | zapnuté | slovenské popisky, väčšie meno, zvýraznené odhlásenie |
| Farebné stavové tlačidlá | zapnuté | výroba zelená, prestoj oranžový, chyba červená |
| Vyhľadávanie zákaziek | zapnuté | hľadá naprieč všetkými pracoviskami naraz |
| Blokovanie tlačidla Späť | zapnuté | zabráni nechcenému vypadnutiu z aplikácie |
| Panel prepínania používateľov | vypnuté | má zmysel len na zdieľanom termináli |
| Čiarový skener a RFID karty | vypnuté | vyžaduje hardvér |
| Tlačidlo výkresu | vypnuté | vyžaduje firemnú sieť a Excel |
| Ladiaci výpis do konzoly | vypnuté | len na hľadanie chýb |

Zmena sa uloží tlačidlom **Uložiť a obnoviť stránku**.

### Citlivé údaje

**Heslá kolegov a ID kariet nie sú v kóde** a nikdy sem nepatria — repozitár je
verejný. Zadávajú sa v paneli nastavení v sekcii *Používatelia* a ukladajú sa
len lokálne do úložiska Tampermonkey na danom počítači. To isté platí pre adresu
a kľúč služby výkresov.

---

## Ako to vnútri funguje

Celý skript stojí na troch spoločných častiach, ktoré zdieľajú všetky moduly:

- **`DomWatch`** — jeden `MutationObserver` pre celú stránku. Aplikácia je SAP
  UI5 a prekresľuje si DOM sama, takže moduly musia svoje prvky dokladať znova.
  Všetky prekreslenia sa zlučujú do jednej dávky.
- **`XhrBus`** — jedno odpočúvanie volaní na `/client/1.0/executeBO`. Prepisuje
  `XMLHttpRequest` v kontexte stránky a odpovede rozposiela modulom, ktoré sa
  prihlásili.
- **`UserSwitch`** — prepnutie používateľa cez SAP dialóg. Používa ho bočný
  panel aj čítačka kariet, preto je mimo oboch modulov.

Moduly medzi sebou komunikujú cez objekt `shared` (index zákaziek, aktuálna
operácia), nie cez globálne premenné na `window`.

---

## Pridanie nového vylepšenia

1. Napísať funkciu `modNiecoNove()` v sekcii **3. MODULY**
2. Pridať jeden riadok do zoznamu `MODULES` (`id`, `name`, `desc`, `def`, `run`)
3. Zvýšiť `@version` v hlavičke — **inak Tampermonkey aktualizáciu nestiahne**
4. Commit a push

Modul sa automaticky objaví v paneli nastavení. Ak niektorý modul spadne,
ostatné bežia ďalej — spúšťajú sa každý vo vlastnom `try/catch`.

---

## Staršie samostatné skripty

Pôvodných 8 samostatných skriptov v koreni repozitára (`enhanced-header.user.js`
a spol.) tu zostáva pre prípad, že ich má niekto ešte nainštalované. **Pre nové
inštalácie použi `pda-suite.user.js`.**

Ak prechádzaš zo starých skriptov: nainštaluj suite a staré v Tampermonkey
zmaž, inak by niektoré veci bežali dvakrát.

Pôvodný základ pochádza z repozitára `Dan1elG94/HF-Slovakia-PDA-scripts`.
