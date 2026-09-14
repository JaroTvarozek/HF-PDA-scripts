# HF PDA scripts

Používateľské skripty (userscripty) pre aplikáciu **PDA** bežiacu na
`https://hf.simplifier.cloud/appDirect/PDA/`.

Skripty nemenia samotnú aplikáciu na serveri — bežia až v prehliadači a dopĺňajú
do nej UI (slovenské popisky, farebné stavové tlačidlá, vyhľadávanie zákaziek
naprieč pracoviskami, načítanie výkresov a pod.).

Pôvodný základ pochádza z repozitára `Dan1elG94/HF-Slovakia-PDA-scripts`.
Táto kópia je samostatná — skripty sa aktualizujú odtiaľto.

---

## Inštalácia

### 1. Tampermonkey

Chrome sám userscripty spúšťať nevie, treba naň správcu:

https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo

### 2. Povoliť spúšťanie skriptov ⚠️

Novšie verzie Chromu majú userscripty vypnuté, kým sa výslovne nepovolia:

1. `chrome://extensions`
2. dlaždica **Tampermonkey** → **Podrobnosti**
3. zapnúť **„Povoliť používateľské skripty"** (*Allow user scripts*)

### 3. Nainštalovať skripty

Klik na odkaz → otvorí sa inštalačná stránka Tampermonkey → **Inštalovať**.

| Skript | Čo robí |
|---|---|
| [enhanced-header](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/enhanced-header.user.js) | Slovenské popisky v hornej lište, väčšie meno používateľa, zvýraznené odhlásenie |
| [color-and-reorder-buttons](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/color-and-reorder-buttons.user.js) | Farebné a zoradené stavové tlačidlá (výroba / prestoj / chyba) |
| [crosscenter-search](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/crosscenter-search.user.js) | Vyhľadávanie zákaziek naprieč všetkými pracoviskami |
| [prevent-back-button](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/prevent-back-button.user.js) | Zablokuje tlačidlo Späť, aby aplikácia nevypadla |
| [users-panel](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/users-panel.user.js) | Bočný panel na rýchle prepnutie používateľa — **len pre zdieľaný dielenský terminál** |
| [input-listener](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/input-listener.user.js) | Čítanie čiarového skenera a RFID kariet — **vyžaduje hardvér** |
| [order-drawing-button](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/order-drawing-button.user.js) | Číslo výkresu k zákazke + prehliadač výkresov — **vyžaduje firemnú sieť** |
| [test](https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/test.user.js) | Ladiaci výpis sieťovej komunikácie do konzoly |

### 4. Overenie

Otvoriť `https://hf.simplifier.cloud/appDirect/PDA/` a dať **Ctrl + F5**.
Na ikone Tampermonkey sa objaví číslo = počet aktívnych skriptov na stránke.

---

## Ako to funguje

Všetky skripty stoja na dvoch princípoch:

- **`MutationObserver`** — aplikácia je SAP UI5, ktoré si DOM prekresľuje samo,
  takže skripty sledujú zmeny a svoje prvky dokladajú znova.
- **Odpočúvanie XHR** — volania na `/client/1.0/executeBO` sa čítajú cez
  prepísané `XMLHttpRequest.prototype.open/send`; odtiaľ sa berú dáta
  o zákazkách a o práve otvorenej operácii.

Skripty si medzi sebou podávajú dáta cez globálne premenné
(`window.PDA_ORDERS_INDEX`, `window.PDA_USERS`, `window.PDA_switchToUser`,
`window.PDA_CURRENT_OPERATION`).

---

## Úpravy

Po zmene skriptu treba **zvýšiť `@version`** v hlavičke, inak Tampermonkey
aktualizáciu nestiahne. Potom stačí commit a push — Tampermonkey si novú verziu
natiahne sám (kontroluje periodicky, ručne cez *Dashboard → Pomôcky → Aktualizovať*).
