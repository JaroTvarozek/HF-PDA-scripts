// ==UserScript==
// @name         PDA Suite (HF Slovakia)
// @namespace    http://tampermonkey.net/
// @version      1.7.0
// @description  Vsetky vylepsenia PDA v jednom skripte + panel na zapinanie a vypinanie jednotlivych modulov
// @author       Gabris, Tvarozek
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/pda-suite.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/pda-suite.user.js
// @match        https://hf.simplifier.cloud/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @run-at       document-start
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      172.16.77.134
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

/*
 * ============================================================================
 *  PDA Suite - jeden skript namiesto osmich samostatnych
 * ============================================================================
 *
 *  Ako to funguje:
 *    - kazde vylepsenie je samostatny MODUL (dole v sekcii MODULY)
 *    - ktore moduly bezia sa nastavuje v paneli: ozubene koliesko vpravo dole
 *      na stranke PDA, alebo cez ikonu Tampermonkey -> "Nastavenia PDA Suite"
 *    - nastavenia sa ukladaju lokalne v Tampermonkey (GM storage), takze
 *      prezijeu aktualizaciu skriptu a NIE su sucastou repozitara
 *
 *  Citlive udaje (hesla kolegov, adresa sluzby vykresov) sa zadavaju
 *  v tom paneli - zamerne nie su v kode, lebo repozitar je verejny.
 *
 *  Pridanie noveho vylepsenia = pridat funkciu + jeden riadok do zoznamu
 *  MODULES. Nic ine netreba.
 * ============================================================================
 */

(function () {
    'use strict';

    const LOG = '[PDA Suite]';
    const W = unsafeWindow;

    /* ========================================================================
     *  1. ULOZISKO NASTAVENI
     * ====================================================================== */

    const KEY_MODULES = 'pda_modules_v1';
    const KEY_USERS = 'pda_users_v1';
    const KEY_PDM = 'pda_pdm_v1';
    const KEY_EXCEL = 'pda_excel_v1';
    const KEY_GROUPS = 'pda_groups_v1';

    function loadJson(key, fallback) {
        try {
            const raw = GM_getValue(key, null);
            if (raw === null || raw === undefined) return fallback;
            return typeof raw === 'string' ? JSON.parse(raw) : raw;
        } catch (e) {
            console.warn(LOG, 'nepodarilo sa nacitat nastavenie', key, e);
            return fallback;
        }
    }

    function saveJson(key, value) {
        try {
            GM_setValue(key, JSON.stringify(value));
        } catch (e) {
            console.warn(LOG, 'nepodarilo sa ulozit nastavenie', key, e);
        }
    }

    const settings = {
        modules: loadJson(KEY_MODULES, {}),
        users: loadJson(KEY_USERS, []),
        pdm: loadJson(KEY_PDM, { base: 'http://172.16.77.134:9000', key: '' }),
        // url prazdna = subor sa vybera rucne cez tlacidlo "Vybrať Excel"
        excel: loadJson(KEY_EXCEL, { url: '', colOrder: 'H', colDrawing: 'AH', colVersion: 'AI' }),
        // kategorie pracovisk pre uvodny prehlad: Nazov|Podtitul|VZOR,VZOR,...
        groups: loadJson(KEY_GROUPS, [
            'Assembly|Finálna montáž a podzostavy|MONTAZ,ASSEMBLY,PODZOST,MONT',
            'Welding|Zváranie a príprava|ZVAR,TIG,MIG,WELD',
            'Machining|CNC a konvenčné obrábanie|CNC,FREZ,SUSTR,BRUS,LMS,HMS,K-TEC,VRTA,PILA,HOBL,HEDELL',
            'Quality Control|Kontrola a meranie|OTK,KONTROL,MERAN,QC,KVALIT',
        ].join('\n')),
    };

    // "H" -> 7, "AH" -> 33 (vracia 0-based index stlpca ako ho vidi XLSX)
    function colToIndex(letters, fallbackIndex) {
        const s = String(letters || '').trim().toUpperCase();
        if (!/^[A-Z]{1,3}$/.test(s)) return fallbackIndex;
        let n = 0;
        for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
    }

    function isModuleOn(mod) {
        const stored = settings.modules[mod.id];
        return stored === undefined ? mod.def : !!stored;
    }

    /* ========================================================================
     *  2. ZDIELANE POMOCKY
     * ====================================================================== */

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    function onReady(fn) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', fn, { once: true });
        } else {
            fn();
        }
    }

    function waitFor(checkFn, { interval = 150, timeout = 10000 } = {}) {
        return new Promise((resolve, reject) => {
            const start = Date.now();
            const timer = setInterval(() => {
                let result = null;
                try { result = checkFn(); } catch (e) { /* ignore */ }
                if (result) {
                    clearInterval(timer);
                    resolve(result);
                } else if (Date.now() - start > timeout) {
                    clearInterval(timer);
                    reject(new Error('waitFor timeout'));
                }
            }, interval);
        });
    }

    // --- pristup k SAP UI5 controlom (aplikacia bezi v kontexte stranky) ---

    function getControl(controlId) {
        try {
            if (W.sap && W.sap.ui && W.sap.ui.getCore) return W.sap.ui.getCore().byId(controlId);
        } catch (e) { /* ignore */ }
        return null;
    }

    function resolveControl(el) {
        while (el && el.nodeType === 1) {
            try {
                if (W.jQuery && typeof W.jQuery(el).control === 'function') {
                    const arr = W.jQuery(el).control();
                    if (arr && arr.length && arr[0]) return arr[0];
                }
                if (W.sap && W.sap.ui && W.sap.ui.core && W.sap.ui.core.Element &&
                    typeof W.sap.ui.core.Element.closestTo === 'function') {
                    const c = W.sap.ui.core.Element.closestTo(el);
                    if (c) return c;
                }
                if (el.id && W.sap && W.sap.ui && W.sap.ui.getCore) {
                    const c = W.sap.ui.getCore().byId(el.id);
                    if (c) return c;
                }
            } catch (e) { /* ignore */ }
            el = el.parentElement;
        }
        return null;
    }

    function pressElement(domEl) {
        const control = resolveControl(domEl);
        if (control && typeof control.firePress === 'function') {
            control.firePress();
            return;
        }
        domEl.click();
    }

    function setControlValue(controlId, value) {
        try {
            const control = getControl(controlId);
            if (control && typeof control.setValue === 'function') {
                control.setValue(value);
                if (typeof control.fireLiveChange === 'function') control.fireLiveChange({ value });
                if (typeof control.fireChange === 'function') control.fireChange({ value, newValue: value });
                return true;
            }
        } catch (e) {
            console.warn(LOG, 'setControlValue zlyhalo pre', controlId, e);
        }
        const inputEl = document.getElementById(controlId + '-inner');
        if (inputEl) {
            const nativeSetter = Object.getOwnPropertyDescriptor(W.HTMLInputElement.prototype, 'value').set;
            nativeSetter.call(inputEl, value);
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        return false;
    }

    /* ------------------------------------------------------------------
     *  Jeden spolocny sledovac DOM namiesto piatich samostatnych.
     *  Aplikacia je SAP UI5 a prekresluje sa sama, takze moduly musia
     *  svoje prvky dokladat opakovane. Vsetko sa zbiera do jednej davky.
     * ---------------------------------------------------------------- */
    const DomWatch = (function () {
        const callbacks = [];
        let started = false;
        let scheduled = false;

        function flush() {
            scheduled = false;
            for (const fn of callbacks) {
                try { fn(); } catch (e) { console.warn(LOG, 'chyba v DOM callbacku', e); }
            }
        }

        function schedule() {
            if (scheduled) return;
            scheduled = true;
            setTimeout(flush, 120);
        }

        function start() {
            if (started || !document.body) return;
            started = true;
            new MutationObserver(schedule).observe(document.body, {
                childList: true, subtree: true, characterData: true,
            });
        }

        return {
            add(fn) {
                callbacks.push(fn);
                onReady(() => { start(); schedule(); });
            },
            poke: schedule,
        };
    })();

    /* ------------------------------------------------------------------
     *  Jedno spolocne odpocuvanie sietovej komunikacie namiesto troch.
     *  Prepisuje XMLHttpRequest v kontexte stranky a rozposiela vysledky
     *  vsetkym prihlasenym modulom.
     * ---------------------------------------------------------------- */
    const XhrBus = (function () {
        const TARGET = '/client/1.0/executeBO';
        const listeners = [];
        let patched = false;

        function patch() {
            if (patched) return;
            patched = true;

            const proto = W.XMLHttpRequest.prototype;
            const originalOpen = proto.open;
            const originalSend = proto.send;

            proto.open = function (method, url, ...rest) {
                this.__pdaUrl = url;
                this.__pdaMethod = method;
                return originalOpen.call(this, method, url, ...rest);
            };

            proto.send = function (body) {
                const url = this.__pdaUrl || '';
                if (url.indexOf(TARGET) !== -1) {
                    const xhr = this;
                    xhr.addEventListener('load', function () {
                        let requestJson = null;
                        let responseJson = null;
                        try { requestJson = JSON.parse(body); } catch (e) { /* ignore */ }
                        try { responseJson = JSON.parse(xhr.responseText); } catch (e) { /* ignore */ }

                        const event = {
                            url,
                            requestRaw: body,
                            request: requestJson,
                            response: responseJson,
                            responseRaw: xhr.responseText,
                        };
                        for (const fn of listeners) {
                            try { fn(event); } catch (e) { console.warn(LOG, 'chyba v XHR callbacku', e); }
                        }
                    });
                }
                return originalSend.apply(this, arguments);
            };
        }

        return {
            subscribe(fn) {
                patch();
                listeners.push(fn);
            },
        };
    })();

    /* ------------------------------------------------------------------
     *  Dekodovanie vstupu zo skenera.
     *  Skener posiela cisla ako slovenske znaky (SK klavesnica), prvy znak
     *  "J" je marker zariadenia a bodka zastupuje pomlcku.
     * ---------------------------------------------------------------- */
    const SCANNER_CHAR_MAP = {
        '+': '1', 'ľ': '2', 'š': '3', 'č': '4', 'ť': '5',
        'ž': '6', 'ý': '7', 'á': '8', 'í': '9', 'é': '0',
    };

    function decodeScannerInput(raw) {
        // POZOR: prvy znak sa odstrani LEN ak je to naozaj marker "J".
        // (v povodnych skriptoch sa na jednom mieste odrezaval vzdy, co
        //  rucne napisanemu cislu zjedlo prvu cislicu)
        const stripped = raw.length > 0 && raw[0] === 'J' ? raw.slice(1) : raw;

        let out = '';
        for (const ch of stripped) {
            if (ch === '.') out += '-';
            else if (Object.prototype.hasOwnProperty.call(SCANNER_CHAR_MAP, ch)) out += SCANNER_CHAR_MAP[ch];
            else out += ch;
        }
        return out;
    }

    function padOperationPart(value) {
        const i = value.indexOf('-');
        if (i === -1) return value;
        return value.slice(0, i) + '-' + value.slice(i + 1).padStart(4, '0');
    }

    /* ------------------------------------------------------------------
     *  Prepinanie pouzivatela - zdielana sluzba.
     *  Pouziva ju aj bocny panel, aj citacka kariet, takze je tu zvlast
     *  a nie je viazana na to, ci je panel zapnuty.
     * ---------------------------------------------------------------- */
    const UserSwitch = (function () {
        const CHANGE_USER_SUFFIX = 'Button_ChangeUser';
        const USER_COMBO_ID = 'Popups--User_ComboBox';
        const PASSWORD_INPUT_ID = 'Popups--EmployeeNo_Input';
        const CONFIRM_BUTTON_ID = 'Popups--UserDialog_Button_close';
        const DIALOG_WAIT_TIMEOUT = 10000;
        const BEFORE_CONFIRM_DELAY = 100;

        async function selectComboBoxItemByText(controlId, text, { timeout = 5000 } = {}) {
            let control;
            try {
                control = await waitFor(() => {
                    const c = getControl(controlId);
                    return c && typeof c.getItems === 'function' && c.getItems().length > 0 ? c : null;
                }, { timeout });
            } catch (e) {
                console.warn(LOG, 'polozky v zozname pouzivatelov sa nenacitali vcas', e);
                return false;
            }

            const items = control.getItems();
            const match =
                items.find((it) => it.getText() === text) ||
                items.find((it) => it.getText().toLowerCase() === text.toLowerCase());

            if (!match) {
                console.warn(LOG, 'pouzivatel sa v zozname nenasiel:', text,
                    'dostupni:', items.map((i) => i.getText()));
                return false;
            }

            control.setSelectedItem(match);
            if (typeof control.fireSelectionChange === 'function') control.fireSelectionChange({ selectedItem: match });
            if (typeof control.fireChange === 'function') control.fireChange({ value: match.getText() });
            return true;
        }

        async function fillLoginPopup(username, password) {
            try {
                await waitFor(
                    () => document.getElementById(USER_COMBO_ID + '-inner') &&
                          document.getElementById(PASSWORD_INPUT_ID + '-inner'),
                    { timeout: DIALOG_WAIT_TIMEOUT }
                );

                const selected = await selectComboBoxItemByText(USER_COMBO_ID, username);
                if (!selected) setControlValue(USER_COMBO_ID, username);

                setControlValue(PASSWORD_INPUT_ID, password);
                await sleep(BEFORE_CONFIRM_DELAY);

                const confirmBtn = await waitFor(() => document.getElementById(CONFIRM_BUTTON_ID), { timeout: 5000 });
                pressElement(confirmBtn);
            } catch (e) {
                console.warn(LOG, 'dialog na zmenu pouzivatela sa neobjavil vcas', e);
            }
        }

        return {
            to(username, password) {
                const btn = document.querySelector('[id$="' + CHANGE_USER_SUFFIX + '"]');
                if (!btn) {
                    console.warn(LOG, 'tlacidlo na zmenu pouzivatela sa na obrazovke nenaslo');
                    return;
                }
                pressElement(btn);
                fillLoginPopup(username, password);
            },
            findByCardId(cardId) {
                return settings.users.find((u) => u.cardId && u.cardId === cardId) || null;
            },
        };
    })();

    // zdielany stav medzi modulmi (nahradza povodne window.PDA_* premenne)
    const shared = {
        ordersIndex: [],
        ordersIndexUpdatedAt: null,
        currentOperation: null,
        fillSearchInput: null,   // doplni modul vyhladavania, ak bezi
        intentionalReload: false, // nastavi panel nastaveni pred location.reload()
    };

    /* ========================================================================
     *  3. MODULY
     * ====================================================================== */

    /* ---------------------- 3.1 Vylepsena hlavicka ---------------------- */

    function modEnhancedHeader() {
        const USERNAME_SUFFIX = 'Label_Username-bdi';
        const BUTTONS = [
            { suffix: 'Button_HomeScreen-img', label: 'Pracoviská' },
            { suffix: 'Button_Reporting-img', label: 'Reporty' },
            { suffix: 'Button_Message-img', label: 'Správy' },
            { suffix: 'Button_Schedule-img', label: 'Rozvrh' },
            { suffix: 'Button_Settings-img', label: 'Admin' },
            { suffix: 'Button_ChangeUser-img', label: 'Zmena používateľa' },
            { suffix: 'Button_Logout-img', label: 'Odhlásenie' },
        ];
        const LOGOUT_BUTTON_SUFFIX = 'Button_Logout-inner';

        function apply() {
            document.querySelectorAll('[id$="' + USERNAME_SUFFIX + '"]').forEach((el) => {
                if (el.dataset.pdaUsernameStyled) return;
                el.style.fontSize = '1.8rem';
                el.style.fontWeight = 'bold';
                el.style.color = '#1f1f1f';
                el.dataset.pdaUsernameStyled = '1';
            });

            BUTTONS.forEach(({ suffix, label }) => {
                document.querySelectorAll('[id$="' + suffix + '"]').forEach((img) => {
                    if (img.dataset.pdaTextAdded) return;
                    const span = document.createElement('span');
                    span.className = 'sapMBtnContent';
                    span.style.marginRight = '0.6rem';
                    span.innerHTML = '<bdi>' + label + '</bdi>';
                    img.insertAdjacentElement('afterend', span);
                    img.dataset.pdaTextAdded = '1';
                });
            });

            document.querySelectorAll('[id$="' + LOGOUT_BUTTON_SUFFIX + '"]').forEach((btn) => {
                if (btn.dataset.pdaLogoutStyled) return;
                btn.style.backgroundColor = '#d67a74';
                btn.style.borderRadius = '4px';
                btn.dataset.pdaLogoutStyled = '1';
            });
        }

        DomWatch.add(apply);
    }

    /* ------------------ 3.2 Farebne stavove tlacidla -------------------- */

    function modStatusButtons() {
        const CONTAINER_ID = 'WorkcenterDetail--Order_Status_Flexbox';
        const COLORS = { productive: '#4e9041', downtime: '#d66c37', fault: '#d04040' };
        const PRIORITY = { productive: 0, downtime: 1, fault: 2 };
        const STATUS_MAP = {
            'Výroba': 'productive',
            'Upinanie': 'productive',
            'Programovanie': 'downtime',
            'Upratovanie stola': 'downtime',
            'Meranie v Procese s OTK': 'downtime',
            'Chyba programu': 'fault',
        };

        function textOf(btn) {
            const el = btn.querySelector('.sapMBtnContent bdi, .sapMBtnContent');
            return el ? el.textContent.trim() : '';
        }

        function styleButton(btn) {
            const inner = btn.querySelector('.sapMBtnInner');
            [btn, inner].forEach((el) => el && el.style.setProperty('border-radius', '10px', 'important'));

            const category = STATUS_MAP[textOf(btn)];
            if (!category) return;
            const color = COLORS[category];
            [btn, inner].forEach((el) => {
                if (!el) return;
                el.style.setProperty('background-color', color, 'important');
                el.style.setProperty('border-color', color, 'important');
            });
        }

        function apply() {
            const container = document.getElementById(CONTAINER_ID);
            if (!container) return;
            const buttons = Array.from(container.children).filter((el) => el.classList.contains('statusBtn'));
            if (buttons.length === 0) return;

            buttons.forEach(styleButton);

            const priorityOf = (btn) => {
                const p = PRIORITY[STATUS_MAP[textOf(btn)]];
                return p === undefined ? 1.5 : p;
            };
            const sorted = [...buttons].sort((a, b) => priorityOf(a) - priorityOf(b));
            if (sorted.every((btn, i) => buttons[i] === btn)) return;
            sorted.forEach((btn) => container.appendChild(btn));
        }

        DomWatch.add(apply);
    }

    /* ------------------ 3.3 Blokovanie tlacidla Spat -------------------- */

    function modPreventBack() {
        function lockCurrentState() {
            W.history.pushState(W.history.state, document.title, location.href);
        }

        onReady(() => {
            lockCurrentState();
            W.addEventListener('popstate', lockCurrentState);
            W.addEventListener('beforeunload', (e) => {
                // obnovenie stranky vyvolane panelom nastaveni sa nema pytat
                if (shared.intentionalReload) return;
                e.preventDefault();
                e.returnValue = '';
            });
        });
    }

    /* --------------- 3.4 Vyhladavanie naprieč pracoviskami -------------- */

    function modCrossSearch() {
        const BO_METHOD = 'getOperationListTempForWorkcenters';
        const WORKCENTER_TILE_SELECTOR = '[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]';
        const CONTENT_ID = 'Main--Workcenter_Panel-content';
        const HOME_BUTTON_ID = 'Main--Button_HomeScreen';
        const PANEL_ID = 'Main--Workcenter_Panel';
        const PARENT_SECTION_ID = 'Main--MainPage-cont';
        const TASKS_PANEL_ID = 'Main--Tasks_Panel';
        const SIDEBAR_ID = '__pda_search_sidebar__';
        const UI_ID = '__pda_custom_search_ui__';
        const TILE_ID_PREFIX = 'Main--Workcenter_Toolbar-';
        const LIST_ID_PREFIX = 'Main--List2-';
        const TILE_WAIT_TIMEOUT = 15000;
        const SETTLE_DELAY = 350;

        let opening = false;
        let renderFn = null;
        let lastAutoOpenedKey = null;

        XhrBus.subscribe((ev) => {
            if (!ev.request || ev.request.BOMethod !== BO_METHOD) return;
            if (!ev.response || !ev.response.result) return;

            const workcenters = ev.response.result.aOperationListTempForWorkcenter || [];
            const flat = [];
            workcenters.forEach((wc) => {
                (wc.operationList || []).forEach((op) => {
                    flat.push({
                        workcenter: wc.workcenterDescription,
                        workcenterCode: wc.workcenter,
                        salesOrderNo: op.salesOrderNo,
                        salesOrderItem: op.salesOrderItem,
                        productionOrderNo: op.productionOrderNo,
                        operationNo: op.operationNo,
                        sequenceNo: op.sequenceNo,
                        materialNo: op.materialNo,
                        material: op.material,
                        description: op.description,
                        descriptionShort: op.descriptionShort,
                        status: op.status,
                        confirmed: op.confirmed,
                    });
                });
            });

            shared.ordersIndex = flat;
            shared.ordersIndexUpdatedAt = new Date();
            console.log(LOG, 'index zakaziek naplneny:', flat.length);

            if (renderFn) {
                const input = document.querySelector('#' + UI_ID + ' input.sapMSFI');
                renderFn(input ? input.value : '');
            }
        });

        function getWorkcenterName(tile) {
            const suffix = tile.id.replace(TILE_ID_PREFIX, '');
            const nameEl = document.getElementById('Main--WorkcenterDescription_Label-' + suffix + '-bdi');
            return nameEl ? nameEl.textContent.trim() : tile.id;
        }

        function getListIdForTile(tile) {
            if (!tile.id || tile.id.indexOf(TILE_ID_PREFIX) !== 0) return null;
            return LIST_ID_PREFIX + tile.id.slice(TILE_ID_PREFIX.length);
        }

        function extractOrderText(li) {
            const titleEl = li.querySelector('[id*="Title"][id$="-inner"]');
            return titleEl ? titleEl.textContent.trim() : '';
        }

        function formatProductionOrder(item) {
            return item.productionOrderNo + ' - ' + item.operationNo + ' - ' + item.sequenceNo;
        }

        function getItemKey(item) {
            return [item.workcenter, item.productionOrderNo, item.operationNo, item.sequenceNo].join('|');
        }

        function setStatus(msg) {
            const el = document.getElementById('__pda_status__');
            if (el) el.textContent = msg;
        }

        function fireListItemPress(li) {
            const itemControl = resolveControl(li);
            if (!itemControl) return false;

            let listControl = itemControl;
            while (listControl && typeof listControl.fireItemPress !== 'function') {
                listControl = listControl.getParent && listControl.getParent();
            }
            if (!listControl) return false;

            let itemForEvent = itemControl;
            while (itemForEvent && itemForEvent.getParent && itemForEvent.getParent() !== listControl) {
                itemForEvent = itemForEvent.getParent();
            }
            if (!itemForEvent) itemForEvent = itemControl;

            listControl.fireItemPress({ listItem: itemForEvent, srcControl: itemForEvent });
            return true;
        }

        async function openItem(item) {
            if (opening) return;
            opening = true;
            try {
                setStatus('Otváram zákazku...');

                const tiles = Array.from(document.querySelectorAll(WORKCENTER_TILE_SELECTOR));
                const tile = tiles.find((t) => getWorkcenterName(t) === item.workcenter);
                if (!tile) { setStatus('Pracovisko sa nenašlo.'); return; }

                const listId = getListIdForTile(tile);
                let list = listId ? document.getElementById(listId) : null;

                if (!list) {
                    pressElement(tile);
                    list = listId ? await waitFor(() => document.getElementById(listId), { timeout: TILE_WAIT_TIMEOUT }) : null;
                }
                if (!list) { setStatus('Zoznam zákaziek pre toto pracovisko sa nenašiel.'); return; }

                list.scrollIntoView({ block: 'center' });
                await sleep(SETTLE_DELAY);

                const targetText = formatProductionOrder(item);
                let li = Array.from(list.querySelectorAll('li')).find((el) => extractOrderText(el) === targetText);
                if (!li) {
                    await sleep(SETTLE_DELAY);
                    li = Array.from(list.querySelectorAll('li')).find((el) => extractOrderText(el) === targetText);
                }
                if (!li) { setStatus('Konkrétna zákazka sa nenašla, otvorené je aspoň pracovisko.'); return; }

                li.scrollIntoView({ block: 'center' });
                const fired = fireListItemPress(li);
                if (!fired) li.click();
                setStatus(fired ? 'Otvorené: ' + targetText : 'Chyba pri otváraní zákazky.');
            } catch (e) {
                console.warn(LOG, 'chyba pri otvarani polozky', e);
                setStatus('Chyba pri otváraní zákazky.');
            } finally {
                opening = false;
            }
        }

        function matchesTerm(it, term) {
            const dashIndex = term.indexOf('-');
            if (dashIndex !== -1) {
                const orderPart = term.slice(0, dashIndex).trim();
                const opPart = term.slice(dashIndex + 1).trim();
                const orderMatch = !orderPart || (it.productionOrderNo || '').toLowerCase().includes(orderPart);
                const opMatch = !opPart || (it.operationNo || '').toLowerCase().includes(opPart);
                return orderMatch && opMatch;
            }
            return (
                (it.productionOrderNo || '').toLowerCase().includes(term) ||
                (it.salesOrderNo || '').toLowerCase().includes(term) ||
                (it.materialNo || '').toLowerCase().includes(term) ||
                (it.material || '').toLowerCase().includes(term)
            );
        }

        function buildSearchUI(sidebar) {
            const list = document.createElement('div');
            list.id = UI_ID;
            list.className = 'sapMList sapMListBGSolid';
            list.style.width = '100%';

            const header = document.createElement('div');
            header.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBStandard sapMTB-Transparent-CTX sapMListHdr sapMListHdrTBar sapMTBHeader-CTX';
            const headerTitle = document.createElement('div');
            headerTitle.className = 'sapMTitle sapMTitleStyleAuto sapMTitleNoWrap sapUiSelectable sapMTitleMaxWidth sapMTitleTB sapMBarChild sapMTBShrinkItem';
            headerTitle.innerHTML = '<span dir="auto">Číslo zákazky</span>';
            header.appendChild(headerTitle);

            const tbContainer = document.createElement('div');
            tbContainer.className = 'sapMListInfoTBarContainer';
            const toolbar = document.createElement('div');
            toolbar.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBClear sapMTB-Transparent-CTX sapMListInfoTBar';
            const sf = document.createElement('div');
            sf.className = 'sapMSF sapMSFVal sapMBarChild sapMTBShrinkItem';
            sf.style.width = '100%';

            const form = document.createElement('form');
            form.className = 'sapMSFF';
            form.addEventListener('submit', (e) => e.preventDefault());

            const input = document.createElement('input');
            input.type = 'search';
            input.autocomplete = 'off';
            input.placeholder = 'Číslo zákazky / materiál';
            input.className = 'sapMSFI';

            const resetDiv = document.createElement('div');
            resetDiv.title = 'Resetovať';
            resetDiv.className = 'sapMSFR sapMSFB';
            resetDiv.addEventListener('click', () => { input.value = ''; render(''); });

            const searchDiv = document.createElement('div');
            searchDiv.title = 'Hľadať';
            searchDiv.className = 'sapMSFS sapMSFB';

            form.appendChild(input);
            form.appendChild(resetDiv);
            form.appendChild(searchDiv);
            sf.appendChild(form);
            toolbar.appendChild(sf);
            tbContainer.appendChild(toolbar);

            const status = document.createElement('div');
            status.id = '__pda_status__';
            status.style.fontSize = '0.72rem';
            status.style.color = '#888';
            status.style.padding = '0.2rem 0.5rem';

            const resultsUl = document.createElement('ul');
            resultsUl.className = 'sapMListItems sapMListUl sapMListHighlight sapMListShowSeparatorsAll sapMListModeSingleSelectMaster';
            resultsUl.setAttribute('role', 'listbox');
            resultsUl.tabIndex = 0;

            function renderItem(it) {
                const li = document.createElement('li');
                li.tabIndex = 0;
                li.setAttribute('role', 'option');
                li.className = 'sapMLIB sapMLIB-CTX sapMLIBShowSeparator sapMLIBTypeActive sapMLIBActionable sapMLIBHoverable sapMLIBFocusable sapMCLI sapUiTinyMargin';
                li.innerHTML =
                    '<div class="sapMLIBContent">' +
                    '<div class="sapMFlexBoxFit sapMFlexBox sapMHBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent" style="height:100%;width:100%;">' +
                    '<div class="sapMFlexBox sapMVBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent sapMFlexItem" style="height:100%;width:100%;">' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="font-weight:bold;text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + formatProductionOrder(it) + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.workcenter || '') + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.material || '') + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.descriptionShort || '') + '</bdi></span></span>' +
                    '</div></div></div>';

                li.addEventListener('click', () => openItem(it));
                li.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openItem(it); }
                });
                return li;
            }

            function idleStatus(index) {
                return 'Index: ' + index.length + ' zákaziek' +
                    (shared.ordersIndexUpdatedAt
                        ? ' (aktualiz. ' + shared.ordersIndexUpdatedAt.toLocaleTimeString() + ')'
                        : ' (čaká sa na načítanie aplikácie)');
            }

            function render(filterText) {
                resultsUl.innerHTML = '';
                const term = (filterText || '').trim().toLowerCase();
                const index = shared.ordersIndex || [];

                if (!term) {
                    setStatus(idleStatus(index));
                    lastAutoOpenedKey = null;
                    return;
                }

                const matches = index.filter((it) => matchesTerm(it, term));
                setStatus(matches.length + ' výsledok/-ov (z ' + index.length + ' položiek)');
                matches.slice(0, 200).forEach((it) => resultsUl.appendChild(renderItem(it)));

                if (matches.length === 1) {
                    const key = getItemKey(matches[0]);
                    if (key !== lastAutoOpenedKey) {
                        lastAutoOpenedKey = key;
                        openItem(matches[0]);
                        setTimeout(() => {
                            input.value = '';
                            resultsUl.innerHTML = '';
                            setStatus(idleStatus(index));
                            lastAutoOpenedKey = null;
                        }, 500);
                    }
                } else {
                    lastAutoOpenedKey = null;
                }
            }

            input.addEventListener('input', () => render(input.value));
            input.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const decoded = padOperationPart(decodeScannerInput(input.value));
                input.value = decoded;
                render(decoded);
            });

            list.appendChild(header);
            list.appendChild(tbContainer);
            list.appendChild(status);
            list.appendChild(resultsUl);
            sidebar.appendChild(list);

            render('');
            renderFn = render;

            // sprístupní vyplnenie vyhľadávania pre modul skenera
            shared.fillSearchInput = (value) => {
                input.value = value;
                render(value);
            };
        }

        function ensureLayout() {
            if (opening) return;
            if (!document.getElementById(HOME_BUTTON_ID) || !document.getElementById(CONTENT_ID)) return;

            const parentSection = document.getElementById(PARENT_SECTION_ID);
            const tasksPanel = document.getElementById(TASKS_PANEL_ID);
            const workcenterPanel = document.getElementById(PANEL_ID);
            if (!parentSection || !tasksPanel || !workcenterPanel) return;

            let container = document.getElementById(SIDEBAR_ID);
            if (!container) {
                container = document.createElement('div');
                container.id = SIDEBAR_ID;
                container.style.width = '100%';
                container.style.boxSizing = 'border-box';
                container.style.maxHeight = '80vh';
                container.style.overflowY = 'auto';
                container.style.margin = '10px 0';
                container.style.backgroundColor = '#ffffff';
                container.style.borderRadius = '10px';
            }

            if (container.nextElementSibling !== workcenterPanel || container.parentElement !== parentSection) {
                parentSection.insertBefore(container, workcenterPanel);
            }

            if (!document.getElementById(UI_ID)) buildSearchUI(container);
        }

        DomWatch.add(ensureLayout);
    }

    /* ------------------ 3.5 Bocny panel pouzivatelov -------------------- */

    function modUsersPanel() {
        const PARENT_ID = 'Main';
        const CONTENT_ID = 'Main--MainPage';
        const WRAPPER_ID = '__pda_main_wrapper__';
        const PANEL_ID = '__pda_user_switch_panel__';
        const STYLE_TAG_ID = '__pda_panel_styles__';
        const PANEL_WIDTH = '280px';

        function injectStyles() {
            if (document.getElementById(STYLE_TAG_ID)) return;
            const style = document.createElement('style');
            style.id = STYLE_TAG_ID;
            style.textContent =
                '.pda-user-btn { border: 2px solid transparent; }' +
                '.pda-user-btn:hover { border: 2px solid #e9e9f0; }';
            document.head.appendChild(style);
        }

        function buildUserButtons(panel) {
            if (settings.users.length === 0) {
                const hint = document.createElement('div');
                hint.textContent = 'Zatiaľ nie sú zadaní žiadni používatelia. Doplň ich v nastaveniach (ozubené koliesko vpravo dole).';
                hint.style.color = '#c9c9dd';
                hint.style.fontSize = '0.8rem';
                hint.style.lineHeight = '1.4';
                panel.appendChild(hint);
                return;
            }

            settings.users.forEach(({ username, password }) => {
                const btn = document.createElement('button');
                btn.textContent = username;
                btn.classList.add('pda-user-btn');
                btn.style.display = 'block';
                btn.style.width = '100%';
                btn.style.padding = '0.5rem 0.8rem';
                btn.style.marginBottom = '0.4rem';
                btn.style.cursor = 'pointer';
                btn.style.textAlign = 'left';
                btn.style.borderRadius = '8px';
                btn.style.color = '#e9e9f0';
                btn.style.backgroundColor = '#222252';
                btn.addEventListener('click', () => UserSwitch.to(username, password));
                panel.appendChild(btn);
            });
        }

        function ensureLayout() {
            const parent = document.getElementById(PARENT_ID);
            const content = document.getElementById(CONTENT_ID);
            if (!parent || !content) return;

            let wrapper = document.getElementById(WRAPPER_ID);
            let panel = document.getElementById(PANEL_ID);
            if (wrapper && wrapper.contains(content) && panel) return;

            if (!wrapper) {
                wrapper = document.createElement('div');
                wrapper.id = WRAPPER_ID;
                wrapper.style.display = 'flex';
                wrapper.style.flexDirection = 'row';
                wrapper.style.alignItems = 'stretch';
                wrapper.style.width = '100%';
                wrapper.style.height = '100%';
            }

            if (!panel) {
                panel = document.createElement('div');
                panel.id = PANEL_ID;
                panel.style.flex = '0 0 ' + PANEL_WIDTH;
                panel.style.maxWidth = PANEL_WIDTH;
                panel.style.marginRight = '4px';
                panel.style.background = '#313175';
                panel.style.overflowY = 'auto';
                panel.style.padding = '0.75rem';
                panel.style.boxSizing = 'border-box';

                const title = document.createElement('div');
                title.textContent = 'Rýchla zmena používateľa';
                title.style.fontWeight = 'bold';
                title.style.marginBottom = '0.5rem';
                title.style.color = '#e9e9f0';
                panel.appendChild(title);

                buildUserButtons(panel);
            }

            if (!wrapper.contains(content)) {
                parent.insertBefore(wrapper, content);
                wrapper.appendChild(panel);
                wrapper.appendChild(content);
                content.style.flex = '1 1 auto';
                content.style.minWidth = '0';
            }
        }

        onReady(injectStyles);
        DomWatch.add(ensureLayout);
    }

    /* ------------------- 3.6 Skener a RFID citacka ---------------------- */

    function modInputListener() {
        const HIDDEN_INPUT_ID = '__pda_hidden_scanner_input__';
        const CARD_ID_LENGTH = 10;
        const REFOCUS_INTERVAL = 1000;

        function handleRawInput(raw) {
            const decoded = decodeScannerInput(raw);

            if (decoded.length === CARD_ID_LENGTH) {
                const user = UserSwitch.findByCardId(decoded);
                if (user) {
                    console.log(LOG, 'karta patrí používateľovi:', user.username);
                    UserSwitch.to(user.username, user.password);
                } else {
                    console.warn(LOG, 'karta nerozpoznaná, ID:', decoded);
                }
                return;
            }

            const finalValue = padOperationPart(decoded);
            console.log(LOG, 'rozpoznané číslo zákazky:', finalValue);
            if (typeof shared.fillSearchInput === 'function') {
                shared.fillSearchInput(finalValue);
            } else {
                console.warn(LOG, 'modul vyhľadávania nebeží, zákazka nebola vložená');
            }
        }

        function createHiddenInput() {
            let input = document.getElementById(HIDDEN_INPUT_ID);
            if (input) return input;

            input = document.createElement('input');
            input.id = HIDDEN_INPUT_ID;
            input.type = 'text';
            input.autocomplete = 'off';
            input.style.position = 'fixed';
            input.style.top = '0';
            input.style.left = '0';
            input.style.width = '1px';
            input.style.height = '1px';
            input.style.opacity = '0';
            input.style.border = 'none';
            input.style.padding = '0';
            input.style.margin = '0';
            input.style.pointerEvents = 'none';

            input.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const raw = input.value;
                input.value = '';
                if (raw) handleRawInput(raw);
            });

            document.body.appendChild(input);
            return input;
        }

        function ensureFocus(input) {
            const active = document.activeElement;
            const isButton = !!active && (active.tagName === 'BUTTON' || active.getAttribute('role') === 'button');
            // nekradne focus, ked pouzivatel pise do policka alebo je otvoreny panel nastaveni
            if (document.getElementById('__pda_settings_overlay__')) return;
            if (!active || active === document.body || isButton) input.focus();
        }

        onReady(() => {
            const input = createHiddenInput();
            input.focus();
            setInterval(() => ensureFocus(input), REFOCUS_INTERVAL);
        });
    }

    /* ---------------------- 3.7 Tlacidlo vykresu ------------------------ */

    function modDrawingButton() {
        const DB_NAME = 'pda_drawing_db';
        const STORE_NAME = 'handles';
        const HANDLE_KEY = 'excel_file_handle';

        // stlpce sa daju prestavit v nastaveniach (predvolene H, AH, AI)
        const COL_ORDER_NO = colToIndex(settings.excel.colOrder, 7);
        const COL_DRAWING_NO = colToIndex(settings.excel.colDrawing, 33);
        const COL_VERSION = colToIndex(settings.excel.colVersion, 34);

        const CONTAINER_ID = 'WorkcenterDetail--Order_FlexBox';
        const WRAPPER_ID = '__pda_order_drawing_wrapper__';
        const BUTTON_ID = '__pda_order_drawing_button__';
        const LOAD_BUTTON_ID = '__pda_order_drawing_load_button__';

        let drawingIndex = {};
        let currentDrawingInfo = null;
        let pendingHandle = null;
        let currentLoadState = 'checking';

        const supportsFsAccess = typeof W.showOpenFilePicker === 'function';

        // --- IndexedDB (aby vyber suboru prezil obnovenie stranky) ---

        function openHandleDb() {
            return new Promise((resolve, reject) => {
                const req = W.indexedDB.open(DB_NAME, 1);
                req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            });
        }

        async function saveHandle(handle) {
            const db = await openHandleDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readwrite');
                tx.objectStore(STORE_NAME).put(handle, HANDLE_KEY);
                tx.oncomplete = () => resolve();
                tx.onerror = () => reject(tx.error);
                tx.onabort = () => reject(tx.error);
            });
        }

        async function loadHandle() {
            const db = await openHandleDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readonly');
                const req = tx.objectStore(STORE_NAME).get(HANDLE_KEY);
                req.onsuccess = () => resolve(req.result || null);
                req.onerror = () => reject(req.error);
            });
        }

        // --- Excel ---

        /*
         * Cislo vyrobnej zakazky sa v Exceli a v PDA nepise rovnako:
         *   PDA:   001600089585
         *   Excel: '1600089585   (apostrof = textova bunka, bez uvodnych nul)
         * Povodny skript to riesil tak, ze natvrdo odrezal prve dva znaky a
         * pridal apostrof - staci mala zmena formatu v Exceli a nenajde nic.
         * Preto sa obe strany prevedu na rovnaky tvar: bez apostrofu, bez
         * medzier a bez uvodnych nul.
         */
        // '001600108008' aj "'1600108008" -> '1600108008' (len cislice, bez uvodnych nul)
        function normalizeOrderKey(value) {
            return String(value === undefined || value === null ? '' : value)
                .replace(/\D/g, '')
                .replace(/^0+/, '');
        }

        // Stlpce sa hladaju podla NAZVU v hlavicke (AutomatedOQ180.xlsx ich ma
        // stale rovnake), pismena z nastaveni su len zaloha pre iny subor.
        const HEADER_NAMES = {
            order: 'production order number',
            drawing: 'document number (material of production order)',
            version: 'document version (material of production order)',
            material: 'material number (production order)',
        };

        // 'SIEHE DIS', 'B.V.', 'O.Z.' nie su cisla vykresov - cislo vykresu ma vzdy cislicu
        function isDrawingNo(v) {
            return /\d/.test(String(v || ''));
        }

        function cellText(row, col) {
            if (col < 0 || row[col] === undefined || row[col] === null) return '';
            return String(row[col]).trim().replace(/^'+/, '');
        }

        function findColumns(headerRow) {
            const hdr = (headerRow || []).map((h) => String(h === undefined || h === null ? '' : h).trim().toLowerCase());
            const cols = {
                order: hdr.indexOf(HEADER_NAMES.order),
                drawing: hdr.indexOf(HEADER_NAMES.drawing),
                version: hdr.indexOf(HEADER_NAMES.version),
                material: hdr.indexOf(HEADER_NAMES.material),
            };
            if (cols.order === -1 || cols.drawing === -1) {
                console.log(LOG, 'Excel: hlavičky stĺpcov sa nenašli, používam písmená z nastavení');
                return { order: COL_ORDER_NO, drawing: COL_DRAWING_NO, version: COL_VERSION, material: -1 };
            }
            return cols;
        }

        function buildDrawingIndex(rows) {
            const cols = findColumns(rows[0]);
            const index = {};
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;
                const key = normalizeOrderKey(row[cols.order]);
                if (!key || index[key]) continue;   // prvy vyskyt vyhrava (v Python verzii overene: 0 konfliktov)
                const vyk = cellText(row, cols.drawing);
                const valid = isDrawingNo(vyk);
                index[key] = {
                    drawingNo: valid ? vyk : '',
                    version: valid ? cellText(row, cols.version) : '',
                    materialNo: cellText(row, cols.material).replace(/^0+/, ''),
                };
            }
            return index;
        }

        function parseWorkbook(bytes) {
            const workbook = XLSX.read(bytes, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            if (!sheet) {
                console.warn(LOG, 'v exceli sa nenašiel žiadny list');
                return false;
            }
            drawingIndex = buildDrawingIndex(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
            console.log(LOG, 'index výkresov vytvorený, záznamov:', Object.keys(drawingIndex).length);
            updateLoadButtonState('loaded');
            if (shared.currentOperation) applyForOperation(shared.currentOperation);
            return true;
        }

        async function loadExcelFromFile(file) {
            parseWorkbook(new Uint8Array(await file.arrayBuffer()));
        }

        // Excel stiahnuty z adresy (http://...). GM_xmlhttpRequest obchadza CORS,
        // takze sa da siahnut aj na interny server.
        function fetchExcel(url) {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url,
                    responseType: 'arraybuffer',
                    onload: (r) => {
                        if (r.status < 200 || r.status >= 300) {
                            reject(new Error('HTTP ' + r.status));
                            return;
                        }
                        resolve(r.response);
                    },
                    onerror: () => reject(new Error('spojenie so serverom zlyhalo')),
                    ontimeout: () => reject(new Error('server neodpovedal včas')),
                });
            });
        }

        async function loadExcelFromUrl(url) {
            updateLoadButtonState('loading');
            try {
                const buffer = await fetchExcel(url);
                if (!buffer) throw new Error('prázdna odpoveď');
                parseWorkbook(new Uint8Array(buffer));
            } catch (e) {
                console.warn(LOG, 'Excel sa nepodarilo stiahnuť z adresy', url, e);
                updateLoadButtonState('url-error');
            }
        }

        async function pickFileAndRemember() {
            try {
                const [handle] = await W.showOpenFilePicker({
                    types: [{
                        description: 'Excel',
                        accept: {
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                            'application/vnd.ms-excel': ['.xls'],
                        },
                    }],
                    multiple: false,
                });
                await saveHandle(handle);
                updateLoadButtonState('loading');
                await loadExcelFromFile(await handle.getFile());
            } catch (err) {
                if (err && err.name === 'AbortError') return;
                console.warn(LOG, 'chyba pri výbere súboru', err);
                updateLoadButtonState('error');
            }
        }

        async function tryAutoLoad() {
            if (!supportsFsAccess) { updateLoadButtonState('unsupported'); return; }

            let handle = null;
            try { handle = await loadHandle(); } catch (e) { /* ignore */ }
            if (!handle) { updateLoadButtonState('nofile'); return; }

            let permission;
            try {
                permission = await handle.queryPermission({ mode: 'read' });
            } catch (e) {
                updateLoadButtonState('needs-permission', handle);
                return;
            }

            if (permission === 'granted') {
                updateLoadButtonState('loading');
                try { await loadExcelFromFile(await handle.getFile()); }
                catch (e) { updateLoadButtonState('error'); }
                return;
            }
            updateLoadButtonState('needs-permission', handle);
        }

        async function confirmPermissionAndLoad(handle) {
            try {
                const permission = await handle.requestPermission({ mode: 'read' });
                if (permission !== 'granted') { updateLoadButtonState('needs-permission', handle); return; }
                updateLoadButtonState('loading');
                await loadExcelFromFile(await handle.getFile());
            } catch (e) {
                updateLoadButtonState('error');
            }
        }

        // --- napojenie na aktualne otvorenu operaciu ---

        /*
         * Prave otvorena operacia sa cita priamo z premennej appky
         * getGlobals().getVar('oSelectedWorkcenterOperation') - rovnako ako v
         * Python verzii. Povodne odpocuvanie siete cakalo na result.operation,
         * ktore appka neposiela spolahlivo, preto tlacidlo ostavalo prazdne.
         */
        function readSelectedOperation() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                const op = main && main.getController().getGlobals().getVar('oSelectedWorkcenterOperation');
                if (!op || !op.productionOrderNo) return null;
                return {
                    workcenter: op.workcenterDescription,
                    workcenterCode: op.workcenter,
                    productionOrderNo: String(op.productionOrderNo || ''),
                    salesOrderNo: String(op.salesOrderNo || ''),
                    operationNo: String(op.operationNo || ''),
                    sequenceNo: op.sequenceNo,
                    materialNo: String(op.materialNo || '').replace(/^0+/, ''),
                    material: op.material,
                };
            } catch (e) {
                return null;
            }
        }

        let lastOperationKey = null;

        function pollSelectedOperation() {
            const cur = readSelectedOperation();
            const key = cur ? cur.productionOrderNo + '|' + cur.operationNo + '|' + cur.sequenceNo : '';
            if (key === lastOperationKey) return;
            lastOperationKey = key;
            shared.currentOperation = cur;
            if (cur) applyForOperation(cur);
            else updateDisplay(null);
        }

        onReady(() => setInterval(pollSelectedOperation, 500));

        function stripLeadingApostrophe(value) {
            if (!value) return value;
            return value.charAt(0) === "'" ? value.slice(1) : value;
        }

        function salesOrderOf(current) {
            return String((current && current.salesOrderNo) || '').trim();
        }

        // "3-55.2-06.74-075_001_A_0.pdf" -> "3-55.2-06.74-075"
        function drawingNoFromResult(v) {
            if (v.cislo_vykresu) return v.cislo_vykresu;
            if (v.kluc && v.kluc_typ === 'vykres') return v.kluc;
            return String(v.nazov || '').replace(/\.(pdf|tiff?)$/i, '').split('_')[0];
        }

        // Ked sluzba vrati viac suborov, najrelevantnejsi je ten, ktory lezi
        // v priecinku tejto zakazky a ktoremu sedi revizia so SAP verziou.
        function bestResult(list) {
            return list.find((v) => v.v_zakazke && v.zhoda_revizie === true) ||
                   list.find((v) => v.v_zakazke) ||
                   list.find((v) => v.zhoda_revizie === true) ||
                   list[0] || null;
        }

        let lookupToken = 0;

        /*
         * Cislo vykresu sa zistuje v dvoch krokoch:
         *   1. z Excelu podla cisla vyrobnej zakazky (ak je Excel nacitany)
         *   2. inak priamo zo sluzby vykresov podla cisla materialu, ktore
         *      PDA uz pozna - Excel teda nie je podmienkou
         * Zakaznicka zakazka ide do parametra "path", vdaka comu sluzba oznaci
         * vykresy lezice priamo v tejto zakazke (v_zakazke = true).
         */
        async function applyForOperation(current) {
            const token = ++lookupToken;

            const fromExcel = drawingIndex[normalizeOrderKey(current.productionOrderNo)];
            const excelDrawingNo = fromExcel ? stripLeadingApostrophe(fromExcel.drawingNo) : '';
            if (excelDrawingNo) {
                updateDisplay({
                    drawingNo: excelDrawingNo,
                    version: stripLeadingApostrophe(fromExcel.version) || '',
                    searchTerm: excelDrawingNo,
                    source: 'excel',
                });
                return;
            }

            updateDisplay({ drawingNo: '…', version: '', searchTerm: '', source: 'hladam' });

            // 2) sluzba vykresov pozna denny export -> kluc priamo pre tuto vyrobnu zakazku
            try {
                const z = await pdmOrderLookup(current.productionOrderNo);
                if (token !== lookupToken) return;
                if (z && z.kluc) {
                    const isDrawing = z.typ !== 'material';
                    updateDisplay({
                        drawingNo: isDrawing ? z.kluc : 'SAP ' + z.kluc,
                        version: isDrawing ? (z.revizia || '') : '',
                        searchTerm: z.kluc,
                        source: 'server',
                        count: z.pocet || 0,
                    });
                    return;
                }
            } catch (e) {
                if (token !== lookupToken) return;
                console.log(LOG, 'služba nepozná zákazku (alebo ešte nemá /zakazka), skúšam materiál:', e.message);
            }

            // 3) zaloha: material - najprv z Excelu (ak zakazku pozna, ale vykres tam nema), inak z operacie
            const material = ((fromExcel && fromExcel.materialNo) || String(current.materialNo || '')).trim();
            if (!material) { updateDisplay(null); return; }

            updateDisplay({ drawingNo: '…', version: '', searchTerm: material, source: 'hladam' });

            try {
                const d = await pdmSearch(material, { path: salesOrderOf(current) });
                if (token !== lookupToken) return; // medzitym sa otvorila ina operacia

                const best = bestResult(d.vysledky || []);
                if (!best) { updateDisplay(null); return; }

                updateDisplay({
                    drawingNo: drawingNoFromResult(best),
                    version: best.revizia || '',
                    searchTerm: material,
                    source: 'pdm',
                    count: d.pocet || 0,
                });
            } catch (e) {
                if (token !== lookupToken) return;
                console.warn(LOG, 'hľadanie výkresu zlyhalo', e);
                updateDisplay(null);
            }
        }

        function updateDisplay(info) {
            const valueEl = document.getElementById('__pda_order_drawing_value__');
            const revisionEl = document.getElementById('__pda_order_drawing_revision__');
            if (!valueEl || !revisionEl) return;

            if (!info) {
                currentDrawingInfo = null;
                valueEl.textContent = '—';
                revisionEl.textContent = '';
                return;
            }

            currentDrawingInfo = info;
            valueEl.textContent = info.drawingNo || '—';

            if (info.source === 'hladam') {
                revisionEl.textContent = 'hľadám…';
            } else if (info.version) {
                revisionEl.textContent = 'rev. ' + info.version + (info.count > 1 ? ' · ' + info.count + ' súb.' : '');
            } else {
                revisionEl.textContent = info.count > 1 ? info.count + ' súbory' : '';
            }
        }

        // --- sluzba "Mapa vykresov" (PDM) ---

        /*
         * Vyrobna zakazka -> kluc z denneho exportu. Sluzba vykresov cita AutomatedOQ180.xlsx
         * raz denne a ku kazdemu vykresu drzi zoznam zakaziek; endpoint /zakazka/<cislo>
         * to vrati jednym dopytom. 404 = zakazka v exporte nie je.
         */
        function pdmOrderLookup(productionOrderNo) {
            const cislo = normalizeOrderKey(productionOrderNo);
            if (!cislo) return Promise.resolve(null);
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + '/zakazka/' + encodeURIComponent(cislo),
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    timeout: 15000,
                    onload: (r) => {
                        if (r.status === 404) { resolve(null); return; }
                        if (r.status < 200 || r.status >= 300) { reject(new Error('HTTP ' + r.status)); return; }
                        try { resolve(JSON.parse(r.responseText)); }
                        catch (e) { reject(new Error('neplatná odpoveď služby')); }
                    },
                    onerror: () => reject(new Error('spojenie so službou zlyhalo')),
                    ontimeout: () => reject(new Error('služba neodpovedala včas')),
                });
            });
        }

        function pdmSearch(cislo, { path = '', live = false } = {}) {
            return new Promise((resolve, reject) => {
                const params = new URLSearchParams({ q: cislo });
                if (path) params.set('path', path);
                if (live) params.set('live', '1');
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + '/search?' + params,
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    onload: (response) => {
                        if (response.status < 200 || response.status >= 300) {
                            reject(new Error('Služba výkresov vrátila HTTP ' + response.status));
                            return;
                        }
                        try { resolve(JSON.parse(response.responseText)); }
                        catch (e) { reject(new Error('Neplatná odpoveď zo služby výkresov')); }
                    },
                    onerror: () => reject(new Error('Služba výkresov neodpovedala (chyba spojenia)')),
                    ontimeout: () => reject(new Error('Služba výkresov neodpovedala včas')),
                });
            });
        }

        /*
         * Zhoda revizie z Excelu s reviziou v nazve suboru - rovnake pravidlo ako
         * v sluzbe (Python verzia): pismeno musi sediet, index len ak ho poznaju
         * obaja. 'BC' ~ 'C' ano, 'BC' ~ 'B' nie, '0A' ~ 'A' ano.
         */
        function revisionMatches(label, rev) {
            label = String(label || '').toUpperCase();
            rev = String(rev || '').toUpperCase();
            if (!label || !rev || label === 'BEZ REV.') return false;
            const l1 = label.slice(-1), i1 = label.length > 1 ? label.slice(0, -1) : '';
            const l2 = rev.slice(-1), i2 = rev.length > 1 ? rev.slice(0, -1) : '';
            return l1 === l2 && (!i1 || !i2 || i1 === i2);
        }

        // Hlada sa LEN podla cisla vykresu (bez revizie); revizia z Excelu sa
        // v zozname iba zvyrazni, rozhodnutie ostava na cloveku.
        function pdmOpenDialog(cislo, { path = '', rev = '' } = {}) {
            const overlay = document.createElement('div');
            overlay.id = '__pda_pdm_overlay__';
            overlay.style.cssText =
                'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;' +
                'display:flex;align-items:center;justify-content:center;font-family:inherit;';

            overlay.innerHTML =
                '<div style="background:#fff;border-radius:10px;max-width:900px;width:92%;max-height:80vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,.3)">' +
                '<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#5b9bd5;color:#fff">' +
                '<b>Výkresy — ' + cislo + '</b>' +
                '<button data-pda-zavrit type="button" style="background:none;border:0;color:#fff;font-size:22px;line-height:1;cursor:pointer">×</button>' +
                '</div><div data-pda-telo style="padding:14px 16px;overflow:auto">Hľadám…</div></div>';

            const close = () => { overlay.remove(); document.removeEventListener('keydown', onEsc); };
            const onEsc = (e) => { if (e.key === 'Escape') close(); };

            overlay.querySelector('[data-pda-zavrit]').addEventListener('click', close);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
            document.addEventListener('keydown', onEsc);
            document.body.appendChild(overlay);

            const telo = overlay.querySelector('[data-pda-telo]');

            const TH = 'style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5"';
            const TD = 'style="padding:6px;border-bottom:1px solid #eee"';

            function run(live) {
                telo.innerHTML = live
                    ? '<p>Hľadám naživo v PDM… (môže to trvať pár sekúnd)</p>'
                    : '<p>Hľadám…</p>';

                pdmSearch(cislo, { path, live }).then((d) => {
                    if (!d.pocet) {
                        telo.innerHTML =
                            '<p>Pre <b>' + cislo + '</b> sa nenašiel žiadny PDF ani TIFF výkres.</p>' +
                            (live ? '' :
                                '<p style="color:#777;font-size:12px">Denná mapa sa obnovuje raz za deň — čerstvo pridaný výkres v nej ešte nemusí byť.</p>' +
                                '<button data-pda-live type="button" style="background:#5b9bd5;color:#fff;border:0;border-radius:7px;padding:8px 15px;cursor:pointer;font:inherit;font-size:.88rem">Hľadať naživo v PDM</button>');
                        const liveBtn = telo.querySelector('[data-pda-live]');
                        if (liveBtn) liveBtn.addEventListener('click', () => run(true));
                        return;
                    }

                    // zhoda revizie: primarne zo sluzby (pocita ju zo SAP verzie), inak lokalne z Excelu
                    const rows = (d.vysledky || []).map((v) => {
                        const known = v.zhoda_revizie === true || v.zhoda_revizie === false;
                        return Object.assign({}, v, { zhoda: known ? v.zhoda_revizie : (rev ? revisionMatches(v.revizia, rev) : null) });
                    });
                    rows.sort((a, b) => ((b.zhoda === true) - (a.zhoda === true)) || String(a.nazov).localeCompare(String(b.nazov)));

                    telo.innerHTML =
                        '<table style="width:100%;border-collapse:collapse;font-size:14px"><thead><tr>' +
                        '<th ' + TH + '>Názov</th><th ' + TH + '>Revízia</th>' +
                        '<th ' + TH + '>Stav</th><th ' + TH + '>Ver.</th>' +
                        '</tr></thead><tbody>' +
                        rows.map((v) => {
                            const bg = v.zhoda === true ? '#e6f4ea' : (v.v_zakazke ? '#f3f7fd' : 'transparent');
                            const mark = v.v_zakazke ? ' <span style="font-size:10px;color:#2f6fbf;border:1px solid #b9cbe8;border-radius:9px;padding:1px 6px">v zákazke</span>' : '';
                            return '<tr data-id="' + v.id + '" style="cursor:pointer;background:' + bg + '">' +
                                '<td ' + TD + '><b>' + v.nazov + '</b>' + mark + '</td>' +
                                '<td ' + TD + '>' + (v.revizia || '') + '</td>' +
                                '<td ' + TD + '>' + (v.stav || '') + '</td>' +
                                '<td ' + TD + '>' + (v.verzia != null ? v.verzia : '') + '</td></tr>';
                        }).join('') +
                        '</tbody></table>' +
                        '<p data-pda-info style="color:#777;font-size:12px;margin-top:10px">' + d.pocet + ' výkresov · ' +
                        (d.zdroj === 'mapa' ? 'z dennej mapy' : 'naživo z PDM') +
                        (d.mapa_z ? ' (' + d.mapa_z + ')' : '') +
                        ' · zelený riadok = revízia sedí so SAP · kliknutím sa výkres otvorí</p>';

                    telo.querySelectorAll('tr[data-id]').forEach((tr) => {
                        tr.addEventListener('click', () => {
                            const info = telo.querySelector('[data-pda-info]');
                            // vydaj suboru trva 0,4-4 s, treba to dat najavo
                            if (info) info.textContent = 'Otváram výkres… (ťahá sa z PDM, môže to chvíľu trvať)';
                            W.open(settings.pdm.base + '/file/' + tr.dataset.id, '_blank');
                        });
                    });
                }).catch((err) => {
                    telo.innerHTML = '<p style="color:#b00">Služba výkresov neodpovedala.<br>' + err.message +
                        '<br><span style="color:#777;font-size:12px">Adresa: ' + settings.pdm.base + '</span></p>';
                });
            }

            run(false);
        }

        // --- tlacidla ---

        function renderLoadButtonState(loadButton, state) {
            if (state === 'loaded') { loadButton.style.display = 'none'; return; }
            loadButton.style.display = '';

            const states = {
                loading: ['Načítavam…', '#f5f5f5', '#ccc'],
                'needs-permission': ['Povoliť prístup k Excelu', '#fff4e5', '#f9a825'],
                error: ['Chyba, skús znova', '#fdecea', '#e53935'],
                'url-error': ['Excel sa nestiahol — skús znova', '#fdecea', '#e53935'],
                'url-invalid': ['Adresa Excelu nie je http — klik otvorí nastavenia', '#fff4e5', '#f9a825'],
                unsupported: ['Prehliadač nepodporuje zapamätanie', '#fdecea', '#e53935'],
                // Excel uz nie je podmienkou - vykres sa najde aj podla materialu
                nofile: ['Excel (nepovinné)', '#f5f5f5', '#ccc'],
            };
            const [text, bg, border] = states[state] || states.nofile;
            loadButton.textContent = text;
            loadButton.style.backgroundColor = bg;
            loadButton.style.borderColor = border;
        }

        function updateLoadButtonState(state, handle) {
            currentLoadState = state;
            pendingHandle = state === 'needs-permission' ? handle : null;
            const loadButton = document.getElementById(LOAD_BUTTON_ID);
            if (loadButton) renderLoadButtonState(loadButton, state);
        }

        function buildWrapper() {
            const wrapper = document.createElement('div');
            wrapper.id = WRAPPER_ID;
            wrapper.style.cssText = 'display:flex;justify-content:flex-end;align-items:center;width:100%;box-sizing:border-box;margin-bottom:8px;';

            const loadButton = document.createElement('button');
            loadButton.id = LOAD_BUTTON_ID;
            loadButton.type = 'button';
            loadButton.style.cssText = 'padding:6px 10px;border:1px solid #ccc;border-radius:6px;background:#f5f5f5;cursor:pointer;font-size:0.75rem;margin-right:8px;align-self:center;';
            loadButton.addEventListener('click', () => {
                const url = String(settings.excel.url || '').trim();
                if (url && !/^https?:\/\//i.test(url)) openSettings();
                else if (url) loadExcelFromUrl(url);
                else if (pendingHandle) confirmPermissionAndLoad(pendingHandle);
                else pickFileAndRemember();
            });

            const button = document.createElement('button');
            button.id = BUTTON_ID;
            button.type = 'button';
            button.style.cssText = 'display:flex;flex-direction:column;align-items:flex-end;gap:2px;padding:6px 12px;border:1px solid #5b9bd5;border-radius:6px;background:#fff;cursor:pointer;width:150px;flex-shrink:0;margin-right:8px;';

            const label = document.createElement('span');
            label.textContent = 'VÝKRES';
            label.style.cssText = 'font-size:0.65rem;color:#888;letter-spacing:0.05em;';

            const value = document.createElement('span');
            value.id = '__pda_order_drawing_value__';
            value.textContent = '—';
            value.style.cssText = 'font-size:0.9rem;font-weight:bold;color:#000;';

            const revision = document.createElement('span');
            revision.id = '__pda_order_drawing_revision__';
            revision.style.cssText = 'font-size:0.75rem;color:#555;';

            button.appendChild(label);
            button.appendChild(value);
            button.appendChild(revision);
            button.addEventListener('click', () => {
                const path = salesOrderOf(shared.currentOperation);

                // ak uz vieme cislo vykresu (alebo aspon material), hladame podla neho
                if (currentDrawingInfo && currentDrawingInfo.searchTerm) {
                    pdmOpenDialog(currentDrawingInfo.searchTerm, {
                        path,
                        rev: (currentDrawingInfo.source === 'excel' || currentDrawingInfo.source === 'server')
                            ? currentDrawingInfo.version : '',
                    });
                    return;
                }
                // inak skusime aspon cislo materialu z otvorenej operacie
                const material = shared.currentOperation && String(shared.currentOperation.materialNo || '').trim();
                if (material) pdmOpenDialog(material, { path });
                else console.log(LOG, 'pre aktuálnu zákazku nie je známe ani číslo výkresu, ani materiálu');
            });

            wrapper.appendChild(loadButton);
            wrapper.appendChild(button);
            return wrapper;
        }

        function ensureButton() {
            const container = document.getElementById(CONTAINER_ID);
            if (!container || document.getElementById(WRAPPER_ID)) return;

            container.insertBefore(buildWrapper(), container.firstChild);
            const loadButton = document.getElementById(LOAD_BUTTON_ID);
            if (loadButton) renderLoadButtonState(loadButton, currentLoadState);
            if (shared.currentOperation) applyForOperation(shared.currentOperation);
        }

        DomWatch.add(ensureButton);
        onReady(() => {
            // ak je v nastaveniach adresa, tahame odtial automaticky;
            // inak sa subor vybera rucne a prehliadac si ho pamata
            const url = String(settings.excel.url || '').trim();
            if (url && !/^https?:\/\//i.test(url)) {
                // cesta na disk (C:\... alebo \\server\...) nie je adresa - prehliadac ju neotvori
                console.warn(LOG, 'adresa Excelu nie je http(s), ignorujem:', url);
                updateLoadButtonState('url-invalid');
            } else if (url) {
                loadExcelFromUrl(url);
            } else {
                tryAutoLoad();
            }
        });
    }

    /* ------------------ 3.8 Prehlad pracovisk (uvodna) ------------------ */

    /*
     * Portovane z Python appky (Temporary/PDA/app/pda_action.py, blok OBLASTI).
     * Rozdiel oproti prvej verzii tohto modulu:
     *  - data sa NEcitaju zo siete ani z DOM, ale priamo z premennej appky
     *    getGlobals().getVar('aWorkcenterOperationLists')
     *  - kategoria = pole `area` (Assembly / Welding / Machining / Quality Control)
     *  - "vyraba" = pole `count` > 0 (kolko ludi je prihlasenych na pracovisku)
     *  - klik na stroj vola ten isty handler, aky vola appka pri kliku na kartu
     *  - povodne karty sa skryju jednym CSS pravidlom, v DOM ostavaju
     */
    function modWorkcenterOverview() {
        const OVERVIEW_ID = '__pda_overview__';
        const STYLE_ID = '__pda_overview_styles__';
        const BODY_CLASS = 'pda-overview-on';
        const HOME_BUTTON_ID = 'Main--Button_HomeScreen';
        const PANEL_ID = 'Main--Workcenter_Panel';
        const ORDER = ['Assembly', 'Welding', 'Machining', 'Quality Control'];
        const SUBTITLES = {
            'Assembly': 'Finálna montáž a podzostavy',
            'Welding': 'Zváranie a príprava',
            'Machining': 'CNC a konvenčné obrábanie',
            'Quality Control': 'Kontrola a meranie',
            'Ostatné': 'Nezaradené pracoviská',
        };
        const ICONS = { 'Assembly': '🔧', 'Welding': '🔥', 'Machining': '⚙', 'Quality Control': '🔎' };

        const expanded = new Set();   // rozbalene kategorie
        const filters = {};           // { q, only } pre kazdu kategoriu
        let lastSignature = null;

        function globals() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                return main ? main.getController().getGlobals() : null;
            } catch (e) { return null; }
        }

        function workcenterList() {
            const g = globals();
            if (!g) return [];
            try { return g.getVar('aWorkcenterOperationLists') || []; } catch (e) { return []; }
        }

        function userName() {
            const g = globals();
            try {
                const n = g && g.getVar('oUser', 'username');
                if (n) return String(n);
            } catch (e) { /* ignore */ }
            const el = document.querySelector('[id$="Label_Username-bdi"]');
            return el ? el.textContent.trim() : '';
        }

        function normalize(s) {
            return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
        }

        /*
         * Otvorenie pracoviska presne tak, ako to robi appka po kliku na kartu
         * (handler f_cards_Header3_press). V Python verzii sa zistilo, ze handler
         * po navTo nedobehne - texty casov ostanu prazdne a potvrdzovacie tlacidla
         * zapnute - preto sa to po 1,5 s dorobi rovnako, ako to robi appka.
         */
        function openWorkcenter(wc) {
            try {
                const core = W.sap.ui.getCore();
                const main = core.byId('Main').getController();
                const g = main.getGlobals();
                const app = core.byId('Application');

                if (app && app.getCurrentPage && String(app.getCurrentPage().getId()).indexOf('WorkcenterDetail') === 0) {
                    main.navTo('Main', 'show');
                }
                if (!g.getVar('oSelectedWorkcenterOperation')) g.setVar('oSelectedWorkcenterOperation', {});

                const ctx = { getObject: () => wc };
                const src = { getParent: () => ({ getBindingContext: () => ctx }) };
                main.f_cards_Header3_press({ getSource: () => src });

                setTimeout(() => {
                    try {
                        const op = g.getVar('oSelectedWorkcenterOperation') || {};
                        [['SetupTime_Text', 'setupTimeConf', 'setupTime'],
                         ['MachineTime_Text', 'machineTimeConf', 'machineTime'],
                         ['LaborTime_Text', 'laborTimeConf', 'laborTime']].forEach(([id, conf, plan]) => {
                            const e = core.byId('WorkcenterDetail--' + id);
                            if (e && typeof e.setText === 'function') {
                                e.setText(op.id ? op[conf] + ' / ' + op[plan] + ' min' : '- / - min');
                            }
                        });
                        if (!op.id) {
                            main.setItemValue('WorkcenterDetail', 'Part_Confirm_Button', 'enabled', false);
                            main.setItemValue('WorkcenterDetail', 'Confirm_Button', 'enabled', false);
                        }
                    } catch (e) { /* kozmetika, nie je kriticka */ }
                }, 1500);
            } catch (e) {
                console.warn(LOG, 'otvorenie pracoviska cez handler zlyhalo, skúšam klik na kartu', e);
                const tile = findTile(wc);
                if (tile) pressElement(tile);
            }
        }

        // zaloha: povodna karta v DOM podla nazvu pracoviska
        function findTile(wc) {
            const want = String(wc.workcenterDescription || '').trim();
            const tiles = document.querySelectorAll('[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]');
            for (const t of tiles) {
                const suffix = t.id.replace('Main--Workcenter_Toolbar-', '');
                const el = document.getElementById('Main--WorkcenterDescription_Label-' + suffix + '-bdi');
                if (el && el.textContent.trim() === want) return t;
            }
            return null;
        }

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
body.${BODY_CLASS} #${PANEL_ID} .sapMPanelContent > :not(#${OVERVIEW_ID}) { display: none !important; }
#${OVERVIEW_ID} { font: 14px/1.5 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; padding: 4px 0 10px; }
#${OVERVIEW_ID} .ov-hi { padding: 4px 6px 14px; }
#${OVERVIEW_ID} .ov-hi .k { font-size:.68rem; letter-spacing:.12em; color:#5b7cb5; font-weight:700; }
#${OVERVIEW_ID} .ov-hi .n { font-size:1.7rem; font-weight:700; margin:2px 0 1px; }
#${OVERVIEW_ID} .ov-hi .s { color:#6b7180; font-size:.87rem; }
#${OVERVIEW_ID} .ov-g { border:1px solid #dfe4ec; border-radius:12px; background:#fff; margin-bottom:10px; overflow:hidden; }
#${OVERVIEW_ID} .ov-g.open { border-color:#b9cbe8; box-shadow:0 1px 6px rgba(40,70,130,.08); }
#${OVERVIEW_ID} .ov-gh { display:flex; align-items:center; gap:14px; width:100%; background:none; border:0;
  padding:14px 16px; cursor:pointer; text-align:left; font:inherit; color:inherit; }
#${OVERVIEW_ID} .ov-g.open .ov-gh { background:#f2f6fc; }
#${OVERVIEW_ID} .ov-ic { width:42px; height:42px; border-radius:9px; background:#eaf1fb; flex-shrink:0;
  display:flex; align-items:center; justify-content:center; font-size:21px; }
#${OVERVIEW_ID} .ov-gt { flex:1 1 auto; min-width:0; }
#${OVERVIEW_ID} .ov-gt b { font-size:1.05rem; display:block; }
#${OVERVIEW_ID} .ov-gt span { color:#6b7180; font-size:.85rem; }
#${OVERVIEW_ID} .ov-num { text-align:center; flex-shrink:0; min-width:74px; }
#${OVERVIEW_ID} .ov-num b { display:block; font-size:1.3rem; color:#1d4ed8; line-height:1.1; }
#${OVERVIEW_ID} .ov-num span { font-size:.62rem; letter-spacing:.09em; color:#8b93a3; }
#${OVERVIEW_ID} .ov-num.on b { color:#2f7d43; }
#${OVERVIEW_ID} .ov-ch { color:#9aa3b4; font-size:18px; flex-shrink:0; }
#${OVERVIEW_ID} .ov-gb { padding:4px 16px 16px; border-top:1px solid #e8edf4; }
#${OVERVIEW_ID} .ov-tools { display:flex; gap:8px; justify-content:flex-end; margin:12px 0; flex-wrap:wrap; }
#${OVERVIEW_ID} .ov-tools input, #${OVERVIEW_ID} .ov-tools select {
  padding:6px 10px; border:1px solid #ccd4e0; border-radius:7px; font:inherit; font-size:.85rem; }
#${OVERVIEW_ID} .ov-tools input { width:190px; }
#${OVERVIEW_ID} .ov-chips { display:grid; grid-template-columns:repeat(auto-fill,minmax(215px,1fr)); gap:7px; }
#${OVERVIEW_ID} .ov-chip { display:flex; align-items:center; justify-content:space-between; gap:10px;
  border:1px solid #dfe4ec; border-radius:9px; background:#fff; padding:8px 11px; cursor:pointer;
  font:inherit; text-align:left; transition:border-color .12s, background .12s; }
#${OVERVIEW_ID} .ov-chip:hover { border-color:#5b8def; background:#f6f9ff; }
#${OVERVIEW_ID} .ov-chip b { font-size:.93rem; }
#${OVERVIEW_ID} .ov-txt { display:flex; flex-direction:column; min-width:0; }
#${OVERVIEW_ID} .ov-txt small { font-size:.71rem; color:#6b7180; line-height:1.3;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
#${OVERVIEW_ID} .ov-st { display:flex; align-items:center; gap:5px; font-size:.76rem; color:#1e40af;
  background:#e8f0fe; border-radius:999px; padding:3px 9px; white-space:nowrap; font-weight:600; }
#${OVERVIEW_ID} .ov-dot { width:7px; height:7px; border-radius:50%; background:#2563eb; flex-shrink:0; }
#${OVERVIEW_ID} .ov-chip.run .ov-st { color:#166534; background:#eaf6ee; }
#${OVERVIEW_ID} .ov-chip.run .ov-dot { background:#16a34a; }
#${OVERVIEW_ID} .ov-empty { color:#8b93a3; font-size:.85rem; padding:8px 2px; }
`;
            document.head.appendChild(st);
        }

        function chip(w) {
            const producing = (w.count || 0) > 0;
            const code = String(w.workcenter || '');
            const name = String(w.workcenterDescription || '');

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'ov-chip' + (producing ? ' run' : '');
            btn.title = code + ' — ' + name;

            const txt = document.createElement('span');
            txt.className = 'ov-txt';
            const label = document.createElement('b');
            label.textContent = code || name;
            txt.appendChild(label);
            if (code && name) {
                const sub = document.createElement('small');
                sub.textContent = name.length > 20 ? name.slice(0, 20).trim() + '…' : name;
                txt.appendChild(sub);
            }

            const st = document.createElement('span');
            st.className = 'ov-st';
            const dot = document.createElement('span');
            dot.className = 'ov-dot';
            st.appendChild(dot);
            st.appendChild(document.createTextNode(producing ? 'Vyrába' : 'Nevyrába'));

            btn.appendChild(txt);
            btn.appendChild(st);
            btn.addEventListener('click', () => openWorkcenter(w));
            return btn;
        }

        function groupBody(groupName, items) {
            const state = filters[groupName] || (filters[groupName] = { q: '', only: 'all' });

            const body = document.createElement('div');
            body.className = 'ov-gb';

            const tools = document.createElement('div');
            tools.className = 'ov-tools';
            const q = document.createElement('input');
            q.type = 'search';
            q.placeholder = 'Hľadať pracovisko…';
            q.value = state.q;
            const sel = document.createElement('select');
            sel.innerHTML = '<option value="all">Všetky</option><option value="run">Vyrába</option><option value="idle">Nevyrába</option>';
            sel.value = state.only;
            tools.appendChild(q);
            tools.appendChild(sel);

            const chips = document.createElement('div');
            chips.className = 'ov-chips';

            function paint() {
                chips.innerHTML = '';
                const needle = normalize(state.q);
                const shown = items.filter((w) => {
                    const producing = (w.count || 0) > 0;
                    if (state.only === 'run' && !producing) return false;
                    if (state.only === 'idle' && producing) return false;
                    if (!needle) return true;
                    return normalize(w.workcenter + ' ' + w.workcenterDescription).indexOf(needle) !== -1;
                });
                if (shown.length === 0) {
                    const empty = document.createElement('div');
                    empty.className = 'ov-empty';
                    empty.textContent = 'Nič nezodpovedá zadaniu.';
                    chips.appendChild(empty);
                    return;
                }
                shown.forEach((w) => chips.appendChild(chip(w)));
            }

            q.addEventListener('input', () => { state.q = q.value; paint(); });
            q.addEventListener('click', (e) => e.stopPropagation());
            sel.addEventListener('change', () => { state.only = sel.value; paint(); });

            body.appendChild(tools);
            body.appendChild(chips);
            paint();
            return body;
        }

        function render(host, list) {
            host.innerHTML = '';

            const hi = document.createElement('div');
            hi.className = 'ov-hi';
            hi.innerHTML = '<div class="k">VITAJTE SPÄŤ</div><div class="n"></div>' +
                '<div class="s">Vyberte pracovisko a začnite pracovať.</div>';
            hi.querySelector('.n').textContent = userName() || 'Pracoviská';
            host.appendChild(hi);

            // zoskupenie podla pola `area`; co ho nema, ide do "Ostatne"
            const groups = {};
            list.forEach((w) => {
                const key = w.area || 'Ostatné';
                (groups[key] = groups[key] || []).push(w);
            });
            const order = ORDER.filter((g) => groups[g])
                .concat(Object.keys(groups).filter((g) => ORDER.indexOf(g) === -1).sort());

            order.forEach((gName) => {
                const items = groups[gName].slice().sort((a, b) => String(a.workcenter).localeCompare(String(b.workcenter)));
                const online = items.reduce((acc, w) => acc + (w.count || 0), 0);
                const isOpen = expanded.has(gName);

                const box = document.createElement('div');
                box.className = 'ov-g' + (isOpen ? ' open' : '');

                const head = document.createElement('button');
                head.type = 'button';
                head.className = 'ov-gh';
                head.innerHTML =
                    '<div class="ov-ic">' + (ICONS[gName] || '🏭') + '</div>' +
                    '<div class="ov-gt"><b></b><span></span></div>' +
                    '<div class="ov-num"><b>' + items.length + '</b><span>PRACOVÍSK</span></div>' +
                    '<div class="ov-num on"><b>' + online + '</b><span>ONLINE</span></div>' +
                    '<div class="ov-ch">' + (isOpen ? '⌄' : '›') + '</div>';
                head.querySelector('.ov-gt b').textContent = gName;
                head.querySelector('.ov-gt span').textContent = SUBTITLES[gName] || '';
                head.addEventListener('click', () => {
                    if (expanded.has(gName)) expanded.delete(gName); else expanded.add(gName);
                    lastSignature = null;
                    ensureOverview();
                });
                box.appendChild(head);
                if (isOpen) box.appendChild(groupBody(gName, items));
                host.appendChild(box);
            });
        }

        function ensureOverview() {
            if (!document.getElementById(HOME_BUTTON_ID)) return;   // len uvodna obrazovka

            const panel = document.getElementById(PANEL_ID);
            const content = panel && panel.querySelector('.sapMPanelContent');
            if (!content) return;

            const list = workcenterList();
            if (!list.length) return;   // appka este taha data zo SAP (~25 s)

            injectStyles();
            document.body.classList.add(BODY_CLASS);

            let host = document.getElementById(OVERVIEW_ID);
            if (!host) {
                host = document.createElement('div');
                host.id = OVERVIEW_ID;
                content.insertBefore(host, content.firstChild);
            } else if (host.parentElement !== content) {
                content.insertBefore(host, content.firstChild);
            }

            // prekreslit len ked sa nieco zmenilo (clovek, pocty, rozbalenie)
            const signature = userName() + '#' +
                list.map((w) => w.workcenter + ':' + (w.count || 0)).join('|') + '#' +
                Array.from(expanded).join(',');
            if (signature === lastSignature) return;
            lastSignature = signature;

            render(host, list);
        }

        DomWatch.add(ensureOverview);
        // `count` sa meni bez zmeny DOM, preto aj casovac (rovnako ako Python verzia)
        onReady(() => setInterval(ensureOverview, 1500));
    }

    /* --------------------- 3.9 Ladiaci vypis ---------------------------- */

    function modDebugLog() {
        XhrBus.subscribe((ev) => {
            console.log('--- [PDA executeBO] ---');
            console.log('request:', ev.requestRaw);
            console.log('response:', ev.responseRaw);
            console.log('-----------------------');
        });
    }

    /* ========================================================================
     *  4. ZOZNAM MODULOV
     *     Nove vylepsenie = napisat funkciu vyssie a pridat sem jeden riadok.
     * ====================================================================== */

    const MODULES = [
        {
            id: 'header',
            name: 'Vylepšená hlavička',
            desc: 'Slovenské popisky pri ikonách, väčšie meno používateľa, zvýraznené odhlásenie.',
            def: true,
            run: modEnhancedHeader,
        },
        {
            id: 'statusButtons',
            name: 'Farebné stavové tlačidlá',
            desc: 'Výroba zelená, prestoj oranžový, chyba červená — a zoradené podľa dôležitosti.',
            def: true,
            run: modStatusButtons,
        },
        {
            id: 'crossSearch',
            name: 'Vyhľadávanie zákaziek',
            desc: 'Vyhľadávací riadok, ktorý hľadá naprieč všetkými pracoviskami naraz.',
            def: true,
            run: modCrossSearch,
        },
        {
            id: 'preventBack',
            name: 'Blokovanie tlačidla Späť',
            desc: 'Zabráni nechcenému vypadnutiu z aplikácie cez tlačidlo Späť v prehliadači.',
            def: true,
            run: modPreventBack,
        },
        {
            id: 'usersPanel',
            name: 'Panel rýchleho prepínania používateľov',
            desc: 'Bočný panel s tlačidlami na prihlásenie. Používateľov zadáš nižšie v nastaveniach.',
            def: false,
            needs: 'Zdieľaný terminál',
            run: modUsersPanel,
        },
        {
            id: 'scanner',
            name: 'Čiarový skener a RFID karty',
            desc: 'Číta skener aj čítačku kariet. Karta prepne používateľa, číslo zákazky sa vloží do vyhľadávania.',
            def: false,
            needs: 'Hardvér',
            run: modInputListener,
        },
        {
            id: 'drawing',
            name: 'Tlačidlo výkresu',
            desc: 'Načíta Excel s výkresmi a k otvorenej zákazke ukáže číslo výkresu a revíziu.',
            def: false,
            needs: 'Firemná sieť',
            run: modDrawingButton,
        },
        {
            id: 'overview',
            name: 'Prehľad pracovísk na úvode',
            desc: 'Zbalí úvodnú obrazovku do kategórií (Assembly, Welding, Machining…). Vidíš len čísla strojov a či pracujú; klik otvorí stroj.',
            def: true,
            run: modWorkcenterOverview,
        },
        {
            id: 'debug',
            name: 'Ladiaci výpis do konzoly',
            desc: 'Vypisuje sieťovú komunikáciu aplikácie do konzoly prehliadača. Bežne netreba.',
            def: false,
            run: modDebugLog,
        },
    ];

    /* ========================================================================
     *  5. PANEL NASTAVENI
     * ====================================================================== */

    const OVERLAY_ID = '__pda_settings_overlay__';
    const GEAR_ID = '__pda_settings_gear__';

    function injectSettingsStyles() {
        if (document.getElementById('__pda_settings_styles__')) return;
        const style = document.createElement('style');
        style.id = '__pda_settings_styles__';
        style.textContent = `
#${GEAR_ID} {
  position: fixed; right: 14px; bottom: 14px; z-index: 2147483000;
  width: 38px; height: 38px; border-radius: 50%; border: none;
  background: #313175; color: #fff; font-size: 19px; line-height: 1;
  cursor: pointer; opacity: .55; box-shadow: 0 2px 8px rgba(0,0,0,.3);
  transition: opacity .15s;
}
#${GEAR_ID}:hover { opacity: 1; }
#${OVERLAY_ID} {
  position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 2147483001;
  display: flex; align-items: center; justify-content: center;
  font: 14px/1.5 -apple-system, "Segoe UI", Roboto, sans-serif; color: #1a1a1f;
}
.pda-set-box {
  background: #fff; border-radius: 12px; width: 94%; max-width: 720px;
  max-height: 88vh; display: flex; flex-direction: column; overflow: hidden;
  box-shadow: 0 10px 40px rgba(0,0,0,.35);
}
.pda-set-head {
  display: flex; justify-content: space-between; align-items: center;
  padding: 14px 18px; background: #313175; color: #fff; flex-shrink: 0;
}
.pda-set-head b { font-size: 1.02rem; }
.pda-set-x { background: none; border: 0; color: #fff; font-size: 24px; line-height: 1; cursor: pointer; }
.pda-set-body { padding: 6px 18px 18px; overflow: auto; }
.pda-set-body h3 {
  font-size: .76rem; text-transform: uppercase; letter-spacing: .07em;
  color: #6b7180; margin: 20px 0 8px; font-weight: 600;
}
.pda-mod {
  display: flex; gap: 11px; align-items: flex-start;
  padding: 10px 12px; border: 1px solid #e3e6eb; border-radius: 9px; margin-bottom: 7px;
}
.pda-mod input { margin-top: 3px; width: 17px; height: 17px; flex-shrink: 0; cursor: pointer; }
.pda-mod .nm { font-weight: 600; }
.pda-mod .ds { color: #5c6370; font-size: .87rem; }
.pda-badge {
  font-size: .66rem; text-transform: uppercase; letter-spacing: .05em;
  background: #eef0f4; color: #5c6370; border-radius: 20px; padding: 2px 8px; margin-left: 7px;
  white-space: nowrap;
}
.pda-set-body table { width: 100%; border-collapse: collapse; }
.pda-set-body th {
  text-align: left; font-size: .74rem; text-transform: uppercase;
  letter-spacing: .05em; color: #6b7180; padding: 0 6px 5px 0; font-weight: 600;
}
.pda-set-body td { padding: 0 6px 6px 0; }
.pda-set-body input[type=text], .pda-set-body input[type=password] {
  width: 100%; box-sizing: border-box; padding: 6px 9px;
  border: 1px solid #ccd1d9; border-radius: 6px; font: inherit; font-size: .9rem;
}
.pda-del { background: #fdecea; border: 1px solid #f0b4ae; color: #b0201a;
  border-radius: 6px; cursor: pointer; padding: 5px 10px; font-size: .82rem; }
.pda-add { background: #eef0f4; border: 1px solid #ccd1d9; border-radius: 7px;
  cursor: pointer; padding: 6px 13px; font-size: .85rem; margin-top: 3px; }
.pda-note { color: #6b7180; font-size: .82rem; margin: 7px 0 0; }
.pda-set-foot {
  display: flex; justify-content: flex-end; gap: 9px; align-items: center;
  padding: 13px 18px; border-top: 1px solid #e3e6eb; flex-shrink: 0; background: #fafbfc;
}
.pda-btn-primary { background: #313175; color: #fff; border: 0; border-radius: 8px;
  padding: 9px 18px; font-weight: 600; cursor: pointer; font-size: .9rem; }
.pda-btn-plain { background: #fff; border: 1px solid #ccd1d9; border-radius: 8px;
  padding: 9px 16px; cursor: pointer; font-size: .9rem; }
`;
        document.head.appendChild(style);
    }

    function openSettings() {
        if (document.getElementById(OVERLAY_ID)) return;
        injectSettingsStyles();

        const draftModules = {};
        MODULES.forEach((m) => { draftModules[m.id] = isModuleOn(m); });
        const draftUsers = settings.users.map((u) => ({ ...u }));
        const draftPdm = { ...settings.pdm };
        const draftExcel = { ...settings.excel };
        const draftGroups = { value: String(settings.groups || '') };

        const overlay = document.createElement('div');
        overlay.id = OVERLAY_ID;

        const box = document.createElement('div');
        box.className = 'pda-set-box';

        // --- hlavicka ---
        const head = document.createElement('div');
        head.className = 'pda-set-head';
        head.innerHTML = '<b>Nastavenia PDA Suite</b>';
        const closeBtn = document.createElement('button');
        closeBtn.className = 'pda-set-x';
        closeBtn.type = 'button';
        closeBtn.textContent = '×';
        head.appendChild(closeBtn);

        // --- telo ---
        const body = document.createElement('div');
        body.className = 'pda-set-body';

        const hMods = document.createElement('h3');
        hMods.textContent = 'Moduly';
        body.appendChild(hMods);

        MODULES.forEach((m) => {
            const row = document.createElement('div');
            row.className = 'pda-mod';

            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = draftModules[m.id];
            cb.addEventListener('change', () => { draftModules[m.id] = cb.checked; });

            const txt = document.createElement('div');
            const nm = document.createElement('div');
            nm.className = 'nm';
            nm.textContent = m.name;
            if (m.needs) {
                const badge = document.createElement('span');
                badge.className = 'pda-badge';
                badge.textContent = m.needs;
                nm.appendChild(badge);
            }
            const ds = document.createElement('div');
            ds.className = 'ds';
            ds.textContent = m.desc;
            txt.appendChild(nm);
            txt.appendChild(ds);

            const lbl = document.createElement('label');
            lbl.style.cssText = 'display:flex;gap:11px;align-items:flex-start;cursor:pointer;';
            lbl.appendChild(cb);
            lbl.appendChild(txt);
            row.appendChild(lbl);
            body.appendChild(row);
        });

        // --- pouzivatelia ---
        const hUsers = document.createElement('h3');
        hUsers.textContent = 'Používatelia pre rýchle prepínanie';
        body.appendChild(hUsers);

        const usersWrap = document.createElement('div');
        body.appendChild(usersWrap);

        function renderUsers() {
            usersWrap.innerHTML = '';
            const table = document.createElement('table');
            table.innerHTML =
                '<thead><tr><th>Meno (presne ako v zozname PDA)</th><th style="width:120px">Osobné číslo</th><th style="width:150px">ID karty</th><th style="width:40px"></th></tr></thead>';
            const tbody = document.createElement('tbody');

            draftUsers.forEach((u, i) => {
                const tr = document.createElement('tr');

                const mk = (field, type) => {
                    const td = document.createElement('td');
                    const inp = document.createElement('input');
                    inp.type = type;
                    inp.value = u[field] || '';
                    inp.addEventListener('input', () => { draftUsers[i][field] = inp.value.trim(); });
                    td.appendChild(inp);
                    return td;
                };

                tr.appendChild(mk('username', 'text'));
                tr.appendChild(mk('password', 'password'));
                tr.appendChild(mk('cardId', 'text'));

                const tdDel = document.createElement('td');
                const del = document.createElement('button');
                del.type = 'button';
                del.className = 'pda-del';
                del.textContent = '×';
                del.title = 'Odstrániť';
                del.addEventListener('click', () => { draftUsers.splice(i, 1); renderUsers(); });
                tdDel.appendChild(del);
                tr.appendChild(tdDel);

                tbody.appendChild(tr);
            });

            table.appendChild(tbody);
            usersWrap.appendChild(table);

            const add = document.createElement('button');
            add.type = 'button';
            add.className = 'pda-add';
            add.textContent = '+ Pridať používateľa';
            add.addEventListener('click', () => {
                draftUsers.push({ username: '', password: '', cardId: '' });
                renderUsers();
            });
            usersWrap.appendChild(add);

            const note = document.createElement('p');
            note.className = 'pda-note';
            note.textContent = 'Tieto údaje sa ukladajú len lokálne v Tampermonkey na tomto počítači — nikdy sa neposielajú na GitHub ani nikam inam.';
            usersWrap.appendChild(note);
        }
        renderUsers();

        // --- hromadne zadanie / prenos na iny pocitac ---
        const ioWrap = document.createElement('div');
        ioWrap.style.cssText = 'margin-top:12px;border:1px dashed #ccd1d9;border-radius:9px;padding:11px 12px;';

        const ioTitle = document.createElement('div');
        ioTitle.textContent = 'Hromadné zadanie / prenos na iný počítač';
        ioTitle.style.cssText = 'font-weight:600;font-size:.88rem;margin-bottom:4px;';

        const ioHelp = document.createElement('div');
        ioHelp.className = 'pda-note';
        ioHelp.style.marginTop = '0';
        const IO_HELP_DEFAULT = 'Jeden používateľ na riadok vo formáte:  Meno;osobné číslo;ID karty   (ID karty môže ostať prázdne)';
        ioHelp.textContent = IO_HELP_DEFAULT;

        const ta = document.createElement('textarea');
        ta.rows = 5;
        ta.spellcheck = false;
        ta.placeholder = 'Ján Novák;12345;0116984be8';
        ta.style.cssText = 'width:100%;box-sizing:border-box;margin-top:7px;padding:8px 10px;' +
            'border:1px solid #ccd1d9;border-radius:6px;font:12px/1.5 ui-monospace,Consolas,monospace;resize:vertical;';

        const ioBtns = document.createElement('div');
        ioBtns.style.cssText = 'display:flex;gap:8px;margin-top:7px;';

        const btnImport = document.createElement('button');
        btnImport.type = 'button';
        btnImport.className = 'pda-add';
        btnImport.style.marginTop = '0';
        btnImport.textContent = 'Načítať z textu';
        btnImport.addEventListener('click', () => {
            const parsed = ta.value
                .split('\n')
                .map((line) => line.trim())
                .filter((line) => line && line.indexOf(';') !== -1)
                .map((line) => {
                    const parts = line.split(';').map((s) => (s || '').trim());
                    return { username: parts[0], password: parts[1] || '', cardId: parts[2] || '' };
                })
                .filter((u) => u.username);

            if (parsed.length === 0) {
                ioHelp.style.color = '#b0201a';
                ioHelp.textContent = 'Nenašiel sa žiadny platný riadok. Formát je: Meno;osobné číslo;ID karty';
                return;
            }

            draftUsers.length = 0;
            parsed.forEach((u) => draftUsers.push(u));
            renderUsers();
            ioHelp.style.color = '';
            ioHelp.textContent = 'Načítaných používateľov: ' + parsed.length + '. Ešte to ulož tlačidlom dole.';
        });

        const btnExport = document.createElement('button');
        btnExport.type = 'button';
        btnExport.className = 'pda-add';
        btnExport.style.marginTop = '0';
        btnExport.textContent = 'Vypísať súčasných';
        btnExport.addEventListener('click', () => {
            ta.value = draftUsers.map((u) => [u.username, u.password, u.cardId].join(';')).join('\n');
            ioHelp.style.color = '';
            ioHelp.textContent = IO_HELP_DEFAULT;
        });

        ioBtns.appendChild(btnImport);
        ioBtns.appendChild(btnExport);
        ioWrap.appendChild(ioTitle);
        ioWrap.appendChild(ioHelp);
        ioWrap.appendChild(ta);
        ioWrap.appendChild(ioBtns);
        body.appendChild(ioWrap);

        // --- vykresy ---
        const hPdm = document.createElement('h3');
        hPdm.textContent = 'Služba výkresov (PDM)';
        body.appendChild(hPdm);

        const pdmTable = document.createElement('table');
        pdmTable.innerHTML = '<thead><tr><th>Adresa služby</th><th style="width:210px">API kľúč (nepovinné)</th></tr></thead>';
        const pdmBody = document.createElement('tbody');
        const pdmTr = document.createElement('tr');

        const tdBase = document.createElement('td');
        const inpBase = document.createElement('input');
        inpBase.type = 'text';
        inpBase.value = draftPdm.base || '';
        inpBase.placeholder = 'http://172.16.77.134:9000';
        inpBase.addEventListener('input', () => { draftPdm.base = inpBase.value.trim(); });
        tdBase.appendChild(inpBase);

        const tdKey = document.createElement('td');
        const inpKey = document.createElement('input');
        inpKey.type = 'password';
        inpKey.value = draftPdm.key || '';
        inpKey.addEventListener('input', () => { draftPdm.key = inpKey.value.trim(); });
        tdKey.appendChild(inpKey);

        pdmTr.appendChild(tdBase);
        pdmTr.appendChild(tdKey);
        pdmBody.appendChild(pdmTr);
        pdmTable.appendChild(pdmBody);
        body.appendChild(pdmTable);

        // --- Excel s vykresmi ---
        const hExcel = document.createElement('h3');
        hExcel.textContent = 'Excel s výkresmi';
        body.appendChild(hExcel);

        const excelUrlTable = document.createElement('table');
        excelUrlTable.innerHTML = '<thead><tr><th>Adresa Excelu (nepovinné)</th></tr></thead>';
        const excelUrlBody = document.createElement('tbody');
        const excelUrlTr = document.createElement('tr');
        const tdUrl = document.createElement('td');
        const inpUrl = document.createElement('input');
        inpUrl.type = 'text';
        inpUrl.value = draftExcel.url || '';
        inpUrl.placeholder = 'http://172.16.77.134:9000/automated180.xlsx';
        inpUrl.addEventListener('input', () => { draftExcel.url = inpUrl.value.trim(); });
        tdUrl.appendChild(inpUrl);
        excelUrlTr.appendChild(tdUrl);
        excelUrlBody.appendChild(excelUrlTr);
        excelUrlTable.appendChild(excelUrlBody);
        body.appendChild(excelUrlTable);

        const excelNote = document.createElement('p');
        excelNote.className = 'pda-note';
        excelNote.innerHTML =
            'Ak sem zadáš adresu začínajúcu <b>http://</b> alebo <b>https://</b>, Excel sa stiahne sám pri každom otvorení aplikácie.<br>' +
            'Ak pole necháš prázdne, súbor sa vyberá ručne tlačidlom <b>„Vybrať Excel"</b> a prehliadač si ho zapamätá.<br>' +
            '<b>Cesta na disku ani sieťový disk sem nepatria</b> (<code>C:\\…</code>, <code>\\\\server\\…</code>) — prehliadač zo zásady ' +
            'nevie otvoriť súbor podľa cesty, vtedy treba tlačidlo na ručný výber.';
        body.appendChild(excelNote);

        const colTable = document.createElement('table');
        colTable.style.marginTop = '9px';
        colTable.innerHTML =
            '<thead><tr><th>Stĺpec — číslo zákazky</th><th>Stĺpec — číslo výkresu</th><th>Stĺpec — verzia</th></tr></thead>';
        const colBody = document.createElement('tbody');
        const colTr = document.createElement('tr');

        [['colOrder', 'H'], ['colDrawing', 'AH'], ['colVersion', 'AI']].forEach(([field, ph]) => {
            const td = document.createElement('td');
            const inp = document.createElement('input');
            inp.type = 'text';
            inp.value = draftExcel[field] || '';
            inp.placeholder = ph;
            inp.addEventListener('input', () => { draftExcel[field] = inp.value.trim().toUpperCase(); });
            td.appendChild(inp);
            colTr.appendChild(td);
        });

        colBody.appendChild(colTr);
        colTable.appendChild(colBody);
        body.appendChild(colTable);

        const colNote = document.createElement('p');
        colNote.className = 'pda-note';
        colNote.textContent = 'Písmená stĺpcov tak, ako ich vidíš v Exceli. Predvolene H, AH a AI. ' +
            'Excel je nepovinný — ak chýba, výkres sa hľadá priamo podľa čísla materiálu.';
        body.appendChild(colNote);


        // --- paticka ---
        const foot = document.createElement('div');
        foot.className = 'pda-set-foot';

        const cancel = document.createElement('button');
        cancel.type = 'button';
        cancel.className = 'pda-btn-plain';
        cancel.textContent = 'Zrušiť';

        const save = document.createElement('button');
        save.type = 'button';
        save.className = 'pda-btn-primary';
        save.textContent = 'Uložiť a obnoviť stránku';

        foot.appendChild(cancel);
        foot.appendChild(save);

        box.appendChild(head);
        box.appendChild(body);
        box.appendChild(foot);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        function close() {
            overlay.remove();
            document.removeEventListener('keydown', onEsc);
        }
        function onEsc(e) { if (e.key === 'Escape') close(); }

        closeBtn.addEventListener('click', close);
        cancel.addEventListener('click', close);
        overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
        document.addEventListener('keydown', onEsc);

        save.addEventListener('click', () => {
            saveJson(KEY_MODULES, draftModules);
            saveJson(KEY_USERS, draftUsers.filter((u) => u.username));
            saveJson(KEY_PDM, draftPdm);
            saveJson(KEY_EXCEL, draftExcel);
            saveJson(KEY_GROUPS, draftGroups.value);
            close();
            shared.intentionalReload = true;
            location.reload();
        });
    }

    function ensureGear() {
        if (document.getElementById(GEAR_ID)) return;
        injectSettingsStyles();
        const gear = document.createElement('button');
        gear.id = GEAR_ID;
        gear.type = 'button';
        gear.title = 'Nastavenia PDA Suite';
        gear.textContent = '⚙';
        gear.addEventListener('click', openSettings);
        document.body.appendChild(gear);
    }

    /* ========================================================================
     *  6. SPUSTENIE
     * ====================================================================== */

    const active = [];
    MODULES.forEach((mod) => {
        if (!isModuleOn(mod)) return;
        try {
            mod.run();
            active.push(mod.id);
        } catch (e) {
            console.warn(LOG, 'modul "' + mod.id + '" sa nepodarilo spustiť', e);
        }
    });

    console.log(LOG, 'aktívne moduly:', active.length ? active.join(', ') : 'žiadne');

    onReady(ensureGear);
    DomWatch.add(ensureGear);

    try {
        GM_registerMenuCommand('Nastavenia PDA Suite', openSettings);
    } catch (e) { /* nedostupne mimo Tampermonkey */ }
})();
