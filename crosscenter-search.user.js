// ==UserScript==
// @name         PDA - Cross-workcenter search (XHR data)
// @namespace    http://tampermonkey.net/
// @version      0.0.6
// @description  Searchbar na vyhladavanie naprieč vsetkymi pracoviskami
// @author       Gabris
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/crosscenter-search.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/crosscenter-search.user.js
// @match        https://hf.simplifier.cloud/appDirect/PDA/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // ---------- XHR intercept: naplnanie PDA_ORDERS_INDEX ----------
    const TARGET_URL_SUBSTRING = '/client/1.0/executeBO';
    const TARGET_BO_METHOD = 'getOperationListTempForWorkcenters';

    window.PDA_ORDERS_INDEX = window.PDA_ORDERS_INDEX || [];
    let lastIndexUpdate = null;

    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
        this._pda_url = url;
        this._pda_method = method;
        return originalOpen.call(this, method, url, ...rest);
    };

    XMLHttpRequest.prototype.send = function (body) {
        const url = this._pda_url || '';

        if (this._pda_url && this._pda_url.includes(TARGET_URL_SUBSTRING)) {
        let parsedBody = null;
        try { parsedBody = JSON.parse(body); } catch (e) { /* ignore */ }

        if (parsedBody && parsedBody.BOMethod === TARGET_BO_METHOD) {
            this.addEventListener('load', function () {
            try {
                const data = JSON.parse(this.responseText);
                const workcenters = data.result.aOperationListTempForWorkcenter || [];

                const flatOrders = [];
                workcenters.forEach((wc) => {
                (wc.operationList || []).forEach((op) => {
                    flatOrders.push({
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

                window.PDA_ORDERS_INDEX = flatOrders;
                lastIndexUpdate = new Date();
                console.log('[PDA orders] index naplneny, pocet zakaziek:', flatOrders.length);
                refreshSearchUIIfPresent();
            } catch (e) {
                console.warn('[PDA orders] chyba pri spracovani response', e);
            }
            });
        }
        }

        return originalSend.apply(this, arguments);
    };

    // ---------- Layout ----------
    const PANEL_ID = 'Main--Workcenter_Panel';
    const CONTENT_ID = 'Main--Workcenter_Panel-content';
    const WORKCENTER_TILE_SELECTOR = '[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]';
    const LIST_UL_ID = 'WorkcenterDetail--Work_List-listUl';
    const HOME_BUTTON_ID = 'Main--Button_HomeScreen';

    const PARENT_SECTION_ID = 'Main--MainPage-cont';
    const TASKS_PANEL_ID = 'Main--Tasks_Panel';
    const SIDEBAR_ID = '__pda_search_sidebar__';
    const UI_ID = '__pda_custom_search_ui__';

    const TILE_WAIT_TIMEOUT = 15000;
    const SETTLE_DELAY = 350;

    const TILE_ID_PREFIX = 'Main--Workcenter_Toolbar-';
    const LIST_ID_PREFIX = 'Main--List2-';

    let opening = false;
    let renderFn = null;
    let lastAutoOpenedKey = null;

    // ---------- Dekodovanie dat zo skenera (SK klavesnica cita cisla ako specialne znaky) ----------
    const SCANNER_CHAR_MAP = {
        '+': '1',
        'ľ': '2',
        'š': '3',
        'č': '4',
        'ť': '5',
        'ž': '6',
        'ý': '7',
        'á': '8',
        'í': '9',
        'é': '0',
    };

    function decodeScannerInput(raw) {
        // prvy znak (napr. "J") je marker skenera, nie je sucastou dat
        const stripped = raw.slice(1);

        let translated = '';
        for (const ch of stripped) {
            if (ch === '.') {
                translated += '-';
            } else if (SCANNER_CHAR_MAP.hasOwnProperty(ch)) {
                translated += SCANNER_CHAR_MAP[ch];
            } else {
                translated += ch; // neznamy znak - necha tak ako je
            }
        }

        const dashIndex = translated.indexOf('-');
        if (dashIndex === -1) return translated;

        const orderPart = translated.slice(0, dashIndex);
        const opPart = translated.slice(dashIndex + 1);
        const paddedOpPart = opPart.padStart(4, '0');

        return orderPart + '-' + paddedOpPart;
    }

    function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

    function waitFor(checkFn, { interval = 200, timeout = 10000 } = {}) {
        return new Promise((resolve, reject) => {
        const start = Date.now();
        const timer = setInterval(() => {
            const result = checkFn();
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

    function pressElement(domEl) {
        let control = null;
        try {
        if (window.sap && sap.ui && sap.ui.core && typeof sap.ui.core.Element.closestTo === 'function') {
            control = sap.ui.core.Element.closestTo(domEl);
        }
        if (!control && window.sap && sap.ui && sap.ui.getCore) {
            control = sap.ui.getCore().byId(domEl.id);
        }
        } catch (e) { /* ignore */ }
        if (control && typeof control.firePress === 'function') {
        control.firePress();
        return;
        }
        domEl.click();
    }

    function getWorkcenterName(tile) {
        const suffix = tile.id.replace('Main--Workcenter_Toolbar-', '');
        const nameEl = document.getElementById(`Main--WorkcenterDescription_Label-${suffix}-bdi`);
        return nameEl ? nameEl.textContent.trim() : tile.id;
    }

    function extractField(li, marker) {
        const el = li.querySelector(`[id*="${marker}"][id$="-bdi"]`);
        return el ? el.textContent.trim() : '';
    }

    function formatProductionOrder(item) {
        return `${item.productionOrderNo} - ${item.operationNo} - ${item.sequenceNo}`;
    }

    function getItemKey(item) {
        return item.workcenter + '|' + item.productionOrderNo + '|' + item.operationNo + '|' + item.sequenceNo;
    }

    function setStatus(msg) {
        const el = document.getElementById('__pda_status__');
        if (el) el.textContent = msg;
    }

    function getListIdForTile(tile) {
        if (!tile.id || !tile.id.startsWith(TILE_ID_PREFIX)) return null;
        const suffix = tile.id.slice(TILE_ID_PREFIX.length); // napr. "Main--ui_layout_Grid3-5"
        return LIST_ID_PREFIX + suffix; // "Main--List2-Main--ui_layout_Grid3-5"
    }

    function extractOrderText(li) {
        // Title control vnoreny v li, ID koncici na "-inner" (rovnaky pattern ako inde v apke)
        const titleEl = li.querySelector('[id*="Title"][id$="-inner"]');
        return titleEl ? titleEl.textContent.trim() : '';
    }

    function pdaResolveControl(el) {
        while (el && el.nodeType === 1) {
            if (window.jQuery && typeof jQuery(el).control === 'function') {
            const arr = jQuery(el).control();
            if (arr && arr.length && arr[0]) return arr[0];
            }
            if (window.sap && sap.ui && sap.ui.core && sap.ui.core.Element && typeof sap.ui.core.Element.closestTo === 'function') {
            const c = sap.ui.core.Element.closestTo(el);
            if (c) return c;
            }
            if (el.id && window.sap && sap.ui && sap.ui.getCore) {
            const c = sap.ui.getCore().byId(el.id);
            if (c) return c;
            }
            el = el.parentElement;
        }
        
        return null;
    }

    function fireListItemPress(li) {
        let itemControl = pdaResolveControl(li);
        if (!itemControl) return false;

        // najdi rodicovsky control, ktory realne ma "itemPress" event (typicky List_1_96)
        let listControl = itemControl;
        while (listControl && typeof listControl.fireItemPress !== 'function') {
            listControl = listControl.getParent && listControl.getParent();
        }
        if (!listControl) return false;

        // najdi ten control, ktory je priamym childom listControl (CustomListItem)
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
            if (!tile) {
                setStatus('Pracovisko sa nenašlo.');
                return;
            }

            const listId = getListIdForTile(tile);
            let list = listId ? document.getElementById(listId) : null;

            // Ak zoznam este nie je v DOM (panel nemusi byt rozbaleny), klikneme na dlazdicu a pockame
            if (!list) {
                pressElement(tile);
                list = listId
                    ? await waitFor(() => document.getElementById(listId), { timeout: TILE_WAIT_TIMEOUT })
                    : null;
            }

            if (!list) {
                setStatus('Zoznam zákaziek pre toto pracovisko sa nenašiel.');
                return;
            }

            list.scrollIntoView({ block: 'center' });
            await sleep(SETTLE_DELAY);

            const targetText = formatProductionOrder(item);
            let li = Array.from(list.querySelectorAll('li')).find(
                (el) => extractOrderText(el) === targetText
            );

            // Ak sa zoznam este dorenderoval po kliku na dlazdicu, dame mu druhu sancu
            if (!li) {
            await sleep(SETTLE_DELAY);
            li = Array.from(list.querySelectorAll('li')).find(
                (el) => extractOrderText(el) === targetText
            );
            }

            if (!li) {
                setStatus('Konkrétna zákazka sa nenašla, otvorené je aspoň pracovisko.');
                return;
            }

            li.scrollIntoView({ block: 'center' });

            const fired = fireListItemPress(li);
            if (!fired) {
                console.warn('[PDA search] fireItemPress zlyhalo, skusam fallback klik');
                li.click();
            }

            setStatus(fired ? 'Otvorené: ' + targetText : 'Chyba pri otváraní zákazky.');
        } catch (e) {
            console.warn('[PDA search] chyba pri otváraní položky', e);
            setStatus('Chyba pri otváraní zákazky.');
        } finally {
            opening = false;
        }
    }

    function ensureLayout() {
        const parentSection = document.getElementById(PARENT_SECTION_ID);
        const tasksPanel = document.getElementById(TASKS_PANEL_ID);
        const workcenterPanel = document.getElementById(PANEL_ID);

        if (!parentSection || !tasksPanel || !workcenterPanel) return null;

        let searchContainer = document.getElementById(SIDEBAR_ID);
        if (!searchContainer) {
            searchContainer = document.createElement('div');
            searchContainer.id = SIDEBAR_ID;
            searchContainer.style.width = '100%';
            searchContainer.style.boxSizing = 'border-box';
            searchContainer.style.maxHeight = '80vh';
            searchContainer.style.overflowY = 'auto';
            searchContainer.style.margin = '10px 0';
            searchContainer.style.backgroundColor = '#ffffff';
            searchContainer.style.borderRadius = '10px';
        }

        // Vlozime nas searchbar medzi Tasks_Panel a Workcenter_Panel, ak tam este nie je
        if (searchContainer.nextElementSibling !== workcenterPanel || searchContainer.parentElement !== parentSection) {
            parentSection.insertBefore(searchContainer, workcenterPanel);
        }

        if (!document.getElementById(UI_ID)) buildCustomSearchUI(searchContainer);
        return { searchContainer };
    }

    function buildCustomSearchUI(sidebar) {
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

        const searchToolbarContainer = document.createElement('div');
        searchToolbarContainer.className = 'sapMListInfoTBarContainer';
        const searchToolbar = document.createElement('div');
        searchToolbar.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBClear sapMTB-Transparent-CTX sapMListInfoTBar';
        const sf = document.createElement('div');
        sf.className = 'sapMSF sapMSFVal sapMBarChild sapMTBShrinkItem';
        sf.style.width = '100%';
        const form = document.createElement('form');
        form.className = 'sapMSFF';
        form.addEventListener('submit', (e) => e.preventDefault());
        const input = document.createElement('input');
        input.type = 'search';
        input.autocomplete = 'off';
        input.placeholder = 'Order number / material';
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
        searchToolbar.appendChild(sf);
        searchToolbarContainer.appendChild(searchToolbar);

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
            li.innerHTML = `
                <div class="sapMLIBContent">
                <div class="sapMFlexBoxFit sapMFlexBox sapMHBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent" style="height:100%;width:100%;">
                    <div class="sapMFlexBox sapMVBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent sapMFlexItem" style="height:100%;width:100%;">
                    <span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="font-weight:bold;text-align:left;">
                        <span class="sapMLabelTextWrapper"><bdi>${formatProductionOrder(it)}</bdi></span>
                    </span>
                    <span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;">
                        <span class="sapMLabelTextWrapper"><bdi>${it.workcenter}</bdi></span>
                    </span>
                    <span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;">
                        <span class="sapMLabelTextWrapper"><bdi>${it.material || ''}</bdi></span>
                    </span>
                    <span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapMFlexItem" style="text-align:left;">
                        <span class="sapMLabelTextWrapper"><bdi>${it.descriptionShort || ''}</bdi></span>
                    </span>
                    </div>
                </div>
                </div>`;

            li.addEventListener('click', () => openItem(it));
            li.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === '') { e.preventDefault(); openItem(it); } 
            });
            return li;
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

        function handleScannerInput() {
            const decoded = decodeScannerInput(input.value);
            input.value = decoded;
            render(decoded);
        }

        function render(filterText) {
            resultsUl.innerHTML = '';
            const term = filterText.trim().toLowerCase();
            const index = window.PDA_ORDERS_INDEX || [];
            
            if (!term) {
                setStatus(`Index: ${index.length} zákaziek` + (lastIndexUpdate ? ` (aktualiz. ${lastIndexUpdate.toLocaleTimeString()})` : ' (čaká sa na načítanie appky)'));
                lastAutoOpenedKey = null;
                return;
            }

            const matches = index.filter((it) => matchesTerm(it, term));
            setStatus(`${matches.length} výsledok/-ov (z ${index.length} položiek)`);
            matches.slice(0, 200).forEach((it) => resultsUl.appendChild(renderItem(it)));

            if (matches.length === 1) {
                const key = getItemKey(matches[0]);
                if (key !== lastAutoOpenedKey) {
                    lastAutoOpenedKey = key;
                    openItem(matches[0]);

                    // vycistit input a zoznam vysledkov s oneskerenim 500ms
                    setTimeout(() => {
                        input.value = '';
                        resultsUl.innerHTML = '';
                        setStatus(`Index: ${index.length} zákaziek` + (lastIndexUpdate ? ` (aktualiz. ${lastIndexUpdate.toLocaleTimeString()})` : ''));
                        lastAutoOpenedKey = null;
                    }, 500);
                }
            } else {
                lastAutoOpenedKey = null;
            }
        }

        input.addEventListener('input', () => render(input.value));
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                handleScannerInput();
            }
        });

        list.appendChild(header);
        list.appendChild(searchToolbarContainer);
        list.appendChild(status);
        list.appendChild(resultsUl);
        sidebar.appendChild(list);

        render('');
        renderFn = render;
    }

    function refreshSearchUIIfPresent() {
        if (renderFn) {
        const input = document.querySelector(`#${UI_ID} input.sapMSFI`);
        renderFn(input ? input.value : '');
        }
    }

    function isOnMainScreen() {
        return !!(document.getElementById(HOME_BUTTON_ID) && document.getElementById(CONTENT_ID));
    }

    function init() {
        if (!isOnMainScreen()) return;
        ensureLayout();
    }

    let debounceTimer = null;
    function scheduleInit() {
        if (debounceTimer) return;
        debounceTimer = setTimeout(() => {
        debounceTimer = null;
        if (!opening) init();
        }, 500);
    }

    scheduleInit();
    const observer = new MutationObserver(() => {
        if (!document.getElementById(UI_ID) && !opening) scheduleInit();
    });
    observer.observe(document.body, { childList: true, subtree: true });
})();