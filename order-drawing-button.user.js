// ==UserScript==
// @name         PDA - Order drawing button
// @namespace    http://tampermonkey.net/
// @version      0.0.7
// @description  Nacita Excel s vykresmi zo sietoveho disku, sleduje aktualne otvorenu operaciu a zobrazuje cislo vykresu + verziu v tlacidle
// @author       Gabris
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/order-drawing-button.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/order-drawing-button.user.js
// @match        https://hf.simplifier.cloud/appDirect/PDA/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @run-at       document-start
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      172.16.77.134
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==
 
(function () {
    'use strict';
 
    // ---------- Sluzba "Mapa vykresov" (PDM) - konfiguracia ----------
    // Adresu a API kluc nastav tu, podla dokumentacie sluzby.
    const PDM_BASE = 'http://172.16.77.134:9000';
    const PDM_KEY = ''; // ak sluzba vyzaduje X-API-Key, doplň ho sem
 
    // ---------- IndexedDB: ulozenie file handle-u, aby prezil aj refresh stranky ----------
 
    const DB_NAME = 'pda_drawing_db';
    const STORE_NAME = 'handles';
    const HANDLE_KEY = 'excel_file_handle';
 
    function openHandleDb() {
        return new Promise((resolve, reject) => {
            const req = unsafeWindow.indexedDB.open(DB_NAME, 1);
            req.onupgradeneeded = () => {
                req.result.createObjectStore(STORE_NAME);
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }
 
    async function saveHandle(handle) {
        console.log('[PDA drawing-button] saveHandle: zacinam ukladat handle', handle);
        const db = await openHandleDb();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, 'readwrite');
            const putReq = tx.objectStore(STORE_NAME).put(handle, HANDLE_KEY);
            putReq.onerror = () => {
                console.warn('[PDA drawing-button] saveHandle: put() zlyhal', putReq.error);
            };
            tx.oncomplete = () => {
                console.log('[PDA drawing-button] saveHandle: transakcia uspesne dokoncena, handle by mal byt ulozeny');
                resolve();
            };
            tx.onerror = () => {
                console.warn('[PDA drawing-button] saveHandle: transakcia zlyhala', tx.error);
                reject(tx.error);
            };
            tx.onabort = () => {
                console.warn('[PDA drawing-button] saveHandle: transakcia bola abortnuta', tx.error);
                reject(tx.error);
            };
        });
    }
 
    async function loadHandle() {
        console.log('[PDA drawing-button] loadHandle: hladam ulozeny handle v IndexedDB');
        const db = await openHandleDb();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, 'readonly');
            const req = tx.objectStore(STORE_NAME).get(HANDLE_KEY);
            req.onsuccess = () => {
                console.log('[PDA drawing-button] loadHandle: vysledok z IndexedDB =', req.result);
                resolve(req.result || null);
            };
            req.onerror = () => {
                console.warn('[PDA drawing-button] loadHandle: get() zlyhal', req.error);
                reject(req.error);
            };
        });
    }
 
    // ---------- Excel: parsovanie a vytvorenie indexu cislo zakazky -> vykres/verzia ----------
 
    const SHEET_NAME = null; // null = prvy list v subore
 
    const COL_ORDER_NO = 7;    // stlpec H
    const COL_DRAWING_NO = 33; // stlpec AH
    const COL_VERSION = 34;    // stlpec AI
 
    let drawingIndex = {};
    let excelLoaded = false;
 
    function buildDrawingIndex(rows) {
        const index = {};
 
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;
 
            const orderNoRaw = row[COL_ORDER_NO];
            if (orderNoRaw === undefined || orderNoRaw === null || orderNoRaw === '') continue;
 
            const key = String(orderNoRaw).trim();
            if (!key) continue;
 
            index[key] = {
                drawingNo: row[COL_DRAWING_NO] != null ? String(row[COL_DRAWING_NO]).trim() : '',
                version: row[COL_VERSION] != null ? String(row[COL_VERSION]).trim() : '',
            };
        }
 
        console.log('[PDA drawing-button] ukazka prvych 5 klucov v indexe:', Object.keys(index).slice(0, 5));
        return index;
    }
 
    async function loadExcelFromFile(file) {
        const buffer = await file.arrayBuffer();
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = SHEET_NAME || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            console.warn('[PDA drawing-button] list', sheetName, 'sa v exceli nenasiel');
            return;
        }
 
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        drawingIndex = buildDrawingIndex(rows);
        unsafeWindow.PDA_DRAWING_INDEX = drawingIndex;
        excelLoaded = true;
 
        console.log('[PDA drawing-button] index vytvoreny, pocet zaznamov:', Object.keys(drawingIndex).length);
        updateLoadButtonState('loaded');
 
        if (unsafeWindow.PDA_CURRENT_OPERATION) {
            applyDrawingForCurrentOperation(unsafeWindow.PDA_CURRENT_OPERATION);
        }
    }
 
    async function loadExcelFromHandle(handle) {
        const file = await handle.getFile();
        await loadExcelFromFile(file);
    }
 
    function findDrawingByOrderNo(orderNo) {
        return drawingIndex[orderNo] || null;
    }
    unsafeWindow.PDA_findDrawingByOrderNo = findDrawingByOrderNo;
 
    // ---------- Vyber a znovupouzitie suboru cez File System Access API ----------
 
    const supportsFsAccess = typeof unsafeWindow.showOpenFilePicker === 'function';
 
    async function pickFileAndRemember() {
        try {
            const [handle] = await unsafeWindow.showOpenFilePicker({
                types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] } }],
                excludeAcceptAllOption: false,
                multiple: false,
            });
            await saveHandle(handle);
            updateLoadButtonState('loading');
            await loadExcelFromHandle(handle);
        } catch (err) {
            if (err && err.name === 'AbortError') return; // pouzivatel zavrel dialog
            console.warn('[PDA drawing-button] chyba pri vybere/nacitani suboru', err);
            updateLoadButtonState('error');
        }
    }
 
    async function tryAutoLoadFromStoredHandle() {
        console.log('[PDA drawing-button] tryAutoLoadFromStoredHandle: start, supportsFsAccess =', supportsFsAccess);
 
        if (!supportsFsAccess) {
            updateLoadButtonState('unsupported');
            return;
        }
 
        let handle;
        try {
            handle = await loadHandle();
        } catch (err) {
            console.warn('[PDA drawing-button] chyba pri citani ulozeneho handle-u z IndexedDB', err);
        }
 
        if (!handle) {
            console.log('[PDA drawing-button] ziadny ulozeny handle sa nenasiel, treba vybrat subor manualne');
            updateLoadButtonState('nofile');
            return;
        }
 
        console.log('[PDA drawing-button] najdeny ulozeny handle:', handle, 'nazov suboru:', handle.name);
 
        let permission;
        try {
            permission = await handle.queryPermission({ mode: 'read' });
            console.log('[PDA drawing-button] stav opravnenia (queryPermission) =', permission);
        } catch (err) {
            console.warn('[PDA drawing-button] chyba pri kontrole opravnenia', err);
            updateLoadButtonState('needs-permission', handle);
            return;
        }
 
        if (permission === 'granted') {
            updateLoadButtonState('loading');
            try {
                await loadExcelFromHandle(handle);
            } catch (err) {
                console.warn('[PDA drawing-button] chyba pri nacitani zapamataneho suboru', err);
                updateLoadButtonState('error');
            }
            return;
        }
 
        // 'prompt' -> potrebny je jeden klik na potvrdenie (bez opatovneho prehliadania suborov)
        updateLoadButtonState('needs-permission', handle);
    }
 
    async function confirmPermissionAndLoad(handle) {
        try {
            const permission = await handle.requestPermission({ mode: 'read' });
            if (permission !== 'granted') {
                updateLoadButtonState('needs-permission', handle);
                return;
            }
            updateLoadButtonState('loading');
            await loadExcelFromHandle(handle);
        } catch (err) {
            console.warn('[PDA drawing-button] chyba pri potvrdzovani opravnenia', err);
            updateLoadButtonState('error');
        }
    }
 
    // ---------- XHR intercept: sledovanie aktualne otvorenej operacie ----------
    const TARGET_URL_SUBSTRING = '/client/1.0/executeBO';
 
    unsafeWindow.PDA_CURRENT_OPERATION = unsafeWindow.PDA_CURRENT_OPERATION || null;
 
    const originalOpen = unsafeWindow.XMLHttpRequest.prototype.open;
    const originalSend = unsafeWindow.XMLHttpRequest.prototype.send;
 
    unsafeWindow.XMLHttpRequest.prototype.open = function (method, url, ...rest) {
        this._pdaDrawing_url = url;
        return originalOpen.call(this, method, url, ...rest);
    };
 
    unsafeWindow.XMLHttpRequest.prototype.send = function (body) {
        const url = this._pdaDrawing_url || '';
 
        if (url.includes(TARGET_URL_SUBSTRING)) {
            this.addEventListener('load', function () {
                try {
                    const data = JSON.parse(this.responseText);
                    if (data && data.result && data.result.operation) {
                        handleCurrentOperationResponse(data.result.operation);
                    }
                } catch (e) {
                    // ignorujeme odpovede, ktore nie su JSON alebo nas nezaujimaju
                }
            });
        }
 
        return originalSend.apply(this, arguments);
    };
 
    function handleCurrentOperationResponse(operation) {
        const current = {
            workcenter: operation.workcenterDescription,
            workcenterCode: operation.workcenter,
            productionOrderNo: operation.productionOrderNo,
            operationNo: operation.operationNo,
            sequenceNo: operation.sequenceNo,
            materialNo: operation.materialNo,
            material: operation.material,
        };
 
        unsafeWindow.PDA_CURRENT_OPERATION = current;
        applyDrawingForCurrentOperation(current);
    }
 
    function applyDrawingForCurrentOperation(current) {
        const orderNoForLookup = "'" + (current.productionOrderNo || '').slice(2);
        const drawing = findDrawingByOrderNo(orderNoForLookup);
 
        updateDrawingButtonDisplay(drawing);
    }
 
    function stripLeadingApostrophe(value) {
        if (!value) return value;
        return value.charAt(0) === "'" ? value.slice(1) : value;
    }
 
    let currentDrawingInfo = null; // { drawingNo, version } zobrazene na tlacidle, ocistene od uvodneho apostrofu
 
    function updateDrawingButtonDisplay(drawing) {
        const valueEl = document.getElementById('__pda_order_drawing_value__');
        const revisionEl = document.getElementById('__pda_order_drawing_revision__');
        if (!valueEl || !revisionEl) return;
 
        if (drawing) {
            const cleanDrawingNo = stripLeadingApostrophe(drawing.drawingNo) || '—';
            const cleanVersion = stripLeadingApostrophe(drawing.version) || '';
 
            currentDrawingInfo = { drawingNo: cleanDrawingNo, version: cleanVersion };
 
            valueEl.textContent = cleanDrawingNo;
            revisionEl.textContent = cleanVersion ? 'rev. ' + cleanVersion : '';
        } else {
            currentDrawingInfo = null;
 
            valueEl.textContent = '—';
            revisionEl.textContent = '';
        }
    }
 
    // ---------- Sluzba "Mapa vykresov" (PDM) - hladanie a modalne okno so zoznamom ----------
 
    function pdmHladaj(cislo) {
        return new Promise((resolve, reject) => {
            const params = new URLSearchParams({ q: cislo });
            GM_xmlhttpRequest({
                method: 'GET',
                url: `${PDM_BASE}/search?${params}`,
                headers: PDM_KEY ? { 'X-API-Key': PDM_KEY } : {},
                onload: function (response) {
                    if (response.status < 200 || response.status >= 300) {
                        reject(new Error(`Sluzba vykresov vratila HTTP ${response.status}`));
                        return;
                    }
                    try {
                        resolve(JSON.parse(response.responseText));
                    } catch (e) {
                        reject(new Error('Neplatna odpoved zo sluzby vykresov'));
                    }
                },
                onerror: function () {
                    reject(new Error('Sluzba vykresov neodpovedala (chyba spojenia)'));
                },
                ontimeout: function () {
                    reject(new Error('Sluzba vykresov neodpovedala vcas (timeout)'));
                },
            });
        });
    }
 
    function pdmOtvorOkno(cislo) {
        const overlay = document.createElement('div');
        overlay.id = '__pda_pdm_overlay__';
        overlay.style.cssText =
            'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;' +
            'display:flex;align-items:center;justify-content:center;font-family:inherit;';
 
        overlay.innerHTML = `
            <div style="background:#ffffff;border-radius:10px;max-width:900px;width:92%;max-height:80vh;
                        display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,.3)">
                <div style="display:flex;justify-content:space-between;align-items:center;
                            padding:12px 16px;background:#5b9bd5;color:#ffffff">
                    <b>Výkresy — ${cislo}</b>
                    <button data-pda-zavrit type="button"
                            style="background:none;border:0;color:#ffffff;font-size:22px;line-height:1;cursor:pointer">×</button>
                </div>
                <div data-pda-telo style="padding:14px 16px;overflow:auto">Hľadám…</div>
            </div>`;
 
        const zavri = () => {
            overlay.remove();
            document.removeEventListener('keydown', naEsc);
        };
        const naEsc = (e) => {
            if (e.key === 'Escape') zavri();
        };
 
        overlay.querySelector('[data-pda-zavrit]').addEventListener('click', zavri);
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) zavri();
        });
        document.addEventListener('keydown', naEsc);
        document.body.appendChild(overlay);
 
        const telo = overlay.querySelector('[data-pda-telo]');
 
        pdmHladaj(cislo)
            .then((d) => {
                if (!d.pocet) {
                    telo.innerHTML = `<p>Pre <b>${cislo}</b> sa nenašiel žiadny PDF ani TIFF výkres.</p>`;
                    return;
                }
 
                telo.innerHTML = `
                    <table style="width:100%;border-collapse:collapse;font-size:14px">
                        <thead>
                            <tr>
                                <th style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5">Názov</th>
                                <th style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5">Revízia</th>
                                <th style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5">Stav</th>
                                <th style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5">Ver.</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${d.vysledky
                                .map(
                                    (v) => `
                                <tr data-id="${v.id}" style="cursor:pointer;background:${v.zhoda_revizie === true ? '#e6f4ea' : 'transparent'}">
                                    <td style="padding:6px;border-bottom:1px solid #eeeeee"><b>${v.nazov}</b></td>
                                    <td style="padding:6px;border-bottom:1px solid #eeeeee">${v.revizia || ''}</td>
                                    <td style="padding:6px;border-bottom:1px solid #eeeeee">${v.stav || ''}</td>
                                    <td style="padding:6px;border-bottom:1px solid #eeeeee">${v.verzia != null ? v.verzia : ''}</td>
                                </tr>`
                                )
                                .join('')}
                        </tbody>
                    </table>
                    <p style="color:#777777;font-size:12px;margin-top:10px">
                        ${d.pocet} výkresov · ${d.zdroj === 'mapa' ? 'z dennej mapy' : 'naživo z PDM'}
                        ${d.mapa_z ? ' (' + d.mapa_z + ')' : ''} · kliknutím sa výkres otvorí
                    </p>`;
 
                telo.querySelectorAll('tr[data-id]').forEach((tr) => {
                    tr.addEventListener('click', () => {
                        unsafeWindow.open(`${PDM_BASE}/file/${tr.dataset.id}`, '_blank');
                    });
                });
            })
            .catch((err) => {
                console.warn('[PDA drawing-button] chyba pri hladani vykresov', err);
                telo.innerHTML = `<p style="color:#b00000">Služba výkresov neodpovedala.<br>${err.message}</p>`;
            });
    }
 
    // ---------- Vykreslenie tlacidla + tlacidlo na (opatovny) vyber/potvrdenie suboru ----------
    const CONTAINER_ID = 'WorkcenterDetail--Order_FlexBox';
    const WRAPPER_ID = '__pda_order_drawing_wrapper__';
    const BUTTON_ID = '__pda_order_drawing_button__';
    const LOAD_BUTTON_ID = '__pda_order_drawing_load_button__';
 
    let pendingHandle = null;
    let currentLoadState = 'checking'; // zapamatany stav nezavisly od toho, ci tlacidlo uz existuje v DOM
 
    function buildButton() {
        const button = document.createElement('button');
        button.id = BUTTON_ID;
        button.type = 'button';
 
        button.style.display = 'flex';
        button.style.flexDirection = 'column';
        button.style.alignItems = 'flex-end';
        button.style.gap = '2px';
        button.style.padding = '6px 12px';
        button.style.border = '1px solid #5b9bd5';
        button.style.borderRadius = '6px';
        button.style.backgroundColor = '#ffffff';
        button.style.cursor = 'pointer';
        button.style.width = '150px';
        button.style.flexShrink = '0';
        button.style.marginRight = '8px';
 
        const label = document.createElement('span');
        label.textContent = 'VÝKRES';
        label.style.fontSize = '0.65rem';
        label.style.color = '#888888';
        label.style.letterSpacing = '0.05em';
 
        const value = document.createElement('span');
        value.id = '__pda_order_drawing_value__';
        value.textContent = '—';
        value.style.fontSize = '0.9rem';
        value.style.fontWeight = 'bold';
        value.style.color = '#000000';
 
        const revision = document.createElement('span');
        revision.id = '__pda_order_drawing_revision__';
        revision.textContent = '';
        revision.style.fontSize = '0.75rem';
        revision.style.color = '#555555';
 
        button.appendChild(label);
        button.appendChild(value);
        button.appendChild(revision);
 
        button.addEventListener('click', () => {
            if (currentDrawingInfo && currentDrawingInfo.drawingNo) {
                pdmOtvorOkno(currentDrawingInfo.drawingNo);
            } else {
                console.log('[PDA drawing-button] pre aktualnu zakazku nie je znamy ziadny vykres');
            }
        });
 
        return button;
    }
 
    function buildLoadButton() {
        const loadButton = document.createElement('button');
        loadButton.id = LOAD_BUTTON_ID;
        loadButton.type = 'button';
 
        loadButton.style.padding = '6px 10px';
        loadButton.style.border = '1px solid #cccccc';
        loadButton.style.borderRadius = '6px';
        loadButton.style.backgroundColor = '#f5f5f5';
        loadButton.style.cursor = 'pointer';
        loadButton.style.fontSize = '0.75rem';
        loadButton.style.marginRight = '8px';
        loadButton.style.alignSelf = 'center';
 
        loadButton.addEventListener('click', () => {
            if (pendingHandle) {
                confirmPermissionAndLoad(pendingHandle);
            } else {
                pickFileAndRemember();
            }
        });
 
        return loadButton;
    }
 
    function updateLoadButtonState(state, handle) {
        currentLoadState = state;
        pendingHandle = state === 'needs-permission' ? handle : null;
 
        const loadButton = document.getElementById(LOAD_BUTTON_ID);
        if (!loadButton) {
            console.log('[PDA drawing-button] updateLoadButtonState: tlacidlo este neexistuje v DOM, stav', state, 'si len zapamatavam');
            return;
        }
 
        renderLoadButtonState(loadButton, state);
    }
 
    function renderLoadButtonState(loadButton, state) {
        // po uspesnom nacitani sa tlacidlo uplne skryje - je to poistka, aby nikto omylom neklikal
        // na vyber/zmenu suboru; zmena suboru je zamerne mozna len manualne cez DevTools (vymazanie IndexedDB)
        if (state === 'loaded') {
            loadButton.style.display = 'none';
            return;
        }
        loadButton.style.display = '';
 
        switch (state) {
            case 'loading':
                loadButton.textContent = 'Načítavam…';
                loadButton.style.backgroundColor = '#f5f5f5';
                loadButton.style.borderColor = '#cccccc';
                break;
            case 'needs-permission':
                loadButton.textContent = 'Povoliť prístup k Excelu';
                loadButton.style.backgroundColor = '#fff4e5';
                loadButton.style.borderColor = '#f9a825';
                break;
            case 'error':
                loadButton.textContent = 'Chyba, skús znova';
                loadButton.style.backgroundColor = '#fdecea';
                loadButton.style.borderColor = '#e53935';
                break;
            case 'unsupported':
                loadButton.textContent = 'Prehliadač nepodporuje zapamätanie';
                loadButton.style.backgroundColor = '#fdecea';
                loadButton.style.borderColor = '#e53935';
                break;
            case 'nofile':
            default:
                loadButton.textContent = 'Vybrať Excel';
                loadButton.style.backgroundColor = '#f5f5f5';
                loadButton.style.borderColor = '#cccccc';
                break;
        }
    }
 
    function buildWrapper() {
        const wrapper = document.createElement('div');
        wrapper.id = WRAPPER_ID;
        wrapper.style.display = 'flex';
        wrapper.style.justifyContent = 'flex-end';
        wrapper.style.alignItems = 'center';
        wrapper.style.width = '100%';
        wrapper.style.boxSizing = 'border-box';
        wrapper.style.marginBottom = '8px';
 
        wrapper.appendChild(buildLoadButton());
        wrapper.appendChild(buildButton());
        return wrapper;
    }
 
    function ensureButton() {
        const container = document.getElementById(CONTAINER_ID);
        if (!container) return;
        if (document.getElementById(WRAPPER_ID)) return;
 
        const wrapper = buildWrapper();
        container.insertBefore(wrapper, container.firstChild);
 
        const loadButton = document.getElementById(LOAD_BUTTON_ID);
        if (loadButton) {
            renderLoadButtonState(loadButton, currentLoadState);
        }
 
        if (unsafeWindow.PDA_CURRENT_OPERATION) {
            applyDrawingForCurrentOperation(unsafeWindow.PDA_CURRENT_OPERATION);
        }
    }
 
    let debounceTimer = null;
    function scheduleEnsure() {
        if (debounceTimer) return;
        debounceTimer = setTimeout(() => {
            debounceTimer = null;
            ensureButton();
        }, 300);
    }
 
    scheduleEnsure();
    const observer = new MutationObserver(() => {
        if (!document.getElementById(WRAPPER_ID)) scheduleEnsure();
    });
    observer.observe(document.body, { childList: true, subtree: true });
 
    tryAutoLoadFromStoredHandle();
})();