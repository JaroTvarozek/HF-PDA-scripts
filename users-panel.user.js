// ==UserScript==
// @name         PDA - Users panel
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  Side panel for fast login
// @author       Gabris
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/users-panel.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/users-panel.user.js
// @match        https://hf.simplifier.cloud/appDirect/PDA/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const PARENT_ID = 'Main';
  const CONTENT_ID = 'Main--MainPage';

  const WRAPPER_ID = '__pda_main_wrapper__';
  const PANEL_ID = '__pda_user_switch_panel__';

  const PANEL_WIDTH = '280px';
  const PANEL_BG_COLOR = '#313175';
  const TEXT_COLOR = '#e9e9f0';
  const BUTTON_BORDER_RADIUS = '8px';
  const BUTTON_BG_COLOR = '#222252';

  const CHANGE_USER_SUFFIX = 'Button_ChangeUser';

  const USER_COMBO_ID = 'Popups--User_ComboBox';
  const PASSWORD_INPUT_ID = 'Popups--EmployeeNo_Input';
  const CONFIRM_BUTTON_ID = 'Popups--UserDialog_Button_close';

  // --- Zoznam pouzivatelov na rychle prepinanie ---
  // POZOR: password a cardId su zamerne prazdne - tento subor je vo verejnom
  // repozitari. Realne hodnoty doplnaj az lokalne v Tampermonkey
  // (Dashboard -> tento skript -> Editor). Necommituj ich.
  const USERS = [
    { username: 'Daniel Gabris', password: '', cardId: '' },
    { username: 'Peter Gabor', password: '', cardId: '' },
    { username: 'Lubomir Zilka', password: '', cardId: '' },
    { username: 'Robert Ziacek', password: '', cardId: '' },
    { username: 'Vladislav Dlhy', password: '', cardId: '' },
  ];

  const DIALOG_WAIT_TIMEOUT = 10000;
  const BEFORE_CONFIRM_DELAY = 100;

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  const STYLE_TAG_ID = '__pda_panel_styles__';

  function injectStyles() {
        if (document.getElementById(STYLE_TAG_ID)) return;
        const style = document.createElement('style');
        style.id = STYLE_TAG_ID;
        style.textContent = `
    .pda-user-btn { border: 2px solid transparent; }
    .pda-user-btn:hover { border: 2px solid #e9e9f0; }
  `;
    document.head.appendChild(style);
  }

  function waitFor(checkFn, { interval = 150, timeout = 10000 } = {}) {
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

  function getControl(controlId) {
    try {
      if (window.sap && sap.ui && sap.ui.getCore) {
        return sap.ui.getCore().byId(controlId);
      }
    } catch (e) { /* ignore */ }
    return null;
  }

  async function selectComboBoxItemByText(controlId, text, { timeout = 5000 } = {}) {
    let control;
    try {
      control = await waitFor(() => {
        const c = getControl(controlId);
        return c && typeof c.getItems === 'function' && c.getItems().length > 0 ? c : null;
      }, { timeout });
    } catch (e) {
      console.warn('[PDA panel] ComboBox items sa nenacitali vcas pre', controlId, e);
      return false;
    }

    const items = control.getItems();
    const match =
      items.find((it) => it.getText() === text) ||
      items.find((it) => it.getText().toLowerCase() === text.toLowerCase());

    if (!match) {
      console.warn(
        '[PDA panel] Nenasla sa polozka v ComboBoxe pre text:',
        text,
        'dostupne polozky:',
        items.map((i) => i.getText())
      );
      return false;
    }

    control.setSelectedItem(match);
    if (typeof control.fireSelectionChange === 'function') {
      control.fireSelectionChange({ selectedItem: match });
    }
    if (typeof control.fireChange === 'function') {
      control.fireChange({ value: match.getText() });
    }
    return true;
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
      console.warn('[PDA panel] setControlValue zlyhalo pre', controlId, e);
    }
    const inputEl = document.getElementById(controlId + '-inner');
    if (inputEl) {
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeSetter.call(inputEl, value);
      inputEl.dispatchEvent(new Event('input', { bubbles: true }));
      inputEl.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }
    return false;
  }

  async function fillLoginPopup(username, password) {
    try {
      await waitFor(
        () => document.getElementById(USER_COMBO_ID + '-inner') && document.getElementById(PASSWORD_INPUT_ID + '-inner'),
        { timeout: DIALOG_WAIT_TIMEOUT }
      );

      const selected = await selectComboBoxItemByText(USER_COMBO_ID, username);
      if (!selected) {
        console.warn('[PDA panel] pouzivam fallback setValue - vyber v zozname sa nepodaril');
        setControlValue(USER_COMBO_ID, username);
      }

      setControlValue(PASSWORD_INPUT_ID, password);

      await sleep(BEFORE_CONFIRM_DELAY);

      const confirmBtn = await waitFor(() => document.getElementById(CONFIRM_BUTTON_ID), { timeout: 5000 });
      pressElement(confirmBtn);
    } catch (e) {
      console.warn('[PDA panel] popup na zmenu pouzivatela sa neobjavil vcas alebo zlyhalo potvrdenie', e);
    }
  }

  function switchToUser(username, password) {
    const btn = document.querySelector(`[id$="${CHANGE_USER_SUFFIX}"]`);
    if (!btn) {
      console.warn('[PDA panel] Button_ChangeUser sa nenasiel na obrazovke');
      return;
    }
    pressElement(btn);
    fillLoginPopup(username, password);
  }

  function buildUserButtons(panel) {
  USERS.forEach(({ username, password }) => {
    const btn = document.createElement('button');
    btn.textContent = username;
    btn.classList.add('pda-user-btn');
    btn.style.display = 'block';
    btn.style.width = '100%';
    btn.style.padding = '0.5rem 0.8rem';
    btn.style.marginBottom = '0.4rem';
    btn.style.cursor = 'pointer';
    btn.style.textAlign = 'left';
    btn.style.borderRadius = BUTTON_BORDER_RADIUS;
    btn.style.color = TEXT_COLOR;
    btn.style.backgroundColor = BUTTON_BG_COLOR;
    btn.addEventListener('click', () => switchToUser(username, password));
    panel.appendChild(btn);
    });
  }

  function ensureLayout() {
    let wrapper = document.getElementById(WRAPPER_ID);
    let panel = document.getElementById(PANEL_ID);

    const parent = document.getElementById(PARENT_ID);
    const content = document.getElementById(CONTENT_ID);
    if (!parent || !content) return null;

    if (wrapper && wrapper.contains(content) && panel) {
      return { wrapper, panel };
    }

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
      panel.style.flex = `0 0 ${PANEL_WIDTH}`;
      panel.style.maxWidth = PANEL_WIDTH;
      panel.style.marginRight = '4px';
      panel.style.background = PANEL_BG_COLOR;
      panel.style.overflowY = 'auto';
      panel.style.padding = '0.75rem';
      panel.style.boxSizing = 'border-box';

      const title = document.createElement('div');
      title.textContent = 'Rýchla zmena používateľa';
      title.style.fontWeight = 'bold';
      title.style.marginBottom = '0.5rem';
      title.style.color = TEXT_COLOR;
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

    return { wrapper, panel };
  }

  ensureLayout();
  injectStyles();

  // Global functions and variables
  window.PDA_USERS = USERS;
  window.PDA_switchToUser = switchToUser;
  //------------------------------

  const observer = new MutationObserver(ensureLayout);
  observer.observe(document.body, { childList: true, subtree: true });
})();