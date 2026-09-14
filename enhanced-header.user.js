// ==UserScript==
// @name         PDA - Enhanced Header
// @namespace    http://tampermonkey.net/
// @version      1.0.4
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/enhanced-header.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/enhanced-header.user.js
// @description  Bigger name, added text to buton and highlighted logout button
// @author       Gabris
// @match        https://hf.simplifier.cloud/appDirect/PDA/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const USERNAME_SUFFIX = 'Label_Username-bdi';
  const USERNAME_FONT_SIZE = '1.8rem';
  const USERNAME_FONT_WEIGHT = 'bold';
  const USERNAME_FONT_COLOR = '#1f1f1f';
  const BUTTON_TEXT_MARGIN_RIGHT = '0.6rem';

  const BUTTONS = [
    { suffix: 'Button_HomeScreen-img', label: 'Pracoviská' },
    { suffix: 'Button_Reporting-img', label: 'Reporty' },
    { suffix: 'Button_Message-img', label: 'Správy' },
    { suffix: 'Button_Schedule-img', label: 'Rozvrh' },
    { suffix: 'Button_Settings-img', label: 'Admin' },
    { suffix: 'Button_ChangeUser-img', label: 'Zmena používateľa' },
    { suffix: 'Button_Logout-img', label: 'Odhlásenie', bgColor: "" },
  ];

  const LOGOUT_BUTTON_SUFFIX = 'Button_Logout-inner';
  const LOGOUT_BG_COLOR = '#d67a74';

  function applyUsernameStyle() {
    document.querySelectorAll(`[id$="${USERNAME_SUFFIX}"]`).forEach((el) => {
      if (
        el.style.fontSize === USERNAME_FONT_SIZE &&
        el.style.fontWeight === USERNAME_FONT_WEIGHT &&
        el.style.color === 'rgb(0, 0, 0)'
      ) return;
      el.style.fontSize = USERNAME_FONT_SIZE;
      el.style.fontWeight = USERNAME_FONT_WEIGHT;
      el.style.color = USERNAME_FONT_COLOR;
    });
  }

  function applyTextOnBtns() {
    BUTTONS.forEach(({ suffix, label }) => {
      document.querySelectorAll(`[id$="${suffix}"]`).forEach((img) => {
        if (img.dataset.pdaTextAdded) return;

        const textSpan = document.createElement('span');
        textSpan.className = 'sapMBtnContent';
        textSpan.style.marginRight = BUTTON_TEXT_MARGIN_RIGHT;
        textSpan.innerHTML = `<bdi>${label}</bdi>`;

        img.insertAdjacentElement('afterend', textSpan);
        img.dataset.pdaTextAdded = '1';
      });
    });
  }

  function applyLogoutBtnStyle() {
    document.querySelectorAll(`[id$="${LOGOUT_BUTTON_SUFFIX}"]`).forEach((btn) => {
      if (btn.dataset.pdaLogoutStyled) return;

      btn.style.backgroundColor = LOGOUT_BG_COLOR;
      btn.style.borderRadius = '4px';

      btn.dataset.pdaLogoutStyled = '1';
    });
  }

  function applyChanges() {
      applyUsernameStyle();
      applyTextOnBtns();
      applyLogoutBtnStyle();
  }

  applyChanges();

  const observer = new MutationObserver(applyChanges);
  observer.observe(document.body, { childList: true, subtree: true, characterData: true });
})();