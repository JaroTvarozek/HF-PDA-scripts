// ==UserScript==
// @name         PDA - Color and reorder buttons
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  Color and reorder status buttons
// @author       Gabris
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/color-and-reorder-buttons.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/color-and-reorder-buttons.user.js
// @match        https://hf.simplifier.cloud/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const CONTAINER_ID = 'WorkcenterDetail--Order_Status_Flexbox';

  const COLORS = {
    productive: '#4e9041',
    downtime: '#d66c37',
    fault: '#d04040',
  };

  const BORDER_RADIUS = '10px';

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
    [btn, inner].forEach((el) => el && el.style.setProperty('border-radius', BORDER_RADIUS, 'important'));

    const category = STATUS_MAP[textOf(btn)];
    if (!category) return;
    const color = COLORS[category];
    [btn, inner].forEach((el) => {
      if (!el) return;
      el.style.setProperty('background-color', color, 'important');
      el.style.setProperty('border-color', color, 'important');
    });
  }

  function getButtons() {
    const container = document.getElementById(CONTAINER_ID);
    if (!container) return { container: null, buttons: [] };
    const buttons = Array.from(container.children).filter((el) => el.classList.contains('statusBtn'));
    return { container, buttons };
  }

  function reorderButtons(container, buttons) {
    if (!container || buttons.length === 0) return;

    const priorityOf = (btn) => PRIORITY[STATUS_MAP[textOf(btn)]] ?? 1.5;
    const sorted = [...buttons].sort((a, b) => priorityOf(a) - priorityOf(b));

    const alreadyOrdered = sorted.every((btn, i) => buttons[i] === btn);
    if (alreadyOrdered) return;

    sorted.forEach((btn) => container.appendChild(btn));
  }

  function updateAll() {
    const { container, buttons } = getButtons();
    buttons.forEach(styleButton);
    reorderButtons(container, buttons);
  }

  updateAll();

  const observer = new MutationObserver(() => updateAll());
  observer.observe(document.body, { childList: true, subtree: true });
})();