// ==UserScript==
// @name         PDA - Prevent back button
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  Prevent browser related back button click
// @author       Gabris
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/prevent-back-button.user.js
// @downloadURL  https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/prevent-back-button.user.js
// @match        https://hf.simplifier.cloud/appDirect/PDA/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const HISTORY_LENGTH_AT_START = history.length;

  function lockCurrentState() {
    history.pushState(history.state, document.title, location.href);
  }

  lockCurrentState();
  window.addEventListener('popstate', function () {
    lockCurrentState();
  });

  if (history.length <= HISTORY_LENGTH_AT_START + 1) {
    window.addEventListener('beforeunload', function (e) {
      e.preventDefault();
      e.returnValue = '';
    });
  }
})();