// ==UserScript==
// @name         Test
// @namespace    http://tampermonkey.net/
// @version      0.0.1
// @description  Zachyti vsetky volania na /client/1.0/executeBO a vypise request+response do konzoly
// @author       Gabris
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/test.user.js
// @updateURL    https://github.com/JaroTvarozek/HF-PDA-scripts/raw/refs/heads/main/test.user.js
// @match        https://hf.simplifier.cloud/appDirect/PDA/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const TARGET_URL_SUBSTRING = '/client/1.0/executeBO';

  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSend = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    this._pda_url = url;
    this._pda_method = method;
    return originalOpen.call(this, method, url, ...rest);
  };

  XMLHttpRequest.prototype.send = function (body) {
    if (this._pda_url && this._pda_url.includes(TARGET_URL_SUBSTRING)) {
      const requestBody = body;
      this.addEventListener('load', function () {
        console.log('--- [PDA executeBO] ---');
        console.log('request body:', requestBody);
        console.log('response:', this.responseText);
        console.log('------------------------');
      });
    }
    return originalSend.apply(this, arguments);
  };
})();