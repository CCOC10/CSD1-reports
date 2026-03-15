/* csd1-cache.js — Shared SWR data cache (sessionStorage) */
(function () {
  'use strict';
  if (window.CSD1Cache) return;
  var PFX = 'CSD1_SWR::';

  window.CSD1Cache = {
    get: function (key) {
      try {
        var raw = sessionStorage.getItem(PFX + key);
        return raw ? JSON.parse(raw) : null;
      } catch (_) { return null; }
    },
    set: function (key, data) {
      try {
        sessionStorage.setItem(PFX + key, JSON.stringify({ data: data, ts: Date.now() }));
      } catch (_) {}
    },
    del: function (key) {
      try { sessionStorage.removeItem(PFX + key); } catch (_) {}
    },
    clear: function (prefix) {
      try {
        Object.keys(sessionStorage).forEach(function (k) {
          if (k.startsWith(PFX + (prefix || ''))) sessionStorage.removeItem(k);
        });
      } catch (_) {}
    }
  };
})();
