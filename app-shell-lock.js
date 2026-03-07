(function () {
  'use strict';

  var VIEWPORT_CONTENT = 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover';

  function isTouchDevice() {
    try {
      if (window.matchMedia && window.matchMedia('(pointer: coarse)').matches) return true;
    } catch (_) {}
    return !!('ontouchstart' in window || (window.navigator && window.navigator.maxTouchPoints > 0));
  }

  function isIOS() {
    var ua = window.navigator && window.navigator.userAgent ? window.navigator.userAgent : '';
    return /iPad|iPhone|iPod/.test(ua) || (ua.indexOf('Macintosh') !== -1 && window.navigator.maxTouchPoints > 1);
  }

  function ensureMeta(name, content) {
    var node = document.querySelector('meta[name="' + name + '"]');
    if (!node) {
      node = document.createElement('meta');
      node.setAttribute('name', name);
      document.head.appendChild(node);
    }
    if (content != null) node.setAttribute('content', content);
    return node;
  }

  function enforceMobileMeta() {
    if (!isTouchDevice()) {
      ensureMeta('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
      return;
    }
    ensureMeta('viewport', VIEWPORT_CONTENT);
    ensureMeta('apple-mobile-web-app-capable', 'yes');
    ensureMeta('apple-mobile-web-app-status-bar-style', 'black-translucent');
    ensureMeta('mobile-web-app-capable', 'yes');
  }

  function isStandaloneMode() {
    var standalone = false;
    try {
      standalone =
        !!(window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) ||
        !!(window.matchMedia && window.matchMedia('(display-mode: fullscreen)').matches) ||
        window.navigator.standalone === true;
    } catch (_) {
      standalone = window.navigator.standalone === true;
    }
    return !!standalone;
  }

  function applyStandaloneClass() {
    var standalone = isStandaloneMode();
    document.documentElement.classList.toggle('is-standalone-app', standalone);
    if (document.body) document.body.classList.toggle('is-standalone-app', standalone);
  }

  function lockZoomOnly() {
    if (!isTouchDevice()) return;

    // iOS Safari pinch-zoom gestures; does not block normal page scroll.
    if (isIOS()) {
      ['gesturestart', 'gesturechange', 'gestureend'].forEach(function (type) {
        document.addEventListener(type, function (e) {
          e.preventDefault();
        }, { passive: false });
      });
    }
  }

  function bindDisplayModeChange() {
    try {
      var mm = window.matchMedia('(display-mode: standalone)');
      if (!mm) return;
      if (typeof mm.addEventListener === 'function') {
        mm.addEventListener('change', applyStandaloneClass);
      } else if (typeof mm.addListener === 'function') {
        mm.addListener(applyStandaloneClass);
      }
    } catch (_) {}
  }

  enforceMobileMeta();
  lockZoomOnly();
  bindDisplayModeChange();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyStandaloneClass, { once: true });
  } else {
    applyStandaloneClass();
  }

  window.addEventListener('resize', applyStandaloneClass, { passive: true });
})();
