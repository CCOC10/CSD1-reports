(function () {
  'use strict';

  var VIEWPORT_CONTENT = 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover';

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

  function lockZoomGestures() {
    var block = function (e) {
      e.preventDefault();
    };

    ['gesturestart', 'gesturechange', 'gestureend'].forEach(function (type) {
      document.addEventListener(type, block, { passive: false });
    });

    document.addEventListener(
      'wheel',
      function (e) {
        if (e.ctrlKey) e.preventDefault();
      },
      { passive: false }
    );

    document.addEventListener(
      'touchmove',
      function (e) {
        if (e.touches && e.touches.length > 1) e.preventDefault();
      },
      { passive: false }
    );

    var lastTouchEnd = 0;
    document.addEventListener(
      'touchend',
      function (e) {
        var now = Date.now();
        if (now - lastTouchEnd <= 300) e.preventDefault();
        lastTouchEnd = now;
      },
      { passive: false }
    );

    document.documentElement.style.touchAction = 'manipulation';
    document.documentElement.style.overscrollBehavior = 'none';
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
  lockZoomGestures();
  bindDisplayModeChange();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyStandaloneClass, { once: true });
  } else {
    applyStandaloneClass();
  }

  window.addEventListener('resize', applyStandaloneClass, { passive: true });
})();
