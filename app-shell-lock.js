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

  function isEditableField(el) {
    if (!el || typeof el !== 'object') return false;
    var tag = (el.tagName || '').toLowerCase();
    if (!tag) return false;
    if (el.disabled || el.readOnly) return false;
    if (tag === 'textarea') return true;
    if (tag === 'input') {
      var type = (el.type || 'text').toLowerCase();
      return ['text', 'search', 'email', 'number', 'password', 'tel', 'url', 'date', 'datetime-local', 'time', 'month', 'week'].indexOf(type) !== -1;
    }
    return !!el.isContentEditable;
  }

  function ensureKeyboardNavStyle() {
    if (document.getElementById('keyboardNavHideStyle')) return;
    var style = document.createElement('style');
    style.id = 'keyboardNavHideStyle';
    style.textContent = [
      'body.keyboard-open .mobile-mini-nav {',
      '  display: none !important;',
      '  opacity: 0 !important;',
      '  pointer-events: none !important;',
      '  transform: translateY(120%) !important;',
      '}',
      'body.keyboard-open .mobile-nav-fade {',
      '  opacity: 0 !important;',
      '  pointer-events: none !important;',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  function bindKeyboardAwareMobileNav() {
    if (!isTouchDevice()) return;
    ensureKeyboardNavStyle();

    var keyboardOpen = false;
    var baseHeight = window.innerHeight || 0;

    function setKeyboardOpen(open) {
      var value = !!open;
      if (value === keyboardOpen) return;
      keyboardOpen = value;
      if (document.body) document.body.classList.toggle('keyboard-open', value);
    }

    function refreshBaseHeight() {
      var current = window.innerHeight || 0;
      if (current > baseHeight) baseHeight = current;
    }

    function refreshKeyboardState() {
      refreshBaseHeight();
      var activeEditable = isEditableField(document.activeElement);
      var viewportOpen = false;
      if (window.visualViewport && activeEditable) {
        var vvHeight = window.visualViewport.height || 0;
        var diff = (baseHeight || window.innerHeight || 0) - vvHeight;
        viewportOpen = diff > 120;
      }
      setKeyboardOpen(activeEditable || viewportOpen);
    }

    document.addEventListener('focusin', function () {
      refreshKeyboardState();
    }, true);
    document.addEventListener('focusout', function () {
      setTimeout(refreshKeyboardState, 40);
    }, true);
    window.addEventListener('orientationchange', function () {
      setTimeout(function () {
        baseHeight = window.innerHeight || baseHeight;
        refreshKeyboardState();
      }, 120);
    }, { passive: true });
    if (window.visualViewport && typeof window.visualViewport.addEventListener === 'function') {
      window.visualViewport.addEventListener('resize', refreshKeyboardState, { passive: true });
    }

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', refreshKeyboardState, { once: true });
    } else {
      refreshKeyboardState();
    }
  }

  enforceMobileMeta();
  lockZoomOnly();
  bindDisplayModeChange();
  bindKeyboardAwareMobileNav();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyStandaloneClass, { once: true });
  } else {
    applyStandaloneClass();
  }

  window.addEventListener('resize', applyStandaloneClass, { passive: true });
})();
