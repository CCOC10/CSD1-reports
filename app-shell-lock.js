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

  function ensurePageLoadingHideStyle() {
    if (document.getElementById('pageLoadingNavHideStyle')) return;
    var style = document.createElement('style');
    style.id = 'pageLoadingNavHideStyle';
    style.textContent = [
      'html.shell-page-loading .mobile-mini-nav,',
      'body.shell-page-loading .mobile-mini-nav,',
      'html.mobile-loading-data .mobile-mini-nav,',
      'body.mobile-loading-data .mobile-mini-nav {',
      '  display: none !important;',
      '  opacity: 0 !important;',
      '  visibility: hidden !important;',
      '  pointer-events: none !important;',
      '}',
      'html.shell-page-loading button#fabBtn.fab-btn,',
      'body.shell-page-loading button#fabBtn.fab-btn,',
      'html.mobile-loading-data button#fabBtn.fab-btn,',
      'body.mobile-loading-data button#fabBtn.fab-btn {',
      '  display: none !important;',
      '  opacity: 0 !important;',
      '  visibility: hidden !important;',
      '  pointer-events: none !important;',
      '}',
      'html.shell-page-loading .floating-profile-fab,',
      'body.shell-page-loading .floating-profile-fab,',
      'html.mobile-loading-data .floating-profile-fab,',
      'body.mobile-loading-data .floating-profile-fab {',
      '  display: none !important;',
      '  opacity: 0 !important;',
      '  visibility: hidden !important;',
      '  pointer-events: none !important;',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  function forceHideShellFloatingNav(active) {
    var shouldHide = !!active;
    var targets = document.querySelectorAll('.mobile-mini-nav, button#fabBtn.fab-btn, .floating-profile-fab');
    for (var i = 0; i < targets.length; i++) {
      var el = targets[i];
      if (!el || !el.style) continue;
      if (shouldHide) {
        if (el.getAttribute('data-shell-forced-hidden') !== '1') {
          el.setAttribute('data-shell-prev-display', el.style.getPropertyValue('display') || '');
          el.setAttribute('data-shell-prev-display-priority', el.style.getPropertyPriority('display') || '');
          el.setAttribute('data-shell-prev-opacity', el.style.getPropertyValue('opacity') || '');
          el.setAttribute('data-shell-prev-opacity-priority', el.style.getPropertyPriority('opacity') || '');
          el.setAttribute('data-shell-prev-visibility', el.style.getPropertyValue('visibility') || '');
          el.setAttribute('data-shell-prev-visibility-priority', el.style.getPropertyPriority('visibility') || '');
          el.setAttribute('data-shell-prev-pointer-events', el.style.getPropertyValue('pointer-events') || '');
          el.setAttribute('data-shell-prev-pointer-events-priority', el.style.getPropertyPriority('pointer-events') || '');
        }
        el.style.setProperty('display', 'none', 'important');
        el.style.setProperty('opacity', '0', 'important');
        el.style.setProperty('visibility', 'hidden', 'important');
        el.style.setProperty('pointer-events', 'none', 'important');
        el.setAttribute('data-shell-forced-hidden', '1');
      } else if (el.getAttribute('data-shell-forced-hidden') === '1') {
        var prevDisplay = el.getAttribute('data-shell-prev-display') || '';
        var prevDisplayPriority = el.getAttribute('data-shell-prev-display-priority') || '';
        var prevOpacity = el.getAttribute('data-shell-prev-opacity') || '';
        var prevOpacityPriority = el.getAttribute('data-shell-prev-opacity-priority') || '';
        var prevVisibility = el.getAttribute('data-shell-prev-visibility') || '';
        var prevVisibilityPriority = el.getAttribute('data-shell-prev-visibility-priority') || '';
        var prevPointerEvents = el.getAttribute('data-shell-prev-pointer-events') || '';
        var prevPointerEventsPriority = el.getAttribute('data-shell-prev-pointer-events-priority') || '';

        if (prevDisplay) el.style.setProperty('display', prevDisplay, prevDisplayPriority);
        else el.style.removeProperty('display');
        if (prevOpacity) el.style.setProperty('opacity', prevOpacity, prevOpacityPriority);
        else el.style.removeProperty('opacity');
        if (prevVisibility) el.style.setProperty('visibility', prevVisibility, prevVisibilityPriority);
        else el.style.removeProperty('visibility');
        if (prevPointerEvents) el.style.setProperty('pointer-events', prevPointerEvents, prevPointerEventsPriority);
        else el.style.removeProperty('pointer-events');

        el.removeAttribute('data-shell-prev-display');
        el.removeAttribute('data-shell-prev-display-priority');
        el.removeAttribute('data-shell-prev-opacity');
        el.removeAttribute('data-shell-prev-opacity-priority');
        el.removeAttribute('data-shell-prev-visibility');
        el.removeAttribute('data-shell-prev-visibility-priority');
        el.removeAttribute('data-shell-prev-pointer-events');
        el.removeAttribute('data-shell-prev-pointer-events-priority');
        el.removeAttribute('data-shell-forced-hidden');
      }
    }
  }

  function setShellPageLoading(active) {
    var value = !!active;
    document.documentElement.classList.toggle('shell-page-loading', value);
    if (document.body) {
      document.body.classList.toggle('shell-page-loading', value);
    }
    forceHideShellFloatingNav(value);
  }

  function bindPageLoadingHideState() {
    ensurePageLoadingHideStyle();
    setShellPageLoading(document.readyState !== 'complete');

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', function () {
        setShellPageLoading(true);
      }, { once: true });
    }

    window.addEventListener('load', function () {
      setShellPageLoading(false);
    }, { once: true });

    window.addEventListener('pageshow', function () {
      setShellPageLoading(false);
    }, { once: true });
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

  function ensureFloatingLayerStyle() {
    if (document.getElementById('floatingLayerStyle')) return;
    var style = document.createElement('style');
    style.id = 'floatingLayerStyle';
    style.textContent = [
      '.toast,',
      '#rankToast.toast {',
      '  z-index: 12020 !important;',
      '}',
      '.fab-menu,',
      '.fab-menu .fab-menu-items {',
      '  z-index: 12010 !important;',
      '}',
      '.fab-menu .fab-menu-items {',
      '  background: rgba(2, 6, 23, 0.72) !important;',
      '  border-color: rgba(255, 212, 0, 0.32) !important;',
      '  box-shadow: 0 14px 36px rgba(2, 6, 23, 0.48) !important;',
      '  backdrop-filter: blur(16px) saturate(145%) !important;',
      '  -webkit-backdrop-filter: blur(16px) saturate(145%) !important;',
      '}',
      '[data-theme="light"] .fab-menu .fab-menu-items {',
      '  background: rgba(255, 255, 255, 0.76) !important;',
      '  border-color: rgba(0, 38, 75, 0.28) !important;',
      '  box-shadow: 0 14px 30px rgba(0, 38, 75, 0.2) !important;',
      '}',
      '@media (max-width: 900px) {',
      '  .toast,',
      '  #rankToast.toast {',
      '    bottom: calc(env(safe-area-inset-bottom) + 5.35rem) !important;',
      '    max-width: min(92vw, 560px);',
      '    background: rgba(2, 6, 23, 0.72) !important;',
      '    border-color: rgba(255, 212, 0, 0.34) !important;',
      '    box-shadow: 0 14px 34px rgba(2, 6, 23, 0.46) !important;',
      '    backdrop-filter: blur(14px) saturate(140%) !important;',
      '    -webkit-backdrop-filter: blur(14px) saturate(140%) !important;',
      '  }',
      '  [data-theme="light"] .toast,',
      '  [data-theme="light"] #rankToast.toast {',
      '    background: rgba(255, 255, 255, 0.74) !important;',
      '    border-color: rgba(0, 38, 75, 0.3) !important;',
      '    color: #00264B !important;',
      '    box-shadow: 0 14px 30px rgba(0, 38, 75, 0.2) !important;',
      '  }',
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
  bindPageLoadingHideState();
  ensureFloatingLayerStyle();
  bindKeyboardAwareMobileNav();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyStandaloneClass, { once: true });
  } else {
    applyStandaloneClass();
  }

  window.addEventListener('resize', applyStandaloneClass, { passive: true });
})();
