/* page-transition.js — smooth in-app navigation with a top loading bar.
   Load this as the FIRST <script> in <head> (no defer/async). */
(function () {
  'use strict';
  if (window.__CSD1_PT__) return;
  window.__CSD1_PT__ = true;

  var navTimer = null;
  var progressBar = null;
  var progressFrame = null;
  var prefetched = Object.create(null);

  var s = document.createElement('style');
  s.textContent =
    'body{animation:csd1PgIn 0.16s ease;}' +
    '@keyframes csd1PgIn{from{opacity:0.96;transform:translateY(3px)}to{opacity:1;transform:none}}' +
    '.csd1-pg-loading{pointer-events:none!important;}' +
    '.csd1-progress{position:fixed;top:0;left:0;width:100%;height:3px;z-index:9999;pointer-events:none;opacity:0;transition:opacity 0.14s ease;}' +
    '.csd1-progress.show{opacity:1;}' +
    '.csd1-progress__bar{width:100%;height:100%;transform-origin:left center;transform:scaleX(0);background:linear-gradient(90deg,#FFD400 0%,#fff2a8 55%,#FFD400 100%);box-shadow:0 0 12px rgba(255,212,0,0.45);}';
  (document.head || document.documentElement).appendChild(s);

  function ensureProgressBar() {
    if (progressBar) return progressBar;
    progressBar = document.createElement('div');
    progressBar.className = 'csd1-progress';
    progressBar.setAttribute('aria-hidden', 'true');

    var inner = document.createElement('div');
    inner.className = 'csd1-progress__bar';
    progressBar.appendChild(inner);

    (document.body || document.documentElement).appendChild(progressBar);
    return progressBar;
  }

  function stopProgressAnimation() {
    if (progressFrame) {
      cancelAnimationFrame(progressFrame);
      progressFrame = null;
    }
  }

  function setProgress(value) {
    var bar = ensureProgressBar().firstChild;
    bar.style.transform = 'scaleX(' + Math.max(0, Math.min(1, value)) + ')';
  }

  function animateProgress() {
    var start = null;

    stopProgressAnimation();
    ensureProgressBar().classList.add('show');
    setProgress(0.08);

    function step(timestamp) {
      if (start == null) start = timestamp;
      var elapsed = timestamp - start;
      var progress = 0.08 + (Math.min(elapsed, 220) / 220) * 0.62;
      setProgress(progress);
      if (elapsed < 220) {
        progressFrame = requestAnimationFrame(step);
        return;
      }
      progressFrame = null;
    }

    progressFrame = requestAnimationFrame(step);
  }

  function finishProgress() {
    stopProgressAnimation();
    if (!progressBar) return;
    setProgress(0.94);
  }

  function getPrefetchTargetFromElement(element) {
    var current = element;
    while (current && current !== document.documentElement) {
      if (current.tagName === 'A' && current.getAttribute('href')) {
        return current.getAttribute('href');
      }

      if (current.hasAttribute && current.hasAttribute('data-prefetch')) {
        return current.getAttribute('data-prefetch');
      }

      if (current.hasAttribute && current.hasAttribute('onclick')) {
        var handler = current.getAttribute('onclick') || '';
        var match = handler.match(/navigateInApp\s*\(\s*(?:event|null)?\s*,\s*['\"]([^'\"]+)['\"]/i);
        if (match && match[1]) {
          return match[1];
        }
      }

      current = current.parentElement;
    }
    return '';
  }

  function resolveUrl(target) {
    try {
      return new URL(String(target || ''), window.location.href);
    } catch (_) {
      return null;
    }
  }

  function isPrefetchable(url) {
    return !!url && url.origin === window.location.origin && url.pathname !== window.location.pathname;
  }

  function prefetch(target) {
    var url = resolveUrl(target);
    if (!isPrefetchable(url)) return;
    var cacheKey = url.href;
    if (prefetched[cacheKey]) return;
    prefetched[cacheKey] = true;

    try {
      fetch(url.href, {
        method: 'GET',
        credentials: 'same-origin',
        mode: 'cors'
      }).catch(function () {});
    } catch (_) {}
  }

  function handlePrefetchEvent(event) {
    prefetch(getPrefetchTargetFromElement(event.target));
  }

  function setupPrefetch() {
    document.addEventListener('pointerover', handlePrefetchEvent, { passive: true });
    document.addEventListener('touchstart', handlePrefetchEvent, { passive: true });
    document.addEventListener('focusin', handlePrefetchEvent, { passive: true });
  }

  function wrapNav() {
    var orig = window.navigateInApp;
    if (typeof orig !== 'function' || orig._csd1pt) return;

    window.navigateInApp = function (evt, path) {
      if (evt) evt.preventDefault();
      prefetch(path);
      if (document.body) document.body.classList.add('csd1-pg-loading');
      if (navTimer) clearTimeout(navTimer);
      animateProgress();
      navTimer = setTimeout(function () {
        finishProgress();
        orig(null, path);
      }, 150);
      return false;
    };

    window.navigateInApp._csd1pt = true;
  }

  function resetAfterLoad() {
    if (!progressBar) return;
    stopProgressAnimation();
    setProgress(1);
    progressBar.classList.remove('show');
    setTimeout(function () {
      if (!progressBar) return;
      setProgress(0);
    }, 150);
    if (document.body) document.body.classList.remove('csd1-pg-loading');
  }

  setupPrefetch();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wrapNav);
  } else {
    wrapNav();
  }

  window.addEventListener('pageshow', resetAfterLoad);
})();
