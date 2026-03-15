(function () {
  'use strict';

  if (window.__CSD1_CONNECT_MOBILE_NAV_READY__) return;
  window.__CSD1_CONNECT_MOBILE_NAV_READY__ = true;

  var REPORT_GROUP_PAGES = {
    'actionreport.html': true,
    'dashboard.html': true,
    'compare.html': true,
    'rank.html': true,
    'history.html': true,
    'notification.html': true
  };

  function getCurrentPage() {
    try {
      var raw = window.location.pathname.split('/').pop() || '';
      return decodeURIComponent(raw || '').trim().toLowerCase();
    } catch (_) {
      return '';
    }
  }

  function getActiveKey(page) {
    if (page === 'index.html') return 'home';
    if (page === 'csd1 connect.html') return 'request';
    if (page === 'connect-status.html') return 'status';
    if (REPORT_GROUP_PAGES[page]) return 'report';
    return '';
  }

  function ensureStyle() {
    if (document.getElementById('csd1ConnectMobileNavStyle')) return;
    var style = document.createElement('style');
    style.id = 'csd1ConnectMobileNavStyle';
    style.textContent = [
      'body.has-unified-mobile-nav .mobile-nav-fade {',
      '  display: none !important;',
      '  opacity: 0 !important;',
      '  visibility: hidden !important;',
      '  pointer-events: none !important;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav {',
      '  display: none;',
      '  position: fixed !important;',
      '  left: max(0.5rem, env(safe-area-inset-left)) !important;',
      '  right: max(0.5rem, env(safe-area-inset-right)) !important;',
      '  bottom: calc(env(safe-area-inset-bottom) + 0.35rem) !important;',
      '  height: 56px;',
      '  z-index: 4310;',
      '  margin: 0 !important;',
      '  transform: none !important;',
      '  translate: none !important;',
      '  background: rgba(6,14,35,0.94) !important;',
      '  backdrop-filter: blur(16px);',
      '  -webkit-backdrop-filter: blur(16px);',
      '  border: 1px solid rgba(255,255,255,0.1) !important;',
      '  border-radius: 18px;',
      '  align-items: center;',
      '  justify-content: space-around;',
      '  padding: 0 0.45rem;',
      '  box-shadow: 0 8px 30px rgba(0,0,0,0.4);',
      '  isolation: isolate;',
      '}',
      '[data-theme="light"] .mobile-mini-nav.csd1-unified-mobile-nav {',
      '  background: rgba(255,255,255,0.97) !important;',
      '  border-color: rgba(0,38,75,0.15) !important;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings {',
      '  display: flex;',
      '  flex-direction: column;',
      '  align-items: center;',
      '  justify-content: center;',
      '  gap: 2px;',
      '  flex: 1;',
      '  min-width: 0;',
      '  height: 100%;',
      '  padding: 0.24rem 0;',
      '  border-radius: 10px;',
      '  color: var(--text-muted);',
      '  text-decoration: none;',
      '  font-size: 0.56rem;',
      '  font-weight: 600;',
      '  line-height: 1.02;',
      '  text-align: center;',
      '  transition: color 0.2s;',
      '  cursor: pointer;',
      '  background: none;',
      '  border: none;',
      '  font-family: inherit;',
      '  appearance: none;',
      '  -webkit-appearance: none;',
      '  box-shadow: none !important;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link span {',
      '  display: block;',
      '  white-space: normal;',
      '  word-break: keep-all;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link:hover,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link:focus-visible,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link:active,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings:hover,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings:focus-visible,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings:active {',
      '  color: #FFD400;',
      '  background: none !important;',
      '  box-shadow: none !important;',
      '  border-color: transparent !important;',
      '  outline: none;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link.active,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings.active {',
      '  color: #FFD400;',
      '  background: none !important;',
      '  box-shadow: none !important;',
      '  border-color: transparent !important;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-link svg,',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings svg {',
      '  width: 22px;',
      '  height: 22px;',
      '  stroke: currentColor;',
      '  stroke-width: 2;',
      '  fill: none;',
      '  stroke-linecap: round;',
      '  stroke-linejoin: round;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-settings {',
      '  flex: 0 0 54px;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-profile-avatar {',
      '  width: 30px;',
      '  height: 30px;',
      '  border-radius: 50%;',
      '  display: inline-flex;',
      '  align-items: center;',
      '  justify-content: center;',
      '  overflow: hidden;',
      '  background: linear-gradient(135deg, #FFD400, #f59e0b);',
      '  color: #0f172a;',
      '  font-size: 0.82rem;',
      '  font-weight: 700;',
      '  box-shadow: 0 2px 10px rgba(255,212,0,0.35), 0 0 0 2px rgba(255,212,0,0.18);',
      '  text-transform: uppercase;',
      '}',
      '.mobile-mini-nav.csd1-unified-mobile-nav .mobile-mini-profile-avatar img {',
      '  width: 100%;',
      '  height: 100%;',
      '  object-fit: cover;',
      '  border-radius: 50%;',
      '}',
      '@media (max-width: 900px) {',
      '  .mobile-mini-nav.csd1-unified-mobile-nav {',
      '    display: flex !important;',
      '  }',
      '}',
      '@media (min-width: 901px) {',
      '  .mobile-mini-nav.csd1-unified-mobile-nav {',
      '    display: none !important;',
      '  }',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  function buildLink(key, href, labelHtml, iconHtml, activeKey) {
    var isActive = key === activeKey;
    return [
      '<a class="mobile-mini-link',
      isActive ? ' active' : '',
      '" href="',
      href,
      '" title="',
      labelHtml.replace(/<br\s*\/?>/gi, ' '),
      '" aria-label="',
      labelHtml.replace(/<br\s*\/?>/gi, ' '),
      '"',
      isActive ? ' aria-current="page"' : '',
      '>',
      iconHtml,
      '<span>',
      labelHtml,
      '</span>',
      '</a>'
    ].join('');
  }

  function buildProfileButton() {
    return [
      '<button class="mobile-mini-link mobile-mini-settings" type="button" title="โปรไฟล์" aria-label="โปรไฟล์">',
      '  <span class="mobile-mini-profile-avatar" id="mobileNavProfileAvatar">U</span>',
      '</button>'
    ].join('');
  }

  function buildMarkup(activeKey) {
    return [
      buildLink('home', 'index.html', 'หน้าหลัก', '<svg viewBox="0 0 24 24" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9.5L12 3l9 6.5V20a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"></path><polyline points="9 21 9 12 15 12 15 21"></polyline></svg>', activeKey),
      buildLink('request', 'CSD1 connect.html', 'ขอข้อมูล', '<svg viewBox="0 0 24 24" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"></rect><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line></svg>', activeKey),
      buildLink('status', 'connect-status.html', 'สถานะ', '<svg viewBox="0 0 24 24" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>', activeKey),
      buildLink('report', 'dashboard.html', 'แดชบอร์ด', '<svg viewBox="0 0 24 24" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="13" width="4" height="8" rx="1"></rect><rect x="10" y="9" width="4" height="12" rx="1"></rect><rect x="17" y="5" width="4" height="16" rx="1"></rect></svg>', activeKey),
      buildProfileButton()
    ].join('');
  }

  function bindNavActions(nav) {
    nav.querySelectorAll('a.mobile-mini-link').forEach(function (link) {
      link.addEventListener('click', function (evt) {
        evt.preventDefault();
        var href = link.getAttribute('href');
        if (typeof window.navigateInApp === 'function') {
          window.navigateInApp(evt, href);
        } else {
          window.location.assign(href);
        }
      });
    });

    var profileBtn = nav.querySelector('.mobile-mini-settings');
    if (!profileBtn) return;
    profileBtn.addEventListener('click', function (evt) {
      if (typeof window.handleAvatarAccess === 'function') {
        window.handleAvatarAccess(evt);
        return;
      }
      window.location.assign('index.html');
    });
  }

  function mount() {
    ensureStyle();

    var page = getCurrentPage();
    var activeKey = getActiveKey(page);
    var nav = document.querySelector('.mobile-mini-nav');

    if (!nav) {
      nav = document.createElement('nav');
      nav.className = 'mobile-mini-nav';
      nav.setAttribute('aria-label', 'เมนูหลักมือถือ');
    }

    nav.className = 'mobile-mini-nav csd1-unified-mobile-nav';
    nav.innerHTML = buildMarkup(activeKey);

    if (nav.parentNode !== document.body) {
      document.body.appendChild(nav);
    }

    document.body.classList.add('has-unified-mobile-nav');
    bindNavActions(nav);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mount, { once: true });
  } else {
    mount();
  }
})();
