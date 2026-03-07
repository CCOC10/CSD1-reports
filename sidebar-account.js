(function () {
  'use strict';

  const state = {
    profileOpen: false,
    notificationUnread: 0,
    notificationTotal: 0,
    notificationLastFetch: 0,
    notificationUserKey: '',
    notificationInFlight: null
  };

  const SIDEBAR_REFRESH_DEFAULT_SHEET_URL = 'https://script.google.com/macros/s/AKfycbytUmPr668UwcPsbmwtZm9wSL3qnhDjYJG8b7DzIiqxqyo6vKLVtDgGdmdvA2hE00xX9Q/exec';

  function safeParseUser() {
    try {
      return JSON.parse(localStorage.getItem('CSD1_USER') || '{}');
    } catch (_) {
      return {};
    }
  }

  function normalizeText(value) {
    return String(value == null ? '' : value).trim();
  }

  function getSessionToken(user) {
    if (!user || typeof user !== 'object') return '';
    return normalizeText(
      user.sessionToken ||
      user.session_token ||
      (user.session && user.session.token) ||
      user.token ||
      user.accessToken ||
      ''
    );
  }

  function isLoggedIn(user) {
    return !!(normalizeText(user && user.email) && getSessionToken(user));
  }

  function isAdminRole(user) {
    return normalizeText(user && user.role).toLowerCase() === 'admin';
  }

  function isStandardUserRole(user) {
    return normalizeText(user && user.role).toLowerCase() === 'user';
  }

  function setAvatar(el, displayName, photoUrl) {
    if (!el) return;
    const name = normalizeText(displayName) || 'ผู้ใช้';
    const initial = name.charAt(0).toUpperCase() || 'U';
    el.textContent = '';

    if (photoUrl) {
      const img = document.createElement('img');
      img.src = photoUrl;
      img.alt = name;
      img.referrerPolicy = 'no-referrer';
      img.addEventListener('error', function () {
        el.textContent = initial;
      });
      el.appendChild(img);
    } else {
      el.textContent = initial;
    }
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function setDisabled(id, disabled) {
    const el = document.getElementById(id);
    if (!el) return;
    el.disabled = !!disabled;
  }

  function setHidden(id, hidden) {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.toggle('is-hidden', !!hidden);
  }

  function createPresenceRow(rowClass, textClass) {
    const row = document.createElement('span');
    row.className = rowClass;

    const dot = document.createElement('span');
    dot.className = 'online-dot';
    dot.setAttribute('data-presence-dot', '');

    const text = document.createElement('span');
    text.className = textClass;
    text.setAttribute('data-presence-text', '');
    text.textContent = 'ออนไลน์';

    row.appendChild(dot);
    row.appendChild(text);
    return row;
  }

  function ensurePresenceBadges() {
    const sidebarMeta = document.querySelector('#sidebarAccountBtn .sidebar-account-meta');
    if (sidebarMeta && !sidebarMeta.querySelector('.sidebar-account-presence')) {
      sidebarMeta.appendChild(createPresenceRow('sidebar-account-presence', 'sidebar-account-presence-text'));
    }

    const mobileText = document.querySelector('.mobile-account-btn .mobile-account-text');
    if (mobileText && !mobileText.querySelector('.mobile-account-presence')) {
      mobileText.appendChild(createPresenceRow('mobile-account-presence', 'mobile-account-presence-text'));
    }

    // Inject avatar-badge dot directly onto avatar images
    ['sidebarAccountAvatar', 'mobileAccountAvatar'].forEach(function (id) {
      var avatarEl = document.getElementById(id);
      if (!avatarEl) return;
      if (!avatarEl.querySelector('.avatar-presence-dot')) {
        var badge = document.createElement('span');
        badge.className = 'online-dot avatar-presence-dot';
        badge.setAttribute('data-presence-dot', '');
        badge.setAttribute('aria-hidden', 'true');
        avatarEl.appendChild(badge);
      }
    });
  }

  function applyPresence(logged) {
    const online = !!logged;
    const text = online ? 'ออนไลน์' : 'ออฟไลน์';

    document.querySelectorAll('[data-presence-text]').forEach(function (el) {
      el.textContent = text;
    });

    document.querySelectorAll('[data-presence-dot]').forEach(function (el) {
      el.classList.toggle('is-offline', !online);
    });

  }

  function ensureSidebarProfileTop() {
    const sidebar = document.getElementById('appSidebar');
    const accountBtn = document.getElementById('sidebarAccountBtn');
    if (!sidebar || !accountBtn) return;

    let top = document.getElementById('sidebarProfileTop');
    if (!top) {
      top = document.createElement('div');
      top.id = 'sidebarProfileTop';
      top.className = 'sidebar-profile-top';
    }

    let divider = document.getElementById('sidebarProfileDivider');
    if (!divider) {
      divider = document.createElement('div');
      divider.id = 'sidebarProfileDivider';
      divider.className = 'sidebar-profile-divider';
    }

    const head = sidebar.querySelector('.app-sidebar-head');
    if (head) {
      sidebar.insertBefore(top, head);
      sidebar.insertBefore(divider, head);
    } else {
      if (!top.parentElement) {
        sidebar.prepend(top);
      }
      if (!divider.parentElement) {
        sidebar.appendChild(divider);
      }
    }

    if (accountBtn.parentElement !== top) {
      top.appendChild(accountBtn);
    }
  }

  function closePanels() {
    if (typeof window.closeAppSidebar === 'function') window.closeAppSidebar();
    if (typeof window.closeMobileSettings === 'function') window.closeMobileSettings();
  }

  function navigateTo(path) {
    if (typeof window.navigateInApp === 'function') {
      try {
        window.navigateInApp(null, path);
        return;
      } catch (_) {}
    }
    window.location.assign(path);
  }

  function navigateFromProfile(path) {
    window.closeProfileSheet();
    closePanels();
    navigateTo(path);
  }

  function toast(message) {
    if (typeof window.showToast === 'function') {
      window.showToast(message);
      return;
    }
    try {
      window.alert(message);
    } catch (_) {}
  }

  function formatUnit(user) {
    return normalizeText(user && user.unit) || '-';
  }

  function formatRole(user) {
    const role = normalizeText(user && user.role).toLowerCase();
    if (role === 'admin') return 'Admin';
    if (role === 'user') return 'User';
    return role ? role.toUpperCase() : '-';
  }

  function normalizeStatusForAlert(value) {
    const status = normalizeText(value).toLowerCase();
    if (!status) return 'new';
    if (status === 'done' || status === 'completed' || status === 'resolved' || status === 'closed') return 'done';
    if (status === 'processing' || status === 'in_progress' || status === 'in-progress' || status === 'working') return 'processing';
    if (status === 'rejected' || status === 'cancelled' || status === 'canceled') return 'rejected';
    return 'new';
  }

  function getSidebarSheetUrl() {
    if (typeof window.getSheetUrl === 'function') {
      try {
        const resolved = normalizeText(window.getSheetUrl());
        if (resolved) return resolved;
      } catch (_) {}
    }
    const stored = normalizeText(localStorage.getItem('CSD1_SHEET_URL'));
    if (stored) {
      if (/^https?:\/\//i.test(stored)) return stored;
      return 'https://' + stored.replace(/^\/+/, '');
    }
    return SIDEBAR_REFRESH_DEFAULT_SHEET_URL;
  }

  function setNotificationState(unread, total) {
    const unreadSafe = Number.isFinite(Number(unread)) ? Math.max(0, Math.floor(Number(unread))) : 0;
    const totalSafe = Number.isFinite(Number(total)) ? Math.max(0, Math.floor(Number(total))) : 0;
    state.notificationUnread = unreadSafe;
    state.notificationTotal = totalSafe;

    const noticeBtn = document.getElementById('profileActionNotification');
    const badge = document.getElementById('profileNotificationBadge');
    if (!noticeBtn) return;
    const hasUnread = unreadSafe > 0;
    noticeBtn.classList.toggle('has-unread', hasUnread);
    if (badge) {
      badge.hidden = !hasUnread;
      badge.textContent = hasUnread ? (unreadSafe > 99 ? '99+' : String(unreadSafe)) : '';
    }
  }

  function refreshNotificationIndicator(options) {
    const opts = options && typeof options === 'object' ? options : {};
    const force = !!opts.force;
    const user = opts.user || safeParseUser();

    const adminRole = isAdminRole(user);
    const standardUserRole = isStandardUserRole(user);

    if (!isLoggedIn(user) || (!adminRole && !standardUserRole)) {
      state.notificationUserKey = '';
      state.notificationLastFetch = 0;
      state.notificationInFlight = null;
      setNotificationState(0, 0);
      return Promise.resolve({ unread: 0, total: 0 });
    }

    const roleKey = adminRole ? 'admin' : 'user';
    const userKey = roleKey + '|' + normalizeText(user.email).toLowerCase() + '|' + normalizeText(getSessionToken(user));
    if (state.notificationUserKey !== userKey) {
      state.notificationUserKey = userKey;
      state.notificationLastFetch = 0;
      state.notificationInFlight = null;
      setNotificationState(0, 0);
    }

    const now = Date.now();
    const staleMs = 30000;
    if (!force && state.notificationLastFetch && (now - state.notificationLastFetch) < staleMs) {
      return Promise.resolve({ unread: state.notificationUnread, total: state.notificationTotal });
    }

    if (state.notificationInFlight) return state.notificationInFlight;

    const sheetUrl = getSidebarSheetUrl();
    if (!sheetUrl) return Promise.resolve({ unread: state.notificationUnread, total: state.notificationTotal });

    const requestUrl = sheetUrl + (sheetUrl.indexOf('?') >= 0 ? '&' : '?') + 'ts=' + now;
    const payload = {
      action: adminRole ? 'getAdminNotifications' : 'getUnitHistory',
      email: user.email,
      sessionToken: getSessionToken(user)
    };

    state.notificationInFlight = fetch(requestUrl, {
      method: 'POST',
      cache: 'no-store',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    })
      .then(function (response) {
        return response.text().then(function (rawText) {
          let parsed = {};
          try {
            parsed = rawText ? JSON.parse(rawText) : {};
          } catch (_) {
            parsed = {};
          }
          return { ok: response.ok, parsed: parsed };
        });
      })
      .then(function (res) {
        if (!res.ok || !res.parsed || res.parsed.status !== 'success') {
          return { unread: state.notificationUnread, total: state.notificationTotal };
        }
        const rows = Array.isArray(res.parsed.data) ? res.parsed.data : [];
        let unread = 0;
        let total = 0;
        if (adminRole) {
          unread = rows.filter(function (row) {
            return normalizeStatusForAlert(row && row.status) === 'new';
          }).length;
          total = rows.length;
        } else {
          const requesterEmail = normalizeText(user.email).toLowerCase();
          const myRequests = rows.filter(function (row) {
            const requestId = normalizeText(row && row.changeRequestId);
            if (!requestId) return false;
            const byEmail = normalizeText(row && row.changeRequestedByEmail).toLowerCase();
            if (byEmail) return byEmail === requesterEmail;
            const fallbackReporter = normalizeText(row && row['อีเมล']).toLowerCase();
            return fallbackReporter === requesterEmail;
          });
          total = myRequests.length;
          unread = myRequests.filter(function (row) {
            const status = normalizeStatusForAlert(row && row.changeRequestStatus);
            return status === 'done' || status === 'rejected';
          }).length;
        }
        setNotificationState(unread, total);
        state.notificationLastFetch = Date.now();
        return { unread: unread, total: total };
      })
      .catch(function () {
        return { unread: state.notificationUnread, total: state.notificationTotal };
      })
      .finally(function () {
        state.notificationInFlight = null;
      });

    return state.notificationInFlight;
  }

  function getCurrentThemeMode() {
    const attrMode = normalizeText(document.documentElement.getAttribute('data-theme')).toLowerCase();
    if (attrMode === 'light' || attrMode === 'dark') return attrMode;
    const storedMode = normalizeText(localStorage.getItem('CSD1_THEME')).toLowerCase();
    if (storedMode === 'light' || storedMode === 'dark') return storedMode;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  function setThemeMode(mode) {
    const normalized = mode === 'light' ? 'light' : 'dark';
    if (typeof window.setTheme === 'function') {
      window.setTheme(normalized);
      return;
    }
    if (typeof window.applyTheme === 'function') {
      try {
        window.applyTheme(normalized);
      } catch (_) {}
    }
    try {
      localStorage.setItem('CSD1_THEME', normalized);
    } catch (_) {}
    document.documentElement.setAttribute('data-theme', normalized);
    const isDark = normalized === 'dark';
    const sidebarToggle = document.getElementById('sidebarThemeToggle');
    if (sidebarToggle) sidebarToggle.checked = isDark;
    const mobileToggle = document.getElementById('mobileThemeToggle');
    if (mobileToggle) mobileToggle.checked = isDark;
  }

  function syncProfileThemeToggle() {
    const toggle = document.getElementById('profileThemeToggle');
    if (!toggle) return;
    toggle.checked = getCurrentThemeMode() === 'dark';
  }

  function ensureProfileSheet() {
    if (document.getElementById('profileSheetOverlay') && document.getElementById('profileSheet')) return;

    const overlay = document.createElement('div');
    overlay.id = 'profileSheetOverlay';
    overlay.className = 'profile-sheet-overlay';

    const sheet = document.createElement('section');
    sheet.id = 'profileSheet';
    sheet.className = 'profile-sheet';
    sheet.setAttribute('role', 'dialog');
    sheet.setAttribute('aria-modal', 'true');
    sheet.setAttribute('aria-label', 'ข้อมูลบัญชีผู้ใช้งาน');
    sheet.innerHTML = [
      '<div class="profile-sheet-header">',
      '  <div class="profile-sheet-title">โปรไฟล์ผู้ใช้งาน</div>',
      '  <button class="profile-sheet-close" id="profileSheetCloseBtn" type="button" aria-label="ปิดหน้าต่าง">&times;</button>',
      '</div>',
      '<div class="profile-sheet-top">',
      '  <span class="profile-sheet-avatar" id="profileSheetAvatar">U</span>',
      '  <div class="profile-sheet-name" id="profileSheetName">ผู้ใช้งาน</div>',
      '  <div class="profile-sheet-sub" id="profileSheetSub">-</div>',
      '</div>',
      '<div class="profile-sheet-stats">',
      '  <div class="profile-stat-card">',
      '    <div class="profile-stat-label">ชป.</div>',
      '    <div class="profile-stat-value" id="profileStatUnit">-</div>',
      '  </div>',
      '  <div class="profile-stat-card">',
      '    <div class="profile-stat-label">สิทธิ์</div>',
      '    <div class="profile-stat-value" id="profileStatRole">-</div>',
      '  </div>',
      '</div>',
      '<div class="profile-sheet-actions">',
      '  <div class="profile-theme-switch-row">',
      '    <label class="app-theme-switch" for="profileThemeToggle" title="สลับโหมดสว่าง/มืด">',
      '      <input type="checkbox" id="profileThemeToggle" aria-label="สลับโหมดมืด">',
      '      <span class="app-theme-switch-track">',
      '        <span class="app-theme-switch-thumb">',
      '          <svg class="app-theme-switch-sun" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4.2"></circle><line x1="12" y1="2.8" x2="12" y2="5.2"></line><line x1="12" y1="18.8" x2="12" y2="21.2"></line><line x1="2.8" y1="12" x2="5.2" y2="12"></line><line x1="18.8" y1="12" x2="21.2" y2="12"></line><line x1="5.5" y1="5.5" x2="7.2" y2="7.2"></line><line x1="16.8" y1="16.8" x2="18.5" y2="18.5"></line><line x1="16.8" y1="7.2" x2="18.5" y2="5.5"></line><line x1="5.5" y1="18.5" x2="7.2" y2="16.8"></line></svg>',
      '          <svg class="app-theme-switch-moon" viewBox="0 0 24 24" aria-hidden="true"><path d="M20.5 14.6A8.8 8.8 0 1 1 11 3.5a6.9 6.9 0 0 0 9.5 11.1z"></path></svg>',
      '        </span>',
      '      </span>',
      '    </label>',
      '  </div>',
      '  <button class="profile-action-btn with-action-icon" id="profileActionHistory" type="button">',
      '    <svg class="action-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">',
      '      <path fill-rule="evenodd" clip-rule="evenodd" d="M5.01112 11.5747L6.29288 10.2929C6.68341 9.90236 7.31657 9.90236 7.7071 10.2929C8.09762 10.6834 8.09762 11.3166 7.7071 11.7071L4.7071 14.7071C4.51956 14.8946 4.26521 15 3.99999 15C3.73477 15 3.48042 14.8946 3.29288 14.7071L0.292884 11.7071C-0.0976406 11.3166 -0.0976406 10.6834 0.292884 10.2929C0.683408 9.90236 1.31657 9.90236 1.7071 10.2929L3.0081 11.5939C3.22117 6.25933 7.61317 2 13 2C18.5229 2 23 6.47715 23 12C23 17.5228 18.5229 22 13 22C9.85817 22 7.05429 20.5499 5.22263 18.2864C4.87522 17.8571 4.94163 17.2274 5.37096 16.88C5.80028 16.5326 6.42996 16.599 6.77737 17.0283C8.24562 18.8427 10.4873 20 13 20C17.4183 20 21 16.4183 21 12C21 7.58172 17.4183 4 13 4C8.72441 4 5.23221 7.35412 5.01112 11.5747ZM13 5C13.5523 5 14 5.44772 14 6V11.5858L16.7071 14.2929C17.0976 14.6834 17.0976 15.3166 16.7071 15.7071C16.3166 16.0976 15.6834 16.0976 15.2929 15.7071L12.2929 12.7071C12.1054 12.5196 12 12.2652 12 12V6C12 5.44772 12.4477 5 13 5Z" fill="currentColor"></path>',
      '    </svg>',
      '    <span class="action-label">ตรวจสอบ แก้ไขข้อมูล</span>',
      '  </button>',
      '  <button class="profile-action-btn admin with-action-icon" id="profileActionNotification" type="button">',
      '    <svg class="action-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">',
      '      <path d="M6 10.5a6 6 0 1 1 12 0v3.2l1.2 2.2a1 1 0 0 1-.88 1.48H5.68a1 1 0 0 1-.88-1.48L6 13.7v-3.2Z" stroke="currentColor" stroke-width="2" stroke-linejoin="round"></path>',
      '      <path d="M10 19a2 2 0 0 0 4 0" stroke="currentColor" stroke-width="2" stroke-linecap="round"></path>',
      '    </svg>',
      '    <span class="action-label">แจ้งเตือนคำขอแก้ไข</span>',
      '    <span class="profile-action-badge" id="profileNotificationBadge" hidden>0</span>',
      '  </button>',
      '  <button class="profile-action-btn danger with-action-icon" id="profileActionLogout" type="button">',
      '    <svg class="action-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">',
      '      <g clip-path="url(#clip0_429_11067)">',
      '        <path d="M15 4.00098H5V18.001C5 19.1055 5.89543 20.001 7 20.001H15" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>',
      '        <path d="M16 15.001L19 12.001M19 12.001L16 9.00098M19 12.001H9" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>',
      '      </g>',
      '      <defs>',
      '        <clipPath id="clip0_429_11067">',
      '          <rect width="24" height="24" fill="white" transform="translate(0 0.000976562)"></rect>',
      '        </clipPath>',
      '      </defs>',
      '    </svg>',
      '    <span class="action-label">ออกจากระบบ</span>',
      '  </button>',
      '</div>'
    ].join('');

    document.body.appendChild(overlay);
    document.body.appendChild(sheet);

    overlay.addEventListener('click', function () {
      window.closeProfileSheet();
    });

    const closeBtn = document.getElementById('profileSheetCloseBtn');
    if (closeBtn) {
      closeBtn.addEventListener('click', function () {
        window.closeProfileSheet();
      });
    }

    const historyBtn = document.getElementById('profileActionHistory');
    if (historyBtn) {
      historyBtn.addEventListener('click', function () {
        window.closeProfileSheet();
        if (typeof window.openHistoryPage === 'function') {
          window.openHistoryPage();
          return;
        }
        navigateFromProfile('history.html');
      });
    }

    const profileThemeToggle = document.getElementById('profileThemeToggle');
    if (profileThemeToggle) {
      profileThemeToggle.addEventListener('change', function (evt) {
        const isDark = !!(evt && evt.target && evt.target.checked);
        setThemeMode(isDark ? 'dark' : 'light');
        syncProfileThemeToggle();
      });
    }

    const noticeBtn = document.getElementById('profileActionNotification');
    if (noticeBtn) {
      noticeBtn.addEventListener('click', function () {
        window.closeProfileSheet();
        const latestUser = safeParseUser();
        if (isAdminRole(latestUser)) {
          if (typeof window.openNotificationPage === 'function') {
            window.openNotificationPage();
            return;
          }
          navigateFromProfile('notification.html');
          return;
        }
        if (typeof window.openHistoryPage === 'function') {
          window.openHistoryPage();
          return;
        }
        navigateFromProfile('history.html');
      });
    }

    const logoutBtn = document.getElementById('profileActionLogout');
    if (logoutBtn) {
      logoutBtn.addEventListener('click', function (evt) {
        window.handleSidebarLogout(evt);
      });
    }
  }

  function positionSheetNearTrigger(sheet, triggerEl) {
    const MARGIN = 15;
    const GAP = 8;
    const sheetW = Math.min(360, window.innerWidth - MARGIN * 2);
    sheet.style.width = sheetW + 'px';
    if (!triggerEl) return;
    const rect = triggerEl.getBoundingClientRect();
    let left = rect.left + rect.width / 2 - sheetW / 2;
    left = Math.max(MARGIN, Math.min(left, window.innerWidth - sheetW - MARGIN));
    sheet.style.left = left + 'px';
    sheet.style.right = 'auto';
    sheet.style.top = 'auto';
    sheet.style.bottom = (window.innerHeight - rect.top + GAP) + 'px';
  }

  function resetSheetPosition(sheet) {
    sheet.style.left = '';
    sheet.style.right = '';
    sheet.style.top = '';
    sheet.style.bottom = '';
    sheet.style.width = '';
  }

  function setProfileVisibility(show, triggerEl) {
    ensureProfileSheet();
    const overlay = document.getElementById('profileSheetOverlay');
    const sheet = document.getElementById('profileSheet');
    if (!overlay || !sheet) return;

    const visible = !!show;
    state.profileOpen = visible;
    if (visible) {
      positionSheetNearTrigger(sheet, triggerEl);
    } else {
      resetSheetPosition(sheet);
    }
    overlay.classList.toggle('show', visible);
    sheet.classList.toggle('show', visible);
    document.body.classList.toggle('profile-sheet-open', visible);
  }

  function refreshProfileSheet(user) {
    ensureProfileSheet();
    const currentUser = user || safeParseUser();
    if (!isLoggedIn(currentUser)) {
      setProfileVisibility(false);
      return;
    }

    const displayName = normalizeText(currentUser.name) || normalizeText(currentUser.email) || 'ผู้ใช้';
    const unit = formatUnit(currentUser);
    const role = formatRole(currentUser);
    const subLine = normalizeText(currentUser.email) || normalizeText(currentUser.position) || '-';
    const photoUrl = normalizeText(currentUser.avatarUrl || currentUser.googlePhoto);

    setAvatar(document.getElementById('profileSheetAvatar'), displayName, photoUrl);
    setText('profileSheetName', displayName);
    setText('profileSheetSub', subLine);
    setText('profileStatUnit', unit);
    setText('profileStatRole', role);

    const noticeBtn = document.getElementById('profileActionNotification');
    if (noticeBtn) {
      const canViewNotice = isAdminRole(currentUser) || isStandardUserRole(currentUser);
      noticeBtn.style.display = canViewNotice ? '' : 'none';
      const noticeLabel = noticeBtn.querySelector('.action-label');
      if (noticeLabel) {
        noticeLabel.textContent = isAdminRole(currentUser) ? 'แจ้งเตือนคำขอแก้ไข' : 'ผลคำร้องของฉัน';
      }
    }
    syncProfileThemeToggle();
    setNotificationState(state.notificationUnread, state.notificationTotal);
    refreshNotificationIndicator({ user: currentUser, force: false });
  }

  window.closeProfileSheet = function closeProfileSheet() {
    setProfileVisibility(false);
  };

  window.openProfileSheet = function openProfileSheet(triggerEl) {
    const user = safeParseUser();
    if (!isLoggedIn(user)) return;
    refreshProfileSheet(user);
    setProfileVisibility(true, triggerEl);
  };

  window.refreshSidebarAccount = function refreshSidebarAccount() {
    const user = safeParseUser();
    const logged = isLoggedIn(user);
    const displayName = logged ? (normalizeText(user.name) || normalizeText(user.email)) : 'เข้าสู่ระบบ';
    const subText = logged ? formatUnit(user) : 'แตะเพื่อเข้าสู่ระบบ';
    const photoUrl = normalizeText(user.avatarUrl || user.googlePhoto);

    ensureSidebarProfileTop();
    ensurePresenceBadges();

    setAvatar(document.getElementById('sidebarAccountAvatar'), displayName, photoUrl);
    setAvatar(document.getElementById('mobileAccountAvatar'), displayName, photoUrl);

    setText('sidebarAccountName', displayName || 'เข้าสู่ระบบ');
    setText('mobileAccountName', displayName || 'เข้าสู่ระบบ');
    setText('sidebarAccountSub', subText);
    setText('mobileAccountSub', subText);

    setDisabled('sidebarLogoutBtn', !logged);
    setDisabled('mobileLogoutBtn', !logged);
    setHidden('sidebarLogoutBtn', true);
    setHidden('mobileLogoutBtn', true);

    refreshProfileSheet(user);
    applyPresence(logged);
    refreshNotificationIndicator({ user: user, force: false });
  };

  window.handleAvatarAccess = function handleAvatarAccess(evt) {
    if (evt) evt.preventDefault();
    const user = safeParseUser();
    if (!isLoggedIn(user)) {
      window.closeProfileSheet();
      closePanels();
      toast('กรุณาเข้าสู่ระบบ');
      navigateTo('index.html');
      return;
    }

    closePanels();
    refreshProfileSheet(user);
    refreshNotificationIndicator({ user: user, force: true });
    const triggerEl = evt && (evt.currentTarget || evt.target);
    window.openProfileSheet(triggerEl);
  };

  window.handleSidebarLogout = function handleSidebarLogout(evt) {
    if (evt) evt.preventDefault();
    const user = safeParseUser();

    if (!isLoggedIn(user)) {
      window.closeProfileSheet();
      closePanels();
      toast('ยังไม่ได้เข้าสู่ระบบ');
      navigateTo('index.html');
      return;
    }

    window.closeProfileSheet();
    if (typeof window.handleLogout === 'function') {
      window.handleLogout();
      if (typeof window.refreshSidebarAccount === 'function') window.refreshSidebarAccount();
      return;
    }

    localStorage.removeItem('CSD1_USER');
    if (typeof window.refreshSidebarAccount === 'function') window.refreshSidebarAccount();
    closePanels();
    toast('ออกจากระบบแล้ว');
    setTimeout(function () {
      navigateTo('index.html');
    }, 180);
  };

  document.addEventListener('DOMContentLoaded', function () {
    if (typeof window.refreshSidebarAccount === 'function') {
      window.refreshSidebarAccount();
    }
    refreshNotificationIndicator({ user: safeParseUser(), force: true });
  });

  window.addEventListener('storage', function () {
    if (typeof window.refreshSidebarAccount === 'function') {
      window.refreshSidebarAccount();
    }
    refreshNotificationIndicator({ user: safeParseUser(), force: true });
  });

  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState !== 'visible') return;
    refreshNotificationIndicator({ user: safeParseUser(), force: true });
  });

  window.addEventListener('focus', function () {
    refreshNotificationIndicator({ user: safeParseUser(), force: true });
  });

  document.addEventListener('keydown', function (evt) {
    if (!state.profileOpen) return;
    if (evt.key === 'Escape') {
      window.closeProfileSheet();
    }
  });

  setInterval(function () {
    if (typeof window.refreshSidebarAccount === 'function') {
      window.refreshSidebarAccount();
    }
  }, 3000);
})();
