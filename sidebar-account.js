(function () {
  'use strict';

  const state = {
    profileOpen: false
  };

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

    const stat = document.getElementById('profileStatOnlineValue');
    if (stat) {
      stat.classList.toggle('is-offline', !online);
    }
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
      '  <div class="profile-stat-card">',
      '    <div class="profile-stat-label">สถานะ</div>',
      '    <div class="profile-stat-value profile-stat-online" id="profileStatOnlineValue">',
      '      <span class="online-dot profile-stat-online-dot" data-presence-dot></span>',
      '      <span class="profile-stat-online-text" data-presence-text>ออนไลน์</span>',
      '    </div>',
      '  </div>',
      '</div>',
      '<div class="profile-sheet-actions">',
      '  <button class="profile-action-btn" id="profileActionHistory" type="button">ประวัติการกรอกข้อมูล</button>',
      '  <button class="profile-action-btn admin" id="profileActionNotification" type="button">แจ้งเตือนคำขอแก้ไข</button>',
      '  <button class="profile-action-btn danger" id="profileActionLogout" type="button">ออกจากระบบ</button>',
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

    const noticeBtn = document.getElementById('profileActionNotification');
    if (noticeBtn) {
      noticeBtn.addEventListener('click', function () {
        window.closeProfileSheet();
        if (typeof window.openNotificationPage === 'function') {
          window.openNotificationPage();
          return;
        }
        navigateFromProfile('notification.html');
      });
    }

    const logoutBtn = document.getElementById('profileActionLogout');
    if (logoutBtn) {
      logoutBtn.addEventListener('click', function (evt) {
        window.handleSidebarLogout(evt);
      });
    }
  }

  function setProfileVisibility(show) {
    ensureProfileSheet();
    const overlay = document.getElementById('profileSheetOverlay');
    const sheet = document.getElementById('profileSheet');
    if (!overlay || !sheet) return;

    const visible = !!show;
    state.profileOpen = visible;
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
      noticeBtn.style.display = isAdminRole(currentUser) ? '' : 'none';
    }
  }

  window.closeProfileSheet = function closeProfileSheet() {
    setProfileVisibility(false);
  };

  window.openProfileSheet = function openProfileSheet() {
    const user = safeParseUser();
    if (!isLoggedIn(user)) return;
    refreshProfileSheet(user);
    setProfileVisibility(true);
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
    window.openProfileSheet();
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
  });

  window.addEventListener('storage', function () {
    if (typeof window.refreshSidebarAccount === 'function') {
      window.refreshSidebarAccount();
    }
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
