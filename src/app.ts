import {
  createRegistration,
  updateRegistration,
  deleteRegistration,
  clearAllRegistrations,
  checkPhoneTaken,
  subscribeToRegistrations,
  fetchAllRegistrations,
  signInAdminWithGoogle,
  signOutAdmin,
  listenAuthState,
  ADMIN_EMAIL,
  RegistrationRecord,
  getCachedAccessToken,
  getSheetsConfig,
  saveSheetsConfig,
  GoogleSheetsConfig,
} from './firebase';
import {
  createEventSpreadsheet,
  syncAllRegistrationsToSpreadsheet,
  appendRegistrationToSpreadsheet,
  getSpreadsheetInfo,
} from './sheets';

(function () {
  const $ = function <T extends HTMLElement = HTMLElement>(id: string): T {
    const el = document.getElementById(id);
    if (!el) throw new Error(`Element with id "${id}" not found`);
    return el as T;
  };

  const MSG = {
    name: '이름을 한글 또는 영문으로 입력해 주세요.',
    phone: '01로 시작하는 휴대폰 번호를 입력해 주세요.',
    email: '이메일 주소를 정확히 입력해 주세요.',
    count: '1명 이상 입력해 주세요.',
    agree: '개인정보 수집, 이용에 동의해 주세요.',
    dup: '이미 신청된 번호입니다.',
  };

  /* ========== 로컬 캐시 및 파이어베이스 상태 동기화 ========== */
  const KEY = 'event-form-db-v1';
  let memoryRecords: RegistrationRecord[] = [];
  let unsubscribeLive: (() => void) | null = null;
  let currentAdminUser: any = null;
  let isCurrentAdmin = false;

  // 로컬스토리지 백업 불러오기 (오프라인 및 초기 캐시용)
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed.rows)) {
        memoryRecords = parsed.rows;
      }
    }
  } catch (e) {
    console.warn('Local cache load error:', e);
  }

  function saveLocalCache(rows: RegistrationRecord[]) {
    try {
      localStorage.setItem(KEY, JSON.stringify({ seq: rows.length, rows }));
    } catch (e) {
      // ignore
    }
  }

  /* ========== 입력 규칙 ========== */
  const VAL = {
    name: function (v: string) {
      return v.trim() && /^[가-힣A-Za-z]{2,20}$/.test(v) ? '' : MSG.name;
    },
    phone: function (v: string) {
      return /^01\d-\d{3,4}-\d{4}$/.test(v) ? '' : MSG.phone;
    },
    email: function (v: string) {
      const ok = v && v.split('@').length === 2 && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v) && !/\.\./.test(v);
      return ok ? '' : MSG.email;
    },
    count: function (v: string) {
      const n = parseInt(v, 10);
      return !n || n < 1 ? MSG.count : '';
    },
  };

  function fmtPhone(v: string) {
    const d = v.replace(/\D/g, '').slice(0, 11);
    if (d.length < 4) return d;
    if (d.length < 8) return d.slice(0, 3) + '-' + d.slice(3);
    return d.slice(0, 3) + '-' + d.slice(3, d.length - 4) + '-' + d.slice(d.length - 4);
  }

  /* Scope helper for input forms */
  function Scope(p: string, hasStepper: boolean) {
    const s = {
      tried: false,
      composing: false,
      timers: {} as Record<string, any>,
      el: function (k: string) {
        return $(p + k) as HTMLInputElement;
      },
      setMsg: function (k: string, text?: string, hint?: boolean) {
        const m = $(p + k + '-msg');
        m.textContent = text || '';
        m.className = 'msg' + (hint ? ' hint' : '');
      },
      setBad: function (k: string, b: boolean) {
        const e = s.el(k);
        const box = k === 'count' && hasStepper ? $(p + 'stepper') : e;
        box.classList.toggle('bad', b);
        e.setAttribute('aria-invalid', b ? 'true' : 'false');
      },
      flash: function (k: string, text: string) {
        clearTimeout(s.timers[k]);
        s.setMsg(k, text, true);
        s.timers[k] = setTimeout(function () {
          if ($(p + k + '-msg').className.indexOf('hint') > -1) s.setMsg(k, '');
        }, 3000);
      },
      check: function (k: 'name' | 'phone' | 'email' | 'count', show?: boolean) {
        const m = VAL[k](s.el(k).value);
        if (show) {
          s.setMsg(k, m);
          s.setBad(k, !!m);
        }
        return m;
      },
      checkAll: function () {
        let first: string | null = null;
        (['name', 'phone', 'email', 'count'] as const).forEach(function (k) {
          if (s.check(k, true) && !first) first = k;
        });
        return first;
      },
      values: function () {
        return {
          name: s.el('name').value.trim(),
          phone: s.el('phone').value,
          email: s.el('email').value.trim().toLowerCase(),
          count: parseInt(s.el('count').value, 10),
          memo: ($(p + 'memo') as HTMLTextAreaElement).value.trim(),
        };
      },
      clean: function () {
        s.el('name').value = s.el('name').value.trim();
        s.el('email').value = s.el('email').value.trim().toLowerCase();
        ($(p + 'memo') as HTMLTextAreaElement).value = ($(p + 'memo') as HTMLTextAreaElement).value.trim();
      },
      clearMsgs: function () {
        (['name', 'phone', 'email', 'count'] as const).forEach(function (k) {
          s.setMsg(k, '');
          s.setBad(k, false);
        });
        s.tried = false;
      },
      refreshCount: function () {
        const memoEl = $(p + 'memo') as HTMLTextAreaElement;
        $(p + 'memo-count').textContent = memoEl.value.length + '/500';
      },
    };

    const name = s.el('name');
    function cleanName() {
      const v = name.value;
      const c = v.replace(/[^가-힣ㄱ-ㅎㅏ-ㅣA-Za-z ]/g, '');
      if (c !== v) {
        name.value = c;
        s.flash('name', '한글과 영문만 입력할 수 있습니다.');
      } else if (s.tried) {
        s.check('name', true);
      }
    }
    name.addEventListener('compositionstart', function () {
      s.composing = true;
    });
    name.addEventListener('compositionend', function () {
      s.composing = false;
      cleanName();
    });
    name.addEventListener('input', function () {
      if (!s.composing) cleanName();
    });
    name.addEventListener('blur', function () {
      name.value = name.value.trim();
      s.check('name', true);
    });

    const phone = s.el('phone');
    phone.addEventListener('input', function () {
      phone.value = fmtPhone(phone.value);
      if (s.tried) s.check('phone', true);
    });
    phone.addEventListener('blur', function () {
      s.check('phone', true);
    });

    const email = s.el('email');
    email.addEventListener('input', function () {
      const v = email.value;
      const c = v.replace(/[^A-Za-z0-9@._+\-]/g, '').slice(0, 100);
      if (c !== v) email.value = c;
      if (s.tried) s.check('email', true);
    });
    email.addEventListener('blur', function () {
      email.value = email.value.trim().toLowerCase();
      s.check('email', true);
    });

    const count = s.el('count');
    count.addEventListener('input', function () {
      count.value = count.value.replace(/\D/g, '').slice(0, 2);
      if (s.tried) s.check('count', true);
    });
    count.addEventListener('blur', function () {
      if (count.value !== '') count.value = String(parseInt(count.value, 10));
      if (s.tried || count.value === '0') s.check('count', true);
    });

    const memo = $(p + 'memo') as HTMLTextAreaElement;
    memo.addEventListener('input', function () {
      if (memo.value.length > 500) memo.value = memo.value.slice(0, 500);
      $(p + 'memo-count').textContent = memo.value.length + '/500';
    });
    memo.addEventListener('blur', function () {
      memo.value = memo.value.trim();
      $(p + 'memo-count').textContent = memo.value.length + '/500';
    });

    return s;
  }

  /* ========== 팝업 공통 ========== */
  function openDlg(id: string) {
    const d = $(id) as HTMLDialogElement;
    if (d.showModal) d.showModal();
    else d.setAttribute('open', '');
  }
  function closeDlg(id: string) {
    const d = $(id) as HTMLDialogElement;
    if (d.close) d.close();
    else d.removeAttribute('open');
  }

  let onConfirm: (() => Promise<void>) | null = null;
  function askDelete(text: string, fn: () => Promise<void>) {
    $('del-t').textContent = text;
    $('del-yes').textContent = '삭제';
    onConfirm = fn;
    openDlg('dlg-del');
  }
  function askConfirm(text: string, btnText: string, fn: () => Promise<void>) {
    $('del-t').textContent = text;
    $('del-yes').textContent = btnText;
    onConfirm = fn;
    openDlg('dlg-del');
  }
  $('del-no').addEventListener('click', function () {
    onConfirm = null;
    closeDlg('dlg-del');
  });
  $('del-yes').addEventListener('click', async function () {
    const fn = onConfirm;
    onConfirm = null;
    closeDlg('dlg-del');
    if (fn) await fn();
  });

  /* ========== 화면 라우팅 및 뷰 전환 (/ vs /admin) ========== */
  const views = { form: $('v-form'), done: $('v-done'), admin: $('v-admin') };

  function isPathAdmin(): boolean {
    const p = window.location.pathname.toLowerCase();
    const s = window.location.search.toLowerCase();
    const h = window.location.hash.toLowerCase();
    return p === '/admin' || p.startsWith('/admin/') || s.includes('admin') || h === '#admin';
  }

  function navigateTo(path: string) {
    if (window.location.pathname !== path) {
      window.history.pushState(null, '', path);
    }
    syncRoute();
  }

  window.addEventListener('popstate', function () {
    syncRoute();
  });

  function syncRoute() {
    const isAdmin = isPathAdmin();
    const adminBadge = document.getElementById('admin-badge');
    const btnBackForm = document.getElementById('btn-back-form');
    const siteFooter = document.getElementById('site-footer');
    const adminGate = document.getElementById('admin-gate');
    const adminDashboard = document.getElementById('admin-dashboard-body');

    if (isAdmin) {
      // 주최자 관리자 페이지 (/admin)
      if (adminBadge) adminBadge.style.display = 'inline-block';
      if (btnBackForm) btnBackForm.style.display = 'inline-flex';
      if (siteFooter) siteFooter.style.display = 'none';

      views.form.hidden = true;
      views.done.hidden = true;
      views.admin.hidden = false;

      if (isCurrentAdmin) {
        if (adminGate) adminGate.hidden = true;
        if (adminDashboard) adminDashboard.hidden = false;
        renderAdmin();
        startLiveSubscription();
        loadAndRenderSheetsConfig();
      } else {
        if (adminGate) adminGate.hidden = false;
        if (adminDashboard) adminDashboard.hidden = true;
        if (unsubscribeLive) {
          unsubscribeLive();
          unsubscribeLive = null;
        }
      }
    } else {
      // 일반 참가 신청 페이지 (/)
      if (adminBadge) adminBadge.style.display = 'none';
      if (btnBackForm) btnBackForm.style.display = 'none';
      if (siteFooter) siteFooter.style.display = 'block';

      views.admin.hidden = true;
      if (unsubscribeLive) {
        unsubscribeLive();
        unsubscribeLive = null;
      }

      if (!views.done.hidden) {
        views.form.hidden = true;
      } else {
        views.form.hidden = false;
      }
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function showView(v: 'form' | 'done' | 'admin') {
    if (v === 'admin') {
      navigateTo('/admin');
      return;
    }

    if (isPathAdmin()) {
      window.history.pushState(null, '', '/');
    }

    views.admin.hidden = true;
    views.form.hidden = v !== 'form';
    views.done.hidden = v !== 'done';
    syncRoute();
  }

  // Google Login for Admin
  async function handleGoogleLogin(msgElementId: string) {
    const msgEl = document.getElementById(msgElementId);
    if (msgEl) {
      msgEl.textContent = '';
      msgEl.className = 'msg';
    }
    try {
      toast('Google 계정 인증 중...');
      const { user } = await signInAdminWithGoogle();
      const email = user.email?.toLowerCase();
      if (email === ADMIN_EMAIL.toLowerCase()) {
        isCurrentAdmin = true;
        closeDlg('dlg-pw');
        toast(`관리자(${user.displayName || user.email})로 로그인되었습니다.`);
        syncRoute();
        loadAndRenderSheetsConfig();
      } else {
        if (msgEl) {
          msgEl.textContent = `권한이 없습니다. 등록된 관리자(${ADMIN_EMAIL}) 계정으로 로그인해 주세요.`;
          msgEl.className = 'msg bad';
        }
        toast('관리자 권한이 없는 계정입니다.');
      }
    } catch (err: any) {
      console.error('Google Sign-in Error:', err);
      if (err.code !== 'auth/popup-closed-by-user' && msgEl) {
        msgEl.textContent = '로그인 중 오류가 발생했습니다: ' + (err.message || '인증 실패');
        msgEl.className = 'msg bad';
      }
    }
  }

  /* ========== 구글 스프레드시트 연동 상태 및 로직 ========== */
  let currentSheetsConfig: GoogleSheetsConfig | null = null;

  async function ensureAccessToken(): Promise<string> {
    let token = getCachedAccessToken();
    if (!token) {
      toast('구글 스프레드시트 작업을 위해 Google 계정 인증이 필요합니다.');
      const res = await signInAdminWithGoogle();
      token = res.accessToken;
      if (!token) {
        throw new Error('Google Workspace 접근 권한 토큰을 획득하지 못했습니다.');
      }
    }
    return token;
  }

  function renderSheetsStatus() {
    const unlinkedBox = document.getElementById('sheets-unlinked-box');
    const linkedBox = document.getElementById('sheets-linked-box');
    const statusBadge = document.getElementById('sheets-status-badge');
    const titleEl = document.getElementById('linked-sheet-title');
    const urlEl = document.getElementById('linked-sheet-url') as HTMLAnchorElement | null;
    const msgEl = document.getElementById('sheets-sync-status-msg');

    if (!unlinkedBox || !linkedBox || !statusBadge) return;

    if (currentSheetsConfig && currentSheetsConfig.spreadsheetId) {
      unlinkedBox.style.display = 'none';
      linkedBox.style.display = 'block';

      statusBadge.textContent = '연동 활성화 (실시간)';
      statusBadge.style.background = '#dcfce7';
      statusBadge.style.color = '#15803d';

      if (titleEl) titleEl.textContent = currentSheetsConfig.sheetTitle || '이벤트 참가 신청 명단';
      if (urlEl) {
        urlEl.href = currentSheetsConfig.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${currentSheetsConfig.spreadsheetId}/edit`;
      }
    } else {
      unlinkedBox.style.display = 'flex';
      linkedBox.style.display = 'none';

      statusBadge.textContent = '미연동';
      statusBadge.style.background = '#f1f5f9';
      statusBadge.style.color = '#64748b';
      if (msgEl) msgEl.textContent = '';
    }
  }

  async function loadAndRenderSheetsConfig() {
    if (!isCurrentAdmin) return;
    try {
      currentSheetsConfig = await getSheetsConfig();
      renderSheetsStatus();
    } catch (e) {
      console.warn('Failed to load Google Sheets config:', e);
    }
  }

  // Google Sheets buttons
  const btnCreateSheet = document.getElementById('btn-create-sheet');
  if (btnCreateSheet) {
    btnCreateSheet.addEventListener('click', function () {
      askConfirm('관리자 Google 드라이브에 새 스프레드시트를 생성하고 참가 신청 데이터를 연동할까요?', '생성 및 연동', async function () {
        try {
          toast('Google 계정 인증 확인 중...');
          const token = await ensureAccessToken();

          toast('구글 스프레드시트 생성 중...');
          const nowStr = new Date().toISOString().slice(0, 10);
          const sheetTitle = `이벤트 참가 신청 명단 (${nowStr})`;
          const created = await createEventSpreadsheet(token, sheetTitle);

          if (memoryRecords.length > 0) {
            toast('신청 데이터를 시트에 기록 중...');
            await syncAllRegistrationsToSpreadsheet(token, created.spreadsheetId, memoryRecords);
          }

          const newConfig: GoogleSheetsConfig = {
            spreadsheetId: created.spreadsheetId,
            spreadsheetUrl: created.spreadsheetUrl,
            sheetTitle: created.title,
            autoSync: true,
            updatedAt: new Date().toISOString(),
          };

          await saveSheetsConfig(newConfig);
          currentSheetsConfig = newConfig;
          renderSheetsStatus();

          const msgEl = document.getElementById('sheets-sync-status-msg');
          if (msgEl) {
            msgEl.textContent = `시트 생성 완료! (${memoryRecords.length}건 동기화됨)`;
          }
          toast('구글 스프레드시트가 성공적으로 생성 및 연동되었습니다!');
        } catch (err: any) {
          console.error('Create sheet error:', err);
          toast('스프레드시트 생성 실패: ' + (err.message || ''));
        }
      });
    });
  }

  const btnSyncAllSheets = document.getElementById('btn-sync-all-sheets');
  if (btnSyncAllSheets) {
    btnSyncAllSheets.addEventListener('click', function () {
      if (!currentSheetsConfig?.spreadsheetId) {
        toast('연동된 구글 스프레드시트가 없습니다.');
        return;
      }

      askConfirm(
        `현재 저장된 신청 데이터 ${memoryRecords.length}건을 구글 스프레드시트에 동기화(업데이트)할까요?`,
        '동기화 실행',
        async function () {
          try {
            toast('Google 계정 인증 확인 중...');
            const token = await ensureAccessToken();

            toast('구글 스프레드시트에 전체 데이터 동기화 중...');
            const count = await syncAllRegistrationsToSpreadsheet(
              token,
              currentSheetsConfig!.spreadsheetId,
              memoryRecords
            );

            const msgEl = document.getElementById('sheets-sync-status-msg');
            if (msgEl) {
              msgEl.textContent = `최근 전체 동기화: ${new Date().toLocaleTimeString('ko-KR')} (${count}건 완료)`;
            }
            toast(`구글 스프레드시트에 ${count}건 동기화 완료!`);
          } catch (err: any) {
            console.error('Sync all to sheets error:', err);
            toast('스프레드시트 동기화 실패: ' + (err.message || ''));
          }
        }
      );
    });
  }

  const btnUnlinkSheet = document.getElementById('btn-unlink-sheet');
  if (btnUnlinkSheet) {
    btnUnlinkSheet.addEventListener('click', function () {
      askConfirm(
        '구글 스프레드시트 연동을 해제할까요? (스프레드시트 파일 자체는 구글 드라이브에 보존됩니다)',
        '연동 해제',
        async function () {
          try {
            await saveSheetsConfig(null);
            currentSheetsConfig = null;
            renderSheetsStatus();
            toast('구글 스프레드시트 연동이 해제되었습니다.');
          } catch (err: any) {
            console.error('Unlink sheet error:', err);
            toast('연동 해제 실패: ' + (err.message || ''));
          }
        }
      );
    });
  }

  const btnOpenConnectDlg = document.getElementById('btn-open-connect-dlg');
  if (btnOpenConnectDlg) {
    btnOpenConnectDlg.addEventListener('click', function () {
      const inp = document.getElementById('cs-input') as HTMLInputElement;
      const msg = document.getElementById('cs-msg');
      if (inp) inp.value = '';
      if (msg) msg.textContent = '';
      openDlg('dlg-connect-sheet');
    });
  }

  const csCancel = document.getElementById('cs-cancel');
  if (csCancel) {
    csCancel.addEventListener('click', function () {
      closeDlg('dlg-connect-sheet');
    });
  }

  const csSubmit = document.getElementById('cs-submit');
  if (csSubmit) {
    csSubmit.addEventListener('click', async function () {
      const inp = document.getElementById('cs-input') as HTMLInputElement;
      const msg = document.getElementById('cs-msg');
      if (!inp || !msg) return;

      const rawVal = inp.value.trim();
      if (!rawVal) {
        msg.textContent = '스프레드시트 URL 또는 ID를 입력해 주세요.';
        return;
      }

      let sheetId = rawVal;
      const urlMatch = rawVal.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (urlMatch && urlMatch[1]) {
        sheetId = urlMatch[1];
      }

      try {
        msg.textContent = '스프레드시트 권한 확인 중...';
        const token = await ensureAccessToken();

        const info = await getSpreadsheetInfo(token, sheetId);

        const newConfig: GoogleSheetsConfig = {
          spreadsheetId: sheetId,
          spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${sheetId}/edit`,
          sheetTitle: info.title || '연결된 구글 스프레드시트',
          autoSync: true,
          updatedAt: new Date().toISOString(),
        };

        await saveSheetsConfig(newConfig);
        currentSheetsConfig = newConfig;

        toast('시트 헤더 및 데이터 동기화 중...');
        await syncAllRegistrationsToSpreadsheet(token, sheetId, memoryRecords);

        closeDlg('dlg-connect-sheet');
        renderSheetsStatus();
        toast('기존 구글 스프레드시트가 성공적으로 연결되었습니다!');
      } catch (err: any) {
        console.error('Connect sheet error:', err);
        msg.textContent = '연결 실패: ' + (err.message || '스프레드시트에 접근할 수 없습니다.');
      }
    });
  }

  const adminGateLoginBtn = document.getElementById('admin-gate-login-btn');
  if (adminGateLoginBtn) {
    adminGateLoginBtn.addEventListener('click', function () {
      handleGoogleLogin('admin-gate-msg');
    });
  }

  const pwGoogleBtn = document.getElementById('pw-google-btn');
  if (pwGoogleBtn) {
    pwGoogleBtn.addEventListener('click', function () {
      handleGoogleLogin('pw-msg');
    });
  }

  const pwNoBtn = document.getElementById('pw-no');
  if (pwNoBtn) {
    pwNoBtn.addEventListener('click', function () {
      closeDlg('dlg-pw');
    });
  }

  // Navigation link events
  const linkAdmin = document.getElementById('link-admin');
  if (linkAdmin) {
    linkAdmin.addEventListener('click', function (e) {
      e.preventDefault();
      navigateTo('/admin');
    });
  }

  const btnBackForm = document.getElementById('btn-back-form');
  if (btnBackForm) {
    btnBackForm.addEventListener('click', function (e) {
      e.preventDefault();
      navigateTo('/');
    });
  }

  const brandLink = document.getElementById('brand-link');
  if (brandLink) {
    brandLink.addEventListener('click', function (e) {
      e.preventDefault();
      navigateTo('/');
    });
  }

  document.querySelectorAll('.nav-link-btn').forEach(function (el) {
    el.addEventListener('click', function (e) {
      e.preventDefault();
      navigateTo('/');
    });
  });

  // Header auth badge
  const navActions = document.getElementById('nav-actions') || document.querySelector('header.top');
  const authBadge = document.createElement('div');
  authBadge.id = 'auth-status-badge';
  authBadge.style.cssText = 'display:inline-flex;align-items:center;gap:8px;font-size:12px;color:#475569;';
  if (navActions) {
    navActions.appendChild(authBadge);
  }

  listenAuthState(function (user, isAdminUser) {
    currentAdminUser = user;
    isCurrentAdmin = isAdminUser;
    if (isAdminUser) {
      authBadge.innerHTML = `
        <span style="display:inline-flex;align-items:center;gap:4px;padding:3px 8px;background:#eff6ff;color:#1d4ed8;border-radius:12px;font-weight:600;">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
          ${user?.email} (주최자 관리자)
        </span>
        <button type="button" id="btn-admin-logout" style="background:none;border:none;color:#94a3b8;cursor:pointer;text-decoration:underline;">로그아웃</button>
      `;
      const logoutBtn = $('btn-admin-logout');
      logoutBtn.onclick = async function () {
        if (unsubscribeLive) {
          unsubscribeLive();
          unsubscribeLive = null;
        }
        await signOutAdmin();
        isCurrentAdmin = false;
        toast('로그아웃되었습니다.');
        syncRoute();
      };
      syncRoute();
    } else if (user) {
      authBadge.innerHTML = `
        <span style="padding:3px 8px;background:#fef2f2;color:#b91c1c;border-radius:12px;font-size:11px;">일반 계정 (${user.email})</span>
        <button type="button" id="btn-admin-logout" style="background:none;border:none;color:#94a3b8;cursor:pointer;text-decoration:underline;">로그아웃</button>
      `;
      const logoutBtn = $('btn-admin-logout');
      logoutBtn.onclick = async function () {
        if (unsubscribeLive) {
          unsubscribeLive();
          unsubscribeLive = null;
        }
        await signOutAdmin();
        isCurrentAdmin = false;
        toast('로그아웃되었습니다.');
        syncRoute();
      };
      syncRoute();
    } else {
      authBadge.innerHTML = ``;
      syncRoute();
    }
  });

  let toastTimer: any;
  function toast(text: string) {
    const t = $('toast');
    t.textContent = text;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      t.hidden = true;
    }, 2500);
  }

  /* ========== 실시간 데이터 리스너 ========== */
  function startLiveSubscription() {
    if (!isCurrentAdmin) {
      console.log('Admin not authenticated. Skipping registration subscription.');
      return;
    }

    if (unsubscribeLive) {
      unsubscribeLive();
      unsubscribeLive = null;
    }

    try {
      unsubscribeLive = subscribeToRegistrations(
        function (records) {
          memoryRecords = records;
          saveLocalCache(records);
          renderAdmin();
        },
        async function (error) {
          console.warn('Live subscription error:', error);
          if (!isCurrentAdmin) return;
          try {
            const records = await fetchAllRegistrations();
            memoryRecords = records;
            saveLocalCache(records);
            renderAdmin();
          } catch (e) {
            renderAdmin();
          }
        }
      );
    } catch (e) {
      console.warn('Realtime subscription failed:', e);
      renderAdmin();
    }
  }

  /* ========== 신청서 (Create) ========== */
  const F = Scope('f-', true);
  let busy = false;
  const cAll = $('c-all') as HTMLInputElement;
  const cAgree = $('c-agree') as HTMLInputElement;
  const cMkt = $('c-mkt') as HTMLInputElement;

  function checkAgree(show: boolean) {
    const m = cAgree.checked ? '' : MSG.agree;
    if (show) {
      $('c-agree-msg').textContent = m;
      $('consents').classList.toggle('bad', !!m);
    }
    return m;
  }
  function syncAll() {
    cAll.checked = cAgree.checked && cMkt.checked;
  }
  cAll.addEventListener('change', function () {
    cAgree.checked = cAll.checked;
    cMkt.checked = cAll.checked;
    if (F.tried) checkAgree(true);
  });
  cAgree.addEventListener('change', function () {
    syncAll();
    if (F.tried) checkAgree(true);
  });
  cMkt.addEventListener('change', syncAll);

  function step(d: number) {
    const c = $('f-count') as HTMLInputElement;
    const n = parseInt(c.value, 10) || 0;
    c.value = String(Math.min(99, Math.max(1, (n || 1) + d)));
    F.check('count', true);
  }
  $('f-minus').addEventListener('click', function () {
    step(-1);
  });
  $('f-plus').addEventListener('click', function () {
    step(1);
  });

  async function submitForm() {
    if (busy) return;
    F.tried = true;
    F.clean();
    const first = F.checkAll();
    if (checkAgree(true) && !first) {
      // agree is first
      $('consents').scrollIntoView({ behavior: 'smooth', block: 'center' });
      return;
    }
    if (first) {
      const target = F.el(first);
      target.scrollIntoView({ behavior: 'smooth', block: 'center' });
      target.focus({ preventScroll: true });
      return;
    }

    const v = F.values();
    const btn = $('submit') as HTMLButtonElement;
    busy = true;
    btn.disabled = true;
    btn.textContent = '중복 확인 중...';

    try {
      // 파이어베이스 phone_index에서 중복 여부 확인
      const isTaken = await checkPhoneTaken(v.phone);
      if (isTaken) {
        openDlg('dlg-dup');
        busy = false;
        btn.disabled = false;
        btn.textContent = '신청하기';
        return;
      }

      btn.textContent = '파이어베이스 저장 중...';
      const nextSeq = memoryRecords.length + 1;
      const createdRecord = await createRegistration(
        {
          name: v.name,
          phone: v.phone,
          email: v.email,
          count: v.count,
          memo: v.memo,
          mkt: cMkt.checked,
          agree: true,
        },
        nextSeq
      );

      // 로컬 메모리에도 추가
      memoryRecords.unshift(createdRecord);
      saveLocalCache(memoryRecords);

      showDone(createdRecord);
      toast('신청이 파이어베이스에 정상 등록되었습니다.');

      // 구글 스프레드시트 연동 중이면 자동 행 추가
      if (currentSheetsConfig?.spreadsheetId) {
        const token = getCachedAccessToken();
        if (token) {
          appendRegistrationToSpreadsheet(token, currentSheetsConfig.spreadsheetId, createdRecord).catch(e => {
            console.warn('Auto sync to Google Sheet error:', e);
          });
        }
      }
    } catch (err: any) {
      console.error('Registration submission failed:', err);
      toast('접수 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.');
    } finally {
      busy = false;
      btn.disabled = false;
      btn.textContent = '신청하기';
    }
  }

  $('submit').addEventListener('click', submitForm);
  $('form').addEventListener('keydown', function (e) {
    const el = e.target as HTMLElement;
    if (e.key === 'Enter' && el.tagName === 'INPUT' && (el as HTMLInputElement).type !== 'checkbox') {
      e.preventDefault();
      submitForm();
    }
  });

  function sumRow(k: string, v: string) {
    const r = document.createElement('div');
    r.className = 'r';
    const a = document.createElement('span');
    a.className = 'k';
    a.textContent = k;
    const b = document.createElement('span');
    b.className = 'v';
    b.textContent = v;
    r.appendChild(a);
    r.appendChild(b);
    return r;
  }

  function showDone(v: any) {
    $('d-email').textContent = v.email;
    const s = $('sum');
    s.innerHTML = '';
    s.appendChild(sumRow('이름', v.name));
    s.appendChild(sumRow('휴대폰 번호', v.phone));
    s.appendChild(sumRow('이메일', v.email));
    s.appendChild(sumRow('참석 인원', v.count + '명'));
    if (v.memo) s.appendChild(sumRow('전달 사항', v.memo));
    $('st1').classList.remove('on');
    $('st2').classList.add('on');
    showView('done');
  }

  $('again').addEventListener('click', function () {
    ['f-name', 'f-phone', 'f-email', 'f-memo'].forEach(function (id) {
      ($(id) as HTMLInputElement).value = '';
    });
    ($('f-count') as HTMLInputElement).value = '1';
    F.refreshCount();
    cAll.checked = cAgree.checked = cMkt.checked = false;
    F.clearMsgs();
    checkAgree(false);
    $('c-agree-msg').textContent = '';
    $('consents').classList.remove('bad');
    $('st2').classList.remove('on');
    $('st1').classList.add('on');
    showView('form');
  });

  $('dup-ok').addEventListener('click', function () {
    closeDlg('dlg-dup');
    F.el('phone').focus();
  });

  /* ========== 관리자 대시보드 로직 ========== */
  function maskName(n: string) {
    const a = Array.from(n || '');
    if (a.length <= 1) return n || '';
    if (a.length === 2) return a[0] + '*';
    return a[0] + new Array(a.length - 1).join('*') + a[a.length - 1];
  }
  function maskPhone(p: string) {
    const m = /^(\d{3})-(\d{3,4})-(\d{4})$/.exec(p || '');
    return m ? m[1] + '-' + new Array(m[2].length + 1).join('*') + '-' + m[3] : p ? '***' : '';
  }
  function maskEmail(e: string) {
    const i = (e || '').indexOf('@');
    if (i < 1) return e ? '***' : '';
    const id = e.slice(0, i);
    const dom = e.slice(i);
    const keep = id.length <= 2 ? 1 : 2;
    return id.slice(0, keep) + new Array(Math.max(1, id.length - keep) + 1).join('*') + dom;
  }
  function maskMemo(s: string) {
    return (s || '')
      .replace(/[A-Za-z0-9._+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]+/g, '***')
      .replace(/\d[\d\-]{2,}\d/g, '***');
  }

  const H9 = 9 * 3600 * 1000;
  function shifted(ms: number) {
    return new Date(ms + H9);
  }
  function fmt(iso: string) {
    if (!iso) return '-';
    return shifted(Date.parse(iso)).toISOString().slice(0, 16).replace('T', ' ');
  }
  function monthsAgo(k: number) {
    const n = shifted(Date.now());
    return Date.UTC(n.getUTCFullYear(), n.getUTCMonth() - k, n.getUTCDate(), n.getUTCHours(), n.getUTCMinutes()) - H9;
  }
  const PAGE = 30;
  let page = 1;

  function dayStart(s: string) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) - H9 : null;
  }

  function baseRows() {
    let rows = memoryRecords.slice();
    const per = ($('a-period') as HTMLSelectElement).value;
    const msg = $('a-range-msg');
    msg.textContent = '';
    if (per === '6m') {
      const cutoff = monthsAgo(6);
      rows = rows.filter(function (x) {
        return Date.parse(x.createdAt) >= cutoff;
      });
    } else if (per === 'custom') {
      const f = dayStart(($('a-from') as HTMLInputElement).value);
      const e = dayStart(($('a-to') as HTMLInputElement).value);
      if (f !== null && e !== null && f > e) {
        msg.textContent = '시작일이 종료일보다 늦습니다.';
        return [];
      }
      rows = rows.filter(function (x) {
        const ts = Date.parse(x.createdAt);
        return (f === null || ts >= f) && (e === null || ts < e + 86400000);
      });
    }
    return rows;
  }

  function filteredRows() {
    const q = ($('a-q') as HTMLInputElement).value.trim().toLowerCase();
    return baseRows()
      .filter(function (x) {
        return !q || [x.name, x.phone, x.email, x.memo].join(' ').toLowerCase().indexOf(q) > -1;
      })
      .sort(function (a, b) {
        return Date.parse(b.createdAt) - Date.parse(a.createdAt);
      });
  }

  function renderAdmin() {
    renderSheetsStatus();
    const all = memoryRecords;
    const base = baseRows();
    const q = ($('a-q') as HTMLInputElement).value.trim().toLowerCase();
    const rows = filteredRows();
    $('a-range').hidden = ($('a-period') as HTMLSelectElement).value !== 'custom';

    let people = 0;
    let mkt = 0;
    base.forEach(function (x) {
      people += x.count || 0;
      if (x.mkt) mkt++;
    });
    $('s-count').textContent = String(base.length);
    $('s-people').textContent = String(people);
    $('s-mkt').textContent = String(mkt);

    /* 최근 6개월을 한 달씩 6칸으로 나눈 그래프 */
    const now = Date.now();
    const recent = all.filter(function (x) {
      return Date.parse(x.createdAt) >= monthsAgo(6);
    });
    const buckets = [];
    let max = 1;
    for (let i = 0; i < 6; i++) {
      const from = monthsAgo(6 - i);
      const to = monthsAgo(5 - i);
      const last = i === 5;
      const c = recent.filter(function (x) {
        const t = Date.parse(x.createdAt);
        return t >= from && (last ? t <= now + 1000 : t < to);
      }).length;
      const f = shifted(from);
      const e = shifted(last ? now : to - 86400000);
      buckets.push({
        n: c,
        label: e.getUTCMonth() + 1 + '월',
        range: f.getUTCMonth() + 1 + '/' + f.getUTCDate() + '~' + (e.getUTCMonth() + 1) + '/' + e.getUTCDate(),
      });
      if (c > max) max = c;
    }
    const chart = $('chart');
    chart.innerHTML = '';
    buckets.forEach(function (b) {
      const d = document.createElement('div');
      d.className = 'bar';
      d.title = b.range + ' ' + b.n + '건';
      const a = document.createElement('div');
      a.className = 'n';
      a.textContent = String(b.n);
      const bar = document.createElement('div');
      bar.className = 'b';
      bar.style.height = Math.round((b.n / max) * 18) + 'px';
      const l = document.createElement('div');
      l.className = 'l';
      l.textContent = b.label;
      d.appendChild(a);
      d.appendChild(bar);
      d.appendChild(l);
      chart.appendChild(d);
    });

    /* 명단 렌더링 */
    const pages = Math.max(1, Math.ceil(rows.length / PAGE));
    if (page > pages) page = pages;
    const body = $('rows');
    body.innerHTML = '';
    rows.slice((page - 1) * PAGE, page * PAGE).forEach(function (x, idx) {
      const tr = document.createElement('tr');
      function cell(text: string, cls?: string) {
        const td = document.createElement('td');
        if (cls) td.className = cls;
        td.textContent = text;
        tr.appendChild(td);
        return td;
      }
      const displayId = String(x.id || (rows.length - ((page - 1) * PAGE + idx)));
      cell(displayId, 'n');
      cell(fmt(x.createdAt), 't');
      cell(maskName(x.name));
      cell(maskPhone(x.phone), 'c');
      cell(maskEmail(x.email));
      cell(x.count + '명', 'c');
      const ag = cell(x.agree ? '동의' : '-', 'c');
      if (x.agreedAt) ag.title = fmt(x.agreedAt) + ' 동의';
      cell(x.mkt ? '동의' : '미동의', 'c');
      cell(maskMemo(x.memo) || '-', 'm');
      cell(x.updatedAt ? fmt(x.updatedAt) : '-', 't');

      const act = cell('', 'a');
      const b1 = document.createElement('button');
      b1.type = 'button';
      b1.className = 'rbtn';
      b1.textContent = '수정';
      b1.setAttribute('data-act', 'edit');
      b1.setAttribute('data-id', x.docId);
      b1.setAttribute('aria-label', '신청 ' + x.docId + ' 수정');

      const b2 = document.createElement('button');
      b2.type = 'button';
      b2.className = 'rbtn';
      b2.textContent = '삭제';
      b2.setAttribute('data-act', 'del');
      b2.setAttribute('data-id', x.docId);
      b2.setAttribute('aria-label', '신청 ' + x.docId + ' 삭제');

      act.appendChild(b1);
      act.appendChild(b2);
      body.appendChild(tr);
    });

    const empty = $('empty');
    if (rows.length) {
      empty.hidden = true;
    } else {
      empty.hidden = false;
      empty.textContent = !all.length
        ? '아직 신청한 사람이 없습니다 (파이어베이스 실시간 연동 중).'
        : q
        ? '검색 결과가 없습니다.'
        : '최근 6개월 안에 신청한 사람이 없습니다.';
    }
    $('a-cnt').textContent = rows.length + '건';
    ($('a-clear') as HTMLButtonElement).disabled = all.length === 0;
    ($('a-csv') as HTMLButtonElement).textContent = 'CSV 다운로드 (' + rows.length + '건)';
    ($('a-csv') as HTMLButtonElement).disabled = rows.length === 0;
    $('pinfo').textContent = page + ' / ' + pages;
    ($('prev') as HTMLButtonElement).disabled = page <= 1;
    ($('next') as HTMLButtonElement).disabled = page >= pages;
  }

  $('a-period').addEventListener('change', function () {
    page = 1;
    renderAdmin();
  });
  ['a-from', 'a-to'].forEach(function (id) {
    $(id).addEventListener('change', function () {
      page = 1;
      renderAdmin();
    });
  });

  /* CSV 다운로드 */
  function csvCell(v: any) {
    const s = String(v == null ? '' : v);
    if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
    return '"' + s.replace(/"/g, '""') + '"';
  }
  function buildCsv(rows: RegistrationRecord[]) {
    const lines = [
      ['신청 번호', '이름', '참석 인원', '전달 사항', '이벤트 소식 수신 동의'].map(csvCell).join(','),
    ];
    rows.forEach(function (x, idx) {
      lines.push(
        [x.id || idx + 1, x.name, x.count, x.memo, x.mkt ? '동의' : '미동의']
          .map(csvCell)
          .join(',')
      );
    });
    return lines.join('\r\n');
  }
  function downloadCsv() {
    const rows = filteredRows();
    if (!rows.length) return;
    const text = buildCsv(rows);
    const blob = new Blob(['\uFEFF' + text], { type: 'text/csv;charset=utf-8' });
    const d = shifted(Date.now()).toISOString().slice(2, 10).replace(/-/g, '');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = d + '_신청명단_홍지영.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
    }, 2000);
    toast(rows.length + '건을 내려받았습니다.');
  }
  $('a-csv').addEventListener('click', downloadCsv);
  $('a-q').addEventListener('input', function () {
    page = 1;
    renderAdmin();
  });
  $('prev').addEventListener('click', function () {
    page--;
    renderAdmin();
  });
  $('next').addEventListener('click', function () {
    page++;
    renderAdmin();
  });

  /* 추가 (Create), 수정 (Update) */
  const E = Scope('e-', false);
  let editingDocId: string | null = null;

  function openEdit(docId: string | null) {
    editingDocId = docId;
    const rec = docId == null ? null : memoryRecords.find(x => x.docId === docId);
    $('e-title').textContent = rec ? '신청 수정 (ID ' + rec.id + ')' : '신청 추가';
    E.el('name').value = rec ? rec.name : '';
    E.el('phone').value = rec ? rec.phone : '';
    E.el('email').value = rec ? rec.email : '';
    E.el('count').value = rec ? String(rec.count) : '1';
    ($('e-memo') as HTMLTextAreaElement).value = rec ? rec.memo : '';
    ($('e-mkt') as HTMLInputElement).checked = rec ? !!rec.mkt : false;
    ($('e-agree') as HTMLInputElement).checked = rec ? !!rec.agree : false;
    $('e-agree-msg').textContent = '';
    E.clearMsgs();
    E.refreshCount();
    openDlg('dlg-edit');
    E.el('name').focus();
  }

  async function saveEdit() {
    E.tried = true;
    E.clean();
    let first: string | null = E.checkAll();
    const v = E.values();
    const mkt = ($('e-mkt') as HTMLInputElement).checked;
    const agree = ($('e-agree') as HTMLInputElement).checked;
    $('e-agree-msg').textContent = '';

    const saveBtn = $('e-save') as HTMLButtonElement;
    saveBtn.disabled = true;
    saveBtn.textContent = '저장 중...';

    try {
      if (!first && (await checkPhoneTaken(v.phone, editingDocId))) {
        E.setMsg('phone', MSG.dup);
        E.setBad('phone', true);
        first = 'phone';
      }
      if (!agree) {
        $('e-agree-msg').textContent = '개인정보 수집, 이용에 동의해 주세요.';
        if (!first) ($('e-agree') as HTMLInputElement).focus();
      }
      if (first) {
        E.el(first as any).focus();
        return;
      }
      if (!agree) return;

      if (editingDocId == null) {
        // 새 신청 추가
        const nextSeq = memoryRecords.length + 1;
        const newRecord = await createRegistration(
          {
            name: v.name,
            phone: v.phone,
            email: v.email,
            count: v.count,
            memo: v.memo,
            mkt: mkt,
            agree: agree,
          },
          nextSeq
        );
        memoryRecords.unshift(newRecord);
        saveLocalCache(memoryRecords);
        toast('신청을 파이어베이스에 추가했습니다.');
      } else {
        // 기존 신청 수정
        const oldRecord = memoryRecords.find(x => x.docId === editingDocId);
        if (oldRecord) {
          await updateRegistration(
            editingDocId,
            {
              name: v.name,
              phone: v.phone,
              email: v.email,
              count: v.count,
              memo: v.memo,
              mkt: mkt,
              agree: agree,
            },
            oldRecord
          );
          // 메모리 업데이트
          oldRecord.name = v.name;
          oldRecord.phone = v.phone;
          oldRecord.email = v.email;
          oldRecord.count = v.count;
          oldRecord.memo = v.memo;
          oldRecord.mkt = mkt;
          oldRecord.agree = agree;
          oldRecord.updatedAt = new Date().toISOString();
          saveLocalCache(memoryRecords);
          toast('신청을 파이어베이스에 수정했습니다.');
        }
      }
      closeDlg('dlg-edit');
      renderAdmin();

      // 구글 스프레드시트 연동 중이면 변경된 데이터 시트 동기화
      if (currentSheetsConfig?.spreadsheetId) {
        const token = getCachedAccessToken();
        if (token) {
          syncAllRegistrationsToSpreadsheet(token, currentSheetsConfig.spreadsheetId, memoryRecords).catch(e => {
            console.warn('Auto sync on edit error:', e);
          });
        }
      }
    } catch (err: any) {
      console.error('Save edit failed:', err);
      toast('저장 중 오류가 발생했습니다: ' + (err.message || ''));
    } finally {
      saveBtn.disabled = false;
      saveBtn.textContent = '저장';
    }
  }

  $('a-add').addEventListener('click', function () {
    openEdit(null);
  });
  $('e-save').addEventListener('click', saveEdit);
  $('e-cancel').addEventListener('click', function () {
    closeDlg('dlg-edit');
  });
  $('dlg-edit').addEventListener('keydown', function (e) {
    const el = e.target as HTMLElement;
    if (e.key === 'Enter' && el.tagName === 'INPUT' && (el as HTMLInputElement).type !== 'checkbox') {
      e.preventDefault();
      saveEdit();
    }
  });

  $('rows').addEventListener('click', function (e) {
    const target = e.target as HTMLElement;
    const b = target.closest ? (target.closest('button[data-act]') as HTMLButtonElement) : null;
    if (!b) return;
    const docId = b.getAttribute('data-id');
    if (!docId) return;
    const rec = memoryRecords.find(x => x.docId === docId);
    if (!rec) return;

    if (b.getAttribute('data-act') === 'edit') {
      openEdit(docId);
    } else {
      askDelete('신청 ID ' + rec.id + ' (' + maskName(rec.name) + ')을 파이어베이스에서 삭제할까요?', async function () {
        try {
          toast('파이어베이스에서 삭제 중...');
          await deleteRegistration(docId, rec.phone);
          memoryRecords = memoryRecords.filter(x => x.docId !== docId);
          saveLocalCache(memoryRecords);
          toast('신청을 삭제했습니다.');
          renderAdmin();

          // 구글 스프레드시트 동기화
          if (currentSheetsConfig?.spreadsheetId) {
            const token = getCachedAccessToken();
            if (token) {
              syncAllRegistrationsToSpreadsheet(token, currentSheetsConfig.spreadsheetId, memoryRecords).catch(e => {
                console.warn('Auto sync on delete error:', e);
              });
            }
          }
        } catch (err: any) {
          console.error(err);
          toast('삭제 실패: ' + (err.message || ''));
        }
      });
    }
  });

  $('a-clear').addEventListener('click', function () {
    const n = memoryRecords.length;
    askDelete('저장된 신청 ' + n + '건을 파이어베이스에서 모두 삭제할까요?', async function () {
      try {
        toast('파이어베이스 전체 데이터 삭제 중...');
        await clearAllRegistrations(memoryRecords);
        memoryRecords = [];
        saveLocalCache(memoryRecords);
        page = 1;
        toast('모든 신청을 삭제했습니다.');
        renderAdmin();

        // 구글 스프레드시트 동기화
        if (currentSheetsConfig?.spreadsheetId) {
          const token = getCachedAccessToken();
          if (token) {
            syncAllRegistrationsToSpreadsheet(token, currentSheetsConfig.spreadsheetId, []).catch(e => {
              console.warn('Auto sync on clear error:', e);
            });
          }
        }
      } catch (err: any) {
        console.error(err);
        toast('전체 삭제 실패: ' + (err.message || ''));
      }
    });
  });

  // 초기 라우트 동기화 (URL 경로 확인하여 / 또는 /admin 화면 렌더링)
  syncRoute();
})();
