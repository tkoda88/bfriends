// ── 設定 ──
const CONFIG = {
  LIFF_ID: '2009966306-JOCQOOvS',
  API_URL: 'https://script.google.com/macros/s/AKfycbzvczSyFTCPTtA4PA69qoiijvyhp-AVx-77VbJJil5GzzMxtKFA4ASLHz__E0apX30/exec',
};

const PAGES = ['schedule', 'teams', 'register', 'mypage', 'docs', 'admin'];

const S = { idToken:'', lineProfileName:'', member:null, schedule:[], isAdmin:false, memberNames:[], memberFurigana:[] };

// ── API呼び出し ──
async function api(action, params = {}) {
  const url = CONFIG.API_URL;
  const payload = JSON.stringify({ action, ...params });
  try {
    const res = await fetch(url, {
      method: 'POST', mode: 'cors', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: payload,
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return await res.json();
  } catch (e) {
    const res2 = await fetch(url + '?payload=' + encodeURIComponent(payload), { method: 'GET', mode: 'cors', redirect: 'follow' });
    if (!res2.ok) throw new Error('HTTP ' + res2.status);
    return await res2.json();
  }
}

// ── ページ切り替え ──
function currentPage() {
  const p = new URLSearchParams(location.search).get('page');
  return PAGES.includes(p) ? p : 'schedule';
}

function showPage(page) {
  PAGES.forEach(p => el('page-' + p).classList.toggle('hidden', p !== page));
}

// ── 初期化 ──
document.addEventListener('DOMContentLoaded', () => {
  el('saveButton').addEventListener('click', doSave);
  el('generateTeamsBtn').addEventListener('click', doGenerateTeams);
  el('exportListBtn').addEventListener('click', doExportList);
  el('saveResultsBtn').addEventListener('click', doSaveResults);
  el('refreshWinRateBtn').addEventListener('click', doRefreshWinRates);
  el('refreshDocsBtn').addEventListener('click', doLoadDocs);
  el('editProfileBtn').addEventListener('click', () => toggleProfileEdit(true));
  el('cancelProfileBtn').addEventListener('click', () => toggleProfileEdit(false));
  el('saveProfileBtn').addEventListener('click', doSaveProfile);
  el('adminRegBtn').addEventListener('click', doAdminRegister);
  el('adminLoadSaturdaysBtn').addEventListener('click', loadSaturdays);
  el('adminRegisterDatesBtn').addEventListener('click', doRegisterDates);
  el('adminGenerateReportBtn').addEventListener('click', doGenerateReport);
  // モーダル
  el('modalNameSearch').addEventListener('input', filterModalNames);
  el('modalLinkBtn').addEventListener('click', doModalLink);
  el('modalToNewBtn').addEventListener('click', () => showModalPane('new'));
  el('modalBackBtn').addEventListener('click', () => showModalPane('link'));
  el('modalNewBtn').addEventListener('click', doModalNew);
  showPage(currentPage());
  initLiff();
});

async function initLiff() {
  try {
    showMsg('LIFF初期化中…', 'info');
    await liff.init({ liffId: CONFIG.LIFF_ID });
    if (!liff.isLoggedIn()) { liff.login(); return; }
    S.idToken = liff.getIDToken();
    if (!S.idToken) { showMsg('IDトークンが取得できませんでした。', 'error'); return; }
    const d = liff.getDecodedIDToken();
    S.lineProfileName = d && d.name ? d.name : '';
    showMsg('LIFF認証済。API呼び出し中…', 'info');
    loadSession();
  } catch(e) {
    showMsg('LIFF初期化エラー: ' + (e.message || JSON.stringify(e)), 'error');
  }
}

async function loadSession() {
  showMsg('読み込み中…', 'info');
  try {
    const r = await api('initializeSession', { idToken: S.idToken });
    if (!r.ok && !r.memberFound) { showMsg(r.message || 'エラー', 'error'); return; }
    S.member = r.member;
    S.schedule = r.schedule || [];
    S.isAdmin = r.isAdmin;

    const name = (r.member && r.member.fullName) || r.lineProfileName || S.lineProfileName || '';
    el('headerName').textContent = name;

    const page = currentPage();

    if (r.memberFound && r.member) {
      if (page !== 'teams') hide('messageCard');
      if (page === 'schedule') {
        renderSchedule(r.schedule || []);
        if (r.isAdmin) { show('adminCard'); loadAdminView(); }
        if (r.isAdmin) el('exportListBtn').classList.remove('hidden');
      } else if (page === 'teams') {
        S.schedule = r.schedule || [];
        populateTeamSelect();
        doRefreshWinRates();
        if (r.isAdmin) el('exportListBtn').classList.remove('hidden');
        showMsg('予定を選んでチーム編成してください。', 'info');
      } else if (page === 'register') {
        if (r.isAdmin) show('adminRegisterPane');
      } else if (page === 'docs') {
        doLoadDocs();
      } else if (page === 'mypage') {
        doLoadMyPage();
      } else if (page === 'admin') {
        doLoadAdminPage();
      }
    } else {
      // 未登録: モーダル表示
      hide('messageCard');
      await loadMemberNames();
      showModal();
    }
  } catch(e) {
    showMsg('APIエラー: ' + (e.message || String(e)), 'error');
  }
}

// ── 初回登録モーダル ──
function showModal() {
  show('modalOverlay');
  showModalPane('link');
}

function hideModal() {
  hide('modalOverlay');
}

function showModalPane(pane) {
  el('modalLinkPane').classList.toggle('hidden', pane !== 'link');
  el('modalNewPane').classList.toggle('hidden', pane !== 'new');
  el('modalTitle').textContent = pane === 'link' ? 'はじめに登録してください' : '新規登録';
  el('modalDesc').textContent  = pane === 'link'
    ? '氏名またはふりがなで検索して名簿と紐付けてください。名簿にない方は新規登録へ進んでください。'
    : '必須項目を入力して登録してください。';
}

function normalizeForSearch(s) { return s.replace(/\s+/g, ''); }
function normalizeNameSpaces(s) { return s.replace(/\s+/, ' ').trim(); }

function filterModalNames() {
  const q = normalizeForSearch(el('modalNameSearch').value.trim());
  const list = el('modalNameList');
  el('modalLinkBtn').disabled = true;
  el('modalLinkBtn').dataset.name = '';
  el('modalSelectedName').textContent = '';
  if (!q || q.length < 2) { hide('modalNameList'); list.innerHTML = ''; return; }
  const filtered = S.memberNames.filter((n, i) =>
    normalizeForSearch(n).includes(q) || (S.memberFurigana[i] && normalizeForSearch(S.memberFurigana[i]).includes(q))
  );
  if (filtered.length === 0) {
    hide('modalNameList');
    el('modalSelectedName').textContent = '該当なし。新規の方は「名簿にない方はこちら」をタップしてください。';
  } else if (filtered.length > 3) {
    hide('modalNameList'); list.innerHTML = '';
    el('modalSelectedName').textContent = filtered.length + '件該当。もう少し絞り込んでください。';
  } else {
    list.innerHTML = '';
    filtered.forEach(n => {
      const d = document.createElement('div');
      d.className = 'name-item';
      d.textContent = n;
      d.addEventListener('click', () => {
        el('modalSelectedName').textContent = '選択: ' + n;
        el('modalLinkBtn').disabled = false;
        el('modalLinkBtn').dataset.name = n;
        list.querySelectorAll('.name-item').forEach(i => i.classList.toggle('selected', i.textContent === n));
      });
      list.appendChild(d);
    });
    show('modalNameList');
  }
}

function checkRequiredFields() {
  const p = (S.myPageData && S.myPageData.profile) || S.member;
  return !!(p && p.furigana && p.gender && p.birthDate && p.mobilePhone && p.address);
}

function redirectToMypageEdit() {
  showMsg('必須項目が未入力です。マイページで情報を入力してください。', 'error');
  history.replaceState(null, '', '?page=mypage');
  showPage('mypage');
  doLoadMyPage().then(() => toggleProfileEdit(true));
}

async function doModalLink() {
  const name = el('modalLinkBtn').dataset.name;
  if (!name) return;
  disableAll(true);
  const r = await api('registerMember', { idToken: S.idToken, fullName: name });
  disableAll(false);
  if (!r.ok) { showMsg(r.message, 'error'); return; }
  S.member = r.member;
  el('headerName').textContent = r.member.fullName || '';
  hideModal();
  if (!checkRequiredFields()) { redirectToMypageEdit(); return; }
  showMsg('登録完了しました。', 'success');
  loadSession();
}

async function doModalNew() {
  const fullName    = normalizeNameSpaces(el('mNewFullName').value);
  const furigana    = normalizeNameSpaces(el('mNewFurigana').value);
  const gender      = el('mNewGender').value;
  const birthDate   = el('mNewBirthDate').value;
  const mobilePhone = el('mNewMobilePhone').value.trim();
  const address     = el('mNewAddress').value.trim();
  if (!fullName || !furigana || !gender || !birthDate || !mobilePhone || !address) {
    showMsg('必須項目をすべて入力してください。', 'error'); return;
  }
  disableAll(true);
  const r = await api('registerNewMember', { idToken: S.idToken, fullName, furigana, gender, birthDate, mobilePhone, address });
  disableAll(false);
  if (!r.ok) { showMsg(r.message, 'error'); return; }
  S.member = r.member;
  el('headerName').textContent = r.member.fullName || '';
  hideModal();
  if (!checkRequiredFields()) { redirectToMypageEdit(); return; }
  showMsg('新規登録が完了しました。', 'success');
  loadSession();
}

// ── 管理ページ ──
async function doLoadAdminPage() {
  if (!S.isAdmin) {
    el('adminAccessMsg').textContent = '管理者のみアクセスできます。';
    return;
  }
  el('adminPageCard').classList.add('hidden');
  show('adminScheduleCard');
  show('adminReportCard');
  // 現在月をデフォルト設定
  const now = new Date();
  const ym = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  el('adminMonthPicker').value = ym;
  el('adminMonthPicker').min = ym;
  el('adminReportMonth').value = ym;
  loadExistingSchedule();
}

async function loadExistingSchedule() {
  const r = await api('getAdminPage', { idToken: S.idToken });
  if (!r.ok) return;
  const root = el('adminExistingSchedule');
  const today = new Date();
  today.setHours(0,0,0,0);
  const future = r.sessions.filter(s => new Date(s.eventDate) >= today);
  if (!future.length) { root.innerHTML = '<p class="muted">登録済みの予定はありません。</p>'; return; }
  root.innerHTML = '<h3>登録済みの予定</h3>' +
    future.sort((a, b) => a.eventDate.localeCompare(b.eventDate)).map(s =>
      '<div class="schedule-row">' +
        '<span>' + esc(s.eventDate) + ' ' + esc(s.title) + '</span>' +
        '<button class="danger small" data-sid="' + esc(s.sessionId) + '">削除</button>' +
      '</div>'
    ).join('');
  root.querySelectorAll('[data-sid]').forEach(btn => {
    btn.addEventListener('click', async () => {
      if (!confirm(btn.previousElementSibling.textContent + ' を削除しますか？')) return;
      disableAll(true);
      const r = await api('deleteScheduleSession', { idToken: S.idToken, sessionId: btn.dataset.sid });
      disableAll(false);
      showMsg(r.message, r.ok ? 'success' : 'error');
      if (r.ok) loadExistingSchedule();
    });
  });
}

function loadSaturdays() {
  const ym = el('adminMonthPicker').value;
  if (!ym) { showMsg('月を選択してください。', 'error'); return; }
  const [y, m] = ym.split('-').map(Number);
  const today = new Date(); today.setHours(0,0,0,0);
  const saturdays = [];
  const d = new Date(y, m - 1, 1);
  while (d.getMonth() === m - 1) {
    if (d.getDay() === 6 && d >= today) saturdays.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }
  const root = el('adminSaturdayList');
  if (!saturdays.length) {
    root.innerHTML = '<p class="muted">対象の土曜日はありません。</p>';
    hide('adminRegisterDatesActions'); return;
  }
  root.innerHTML = saturdays.map(s => {
    const ds = s.getFullYear() + '-' + String(s.getMonth()+1).padStart(2,'0') + '-' + String(s.getDate()).padStart(2,'0');
    return '<label class="saturday-item"><input type="checkbox" value="' + ds + '" checked><span>' + ds + '</span></label>';
  }).join('');
  show('adminRegisterDatesActions');
}

async function doRegisterDates() {
  const dates = Array.from(el('adminSaturdayList').querySelectorAll('input:checked')).map(i => i.value);
  if (!dates.length) { showMsg('日程を選択してください。', 'error'); return; }
  disableAll(true);
  const r = await api('addScheduleSessions', { idToken: S.idToken, dates });
  disableAll(false);
  showMsg(r.message, r.ok ? 'success' : 'error');
  if (r.ok) { el('adminSaturdayList').innerHTML = ''; hide('adminRegisterDatesActions'); loadExistingSchedule(); }
}

async function doGenerateReport() {
  const monthKey = el('adminReportMonth').value;
  if (!monthKey) { showMsg('対象月を選択してください。', 'error'); return; }
  disableAll(true);
  const r = await api('generateActivityReport', { idToken: S.idToken, monthKey });
  disableAll(false);
  showMsg(r.ok ? r.message + 'スプレッドシートの ' + esc(r.sheetName) + ' を確認してください。' : r.message, r.ok ? 'success' : 'error');
}

// ── 管理者代理登録 ──
async function doAdminRegister() {
  const fullName = normalizeNameSpaces(el('adminRegFullName').value);
  if (!fullName) { showMsg('氏名は必須です。', 'error'); return; }
  disableAll(true);
  const r = await api('adminRegisterMember', {
    idToken:     S.idToken,
    fullName,
    furigana:    normalizeNameSpaces(el('adminRegFurigana').value),
    gender:      el('adminRegGender').value,
    birthDate:   el('adminRegBirthDate').value,
    mobilePhone: el('adminRegMobilePhone').value.trim(),
    address:     el('adminRegAddress').value.trim(),
  });
  disableAll(false);
  showMsg(r.message, r.ok ? 'success' : 'error');
  if (r.ok) {
    ['adminRegFullName','adminRegFurigana','adminRegBirthDate','adminRegMobilePhone','adminRegAddress']
      .forEach(id => { el(id).value = ''; });
    el('adminRegGender').value = '';
  }
}

// ── 名簿から名前リスト取得 ──
async function loadMemberNames() {
  const r = await api('getMemberNames', { idToken: S.idToken });
  S.memberNames = r.names || [];
  S.memberFurigana = r.furigana || [];
}

// ── 出欠保存 ──
async function doSave() {
  if (!S.member) { showMsg('先に登録を完了してください。', 'error'); return; }
  disableAll(true);
  const r = await api('submitAvailability', { idToken: S.idToken, answers: collectAnswers() });
  disableAll(false);
  if (!r.ok) { showMsg(r.message || '保存失敗', 'error'); return; }
  S.schedule = r.schedule || [];
  renderSchedule(r.schedule || []);
  el('lastSavedAt').textContent = '最終保存: ' + r.savedAt;
  showMsg(r.message || '保存しました。', 'success');
  if (S.isAdmin) loadAdminView();
  populateTeamSelect();
}

// ── スケジュール描画 ──
function renderSchedule(months) {
  const root = el('scheduleRoot');
  root.innerHTML = '';
  if (!months.length) { root.innerHTML = '<p class="muted">現在回答可能な予定はありません。</p>'; return; }
  const today = new Date().toISOString().slice(0, 10);
  months.forEach(month => {
    const sec = document.createElement('section');
    sec.className = 'month-section';
    sec.innerHTML = '<h3>' + esc(month.monthKey) + '</h3>';
    month.sessions.forEach(s => {
      const art = document.createElement('article');
      art.className = 'session-card';
      art.dataset.sessionId = s.sessionId;
      const isToday = s.eventDate === today;
      const alreadyAttended = s.attendees && s.attendees.some(a => a.lineId === (S.member && S.member.lineId));
      const selfCheckBtn = isToday
        ? '<button class="' + (alreadyAttended ? 'secondary' : 'primary') + ' self-attend-btn" data-sid="' + esc(s.sessionId) + '">' +
          (alreadyAttended ? '当日参加済' : '当日参加する') + '</button>'
        : '';
      art.innerHTML =
        '<div class="session-header"><div>' +
          '<div class="session-date">' + esc(s.eventDate) + (isToday ? ' <span class="pill today-pill">本日</span>' : '') + '</div>' +
          '<div class="session-title">' + esc(s.title) + '</div>' +
          (s.note ? '<div class="muted">' + esc(s.note) + '</div>' : '') +
        '</div><div class="status-block">' +
          '<span class="pill">' + esc(s.statusLabel) + '</span>' +
          '<span class="muted">参加 ' + s.counts.yes + ' / 不参加 ' + s.counts.no + ' / 未定 ' + s.counts.undecided + '</span>' +
        '</div></div>' +
        '<div class="answer-row">' +
          '<label><input type="radio" name="a-' + s.sessionId + '" value="yes"' + (s.myAnswer === 'yes' || s.myAnswer === 'attended' ? ' checked' : '') + '> 参加</label>' +
          '<label><input type="radio" name="a-' + s.sessionId + '" value="no"' + (s.myAnswer === 'no' ? ' checked' : '') + '> 不参加</label>' +
          '<label><input type="radio" name="a-' + s.sessionId + '" value="undecided"' + (!s.myAnswer || s.myAnswer === 'undecided' ? ' checked' : '') + '> 未定</label>' +
        '</div>' +
        '<label class="note-field"><span>備考</span><input type="text" class="session-note-input" placeholder="任意" value="' + esc(s.myNote || '') + '"/></label>' +
        (selfCheckBtn ? '<div class="actions">' + selfCheckBtn + '</div>' : '');
      sec.appendChild(art);
    });
    root.appendChild(sec);
  });
  // 当日参加ボタンのイベント
  root.querySelectorAll('.self-attend-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      disableAll(true);
      const r = await api('markSelfAttendance', { idToken: S.idToken, sessionId: btn.dataset.sid });
      disableAll(false);
      if (r.ok) { S.schedule = r.schedule || []; renderSchedule(r.schedule || []); showMsg(r.message, 'success'); }
      else { showMsg(r.message, 'error'); }
    });
  });
}

function collectAnswers() {
  return Array.from(document.querySelectorAll('.session-card')).map(c => ({
    sessionId: c.dataset.sessionId,
    answer: (c.querySelector('input[name="a-' + c.dataset.sessionId + '"]:checked') || {}).value || 'undecided',
    note: c.querySelector('.session-note-input').value,
  }));
}

// ── 管理者: 当日参加チェック ──
async function loadAdminView() {
  const mk = S.schedule && S.schedule[0] ? S.schedule[0].monthKey : '';
  if (!mk) return;
  const r = await api('getAdminData', { idToken: S.idToken, monthKey: mk });
  if (r.ok) renderAdmin(r.schedule || []);
}

function renderAdmin(months) {
  const root = el('adminRoot');
  root.innerHTML = '';
  months.forEach(month => {
    month.sessions.forEach(s => {
      const box = document.createElement('div');
      box.className = 'admin-box';
      const list = (s.attendees||[]).map(m =>
        '<label class="admin-check"><input type="checkbox" data-uid="' + esc(m.lineId) + '"' +
        (m.answer==='attended'?' checked':'') + '><span>' + esc(m.fullName) + '</span></label>'
      ).join('');
      box.innerHTML =
        '<div class="session-header"><div>' +
          '<div class="session-date">' + esc(s.eventDate) + '</div>' +
          '<div class="session-title">' + esc(s.title) + '</div>' +
        '</div>' +
        '<button class="secondary" data-action="mark" data-sid="' + esc(s.sessionId) + '">当日参加を保存</button></div>' +
        '<div class="admin-list">' + (list || '<span class="muted">参加予定者なし</span>') + '</div>';
      root.appendChild(box);
    });
  });
  root.querySelectorAll('[data-action="mark"]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const box = btn.closest('.admin-box');
      const ids = Array.from(box.querySelectorAll('input[type="checkbox"]:checked')).map(i => i.dataset.uid);
      disableAll(true);
      const r = await api('markAttendance', { idToken: S.idToken, sessionId: btn.dataset.sid, presentLineUserIds: ids });
      disableAll(false);
      if (r.ok) { S.schedule = r.schedule || []; renderSchedule(r.schedule || []); renderAdmin(r.schedule || []); showMsg(r.message, 'success'); }
      else { showMsg(r.message, 'error'); }
    });
  });
}

// ── チーム編成 ──
let _currentGameNumber = 1;
let _prevRestIds = [];
let _currentSessionId = '';
let _loadingTeam = false;

function populateTeamSelect() {
  const sel = el('teamSessionSelect');
  sel.innerHTML = '';
  (S.schedule||[]).forEach(m => {
    m.sessions.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s.sessionId;
      opt.textContent = s.eventDate + ' ' + s.title + '（参加' + s.counts.yes + '名）';
      sel.appendChild(opt);
    });
  });
  // changeイベントは1回のみ登録
  sel.onchange = () => {
    _currentSessionId = '';
    _loadingTeam = false;
    el('teamPlayerList').innerHTML = '';
    el('teamResult').innerHTML = '';
    hide('resultCard');
    loadTeamPlayerList();
  };
  if (sel.options.length) loadTeamPlayerList();
}

async function loadTeamPlayerList() {
  if (_loadingTeam) return;
  _loadingTeam = true;
  try {
  const sessionId = el('teamSessionSelect').value;
  if (!sessionId) { _loadingTeam = false; return; }
  if (sessionId !== _currentSessionId) {
    _currentSessionId = sessionId;
    _prevRestIds = [];
    // バックエンドから次のゲーム番号を取得
    const gn = await api('getLatestGameNumber', { idToken: S.idToken, sessionId });
    _currentGameNumber = gn.ok ? gn.nextGameNumber : 1;
    // 編成済みで結果未入力のゲームがあれば自動表示してここで終了
    if (gn.ok && gn.pendingGame) {
      el('teamPlayerList').innerHTML = '';
      renderGameResult(gn.pendingGame, []);
      show('resultCard');
      hide('generateTeamsBtn');
      el('gameNumberLabel').textContent = _currentGameNumber;
      return;
    }
  }
  el('gameNumberLabel').textContent = _currentGameNumber;
  show('generateTeamsBtn');

  // 休憩者自動提案を取得
  const r = await api('suggestRest', { idToken: S.idToken, sessionId, prevRestIds: _prevRestIds });
  const suggestedRestIds = r.ok ? r.suggestedRestIds : [];

  // 参加者一覧を構築
  const session = (S.schedule||[]).flatMap(m => m.sessions).find(s => s.sessionId === sessionId);
  if (!session || !session.attendees.length) {
    el('teamPlayerList').innerHTML = '<p class="muted">参加予定者がいません。</p>';
    return;
  }
  el('teamPlayerList').innerHTML =
    '<p class="muted" style="margin-bottom:8px">休憩する方のチェックを外してください。システムが休憩候補を自動提案しています。</p>' +
    session.attendees.map(a =>
      '<label class="player-check-item">' +
        '<input type="checkbox" class="player-check" value="' + esc(a.lineId) + '"' +
        (suggestedRestIds.includes(a.lineId) ? '' : ' checked') + '>' +
        '<span>' + esc(a.fullName) + (a.isTrial ? ' <span class="trial-badge">体験</span>' : '') + '</span>' +
      '</label>'
    ).join('');
  } finally { _loadingTeam = false; }
}

async function doGenerateTeams() {
  const sessionId = el('teamSessionSelect').value;
  if (!sessionId) { showMsg('予定を選択してください。', 'error'); return; }
  const checkedIds = Array.from(el('teamPlayerList').querySelectorAll('.player-check:checked')).map(i => i.value);
  const allIds = Array.from(el('teamPlayerList').querySelectorAll('.player-check')).map(i => i.value);
  const restLineIds = allIds.filter(id => !checkedIds.includes(id));
  disableAll(true);
  const r = await api('generateTeams', { idToken: S.idToken, sessionId, gameNumber: _currentGameNumber, restLineIds });
  disableAll(false);
  if (!r.ok) { showMsg(r.message, 'error'); return; }
  renderGameResult(r.game, r.restPlayers);
  show('resultCard');
  showMsg(r.message, 'success');
  el('messageCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function renderGameResult(game, restPlayers) {
  const root = el('teamResult');
  let html = '<div class="game-block">';
  html += '<div class="game-title">第' + game.gameNumber + 'ゲーム</div>';
  html += '<div class="matchup">' + game.matchups.map(esc).join(' / ') + '</div>';
  if (restPlayers && restPlayers.length) {
    html += '<div class="rest-players">休憩: ' + restPlayers.map(p => esc(p.fullName)).join('、') + '</div>';
  }
  html += '<div class="teams-grid">';
  game.teams.forEach(t => {
    html += '<div class="team-box"><div class="team-name">' + esc(t.name) + '</div>';
    html += '<div class="team-meta">' + esc(t.genderSummary) + ' / ' + esc(t.ageSummary) + ' / 勝率' + esc(t.avgWinRate) + '</div>';
    t.members.forEach(m => {
      const wr = m.totalGames > 0 ? ' 勝率' + m.winRate + '%(得点差' + (m.avgScoreDiff >= 0 ? '+' : '') + Math.round(m.avgScoreDiff || 0) + ')' : '';
      const trial = m.isTrial ? ' <span class="trial-badge">体験</span>' : '';
      html += '<div class="team-member">' + esc(m.fullName) + trial + ' <span class="muted">' + esc(m.gender) + (m.ageApril1 ? ' ' + m.ageApril1 + '歳' : '') + wr + '</span></div>';
    });
    html += '</div>';
  });
  html += '</div></div>';
  root.innerHTML = html;
  renderResultInputs(game);
}

// ── 試合結果入力 ──
function renderResultInputs(game) {
  const root = el('resultRoot');
  root.innerHTML = '';
  game.matchups.forEach((mu, mi) => {
    const teamNames = game.teams.map(t => t.name);
    let tA = teamNames[0], tB = teamNames[1];
    if (game.numTeams === 4) { tA = teamNames[mi*2]; tB = teamNames[mi*2+1]; }
    const match = mu.match(/(.+?)\s*vs\s*(.+?)（/);
    if (match) { tA = match[1].trim(); tB = match[2].trim(); }
    const div = document.createElement('div');
    div.className = 'game-block';
    div.innerHTML =
      '<div class="game-title">第' + game.gameNumber + 'ゲーム' + (game.numTeams === 4 ? '（コート' + (mi+1) + '）' : '') + '</div>' +
      '<div class="score-row">' +
        '<span class="score-team">' + esc(tA) + '</span>' +
        '<input type="number" class="score-input" data-game="' + game.gameNumber + '" data-match="' + mi + '" data-side="A" data-team="' + esc(tA) + '" min="0" max="99" placeholder="0"/>' +
        '<span class="score-vs">-</span>' +
        '<input type="number" class="score-input" data-game="' + game.gameNumber + '" data-match="' + mi + '" data-side="B" data-team="' + esc(tB) + '" min="0" max="99" placeholder="0"/>' +
        '<span class="score-team">' + esc(tB) + '</span>' +
      '</div>';
    root.appendChild(div);
  });
}

async function doSaveResults() {
  const sessionId = el('teamSessionSelect').value;
  if (!sessionId) { showMsg('予定を選択してください。', 'error'); return; }
  const inputs = document.querySelectorAll('.score-input');
  const map = {};
  inputs.forEach(inp => {
    const key = inp.dataset.game + '|' + inp.dataset.match;
    if (!map[key]) map[key] = { gameNumber: Number(inp.dataset.game) };
    if (inp.dataset.side === 'A') { map[key].teamA = inp.dataset.team; map[key].scoreA = inp.value; }
    else { map[key].teamB = inp.dataset.team; map[key].scoreB = inp.value; }
  });
  const results = Object.values(map).filter(r => r.scoreA !== '' || r.scoreB !== '');
  const sel = el('teamSessionSelect');
  const evtDate = (sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].textContent : '').split(' ')[0] || '';
  results.forEach(r => r.eventDate = evtDate);
  if (!results.length) { showMsg('スコアを入力してください。', 'error'); return; }
  disableAll(true);
  const r = await api('submitGameResults', { idToken: S.idToken, sessionId, results });
  disableAll(false);
  if (r.ok) {
    _prevRestIds = Array.from(el('teamPlayerList').querySelectorAll('.player-check:not(:checked)')).map(i => i.value);
    _currentGameNumber++;
    el('gameNumberLabel').textContent = _currentGameNumber;
    el('teamResult').innerHTML = '';
    hide('resultCard');
    doRefreshWinRates();
    await loadTeamPlayerList();
    showMsg(r.message + ' 次のゲームの参加者を選択してください。', 'success');
    el('messageCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  } else { showMsg(r.message, 'error'); }
}

// ── 勝率一覧 ──
async function doRefreshWinRates() {
  const r = await api('getWinRates', { idToken: S.idToken });
  if (r.ok) renderWinRates(r.winRates || {});
}

function renderWinRates(winRates) {
  const root = el('winRateRoot');
  const entries = Object.entries(winRates).sort((a, b) => b[1].winRate - a[1].winRate);
  if (!entries.length) { root.innerHTML = '<p class="muted">まだ試合結果がありません。</p>'; return; }
  let html = '<table class="wr-table"><tr><th>氏名</th><th>試合</th><th>勝</th><th>負</th><th>分</th><th>勝率</th></tr>';
  entries.forEach(([name, s]) => {
    html += '<tr><td>' + esc(name) + '</td><td>' + s.games + '</td><td>' + s.wins + '</td><td>' + s.losses + '</td><td>' + s.draws + '</td><td><strong>' + s.winRate + '%</strong></td></tr>';
  });
  html += '</table>';
  root.innerHTML = html;
}

// ── マイページ ──
let _chartDonut = null, _chartLine = null;

async function doLoadMyPage() {
  const r = await api('getMyPage', { idToken: S.idToken });
  if (!r.ok) { showMsg(r.message || 'エラー', 'error'); return; }
  S.myPageData = r;
  renderProfileView(r.profile);
  renderStats(r.stats, r.recentGames || []);
  toggleProfileEdit(false);
}

function renderProfileView(p) {
  el('profileView').innerHTML = [
    ['氏名',       p.fullName],
    ['ふりがな',     p.furigana    || '未登録'],
    ['性別',       p.gender      || '-'],
    ['生年月日',     p.birthDate   || '-'],
    ['年齢(4/1)',   p.ageApril1   ? p.ageApril1 + '歳' : '-'],
    ['携帯電話',     p.mobilePhone || '未登録'],
    ['自宅電話',     p.homePhone   || '-'],
    ['住所',       p.address     || '未登録'],
    ['LINE表示名',  p.lineName    || '-'],
  ].map(([label, val]) =>
    '<div class="profile-row"><span class="profile-label">' + esc(label) + '</span><span class="profile-val">' + esc(val) + '</span></div>'
  ).join('');
}

function renderStats(s, recentGames) {
  el('myStatsRoot').innerHTML =
    '<div class="stats-grid">' +
      '<div class="stat-box"><div class="stat-num">' + s.games   + '</div><div class="stat-label">試合</div></div>' +
      '<div class="stat-box"><div class="stat-num">' + s.wins    + '</div><div class="stat-label">勝</div></div>' +
      '<div class="stat-box"><div class="stat-num">' + s.losses  + '</div><div class="stat-label">負</div></div>' +
      '<div class="stat-box"><div class="stat-num">' + s.draws   + '</div><div class="stat-label">分</div></div>' +
      '<div class="stat-box accent"><div class="stat-num">' + s.winRate + '%</div><div class="stat-label">勝率</div></div>' +
    '</div>';

  if (_chartDonut) _chartDonut.destroy();
  if (_chartLine)  _chartLine.destroy();

  _chartDonut = new Chart(el('chartDonut'), {
    type: 'doughnut',
    data: {
      labels: ['勝', '負', '分'],
      datasets: [{ data: [s.wins, s.losses, s.draws], backgroundColor: ['#0f62fe','#d93025','#aab4c8'], borderWidth: 0 }],
    },
    options: { plugins: { legend: { position: 'bottom' } }, cutout: '65%' },
  });

  const labels = recentGames.map((g, i) => (i + 1) + '戦');
  const data   = recentGames.map(g => g.outcome === 'win' ? 1 : g.outcome === 'draw' ? 0.5 : 0);
  _chartLine = new Chart(el('chartLine'), {
    type: 'line',
    data: {
      labels,
      datasets: [{ label: '直近の勝敗', data, borderColor: '#0f62fe', backgroundColor: 'rgba(15,98,254,.1)', pointBackgroundColor: data.map(v => v === 1 ? '#0f62fe' : v === 0.5 ? '#aab4c8' : '#d93025'), tension: 0.3, fill: true }],
    },
    options: {
      scales: { y: { min: 0, max: 1, ticks: { callback: v => v === 1 ? '勝' : v === 0.5 ? '分' : '負' } } },
      plugins: { legend: { display: false } },
    },
  });
}

function toggleProfileEdit(editing) {
  el('profileView').classList.toggle('hidden', editing);
  el('profileEdit').classList.toggle('hidden', !editing);
  el('editProfileBtn').classList.toggle('hidden', editing);
  if (editing && S.myPageData) {
    const p = S.myPageData.profile;
    el('editFullName').value    = p.fullName    || '';
    el('editFurigana').value    = p.furigana    || '';
    el('editGender').value      = p.gender      || '男';
    el('editBirthDate').value   = p.birthDate   || '';
    el('editMobilePhone').value = p.mobilePhone || '';
    el('editAddress').value     = p.address     || '';
    el('editHomePhone').value   = p.homePhone   || '';
  }
}

async function doSaveProfile() {
  const fullName    = normalizeNameSpaces(el('editFullName').value);
  const furigana    = normalizeNameSpaces(el('editFurigana').value);
  const gender      = el('editGender').value;
  const birthDate   = el('editBirthDate').value;
  const mobilePhone = el('editMobilePhone').value.trim();
  const address     = el('editAddress').value.trim();
  if (!fullName || !furigana || !gender || !birthDate || !mobilePhone || !address) {
    showMsg('必須項目をすべて入力してください。', 'error'); return;
  }
  disableAll(true);
  const r = await api('updateProfile', {
    idToken: S.idToken, fullName, furigana, gender, birthDate, mobilePhone, address,
    homePhone: el('editHomePhone').value.trim(),
  });
  disableAll(false);
  showMsg(r.message, r.ok ? 'success' : 'error');
  if (r.ok) doLoadMyPage();
}

// ── 共有資料 ──
async function doLoadDocs() {
  const r = await api('getDocs', { idToken: S.idToken });
  const root = el('docsRoot');
  if (!r.ok) { root.innerHTML = '<p class="muted">資料の取得に失敗しました。</p>'; showMsg('エラー: ' + (r.message || ''), 'error'); return; }
  if (!r.docs.length) { root.innerHTML = '<p class="muted">現在公開中の資料はありません。</p>'; return; }
  root.innerHTML = r.docs.map(d =>
    '<a class="doc-item" href="' + esc(d.url) + '" target="_blank" rel="noopener">' +
      '<span class="doc-icon">📄</span>' +
      '<span class="doc-info"><span class="doc-name">' + esc(d.name) + '</span>' +
      '<span class="doc-date muted">' + esc(d.updatedAt) + '</span></span>' +
      '<span class="doc-arrow">›</span>' +
    '</a>'
  ).join('');
}

// ── 参加者リスト出力 ──
async function doExportList() {
  const sessionId = el('teamSessionSelect').value;
  if (!sessionId) { showMsg('予定を選択してください。', 'error'); return; }
  disableAll(true);
  const r = await api('generateExportList', { idToken: S.idToken, sessionId });
  disableAll(false);
  showMsg(r.message, r.ok ? 'success' : 'error');
}

// ── ユーティリティ ──
function el(id) { return document.getElementById(id); }
function show(id) { el(id).classList.remove('hidden'); }
function hide(id) { el(id).classList.add('hidden'); }
function disableAll(d) { document.querySelectorAll('button').forEach(b => b.disabled = d); }
function showMsg(t, type) { const a = el('messageArea'); a.className = 'message ' + (type||'info'); a.textContent = t; if (t) show('messageCard'); }
function esc(v) { return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
