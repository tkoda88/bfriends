/* =====================================================
   Bフレンズ ソフトバレー 参加管理アプリ
   Code.gs — メインロジック

   スプレッドシート構成:
     ROSTER_SPREADSHEET_ID → 「Bフレンズ名簿」(Members のみ)
     SPREADSHEET_ID        → 運用用 (Schedule, Responses, GameSets, GameResults, ExportTemplate)
   ===================================================== */

const SHEET_NAMES = {
  MEMBERS:     'Members',
  SCHEDULE:    'Schedule',
  RESPONSES:   'Responses',
  GAMESETS:    'GameSets',
  GAMERESULTS: 'GameResults',
  EXPORT:      'ExportTemplate',
};

/* Members ヘッダーは名簿スプレッドシートの実際のヘッダー名をそのまま使用。
   コード内では MC.XXX で参照する。 */
const MC = {
  STATUS:       '入会',
  NO:           'Ｎｏ',
  FULL_NAME:    '氏 名',
  FURIGANA:     'ふりがな',
  ADDRESS:      '住 所',
  HOME_PHONE:   '自宅電話',
  MOBILE_PHONE: '携帯電話',
  GENDER:       '性別',
  BIRTH_DATE:   '生年月日',
  AGE_APRIL1:   '4/1年齢',
  INSURANCE:    'R７保険',
  CONTACT:      '連絡方法',
  LINE_ID:      'LINE ID',
  LINE_NAME:    'LINE表示名',
  NOTE:         'メモ',
};

const HEADERS = {
  Schedule: [
    'sessionId','monthKey','eventDate','title',
    'minAttendees','maxAttendees','openForResponse','note'
  ],
  Responses: [
    'lineId','no','fullName','sessionId','monthKey',
    'eventDate','title','answer','note','lineDisplayName','submittedAt'
  ],
  GameSets: [
    'sessionId','eventDate','gameNumber','teamName',
    'no','fullName','gender','ageApril1','createdAt'
  ],
  GameResults: [
    'sessionId','eventDate','gameNumber',
    'teamA','scoreA','teamB','scoreB',
    'winTeam','createdAt'
  ],
};

// ── Web App API ──
// GitHub Pages から fetch で呼び出される。
// GET: ?action=xxx&param=yyy
// POST: JSON body { action, ...params }

function doGet(e) {
  // GETフォールバック: ?payload=JSON または ?action=xxx
  const params = e ? e.parameter : {};
  if (params.payload) {
    const body = JSON.parse(decodeURIComponent(params.payload));
    return handleRequest_(body.action || '', body);
  }
  const action = params.action || '';
  return handleRequest_(action, params);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const action = body.action || '';
  return handleRequest_(action, body);
}

function handleRequest_(action, params) {
  try {
    let result;
    switch (action) {
      case 'initializeSession':       result = initializeSession(params.idToken); break;
      case 'getMemberNames':          result = getMemberNamesForRegistration(params.idToken); break;
      case 'registerMember':          result = registerMember(params); break;
      case 'registerNewMember':       result = registerNewMember(params); break;
      case 'submitAvailability':      result = submitAvailability(params); break;
      case 'getAdminData':            result = getAdminData(params.idToken, params.monthKey); break;
      case 'markAttendance':          result = markAttendance(params); break;
      case 'markSelfAttendance':       result = markSelfAttendance(params); break;
      case 'generateTeams':           result = generateTeams(params); break;
      case 'suggestRest':              result = suggestRestAction(params); break;
      case 'getLatestGameNumber':       result = getLatestGameNumber(params); break;
      case 'submitGameResults':       result = submitGameResults(params); break;
      case 'getWinRates':             result = getWinRates(params.idToken); break;
      case 'generateExportList':      result = generateExportList(params); break;
      case 'getDocs':                 result = getDocs(params.idToken); break;
      case 'getMyPage':               result = getMyPage(params.idToken); break;
      case 'updateProfile':           result = updateProfile(params); break;
      case 'adminRegisterMember':     result = adminRegisterMember(params); break;
      case 'getAdminPage':             result = getAdminPage(params.idToken); break;
      case 'addScheduleSessions':      result = addScheduleSessions(params); break;
      case 'deleteScheduleSession':    result = deleteScheduleSession(params); break;
      case 'generateActivityReport':   result = generateActivityReport(params); break;
      default:                        result = { ok: false, message: 'Unknown action: ' + action };
    }
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ ok: false, message: err.message || String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 初期セットアップ（運用スプレッドシート側のみ） ──

function setupSheets() {
  const ss = getOpsSS_();
  ['Schedule','Responses','GameSets','GameResults'].forEach(name => {
    createSheetIfMissing_(ss, name, HEADERS[name]);
  });
  if (!ss.getSheetByName(SHEET_NAMES.EXPORT)) {
    const s = ss.insertSheet(SHEET_NAMES.EXPORT);
    s.getRange('A1').setValue('（提出用フォーマットをここに貼り付けてください）');
  }
  // Schedule にサンプル
  const sch = ss.getSheetByName(SHEET_NAMES.SCHEDULE);
  if (sch.getLastRow() === 1) {
    const today = new Date();
    const rows = [];
    for (let i = 1; i <= 4; i++) {
      const d = new Date(today.getFullYear(), today.getMonth(), 7 * i + 1);
      rows.push([
        'S' + Utilities.getUuid().slice(0, 8),
        Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM'),
        Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'),
        '定例活動 ' + i, 6, 24, true, ''
      ]);
    }
    sch.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { ok: true, message: 'シート初期化完了', url: ss.getUrl() };
}

// ── LIFF セッション初期化 ──

function initializeSession(idToken) {
  const profile  = verifyIdToken_(idToken);
  const member   = findMemberByLineId_(profile.sub);
  const schedule = buildScheduleView_(null, profile.sub);
  const isAdmin  = checkAdmin_(profile.sub);
  return {
    ok: true,
    appTitle: PropertiesService.getScriptProperties().getProperty('APP_TITLE') || 'Bフレンズ 参加管理',
    memberFound: !!member,
    member: member ? {
      no: member[MC.NO], fullName: member[MC.FULL_NAME],
      gender: member[MC.GENDER], ageApril1: member[MC.AGE_APRIL1],
      furigana: member[MC.FURIGANA], mobilePhone: member[MC.MOBILE_PHONE],
      address: member[MC.ADDRESS], birthDate: member[MC.BIRTH_DATE], lineId: profile.sub,
    } : null,
    lineProfileName: profile.name || '',
    schedule,
    isAdmin,
    nowJst: fmtJst_(new Date()),
  };
}

// ── 名簿の名前リスト取得（初回登録用）──

function getMemberNamesForRegistration(idToken) {
  verifyIdToken_(idToken);
  const rows = readObjects_(getRosterSheet_());
  const active = rows.filter(r => !r[MC.LINE_ID] && isActiveMember_(r));
  return {
    ok: true,
    names: active.map(r => r[MC.FULL_NAME]).filter(Boolean),
    furigana: active.map(r => r[MC.FURIGANA] || ''),
  };
}

// ── 既存メンバー紐付け（名簿にある人が LINE userId を登録）──

function registerMember(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const fullName = String(payload.fullName || '').trim();
    if (!fullName) throw new Error('氏名は必須です。');

    const sheet = getRosterSheet_();
    const rows = readObjects_(sheet);

    const normalize = s => String(s||'').replace(/[\s　]+/g, '');
    const idx = rows.findIndex(r => normalize(r[MC.FULL_NAME]) === normalize(fullName));
    if (idx < 0) throw new Error('名簿に「' + fullName + '」が見つかりません。');

    if (!isActiveMember_(rows[idx])) throw new Error('この名前は現在休会または退会中です。');

    const headers = getSheetHeaders_(sheet);
    setColValue_(sheet, headers, idx + 2, MC.LINE_ID, profile.sub);
    setColValue_(sheet, headers, idx + 2, MC.LINE_NAME, profile.name || '');

    const member = rows[idx];
    return {
      ok: true,
      member: { no: member[MC.NO], fullName: member[MC.FULL_NAME], gender: member[MC.GENDER], ageApril1: member[MC.AGE_APRIL1] },
      lineProfileName: profile.name || '',
    };
  } finally { lock.releaseLock(); }
}

// ── 新規メンバー登録（名簿に新規行を追加 + LINE userId も同時登録）──

function registerNewMember(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const fullName = String(payload.fullName || '').trim();
    const furigana = String(payload.furigana || '').trim();
    const gender   = String(payload.gender || '').trim();
    const birthDate = String(payload.birthDate || '').trim();
    const mobilePhone = String(payload.mobilePhone || '').trim();

    if (!fullName) throw new Error('氏名は必須です。');
    if (!gender)   throw new Error('性別は必須です。');
    if (!birthDate) throw new Error('生年月日は必須です。');
    if (!mobilePhone) throw new Error('携帯電話は必須です。');
    const address = String(payload.address || '').trim();
    if (!address) throw new Error('住所は必須です。');

    const sheet = getRosterSheet_();
    const rows = readObjects_(sheet);
    const headers = getSheetHeaders_(sheet);

    // 重複チェック
    const normalize = s => String(s||'').replace(/[\s　]+/g, '');
    if (rows.some(r => normalize(r[MC.FULL_NAME]) === normalize(fullName))) {
      throw new Error('「' + fullName + '」は既に名簿に登録されています。既存メンバーの場合は名前を選択して紐付けしてください。');
    }

    // 次のNo を算出
    const maxNo = rows.reduce((mx, r) => Math.max(mx, Number(r[MC.NO]) || 0), 0);
    const newNo = maxNo + 1;

    // 4/1年齢を計算
    let ageApril1 = '';
    if (birthDate) {
      const bd = new Date(birthDate);
      const now = new Date();
      const aprilYear = now.getMonth() < 3 ? now.getFullYear() : now.getFullYear();
      const april1 = new Date(aprilYear, 3, 1);
      ageApril1 = Math.floor((april1 - bd) / (365.25 * 24 * 60 * 60 * 1000));
    }

    // 新規行を追加
    const newRow = headers.map(h => {
      switch(h) {
        case MC.STATUS:       return '入会';
        case MC.NO:           return newNo;
        case MC.FULL_NAME:    return fullName;
        case MC.FURIGANA:     return furigana;
        case MC.GENDER:       return gender;
        case MC.BIRTH_DATE:   return birthDate;
        case MC.AGE_APRIL1:   return ageApril1;
        case MC.MOBILE_PHONE: return mobilePhone;
        case MC.ADDRESS:      return address;
        case MC.CONTACT:      return 'LINE';
        case MC.LINE_ID:      return profile.sub;
        case MC.LINE_NAME:    return profile.name || '';
        case MC.NOTE:         return fmtJst_(new Date()) + ' LIFF新規登録';
        default:              return '';
      }
    });
    appendRowRaw_(sheet, headers, newRow);

    return {
      ok: true,
      member: { no: String(newNo), fullName, gender, ageApril1: String(ageApril1) },
      lineProfileName: profile.name || '',
    };
  } finally { lock.releaseLock(); }
}

// ── 出欠回答 ──

function submitAvailability(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const member = findMemberByLineId_(profile.sub);
    if (!member) return { ok: false, needsRegistration: true, message: '名簿に未登録です。先に初回登録を行ってください。' };

    const answers = Array.isArray(payload.answers) ? payload.answers : [];
    if (!answers.length) throw new Error('回答対象がありません。');

    const scheduleMap = getScheduleMap_();
    const sheet = getOpsSS_().getSheetByName(SHEET_NAMES.RESPONSES);
    const existing = readObjects_(sheet);
    const now = fmtJst_(new Date());

    answers.forEach(a => {
      const ses = scheduleMap[a.sessionId];
      if (!ses || String(ses.openForResponse).toLowerCase() === 'false') return;
      const answer = normalizeAnswer_(a.answer);
      const record = {
        lineId: profile.sub, no: member[MC.NO] || '', fullName: member[MC.FULL_NAME],
        sessionId: ses.sessionId, monthKey: ses.monthKey, eventDate: ses.eventDate,
        title: ses.title, answer, note: String(a.note||'').trim(),
        lineDisplayName: profile.name || '', submittedAt: now,
      };
      const ei = existing.findIndex(r => r.lineId === profile.sub && r.sessionId === ses.sessionId);
      if (ei >= 0) {
        writeRow_(sheet, ei + 2, HEADERS.Responses, record);
        existing[ei] = record;
      } else {
        appendRow_(sheet, HEADERS.Responses, record);
        existing.push(record);
      }
    });

    return { ok: true, memberName: member[MC.FULL_NAME], savedAt: now, schedule: buildScheduleView_(null, profile.sub), message: '参加予定を保存しました。' };
  } finally { lock.releaseLock(); }
}

// ── 管理者: 当日参加チェック ──

function getAdminData(idToken, monthKey) {
  const profile = verifyIdToken_(idToken);
  ensureAdmin_(profile.sub);
  return { ok: true, schedule: buildScheduleView_(monthKey) };
}

function markAttendance(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    ensureAdmin_(profile.sub);
    const sessionId = String(payload.sessionId || '');
    const presentIds = (payload.presentLineUserIds || []).map(v => String(v).trim()).filter(Boolean);
    const sheet = getOpsSS_().getSheetByName(SHEET_NAMES.RESPONSES);
    const rows = readObjects_(sheet);
    rows.forEach((r, i) => {
      if (r.sessionId !== sessionId) return;
      const newAns = presentIds.includes(r.lineId) ? 'attended' : r.answer;
      if (newAns !== r.answer) {
        r.answer = newAns;
        r.submittedAt = fmtJst_(new Date());
        writeRow_(sheet, i + 2, HEADERS.Responses, r);
      }
    });
    // 活動報告書を自動更新
    updateActivityReport_(sessionId);
    return { ok: true, schedule: buildScheduleView_(null, profile.sub), message: '当日参加を更新しました。' };
  } finally { lock.releaseLock(); }
}

// ── チーム編成（1ゲームずつ）──

function generateTeams(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const caller = findMemberByLineId_(profile.sub);
    if (!caller) throw new Error('名簿に未登録です。');

    const sessionId = String(payload.sessionId || '');
    const gameNumber = Number(payload.gameNumber || 1);
    const restLineIds = Array.isArray(payload.restLineIds) ? payload.restLineIds : [];

    const opsSS = getOpsSS_();
    const responses = readObjects_(opsSS.getSheetByName(SHEET_NAMES.RESPONSES));
    const attendees = responses.filter(r => r.sessionId === sessionId && (r.answer === 'yes' || r.answer === 'attended'));
    if (attendees.length < 6) throw new Error('参加者が6名未満のためチーム編成できません。');

    const members = readObjects_(getRosterSheet_());
    const memberMap = {};
    members.forEach(m => { if (m[MC.LINE_ID]) memberMap[m[MC.LINE_ID]] = m; });
    const winRates = calcWinRates_();

    const players = attendees
      .filter(a => !restLineIds.includes(a.lineId))
      .map(a => {
        const m = memberMap[a.lineId] || {};
        const wr = winRates[a.fullName];
        return {
          no: a.no || m[MC.NO] || '',
          fullName: a.fullName,
          lineId: a.lineId,
          gender: m[MC.GENDER] || '不明',
          ageApril1: Number(m[MC.AGE_APRIL1]) || 0,
          winRate: wr ? wr.winRate : 50,
          totalGames: wr ? wr.games : 0,
          avgScoreDiff: wr ? wr.avgScoreDiff : 0,
          isTrial: String(m[MC.STATUS] || '').trim() === '体験',
        };
      });

    if (players.length < 6) throw new Error('休桬者を除いた参加者が6名未満です。');

    // 前ゲームのペアカウント・対戦カウントを構築
    const gsSheet = opsSS.getSheetByName(SHEET_NAMES.GAMESETS);
    const prevRows = readObjects_(gsSheet).filter(r => r.sessionId === sessionId);
    const pairCount = {};
    const vsCount = {};
    // 過去全ゲームの同チームペアカウント
    const prevTeamsByGame = {};
    prevRows.forEach(r => {
      const gk = r.gameNumber;
      if (!prevTeamsByGame[gk]) prevTeamsByGame[gk] = {};
      if (!prevTeamsByGame[gk][r.teamName]) prevTeamsByGame[gk][r.teamName] = [];
      prevTeamsByGame[gk][r.teamName].push(r.fullName);
    });
    Object.values(prevTeamsByGame).forEach(teamMap => {
      Object.values(teamMap).forEach(ms => {
        for (let i = 0; i < ms.length; i++)
          for (let j = i + 1; j < ms.length; j++) {
            const key = [ms[i], ms[j]].sort().join('|');
            pairCount[key] = (pairCount[key] || 0) + 1;
          }
      });
      // 対戦カウント（対戦したチーム名のペア）→ GameResultsから取得
    });
    const grRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.GAMERESULTS))
      .filter(r => r.sessionId === sessionId);
    grRows.forEach(r => {
      const key = [r.teamA, r.teamB].sort().join('|');
      vsCount[key] = (vsCount[key] || 0) + 1;
    });

    const game = makeOneGame(players, gameNumber, pairCount, vsCount);

    // 同ゲーム番号の既存データを上書き
    const allRows = readObjects_(gsSheet);
    for (let i = allRows.length - 1; i >= 0; i--) {
      if (allRows[i].sessionId === sessionId && Number(allRows[i].gameNumber) === gameNumber)
        gsSheet.deleteRow(i + 2);
    }
    const now = fmtJst_(new Date());
    const eventDate = attendees[0] ? attendees[0].eventDate : '';
    game.teams.forEach(team => {
      team.members.forEach(p => {
        appendRow_(gsSheet, HEADERS.GameSets, {
          sessionId, eventDate, gameNumber,
          teamName: team.name, no: p.no, fullName: p.fullName,
          gender: p.gender, ageApril1: p.ageApril1, createdAt: now,
        });
      });
    });

    const restPlayers = attendees
      .filter(a => restLineIds.includes(a.lineId))
      .map(a => ({ lineId: a.lineId, fullName: a.fullName }));

    return { ok: true, game: { gameNumber, ...game }, restPlayers, message: '第' + gameNumber + 'ゲームのチーム編成を保存しました。' };
  } finally { lock.releaseLock(); }
}

// ── 休憩者自動提案 ──

function suggestRestAction(payload) {
  const profile = verifyIdToken_(payload.idToken);
  const caller = findMemberByLineId_(profile.sub);
  if (!caller) throw new Error('名簿に未登録です。');
  const sessionId = String(payload.sessionId || '');
  const prevRestIds = Array.isArray(payload.prevRestIds) ? payload.prevRestIds : [];
  const opsSS = getOpsSS_();
  const responses = readObjects_(opsSS.getSheetByName(SHEET_NAMES.RESPONSES));
  const attendees = responses.filter(r => r.sessionId === sessionId && (r.answer === 'yes' || r.answer === 'attended'));
  const members = readObjects_(getRosterSheet_());
  const memberMap = {};
  members.forEach(m => { if (m[MC.LINE_ID]) memberMap[m[MC.LINE_ID]] = m; });
  const allPlayers = attendees.map(a => {
    const m = memberMap[a.lineId] || {};
    return { lineId: a.lineId, fullName: a.fullName, isTrial: String(m[MC.STATUS] || '').trim() === '体験' };
  });
  return { ok: true, suggestedRestIds: suggestRest(allPlayers, prevRestIds) };
}

// ── 休桬者自動提案 ──

function getLatestGameNumber(payload) {
  verifyIdToken_(payload.idToken);
  const sessionId = String(payload.sessionId || '');
  const gsRows = readObjects_(getOpsSS_().getSheetByName(SHEET_NAMES.GAMESETS))
    .filter(r => r.sessionId === sessionId);
  const grRows = readObjects_(getOpsSS_().getSheetByName(SHEET_NAMES.GAMERESULTS))
    .filter(r => r.sessionId === sessionId);
  const maxGame  = gsRows.reduce((mx, r) => Math.max(mx, Number(r.gameNumber) || 0), 0);
  const maxSaved = grRows.reduce((mx, r) => Math.max(mx, Number(r.gameNumber) || 0), 0);

  // 未入力ゲームの編成データを取得
  let pendingGame = null;
  if (maxGame > maxSaved) {
    const gameNum = maxGame;
    const teamMap = {};
    gsRows.filter(r => Number(r.gameNumber) === gameNum).forEach(r => {
      if (!teamMap[r.teamName]) teamMap[r.teamName] = [];
      teamMap[r.teamName].push({ fullName: r.fullName, gender: r.gender, ageApril1: Number(r.ageApril1) || 0, winRate: 0, totalGames: 0, avgScoreDiff: 0, isTrial: false });
    });
    const teams = Object.entries(teamMap)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([name, members]) => ({
        name,
        members,
        genderSummary: '男' + members.filter(m => m.gender === '男').length + ' 女' + members.filter(m => m.gender !== '男').length,
        ageSummary: members.length ? '平均' + Math.round(members.reduce((s, m) => s + m.ageApril1, 0) / members.length) + '歳' : '-',
        avgWinRate: '-',
      }));
    const numTeams = teams.length;
    const matchups = numTeams === 4
      ? ['チームA vs チームB（コート1）', 'チームC vs チームD（コート2）']
      : ['チームA vs チームB'];
    pendingGame = { gameNumber: gameNum, numTeams, teams, matchups };
  }

  const nextGameNumber = pendingGame ? pendingGame.gameNumber : maxGame + 1;
  return { ok: true, nextGameNumber, pendingGame };
}

// ── 試合結果入力（全員実行可能）──

function submitGameResults(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const caller = findMemberByLineId_(profile.sub);
    if (!caller) throw new Error('名簿に未登録です。');
    const sessionId = String(payload.sessionId || '');
    const results = Array.isArray(payload.results) ? payload.results : [];
    if (!results.length) throw new Error('結果データがありません。');

    const sheet = getOpsSS_().getSheetByName(SHEET_NAMES.GAMERESULTS);
    const now = fmtJst_(new Date());
    const eventDate = results[0].eventDate || '';

    results.forEach(r => {
      // 同セッション・同ゲーム番号の既存行だけ削除（他ゲームは残す）
      const rows = readObjects_(sheet);
      for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i].sessionId === sessionId &&
            String(rows[i].gameNumber) === String(r.gameNumber) &&
            rows[i].teamA === r.teamA && rows[i].teamB === r.teamB) {
          sheet.deleteRow(i + 2);
        }
      }
      const scoreA = Number(r.scoreA || 0);
      const scoreB = Number(r.scoreB || 0);
      const winTeam = scoreA > scoreB ? r.teamA : scoreB > scoreA ? r.teamB : 'draw';
      appendRow_(sheet, HEADERS.GameResults, {
        sessionId, eventDate, gameNumber: r.gameNumber,
        teamA: r.teamA, scoreA, teamB: r.teamB, scoreB,
        winTeam, createdAt: now,
      });
    });

    return { ok: true, message: results.length + 'ゲーム分の結果を保存しました。' };
  } finally { lock.releaseLock(); }
}

// ── 勝率データ取得（参加者も閲覧可能）──

function getWinRates(idToken) {
  verifyIdToken_(idToken);
  return { ok: true, winRates: calcWinRates_() };
}

function calcWinRates_() {
  const opsSS = getOpsSS_();
  const gsRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.GAMESETS));
  const grRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.GAMERESULTS));

  const resultMap = {};
  grRows.forEach(r => { resultMap[r.sessionId + '|' + r.gameNumber] = r; });

  const stats = {};
  gsRows.forEach(row => {
    const result = resultMap[row.sessionId + '|' + row.gameNumber];
    if (!result) return;
    const name = row.fullName;
    if (!stats[name]) stats[name] = { wins: 0, losses: 0, draws: 0, games: 0, scoreDiffSum: 0 };
    stats[name].games++;
    const scoreA = Number(result.scoreA) || 0;
    const scoreB = Number(result.scoreB) || 0;
    const myScore  = row.teamName === result.teamA ? scoreA : scoreB;
    const oppScore = row.teamName === result.teamA ? scoreB : scoreA;
    stats[name].scoreDiffSum += myScore - oppScore;
    if (result.winTeam === 'draw') stats[name].draws++;
    else if (result.winTeam === row.teamName) stats[name].wins++;
    else stats[name].losses++;
  });

  const VIRTUAL = 5; // ベイズ補正の仮想試合数
  Object.values(stats).forEach(s => {
    // 補正勝率: 試合数が少ない人は50%に近づく
    s.winRate = Math.round(((s.wins + VIRTUAL * 0.5) / (s.games + VIRTUAL)) * 100);
    // 平均得点差
    s.avgScoreDiff = s.games > 0 ? s.scoreDiffSum / s.games : 0;
  });
  return stats;
}

// ── マイページ ──

function getMyPage(idToken) {
  const profile = verifyIdToken_(idToken);
  const sheet = getRosterSheet_();
  const headers = getSheetHeaders_(sheet);
  const rows = readObjects_(sheet);
  const idx = rows.findIndex(r => r[MC.LINE_ID] === profile.sub && isActiveMember_(r));
  if (idx < 0) throw new Error('名簿に未登録です。');
  const member = rows[idx];

  // LINE表示名を自動更新
  if (profile.name && profile.name !== member[MC.LINE_NAME]) {
    setColValue_(sheet, headers, idx + 2, MC.LINE_NAME, profile.name);
  }

  const winRates = calcWinRates_();
  const wr = winRates[member[MC.FULL_NAME]] || { wins:0, losses:0, draws:0, games:0, winRate:0 };

  // 直近10ゲーム履歴
  const opsSS = getOpsSS_();
  const gsRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.GAMESETS));
  const grRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.GAMERESULTS));
  const resultMap = {};
  grRows.forEach(r => { resultMap[r.sessionId + '|' + r.gameNumber] = r; });
  const myGames = gsRows
    .filter(r => r.fullName === member[MC.FULL_NAME] && resultMap[r.sessionId + '|' + r.gameNumber])
    .slice(-10)
    .map(r => {
      const res = resultMap[r.sessionId + '|' + r.gameNumber];
      const outcome = res.winTeam === 'draw' ? 'draw' : res.winTeam === r.teamName ? 'win' : 'loss';
      return { date: r.eventDate, game: r.gameNumber, outcome };
    });

  return {
    ok: true,
    profile: {
      fullName:    member[MC.FULL_NAME]    || '',
      furigana:    member[MC.FURIGANA]     || '',
      gender:      member[MC.GENDER]       || '',
      birthDate:   member[MC.BIRTH_DATE]   || '',
      ageApril1:   member[MC.AGE_APRIL1]   || '',
      mobilePhone: member[MC.MOBILE_PHONE] || '',
      homePhone:   member[MC.HOME_PHONE]   || '',
      address:     member[MC.ADDRESS]      || '',
      lineName:    profile.name            || '',
    },
    stats: wr,
    recentGames: myGames,
  };
}

function updateProfile(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const sheet = getRosterSheet_();
    const headers = getSheetHeaders_(sheet);
    const rows = readObjects_(sheet);
    const idx = rows.findIndex(r => r[MC.LINE_ID] === profile.sub);
    if (idx < 0) throw new Error('名簿に未登録です。');
    const rowNum = idx + 2;
    const required = ['fullName','furigana','gender','birthDate','address','mobilePhone'];
    for (const f of required) {
      if (payload[f] !== undefined && !String(payload[f]).trim()) throw new Error('必須項目が未入力です。');
    }
    const map = [
      ['fullName',    MC.FULL_NAME],
      ['furigana',    MC.FURIGANA],
      ['gender',      MC.GENDER],
      ['birthDate',   MC.BIRTH_DATE],
      ['address',     MC.ADDRESS],
      ['mobilePhone', MC.MOBILE_PHONE],
      ['homePhone',   MC.HOME_PHONE],
    ];
    map.forEach(([key, col]) => {
      if (payload[key] !== undefined) setColValue_(sheet, headers, rowNum, col, String(payload[key]).trim());
    });
    // 4/1年齢を再計算
    const bd = String(payload.birthDate || rows[idx][MC.BIRTH_DATE] || '').trim();
    if (bd) {
      const bdDate = new Date(bd);
      const now = new Date();
      const april1 = new Date(now.getFullYear(), 3, 1);
      const age = Math.floor((april1 - bdDate) / (365.25 * 24 * 60 * 60 * 1000));
      setColValue_(sheet, headers, rowNum, MC.AGE_APRIL1, age);
    }
    return { ok: true, message: 'プロフィールを更新しました。' };
  } finally { lock.releaseLock(); }
}

// ── 一般メンバー: 自分の当日参加チェック ──

function markSelfAttendance(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    const member = findMemberByLineId_(profile.sub);
    if (!member) throw new Error('名簿に未登録です。');
    const sessionId = String(payload.sessionId || '');
    const schedMap = getScheduleMap_();
    const ses = schedMap[sessionId];
    if (!ses) throw new Error('予定が見つかりません。');
    // 開催日当日のみ許可
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (ses.eventDate !== today) throw new Error('当日参加チェックは開催日当日のみ有効です。');
    const sheet = getOpsSS_().getSheetByName(SHEET_NAMES.RESPONSES);
    const rows = readObjects_(sheet);
    const idx = rows.findIndex(r => r.lineId === profile.sub && r.sessionId === sessionId);
    if (idx >= 0) {
      rows[idx].answer = 'attended';
      rows[idx].submittedAt = fmtJst_(new Date());
      writeRow_(sheet, idx + 2, HEADERS.Responses, rows[idx]);
    } else {
      // 参加予定未登録の場合も新規追加
      const record = {
        lineId: profile.sub, no: member[MC.NO] || '', fullName: member[MC.FULL_NAME],
        sessionId: ses.sessionId, monthKey: ses.monthKey, eventDate: ses.eventDate,
        title: ses.title, answer: 'attended', note: '',
        lineDisplayName: profile.name || '', submittedAt: fmtJst_(new Date()),
      };
      appendRow_(sheet, HEADERS.Responses, record);
    }
    updateActivityReport_(sessionId);
    return { ok: true, schedule: buildScheduleView_(null, profile.sub), message: '当日参加を登録しました。' };
  } finally { lock.releaseLock(); }
}

// ── 管理者専用ページデータ ──

function getAdminPage(idToken) {
  const profile = verifyIdToken_(idToken);
  ensureAdmin_(profile.sub);
  const schedRows = readObjects_(getOpsSS_().getSheetByName(SHEET_NAMES.SCHEDULE));
  return { ok: true, sessions: schedRows };
}

// ── 候補日一括登録 ──

function addScheduleSessions(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    ensureAdmin_(profile.sub);
    const dates = Array.isArray(payload.dates) ? payload.dates : [];
    if (!dates.length) throw new Error('日程が指定されていません。');
    const sheet = getOpsSS_().getSheetByName(SHEET_NAMES.SCHEDULE);
    const existing = readObjects_(sheet).map(r => r.eventDate);
    let added = 0;
    dates.forEach(d => {
      if (existing.includes(d)) return;
      const monthKey = d.slice(0, 7);
      appendRow_(sheet, HEADERS.Schedule, {
        sessionId: 'S' + Utilities.getUuid().slice(0, 8),
        monthKey, eventDate: d, title: '定例活動',
        minAttendees: 3, maxAttendees: 40, openForResponse: true, note: '',
      });
      added++;
    });
    return { ok: true, message: added + '件の日程を登録しました。' };
  } finally { lock.releaseLock(); }
}

// ── 候補日削除 ──

function deleteScheduleSession(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    ensureAdmin_(profile.sub);
    const sessionId = String(payload.sessionId || '');
    const sheet = getOpsSS_().getSheetByName(SHEET_NAMES.SCHEDULE);
    const rows = readObjects_(sheet);
    const idx = rows.findIndex(r => r.sessionId === sessionId);
    if (idx < 0) throw new Error('小指定の予定が見つかりません。');
    sheet.deleteRow(idx + 2);
    return { ok: true, message: '予定を削除しました。' };
  } finally { lock.releaseLock(); }
}

// ── 活動報告書自動更新 ──

function updateActivityReport_(sessionId) {
  const opsSS = getOpsSS_();
  const schedMap = getScheduleMap_();
  const ses = schedMap[sessionId];
  if (!ses) return;
  const monthKey = ses.monthKey;
  const eventDate = ses.eventDate;

  // attended の参加者を取得
  const responses = readObjects_(opsSS.getSheetByName(SHEET_NAMES.RESPONSES));
  const attendees = responses.filter(r => r.sessionId === sessionId && r.answer === 'attended');

  // 名簿から生年月日・ステータス取得
  const members = readObjects_(getRosterSheet_());
  const memberMap = {};
  members.forEach(m => { if (m[MC.LINE_ID]) memberMap[m[MC.LINE_ID]] = m; });

  // 開催日時点の実年齢を計算
  const evtDate = new Date(eventDate);
  function calcAge(birthDateStr) {
    if (!birthDateStr) return null;
    const bd = new Date(birthDateStr);
    let age = evtDate.getFullYear() - bd.getFullYear();
    const m = evtDate.getMonth() - bd.getMonth();
    if (m < 0 || (m === 0 && evtDate.getDate() < bd.getDate())) age--;
    return age;
  }

  // 年齢帯別・会員別集計
  // 列: [0-6会員, 0-6会員外, 7-15会員, 7-15会員外, 16-30会員, 16-30会員外, 31-59会員, 31-59会員外, 60+会員, 60+会員外]
  const counts = [0,0,0,0,0,0,0,0,0,0];
  attendees.forEach(a => {
    const m = memberMap[a.lineId] || {};
    const isTrial = String(m[MC.STATUS] || '').trim() === '体験';
    const age = calcAge(m[MC.BIRTH_DATE]);
    let band;
    if (age === null) band = 3; // 不明なら成人扱い
    else if (age <= 6)  band = 0;
    else if (age <= 15) band = 1;
    else if (age <= 30) band = 2;
    else if (age <= 59) band = 3;
    else                band = 4;
    const col = band * 2 + (isTrial ? 1 : 0);
    counts[col]++;
  });
  const total = counts.reduce((s, v) => s + v, 0);

  // Report_{monthKey} シートを取得 or 作成
  const reportName = 'Report_' + monthKey;
  let rSheet = opsSS.getSheetByName(reportName);
  if (!rSheet) {
    // ExportTemplateからコピー
    const tmpl = opsSS.getSheetByName(SHEET_NAMES.EXPORT);
    rSheet = tmpl.copyTo(opsSS);
    rSheet.setName(reportName);
    // 計行をクリア
    if (rSheet.getLastRow() >= 4) rSheet.getRange(4, 1, rSheet.getLastRow() - 3, 12).clearContent();
  }

  // 該当開催日の行を探す
  const lr = rSheet.getLastRow();
  let targetRow = -1;
  if (lr >= 4) {
    const col1 = rSheet.getRange(4, 1, lr - 3, 1).getValues();
    col1.forEach((v, i) => { if (String(v[0]).trim() === eventDate) targetRow = i + 4; });
  }
  // なければ計行の前に挿入
  if (targetRow < 0) {
    // 計行を探す（A列が「計」の行）
    let keRow = -1;
    if (lr >= 4) {
      const col1 = rSheet.getRange(4, 1, lr - 3, 1).getValues();
      col1.forEach((v, i) => { if (String(v[0]).trim() === '計') keRow = i + 4; });
    }
    if (keRow > 0) {
      rSheet.insertRowBefore(keRow);
      targetRow = keRow;
    } else {
      targetRow = lr + 1;
    }
  }
  rSheet.getRange(targetRow, 1, 1, 12).setValues([[eventDate, ...counts, total]]);

  // 計行を再計算
  const dataLr = rSheet.getLastRow();
  let keRow2 = -1;
  if (dataLr >= 4) {
    const col1 = rSheet.getRange(4, 1, dataLr - 3, 1).getValues();
    col1.forEach((v, i) => { if (String(v[0]).trim() === '計') keRow2 = i + 4; });
  }
  if (keRow2 > 0) {
    const totals = Array(11).fill(0);
    for (let row = 4; row < keRow2; row++) {
      const vals = rSheet.getRange(row, 2, 1, 11).getValues()[0];
      vals.forEach((v, i) => { totals[i] += Number(v) || 0; });
    }
    rSheet.getRange(keRow2, 2, 1, 11).setValues([totals]);
  }
}

function generateActivityReport(payload) {
  const profile = verifyIdToken_(payload.idToken);
  ensureAdmin_(profile.sub);
  const monthKey = String(payload.monthKey || '');
  if (!monthKey) throw new Error('月を指定してください。');
  // その月の全セッションを再生成
  const schedMap = getScheduleMap_();
  const sessions = Object.values(schedMap).filter(s => s.monthKey === monthKey);
  sessions.forEach(s => updateActivityReport_(s.sessionId));
  const reportName = 'Report_' + monthKey;
  const url = getOpsSS_().getUrl() + '#gid=' + (getOpsSS_().getSheetByName(reportName) || {}).getSheetId();
  return { ok: true, message: monthKey + 'の活動報告書を生成しました。', sheetName: reportName };
}

// ── 管理者代理登録 ──

function adminRegisterMember(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const profile = verifyIdToken_(payload.idToken);
    ensureAdmin_(profile.sub);
    const fullName = String(payload.fullName || '').trim();
    if (!fullName) throw new Error('氏名は必須です。');

    const sheet = getRosterSheet_();
    const rows = readObjects_(sheet);
    const headers = getSheetHeaders_(sheet);

    const normalize = s => String(s||'').replace(/[\s　]+/g, '');
    if (rows.some(r => normalize(r[MC.FULL_NAME]) === normalize(fullName))) {
      throw new Error('「' + fullName + '」は既に名簿に登録されています。');
    }

    const maxNo = rows.reduce((mx, r) => Math.max(mx, Number(r[MC.NO]) || 0), 0);
    const newNo = maxNo + 1;

    let ageApril1 = '';
    const birthDate = String(payload.birthDate || '').trim();
    if (birthDate) {
      const bd = new Date(birthDate);
      const april1 = new Date(new Date().getFullYear(), 3, 1);
      ageApril1 = Math.floor((april1 - bd) / (365.25 * 24 * 60 * 60 * 1000));
    }

    const newRow = headers.map(h => {
      switch(h) {
        case MC.STATUS:       return '入会';
        case MC.NO:           return newNo;
        case MC.FULL_NAME:    return fullName;
        case MC.FURIGANA:     return String(payload.furigana || '').trim();
        case MC.GENDER:       return String(payload.gender || '').trim();
        case MC.BIRTH_DATE:   return birthDate;
        case MC.AGE_APRIL1:   return ageApril1;
        case MC.MOBILE_PHONE: return String(payload.mobilePhone || '').trim();
        case MC.ADDRESS:      return String(payload.address || '').trim();
        case MC.NOTE:         return fmtJst_(new Date()) + ' 管理者代理登録';
        default:              return '';
      }
    });
    appendRowRaw_(sheet, headers, newRow);
    return { ok: true, message: '「' + fullName + '」を代理登録しました。' };
  } finally { lock.releaseLock(); }
}

// ── 共有資料一覧 ──

function getDocs(idToken) {
  verifyIdToken_(idToken);
  const folderId = getReqProp_('DOCS_FOLDER_ID');
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.PDF);
  const docs = [];
  while (files.hasNext()) {
    const f = files.next();
    docs.push({
      id: f.getId(),
      name: f.getName(),
      url: 'https://drive.google.com/file/d/' + f.getId() + '/view',
      updatedAt: Utilities.formatDate(f.getLastUpdated(), 'Asia/Tokyo', 'yyyy-MM-dd'),
    });
  }
  docs.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
  return { ok: true, docs };
}

// ── 管理者: 提出用リスト生成 ──

function generateExportList(payload) {
  const profile = verifyIdToken_(payload.idToken);
  ensureAdmin_(profile.sub);
  const sessionId = String(payload.sessionId || '');

  const opsSS = getOpsSS_();
  const responses = readObjects_(opsSS.getSheetByName(SHEET_NAMES.RESPONSES));
  const attendees = responses
    .filter(r => r.sessionId === sessionId && (r.answer === 'yes' || r.answer === 'attended'))
    .sort((a, b) => String(a.fullName).localeCompare(String(b.fullName), 'ja'));

  const members = readObjects_(getRosterSheet_());
  const memberMap = {};
  members.forEach(m => { if (m[MC.LINE_ID]) memberMap[m[MC.LINE_ID]] = m; });

  const list = attendees.map(a => {
    const m = memberMap[a.lineId] || {};
    return { no: a.no || m[MC.NO], fullName: a.fullName, furigana: m[MC.FURIGANA] || '', gender: m[MC.GENDER] || '', ageApril1: m[MC.AGE_APRIL1] || '' };
  });

  const exportSheet = opsSS.getSheetByName(SHEET_NAMES.EXPORT);
  const startRow = 3;
  const eventDate = attendees[0] ? attendees[0].eventDate : '';
  exportSheet.getRange('A1').setValue('参加者リスト — ' + eventDate);
  exportSheet.getRange('A2').setValue('No,氏名,ふりがな,性別,年齢');
  if (exportSheet.getLastRow() >= startRow) {
    exportSheet.getRange(startRow, 1, exportSheet.getLastRow() - startRow + 1, 5).clearContent();
  }
  if (list.length) {
    const data = list.map(l => [l.no, l.fullName, l.furigana, l.gender, l.ageApril1]);
    exportSheet.getRange(startRow, 1, data.length, 5).setValues(data);
  }

  return { ok: true, count: list.length, message: list.length + '名の参加者リストをExportTemplateシートに出力しました。' };
}

// ── スケジュールビュー ──

function buildScheduleView_(monthKey, lineId) {
  const opsSS = getOpsSS_();
  const schedRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.SCHEDULE))
    .filter(r => r.sessionId && String(r.openForResponse).toLowerCase() !== 'false');
  const respRows = readObjects_(opsSS.getSheetByName(SHEET_NAMES.RESPONSES));

  const grouped = {};
  schedRows.forEach(r => {
    if (monthKey && r.monthKey !== monthKey) return;
    const counts = countAnswers_(respRows, r.sessionId);
    const myResp = lineId ? respRows.find(rr => rr.sessionId === r.sessionId && rr.lineId === lineId) : null;
    if (!grouped[r.monthKey]) grouped[r.monthKey] = [];
    grouped[r.monthKey].push({
      sessionId: r.sessionId, monthKey: r.monthKey, eventDate: r.eventDate,
      title: r.title, minAttendees: Number(r.minAttendees||0), maxAttendees: Number(r.maxAttendees||0),
      note: r.note || '', counts,
      statusLabel: buildStatusLabel_(counts, Number(r.minAttendees||0), Number(r.maxAttendees||0)),
      attendees: listAttendees_(respRows, r.sessionId),
      myAnswer: myResp ? myResp.answer : 'undecided',
      myNote: myResp ? myResp.note : '',
    });
  });
  return Object.keys(grouped).sort().map(k => ({
    monthKey: k,
    sessions: grouped[k].sort((a, b) => String(a.eventDate).localeCompare(String(b.eventDate))),
  }));
}

function listAttendees_(respRows, sessionId) {
  return respRows
    .filter(r => r.sessionId === sessionId && (r.answer === 'yes' || r.answer === 'attended'))
    .map(r => ({ lineId: r.lineId, no: r.no, fullName: r.fullName, answer: r.answer }))
    .sort((a, b) => String(a.fullName).localeCompare(String(b.fullName), 'ja'));
}

function countAnswers_(respRows, sessionId) {
  return respRows.reduce((acc, r) => {
    if (r.sessionId !== sessionId) return acc;
    const a = normalizeAnswer_(r.answer);
    if (a === 'yes') acc.yes++;
    else if (a === 'no') acc.no++;
    else if (a === 'undecided') acc.undecided++;
    else if (a === 'attended') { acc.yes++; acc.attended++; }
    return acc;
  }, { yes: 0, no: 0, undecided: 0, attended: 0 });
}

function buildStatusLabel_(counts, min, max) {
  if (max && counts.yes >= max) return '定員到達';
  if (min && counts.yes < min) return '開催判断待ち';
  return '開催見込み';
}

// ── ヘルパー ──

function findMemberByLineId_(lineId) {
  const rows = readObjects_(getRosterSheet_());
  return rows.find(r => r[MC.LINE_ID] === lineId && isActiveMember_(r)) || null;
}

function isActiveMember_(r) {
  const s = String(r[MC.STATUS] || '').trim();
  return s !== '退会' && s !== '休会';
}

function getScheduleMap_() {
  return readObjects_(getOpsSS_().getSheetByName(SHEET_NAMES.SCHEDULE))
    .reduce((m, r) => { if (r.sessionId) m[r.sessionId] = r; return m; }, {});
}

function verifyIdToken_(idToken) {
  if (!idToken) throw new Error('IDトークンを取得できませんでした。');
  const chId = getReqProp_('LINE_LOGIN_CHANNEL_ID');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', contentType: 'application/x-www-form-urlencoded',
    payload: { id_token: idToken, client_id: chId }, muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('LINE IDトークン検証失敗: ' + (body.error_description || body.message || res.getResponseCode()));
  return body;
}

function checkAdmin_(uid) {
  const raw = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_IDS') || '';
  return raw.split(',').map(v => v.trim()).filter(Boolean).includes(uid);
}

function ensureAdmin_(uid) {
  if (!checkAdmin_(uid)) throw new Error('管理者権限がありません。');
}

function clearSessionRows_(sheet, sessionId) {
  const rows = readObjects_(sheet);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i].sessionId === sessionId) sheet.deleteRow(i + 2);
  }
}

function normalizeAnswer_(a) {
  const v = String(a || '').trim().toLowerCase();
  return ['yes','no','undecided','attended'].includes(v) ? v : 'undecided';
}

// ── シート操作 ──

function createSheetIfMissing_(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.setFrozenRows(1);
  }
  return s;
}

function appendRow_(sheet, headers, obj) {
  appendRowRaw_(sheet, headers, headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function appendRowRaw_(sheet, headers, rowValues) {
  const rowNum = sheet.getLastRow() + 1;
  PHONE_COLS_.forEach(col => {
    const ci = headers.indexOf(col);
    if (ci >= 0) sheet.getRange(rowNum, ci + 1).setNumberFormat('@STRING@');
  });
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowValues]);
}

function writeRow_(sheet, rowNum, headers, obj) {
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([headers.map(h => obj[h] !== undefined ? obj[h] : '')]);
}

function readObjects_(sheet) {
  const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const vals = sheet.getRange(1, 1, lr, lc).getDisplayValues();
  const hdr = vals[0];
  return vals.slice(1).filter(r => r.some(c => c !== '')).map(r => hdr.reduce((o, h, i) => { o[h] = r[i]; return o; }, {}));
}

function getSheetHeaders_(sheet) {
  const lc = sheet.getLastColumn();
  if (lc < 1) return [];
  return sheet.getRange(1, 1, 1, lc).getDisplayValues()[0];
}

const PHONE_COLS_ = [MC.MOBILE_PHONE, MC.HOME_PHONE];

function setColValue_(sheet, headers, rowNum, colName, value) {
  const col = headers.indexOf(colName);
  if (col < 0) return;
  const cell = sheet.getRange(rowNum, col + 1);
  if (PHONE_COLS_.includes(colName)) cell.setNumberFormat('@STRING@');
  cell.setValue(value);
}

// ── スプレッドシート取得 ──

function getRosterSS_() { return SpreadsheetApp.openById(getReqProp_('ROSTER_SPREADSHEET_ID')); }
function getRosterSheet_() { return getRosterSS_().getSheetByName(SHEET_NAMES.MEMBERS); }
function getOpsSS_() { return SpreadsheetApp.openById(getReqProp_('SPREADSHEET_ID')); }

function getReqProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('スクリプトプロパティ ' + key + ' が未設定です。');
  return v;
}

function fmtJst_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); }
