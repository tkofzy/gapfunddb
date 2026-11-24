/**
 * スタートアップ支援・公募情報データベース — GAS バックエンド
 *
 * 構成:
 * - Public JSON API (doGet): キャッシュ付き・読み取り専用（internalMemo / sourceType 除外）
 * - Admin UI & RPCs (?page=admin): ADMIN_EMAIL または ADMIN_TOKEN による認証保護
 * - Background Crawler (dailyAutoCrawl): RSS 2.0 / Atom / HTML スクレイピング対応
 */

const SHEET_PROGRAMS = 'Programs';
const SHEET_TARGETS = 'CrawlTargets';
const SHEET_ANALYTICS = 'Analytics';
const CACHE_KEY_PUBLIC = 'PUBLIC_PROGRAMS_JSON_V1';
const CACHE_TTL_SECONDS = 300; // 5分キャッシュ

const ANALYTICS_HEADERS = ['date', 'type', 'label', 'count', 'updatedAt'];
const ALLOWED_ANALYTICS_TYPES = new Set(['pv', 'phase', 'search', 'filter', 'program_click']);

// [EDIT] オーナー設定・カスタマイズ項目
const CONFIG = {
  ADMIN_PAGE_TITLE: '管理画面 | スタートアップ支援データベース',
  DEFAULT_KEYWORDS: '公募|スタートアップ|補助金|アクセラ',
  CLOSING_SOON_DAYS: 14,
  TRIGGER_HOUR_JST: 7,
  DEFAULT_TARGETS: 'researcher,student'
};

const PROGRAM_HEADERS = [
  'id', 'title', 'org', 'phase', 'providerType', 'targets',
  'amount', 'deadline', 'status', 'url', 'summary',
  'internalMemo', 'sourceType', 'updatedAt'
];

const TARGET_HEADERS = [
  'org', 'url', 'phase', 'providerType', 'keywords', 'active',
  'type', 'excludeKeywords', 'lastFetched', 'lastStatus', 'lastError'
];

// ============================================================================
// 1. Web App エントリーポイント (doGet)
// ============================================================================

/**
 * Web アプリの GET リクエストを処理する。
 * - ?page=admin の場合: 管理画面 (admin.html) を返す
 * - それ以外 (デフォルト / ?api=programs): 公開用 JSON (または JSONP) を返す
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  if (params.page === 'admin') {
    const template = HtmlService.createTemplateFromFile('admin');
    return template.evaluate()
      .setTitle(CONFIG.ADMIN_PAGE_TITLE)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }

  const jsonStr = getPublicProgramsJson_();
  const callback = params.callback ? String(params.callback).trim() : '';

  if (callback) {
    if (!/^[a-zA-Z0-9_$.]+$/.test(callback)) {
      const errJson = JSON.stringify({ ok: false, error: 'Invalid callback parameter' });
      return ContentService.createTextOutput(errJson)
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(callback + '(' + jsonStr + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 公開フロントエンドからの軽量アクセス・関心イベント (POST text/plain JSON) を記録する。
 */
function doPost(e) {
  try {
    const rawBody = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
    if (!rawBody) {
      return jsonOutput_({ ok: false, error: 'Empty body' });
    }
    const payload = JSON.parse(rawBody);
    const type = String(payload.type || '').trim();
    if (!ALLOWED_ANALYTICS_TYPES.has(type)) {
      return jsonOutput_({ ok: false, error: 'Invalid event type' });
    }

    let label = String(payload.label || '')
      .replace(/[\x00-\x1F\x7F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80);

    if (!label && type === 'pv') {
      label = 'home';
    }
    if (!label) {
      return jsonOutput_({ ok: false, error: 'Empty label' });
    }

    recordAnalyticsEvent_(type, label);
    return jsonOutput_({ ok: true });
  } catch (err) {
    console.warn('doPost error: ' + err.message);
    return jsonOutput_({ ok: false, error: 'Invalid request' });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * (date, type, label) 単位で Analytics シートにカウントを集約する。
 */
function recordAnalyticsEvent_(type, label) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ANALYTICS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ANALYTICS);
      sheet.appendRow(ANALYTICS_HEADERS);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(ANALYTICS_HEADERS);
    }

    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
    const nowTimeStr = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm');

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        const rowDate = rows[i][0] instanceof Date
          ? Utilities.formatDate(rows[i][0], 'JST', 'yyyy-MM-dd')
          : String(rows[i][0] || '').trim();
        const rowType = String(rows[i][1] || '').trim();
        const rowLabel = String(rows[i][2] || '').trim();
        if (rowDate === todayStr && rowType === type && rowLabel === label) {
          const nextCount = (Number(rows[i][3]) || 0) + 1;
          sheet.getRange(i + 2, 4, 1, 2).setValues([[nextCount, nowTimeStr]]);
          return;
        }
      }
    }

    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[todayStr, type, label, 1, nowTimeStr]]);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 公開用 JSON 文字列をキャッシュから取得、またはスプレッドシートから生成する。
 * internalMemo と sourceType は公開出力から完全に除外する。
 */
function getPublicProgramsJson_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_PUBLIC);
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PROGRAMS);
  const props = PropertiesService.getScriptProperties();
  const lastUpdated = props.getProperty('LAST_SYNC_TIME') ||
    Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm');

  if (!sheet) {
    return JSON.stringify({ ok: true, lastUpdated: lastUpdated, count: 0, programs: [] });
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return JSON.stringify({ ok: true, lastUpdated: lastUpdated, count: 0, programs: [] });
  }

  const headers = rows.shift().map(h => String(h).trim());
  const todayStart = getTodayStartJst_();

  const programs = rows
    .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
    .map(row => {
      const raw = rowToObject_(headers, row);
      const deadline = String(raw.deadline || '');
      const updatedAtStr = String(raw.updatedAt || '');
      const computedStatus = computeProgramStatus_(raw.status, deadline, todayStart);
      return {
        id: String(raw.id || ''),
        title: String(raw.title || ''),
        org: String(raw.org || ''),
        phase: String(raw.phase || ''),
        providerType: String(raw.providerType || ''),
        targets: parseTargetsArray_(raw.targets),
        amount: String(raw.amount || ''),
        deadline: deadline,
        status: computedStatus,
        isNew: computeIsNew_(computedStatus, updatedAtStr, todayStart),
        url: String(raw.url || ''),
        summary: String(raw.summary || ''),
        updatedAt: updatedAtStr
      };
    })
    .reverse();

  const payload = JSON.stringify({
    ok: true,
    lastUpdated: lastUpdated,
    count: programs.length,
    programs: programs
  });

  try {
    cache.put(CACHE_KEY_PUBLIC, payload, CACHE_TTL_SECONDS);
  } catch (e) {
    console.warn('Cache put skipped (size limit): ' + e.message);
  }

  return payload;
}

/**
 * 公開キャッシュを無効化する。
 */
function invalidatePublicCache_() {
  try {
    CacheService.getScriptCache().remove(CACHE_KEY_PUBLIC);
  } catch (e) {
    console.warn('Cache remove failed: ' + e.message);
  }
}

// ============================================================================
// 2. 管理者認証 & 保護された RPC (admin.html から呼び出し)
// ============================================================================

/**
 * [EDIT] 初回セットアップ用ヘルパー（GAS エディタから手動実行可）
 * ADMIN_EMAIL と ADMIN_TOKEN をスクリプトプロパティに保存する。
 */
function setupAdminAuth(adminEmail, adminToken) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_EMAIL') || props.getProperty('ADMIN_TOKEN')) {
    throw new Error('既に管理者認証が設定されています。変更する場合は GAS プロジェクト設定のスクリプトプロパティから更新してください。');
  }
  if (adminEmail) {
    props.setProperty('ADMIN_EMAIL', String(adminEmail).trim().toLowerCase());
  }
  if (adminToken) {
    props.setProperty('ADMIN_TOKEN', String(adminToken).trim());
  }
  return {
    hasEmail: Boolean(props.getProperty('ADMIN_EMAIL')),
    hasToken: Boolean(props.getProperty('ADMIN_TOKEN'))
  };
}

/**
 * 呼び出し元が管理者であることを検証する。
 * ADMIN_EMAIL (Google ログインメール) または ADMIN_TOKEN の一致が必要。
 */
function assertAdmin_(token) {
  const props = PropertiesService.getScriptProperties();
  const configuredEmail = String(props.getProperty('ADMIN_EMAIL') || '').trim().toLowerCase();
  const configuredToken = String(props.getProperty('ADMIN_TOKEN') || '').trim();

  let activeEmail = '';
  try {
    activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    activeEmail = '';
  }

  if (configuredEmail && activeEmail && activeEmail === configuredEmail) {
    return { method: 'email', email: activeEmail };
  }

  if (!configuredToken && !configuredEmail) {
    throw new Error('管理認証が未設定です。GAS エディタで setupAdminAuth("you@example.com", "your-secret-token") を実行してください。');
  }

  const provided = String(token || '').trim();
  if (configuredToken && provided && provided === configuredToken) {
    return { method: 'token' };
  }

  throw new Error('認証エラー: 管理者トークンが無効です。');
}

/**
 * 管理画面用の全データ（internalMemo, sourceType, CrawlTargets 含む）を取得する。
 */
function getAdminData(token) {
  const auth = assertAdmin_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName(SHEET_PROGRAMS);
  const tSheet = ss.getSheetByName(SHEET_TARGETS);
  const props = PropertiesService.getScriptProperties();

  const programs = [];
  if (pSheet && pSheet.getLastRow() > 1) {
    const pRows = pSheet.getDataRange().getValues();
    const pHeaders = pRows.shift().map(h => String(h).trim());
    const todayStart = getTodayStartJst_();

    pRows.forEach((row, idx) => {
      if (row[0] === '' || row[0] === null) return;
      const obj = rowToObject_(pHeaders, row);
      obj.targets = parseTargetsArray_(obj.targets);
      obj.computedStatus = computeProgramStatus_(obj.status, obj.deadline, todayStart);
      obj.rowIndex = idx + 2;
      programs.push(obj);
    });
  }

  const targets = [];
  if (tSheet && tSheet.getLastRow() > 1) {
    ensureTargetHeaders_(tSheet);
    const tRows = tSheet.getDataRange().getValues();
    const tHeaders = tRows.shift().map(h => String(h).trim());
    tRows.forEach((row, idx) => {
      if (!row[0] && !row[1]) return;
      const obj = rowToObject_(tHeaders, row);
      obj.active = String(obj.active).toUpperCase() === 'TRUE';
      obj.type = String(obj.type || 'auto');
      obj.rowIndex = idx + 2;
      targets.push(obj);
    });
  }

  return {
    ok: true,
    authMethod: auth.method,
    lastSync: props.getProperty('LAST_SYNC') || '未実行',
    programs: programs.reverse(),
    targets: targets,
    analytics: getAnalyticsSummary_(ss)
  };
}

/**
 * Analytics シートから PV と各種ランキング（上位10件）を集計する。
 */
function getAnalyticsSummary_(ss) {
  const empty = {
    todayPv: 0,
    last7DaysPv: 0,
    totalPv: 0,
    topPhases: [],
    topSearches: [],
    topFilters: [],
    topPrograms: []
  };

  const sheet = ss.getSheetByName(SHEET_ANALYTICS);
  if (!sheet || sheet.getLastRow() <= 1) return empty;

  const now = new Date();
  const todayStr = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
  const sevenDaysAgo = new Date(now.getTime() - 6 * 86400000);
  const sevenDaysAgoStr = Utilities.formatDate(sevenDaysAgo, 'JST', 'yyyy-MM-dd');

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  let todayPv = 0;
  let last7DaysPv = 0;
  let totalPv = 0;

  const maps = {
    phase: {},
    search: {},
    filter: {},
    program_click: {}
  };

  rows.forEach(row => {
    const dateStr = row[0] instanceof Date
      ? Utilities.formatDate(row[0], 'JST', 'yyyy-MM-dd')
      : String(row[0] || '').trim();
    const type = String(row[1] || '').trim();
    const label = String(row[2] || '').trim();
    const count = Number(row[3]) || 0;
    if (!type || !count) return;

    if (type === 'pv') {
      totalPv += count;
      if (dateStr === todayStr) todayPv += count;
      if (dateStr >= sevenDaysAgoStr && dateStr <= todayStr) last7DaysPv += count;
    } else if (maps[type] && label) {
      maps[type][label] = (maps[type][label] || 0) + count;
    }
  });

  const toTop10 = (map) => Object.keys(map)
    .map(label => ({ label: label, count: map[label] }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  return {
    todayPv: todayPv,
    last7DaysPv: last7DaysPv,
    totalPv: totalPv,
    topPhases: toTop10(maps.phase),
    topSearches: toTop10(maps.search),
    topFilters: toTop10(maps.filter),
    topPrograms: toTop10(maps.program_click)
  };
}

/**
 * プログラムの新規作成または更新を行う（要管理者認証）。
 */
function saveProgramData(item, token) {
  assertAdmin_(token);
  if (!item || !item.title) {
    throw new Error('タイトルは必須です。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PROGRAMS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROGRAMS);
    sheet.appendRow(PROGRAM_HEADERS);
  }

  const rows = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  const targetsStr = Array.isArray(item.targets) ? item.targets.join(',') : String(item.targets || '');

  const buildRow = (id, sourceType) => [
    id,
    String(item.title || '').trim(),
    String(item.org || '').trim(),
    String(item.phase || 'p1'),
    String(item.providerType || 'public'),
    targetsStr,
    String(item.amount || '').trim(),
    String(item.deadline || '').trim(),
    String(item.status || 'open'),
    String(item.url || '').trim(),
    String(item.summary || '').trim(),
    String(item.internalMemo || '').trim(),
    sourceType,
    nowStr
  ];

  if (item.id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(item.id)) {
        const rowData = buildRow(item.id, item.sourceType || String(rows[i][12] || 'manual'));
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        invalidatePublicCache_();
        return getAdminData(token);
      }
    }
  }

  const newId = 'm_' + new Date().getTime();
  const newRow = buildRow(newId, 'manual');
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, newRow.length).setValues([newRow]);
  invalidatePublicCache_();
  return getAdminData(token);
}

/**
 * プログラムを削除する（要管理者認証）。
 */
function deleteProgramData(id, token) {
  assertAdmin_(token);
  if (!id) throw new Error('ID が指定されていません。');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PROGRAMS);
  if (!sheet) throw new Error('Programs シートが見つかりません。');

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      invalidatePublicCache_();
      return getAdminData(token);
    }
  }
  throw new Error('対象のプログラムが見つかりません: ' + id);
}

/**
 * CrawlTargets の追加または更新を行う（要管理者認証）。
 */
function saveCrawlTarget(target, token) {
  assertAdmin_(token);
  if (!target || !target.org || !target.url) {
    throw new Error('機関名と URL は必須です。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_TARGETS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TARGETS);
    sheet.appendRow(TARGET_HEADERS);
  } else {
    ensureTargetHeaders_(sheet);
  }

  const rowValues = [
    String(target.org || '').trim(),
    String(target.url || '').trim(),
    String(target.phase || 'p2'),
    String(target.providerType || 'public'),
    String(target.keywords || '').trim(),
    target.active ? 'TRUE' : 'FALSE',
    String(target.type || 'auto').trim(),
    String(target.excludeKeywords || '').trim()
  ];

  const rowIndex = Number(target.rowIndex);
  if (rowIndex >= 2 && rowIndex <= sheet.getLastRow()) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    const fullRow = rowValues.concat(['', '', '']);
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, fullRow.length).setValues([fullRow]);
  }

  return getAdminData(token);
}

/**
 * CrawlTargets から行を削除する（要管理者認証）。
 */
function deleteCrawlTarget(rowIndex, token) {
  assertAdmin_(token);
  const idx = Number(rowIndex);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TARGETS);
  if (!sheet || idx < 2 || idx > sheet.getLastRow()) {
    throw new Error('無効な行番号です。');
  }
  sheet.deleteRow(idx);
  return getAdminData(token);
}

/**
 * 管理画面からの手動巡回トリガー（要管理者認証）。
 */
function triggerManualCrawl(token) {
  assertAdmin_(token);
  const crawlResult = dailyAutoCrawl();
  const adminData = getAdminData(token);
  adminData.crawlResult = crawlResult;
  return adminData;
}

// ============================================================================
// 3. 高精度データ収集 (dailyAutoCrawl) & トリガー設定
// ============================================================================

/**
 * 毎日自動巡回を実行する（時間主導型トリガー & 管理画面手動実行用）。
 * RSS 2.0 / Atom / HTML ページに対応し、重複排除・日付/金額抽出・バッチ書込を行う。
 */
function dailyAutoCrawl() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn('dailyAutoCrawl: 別の処理が実行中のためスキップしました');
    return { ok: false, skipped: true, reason: 'locked' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_PROGRAMS);
    const tSheet = ss.getSheetByName(SHEET_TARGETS);
    if (!pSheet || !tSheet) {
      return { ok: false, error: 'Programs または CrawlTargets シートが存在しません' };
    }

    ensureTargetHeaders_(tSheet);

    const existingUrlSet = new Set();
    if (pSheet.getLastRow() > 1) {
      const pRows = pSheet.getDataRange().getValues();
      for (let i = 1; i < pRows.length; i++) {
        const norm = normalizeUrl_(String(pRows[i][9] || ''));
        if (norm) existingUrlSet.add(norm);
      }
    }

    const tRows = tSheet.getDataRange().getValues();
    if (tRows.length <= 1) {
      invalidatePublicCache_();
      return { ok: true, addedCount: 0, targetResults: [] };
    }

    const nowDate = new Date();
    const nowStr = Utilities.formatDate(nowDate, 'JST', 'yyyy-MM-dd');
    const nowTimeStr = Utilities.formatDate(nowDate, 'JST', 'yyyy-MM-dd HH:mm');
    const newRows = [];
    const targetStatusUpdates = []; // { rowIndex, lastFetched, lastStatus, lastError }
    const targetResults = [];

    for (let i = 1; i < tRows.length; i++) {
      const row = tRows[i];
      const org = String(row[0] || '').trim();
      const targetUrl = String(row[1] || '').trim();
      const phase = String(row[2] || 'p2').trim();
      const providerType = String(row[3] || 'public').trim();
      const keywords = String(row[4] || '').trim();
      const active = String(row[5] || '').toUpperCase() === 'TRUE';
      const sourceType = String(row[6] || 'auto').trim().toLowerCase();
      const excludeKeywords = String(row[7] || '').trim();

      if (!active || !targetUrl) continue;

      try {
        const res = UrlFetchApp.fetch(targetUrl, {
          muteHttpExceptions: true,
          followRedirects: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (compatible; StartupPortalBot/1.0)' }
        });

        const httpCode = res.getResponseCode();
        if (httpCode < 200 || httpCode >= 300) {
          throw new Error('HTTP ' + httpCode);
        }

        const contentType = String(res.getHeaders()['Content-Type'] || '').toLowerCase();
        const body = res.getContentText();
        const format = detectFeedFormat_(sourceType, contentType, body);

        let entries = [];
        if (format === 'rss') {
          entries = parseRssItems_(body, targetUrl);
        } else if (format === 'atom') {
          entries = parseAtomItems_(body, targetUrl);
        } else {
          entries = parseHtmlItems_(body, targetUrl);
        }

        const includeRegex = buildKeywordRegex_(keywords, CONFIG.DEFAULT_KEYWORDS);
        const excludeRegex = excludeKeywords ? buildKeywordRegex_(excludeKeywords, '') : null;

        let targetAdded = 0;
        entries.forEach(entry => {
          const title = cleanWhitespace_(entry.title);
          const link = String(entry.link || '').trim();
          const summaryText = cleanWhitespace_(entry.summary || '');
          const searchText = title + ' ' + summaryText;

          if (!title || !link || !/^https?:\/\//i.test(link)) return;
          if (!includeRegex.test(searchText)) return;
          if (excludeRegex && excludeRegex.test(searchText)) return;

          const normUrl = normalizeUrl_(link);
          if (!normUrl || existingUrlSet.has(normUrl)) return;

          const deadline = extractDeadline_(searchText) || 'Webページ参照';
          const amount = extractAmount_(searchText) || '公募要領参照';
          const shortSummary = summaryText
            ? summaryText.slice(0, 140) + (summaryText.length > 140 ? '…' : '')
            : '自動巡回により検知された新着情報です。リンク先で詳細を確認してください。';

          const cleanOrg = org
            .replace(/\s*公募新着RSS$/i, '')
            .replace(/\s*新着情報RSS$/i, '')
            .trim();

          newRows.push([
            'auto_' + nowDate.getTime() + '_' + newRows.length,
            title,
            cleanOrg,
            phase || 'p2',
            providerType || 'public',
            CONFIG.DEFAULT_TARGETS,
            amount,
            deadline,
            'open',
            link,
            shortSummary,
            '【自動検知】対象要件・締切を確認してください',
            'auto',
            nowStr
          ]);

          existingUrlSet.add(normUrl);
          targetAdded++;
        });

        targetStatusUpdates.push({
          rowIndex: i + 1,
          lastFetched: nowTimeStr,
          lastStatus: 'ok (' + format + ', +' + targetAdded + ')',
          lastError: ''
        });
        targetResults.push({ org: org, format: format, added: targetAdded, status: 'ok' });
      } catch (e) {
        console.warn('巡回エラー (' + org + '): ' + e.message);
        targetStatusUpdates.push({
          rowIndex: i + 1,
          lastFetched: nowTimeStr,
          lastStatus: 'error',
          lastError: String(e.message || e).slice(0, 200)
        });
        targetResults.push({ org: org, status: 'error', error: e.message });
      }
    }

    if (newRows.length > 0) {
      pSheet.getRange(pSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    targetStatusUpdates.forEach(u => {
      tSheet.getRange(u.rowIndex, 9, 1, 3).setValues([[u.lastFetched, u.lastStatus, u.lastError]]);
    });

    const props = PropertiesService.getScriptProperties();
    props.setProperty('LAST_SYNC_TIME', nowTimeStr);
    props.setProperty('LAST_SYNC', nowTimeStr + ' (新規検知: ' + newRows.length + '件)');
    invalidatePublicCache_();

    return {
      ok: true,
      addedCount: newRows.length,
      lastSync: nowTimeStr,
      targetResults: targetResults
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * [EDIT] 毎日自動巡回トリガーを登録する（GAS エディタから1回手動実行）。
 * 重複トリガーを削除してから再登録する。
 */
function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'dailyAutoCrawl') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('dailyAutoCrawl')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR_JST)
    .inTimezone('Asia/Tokyo')
    .create();

  return 'dailyAutoCrawl トリガーを毎日 ' + CONFIG.TRIGGER_HOUR_JST + '時 (JST) に設定しました。';
}

// ============================================================================
// 4. フィード解析・HTMLスクレイピング・ユーティリティ
// ============================================================================

/**
 * CrawlTargets シートに拡張ヘッダー（7〜11列目）がなければ追加する。
 */
function ensureTargetHeaders_(tSheet) {
  const lastCol = tSheet.getLastColumn();
  if (lastCol >= TARGET_HEADERS.length) return;
  const headerRow = tSheet.getRange(1, 1, 1, TARGET_HEADERS.length).getValues()[0];
  let changed = false;
  for (let c = 0; c < TARGET_HEADERS.length; c++) {
    if (!headerRow[c]) {
      headerRow[c] = TARGET_HEADERS[c];
      changed = true;
    }
  }
  if (changed) {
    tSheet.getRange(1, 1, 1, TARGET_HEADERS.length).setValues([headerRow]);
  }
}

function detectFeedFormat_(configuredType, contentType, body) {
  if (configuredType === 'rss' || configuredType === 'atom' || configuredType === 'html') {
    return configuredType;
  }
  const head = body.slice(0, 1200);
  if (/<feed[\s>]/i.test(head) && /http:\/\/www\.w3\.org\/2005\/Atom|<entry[\s>]/i.test(body.slice(0, 4000))) {
    return 'atom';
  }
  if (/<rss[\s>]|<rdf:RDF[\s>]/i.test(head) || (contentType.includes('xml') && /<channel[\s>]/i.test(head))) {
    return 'rss';
  }
  if (contentType.includes('xml') || head.trim().startsWith('<?xml')) {
    return /<entry[\s>]/i.test(body.slice(0, 4000)) ? 'atom' : 'rss';
  }
  return 'html';
}

function parseRssItems_(xmlText, baseUrl) {
  const doc = XmlService.parse(xmlText);
  const root = doc.getRootElement();
  const items = [];

  const collectItems = (el) => {
    el.getChildren().forEach(child => {
      const name = child.getName();
      if (name === 'item') {
        items.push(child);
      } else if (name === 'channel') {
        collectItems(child);
      }
    });
  };
  collectItems(root);

  return items.map(item => {
    let title = '', link = '', summary = '';
    item.getChildren().forEach(c => {
      const n = c.getName();
      if (n === 'title') title = c.getText();
      else if (n === 'link') link = c.getText();
      else if (n === 'description' || n === 'encoded') {
        if (!summary) summary = stripHtml_(decodeHtmlEntities_(c.getText()));
      }
    });
    return {
      title: decodeHtmlEntities_(title),
      link: resolveUrl_(link.trim(), baseUrl),
      summary: summary
    };
  });
}

function parseAtomItems_(xmlText, baseUrl) {
  const doc = XmlService.parse(xmlText);
  const root = doc.getRootElement();
  const entries = root.getChildren().filter(c => c.getName() === 'entry');

  return entries.map(entry => {
    let title = '', link = '', summary = '';
    entry.getChildren().forEach(c => {
      const n = c.getName();
      if (n === 'title') {
        title = c.getText();
      } else if (n === 'link') {
        const relAttr = c.getAttribute('rel');
        const rel = relAttr ? relAttr.getValue() : 'alternate';
        const hrefAttr = c.getAttribute('href');
        const href = hrefAttr ? hrefAttr.getValue() : c.getText();
        if (!link || rel === 'alternate') {
          link = href;
        }
      } else if (n === 'summary' || n === 'content') {
        if (!summary) summary = stripHtml_(decodeHtmlEntities_(c.getText()));
      }
    });
    return {
      title: decodeHtmlEntities_(title),
      link: resolveUrl_(link.trim(), baseUrl),
      summary: summary
    };
  });
}

function parseHtmlItems_(html, baseUrl) {
  const cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<header[\s\S]*?<\/header>/gi, ' ')
    .replace(/<footer[\s\S]*?<\/footer>/gi, ' ')
    .replace(/<nav[\s\S]*?<\/nav>/gi, ' ');

  const linkRe = /<a\s+[^>]*href\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>([\s\S]*?)<\/a>/gi;
  const ignoreTextRe = /^(?:詳細|こちら|詳しく見る|もっと見る|TOP|トップ|ホーム|戻る|次へ|前へ|お問い合わせ|プライバシーポリシー|サイトマップ|ログイン|English)$/i;

  const results = [];
  let match;
  while ((match = linkRe.exec(cleaned)) !== null) {
    const rawHref = (match[1] || match[2] || match[3] || '').trim();
    if (!rawHref || /^(?:#|javascript:|mailto:|tel:)/i.test(rawHref)) continue;

    const absUrl = resolveUrl_(decodeHtmlEntities_(rawHref), baseUrl);
    if (!absUrl) continue;

    const linkText = cleanWhitespace_(stripHtml_(decodeHtmlEntities_(match[4])));
    if (linkText.length < 6 || ignoreTextRe.test(linkText)) continue;

    const startIdx = Math.max(0, match.index - 140);
    const endIdx = Math.min(cleaned.length, linkRe.lastIndex + 140);
    const context = cleanWhitespace_(stripHtml_(decodeHtmlEntities_(cleaned.slice(startIdx, endIdx))));

    results.push({
      title: linkText,
      link: absUrl,
      summary: context
    });
  }
  return results;
}

function resolveUrl_(href, baseUrl) {
  if (!href) return '';
  if (/^https?:\/\//i.test(href)) return href;

  const m = String(baseUrl).match(/^(https?:)\/\/([^\/?#]+)(\/[^?#]*)?/i);
  if (!m) return '';
  const protocol = m[1];
  const host = m[2];
  const basePath = m[3] || '/';

  if (href.startsWith('//')) return protocol + href;
  if (href.startsWith('/')) return protocol + '//' + host + href;

  const dir = basePath.endsWith('/') ? basePath : basePath.replace(/[^\/]*$/, '');
  const combined = (dir + href).split('/');
  const stack = [];
  combined.forEach(seg => {
    if (seg === '' || seg === '.') return;
    if (seg === '..') {
      if (stack.length > 0) stack.pop();
    } else {
      stack.push(seg);
    }
  });
  return protocol + '//' + host + '/' + stack.join('/');
}

function normalizeUrl_(rawUrl) {
  if (!rawUrl || !/^https?:\/\//i.test(rawUrl)) return '';
  const noHash = rawUrl.split('#')[0];
  const qIdx = noHash.indexOf('?');
  let base = qIdx >= 0 ? noHash.slice(0, qIdx) : noHash;
  const query = qIdx >= 0 ? noHash.slice(qIdx + 1) : '';

  base = base.replace(/^https?:\/\//i, '').toLowerCase();
  if (base.length > 1 && base.endsWith('/')) {
    base = base.slice(0, -1);
  }

  if (!query) return base;
  const trackingParams = new Set([
    'utm_source', 'utm_medium', 'utm_campaign', 'utm_term', 'utm_content',
    'fbclid', 'gclid', 'yclid', '_ga'
  ]);
  const kept = query.split('&')
    .filter(p => {
      if (!p) return false;
      const key = decodeURIComponent(p.split('=')[0] || '').toLowerCase();
      return !trackingParams.has(key);
    })
    .sort();

  return kept.length > 0 ? base + '?' + kept.join('&') : base;
}

function buildKeywordRegex_(rawKeywords, defaultPattern) {
  const source = String(rawKeywords || '').trim();
  if (!source) {
    return defaultPattern ? new RegExp(defaultPattern, 'i') : /$^/;
  }
  const parts = source
    .split(/[,|、\n]+/)
    .map(s => s.trim())
    .filter(Boolean)
    .map(s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));

  if (parts.length === 0) {
    return defaultPattern ? new RegExp(defaultPattern, 'i') : /$^/;
  }
  return new RegExp(parts.join('|'), 'i');
}

function extractDeadline_(text) {
  if (!text) return '';
  const jpMatch = text.match(/(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (jpMatch) {
    return jpMatch[1] + '-' + pad2_(jpMatch[2]) + '-' + pad2_(jpMatch[3]);
  }
  const isoMatch = text.match(/(20\d{2})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (isoMatch) {
    return isoMatch[1] + '-' + pad2_(isoMatch[2]) + '-' + pad2_(isoMatch[3]);
  }
  return '';
}

function extractAmount_(text) {
  if (!text) return '';
  const m = text.match(/(?:上限|最大|補助額|助成額|支援額|総額)?\s*(\d[\d,\.]*\s*(?:億円|万円|千円))/);
  return m ? m[0].trim() : '';
}

function computeProgramStatus_(rawStatus, deadlineStr, todayStart) {
  const s = String(rawStatus || 'open').trim().toLowerCase();
  if (s === 'always' || s === 'closed') return s;

  const deadlineDate = parseDeadlineDate_(deadlineStr);
  if (!deadlineDate) {
    return (s === 'closing_soon' || s === 'open') ? s : 'open';
  }

  const diffDays = (deadlineDate.getTime() - todayStart.getTime()) / 86400000;
  if (diffDays < 0) return 'closed';
  if (diffDays <= CONFIG.CLOSING_SOON_DAYS) return 'closing_soon';
  return 'open';
}

function computeIsNew_(status, updatedAtStr, todayStart) {
  if (status !== 'open' && status !== 'closing_soon') return false;
  const parsedDate = parseDeadlineDate_(updatedAtStr);
  if (!parsedDate) return false;
  const diffDays = Math.floor((todayStart.getTime() - parsedDate.getTime()) / 86400000);
  return diffDays >= 0 && diffDays <= 7;
}

function parseDeadlineDate_(deadlineStr) {
  if (!deadlineStr) return null;
  const s = String(deadlineStr).trim();
  const m = s.match(/^(20\d{2})[-\/年](\d{1,2})[-\/月](\d{1,2})日?$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function getTodayStartJst_() {
  const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  const parts = todayStr.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    if (!h) return;
    const val = row[i];
    if (val instanceof Date) {
      obj[h] = Utilities.formatDate(val, 'JST', 'yyyy-MM-dd');
    } else {
      obj[h] = val !== undefined && val !== null ? String(val) : '';
    }
  });
  return obj;
}

function parseTargetsArray_(val) {
  if (Array.isArray(val)) return val.map(s => String(s).trim()).filter(Boolean);
  if (!val) return [];
  return String(val).split(',').map(s => s.trim()).filter(Boolean);
}

function stripHtml_(str) {
  return String(str || '').replace(/<[^>]+>/g, ' ');
}

function decodeHtmlEntities_(str) {
  return String(str || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, (_, num) => String.fromCharCode(Number(num)))
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function cleanWhitespace_(str) {
  return String(str || '').replace(/\s+/g, ' ').trim();
}

function pad2_(n) {
  const s = String(n);
  return s.length < 2 ? '0' + s : s;
}
