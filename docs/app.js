'use strict';

// ============================================================================
// [EDIT] 設定エリア（ここをご自身の環境に合わせて編集してください）
// ============================================================================
const CONFIG = {
  // [EDIT] Google Apps Script のウェブアプリ公開URLを指定してください
  // 例: 'https://script.google.com/macros/s/AKfycb.../exec?api=programs'
  // 空文字 '' のままにすると、同梱のサンプルデータ (./sample-data.json) を表示します。
  DATA_ENDPOINT: 'https://script.google.com/macros/s/AKfycbyPBTe3iiRLDhA9YjMikHIqJwMCztSLvoMxqUnXgBgVenNROTYJnE3gRLuvbHWDw8NK1w/exec?api=programs',

  // サンプルデータのパス（ローカル確認・未設定時のフォールバック用）
  FALLBACK_ENDPOINT: './sample-data.json',

  // [EDIT] Google Analytics 4 (GA4) を使う場合は測定ID（例: 'G-XXXXXXXXXX'）を設定してください
  // 空文字 '' のままにすると GA4 は読み込まれません。
  GA_MEASUREMENT_ID: '',

  // 簡易アクセス・関心分析（GAS側 Analytics シートへの記録）を有効にするか
  ENABLE_INTERNAL_ANALYTICS: true
};
// ============================================================================

function getTrackEndpoint() {
  const raw = (CONFIG.DATA_ENDPOINT || '').trim();
  if (!raw) return '';
  return raw.replace(/\?api=programs(&|$)/, '?').replace(/\?$/, '');
}

const TRACK_ENDPOINT = getTrackEndpoint();
let searchTrackTimer = null;

function initGA4() {
  const gaId = (CONFIG.GA_MEASUREMENT_ID || '').trim();
  if (!gaId) return;

  window.dataLayer = window.dataLayer || [];
  window.gtag = function gtag() {
    window.dataLayer.push(arguments);
  };
  window.gtag('js', new Date());
  window.gtag('config', gaId, { anonymize_ip: true });

  const script = document.createElement('script');
  script.async = true;
  script.src = 'https://www.googletagmanager.com/gtag/js?id=' + encodeURIComponent(gaId);
  document.head.appendChild(script);
}

function trackEvent(type, label) {
  const cleanLabel = String(label || '').slice(0, 80);

  if (CONFIG.ENABLE_INTERNAL_ANALYTICS && TRACK_ENDPOINT) {
    try {
      const payload = JSON.stringify({ type, label: cleanLabel });
      const sent =
        typeof navigator !== 'undefined' &&
        typeof navigator.sendBeacon === 'function' &&
        navigator.sendBeacon(
          TRACK_ENDPOINT,
          new Blob([payload], { type: 'text/plain;charset=UTF-8' })
        );
      if (!sent) {
        fetch(TRACK_ENDPOINT, {
          method: 'POST',
          mode: 'no-cors',
          keepalive: true,
          headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
          body: payload
        }).catch(() => {});
      }
    } catch {
      // ignore tracking errors
    }
  }

  if (typeof window.gtag === 'function') {
    try {
      window.gtag('event', type, { event_label: cleanLabel });
    } catch {
      // ignore GA4 errors
    }
  }
}

const PHASES = [
  {
    id: 'p1',
    label: '1. アイデア・検証',
    description: '課題仮説の検証、初期プロトタイプや事業アイデアの磨き込み向け'
  },
  {
    id: 'p2',
    label: '2. 起業準備',
    description: '法人設立前後のチーム組成、メンタリング、初期アクセラ向け'
  },
  {
    id: 'p3',
    label: '3. 開発・実証',
    description: 'MVP開発、PoC（概念実証）、試作・技術検証の補助金向け'
  },
  {
    id: 'p4',
    label: '4. 成長・資金調達',
    description: '事業拡大、資本調達、海外展開や大型採択プログラム向け'
  }
];

const STATUS_OPTIONS = [
  { id: 'all', label: 'すべて' },
  { id: 'open', label: '公募中' },
  { id: 'new', label: '新着' },
  { id: 'closing_soon', label: '締切間近' },
  { id: 'always', label: '通年募集' },
  { id: 'closed', label: '終了' }
];

const TARGET_OPTIONS = [
  { id: 'all', label: 'すべて' },
  { id: 'researcher', label: '研究者' },
  { id: 'student', label: '学生' },
  { id: 'other', label: '一般・起業家' }
];

const PROVIDER_OPTIONS = [
  { id: 'all', label: 'すべて' },
  { id: 'public', label: '国・公的機関' },
  { id: 'municipality', label: '自治体・地域PF' },
  { id: 'private', label: '民間・VC・財団' }
];

const PHASE_LABEL_MAP = Object.fromEntries(PHASES.map(p => [p.id, p.label]));
const PHASE_MAP = Object.fromEntries(PHASES.map(p => [p.id, p]));
const PROVIDER_LABEL_MAP = Object.fromEntries(PROVIDER_OPTIONS.map(p => [p.id, p.label]));
const TARGET_LABEL_MAP = {
  researcher: '研究者',
  student: '学生',
  other: '一般・起業家'
};

const state = {
  programs: [],
  lastUpdated: '-',
  selectedPhase: null,
  selectedStatus: 'open',
  selectedTarget: 'all',
  selectedProvider: 'all',
  searchQuery: '',
  activeProgramId: null
};

const els = {};
let lastFocusedElement = null;

document.addEventListener('DOMContentLoaded', () => {
  initGA4();
  cacheElements();
  bindEvents();
  readInitialPhaseFromUrl();
  renderFilterChips();
  loadPrograms();
  trackEvent('pv', 'home');
});

function cacheElements() {
  els.lastUpdatedText = document.getElementById('lastUpdatedText');
  els.demoBanner = document.getElementById('demoBanner');
  els.demoBannerText = document.getElementById('demoBannerText');
  els.searchInput = document.getElementById('searchInput');
  els.clearSearchBtn = document.getElementById('clearSearchBtn');
  els.statusFilters = document.getElementById('statusFilters');
  els.phaseFilters = document.getElementById('phaseFilters');
  els.phaseGuideBar = document.getElementById('phaseGuideBar');
  els.phaseGuideTitle = document.getElementById('phaseGuideTitle');
  els.phaseGuideDesc = document.getElementById('phaseGuideDesc');
  els.targetFilters = document.getElementById('targetFilters');
  els.providerFilters = document.getElementById('providerFilters');
  els.resetFiltersBtn = document.getElementById('resetFiltersBtn');
  els.resultsCount = document.getElementById('resultsCount');
  els.dataTableWrap = document.getElementById('dataTableWrap');
  els.loadingState = document.getElementById('loadingState');
  els.errorState = document.getElementById('errorState');
  els.errorMessage = document.getElementById('errorMessage');
  els.retryBtn = document.getElementById('retryBtn');
  els.emptyState = document.getElementById('emptyState');
  els.emptyResetBtn = document.getElementById('emptyResetBtn');
  els.cardsContainer = document.getElementById('cardsContainer');

  els.detailDrawerOverlay = document.getElementById('detailDrawerOverlay');
  els.detailDrawer = document.getElementById('detailDrawer');
  els.drawerBadges = document.getElementById('drawerBadges');
  els.drawerCloseBtn = document.getElementById('drawerCloseBtn');
  els.drawerFooterCloseBtn = document.getElementById('drawerFooterCloseBtn');
  els.drawerTitle = document.getElementById('drawerTitle');
  els.drawerOrg = document.getElementById('drawerOrg');
  els.drawerAmount = document.getElementById('drawerAmount');
  els.drawerDeadline = document.getElementById('drawerDeadline');
  els.drawerTargets = document.getElementById('drawerTargets');
  els.drawerPhase = document.getElementById('drawerPhase');
  els.drawerProvider = document.getElementById('drawerProvider');
  els.drawerUpdatedAt = document.getElementById('drawerUpdatedAt');
  els.drawerSummaryWrap = document.getElementById('drawerSummaryWrap');
  els.drawerSummary = document.getElementById('drawerSummary');
  els.drawerLink = document.getElementById('drawerLink');
}

function bindEvents() {
  els.searchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value.trim().toLowerCase();
    els.clearSearchBtn.hidden = !state.searchQuery;
    render();
    clearTimeout(searchTrackTimer);
    if (state.searchQuery.length >= 2) {
      const querySnapshot = state.searchQuery;
      searchTrackTimer = setTimeout(() => {
        trackEvent('search', querySnapshot);
      }, 1000);
    }
  });

  els.clearSearchBtn.addEventListener('click', () => {
    state.searchQuery = '';
    els.searchInput.value = '';
    els.clearSearchBtn.hidden = true;
    render();
    els.searchInput.focus();
  });

  els.resetFiltersBtn.addEventListener('click', resetAllFilters);
  els.emptyResetBtn.addEventListener('click', resetAllFilters);
  els.retryBtn.addEventListener('click', loadPrograms);

  els.drawerCloseBtn.addEventListener('click', closeDrawer);
  els.drawerFooterCloseBtn.addEventListener('click', closeDrawer);
  els.detailDrawerOverlay.addEventListener('click', (e) => {
    if (e.target === els.detailDrawerOverlay) {
      closeDrawer();
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !els.detailDrawerOverlay.hidden) {
      closeDrawer();
    }
  });

  els.drawerLink.addEventListener('click', () => {
    const active = state.programs.find(p => String(p.id) === String(state.activeProgramId));
    if (active) {
      trackEvent('program_click', active.title || active.url);
    }
  });
}

function readInitialPhaseFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const paramPhase = params.get('phase');
  const hashPhase = window.location.hash ? window.location.hash.slice(1) : '';
  const candidate = paramPhase || hashPhase;
  if (PHASES.some(p => p.id === candidate)) {
    state.selectedPhase = candidate;
  }
}

function resetAllFilters() {
  state.selectedPhase = null;
  state.selectedStatus = 'open';
  state.selectedTarget = 'all';
  state.selectedProvider = 'all';
  state.searchQuery = '';
  els.searchInput.value = '';
  els.clearSearchBtn.hidden = true;
  renderFilterChips();
  render();
}

async function loadPrograms() {
  setViewState('loading');

  const endpoint = CONFIG.DATA_ENDPOINT ? CONFIG.DATA_ENDPOINT.trim() : '';
  const isUsingFallback = !endpoint;
  const targetUrl = isUsingFallback ? CONFIG.FALLBACK_ENDPOINT : endpoint;

  try {
    const response = await fetch(targetUrl, { method: 'GET' });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const data = await response.json();
    if (!data || !Array.isArray(data.programs)) {
      throw new Error('データ形式が正しくありません');
    }

    state.programs = data.programs;
    state.lastUpdated = data.lastUpdated || '-';
    els.lastUpdatedText.textContent = state.lastUpdated;

    if (isUsingFallback) {
      els.demoBannerText.textContent =
        '現在サンプルデータを表示しています。公開用データを接続するには docs/app.js の DATA_ENDPOINT に GAS Web App URL を設定してください。';
      els.demoBanner.hidden = false;
    } else {
      els.demoBanner.hidden = true;
    }

    render();
  } catch (err) {
    els.resultsCount.textContent = '0件が該当';
    els.errorMessage.textContent =
      '通信環境をご確認のうえ、しばらくしてから再度お試しください。';
    setViewState('error');
  }
}

function setViewState(mode) {
  const showTableWrap = mode === 'loading' || mode === 'ready';
  els.dataTableWrap.hidden = !showTableWrap;
  els.dataTableWrap.style.display = showTableWrap ? '' : 'none';

  const states = [
    [els.loadingState, mode !== 'loading'],
    [els.errorState, mode !== 'error'],
    [els.emptyState, mode !== 'empty'],
    [els.cardsContainer, mode !== 'ready']
  ];
  states.forEach(([el, isHidden]) => {
    el.hidden = isHidden;
    el.style.display = isHidden ? 'none' : '';
  });
}

function normalizeStatus(rawStatus) {
  if (rawStatus === 'closing') return 'closing_soon';
  if (rawStatus === 'upcoming') return 'open';
  if (rawStatus === 'open' || rawStatus === 'closing_soon' || rawStatus === 'always' || rawStatus === 'closed') {
    return rawStatus;
  }
  return 'open';
}

function getStatusMeta(rawStatus) {
  const normalized = normalizeStatus(rawStatus);
  switch (normalized) {
    case 'closing_soon':
      return { label: '締切間近', className: 'badge-closing' };
    case 'always':
      return { label: '通年募集', className: 'badge-always' };
    case 'closed':
      return { label: '終了', className: 'badge-closed' };
    case 'open':
    default:
      return { label: '公募中', className: 'badge-open' };
  }
}

function isSafeHttpUrl(rawUrl) {
  if (typeof rawUrl !== 'string' || !rawUrl.trim()) return false;
  try {
    const parsed = new URL(rawUrl.trim());
    return parsed.protocol === 'https:' || parsed.protocol === 'http:';
  } catch {
    return false;
  }
}

function getSortPriority(program) {
  if (isProgramNew(program)) return 0;
  const norm = normalizeStatus(program.status);
  if (norm === 'open') return 1;
  if (norm === 'closing_soon') return 2;
  if (norm === 'always') return 3;
  if (norm === 'closed') return 4;
  return 99;
}

function parseDeadlineTimestamp(raw) {
  if (typeof raw !== 'string') return Number.POSITIVE_INFINITY;
  const s = raw.trim();
  if (!s) return Number.POSITIVE_INFINITY;

  const isoMatch = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (isoMatch) {
    const [, y, m, d, hh = '23', mm = '59', ss = '59'] = isoMatch;
    return new Date(Number(y), Number(m) - 1, Number(d), Number(hh), Number(mm), Number(ss)).getTime();
  }

  const jpMatch = s.match(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/);
  if (jpMatch) {
    const [, y, m, d] = jpMatch;
    return new Date(Number(y), Number(m) - 1, Number(d), 23, 59, 59).getTime();
  }

  return Number.POSITIVE_INFINITY;
}

function normalizeTargetCode(raw) {
  const val = String(raw || '').trim();
  if (!val) return '';
  const lower = val.toLowerCase();
  if (lower === 'researcher' || val === '研究者') return 'researcher';
  if (lower === 'student' || val === '学生') return 'student';
  if (
    lower === 'other' ||
    lower === 'general' ||
    lower === 'startup' ||
    lower === 'entrepreneur' ||
    val === 'その他' ||
    val === '一般' ||
    val === '起業家' ||
    val === '一般・起業家'
  ) {
    return 'other';
  }
  return val;
}

function getProgramTargetCodes(program) {
  const rawList = Array.isArray(program.targets)
    ? program.targets
    : String(program.targets || '').split(',');
  const codes = [];
  rawList.forEach((raw) => {
    const code = normalizeTargetCode(raw);
    if (code && !codes.includes(code)) {
      codes.push(code);
    }
  });
  return codes.length > 0 ? codes : ['other'];
}

function getProgramTargetLabels(program) {
  return getProgramTargetCodes(program).map(code => TARGET_LABEL_MAP[code] || code);
}

function programMatchesTarget(program, selectedTarget) {
  if (selectedTarget === 'all') return true;
  return getProgramTargetCodes(program).includes(selectedTarget);
}

function parseUpdatedAtTimestamp(raw) {
  if (typeof raw !== 'string') return NaN;
  const m = raw.trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return NaN;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function getTodayStartUtcForJst() {
  const nowJst = new Date(Date.now() + 9 * 3600 * 1000);
  return Date.UTC(nowJst.getUTCFullYear(), nowJst.getUTCMonth(), nowJst.getUTCDate());
}

function isProgramNew(program) {
  const norm = normalizeStatus(program.status);
  if (norm !== 'open' && norm !== 'closing_soon') return false;
  const parsedUtc = parseUpdatedAtTimestamp(program.updatedAt);
  if (Number.isNaN(parsedUtc)) return false;
  const diffDays = Math.floor((getTodayStartUtcForJst() - parsedUtc) / 86400000);
  return diffDays >= 0 && diffDays <= 7;
}

function createNewBadge() {
  const badge = document.createElement('span');
  badge.className = 'badge badge-new';
  const icon = document.createElement('span');
  icon.className = 'material-symbols-rounded badge-icon';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = 'new_releases';
  badge.appendChild(icon);
  badge.appendChild(document.createTextNode('新着'));
  return badge;
}

function matchesNonPhaseFilters(program) {
  if (state.selectedStatus !== 'all') {
    if (state.selectedStatus === 'new') {
      if (!isProgramNew(program)) return false;
    } else {
      const norm = normalizeStatus(program.status);
      if (state.selectedStatus === 'open') {
        if (norm !== 'closing_soon' && norm !== 'open' && norm !== 'always') return false;
      } else if (norm !== state.selectedStatus) {
        return false;
      }
    }
  }

  if (state.selectedTarget !== 'all') {
    if (!programMatchesTarget(program, state.selectedTarget)) return false;
  }

  if (state.selectedProvider !== 'all') {
    if (program.providerType !== state.selectedProvider) return false;
  }

  if (state.searchQuery) {
    const haystack = [
      program.title,
      program.org,
      program.summary,
      program.amount,
      ...getProgramTargetLabels(program)
    ]
      .filter(Boolean)
      .join(' ')
      .toLowerCase();
    if (!haystack.includes(state.searchQuery)) return false;
  }

  return true;
}

function renderFilterChips() {
  renderChipGroup(els.statusFilters, STATUS_OPTIONS, state.selectedStatus, (id) => {
    state.selectedStatus = id;
    renderFilterChips();
    render();
    trackEvent('filter', 'status:' + id);
  });

  renderChipGroup(els.targetFilters, TARGET_OPTIONS, state.selectedTarget, (id) => {
    state.selectedTarget = id;
    renderFilterChips();
    render();
    trackEvent('filter', 'target:' + id);
  });

  renderChipGroup(els.providerFilters, PROVIDER_OPTIONS, state.selectedProvider, (id) => {
    state.selectedProvider = id;
    renderFilterChips();
    render();
    trackEvent('filter', 'provider:' + id);
  });
}

function createChipCheckIcon() {
  const icon = document.createElement('span');
  icon.className = 'material-symbols-rounded chip-check';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = 'check';
  return icon;
}

function renderChipGroup(container, options, activeId, onSelect) {
  container.textContent = '';
  options.forEach((opt) => {
    const isActive = opt.id === activeId;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'chip-btn' + (isActive ? ' is-active' : '');
    btn.setAttribute('aria-pressed', String(isActive));
    if (isActive) {
      btn.appendChild(createChipCheckIcon());
    }
    const labelSpan = document.createElement('span');
    labelSpan.textContent = opt.label;
    btn.appendChild(labelSpan);
    btn.addEventListener('click', () => onSelect(opt.id));
    container.appendChild(btn);
  });
}

function renderPhaseFilters(baseFiltered) {
  els.phaseFilters.textContent = '';

  const countsByPhase = { p1: 0, p2: 0, p3: 0, p4: 0 };
  baseFiltered.forEach((p) => {
    if (Object.prototype.hasOwnProperty.call(countsByPhase, p.phase)) {
      countsByPhase[p.phase] += 1;
    }
  });

  const allBtn = document.createElement('button');
  allBtn.type = 'button';
  const allActive = !state.selectedPhase;
  allBtn.className = 'chip-btn' + (allActive ? ' is-active' : '');
  allBtn.setAttribute('aria-pressed', String(allActive));
  if (allActive) {
    allBtn.appendChild(createChipCheckIcon());
  }
  const allLabel = document.createElement('span');
  allLabel.textContent = 'すべて';
  allBtn.appendChild(allLabel);
  const allCount = document.createElement('span');
  allCount.className = 'chip-count';
  allCount.textContent = String(baseFiltered.length);
  allBtn.appendChild(allCount);
  allBtn.addEventListener('click', () => {
    state.selectedPhase = null;
    render();
  });
  els.phaseFilters.appendChild(allBtn);

  PHASES.forEach((phase) => {
    const isActive = state.selectedPhase === phase.id;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'chip-btn' + (isActive ? ' is-active' : '');
    btn.setAttribute('aria-pressed', String(isActive));
    if (isActive) {
      btn.appendChild(createChipCheckIcon());
    }

    const labelSpan = document.createElement('span');
    labelSpan.textContent = phase.label;
    btn.appendChild(labelSpan);

    const countSpan = document.createElement('span');
    countSpan.className = 'chip-count';
    countSpan.textContent = String(countsByPhase[phase.id]);
    btn.appendChild(countSpan);

    btn.addEventListener('click', () => {
      state.selectedPhase = isActive ? null : phase.id;
      render();
      if (!isActive) {
        trackEvent('phase', phase.label);
      }
    });

    els.phaseFilters.appendChild(btn);
  });

  const activePhase = state.selectedPhase ? PHASE_MAP[state.selectedPhase] : null;
  if (activePhase) {
    els.phaseGuideTitle.textContent = activePhase.label;
    els.phaseGuideDesc.textContent = activePhase.description;
    els.phaseGuideBar.hidden = false;
  } else {
    els.phaseGuideBar.hidden = true;
  }
}

function renderRows(items) {
  els.cardsContainer.textContent = '';

  items.forEach((item) => {
    const normStatus = normalizeStatus(item.status);
    const statusMeta = getStatusMeta(normStatus);
    const isSelected = state.activeProgramId !== null && String(item.id) === String(state.activeProgramId);

    const row = document.createElement('article');
    row.className = 'program-row' + (isSelected ? ' is-selected' : '');
    row.setAttribute('role', 'listitem');
    row.tabIndex = 0;
    row.setAttribute('aria-label', `${item.title || '名称未設定'} の詳細を開く`);

    row.addEventListener('click', () => openDrawer(item));
    row.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openDrawer(item);
      }
    });

    // 1. ステータスバッジ
    const statusCol = document.createElement('div');
    statusCol.className = 'col-status';
    const statusBadge = document.createElement('span');
    statusBadge.className = `badge ${statusMeta.className}`;
    statusBadge.textContent = statusMeta.label;
    statusCol.appendChild(statusBadge);
    if (isProgramNew(item)) {
      statusCol.appendChild(createNewBadge());
    }
    row.appendChild(statusCol);

    // 2. 公募/制度名 + 実施機関/フェーズ
    const mainCol = document.createElement('div');
    mainCol.className = 'col-main';

    const titleEl = document.createElement('h3');
    titleEl.className = 'row-title';
    titleEl.textContent = item.title || '名称未設定';
    mainCol.appendChild(titleEl);

    const subline = document.createElement('div');
    subline.className = 'row-subline';
    const hasOrg = Boolean(item.org);
    const phaseLabel = item.phase && PHASE_LABEL_MAP[item.phase] ? PHASE_LABEL_MAP[item.phase] : '';

    if (hasOrg) {
      const orgSpan = document.createElement('span');
      orgSpan.className = 'row-org';
      orgSpan.textContent = item.org;
      subline.appendChild(orgSpan);
    }
    if (hasOrg && phaseLabel) {
      const sepSpan = document.createElement('span');
      sepSpan.className = 'row-sep';
      sepSpan.textContent = '/';
      subline.appendChild(sepSpan);
    }
    if (phaseLabel) {
      const phaseSpan = document.createElement('span');
      phaseSpan.className = 'row-phase';
      phaseSpan.textContent = phaseLabel;
      subline.appendChild(phaseSpan);
    }
    if (subline.childNodes.length > 0) {
      mainCol.appendChild(subline);
    }
    row.appendChild(mainCol);

    // 3. 支援金額/期間
    const amountCol = document.createElement('div');
    amountCol.className = 'col-amount';
    const amountMobileLabel = document.createElement('span');
    amountMobileLabel.className = 'col-label-mobile';
    amountMobileLabel.textContent = '支援規模';
    amountCol.appendChild(amountMobileLabel);
    amountCol.appendChild(document.createTextNode(item.amount || '-'));
    row.appendChild(amountCol);

    // 4. 応募締切
    const deadlineCol = document.createElement('div');
    deadlineCol.className = 'col-deadline' + (normStatus === 'closing_soon' ? ' is-urgent' : '');
    const deadlineMobileLabel = document.createElement('span');
    deadlineMobileLabel.className = 'col-label-mobile';
    deadlineMobileLabel.textContent = '応募締切';
    deadlineCol.appendChild(deadlineMobileLabel);
    deadlineCol.appendChild(document.createTextNode(item.deadline || '-'));
    row.appendChild(deadlineCol);

    // 5. 対象要件タグ
    const targetsCol = document.createElement('div');
    targetsCol.className = 'col-targets';
    getProgramTargetLabels(item).forEach((label) => {
      const tag = document.createElement('span');
      tag.className = 'target-tag';
      tag.textContent = label;
      targetsCol.appendChild(tag);
    });
    row.appendChild(targetsCol);

    els.cardsContainer.appendChild(row);
  });
}

function openDrawer(program) {
  lastFocusedElement = document.activeElement;
  state.activeProgramId = program.id;

  const normStatus = normalizeStatus(program.status);
  const statusMeta = getStatusMeta(normStatus);

  els.drawerBadges.textContent = '';

  const statusBadge = document.createElement('span');
  statusBadge.className = `badge ${statusMeta.className}`;
  statusBadge.textContent = statusMeta.label;
  els.drawerBadges.appendChild(statusBadge);

  if (isProgramNew(program)) {
    els.drawerBadges.appendChild(createNewBadge());
  }

  const phaseLabel = program.phase && PHASE_LABEL_MAP[program.phase] ? PHASE_LABEL_MAP[program.phase] : '';
  if (phaseLabel) {
    const phaseBadge = document.createElement('span');
    phaseBadge.className = 'badge badge-phase';
    phaseBadge.textContent = phaseLabel;
    els.drawerBadges.appendChild(phaseBadge);
  }

  const providerLabel = PROVIDER_LABEL_MAP[program.providerType] || '';
  if (providerLabel) {
    const providerBadge = document.createElement('span');
    providerBadge.className = 'badge badge-provider';
    providerBadge.textContent = providerLabel;
    els.drawerBadges.appendChild(providerBadge);
  }

  els.drawerTitle.textContent = program.title || '名称未設定';
  els.drawerOrg.textContent = program.org || '実施機関未設定';
  els.drawerAmount.textContent = program.amount || '-';
  els.drawerDeadline.textContent = program.deadline || '-';
  els.drawerTargets.textContent = getProgramTargetLabels(program).join(' / ');
  els.drawerPhase.textContent = phaseLabel || '-';
  els.drawerProvider.textContent = providerLabel || '-';
  els.drawerUpdatedAt.textContent = program.updatedAt || '-';

  if (program.summary) {
    els.drawerSummary.textContent = program.summary;
    els.drawerSummaryWrap.hidden = false;
  } else {
    els.drawerSummary.textContent = '';
    els.drawerSummaryWrap.hidden = true;
  }

  if (isSafeHttpUrl(program.url)) {
    els.drawerLink.href = program.url.trim();
    els.drawerLink.hidden = false;
    els.drawerLink.style.display = '';
  } else {
    els.drawerLink.removeAttribute('href');
    els.drawerLink.hidden = true;
    els.drawerLink.style.display = 'none';
  }

  els.detailDrawerOverlay.hidden = false;
  els.detailDrawerOverlay.style.display = '';
  document.body.style.overflow = 'hidden';

  render();
  els.drawerCloseBtn.focus();
}

function closeDrawer() {
  if (els.detailDrawerOverlay.hidden) return;
  els.detailDrawerOverlay.hidden = true;
  els.detailDrawerOverlay.style.display = 'none';
  document.body.style.overflow = '';
  state.activeProgramId = null;

  render();

  if (lastFocusedElement && typeof lastFocusedElement.focus === 'function' && document.contains(lastFocusedElement)) {
    lastFocusedElement.focus();
  }
  lastFocusedElement = null;
}

function render() {
  const baseFiltered = state.programs.filter(matchesNonPhaseFilters);
  renderPhaseFilters(baseFiltered);

  const filtered = state.selectedPhase
    ? baseFiltered.filter(p => p.phase === state.selectedPhase)
    : baseFiltered;

  const finalItems = filtered.slice().sort((a, b) => {
    const pa = getSortPriority(a);
    const pb = getSortPriority(b);
    if (pa !== pb) return pa - pb;
    return parseDeadlineTimestamp(a.deadline) - parseDeadlineTimestamp(b.deadline);
  });

  els.resultsCount.textContent = `${finalItems.length}件が該当`;

  const hasAnyFilter =
    Boolean(state.selectedPhase) ||
    state.selectedStatus !== 'open' ||
    state.selectedTarget !== 'all' ||
    state.selectedProvider !== 'all' ||
    Boolean(state.searchQuery);
  els.resetFiltersBtn.hidden = !hasAnyFilter;

  if (finalItems.length === 0) {
    setViewState('empty');
  } else {
    renderRows(finalItems);
    setViewState('ready');
  }
}
