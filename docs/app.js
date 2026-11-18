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
  FALLBACK_ENDPOINT: './sample-data.json'
};
// ============================================================================

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
  { id: 'closing_soon', label: '締切間近' },
  { id: 'always', label: '通年募集' },
  { id: 'closed', label: '終了' }
];

const PROVIDER_OPTIONS = [
  { id: 'all', label: 'すべて' },
  { id: 'public', label: '国・公的機関' },
  { id: 'municipality', label: '自治体・地域PF' },
  { id: 'private', label: '民間・VC・財団' }
];

const PHASE_LABEL_MAP = Object.fromEntries(PHASES.map(p => [p.id, p.label]));

const state = {
  programs: [],
  lastUpdated: '-',
  selectedPhase: null,
  selectedStatus: 'open',
  selectedProvider: 'all',
  searchQuery: ''
};

const els = {};

document.addEventListener('DOMContentLoaded', () => {
  cacheElements();
  bindEvents();
  readInitialPhaseFromUrl();
  renderFilterChips();
  loadPrograms();
});

function cacheElements() {
  els.lastUpdatedText = document.getElementById('lastUpdatedText');
  els.demoBanner = document.getElementById('demoBanner');
  els.demoBannerText = document.getElementById('demoBannerText');
  els.phaseBar = document.getElementById('phaseBar');
  els.clearPhaseBtn = document.getElementById('clearPhaseBtn');
  els.searchInput = document.getElementById('searchInput');
  els.statusFilters = document.getElementById('statusFilters');
  els.providerFilters = document.getElementById('providerFilters');
  els.resetFiltersBtn = document.getElementById('resetFiltersBtn');
  els.resultsCount = document.getElementById('resultsCount');
  els.loadingState = document.getElementById('loadingState');
  els.errorState = document.getElementById('errorState');
  els.errorMessage = document.getElementById('errorMessage');
  els.retryBtn = document.getElementById('retryBtn');
  els.emptyState = document.getElementById('emptyState');
  els.emptyResetBtn = document.getElementById('emptyResetBtn');
  els.cardsContainer = document.getElementById('cardsContainer');
}

function bindEvents() {
  els.searchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value.trim().toLowerCase();
    render();
  });

  els.clearPhaseBtn.addEventListener('click', () => {
    state.selectedPhase = null;
    render();
  });

  els.resetFiltersBtn.addEventListener('click', resetAllFilters);
  els.emptyResetBtn.addEventListener('click', resetAllFilters);
  els.retryBtn.addEventListener('click', loadPrograms);
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
  state.selectedProvider = 'all';
  state.searchQuery = '';
  els.searchInput.value = '';
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
    els.errorMessage.textContent =
      '通信環境をご確認のうえ、しばらくしてから再度お試しください。';
    setViewState('error');
  }
}

function setViewState(mode) {
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

const STATUS_PRIORITY = {
  closing_soon: 0,
  open: 1,
  always: 2,
  closed: 3
};

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

function matchesNonPhaseFilters(program) {
  if (state.selectedStatus !== 'all') {
    const norm = normalizeStatus(program.status);
    if (state.selectedStatus === 'open') {
      if (norm !== 'closing_soon' && norm !== 'open' && norm !== 'always') return false;
    } else if (norm !== state.selectedStatus) {
      return false;
    }
  }
  if (state.selectedProvider !== 'all') {
    if (program.providerType !== state.selectedProvider) return false;
  }
  if (state.searchQuery) {
    const haystack = [
      program.title,
      program.org,
      program.summary,
      program.amount
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
  });

  renderChipGroup(els.providerFilters, PROVIDER_OPTIONS, state.selectedProvider, (id) => {
    state.selectedProvider = id;
    renderFilterChips();
    render();
  });
}

function renderChipGroup(container, options, activeId, onSelect) {
  container.textContent = '';
  options.forEach((opt) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'chip-btn' + (opt.id === activeId ? ' is-active' : '');
    btn.setAttribute('aria-pressed', String(opt.id === activeId));
    btn.textContent = opt.label;
    btn.addEventListener('click', () => onSelect(opt.id));
    container.appendChild(btn);
  });
}

function renderPhaseBar(baseFiltered) {
  els.phaseBar.textContent = '';

  const countsByPhase = { p1: 0, p2: 0, p3: 0, p4: 0 };
  baseFiltered.forEach((p) => {
    if (Object.prototype.hasOwnProperty.call(countsByPhase, p.phase)) {
      countsByPhase[p.phase] += 1;
    }
  });

  PHASES.forEach((phase) => {
    const isActive = state.selectedPhase === phase.id;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'phase-btn' + (isActive ? ' is-active' : '');
    btn.setAttribute('aria-pressed', String(isActive));

    const top = document.createElement('div');
    top.className = 'phase-top';

    const nameEl = document.createElement('span');
    nameEl.className = 'phase-name';
    nameEl.textContent = phase.label;

    const countEl = document.createElement('span');
    countEl.className = 'phase-count';
    countEl.textContent = `${countsByPhase[phase.id]}件`;

    top.appendChild(nameEl);
    top.appendChild(countEl);

    const descEl = document.createElement('p');
    descEl.className = 'phase-desc';
    descEl.textContent = phase.description;

    btn.appendChild(top);
    btn.appendChild(descEl);

    btn.addEventListener('click', () => {
      state.selectedPhase = isActive ? null : phase.id;
      render();
    });

    els.phaseBar.appendChild(btn);
  });

  els.clearPhaseBtn.hidden = !state.selectedPhase;
}

function renderCards(items) {
  els.cardsContainer.textContent = '';

  items.forEach((item) => {
    const card = document.createElement('article');
    card.className = 'program-card';

    const body = document.createElement('div');

    // Card header: status + phase + org
    const head = document.createElement('div');
    head.className = 'card-head';

    const badges = document.createElement('div');
    badges.className = 'card-badges';

    const statusMeta = getStatusMeta(item.status);
    const statusBadge = document.createElement('span');
    statusBadge.className = `badge ${statusMeta.className}`;
    statusBadge.textContent = statusMeta.label;
    badges.appendChild(statusBadge);

    if (item.phase && PHASE_LABEL_MAP[item.phase]) {
      const phaseBadge = document.createElement('span');
      phaseBadge.className = 'badge badge-phase';
      phaseBadge.textContent = PHASE_LABEL_MAP[item.phase];
      badges.appendChild(phaseBadge);
    }

    const orgEl = document.createElement('span');
    orgEl.className = 'card-org';
    orgEl.textContent = item.org || '';

    head.appendChild(badges);
    head.appendChild(orgEl);
    body.appendChild(head);

    // Title
    const titleEl = document.createElement('h3');
    titleEl.className = 'card-title';
    titleEl.textContent = item.title || '名称未設定';
    body.appendChild(titleEl);

    // Summary
    if (item.summary) {
      const summaryEl = document.createElement('p');
      summaryEl.className = 'card-summary';
      summaryEl.textContent = item.summary;
      body.appendChild(summaryEl);
    }

    card.appendChild(body);

    // Footer: amount, deadline, external link
    const foot = document.createElement('div');
    foot.className = 'card-foot';

    const meta = document.createElement('div');
    meta.className = 'card-meta';

    const amountItem = document.createElement('span');
    amountItem.className = 'card-meta-item';
    const amountStrong = document.createElement('strong');
    amountStrong.textContent = '規模: ';
    amountItem.appendChild(amountStrong);
    amountItem.appendChild(document.createTextNode(item.amount || '-'));
    meta.appendChild(amountItem);

    const deadlineItem = document.createElement('span');
    deadlineItem.className = 'card-meta-item';
    const deadlineStrong = document.createElement('strong');
    deadlineStrong.textContent = '締切: ';
    deadlineItem.appendChild(deadlineStrong);
    deadlineItem.appendChild(document.createTextNode(item.deadline || '-'));
    meta.appendChild(deadlineItem);

    foot.appendChild(meta);

    if (isSafeHttpUrl(item.url)) {
      const link = document.createElement('a');
      link.className = 'card-link';
      link.href = item.url.trim();
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      link.textContent = '公式ページを見る →';
      foot.appendChild(link);
    }

    card.appendChild(foot);
    els.cardsContainer.appendChild(card);
  });
}

function render() {
  const baseFiltered = state.programs.filter(matchesNonPhaseFilters);
  renderPhaseBar(baseFiltered);

  const filtered = state.selectedPhase
    ? baseFiltered.filter(p => p.phase === state.selectedPhase)
    : baseFiltered;

  const finalItems = filtered.slice().sort((a, b) => {
    const pa = STATUS_PRIORITY[normalizeStatus(a.status)] ?? 99;
    const pb = STATUS_PRIORITY[normalizeStatus(b.status)] ?? 99;
    if (pa !== pb) return pa - pb;
    return parseDeadlineTimestamp(a.deadline) - parseDeadlineTimestamp(b.deadline);
  });

  els.resultsCount.textContent = `${finalItems.length}件`;

  const hasAnyFilter =
    Boolean(state.selectedPhase) ||
    state.selectedStatus !== 'open' ||
    state.selectedProvider !== 'all' ||
    Boolean(state.searchQuery);
  els.resetFiltersBtn.hidden = !hasAnyFilter;

  if (finalItems.length === 0) {
    setViewState('empty');
  } else {
    renderCards(finalItems);
    setViewState('ready');
  }
}
