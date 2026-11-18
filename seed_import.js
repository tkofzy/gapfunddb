/**
 * 初期シードデータ一括インポート用スクリプト（GAS エディタから直接実行可能）
 *
 * 使い方:
 * 1. `clasp push` 後、GAS エディタで `importAttachedSeedData` を選択して「実行」
 * 2. URL ベースで重複排除するため、何度実行しても冪等（重複登録されない）です
 */

/**
 * 82件の初期公募プログラムと12件の巡回ターゲットをスプレッドシートへ冪等にインポートする。
 * @return {Object} { ok, programsAdded, programsSkipped, targetsAdded, targetsSkipped }
 */
function importAttachedSeedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = seedGetOrInitSheet_(ss, SHEET_PROGRAMS, PROGRAM_HEADERS);
  const tSheet = seedGetOrInitSheet_(ss, SHEET_TARGETS, TARGET_HEADERS);
  ensureTargetHeaders_(tSheet);

  const existingProgramUrls = new Set();
  if (pSheet.getLastRow() > 1) {
    const pRows = pSheet.getDataRange().getValues();
    for (let i = 1; i < pRows.length; i++) {
      const norm = normalizeUrl_(String(pRows[i][9] || ''));
      if (norm) existingProgramUrls.add(norm);
    }
  }

  const existingTargetUrls = new Set();
  if (tSheet.getLastRow() > 1) {
    const tRows = tSheet.getDataRange().getValues();
    for (let i = 1; i < tRows.length; i++) {
      const norm = normalizeUrl_(String(tRows[i][1] || ''));
      if (norm) existingTargetUrls.add(norm);
    }
  }

  const todayStart = getTodayStartJst_();
  const nowStr = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  const seedPrograms = getSeedProgramsData_();
  const newProgramRows = [];
  let programsSkipped = 0;

  seedPrograms.forEach((raw, idx) => {
    const normUrl = normalizeUrl_(raw.url);
    if (!normUrl || existingProgramUrls.has(normUrl)) {
      programsSkipped++;
      return;
    }

    const deadline = normalizeSeedDeadline_(raw.deadline);
    const status = computeProgramStatus_(raw.status, deadline, todayStart);
    const id = 'seed_' + String(idx + 1).padStart(3, '0');

    newProgramRows.push([
      id,
      raw.title,
      raw.org,
      raw.phase,
      normalizeSeedProvider_(raw.providerType),
      normalizeSeedTargets_(raw.targets),
      raw.amount,
      deadline,
      status,
      raw.url,
      raw.summary,
      '',
      'manual',
      nowStr
    ]);
    existingProgramUrls.add(normUrl);
  });

  if (newProgramRows.length > 0) {
    pSheet.getRange(pSheet.getLastRow() + 1, 1, newProgramRows.length, PROGRAM_HEADERS.length)
      .setValues(newProgramRows);
  }

  const seedTargets = getSeedCrawlTargetsData_();
  const newTargetRows = [];
  let targetsSkipped = 0;

  seedTargets.forEach(t => {
    const normUrl = normalizeUrl_(t.url);
    if (!normUrl || existingTargetUrls.has(normUrl)) {
      targetsSkipped++;
      return;
    }
    newTargetRows.push([
      t.org,
      t.url,
      t.phase,
      t.providerType,
      t.keywords,
      'TRUE',
      t.type,
      t.excludeKeywords || '',
      '',
      '',
      ''
    ]);
    existingTargetUrls.add(normUrl);
  });

  if (newTargetRows.length > 0) {
    tSheet.getRange(tSheet.getLastRow() + 1, 1, newTargetRows.length, TARGET_HEADERS.length)
      .setValues(newTargetRows);
  }

  invalidatePublicCache_();

  const result = {
    ok: true,
    programsAdded: newProgramRows.length,
    programsSkipped: programsSkipped,
    targetsAdded: newTargetRows.length,
    targetsSkipped: targetsSkipped
  };
  console.log('importAttachedSeedData result:', JSON.stringify(result));
  return result;
}

/**
 * 管理画面・RPC からシードインポートを実行するための管理者保護ラッパー。
 */
function runSeedImport(token) {
  assertAdmin_(token);
  return importAttachedSeedData();
}

function seedGetOrInitSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function normalizeSeedProvider_(val) {
  const s = String(val || '').trim();
  if (s === '国・公的資金' || s === 'public') return 'public';
  if (s.includes('municipality') || s.includes('自治体')) return 'municipality';
  if (s === 'private') return 'private';
  return 'public';
}

function normalizeSeedTargets_(val) {
  const map = { '研究者': 'researcher', '学生': 'student', 'その他': 'other' };
  return String(val || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .map(s => map[s] || s)
    .join(',');
}

function normalizeSeedDeadline_(val) {
  const s = String(val || '').trim();
  const m = s.match(/^(20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2}:\d{2}$/);
  return m ? m[1] : s;
}

function getSeedProgramsData_() {
  return [
    {
      title: 'JST 大学発新産業創出基金事業「ディープテック・スタートアップ国際展開プログラム（D-Global）」',
      org: '科学技術振興機構（JST）',
      providerType: '国・公的資金',
      targets: '研究者,その他',
      phase: 'p3',
      deadline: '2026-09-30 12:00:00',
      amount: '最大約3億円/件',
      status: 'closed',
      url: 'https://www.jst.go.jp/program/startupkikin/',
      summary: '事業化推進機関との共同申請 / 2026年9月30日受付終了、審査・選考進行中'
    },
    {
      title: 'JST START（大学発新産業創出プログラム）「プロジェクト推進型 SBIRフェーズ1支援」',
      org: '科学技術振興機構（JST）',
      providerType: '国・公的資金',
      targets: '研究者',
      phase: 'p1',
      deadline: '2026-04-30 12:00:00',
      amount: '数千万円規模',
      status: 'closed',
      url: 'https://www.jst.go.jp/start/sbir/call2026.html',
      summary: '2026年度公募は4月30日正午で受付終了'
    },
    {
      title: 'NEDO ディープテック・スタートアップ支援基金「国際共同研究開発」',
      org: '新エネルギー・産業技術総合開発機構（NEDO）',
      providerType: '国・公的資金',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2027-01-20 12:00:00',
      amount: '上限1億円、補助率2/3以内',
      status: 'open',
      url: 'https://www.nedo.go.jp/koubo/CA2_100541.html',
      summary: '2026年10月7日公募開始、2027年1月20日正午締切。上限1億円／件（補助率2/3以内）。'
    },
    {
      title: 'NEDO NEP（NEDO Entrepreneurs Program）開拓コース / 躍進コース',
      org: '新エネルギー・産業技術総合開発機構（NEDO）',
      providerType: '国・公的資金',
      targets: '研究者,学生,その他',
      phase: 'p1',
      deadline: '春期公募終了（次回予告注視）',
      amount: '開拓:最大500万円 / 躍進:最大3000万円',
      status: 'closed',
      url: 'https://www.nedo.go.jp/koubo/',
      summary: '春・秋の定期公募サイクル（次回準備）'
    },
    {
      title: 'GTIE GAPファンド「エクスプロールコース」',
      org: 'Greater Tokyo Innovation Ecosystem（GTIE）',
      providerType: '自治体・大学連携(municipality)',
      targets: '研究者',
      phase: 'p1',
      deadline: '2026年11月下旬募集開始予定',
      amount: '上限6,000万円（2年/3年コース）',
      status: 'open',
      url: 'https://gtie.jp/news/45183/',
      summary: '次回（2027年1期）は2026年11月下旬募集開始予定'
    },
    {
      title: 'GTIE GAPファンド「海外市場開拓実践コース」',
      org: 'Greater Tokyo Innovation Ecosystem（GTIE）',
      providerType: 'municipality',
      targets: '研究者',
      phase: 'p3',
      deadline: '2026年12月募集開始予定',
      amount: '上限2,800万円（最長1年間）',
      status: 'open',
      url: 'https://gtie.jp/news/44958/',
      summary: 'GTIE主幹・共同機関の研究者対象。グローバル展開を目指す技術シーズ向け。'
    },
    {
      title: '東大IPC「1stRound」起業支援プログラム',
      org: '東大IPC / 参画大学共創',
      providerType: 'municipality',
      targets: '研究者,学生,その他',
      phase: 'p2',
      deadline: '2026-11-16 23:59:00',
      amount: '最大1,000万円（非希薄化資金）＋企業連携',
      status: 'closed',
      url: 'https://www.1stround.jp/',
      summary: '24大学6研究機関共催。最大1,000万円非希薄化資金＋ハンズオン支援（第16回受付終了・審査中）'
    },
    {
      title: '東京都 大学発スタートアップ創出支援事業',
      org: '東京都',
      providerType: 'municipality',
      targets: '研究者,学生,その他',
      phase: 'p2',
      deadline: '年度プログラム',
      amount: 'コーディネーター伴走支援・実証支援',
      status: 'open',
      url: 'https://www.startupandglobalfinancialcity.metro.tokyo.lg.jp/startup/initiatives/university-startup-support',
      summary: '都内大学発スタートアップチーム向け伴走・実証支援'
    },
    {
      title: 'BRAVE（技術シーズ事業化推進プログラム）',
      org: 'Beyond Next Ventures株式会社',
      providerType: 'private',
      targets: '研究者,その他',
      phase: 'p1',
      deadline: '公募要領参照',
      amount: '経営人材マッチング・最大数千万円出資機会・伴走支援',
      status: 'always',
      url: 'https://beyondnextventures.com/jp/incubation/brave/',
      summary: '国内最大級の研究成果事業化アクセラレーター（定期実施）'
    },
    {
      title: 'LINK-J ディープテック集中アクセラレーター',
      org: '一般社団法人LINK-J',
      providerType: 'private',
      targets: '研究者,学生,その他',
      phase: 'p2',
      deadline: '2026年度分終了/次回予告',
      amount: '専門家伴走・メンタリング・VCマッチング',
      status: 'closed',
      url: 'https://www.link-j.org/news/article-47340.html',
      summary: 'ライフサイエンス・医療機器・ディープテック向け（次回準備）'
    },
    {
      title: 'FASTAR（スタートアップ成長支援事業）',
      org: '中小企業基盤整備機構',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p2',
      deadline: '2026-12-04 12:00:00',
      amount: '専門家による1年間の伴走支援・資金調達支援',
      status: 'open',
      url: 'https://fastar.smrj.go.jp/program/16/',
      summary: '第16期公募中（12月4日正午締切）。プレシード〜アーリー対象、地域枠新設。'
    },
    {
      title: 'KAWASAKI DEEP TECH ACCELERATOR',
      org: '川崎市 / 民間連携',
      providerType: 'municipality',
      targets: '研究者,その他',
      phase: 'p3',
      deadline: '2026/06/10',
      amount: 'インキュベーション・PoCフィールド提供・専門家伴走',
      status: 'closed',
      url: 'https://www.city.kawasaki.jp/templates/prs/280/0000186806.html',
      summary: '研究開発型・ものづくり・ディープテック特化'
    },
    {
      title: 'TMIP Tokyo GreenTech Challenge',
      org: '一般社団法人大丸有環境軸（TMIP） / 東京都',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '随時/年度公募',
      amount: '大手企業300団体との実証実験・協定金支援・PoC費用',
      status: 'always',
      url: 'https://www.tmip.jp/report/12982',
      summary: '東京都TOKYO SUTEAM重点分野採択プログラム（GX・脱炭素・サーキュラーエコノミー）'
    },
    {
      title: 'DBJスタートアップアクセラレーションアワード2026',
      org: '日本政策投資銀行（DBJ）',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026/07/31',
      amount: '専門家伴走支援・マッチング・DBJ投融資連携',
      status: 'closed',
      url: 'https://ssc.jeri.or.jp/award/',
      summary: '政府系金融機関によるスタートアップ発掘'
    },
    {
      title: 'TRIBUS 2026（リコー アクセラレータープログラム）',
      org: '株式会社リコー',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '公募要領参照',
      amount: '実証実験費用支援・リコーアセット提供・事業共創',
      status: 'open',
      url: 'https://jp.ricoh.com/release/2026/0608_2',
      summary: '大手事業会社による統合型共創プログラム'
    },
    {
      title: 'KEIKYU ACCELERATOR PROGRAM',
      org: '京浜急行電鉄株式会社 / 横須賀市',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-08-24 23:59:00',
      amount: '京急沿線アセットでの実証実験・事業共創',
      status: 'closed',
      url: 'https://www.keikyu.co.jp/report/2026/report2026060901.html',
      summary: 'モビリティ・スマートシティ・地域課題解決（沿線アセット活用型）'
    },
    {
      title: '令和8(2026)年度 社会課題解決スタートアップ企業応援事業',
      org: '栃木県',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '公募要領参照',
      amount: '補助上限500万円（補助率10/10以内）',
      status: 'open',
      url: 'https://www.pref.tochigi.lg.jp/f03/startup/r8_socialsu_support.html',
      summary: '10月13日より2次募集開始（社会課題解決全般）'
    },
    {
      title: '2026年度 第2回 スタートアップ助成',
      org: 'アーツカウンシル東京',
      providerType: 'municipality',
      targets: 'その他,学生',
      phase: 'p2',
      deadline: '2026-09-24 17:00:00',
      amount: '上限100万円/件',
      status: 'closed',
      url: 'https://www.artscouncil-tokyo.jp/uploads/2026/08/2026_2_startup-grant_release_260825.pdf',
      summary: 'アート・クリエイティブ・コンテンツ・文化技術分野'
    },
    {
      title: 'ディープテック・スタートアップ支援基金／大企業等の連携によるディープテック・スタートアップ支援事業',
      org: 'NEDO',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-09-30',
      amount: '大企業アセット連携・実証研究開発助成',
      status: 'closed',
      url: 'https://www.nedo.go.jp/koubo/CA2_100535.html',
      summary: '大企業等の経営資源を活用した実証・スケールアップ支援'
    },
    {
      title: '青山スタートアップアクセラレーションセンター（ASAC）アクセラレーションプログラム',
      org: '東京都',
      providerType: 'municipality',
      targets: 'その他,研究者,学生',
      phase: 'p2',
      deadline: '2026-09-13 23:59:00',
      amount: '5ヶ月間の集中伴走支援・メンタリング・PoC支援',
      status: 'closed',
      url: 'https://www.acceleration-tokyo.metro.tokyo.lg.jp/',
      summary: 'プレシード〜シード対象の集中伴走プログラム'
    },
    {
      title: 'ソーシャルXアクセラレーションプログラム',
      org: '株式会社ソーシャル・エックス / 官民共創パートナー',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-09-30 00:00:00',
      amount: '自治体実証連携・官民共創伴走・マッチング支援',
      status: 'closed',
      url: 'https://sac.socialx.inc/',
      summary: '自治体課題解決と事業成長を両立する官民共創特化アクセラレーター'
    },
    {
      title: 'キャンパスベンチャーグランプリ（CVG）',
      org: '日刊工業新聞社',
      providerType: 'private',
      targets: '学生,研究者',
      phase: 'p1',
      deadline: '2026/09/30',
      amount: '大賞賞金100万円・大臣賞授与・事業化支援',
      status: 'closed',
      url: 'https://cvg.nikkan.co.jp/',
      summary: '国内最大級の学生ビジネスプランコンテスト'
    },
    {
      title: 'TOKYO STARTUP GATEWAY（TSG）',
      org: '東京都 / NPO法人ETIC.',
      providerType: 'municipality',
      targets: '学生,その他,研究者',
      phase: 'p1',
      deadline: '2026/07/14',
      amount: '賞金総額800万円超（最優秀賞300万円）＋アクセラレーション支援',
      status: 'closed',
      url: 'https://tsg.metro.tokyo.lg.jp/',
      summary: '400字からエントリーできる東京都主催のブラッシュアップ型コンテスト'
    },
    {
      title: '日本弁理士会 知財活用ビジネスプランコンテスト',
      org: '日本弁理士会',
      providerType: 'private',
      targets: '研究者,学生,その他',
      phase: 'p2',
      deadline: '2026-09-30 23:59:00',
      amount: '表彰・賞金＋弁理士による知的財産・事業化メンタリング支援',
      status: 'closed',
      url: 'https://www.jpaa.or.jp/about-us/attached_institution/management/bpc/',
      summary: '知的財産権を活用したビジネスプランを評価・支援'
    },
    {
      title: 'Shibuya Startup Support「S-Startups」スタートアップ認定制度',
      org: '渋谷区',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-09-30',
      amount: '渋谷区実証実験フィールド提供・官民連携・海外展開支援',
      status: 'closed',
      url: 'https://www.city.shibuya.tokyo.lg.jp/kusei/hodo/hodo-2026/hodo_20260901.html?id=vsr3h8d75',
      summary: '渋谷を拠点にグローバル成長を目指す認定スタートアップ支援'
    },
    {
      title: '八王子市ビジネスアイデアピッチコンテスト「ハチステ」',
      org: '八王子市 / 株式会社ツクリエ',
      providerType: 'municipality',
      targets: '学生,研究者,その他',
      phase: 'p1',
      deadline: '2026-10-16 23:59:00',
      amount: '賞金・事業化ブラッシュアップ伴走支援・VCマッチング',
      status: 'closed',
      url: 'https://prtimes.jp/main/html/rd/p/000000543.000033548.html',
      summary: '多摩・八王子地域のアイデア・起業家発掘ピッチコンテスト'
    },
    {
      title: '日本ソーシャル・スタートアップAWARD',
      org: '日本フィランソロピック財団',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-09-30',
      amount: '助成金（総額数千万円規模・1件数百万円）＋伴走支援',
      status: 'closed',
      url: 'https://np-foundation.or.jp/information/000411.html',
      summary: '社会課題解決に取り組む創業初期・シードスタートアップ向け助成型アワード'
    },
    {
      title: 'FRONTIER PITCH TOKYO for Startups',
      org: 'FRONTIER PITCH 実行委員会',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-09-30 23:59:00',
      amount: '大手企業・VC審査員への直接ピッチ・出資およびPoC共創機会',
      status: 'closed',
      url: 'https://prtimes.jp/main/html/rd/p/000000390.000069250.html',
      summary: '最先端技術・フロンティア領域のピッチコンテスト'
    },
    {
      title: '三菱UFJ技術育成財団 研究開発助成金',
      org: '三菱UFJ技術育成財団（MU-TECH）',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-20',
      amount: '上限300万円（返済不要助成金・研究開発費の2/3以内）',
      status: 'closed',
      url: 'https://www.mutech.or.jp/subsidy/',
      summary: '設立5年以内等の技術志向型スタートアップ向け研究開発助成'
    },
    {
      title: 'グローバル・スタートアップ・アクセラレーションプログラム（GSAP）',
      org: 'JETRO / 内閣府・経済産業省',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-09-10 23:59:00',
      amount: '世界トップアクセラレーターによる集中ハンズオン・海外展開支援',
      status: 'closed',
      url: 'https://www.jetro.go.jp/services/gsap.html',
      summary: '海外進出を目指すスタートアップ向け集中プログラム'
    },
    {
      title: 'IP BASE AWARD',
      org: '特許庁',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-31 23:59:00',
      amount: '特許庁長官賞授与・知財専門家ネットワーク連携',
      status: 'closed',
      url: 'https://ipbase.go.jp/award/',
      summary: '知財を活用して成長を目指すスタートアップ・支援者表彰制度'
    },
    {
      title: 'サーキュラーエコノミー実証スタートアップコンテスト（CSUP）',
      org: '埼玉県',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-16 23:59:00',
      amount: '埼玉県内での社会実証支援・マッチング・知事賞授与',
      status: 'closed',
      url: 'https://cehub.jp/news/saitama-csup-3rd-startup-business-plan-contest-entries/',
      summary: '資源循環・サーキュラーエコノミー領域の実証コンテスト'
    },
    {
      title: 'OIHスタートアップアクセラレーションプログラム（OSAP）',
      org: '大阪市 / 大阪産業局',
      providerType: 'municipality',
      targets: 'その他',
      phase: 'p4',
      deadline: '2026-10-09 23:59:00',
      amount: '専門家集団による4か月間の伴走支援',
      status: 'closed',
      url: 'https://www.obda.or.jp/press-release/20260908.html',
      summary: 'アーリー期スタートアップ向け集中支援プログラム'
    },
    {
      title: 'STATION Ai オープンイノベーション支援プログラム「SKIP Focus」',
      org: 'STATION Ai株式会社',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-05 23:59:00',
      amount: '実証支援・大手事業会社との協業案件創出・伴走育成',
      status: 'closed',
      url: 'https://prtimes.jp/main/html/rd/p/000000138.000095825.html',
      summary: 'インキュベーション拠点STATION Aiによるオープンイノベーション創出プログラム'
    },
    {
      title: 'KIDスタートアップピッチコンテスト',
      org: '神奈川県',
      providerType: 'municipality',
      targets: 'その他,研究者,学生',
      phase: 'p2',
      deadline: '2026-09-28 17:00:00',
      amount: '賞金授与・県内実証連携・VCマッチング支援',
      status: 'closed',
      url: 'https://startups.pref.kanagawa.jp/event/246698/',
      summary: '神奈川発のスタートアップ発掘・育成ピッチコンテスト'
    },
    {
      title: 'JST大学発新産業創出基金事業 KSAC-GAPファンド',
      org: '関西スタートアップアカデミア・コアリション（KSAC） / JST',
      providerType: 'municipality',
      targets: '研究者,学生',
      phase: 'p1',
      deadline: '2026-10-30 12:00:00',
      amount: 'ステップ1:上限500万〜1,000万円 / ステップ2:上限6,000万円',
      status: 'closed',
      url: 'https://rd.doshisha.ac.jp/rd/news/detail/018-ckIqRF.html',
      summary: '関西圏大学連合によるGAPファンド第4回公募'
    },
    {
      title: '静岡インバウンドツーリズムイノベーション',
      org: '静岡県 / Creww株式会社',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-04 23:59:00',
      amount: '静岡県内での実証実験サポート・事業者アライアンス支援',
      status: 'closed',
      url: 'https://growth.creww.me/program_list',
      summary: '観光・インバウンド領域特化のオープンイノベーションプログラム'
    },
    {
      title: '京信・地域の起業家アワード',
      org: '京都信用金庫',
      providerType: 'private',
      targets: 'その他,研究者,学生',
      phase: 'p2',
      deadline: '2026-09-30 23:59:00',
      amount: '表彰・賞金授与・事業性投融資連携・メンタリング',
      status: 'closed',
      url: 'https://www.kyoto-shinkin.co.jp/business/support/start/award/',
      summary: '地域課題解決・イノベーション創出を目指す起業家表彰'
    },
    {
      title: '池田泉州銀行 イノベーション研究開発助成金／ニュービジネス助成金',
      org: '株式会社池田泉州銀行',
      providerType: 'private',
      targets: '研究者,その他',
      phase: 'p3',
      deadline: '2026-10-20',
      amount: '大賞300万円・優秀賞100万円（返還不要）',
      status: 'closed',
      url: 'https://www.sihd-bk.jp/corporation/collaboration/subsidy/common/pdf/newbusiness_points.pdf',
      summary: '産学連携・技術開発・ニュービジネスを支援する助成金'
    },
    {
      title: '脱炭素化・エネルギー転換に資する我が国技術の国際実証事業',
      org: 'NEDO',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-10-26 12:00:00',
      amount: '海外での技術実証研究開発費用の助成',
      status: 'closed',
      url: 'https://www.nedo.go.jp/koubo/GI2_100001_00028.html',
      summary: '脱炭素・クリーンテック技術の海外市場展開・現地実証事業'
    },
    {
      title: 'GTIE GAPファンド エントリーコース',
      org: 'GTIE',
      providerType: 'municipality',
      targets: '研究者,学生',
      phase: 'p1',
      deadline: '2026-10-07 12:00:00',
      amount: '上限500万円／件（最長1年以内）',
      status: 'closed',
      url: 'https://gtie.jp/news/45531/',
      summary: '首都圏大学連合GTIEによる創業前GAPファンド第6回'
    },
    {
      title: '大企業設備活用スタートアップ協業支援事業',
      org: '東京都 産業労働局',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-12-03',
      amount: '大企業5社の研究設備利用料相当額支援＋伴走支援',
      status: 'open',
      url: 'https://www.metro.tokyo.lg.jp/information/press/2026/09/2026091706',
      summary: '2026年9月17日〜12月3日募集。大企業研究設備を活用した協業支援。'
    },
    {
      title: 'UPGRADE with TOKYO（第61回）',
      org: '東京都 産業労働局 / 狛江市',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-15 18:00:00',
      amount: '都・自治体との協働事業化・実証実験検討',
      status: 'closed',
      url: 'https://upgrade-tokyo.metro.tokyo.lg.jp/',
      summary: '都政課題解決型ピッチイベント'
    },
    {
      title: 'NX-Tech HUB Open Innovation Program 2026',
      org: 'NIPPON EXPRESSホールディングス / eiicon',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-26 23:59:00',
      amount: '実証フィールド提供・PoC費用支援・CVC出資検討',
      status: 'closed',
      url: 'https://eiicon.net/about/nipponexpress-holdings2026/',
      summary: '物流課題×AI・ロボティクス共創プログラム'
    },
    {
      title: '防衛省版SBIR制度',
      org: '防衛省 / 防衛装備庁',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-11-02',
      amount: '各フェーズ研究開発委託費・交付金',
      status: 'closed',
      url: 'https://www.mod.go.jp/atla/sbir/koubo/r08/r8_sbir_koubo_youkou_2610.pdf',
      summary: 'デュアルユース先端技術の研究開発支援制度'
    },
    {
      title: 'TOKYO SUTEAM 令和8年度協定事業者プログラム',
      org: '東京都',
      providerType: 'municipality',
      targets: 'その他',
      phase: 'p2',
      deadline: '各事業者による',
      amount: '各協定事業者による伴走・支援プログラム',
      status: 'open',
      url: 'https://www.metro.tokyo.lg.jp/information/press/2026/09/2026092809',
      summary: '令和8年度協定事業者40者の採択発表。各事業者による新規公募を展開。'
    },
    {
      title: '大分発！ソーシャル・スタートアップ基金',
      org: '日本フィランソロピック財団',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-10-31 23:59:00',
      amount: '1件あたり上限300万円（助成金）',
      status: 'closed',
      url: 'https://np-foundation.or.jp/information/000427.html',
      summary: '地域社会課題解決に取り組むスタートアップ助成金'
    },
    {
      title: 'TOPPAN新事業共創プログラム「co-necto 2026」',
      org: 'TOPPAN株式会社',
      providerType: 'private',
      targets: 'その他',
      phase: 'p3',
      deadline: '2026-10-31 23:59:00',
      amount: 'パートナー企業25社との実証フィールド・アセット提供',
      status: 'closed',
      url: 'https://toppan-conecto.com/',
      summary: 'パートナー企業とのオープンイノベーション・実証実験'
    },
    {
      title: '大学生向け起業プログラム『STARTLINE』3期生',
      org: 'ANRI株式会社',
      providerType: 'private',
      targets: '学生',
      phase: 'p1',
      deadline: '2026-10-31 00:00:00',
      amount: '設立費用100万円(Non-Equity)',
      status: 'closed',
      url: 'https://anri.vc/',
      summary: '投資を前提としない学生向け起業支援プログラム'
    },
    {
      title: '第12回日本アントレプレナー大賞',
      org: '賢者の選択リーダーズ倶楽部',
      providerType: 'private',
      targets: 'その他',
      phase: 'p2',
      deadline: '2027-01-18 23:59:00',
      amount: '大賞100万円、部門賞10万円',
      status: 'open',
      url: 'https://kenja-club.com/lc/entrepreneur',
      summary: '高校生以上・事業化7年未満を対象としたビジネスコンテスト（2027年1月18日締切）'
    },
    {
      title: 'A-IDEAセレクト（第2回公募）',
      org: '愛知県 経済産業局 / eiicon',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-11-09 17:00:00',
      amount: '愛知県認定プロジェクト指定、伴走・実証・マッチング支援',
      status: 'closed',
      url: 'https://corp.eiicon.net/press/B0Pllx2',
      summary: '社会課題起点型プロジェクト認定制度'
    },
    {
      title: 'TOKYO to Global: Femtech & Healthcare Startup US Launchpad',
      org: 'NEXTBLUE / 東京都',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-12-10 23:59:00',
      amount: '20社採択（無料）・最大5社PoC補助・10社シカゴ渡航宿泊補助',
      status: 'open',
      url: 'https://nextblue.com/programs/tokyotoglobal',
      summary: 'フェムテック・ヘルスケア領域の米国展開支援（12月10日締切）'
    },
    {
      title: 'J-StarX Global Bootcamp Series in Kansai',
      org: 'JETRO',
      providerType: 'public',
      targets: '研究者,学生,その他',
      phase: 'p1',
      deadline: '2026-10-04 23:59:00',
      amount: '国内ブートキャンプ＋海外現地滞在（渡航滞在費JETRO負担）',
      status: 'closed',
      url: 'https://www.jetro.go.jp/services/j-starx/global_bootcamp_kansai.html',
      summary: '起業家育成・グローバル進出ブートキャンプ'
    },
    {
      title: 'HSFC GAPファンド（ステップ2）',
      org: 'ノーステック財団 / HSFC',
      providerType: 'municipality',
      targets: '研究者',
      phase: 'p3',
      deadline: '2026-10-29 12:00:00',
      amount: '研究開発費（上限数百万円〜最大数千万円規模）',
      status: 'closed',
      url: 'https://hsfc.jp/events/202605081262/',
      summary: '北海道スタートアップエコシステムによるGAPファンド'
    },
    {
      title: 'K-NIC Startup Hands on Program',
      org: '川崎市、NEDO、K-NIC',
      providerType: 'public',
      targets: '研究者,その他,学生',
      phase: 'p1',
      deadline: '2026-11-11 23:59:00',
      amount: '専門メンター集中伴走支援・NEDO公募申請対策',
      status: 'closed',
      url: 'https://www.k-nic.jp/program/12097/',
      summary: 'ディープテック領域の起業前〜創業初期向け集中伴走プログラム'
    },
    {
      title: '東京大学事業化推進助成制度（第20期）',
      org: '東京大学 産学協創推進本部',
      providerType: 'municipality',
      targets: '研究者',
      phase: 'p1',
      deadline: '公募要領参照',
      amount: '研究開発費助成（ステップ1/2）、事業化検証支援',
      status: 'open',
      url: 'https://www.ducr.u-tokyo.ac.jp/information/updates/index.html',
      summary: '所属研究者向けの事業化推進GAPファンド（10月28日公募開始）'
    },
    {
      title: 'CINOVA YAMANASHI MATCHING',
      org: '山梨県 産業労働部 / eiicon',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-11-17 23:59:00',
      amount: '実証補助金 上限100万円（補助率1/2以内）＋伴走支援',
      status: 'closed',
      url: 'https://yamanashi-matching.eiicon.net/',
      summary: '山梨県内ホスト企業とのオープンイノベーション'
    },
    {
      title: 'ローカルスタートアップ共創プログラム『スクラム・コモンズ』',
      org: '株式会社秋田銀行',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p1',
      deadline: '2027-01-25',
      amount: '秋田銀行が共同創業者として2年間立ち上げ・資金伴走',
      status: 'open',
      url: 'https://prtimes.jp/main/html/rd/p/000000012.000085847.html',
      summary: '銀行が共同創業者として参画する創業前インキュベーション（2027年1月25日締切）'
    },
    {
      title: 'MONO INNOVATION AICHI',
      org: '愛知県 経済産業局 / eiicon',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '公募要領参照',
      amount: '愛知県内モノづくり企業との共創マッチング・実証支援',
      status: 'open',
      url: 'http://aichi.eiicon.net/mono-innovation-aichi2026',
      summary: '製造・ロボティクス・DX/GX等の共創パートナー募集'
    },
    {
      title: 'TOKYO SEED ENGINE（ANRI × TOKYO SUTEAM）',
      org: 'ANRI / 東京都',
      providerType: 'municipality',
      targets: '研究者,学生',
      phase: 'p1',
      deadline: '通年／公募要領参照',
      amount: '事業化検証、GAPファンド申請支援、経営人材マッチング',
      status: 'always',
      url: 'https://prtimes.jp/main/html/rd/p/000000092.000040191.html',
      summary: '研究室発シーズの起業を仕組みで支援するインキュベーション'
    },
    {
      title: 'TOKYO STARTUP LINK（課題提案型マッチング）',
      org: '東京都 産業労働局',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-11-18 23:59:00',
      amount: '都内企業とのマッチング・ピッチ・実証フォローアップ',
      status: 'closing_soon',
      url: 'https://www.metro.tokyo.lg.jp/information/press/2026/10/2026100717',
      summary: 'ホスト企業課題募集後にスタートアップ登壇募集を開始'
    },
    {
      title: 'JR東日本スタートアッププログラム2026（秋募集）',
      org: 'JR東日本 / JR東日本スタートアップ',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-11-13 23:59:00',
      amount: '実証フィールド提供・PoC費用支援・CVC出資検討',
      status: 'closed',
      url: 'https://jrestartup.co.jp/2026-entry-autumn/',
      summary: '地域・デジタル・地球共創のオープンイノベーション'
    },
    {
      title: 'AI-Native Startup Ignition',
      org: 'スペースデータ / 東京都',
      providerType: 'municipality',
      targets: '研究者,その他,学生',
      phase: 'p1',
      deadline: '通年／公募要領参照',
      amount: 'ハッカソン、経営メンタリング、法人設立支援、出資検討',
      status: 'always',
      url: 'https://prtimes.jp/main/html/rd/p/000000219.000080352.html',
      summary: '少人数×AI型起業をハッカソンから法人化まで伴走'
    },
    {
      title: 'ピッチコンテスト「Changing Tomorrow」柏の葉-つくば',
      org: 'LINK-J / アステラス製薬 / 三井不動産',
      providerType: 'private',
      targets: '研究者',
      phase: 'p2',
      deadline: '2026-11-30 17:00:00',
      amount: '優勝賞金1,000万円＋シェアラボ1年間無償利用',
      status: 'closing_soon',
      url: 'https://www.link-j.org/news/article-53347.html',
      summary: '創薬・ライフサイエンス領域のピッチコンテスト（11月30日締切）'
    },
    {
      title: 'GX分野のディープテック・スタートアップ事業開発支援（GX_UPP）',
      org: 'NEDO',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-11-17 12:00:00',
      amount: '補助上限50億円（補助率1/2以内）',
      status: 'closed',
      url: 'https://www.nedo.go.jp/koubo/CA2_100511.html',
      summary: 'GX・資源循環・省エネモビリティの実用化・事業開発支援'
    },
    {
      title: '大企業等のスタートアップ連携・調達加速化事業（PoP / GX_PoP）',
      org: 'NEDO',
      providerType: 'public',
      targets: 'その他',
      phase: 'p4',
      deadline: '2027年1月上旬（予定）',
      amount: '大企業による本格調達に向けた製品検証補助',
      status: 'open',
      url: 'https://www.nedo.go.jp/koubo/CA1_100547.html',
      summary: '11月下旬〜2027年1月上旬公募予定'
    },
    {
      title: 'Global Deeptech Startup Foundry',
      org: 'CIC Japan / 東京都',
      providerType: 'municipality',
      targets: '研究者,その他',
      phase: 'p2',
      deadline: '2026-11-30 17:00:00',
      amount: '研究開発費補助最大150万円＋海外渡航費補助最大50万円',
      status: 'closing_soon',
      url: 'https://prtimes.jp/main/html/rd/p/000000187.000063842.html',
      summary: '先端材料・AI・ロボ・量子等のディープテック支援（11月30日締切）'
    },
    {
      title: 'Business Next Tokyo（ビジネックス東京）',
      org: '東京都 産業労働局',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2027/03/31',
      amount: '大手CVC10社との協業マッチング＋共同PoC補助金上限2,000万円',
      status: 'open',
      url: 'https://business-next.metro.tokyo.lg.jp/',
      summary: '随時マッチング受付、12月14日リバースピッチ開催'
    },
    {
      title: 'AICHI NEXT UNICORN LEAGUE（シーズン9）',
      org: '愛知県 経済産業局 / eiicon',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p2',
      deadline: '2026-11-16 23:59:00',
      amount: '賞金総額1,000万円（大賞500万円）＋メンタリング',
      status: 'closed',
      url: 'https://aichi.eiicon.net/aichi-n-unicorn2026',
      summary: 'サステナビリティ×イノベーションのビジネスプランコンテスト'
    },
    {
      title: 'TCIC IP INNOVATION AWARDS（TIIA）2026',
      org: '東京都 / TCIC',
      providerType: 'municipality',
      targets: 'その他,学生,研究者',
      phase: 'p2',
      deadline: '2026-10-29 17:00:00',
      amount: '最優秀賞300万円＋1年間伴走支援・オフィス利用権',
      status: 'closed',
      url: 'https://tcic.metro.tokyo.lg.jp/tiia2026/',
      summary: 'コンテンツIP・XR・AI等のビジネスプランコンテスト'
    },
    {
      title: '令和8年度 U29起業家ピッチコンテスト in YOXO BOX',
      org: '神奈川県 / 横浜市',
      providerType: 'municipality',
      targets: '学生,その他,研究者',
      phase: 'p1',
      deadline: '2026-12-20 23:59:00',
      amount: 'ピッチ登壇・事業化支援',
      status: 'open',
      url: 'https://www.pref.kanagawa.jp/docs/sr4/kidyoxo/u29pitch2027.html',
      summary: '29歳以下の若手・学生対象（12月20日締切）'
    },
    {
      title: 'スタートアップ育成支援プログラム「J-Startup」（第6次選定）',
      org: '経済産業省 / JETRO・NEDO',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-11-13 17:00:00',
      amount: '海外展開支援・NEDO優遇・政府入札機会拡大',
      status: 'closed',
      url: 'https://www.pwc.com/jp/ja/news-room/2026/j-startup2026.html',
      summary: '政府認定・集中支援プログラム（初の公募制導入）'
    },
    {
      title: 'JST「早暁（そうぎょう）プログラム」第4期 ステージ1',
      org: '科学技術振興機構（JST）',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p1',
      deadline: '2027年3月上旬（予定）',
      amount: 'シーズ探索経費支援・事業化検証・GAPファンド接続',
      status: 'open',
      url: 'https://www.jst.go.jp/program/startupkikin/sogyo-sp/index.html',
      summary: '12月上旬公募開始予定（技術シーズ探索・チーム組成）'
    },
    {
      title: 'スタートアップ社会実装促進事業（PoC Ground Tokyo）第3期',
      org: '東京都 産業労働局',
      providerType: 'municipality',
      targets: 'その他,研究者',
      phase: 'p3',
      deadline: '2026-11-24 12:00:00',
      amount: 'PoC実証実験費用支援（協定金支給）・フィールド提供',
      status: 'closing_soon',
      url: 'https://poc-ground.metro.tokyo.lg.jp/',
      summary: '暑さ対策・DX・GX・女性活躍の4重点領域（11月24日正午締切）'
    },
    {
      title: '世界学生起業家アワード「EO GSEA東京大会」',
      org: 'EO Tokyo Central',
      providerType: 'private',
      targets: '学生',
      phase: 'p2',
      deadline: '2027-01-05 23:59:00',
      amount: '日本大会・世界大会進出権、現役起業家メンタリング',
      status: 'open',
      url: 'https://gsea-japan.org/venue-info/',
      summary: '学生起業家向けピッチアワード（2027年1月5日締切）'
    },
    {
      title: 'SHIMADZUみらい共創チャレンジ',
      org: '株式会社島津製作所',
      providerType: 'private',
      targets: '研究者,その他',
      phase: 'p3',
      deadline: '2026-12-31 23:59:00',
      amount: '年間研究費400万〜500万円＋先端設備利用',
      status: 'open',
      url: 'https://www.shimadzu.co.jp/research_and_development/future_collaboratory/challenge/',
      summary: '産学共同研究・社外パートナー公募（12月31日締切）'
    },
    {
      title: 'アカデミック・アントレプレナーチャレンジ ピッチコンテスト',
      org: '東京農工大学 / NEDO等',
      providerType: 'municipality',
      targets: '学生,研究者',
      phase: 'p1',
      deadline: '2026-12-01 12:00:00',
      amount: '学長賞副賞50万円、企業賞・NEDO賞＋メンタリング',
      status: 'closing_soon',
      url: 'https://www.tuat-flourish.jp/topics/event/3839/',
      summary: '学術知見・研究シーズに基づくピッチ（12月1日正午締切）'
    },
    {
      title: 'シード投資プログラム「Atlas 2027 Winter」',
      org: 'Coreline Ventures株式会社',
      providerType: 'private',
      targets: 'その他,研究者',
      phase: 'p2',
      deadline: '2026-11-30 23:59:00',
      amount: '開始時シード出資＋5ヶ月間集中ハンズオン支援',
      status: 'closing_soon',
      url: 'https://atlas-by-coreline.com/',
      summary: 'プレシード〜シード期向け投資・育成プログラム（11月30日締切）'
    },
    {
      title: '5G GROWTH FIELD ビジネスコンテスト',
      org: '株式会社ユニッジ / 東京都',
      providerType: 'municipality',
      targets: 'その他',
      phase: 'p3',
      deadline: '2027-01-18',
      amount: '最大500万円の実証費用＋社会実装伴走支援',
      status: 'open',
      url: 'https://tokyo-next5g-unidge.co.jp/',
      summary: '次世代通信（5G・6G）×身体拡張（1月18日締切）'
    },
    {
      title: '環境スタートアップ大賞',
      org: '環境省',
      providerType: 'public',
      targets: 'その他,研究者',
      phase: 'p4',
      deadline: '2026-12-11 17:00:00',
      amount: '大臣賞表彰・ピッチ登壇・投資家マッチング',
      status: 'open',
      url: 'https://www.env.go.jp/press/press_05825.html',
      summary: '環境保全・気候変動・資源循環領域（12月11日締切）'
    },
    {
      title: 'Osaka Fintech & Beyond Accelerator 2026',
      org: '大阪府・大阪市 / Plug and Play Japan',
      providerType: 'municipality',
      targets: 'その他',
      phase: 'p4',
      deadline: '2026/12月中旬開始前まで',
      amount: '関西金融機関・大企業との個別マッチング・PoC検討',
      status: 'open',
      url: 'https://contents-jp.plugandplaytechcenter.com/osaka_fintech_accelerator_26',
      summary: 'Fintech・金融イノベーション向けアクセラレーター'
    },
    {
      title: '東北地域スタートアップ成長加速支援事業',
      org: '東北経済産業局',
      providerType: 'public',
      targets: 'その他',
      phase: 'p3',
      deadline: '2026-11-27 17:00:00',
      amount: '参加費無料・事業会社との個別マッチング支援',
      status: 'closing_soon',
      url: 'https://j-net21.smrj.go.jp/news/bt5puv0000011ttm.html',
      summary: '東北6県拠点・進出意向のあるスタートアップ向け（11月27日締切）'
    }
  ];
}

function getSeedCrawlTargetsData_() {
  return [
    { org: 'JST スタートアップ・新産業', url: 'https://www.jst.go.jp/program/startupkikin/', phase: 'p1', providerType: 'public', keywords: '公募,募集,GAP,スタートアップ,SBIR', type: 'html' },
    { org: 'NEDO 公募情報', url: 'https://www.nedo.go.jp/koubo/', phase: 'p3', providerType: 'public', keywords: 'スタートアップ,ディープテック,SBIR,NEP,公募', type: 'html' },
    { org: 'GTIE ニュース', url: 'https://gtie.jp/news/', phase: 'p1', providerType: 'municipality', keywords: 'GAP,公募,募集,スタートアップ', type: 'html' },
    { org: '東京都 創業・スタートアップ施策', url: 'https://www.startupandglobalfinancialcity.metro.tokyo.lg.jp/startup/initiatives/', phase: 'p2', providerType: 'municipality', keywords: '公募,募集,スタートアップ,支援,アクセラ', type: 'html' },
    { org: 'J-Net21 新着ニュース', url: 'https://j-net21.smrj.go.jp/news/', phase: 'p2', providerType: 'public', keywords: 'スタートアップ,起業,補助金,公募,アクセラ', type: 'auto' },
    { org: 'LINK-J ニュース', url: 'https://www.link-j.org/news/', phase: 'p2', providerType: 'private', keywords: 'アクセラ,ピッチ,公募,募集,スタートアップ', type: 'html' },
    { org: '神奈川県 スタートアップイベント・公募', url: 'https://startups.pref.kanagawa.jp/', phase: 'p2', providerType: 'municipality', keywords: '公募,ピッチ,募集,スタートアップ,アクセラ', type: 'html' },
    { org: '日本フィランソロピック財団', url: 'https://np-foundation.or.jp/information/', phase: 'p3', providerType: 'private', keywords: '助成,公募,スタートアップ,アワード', type: 'html' },
    { org: 'Creww プログラム一覧', url: 'https://growth.creww.me/program_list', phase: 'p3', providerType: 'private', keywords: 'アクセラ,共創,募集,イノベーション', type: 'html' },
    { org: '東大IPC 1stRound', url: 'https://www.1stround.jp/', phase: 'p2', providerType: 'municipality', keywords: '1stRound,公募,募集,エントリー', type: 'html' },
    { org: 'Beyond Next Ventures', url: 'https://beyondnextventures.com/jp/incubation/brave/', phase: 'p1', providerType: 'private', keywords: 'BRAVE,募集,エントリー,アクセラ', type: 'html' },
    { org: 'K-NIC プログラム', url: 'https://www.k-nic.jp/program/', phase: 'p1', providerType: 'public', keywords: '公募,募集,プログラム,アクセラ,ハンズオン', type: 'html' }
  ];
}
