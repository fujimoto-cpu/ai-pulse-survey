/*********************************************************
 *  AI活用 月次アンケート  Google Apps Script
 *  - フロントは隠しiframeでGET送信するため doGet/doPost 両対応
 *  - 対象スプシIDをスクリプトに直書き
 *********************************************************/

const SPREADSHEET_ID  = '1TnqowBoVyztIWFa6IbhYA7j4Ohua9IgUjKkTLDcBlEk';
const SHEET_SETTINGS  = '設定';
const SHEET_DASHBOARD = 'ダッシュボード';
const TZ = 'Asia/Tokyo';

// Web App経由でもメニュー経由でも動くようにスプレッドシートを取得
function getSS_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) { /* anonymous web app context */ }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const MONTH_HEADERS = [
  'タイムスタンプ', '氏名',
  'Q1_削減時間(h/月)', 'Q2_1日業務時間変化',
  'Q3_新規創出件数', 'Q4_活用業務', 'Q4_その他自由記述',
  'Q5_AI使用頻度', 'Q6_使いやすさ評価',
  'Q7_困りごと', 'Q7_その他自由記述', 'Q8_困りごと詳細',
  'Q9_インパクト事例', 'Q10_改善要望'
];

const Q2_MAP = {
  '2h_plus' : '2時間以上短縮できた',
  '1-2h'    : '1〜2時間短縮できた',
  '30m-1h'  : '30分〜1時間短縮できた',
  'less_30m': '30分未満の短縮',
  'no_change': 'あまり変わらない'
};
const Q4_MAP = {
  doc: '資料・文書作成', research: '調査・情報収集', code: 'コーディング・開発',
  plan: '企画・アイデア出し', mail: 'メール・文章作成', data: 'データ分析・集計', other: 'その他'
};
const Q5_MAP = {
  daily: '毎日使った', few_week: '週3〜4回', once_week: '週1〜2回',
  few_month: '月に数回', not_used: 'ほぼ使わなかった'
};
const Q7_MAP = {
  no: '特に困っていない', how_to: '使い方・操作がわからない',
  quality: '出力の品質に満足できない', time: '使う時間が確保できない',
  use_case: 'どの業務に使えばよいかわからない', other: 'その他'
};


// ============ ルーティング ============
function doGet(e)  { return handleRequest_(e); }
function doPost(e) { return handleRequest_(e); }

function handleRequest_(e) {
  try {
    let params = {};

    // 1) クエリ ?payload=... または POST bodyの payload を最優先
    if (e && e.parameter && e.parameter.payload) {
      params = JSON.parse(e.parameter.payload);
    } else if (e && e.postData && e.postData.contents) {
      try { params = JSON.parse(e.postData.contents); }
      catch (_) { params = e.parameter || {}; }
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    // 動作確認用：payload無しのGETは alive 応答（バージョンマーカー付き）
    if (!params || !params.name) {
      return jsonOut_({
        ok: true,
        message: 'AI Survey endpoint is alive.',
        version: 'v2-2026-05-21',
        receivedParams: e ? Object.keys(e.parameter || {}) : []
      });
    }

    // q4は配列で来るのでparametersから補正
    if (e && e.parameters && e.parameters.q4 && !Array.isArray(params.q4)) {
      params.q4 = e.parameters.q4;
    }

    const ss = getSS_();
    ensureSettings(ss);

    const now = new Date();
    const monthName = Utilities.formatDate(now, TZ, 'yyyy.MM');
    const sheet = ensureMonthSheet(ss, monthName);

    let q4Arr = params.q4;
    if (!Array.isArray(q4Arr)) q4Arr = q4Arr ? [q4Arr] : [];
    const q4Str = q4Arr.map(v => Q4_MAP[v] || v).join(', ');

    const row = [
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
      params.name || '',
      params.q1 !== '' && params.q1 != null ? Number(params.q1) : '',
      Q2_MAP[params.q2] || '',
      params.q3 !== '' && params.q3 != null ? Number(params.q3) : '',
      q4Str,
      params.q4_other_text || '',
      Q5_MAP[params.q5] || '',
      params.q6 ? Number(params.q6) : '',
      Q7_MAP[params.q7] || '',
      params.q7_other_text || '',
      params.q8 || '',
      params.q9 || '',
      params.q10 || ''
    ];
    sheet.appendRow(row);

    refreshDashboard();

    return jsonOut_({ ok: true, sheet: monthName });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err), stack: err && err.stack });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============ 設定タブ ============
function ensureSettings(ss) {
  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (s) return s;
  s = ss.insertSheet(SHEET_SETTINGS);
  const rows = [
    ['項目', '値', '備考'],
    ['時給単価(円/h)',          5000,  'コスト削減算出に使用'],
    ['AI月額コスト(円)',        50000, 'ROI算出に使用'],
    ['目標_月間削減時間(h)',    100,   ''],
    ['目標_新規創出件数(件)',   30,    ''],
    ['目標_AI浸透率(%)',        70,    '0〜100の数値（％）'],
    ['目標_困ってる人(名)',     2,     '以下が目標'],
    ['対象月',
      Utilities.formatDate(new Date(), TZ, 'yyyy.MM'),
      'ダッシュボードに表示する月（YYYY.MM形式）']
  ];
  s.getRange(1, 1, rows.length, 3).setValues(rows);
  s.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1a0a3e').setFontColor('#ffffff');
  s.setColumnWidth(1, 220);
  s.setColumnWidth(2, 140);
  s.setColumnWidth(3, 360);
  return s;
}

function getSetting(key, fallback) {
  const ss = getSS_();
  const s = ensureSettings(ss);
  const lastRow = s.getLastRow();
  if (lastRow < 2) return fallback;
  const data = s.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const [k, v] of data) {
    if (String(k).trim() === key) return v;
  }
  return fallback;
}


// ============ 月次タブ ============
function ensureMonthSheet(ss, monthName) {
  let s = ss.getSheetByName(monthName);
  if (s) return s;
  s = ss.insertSheet(monthName);
  s.getRange(1, 1, 1, MONTH_HEADERS.length)
    .setValues([MONTH_HEADERS])
    .setFontWeight('bold')
    .setBackground('#7c5cfc')
    .setFontColor('#ffffff');
  s.setFrozenRows(1);
  s.setColumnWidth(1, 150);
  s.setColumnWidth(2, 110);
  s.setColumnWidth(6, 220);
  s.setColumnWidth(12, 260);
  s.setColumnWidth(13, 260);
  s.setColumnWidth(14, 260);
  return s;
}


// ============ ダッシュボード ============
function refreshDashboard() {
  const ss = getSS_();
  let d = ss.getSheetByName(SHEET_DASHBOARD);
  if (!d) {
    d = ss.insertSheet(SHEET_DASHBOARD, 0);
  }
  d.clear();
  d.clearFormats();
  d.getRange(1, 1, 30, 12).breakApart();

  let targetMonth = String(getSetting('対象月', '') || '').trim();
  if (!targetMonth) {
    targetMonth = Utilities.formatDate(new Date(), TZ, 'yyyy.MM');
  }
  const monthSheet = ss.getSheetByName(targetMonth);

  const wage          = Number(getSetting('時給単価(円/h)',       5000))  || 5000;
  const aiCost        = Number(getSetting('AI月額コスト(円)',     50000)) || 50000;
  const targetH       = Number(getSetting('目標_月間削減時間(h)',  100))  || 100;
  const targetCount   = Number(getSetting('目標_新規創出件数(件)', 30))   || 30;
  const targetPenet   = Number(getSetting('目標_AI浸透率(%)',      70))   || 70;
  const targetTrouble = Number(getSetting('目標_困ってる人(名)',   2));

  let totalH = 0, totalCount = 0, dailyUsers = 0, troubled = 0, respondents = 0;
  if (monthSheet && monthSheet.getLastRow() > 1) {
    const data = monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, MONTH_HEADERS.length).getValues();
    respondents = data.length;
    data.forEach(r => {
      totalH     += Number(r[2]) || 0;
      totalCount += Number(r[4]) || 0;
      if (r[7] === Q5_MAP.daily)        dailyUsers++;
      if (r[9] && r[9] !== Q7_MAP.no)   troubled++;
    });
  }
  const monthlyCost   = totalH * wage;
  const yearly        = monthlyCost * 12;
  const roi           = aiCost > 0 ? Math.round(((monthlyCost - aiCost) / aiCost) * 100) : 0;
  const penetRate     = respondents > 0 ? Math.round((dailyUsers / respondents) * 100) : 0;
  const penetVsTarget = targetPenet > 0 ? Math.round((penetRate / targetPenet) * 100) : 0;
  const hourVsTarget  = targetH > 0 ? Math.round((totalH / targetH) * 100) : 0;
  const countVsTarget = targetCount > 0 ? Math.round((totalCount / targetCount) * 100) : 0;
  const troubleStatus = troubled <= targetTrouble ? '✅ OK' : '⚠️ 要注意';

  d.getRange(1, 1, 1, 8).merge()
    .setValue('📊 AI活用ダッシュボード（' + targetMonth + '）')
    .setFontSize(18).setFontWeight('bold');
  d.getRange(2, 1, 1, 8).merge()
    .setValue('回答者数: ' + respondents + '名　/　最終更新: ' +
              Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'))
    .setFontColor('#666666');

  const labelRow = 4, valRow = 5, noteRow = 6;
  d.getRange(labelRow, 1, valRow - labelRow + 2, 8).setBackground('#5b3df5').setFontColor('#ffffff');
  d.getRange(labelRow, 1).setValue('🕒 削減時間（今月）');
  d.getRange(labelRow, 2).setValue('×').setHorizontalAlignment('center');
  d.getRange(labelRow, 3).setValue('💴 時給単価');
  d.getRange(labelRow, 4).setValue('=').setHorizontalAlignment('center');
  d.getRange(labelRow, 5).setValue('💰 月間コスト削減');
  d.getRange(labelRow, 6).setValue('📈 ROI');
  d.getRange(labelRow, 7, 1, 2).merge().setValue('📅 年間換算');
  d.getRange(labelRow, 1, 1, 8).setFontSize(11).setFontWeight('bold');

  d.getRange(valRow, 1).setValue(totalH + ' h');
  d.getRange(valRow, 2).setValue('×').setHorizontalAlignment('center');
  d.getRange(valRow, 3).setValue('¥' + Number(wage).toLocaleString() + ' /h');
  d.getRange(valRow, 4).setValue('=').setHorizontalAlignment('center');
  d.getRange(valRow, 5).setValue('¥' + monthlyCost.toLocaleString());
  d.getRange(valRow, 6).setValue('+' + roi + '%');
  d.getRange(valRow, 7, 1, 2).merge().setValue('¥' + yearly.toLocaleString());
  d.getRange(valRow, 1, 1, 8).setFontSize(18).setFontWeight('bold').setVerticalAlignment('middle');
  d.setRowHeight(valRow, 44);

  d.getRange(noteRow, 6).setValue('vs AI月額 ¥' + aiCost.toLocaleString())
    .setFontSize(10).setFontColor('#cccccc');

  const kpiRow = noteRow + 2;
  const kpis = [
    { title: '🕒 月間削減時間', value: totalH + ' h',
      sub: '目標 ' + targetH + 'h / 先月比 ' + (totalH ? '+' + totalH + 'h' : '-'),
      pct: hourVsTarget + '%', bg: '#e8f9ef', accent: '#1e8e4d' },
    { title: '🌱 新規創出件数', value: totalCount + ' 件',
      sub: '目標 ' + targetCount + '件 / できたタグ合計',
      pct: countVsTarget + '%', bg: '#fdf5e6', accent: '#b9700f' },
    { title: '👥 AI浸透率', value: penetRate + '%',
      sub: '目標 ' + targetPenet + '% / 毎日活用者',
      pct: penetVsTarget + '%', bg: '#eaf3fc', accent: '#1a6cb2' },
    { title: '🆘 困ってる人', value: troubled + ' 名',
      sub: '目標 ' + targetTrouble + '名以下 / 要サポート',
      pct: troubleStatus, bg: '#fdecea', accent: '#b03028' }
  ];

  kpis.forEach((k, i) => {
    const col = 1 + i * 2;
    d.getRange(kpiRow, col).setValue(k.title)
      .setFontWeight('bold').setFontSize(11);
    d.getRange(kpiRow, col + 1).setValue(k.pct)
      .setFontWeight('bold').setFontColor(k.accent)
      .setHorizontalAlignment('right');
    d.getRange(kpiRow + 1, col, 1, 2).merge()
      .setValue(k.value)
      .setFontSize(22).setFontWeight('bold')
      .setVerticalAlignment('middle');
    d.setRowHeight(kpiRow + 1, 50);
    d.getRange(kpiRow + 2, col, 1, 2).merge()
      .setValue(k.sub).setFontSize(10).setFontColor('#666666');
    d.getRange(kpiRow, col, 3, 2).setBackground(k.bg)
      .setBorder(true, true, true, true, false, false, k.accent, SpreadsheetApp.BorderStyle.SOLID);
  });

  for (let c = 1; c <= 8; c++) d.setColumnWidth(c, 135);
  d.setHiddenGridlines(true);
  d.setFrozenRows(0);
}


// ============ メニュー ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI Survey')
    .addItem('初期セットアップ', 'setup')
    .addItem('ダッシュボードを更新', 'manualRefresh')
    .addItem('今月のタブを作成', 'createCurrentMonthSheet')
    .addToUi();
}

function setup() {
  const ss = getSS_();
  ensureSettings(ss);
  createCurrentMonthSheet();
  refreshDashboard();
}

function manualRefresh() {
  refreshDashboard();
}

function createCurrentMonthSheet() {
  const ss = getSS_();
  const monthName = Utilities.formatDate(new Date(), TZ, 'yyyy.MM');
  ensureMonthSheet(ss, monthName);
}
