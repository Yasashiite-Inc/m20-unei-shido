/**
 * M20 品質管理・コンプラ（運営指導対策）GAS
 * スプレッドシートID: 1b_bkfKRvPBaiOtrgMD72XSfCZZ-MApNFpEmM5lLdI20
 *
 * ■訪問介護員シート（研修管理）と加算要件の対応：
 *   E列「留意点会議日」→ 加算要件「利用者情報・留意事項の伝達」「定期会議開催」
 *   F列「初任時研修1」 → 加算要件「研修の実施」
 *   G列「初任時研修2」 → 加算要件「研修の実施」（複数回）
 *   H列「健康診断日」  → 加算要件「全員への健康診断の実施」
 *
 * ■追加シート（メニュー「運営指導対策 > 初期シート作成」で自動作成）
 *   事業所基本情報 / 運営指導履歴 / 書類管理 / インシデント / 人事登録 / 法務相談 / 健診取込ログ
 *
 * ■任意の設定（プロジェクトの設定 > スクリプト プロパティ）
 *   LEGAL_MAIL   … 法務相談の通知先メールアドレス（未設定なら記録のみ）
 *   HEALTH_FOLDER_ID … 健診画像の保存先フォルダID（未設定なら自動作成）
 *   健診OCRを使う場合は「サービス」から Drive API を追加してください。
 *
 * このコードをコード.gsに丸ごと貼り付けてください。
 */

var SPREADSHEET_ID = '1b_bkfKRvPBaiOtrgMD72XSfCZZ-MApNFpEmM5lLdI20';
var TZ = 'Asia/Tokyo';

// 追加シートの定義（シート名と見出し）
var SHEETS = {
  offices:     {name:'事業所基本情報', headers:['事業所名','サービス種別','事業所番号','開設日','指定更新日','指導権者','エリア','管掌役員','利用者数','特定事業所加算','処遇改善加算','地域区分']},
  inspections: {name:'運営指導履歴',   headers:['事業所名','実施日','区分','指導権者','指摘区分','指摘内容','改善報告期限','改善報告日','状態']},
  documents:   {name:'書類管理',       headers:['事業所名','利用者ID','利用者名','契約書','重要事項説明書','個別援助計画書','計画同意日','月額単位数']},
  incidents:   {name:'インシデント',   headers:['発生日','事業所名','区分','重要度','概要','対応状況']},
  staff:       {name:'人事登録',       headers:['登録日','事業所名','氏名','職種','資格','常勤区分','サ責','週勤務時間','異動区分','届出要否','届出期限','届出状況']},
  legal:       {name:'法務相談',       headers:['受付日時','事業所名','相談者','カテゴリ','相談内容','緊急度','状態']},
  healthLog:   {name:'健診取込ログ',   headers:['取込日時','介護員番号','氏名','受診日','ファイルURL','OCR結果']}
};

// ===== メニュー =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('運営指導対策')
    .addItem('初期シート作成（不足分のみ）', 'setupSheets')
    .addToUi();
}

// ===== ウェブアプリ =====
function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';

  // actionパラメータがなければHTML画面を表示
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('運営指導対策（M20 品質管理・コンプラ）')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // actionパラメータがあればJSON APIとして動作
  var result;
  switch(action) {
    case 'all': result = getAppData(); break;
    case 'kasanStatus': result = getKasanStatus(); break;
    case 'kasanDetail': result = getKasanDetail(e.parameter.office); break;
    case 'helperList': result = getHelperList(e.parameter.office); break;
    default: result = {message:'M20 API', actions:['all','kasanStatus','kasanDetail','helperList']};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var result = {success:true};
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}


// ===== 画面用：全データ一括取得 =====
/**
 * 画面の初期表示で1回だけ呼ぶ。シートがないものは null を返し、画面側でサンプル表示にする。
 */
function getAppData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var out = {generatedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), sources:{}};
  try { out.kasan = getKasanStatus(); out.helpers = getKasanDetail('').介護員一覧; out.sources.kasan = 'sheet'; }
  catch (err) { out.kasan = null; out.helpers = null; out.kasanError = String(err); }
  Object.keys(SHEETS).forEach(function(key) {
    if (key === 'healthLog') return;
    var rows = readSheet_(ss, SHEETS[key].name);
    out[key] = rows;
    out.sources[key] = rows ? 'sheet' : 'sample';
  });
  out.checklists = getChecklists_();
  return out;
}


// ===== 加算要件チェック（メイン機能） =====
/**
 * 特定事業所加算の要件充足をスプレッドシートのデータから判定
 *
 * 返り値の構造:
 * {
 *   事業所名: {
 *     総人数,
 *     要件1_情報伝達: { 要件名, 判定基準, 充足率, 充足人数, 判定, 未実施者[] },
 *     要件2_会議開催: { ... },
 *     要件3_健康診断: { ... },
 *     要件4_研修実施: { ... },
 *     総合判定
 *   }
 * }
 */
function getKasanStatus() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findHelperSheet(ss);
  if (!sheet) return {error:'訪問介護員シートが見つかりません'};

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).filter(function(r){ return r[0] !== '' && r[0] !== null; });

  // 列インデックス特定
  var c = findColumns(headers);

  var today = new Date();
  // 留意点会議: 直近の開催から1ヶ月以内なら「今月実施済み」
  var oneMonthAgo = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate());
  // 健康診断: 1年以内に実施
  var oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());

  // 事業所ごとに集計
  var offices = {};
  rows.forEach(function(row) {
    var office = c.office >= 0 ? String(row[c.office]) : '全体';
    if (!offices[office]) {
      offices[office] = {
        総人数: 0,
        要件1_情報伝達: {要件名:'利用者情報・留意事項の伝達（毎回のサービス提供前）', 判定基準:'留意点会議に出席し情報共有を実施', 充足:0, 未実施:[]},
        要件2_会議開催: {要件名:'サ責による定期的な会議の開催', 判定基準:'月1回以上の会議に参加（留意点会議日が直近1ヶ月以内）', 充足:0, 未実施:[]},
        要件3_健康診断: {要件名:'全ての訪問介護員等に対する健康診断の定期的な実施', 判定基準:'1年以内に健康診断を受診', 充足:0, 未実施:[]},
        要件4_研修実施: {要件名:'訪問介護員の資質向上のための研修機会の確保', 判定基準:'1年以内に研修を受講（研修1または研修2）', 充足:0, 未実施:[]}
      };
    }
    var o = offices[office];
    o.総人数++;
    var name = c.name >= 0 ? String(row[c.name]) : '介護員' + o.総人数;

    // 留意点会議日 → 要件1（情報伝達）& 要件2（会議開催）
    var meetingDate = c.meeting >= 0 ? row[c.meeting] : '';
    if (isValidDate(meetingDate)) {
      var mdD = new Date(meetingDate);
      if (mdD > today || mdD >= oneMonthAgo) { o.要件1_情報伝達.充足++; o.要件2_会議開催.充足++; }
      else { o.要件1_情報伝達.未実施.push(name); o.要件2_会議開催.未実施.push(name); }
    } else { o.要件1_情報伝達.未実施.push(name); o.要件2_会議開催.未実施.push(name); }

    // 健康診断 → 要件3（空白=未取得❌、1年超=期限切れ）
    var healthDate = c.health >= 0 ? row[c.health] : '';
    if (isValidDate(healthDate)) {
      if (new Date(healthDate) >= oneYearAgo) { o.要件3_健康診断.充足++; }
      else { o.要件3_健康診断.未実施.push(name); }
    } else { o.要件3_健康診断.未実施.push(name); }

    // 研修1 or 研修2 → 要件4（未来=予定あり→充足、1年以内=実施済→充足）
    var t1 = c.train1 >= 0 ? row[c.train1] : '';
    var t2 = c.train2 >= 0 ? row[c.train2] : '';
    var t1ok = isValidDate(t1) && (new Date(t1) > today || new Date(t1) >= oneYearAgo);
    var t2ok = isValidDate(t2) && (new Date(t2) > today || new Date(t2) >= oneYearAgo);
    if (t1ok || t2ok) { o.要件4_研修実施.充足++; }
    else { o.要件4_研修実施.未実施.push(name); }
  });

  // 充足率・判定を計算して返す
  var result = {};
  Object.keys(offices).forEach(function(key) {
    var o = offices[key];
    var total = o.総人数;

    var r1 = o.要件1_情報伝達;
    r1.充足率 = Math.round(r1.充足 / total * 100) + '%';
    r1.充足人数 = r1.充足 + '/' + total + '名';
    r1.判定 = r1.未実施.length === 0 ? '✅充足' : '⚠️未実施' + r1.未実施.length + '名';

    var r2 = o.要件2_会議開催;
    r2.充足率 = Math.round(r2.充足 / total * 100) + '%';
    r2.充足人数 = r2.充足 + '/' + total + '名';
    r2.判定 = r2.未実施.length === 0 ? '✅充足' : '⚠️未参加' + r2.未実施.length + '名';

    var r3 = o.要件3_健康診断;
    r3.充足率 = Math.round(r3.充足 / total * 100) + '%';
    r3.充足人数 = r3.充足 + '/' + total + '名';
    r3.判定 = r3.未実施.length === 0 ? '✅全員実施済' : '❌未受診' + r3.未実施.length + '名';

    var r4 = o.要件4_研修実施;
    r4.充足率 = Math.round(r4.充足 / total * 100) + '%';
    r4.充足人数 = r4.充足 + '/' + total + '名';
    r4.判定 = r4.未実施.length === 0 ? '✅全員受講済' : '⚠️未受講' + r4.未実施.length + '名';

    // 総合判定（4要件すべてを見る）
    var allOk = r1.未実施.length === 0 && r2.未実施.length === 0 && r3.未実施.length === 0 && r4.未実施.length === 0;
    o.総合判定 = allOk ? '✅特定事業所加算 全要件充足' : '⚠️一部要件に未充足あり';

    result[key] = o;
  });

  return result;
}


// ===== 事業所別の詳細データ取得 =====
/**
 * 特定事業所の全介護員の個別ステータスを返す
 * ?action=kasanDetail&office=小山訪問介護
 */
function getKasanDetail(officeName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findHelperSheet(ss);
  if (!sheet) return {error:'シートなし'};

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).filter(function(r){ return r[0] !== '' && r[0] !== null; });
  var c = findColumns(headers);

  var today = new Date();
  var oneMonthAgo = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate());
  var oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());

  var result = [];
  rows.forEach(function(row) {
    var office = c.office >= 0 ? String(row[c.office]) : '';
    if (officeName && office !== officeName) return;

    var meetingDate = c.meeting >= 0 ? row[c.meeting] : '';
    var healthDate = c.health >= 0 ? row[c.health] : '';
    var t1 = c.train1 >= 0 ? row[c.train1] : '';
    var t2 = c.train2 >= 0 ? row[c.train2] : '';

    result.push({
      介護員番号: String(row[0]),
      氏名: c.name >= 0 ? String(row[c.name]) : '',
      カナ: String(row[2] || ''),
      事業所: office,
      留意点会議日: meetingDate ? formatDate(meetingDate) : '未設定',
      留意点会議_判定: isValidDate(meetingDate) ? (new Date(meetingDate) > today ? '📅予定' : (new Date(meetingDate) >= oneMonthAgo ? '✅' : '⚠️')) : '⚠️',
      研修1日: t1 ? formatDate(t1) : '未設定',
      研修2日: t2 ? formatDate(t2) : '未設定',
      研修_判定: (isValidDate(t1) && (new Date(t1) > today || new Date(t1) >= oneYearAgo)) || (isValidDate(t2) && (new Date(t2) > today || new Date(t2) >= oneYearAgo)) ? '✅' : '⚠️',
      健康診断日: healthDate ? formatDate(healthDate) : '未取得',
      健康診断_判定: isValidDate(healthDate) ? (new Date(healthDate) >= oneYearAgo ? '✅' : '⚠️期限切') : '❌'
    });
  });

  return {事業所: officeName || '全体', 介護員一覧: result, 人数: result.length};
}

// ===== ヘルパー一覧取得 =====
function getHelperList(officeName) {
  return getKasanDetail(officeName);
}


// ===== 書き込み系（画面から呼ぶ） =====

/** 人事情報の登録。届出要否と期限（変更から10日以内）を自動で付ける */
function saveStaff(p) {
  var today = new Date();
  var needs = (p.サ責 === '○' || p.職種 === '管理者' || p.職種 === 'サービス提供責任者') ? '要（変更届）' : '不要';
  var due = needs === '不要' ? '' : Utilities.formatDate(new Date(today.getTime() + 10 * 86400000), TZ, 'yyyy-MM-dd');
  appendRow_('staff', [Utilities.formatDate(today, TZ, 'yyyy-MM-dd'), p.事業所名, p.氏名, p.職種, p.資格, p.常勤区分, p.サ責, Number(p.週勤務時間) || 0, p.異動区分, needs, due, needs === '不要' ? '-' : '未提出']);
  return {ok:true, 届出要否:needs, 届出期限:due};
}

/** 運営指導の履歴を1件追加 */
function saveInspection(p) {
  appendRow_('inspections', [p.事業所名, p.実施日, p.区分, p.指導権者, p.指摘区分, p.指摘内容, p.改善報告期限, p.改善報告日 || '', p.状態 || '改善報告待ち']);
  return {ok:true};
}

/** 事業所基本情報の追加・更新（事業所名＋サービス種別で一致した行を上書き） */
function saveOffice(p) {
  var sh = getOrCreateSheet_('offices');
  var values = sh.getDataRange().getValues();
  var row = SHEETS.offices.headers.map(function(h){ return p[h] === undefined ? '' : p[h]; });
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === p.事業所名 && String(values[i][1]) === p.サービス種別) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return {ok:true, updated:true};
    }
  }
  sh.appendRow(row);
  return {ok:true, updated:false};
}

/** 弁護士カジュアル相談の受付。LEGAL_MAIL が設定されていれば通知する */
function saveLegal(p) {
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  appendRow_('legal', [now, p.事業所名, p.相談者, p.カテゴリ, p.相談内容, p.緊急度, '受付']);
  var to = PropertiesService.getScriptProperties().getProperty('LEGAL_MAIL');
  if (to) {
    MailApp.sendEmail(to, '【法務相談】' + p.カテゴリ + '（' + p.事業所名 + '）',
      '受付: ' + now + '\n事業所: ' + p.事業所名 + '\n相談者: ' + p.相談者 + '\n緊急度: ' + p.緊急度 + '\n\n' + p.相談内容);
  }
  return {ok:true, mailed:!!to};
}

/** 運営指導 確認項目チェックの保存（事業所ごと） */
function saveChecklist(office, state) {
  PropertiesService.getScriptProperties().setProperty('chk_' + office, JSON.stringify(state));
  return {ok:true};
}

function getChecklists_() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var out = {};
  Object.keys(props).forEach(function(k) {
    if (k.indexOf('chk_') === 0) { try { out[k.substring(4)] = JSON.parse(props[k]); } catch (e) {} }
  });
  return out;
}


// ===== 健康診断：スマホ撮影 → OCR → 自動格納 =====
/**
 * 画像を Drive に保存し、OCR で受診日の候補を返す（まだシートには書かない）
 * base64: 画像データ（data: の後ろ部分）
 */
function uploadHealthImage(base64, mimeType, helperNo, helperName) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, '健診_' + helperNo + '_' + helperName + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmm'));
  var file = getHealthFolder_().createFile(blob);
  var text = '';
  var ocrError = '';
  try { text = ocrImage_(blob); } catch (err) { ocrError = String(err); }
  return {fileUrl: file.getUrl(), text: text.substring(0, 2000), dateCandidates: extractDates_(text), ocrError: ocrError};
}

/** 確認済みの受診日を研修管理シートの健康診断列に書き込む */
function confirmHealthDate(helperNo, helperName, dateStr, fileUrl, ocrText) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findHelperSheet(ss);
  var data = sheet.getDataRange().getValues();
  var c = findColumns(data[0]);
  if (c.health < 0) throw new Error('研修管理シートに「健康診断」列がありません');
  var hit = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(helperNo)) {
      sheet.getRange(i + 1, c.health + 1).setValue(new Date(dateStr));
      hit = true;
      break;
    }
  }
  if (!hit) throw new Error('介護員番号 ' + helperNo + ' が見つかりません');
  appendRow_('healthLog', [Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), helperNo, helperName, dateStr, fileUrl || '', (ocrText || '').substring(0, 500)]);
  return {ok:true};
}

function ocrImage_(blob) {
  if (typeof Drive === 'undefined') throw new Error('Drive API（拡張サービス）が未追加のため OCR は使えません。受診日を手入力してください。');
  var docId;
  if (Drive.Files.create) { // Drive API v3
    docId = Drive.Files.create({name: blob.getName(), mimeType: 'application/vnd.google-apps.document'}, blob, {ocrLanguage: 'ja'}).id;
  } else { // Drive API v2
    docId = Drive.Files.insert({title: blob.getName(), mimeType: blob.getContentType()}, blob, {ocr: true, ocrLanguage: 'ja'}).id;
  }
  var text = DocumentApp.openById(docId).getBody().getText();
  DriveApp.getFileById(docId).setTrashed(true); // OCR用の一時ドキュメントはゴミ箱へ
  return text;
}

/** OCR文字列から日付らしいものを抜き出す（西暦・令和の両方） */
function extractDates_(text) {
  var found = [];
  var re = /(令和\s*(\d{1,2})|(20\d{2}))\s*[年\/\.\-]\s*(\d{1,2})\s*[月\/\.\-]\s*(\d{1,2})/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    var y = m[2] ? 2018 + Number(m[2]) : Number(m[3]);
    var d = new Date(y, Number(m[4]) - 1, Number(m[5]));
    if (!isNaN(d.getTime())) {
      var s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      if (found.indexOf(s) < 0) found.push(s);
    }
  }
  return found;
}

function getHealthFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('HEALTH_FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  var folder = DriveApp.createFolder('健診結果_運営指導対策');
  props.setProperty('HEALTH_FOLDER_ID', folder.getId());
  return folder;
}


// ===== 初期シート作成 =====
/** 追加シートのうち、ないものだけ見出し＋サンプル行つきで作る。既存シートには触らない */
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var names = officeNamesFromHelpers_(ss);
  if (!names.length) names = ['小山訪問介護'];
  var created = [];
  var y = new Date().getFullYear();
  var samples = {
    offices: names.map(function(n, i){ return [n, '訪問介護', '09700000' + (10 + i), (y - 8) + '-04-01', (y + (i % 2 ? 1 : 0)) + '-03-31', '栃木県', '県南エリア', '（管掌役員名）', 50, '特定Ⅰ', '処遇Ⅰ', '7級地']; }),
    inspections: [[names[0], (y - 3) + '-02-14', '運営指導', '栃木県', '計画', '訪問介護計画の同意日がサービス開始後になっている利用者がいた', (y - 3) + '-04-14', (y - 3) + '-04-10', '完了']],
    documents: [[names[0], 'U001', '（利用者名）', '○', '○', '', '', 3200]],
    incidents: [[y + '-09-12', names[0], '転倒', '中', '（概要）', '対応済']],
    staff: [], legal: [], healthLog: []
  };
  Object.keys(SHEETS).forEach(function(key) {
    if (ss.getSheetByName(SHEETS[key].name)) return;
    var sh = ss.insertSheet(SHEETS[key].name);
    sh.appendRow(SHEETS[key].headers);
    sh.getRange(1, 1, 1, SHEETS[key].headers.length).setFontWeight('bold').setBackground('#e8f5e9');
    sh.setFrozenRows(1);
    (samples[key] || []).forEach(function(r){ sh.appendRow(r); });
    created.push(SHEETS[key].name);
  });
  var msg = created.length ? '作成したシート: ' + created.join('、') : '追加するシートはありませんでした';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

function officeNamesFromHelpers_(ss) {
  var sheet = findHelperSheet(ss);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var c = findColumns(data[0]);
  if (c.office < 0) return [];
  var seen = {};
  data.slice(1).forEach(function(r){ var v = String(r[c.office] || ''); if (v) seen[v] = true; });
  return Object.keys(seen);
}


// ===== ユーティリティ =====
function findHelperSheet(ss) {
  var sheets = ss.getSheets();
  // 「研修管理」シートを名前で検索
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name === '研修管理') return sheets[i];
  }
  // 見つからなければ左から9番目（0始まりで8）
  if (sheets.length >= 9) return sheets[8];
  // それでもなければヘッダーに「訪問介護員」「氏名」がある最初のシートを探す
  for (var j = 0; j < sheets.length; j++) {
    var h = sheets[j].getRange(1,1,1,10).getValues()[0];
    for (var k = 0; k < h.length; k++) {
      if (String(h[k]).indexOf('訪問介護員') >= 0 || String(h[k]).indexOf('氏名') >= 0) return sheets[j];
    }
  }
  return sheets[0];
}

function findColumns(headers) {
  // 最初に見つかった列を採用する（「氏名カナ」「事業所番号」などで上書きされないように）
  var c = {name:-1, office:-1, meeting:-1, train1:-1, train2:-1, health:-1};
  headers.forEach(function(h, i) {
    var s = String(h);
    if (c.name < 0 && s.indexOf('氏名') >= 0 && s.indexOf('カナ') < 0) c.name = i;
    if (c.office < 0 && (s.indexOf('事業所') >= 0 || s.indexOf('所属') >= 0) && s.indexOf('番号') < 0) c.office = i;
    if (c.meeting < 0 && (s.indexOf('留意点会議') >= 0 || s.indexOf('会議') >= 0)) c.meeting = i;
    if (c.train1 < 0 && s.indexOf('研修1') >= 0) c.train1 = i;
    if (c.train2 < 0 && s.indexOf('研修2') >= 0) c.train2 = i;
    if (c.health < 0 && s.indexOf('健康診断') >= 0) c.health = i;
  });
  return c;
}

function isValidDate(val) {
  if (!val) return false;
  var d = new Date(val);
  return !isNaN(d.getTime());
}

function formatDate(val) {
  if (!val) return '';
  var d = new Date(val);
  if (isNaN(d.getTime())) return String(val);
  return d.getFullYear() + '/' + (d.getMonth()+1) + '/' + d.getDate();
}

/** シートを見出しつきオブジェクト配列で返す。日付は yyyy-MM-dd 文字列に（google.script.run は Date を返せないため） */
function readSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return null;
  var values = sh.getDataRange().getValues();
  if (values.length < 1) return [];
  var headers = values[0].map(String);
  return values.slice(1).filter(function(r){ return r.join('') !== ''; }).map(function(r) {
    var o = {};
    headers.forEach(function(h, i) {
      var v = r[i];
      o[h] = (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : v;
    });
    return o;
  });
}

function getOrCreateSheet_(key) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEETS[key].name);
  if (!sh) { sh = ss.insertSheet(SHEETS[key].name); sh.appendRow(SHEETS[key].headers); sh.setFrozenRows(1); }
  return sh;
}

function appendRow_(key, row) {
  getOrCreateSheet_(key).appendRow(row);
}
