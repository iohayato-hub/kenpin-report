/**
 * ==============================================================================
 * 検品報告・生産性管理システム (Code.gs)
 * ==============================================================================
 * このコードをGoogleスプレッドシートの「拡張機能」>「Apps Script」に貼り付けて
 * 「ウェブアプリとしてデプロイ」してください。
 *
 * 【シート構成】
 * 1. 【商品マスタ】: [商品コード, 品名1, 品名2, サイズ, JANコード]
 * 2. 【検品実績】: [報告日, 作業者名, 商品コード, 品名1, 品名2, サイズ, JANコード, 数量, 送信日時]
 * 3. 【作業時間記録】: [報告日, 作業者名, 作業時間(分), 送信日時]
 * ==============================================================================
 */

var SHEET_NAMES = {
  MASTER: '商品マスタ',
  INSPECTION: '検品実績',
  WORK_TIME: '作業時間記録'
};

/**
 * 初回セットアップ: スプレッドシートに必要なシートと初期ヘッダーを自動生成
 */
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 商品マスタ
  var masterSheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(SHEET_NAMES.MASTER);
    var masterHeaders = ['商品コード', '品名1', '品名2', 'サイズ', 'JANコード'];
    masterSheet.appendRow(masterHeaders);
    masterSheet.getRange(1, 1, 1, masterHeaders.length)
      .setBackground('#4F46E5')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    // 初期サンプルデータを登録
    var sampleItems = [
      ['RMB-001', 'ROMBO', 'BLK WHT XS', 'XS', '4580001230011'],
      ['RMB-002', 'ROMBO', 'BLK WHT S', 'S', '4580001230028'],
      ['RMB-003', 'ROMBO', 'NVY RED M', 'M', '4580001230035'],
      ['ITM-101', 'プレミアムコットンTシャツ', 'ホワイト / クルーネック', 'S', '4901234567890'],
      ['ITM-102', 'プレミアムコットンTシャツ', 'ホワイト / クルーネック', 'M', '4901234567891'],
      ['ITM-103', 'プレミアムコットンTシャツ', 'ホワイト / クルーネック', 'L', '4901234567892'],
      ['ITM-104', 'プレミアムコットンTシャツ', 'ブラック / クルーネック', 'M', '4901234567893'],
      ['ITM-105', 'プレミアムコットンTシャツ', 'ブラック / クルーネック', 'L', '4901234567894'],
      ['ITM-201', 'ストレッチテーパードチノ', 'ベージュ / 防シワ加工', '30inch', '4902345678901'],
      ['ITM-202', 'ストレッチテーパードチノ', 'オリーブ / 防シワ加工', '32inch', '4902345678902'],
      ['ITM-301', '防風ライトマウンテンパーカー', 'フォレストグリーン / 撥水', 'M', '4903456789012'],
      ['ITM-302', '防風ライトマウンテンパーカー', 'コヨーテタン / 撥水', 'L', '4903456789029'],
      ['ITM-401', 'オーガニックソックス 3足組', 'モノトーンアソート', '25-27cm', '4904567890123'],
      ['ITM-501', '14ozキャンバストートバッグ', 'ナチュラル×ネイビー', 'Free', '4905678901234'],
      ['ITM-601', 'ウォッシャブルメリノセーター', 'チャコールグレー / ハイゲージ', 'L', '4906789012345'],
      ['ITM-701', 'リバーシブルバケットハット', 'ブラック×カーキ / UVカット', '58cm', '4907890123456']
    ];
    masterSheet.getRange(2, 1, sampleItems.length, sampleItems[0].length).setValues(sampleItems);
    masterSheet.autoResizeColumns(1, masterHeaders.length);
  }

  // 2. 検品実績
  var inspSheet = ss.getSheetByName(SHEET_NAMES.INSPECTION);
  if (!inspSheet) {
    inspSheet = ss.insertSheet(SHEET_NAMES.INSPECTION);
    var inspHeaders = ['報告日', '作業者名', '商品コード', '品名1', '品名2', 'サイズ', 'JANコード', '数量', '送信日時'];
    inspSheet.appendRow(inspHeaders);
    inspSheet.getRange(1, 1, 1, inspHeaders.length)
      .setBackground('#10B981')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    inspSheet.autoResizeColumns(1, inspHeaders.length);
  }

  // 3. 作業時間記録
  var workSheet = ss.getSheetByName(SHEET_NAMES.WORK_TIME);
  if (!workSheet) {
    workSheet = ss.insertSheet(SHEET_NAMES.WORK_TIME);
    var workHeaders = ['報告日', '作業者名', '作業時間(分)', '送信日時'];
    workSheet.appendRow(workHeaders);
    workSheet.getRange(1, 1, 1, workHeaders.length)
      .setBackground('#6366F1')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    workSheet.autoResizeColumns(1, workHeaders.length);
  }

  return { success: true, message: 'シートの初期化が完了しました。' };
}

/**
 * Webアプリケーション GETリクエストハンドラ
 * - パラメータなし: index.html を配信（GAS上で完結動作）
 * - action=getMaster: 商品マスタをJSON返却
 * - action=getProductivity: 生産性データをJSON返却
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  if (action === 'getMaster') {
    var master = apiGetMaster();
    return createJsonResponse(master);
  }

  if (action === 'getProductivity') {
    var date = e.parameter.date || '';
    var prod = apiGetProductivity(date);
    return createJsonResponse(prod);
  }

  if (action === 'init') {
    var initRes = initializeSheets();
    return createJsonResponse(initRes);
  }

  // デフォルト: GAS HTMLサービスでindex.htmlを直接配信
  try {
    var html = HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('検品報告・生産性管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return html;
  } catch (err) {
    return ContentService.createTextOutput('index.html が見つかりません。GASエディタで「HTMLファイル」を作成し名前を index にしてください。\nエラー: ' + err.message);
  }
}

/**
 * Webアプリケーション POSTリクエストハンドラ
 * 外部Webアプリやスマホブラウザからのデータ登録を受信
 */
function doPost(e) {
  try {
    var body = e.postData ? e.postData.contents : '';
    if (!body) {
      return createJsonResponse({ success: false, error: 'リクエストボディが空です' });
    }

    var payload = JSON.parse(body);
    var action = payload.action || 'submitReport';

    if (action === 'submitReport') {
      var result = apiSubmitReport(payload);
      return createJsonResponse(result);
    }

    if (action === 'getMaster') {
      return createJsonResponse(apiGetMaster());
    }

    return createJsonResponse({ success: false, error: '未対応のaction: ' + action });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message, stack: err.stack });
  }
}

/**
 * JSONレスポンスの生成ヘルパー（CORSヘッダー対応）
 */
function createJsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ==============================================================================
// GAS内部API (google.script.run または REST API双方から呼び出し可能)
// ==============================================================================

/**
 * 商品マスタ全件取得
 */
function apiGetMaster() {
  initializeSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { success: true, items: [] };
  }

  // 1行目はヘッダー [商品コード, 品名1, 品名2, サイズ, JANコード]
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var items = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var code = String(row[0] || '').trim();
    if (code) {
      items.push({
        code: code,
        name1: String(row[1] || '').trim(),
        name2: String(row[2] || '').trim(),
        size: String(row[3] || '').trim(),
        jan: String(row[4] || '').trim()
      });
    }
  }

  return { success: true, items: items, count: items.length };
}

/**
 * 検品実績と作業時間の一括書き込み
 * @param {Object} data { records: Array, workTime: Object }
 */
function apiSubmitReport(data) {
  initializeSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  var records = data.records || [];
  var workTime = data.workTime || null;

  var insertedCount = 0;

  // 1. 検品実績シートへの書き込み
  if (records.length > 0) {
    var inspSheet = ss.getSheetByName(SHEET_NAMES.INSPECTION);
    var rowsToInsert = [];

    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      rowsToInsert.push([
        r.date || todayStr,
        r.worker || '未設定',
        r.code || '',
        r.name1 || '',
        r.name2 || '',
        r.size || '',
        "'" + (r.jan || ''), // JANコードの先頭0落ち防止
        Number(r.quantity) || 1,
        r.timestamp || nowStr
      ]);
    }

    var startRow = inspSheet.getLastRow() + 1;
    inspSheet.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    insertedCount = rowsToInsert.length;
  }

  // 2. 作業時間記録シートへの書き込み（作業時間が入力されている場合）
  var workTimeRecorded = false;
  if (workTime && Number(workTime.minutes) > 0) {
    var workSheet = ss.getSheetByName(SHEET_NAMES.WORK_TIME);
    workSheet.appendRow([
      workTime.date || todayStr,
      workTime.worker || '未設定',
      Number(workTime.minutes),
      workTime.timestamp || nowStr
    ]);
    workTimeRecorded = true;
  }

  return {
    success: true,
    message: '送信完了しました',
    recordsCount: insertedCount,
    workTimeRecorded: workTimeRecorded
  };
}

/**
 * 生産性集計（作業者別）
 * - 1点あたりに掛かった検品時間 (秒/点)
 * - 1時間あたりの検品数量 (個/h)
 */
function apiGetProductivity(filterDate) {
  initializeSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inspSheet = ss.getSheetByName(SHEET_NAMES.INSPECTION);
  var workSheet = ss.getSheetByName(SHEET_NAMES.WORK_TIME);

  // 実績データの集計
  var workerQuantities = {};
  var workerDates = {};
  var inspLastRow = inspSheet.getLastRow();

  if (inspLastRow > 1) {
    var inspValues = inspSheet.getRange(2, 1, inspLastRow - 1, 9).getValues();
    for (var i = 0; i < inspValues.length; i++) {
      var row = inspValues[i];
      var date = String(row[0] || '').substring(0, 10);
      var worker = String(row[1] || '未設定').trim();
      var qty = Number(row[7]) || 0;

      if (filterDate && date !== filterDate) continue;

      if (!workerQuantities[worker]) {
        workerQuantities[worker] = 0;
        workerDates[worker] = {};
      }
      workerQuantities[worker] += qty;
      workerDates[worker][date] = true;
    }
  }

  // 作業時間の集計
  var workerMinutes = {};
  var workLastRow = workSheet.getLastRow();

  if (workLastRow > 1) {
    var workValues = workSheet.getRange(2, 1, workLastRow - 1, 4).getValues();
    for (var j = 0; j < workValues.length; j++) {
      var wrow = workValues[j];
      var wdate = String(wrow[0] || '').substring(0, 10);
      var wworker = String(wrow[1] || '未設定').trim();
      var mins = Number(wrow[2]) || 0;

      if (filterDate && wdate !== filterDate) continue;

      if (!workerMinutes[wworker]) {
        workerMinutes[wworker] = 0;
        if (!workerDates[wworker]) workerDates[wworker] = {};
      }
      workerMinutes[wworker] += mins;
      workerDates[wworker][wdate] = true;
    }
  }

  // 作業者一覧をマージ
  var allWorkers = {};
  for (var w1 in workerQuantities) allWorkers[w1] = true;
  for (var w2 in workerMinutes) allWorkers[w2] = true;

  var summaries = [];
  for (var workerName in allWorkers) {
    var totalItems = workerQuantities[workerName] || 0;
    var totalMins = workerMinutes[workerName] || 0;
    var daysCount = Object.keys(workerDates[workerName] || {}).length;

    // 1点あたり所要時間（秒） = (総作業時間(分) * 60) ÷ 総検品数
    var secPerItem = (totalItems > 0 && totalMins > 0)
      ? Math.round(((totalMins * 60) / totalItems) * 10) / 10
      : 0;

    // 1時間あたり検品数量（個/h） = (総検品数 ÷ 総作業時間(分)) * 60
    var itemsPerHour = (totalMins > 0)
      ? Math.round(((totalItems / totalMins) * 60) * 10) / 10
      : 0;

    // ④ 検品時間点数 (目標基準100個/hを100ptとするスコア)
    var productivityScore = (totalMins > 0)
      ? Math.round((itemsPerHour / 100) * 100)
      : 0;

    summaries.push({
      worker: workerName,
      totalItems: totalItems,
      totalMinutes: totalMins,
      secPerItem: secPerItem,
      itemsPerHour: itemsPerHour,
      productivityScore: productivityScore,
      daysWorked: daysCount
    });
  }

  // 作業者ランキング (検品時間点数 降順)
  var workerRankings = summaries.slice().sort(function(a, b) {
    return b.productivityScore - a.productivityScore;
  });
  for (var rIdx = 0; rIdx < workerRankings.length; rIdx++) {
    workerRankings[rIdx].rank = rIdx + 1;
  }

  // ① & ② 週別・月別の集計
  var weeklyMap = {};
  var monthlyMap = {};

  if (inspLastRow > 1) {
    var allInsp = inspSheet.getRange(2, 1, inspLastRow - 1, 9).getValues();
    for (var k = 0; k < allInsp.length; k++) {
      var dStr = String(allInsp[k][0] || '').substring(0, 10);
      var q = Number(allInsp[k][7]) || 0;
      if (!dStr) continue;

      // 月別キー (YYYY-MM)
      var mKey = dStr.substring(0, 7);
      if (!monthlyMap[mKey]) monthlyMap[mKey] = { quantity: 0, minutes: 0 };
      monthlyMap[mKey].quantity += q;

      // 週別キー (月曜日開始日)
      var dt = new Date(dStr + 'T00:00:00');
      var day = dt.getDay();
      var diff = dt.getDate() - day + (day === 0 ? -6 : 1);
      var mon = new Date(dt.setDate(diff));
      var wKey = Utilities.formatDate(mon, 'Asia/Tokyo', 'yyyy-MM-dd');
      if (!weeklyMap[wKey]) weeklyMap[wKey] = { quantity: 0, minutes: 0 };
      weeklyMap[wKey].quantity += q;
    }
  }

  if (workLastRow > 1) {
    var allWork = workSheet.getRange(2, 1, workLastRow - 1, 4).getValues();
    for (var m = 0; m < allWork.length; m++) {
      var wdStr = String(allWork[m][0] || '').substring(0, 10);
      var wMins = Number(allWork[m][2]) || 0;
      if (!wdStr) continue;

      var wmKey = wdStr.substring(0, 7);
      if (!monthlyMap[wmKey]) monthlyMap[wmKey] = { quantity: 0, minutes: 0 };
      monthlyMap[wmKey].minutes += wMins;

      var wdt = new Date(wdStr + 'T00:00:00');
      var wday = wdt.getDay();
      var wdiff = wdt.getDate() - wday + (wday === 0 ? -6 : 1);
      var wmon = new Date(wdt.setDate(wdiff));
      var wwKey = Utilities.formatDate(wmon, 'Asia/Tokyo', 'yyyy-MM-dd');
      if (!weeklyMap[wwKey]) weeklyMap[wwKey] = { quantity: 0, minutes: 0 };
      weeklyMap[wwKey].minutes += wMins;
    }
  }

  return {
    success: true,
    filterDate: filterDate || '全期間',
    summaries: summaries,
    workerRankings: workerRankings,
    weekly: weeklyMap,
    monthly: monthlyMap,
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
  };
}
