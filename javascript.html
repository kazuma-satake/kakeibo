const ss = SpreadsheetApp.getActiveSpreadsheet();

// WebアプリのGETリクエストを処理
function doGet() {
  initialSetup(); // 初回起動時にシートなどを準備
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle('家計簿アプリ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// 初期設定
function initialSetup() {
  // 「設定」シートの確認と作成
  let configSheet = ss.getSheetByName('設定');
  if (!configSheet) {
    configSheet = ss.insertSheet('設定');
    configSheet.getRange('A1').setValue('ジャンル').setFontWeight('bold');
    const sampleGenres = [['食費'], ['日用品'], ['交通費'], ['交際費'], ['趣味'], ['その他']];
    configSheet.getRange(2, 1, sampleGenres.length, 1).setValues(sampleGenres);
  }

  // 「固定費」シートの確認と作成
  let fixedCostSheet = ss.getSheetByName('固定費');
  if (!fixedCostSheet) {
    fixedCostSheet = ss.insertSheet('固定費');
    fixedCostSheet.getRange('A1:B1').setValues([['項目', '金額']]).setFontWeight('bold');
    const sampleFixedCosts = [['家賃', 80000], ['水道光熱費', 10000], ['通信費', 5000]];
    fixedCostSheet.getRange(2, 1, sampleFixedCosts.length, 2).setValues(sampleFixedCosts);
  }

  // 当月シートの確認と作成
  const today = new Date();
  const currentMonthSheetName = Utilities.formatDate(today, 'JST', 'yyyy-MM');
  let currentMonthSheet = ss.getSheetByName(currentMonthSheetName);
  if (!currentMonthSheet) {
    currentMonthSheet = ss.insertSheet(currentMonthSheetName, 0);
    const headers = [['登録日時', '日付', '場所', '品名', '金額', 'ジャンル']];
    currentMonthSheet.getRange('A1:F1').setValues(headers).setFontWeight('bold');
    currentMonthSheet.setFrozenRows(1);
  }
}

// 他のHTMLファイルをインクルードするための関数
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- フロントエンドから呼び出される関数群 ---

/**
 * データをスプレッドシートに保存します。
 * @param {object} data 保存するデータオブジェクト。
 * @return {object} 処理結果。
 */
function saveData(data) {
  try {
    // 入力された日付 (例: "2025-07-15") からシート名 (例: "2025-07") を生成
    const expenseDate = new Date(data.date);
    const sheetName = Utilities.formatDate(expenseDate, 'JST', 'yyyy-MM');
    
    let sheet = ss.getSheetByName(sheetName);
    
    // もし該当月のシートがなければ、新しく作成する
    if (!sheet) {
      sheet = ss.insertSheet(sheetName, 0);
      const headers = [['登録日時', '日付', '場所', '品名', '金額', 'ジャンル']];
      sheet.getRange('A1:F1').setValues(headers).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    const timestamp = new Date();
    sheet.appendRow([timestamp, data.date, data.place, data.item, data.amount, data.genre]);
    
    return { status: 'success', message: '保存しました。' };
  } catch (e) {
    return { status: 'error', message: '保存に失敗しました: ' + e.message };
  }
}

/**
 * 指定された月のデータを読み込みます。
 * @param {string} month 読み込む月 (例: "2025-08")。
 * @return {Array<Array<string>>} 支出データの配列。
 */
function loadData(month) {
  const sheet = ss.getSheetByName(month);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); // ヘッダー行を削除
  return data.map(row => {
    row[0] = row[0] ? Utilities.formatDate(new Date(row[0]), 'JST', 'yyyy/MM/dd HH:mm:ss') : '';
    row[1] = row[1] ? Utilities.formatDate(new Date(row[1]), 'JST', 'yyyy-MM-dd') : '';
    return row;
  });
}

/**
 * ジャンル一覧を取得します。
 * @return {Array<string>} ジャンルの配列。
 */
function getGenres() {
  const sheet = ss.getSheetByName('設定');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const genres = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return genres.flat();
}

/**
 * 新しいジャンルを追加します。
 * @param {string} genre 追加するジャンル名。
 * @return {object} 処理結果。
 */
function addGenre(genre) {
  if (!genre) return { status: 'error', message: 'ジャンル名が空です。'};
  const sheet = ss.getSheetByName('設定');
  const existingGenres = getGenres();
  if (existingGenres.includes(genre)) {
    return { status: 'error', message: '同じ名前のジャンルが既に存在します。'};
  }
  sheet.appendRow([genre]);
  return { status: 'success', message: 'ジャンルを追加しました。', newGenre: genre };
}

/**
 * 固定費一覧を取得します。
 * @return {Array<Array<string|number>>} 固定費の項目と金額の配列。
 */
function getFixedCosts() {
  const sheet = ss.getSheetByName('固定費');
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data;
}

/**
 * 月ごとのシート名一覧を取得します。
 * @return {Array<string>} シート名の配列。
 */
function getSheetNames() {
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  return sheetNames.filter(name => /^\d{4}-\d{2}$/.test(name)).sort().reverse();
}
