/**
 * Config.gs
 * 設定値・ヘルパー関数
 */

// ===========================================
// シート名定義（定数）
// ===========================================
const SHEET_NAMES = {
  UI_CALENDAR: 'UI_Calendar',
  UI_MEMO: 'UI_Memo',
  DB_EVENTS: 'DB_Events',
  DB_MEMOS: 'DB_Memos',
  DB_ROKUYO: 'DB_Rokuyo',
  DB_LOG: 'DB_Log',
  SETTINGS: 'Settings'
};

// ===========================================
// スクリプトプロパティキー
// ===========================================
const PROP_KEYS = {
  GEMINI_API_KEY: 'GEMINI_API_KEY',
  GEMINI_MODEL_CALENDAR: 'GEMINI_MODEL_CALENDAR',
  GEMINI_MODEL_MEMO: 'GEMINI_MODEL_MEMO'
};

// ===========================================
// デフォルト値
// ===========================================
const DEFAULTS = {
  TIMEZONE: 'Asia/Tokyo',
  GEMINI_MODEL: 'gemini-2.0-flash'
};

// ===========================================
// シート取得ヘルパー
// ===========================================

/**
 * シートを名前で取得
 * @param {string} name - シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`シート "${name}" が見つかりません`);
  }
  return sheet;
}

/**
 * Settingsシートから設定を読み込む
 * @returns {Object} 設定オブジェクト
 */
function getSettings() {
  const sheet = getSheet(SHEET_NAMES.SETTINGS);
  const data = sheet.getRange('A3:B15').getValues();

  const settings = {};
  for (const row of data) {
    if (row[0]) {
      settings[String(row[0]).trim()] = row[1];
    }
  }

  return {
    weekStart: settings['WeekStart'] || 'Sunday',
    showRokuyo: settings['ShowRokuyo'] === true || settings['ShowRokuyo'] === 'TRUE',
    timezone: settings['Timezone'] || DEFAULTS.TIMEZONE,
    theme: settings['Theme'] || 'Pastel',
    googleCalendarId: settings['GoogleCalendarId'] || '',
    geminiApiKey: String(settings['GeminiApiKey'] || '').trim(),
    geminiModelCalendar: String(settings['GeminiModelCalendar'] || '').trim(),
    geminiModelMemo: String(settings['GeminiModelMemo'] || '').trim(),
    calendarName: String(settings['CalendarName'] || '').trim(),
    imageColor: String(settings['ImageColor'] || '').trim()
  };
}

/**
 * スクリプトプロパティから値を取得
 * @param {string} key - プロパティキー
 * @param {string} defaultValue - デフォルト値
 * @returns {string}
 */
function getProperty(key, defaultValue = '') {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(key) || defaultValue;
}

/**
 * Gemini APIキーを取得
 * Settingsシートの GeminiApiKey を優先、なければスクリプトプロパティにフォールバック
 * @returns {string}
 */
function getGeminiApiKey() {
  const settings = getSettings();
  if (settings.geminiApiKey) {
    return settings.geminiApiKey;
  }
  // フォールバック：スクリプトプロパティ
  const key = getProperty(PROP_KEYS.GEMINI_API_KEY);
  if (!key) {
    throw new Error('GeminiApiKey が Settings シートに設定されていません。Settings シートの A列に「GeminiApiKey」、B列に API キーを入力してください。');
  }
  return key;
}

/**
 * Calendar用AIモデルを取得
 * Settingsシートの GeminiModelCalendar を優先
 * @returns {string}
 */
function getCalendarModel() {
  const settings = getSettings();
  return settings.geminiModelCalendar || getProperty(PROP_KEYS.GEMINI_MODEL_CALENDAR, DEFAULTS.GEMINI_MODEL);
}

/**
 * Memo用AIモデルを取得
 * Settingsシートの GeminiModelMemo を優先
 * @returns {string}
 */
function getMemoModel() {
  const settings = getSettings();
  return settings.geminiModelMemo || getProperty(PROP_KEYS.GEMINI_MODEL_MEMO, DEFAULTS.GEMINI_MODEL);
}

// ===========================================
// ID生成ヘルパー
// ===========================================

/**
 * ランダム文字列を生成
 * @param {number} length - 長さ
 * @returns {string}
 */
function randomString(length = 4) {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * 現在時刻のタイムスタンプ文字列を取得
 * @returns {string} YYYYMMDD_HHMMSS形式
 */
function getTimestampString() {
  const now = new Date();
  const settings = getSettings();
  const tz = settings.timezone;

  const formatted = Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss');
  return formatted;
}

/**
 * 現在時刻を取得（フォーマット済み）
 * @returns {string} YYYY-MM-DD HH:MM:SS形式
 */
function getCurrentDateTime() {
  const now = new Date();
  const settings = getSettings();
  const tz = settings.timezone;

  return Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 今日の日付を取得
 * @returns {string} YYYY-MM-DD形式
 */
function getTodayDate() {
  const now = new Date();
  const settings = getSettings();
  const tz = settings.timezone;

  return Utilities.formatDate(now, tz, 'yyyy-MM-dd');
}

/**
 * イベントIDを生成
 * @returns {string} evt_YYYYMMDD_HHMMSS_xxxx
 */
function newEventId() {
  return `evt_${getTimestampString()}_${randomString(4)}`;
}

/**
 * メモIDを生成
 * @returns {string} mem_YYYYMMDD_HHMMSS_xxxx
 */
function newMemoId() {
  return `mem_${getTimestampString()}_${randomString(4)}`;
}

/**
 * ログIDを生成
 * @returns {string} log_YYYYMMDD_HHMMSS_xxxx
 */
function newLogId() {
  return `log_${getTimestampString()}_${randomString(4)}`;
}

// ===========================================
// カラー見本をSettingsシートに書き込む
// ===========================================

/**
 * Settingsシートの D列〜E列にカラー見本を書き込む
 * GASエディタまたはメニューから手動実行
 */
function writeColorGuide() {
  var sheet = getSheet(SHEET_NAMES.SETTINGS);

  // ヘッダー
  var header = sheet.getRange('D2:E2');
  header.setValues([['カラー名', '見本']]);
  header.setFontWeight('bold').setBackground('#e8eaf6');

  // カラーデータ
  var colors = [
    ['青',         '#1565c0'],
    ['ブルー',     '#1565c0'],
    ['紺',         '#283593'],
    ['ネイビー',   '#283593'],
    ['水色',       '#0288d1'],
    ['緑',         '#2e7d32'],
    ['グリーン',   '#2e7d32'],
    ['エメラルド', '#00897b'],
    ['赤',         '#c62828'],
    ['レッド',     '#c62828'],
    ['ピンク',     '#ad1457'],
    ['ローズ',     '#c2185b'],
    ['紫',         '#6a1b9a'],
    ['パープル',   '#6a1b9a'],
    ['オレンジ',   '#e65100'],
    ['橙',         '#e65100'],
    ['茶',         '#4e342e'],
    ['ブラウン',   '#4e342e'],
    ['グレー',     '#455a64'],
    ['灰',         '#455a64'],
    ['黒',         '#212121'],
    ['ブラック',   '#212121'],
    ['インディゴ', '#5c6bc0']
  ];

  // カラー名を書き込み
  var nameData = colors.map(function(c) { return [c[0]]; });
  sheet.getRange(3, 4, colors.length, 1).setValues(nameData);

  // 見本セルに色を付ける
  for (var i = 0; i < colors.length; i++) {
    var row = i + 3;
    var cell = sheet.getRange(row, 5); // E列
    cell.setValue('').setBackground(colors[i][1]);
  }

  // 列幅調整
  sheet.setColumnWidth(4, 100); // D列
  sheet.setColumnWidth(5, 60);  // E列

  // 使い方メモ
  var noteCell = sheet.getRange(3 + colors.length, 4);
  noteCell.setValue('↑ D列の名前をB列のImageColorにコピペ');
  noteCell.setFontSize(9).setFontColor('#999');

  try { showToast('カラー見本を書き込みました', 'Settings', 3); } catch(e) {}
}

// ===========================================
// ユーティリティ
// ===========================================

/**
 * トースト通知を表示
 * @param {string} message - メッセージ
 * @param {string} title - タイトル
 * @param {number} timeout - 表示秒数
 */
function showToast(message, title = 'AIカレンダー', timeout = 3) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}

/**
 * アラートダイアログを表示
 * @param {string} message - メッセージ
 */
function showAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * 確認ダイアログを表示
 * @param {string} message - メッセージ
 * @returns {boolean} OKが押されたらtrue
 */
function showConfirm(message) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認', message, ui.ButtonSet.OK_CANCEL);
  return response === ui.Button.OK;
}