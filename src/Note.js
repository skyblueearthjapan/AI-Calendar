/**
 * Note.js - 固定10タブノートシステム
 * DB_Notes シートを使用した固定ノート管理
 *
 * DB_Notes シート構造:
 *   Row 1: タイトル
 *   Row 2: ヘッダー (note_id, tab_name, note_text, created_at, updated_at)
 *   Row 3+: データ
 */

// カラム定義 (0-indexed)
var NOTE_COLS = {
  note_id:    0,
  tab_name:   1,
  note_text:  2,
  created_at: 3,
  updated_at: 4
};

/**
 * デフォルトのタブ名配列を返す (10個固定)
 * @return {string[]}
 */
function getDefaultNoteNames() {
  return [
    'メモ1', 'メモ2', 'メモ3', 'メモ4', 'メモ5',
    'メモ6', 'メモ7', 'メモ8', 'メモ9', 'メモ10'
  ];
}

/**
 * DB_Notes シートに10件のデフォルトノートを初期化する
 * データが既に存在する場合は何もしない (セットアップ時に1回だけ呼び出す)
 */
function initializeNotes() {
  var sheet = getSheet('DB_Notes');
  var lastRow = sheet.getLastRow();

  // Row 3 以降にデータが存在すれば初期化済みとみなす
  if (lastRow >= 3) {
    return;
  }

  var now = getCurrentDateTime();
  var defaultNames = getDefaultNoteNames();
  var rows = [];

  for (var i = 0; i < defaultNames.length; i++) {
    rows.push([
      i + 1,            // note_id (1-10)
      defaultNames[i],  // tab_name
      '',               // note_text (空)
      now,              // created_at
      now               // updated_at
    ]);
  }

  // Row 3 から10行分を一括書き込み
  sheet.getRange(3, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * 全10件のノートを取得する
 * @return {Object[]} ノートオブジェクトの配列
 */
function getAllNotes() {
  var sheet = getSheet('DB_Notes');
  var lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return [];
  }

  var numRows = lastRow - 2; // Row 3 から
  var data = sheet.getRange(3, 1, numRows, 5).getValues();
  var notes = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    notes.push({
      note_id:    row[NOTE_COLS.note_id],
      tab_name:   row[NOTE_COLS.tab_name],
      note_text:  row[NOTE_COLS.note_text],
      created_at: row[NOTE_COLS.created_at],
      updated_at: row[NOTE_COLS.updated_at]
    });
  }

  return notes;
}

/**
 * 指定された note_id のノートを1件取得する
 * @param {number} noteId - ノートID (1-10)
 * @return {Object|null} ノートオブジェクト、見つからない場合は null
 */
function getNoteById(noteId) {
  var sheet = getSheet('DB_Notes');
  var lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return null;
  }

  var numRows = lastRow - 2;
  var data = sheet.getRange(3, 1, numRows, 5).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[NOTE_COLS.note_id] == noteId) {
      return {
        note_id:    row[NOTE_COLS.note_id],
        tab_name:   row[NOTE_COLS.tab_name],
        note_text:  row[NOTE_COLS.note_text],
        created_at: row[NOTE_COLS.created_at],
        updated_at: row[NOTE_COLS.updated_at]
      };
    }
  }

  return null;
}

/**
 * 指定された note_id のノートテキストを保存する
 * @param {number} noteId - ノートID (1-10)
 * @param {string} text - 保存するテキスト
 */
function saveNoteText(noteId, text) {
  var sheet = getSheet('DB_Notes');
  var lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return;
  }

  var numRows = lastRow - 2;
  var data = sheet.getRange(3, 1, numRows, 1).getValues(); // note_id 列のみ取得

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == noteId) {
      var rowIndex = i + 3; // シート上の実際の行番号
      sheet.getRange(rowIndex, NOTE_COLS.note_text + 1).setValue(text);
      sheet.getRange(rowIndex, NOTE_COLS.updated_at + 1).setValue(getCurrentDateTime());
      return;
    }
  }
}

/**
 * 指定された note_id のタブ名を変更する
 * @param {number} noteId - ノートID (1-10)
 * @param {string} newName - 新しいタブ名
 */
function saveNoteName(noteId, newName) {
  var sheet = getSheet('DB_Notes');
  var lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return;
  }

  var numRows = lastRow - 2;
  var data = sheet.getRange(3, 1, numRows, 1).getValues(); // note_id 列のみ取得

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == noteId) {
      var rowIndex = i + 3; // シート上の実際の行番号
      sheet.getRange(rowIndex, NOTE_COLS.tab_name + 1).setValue(newName);
      sheet.getRange(rowIndex, NOTE_COLS.updated_at + 1).setValue(getCurrentDateTime());
      return;
    }
  }
}
