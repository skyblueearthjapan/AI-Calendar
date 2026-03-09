/**
 * ReservationSync.gs
 * チームスケジュールからの予約キュー読み取り・反映処理
 */

// チームスケジュール側のDB_ReservationQueueカラム定義
const RSV_COLS = {
  RESERVATION_ID: 0, CREATED_AT: 1, MEMBER_NAME: 2, MEMBER_SSID: 3,
  ACTION: 4, EVENT_ID: 5, TITLE: 6, START_DATE: 7, END_DATE: 8,
  START_TIME: 9, END_TIME: 10, ALL_DAY: 11, MEMO: 12, COLOR_KEY: 13,
  STATUS: 14, APPLIED_AT: 15, ERROR_MESSAGE: 16, REQUESTED_BY: 17
};

/**
 * 自分宛てのpending予約を取得
 * @returns {Array} 予約データの配列
 */
function getMyPendingReservations() {
  try {
    const settings = getSettings();
    const teamSSID = settings.teamScheduleSSID;

    if (!teamSSID) {
      return [];
    }

    // 自身のスプレッドシートID
    const mySSID = SpreadsheetApp.getActiveSpreadsheet().getId();

    // チームスケジュールのスプレッドシートを開く
    const teamSS = SpreadsheetApp.openById(teamSSID);
    const sheet = teamSS.getSheetByName('DB_ReservationQueue');

    if (!sheet) {
      console.log('DB_ReservationQueue sheet not found in team schedule');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    const data = sheet.getRange(3, 1, lastRow - 2, 18).getValues();
    const tz = settings.timezone;
    const reservations = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const memberSSID = String(row[RSV_COLS.MEMBER_SSID] || '').trim();
      const status = String(row[RSV_COLS.STATUS] || '').trim();

      // 自分宛て かつ pending のみ
      if (memberSSID === mySSID && status === 'pending') {
        reservations.push({
          reservation_id: String(row[RSV_COLS.RESERVATION_ID] || ''),
          created_at: String(row[RSV_COLS.CREATED_AT] || ''),
          member_name: String(row[RSV_COLS.MEMBER_NAME] || ''),
          action: String(row[RSV_COLS.ACTION] || ''),
          event_id: String(row[RSV_COLS.EVENT_ID] || ''),
          title: String(row[RSV_COLS.TITLE] || ''),
          start_date: formatReservationDate_(row[RSV_COLS.START_DATE], tz),
          end_date: formatReservationDate_(row[RSV_COLS.END_DATE], tz),
          start_time: formatReservationTime_(row[RSV_COLS.START_TIME]),
          end_time: formatReservationTime_(row[RSV_COLS.END_TIME]),
          all_day: row[RSV_COLS.ALL_DAY] === 'TRUE' || row[RSV_COLS.ALL_DAY] === true,
          memo: String(row[RSV_COLS.MEMO] || ''),
          color_key: String(row[RSV_COLS.COLOR_KEY] || 'other'),
          row_index: i + 3 // シート行番号（1-based）
        });
      }
    }

    return reservations;
  } catch (e) {
    console.error('getMyPendingReservations error:', e);
    return [];
  }
}

/**
 * pending予約の件数を取得
 * @returns {number}
 */
function getPendingReservationCount_() {
  try {
    return getMyPendingReservations().length;
  } catch (e) {
    return 0;
  }
}

/**
 * 予約を承認して適用
 * @param {string} reservationId - 予約ID
 * @returns {Object} 結果
 */
function applyReservation(reservationId) {
  try {
    const reservations = getMyPendingReservations();
    const reservation = reservations.find(r => r.reservation_id === reservationId);

    if (!reservation) {
      return { success: false, error: '予約が見つかりません' };
    }

    let result;
    const action = reservation.action;

    if (action === 'create') {
      // 新規予定として登録
      const eventData = {
        title: reservation.title,
        start_date: reservation.start_date,
        end_date: reservation.end_date || reservation.start_date,
        start_time: reservation.start_time || '',
        end_time: reservation.end_time || '',
        all_day: reservation.all_day,
        memo: reservation.memo || '',
        color_key: reservation.color_key || 'other'
      };

      const eventId = insertEventToDB(eventData, 'チームスケジュールからの予約', 'reservation');
      result = { success: true, message: `予定「${reservation.title}」を登録しました`, eventId: eventId };

    } else if (action === 'update') {
      // 既存予定を更新
      if (!reservation.event_id) {
        return { success: false, error: '更新対象のevent_idがありません' };
      }

      const payload = {
        event_id: reservation.event_id,
        title: reservation.title,
        start_date: reservation.start_date,
        end_date: reservation.end_date,
        start_time: reservation.start_time,
        end_time: reservation.end_time,
        all_day: reservation.all_day,
        memo: reservation.memo,
        color_key: reservation.color_key
      };

      result = updateEvent(payload);

    } else if (action === 'delete') {
      // 予定を削除
      if (!reservation.event_id) {
        return { success: false, error: '削除対象のevent_idがありません' };
      }

      result = deleteEvent(reservation.event_id);

    } else {
      return { success: false, error: '不明なアクション: ' + action };
    }

    // チームSSの予約ステータスを更新
    updateTeamReservationStatus_(reservationId, 'applied', '');

    return result;
  } catch (e) {
    console.error('applyReservation error:', e);
    // エラー時もステータス更新を試みる
    try {
      updateTeamReservationStatus_(reservationId, 'error', e.message);
    } catch (e2) {
      console.error('Status update error:', e2);
    }
    return { success: false, error: e.message };
  }
}

/**
 * 予約を拒否
 * @param {string} reservationId - 予約ID
 * @returns {Object} 結果
 */
function rejectReservation(reservationId) {
  try {
    updateTeamReservationStatus_(reservationId, 'rejected', '');
    return { success: true, message: '予約を拒否しました' };
  } catch (e) {
    console.error('rejectReservation error:', e);
    return { success: false, error: e.message };
  }
}

/**
 * チームスケジュールのDB_ReservationQueueのステータスを更新
 * @param {string} reservationId - 予約ID
 * @param {string} status - 新ステータス
 * @param {string} errorMsg - エラーメッセージ
 */
function updateTeamReservationStatus_(reservationId, status, errorMsg) {
  const settings = getSettings();
  const teamSSID = settings.teamScheduleSSID;

  if (!teamSSID) {
    throw new Error('TeamScheduleSSIDが設定されていません');
  }

  const teamSS = SpreadsheetApp.openById(teamSSID);
  const sheet = teamSS.getSheetByName('DB_ReservationQueue');

  if (!sheet) {
    throw new Error('DB_ReservationQueueシートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const data = sheet.getRange(3, 1, lastRow - 2, 18).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][RSV_COLS.RESERVATION_ID]) === reservationId) {
      const rowIdx = i + 3;
      sheet.getRange(rowIdx, RSV_COLS.STATUS + 1).setValue(status);

      if (status === 'applied') {
        const now = Utilities.formatDate(new Date(), settings.timezone, 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(rowIdx, RSV_COLS.APPLIED_AT + 1).setValue(now);
      }

      if (errorMsg) {
        sheet.getRange(rowIdx, RSV_COLS.ERROR_MESSAGE + 1).setValue(errorMsg);
      }
      return;
    }
  }
}

// ===========================================
// ヘルパー関数
// ===========================================

function formatReservationDate_(value, tz) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function formatReservationTime_(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'GMT', 'HH:mm');
  }
  const s = String(value).trim();
  if (!s) return '';
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return String(m[1]).padStart(2, '0') + ':' + m[2];
  return s;
}
