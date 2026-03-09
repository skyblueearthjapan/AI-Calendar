/**
 * SyncToTeam.gs
 * AIカレンダーからチームスケジュールへの手動同期
 */

/**
 * 自分の予定をチームスケジュールのDB_TeamScheduleに同期
 * @returns {Object} 結果
 */
function syncMyEventsToTeam() {
  try {
    const settings = getSettings();
    const teamSSID = settings.teamScheduleSSID;

    if (!teamSSID) {
      return { success: false, error: 'TeamScheduleSSIDが設定されていません。Settingsシートに追加してください。' };
    }

    const mySSID = SpreadsheetApp.getActiveSpreadsheet().getId();
    const tz = settings.timezone;

    // チームスケジュールのSETTINGSからメンバー名を取得
    const teamSS = SpreadsheetApp.openById(teamSSID);
    const teamSettings = teamSS.getSheetByName('SETTINGS');
    if (!teamSettings) {
      return { success: false, error: 'チームスケジュールのSETTINGSシートが見つかりません' };
    }

    const tsData = teamSettings.getDataRange().getValues();
    let myMemberName = '';
    for (let i = 0; i < tsData.length; i++) {
      const key = String(tsData[i][0] || '').trim();
      const val = String(tsData[i][1] || '').trim();
      // Member1_SSID 〜 Member5_SSID で自分のSSIDを探す
      if (key.match(/^Member\d+_SSID$/) && val === mySSID) {
        const num = key.replace('Member', '').replace('_SSID', '');
        // 対応するMemberN_Nameを探す
        for (let j = 0; j < tsData.length; j++) {
          if (String(tsData[j][0]).trim() === 'Member' + num + '_Name') {
            myMemberName = String(tsData[j][1] || '').trim();
            break;
          }
        }
        break;
      }
    }

    if (!myMemberName) {
      return { success: false, error: 'チームスケジュールに自分のメンバー登録が見つかりません' };
    }

    // 自分のDB_Eventsからactive予定を取得
    const eventsSheet = getSheet(SHEET_NAMES.DB_EVENTS);
    const evData = eventsSheet.getDataRange().getValues();
    const myEvents = [];

    for (let i = 2; i < evData.length; i++) {
      const row = evData[i];
      const status = String(row[EVENT_COLS.STATUS] || '').trim().toLowerCase();
      if (status !== 'active') continue;

      let startDate = row[EVENT_COLS.START_DATE];
      let endDate = row[EVENT_COLS.END_DATE];
      let startTime = row[EVENT_COLS.START_TIME];
      let endTime = row[EVENT_COLS.END_TIME];

      if (startDate instanceof Date) startDate = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
      else startDate = String(startDate || '').trim();

      if (endDate instanceof Date) endDate = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');
      else endDate = String(endDate || '').trim();

      if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, 'GMT', 'HH:mm');
      else startTime = String(startTime || '').trim();

      if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, 'GMT', 'HH:mm');
      else endTime = String(endTime || '').trim();

      myEvents.push({
        member_name: myMemberName,
        event_id: String(row[EVENT_COLS.EVENT_ID] || ''),
        title: String(row[EVENT_COLS.TITLE] || ''),
        start_date: startDate,
        end_date: endDate,
        start_time: startTime,
        end_time: endTime,
        all_day: row[EVENT_COLS.ALL_DAY] === 'TRUE' || row[EVENT_COLS.ALL_DAY] === true,
        memo: String(row[EVENT_COLS.MEMO] || ''),
        color_key: String(row[EVENT_COLS.COLOR_KEY] || 'other')
      });
    }

    // チームスケジュールのDB_TeamScheduleに書き込み
    const tsSheet = teamSS.getSheetByName('DB_TeamSchedule');
    if (!tsSheet) {
      return { success: false, error: 'DB_TeamScheduleシートが見つかりません' };
    }

    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    const tsLastRow = tsSheet.getLastRow();

    // 自分のメンバー行だけを削除して書き直す
    if (tsLastRow >= 3) {
      const existingData = tsSheet.getRange(3, 1, tsLastRow - 2, 11).getValues();
      const rowsToDelete = [];
      for (let i = existingData.length - 1; i >= 0; i--) {
        if (String(existingData[i][0]).trim() === myMemberName) {
          rowsToDelete.push(i + 3);
        }
      }
      // 下から削除（行番号ずれ防止）
      for (let r = 0; r < rowsToDelete.length; r++) {
        tsSheet.deleteRow(rowsToDelete[r]);
      }
    }

    // 新しいデータを追加
    if (myEvents.length > 0) {
      const rows = myEvents.map(function(e) {
        return [
          e.member_name, e.event_id, e.title,
          e.start_date, e.end_date, e.start_time, e.end_time,
          e.all_day ? 'TRUE' : 'FALSE', e.memo, e.color_key, now
        ];
      });
      const newLastRow = tsSheet.getLastRow();
      tsSheet.getRange(newLastRow + 1, 1, rows.length, 11).setValues(rows);
    }

    return {
      success: true,
      eventCount: myEvents.length,
      message: myEvents.length + '件の予定をチームスケジュールに同期しました'
    };
  } catch (e) {
    console.error('syncMyEventsToTeam error:', e);
    return { success: false, error: e.message };
  }
}
