const SETTINGS_SHEET_NAME = '設定';  // 設定シートの名前
const DATA_SHEET_NAME = '返却期限';  // 返却期限データのシート名
const ERROR_LOG_SHEET_NAME = 'エラーログ';  // エラーログ出力用シートの名前

/**
 * スプレッドシートから設定情報を取得し、返却期限が明日の本をLINEに通知する
 */
const checkReturnDatesAndNotify = () => {
  try {
    const config = getConfig();  // 設定情報を取得

    if (!config || !config.URL || !config.ACCESS_TOKEN || !config.TO_USER_ID) {
      logError("設定情報が正しく取得できませんでした。設定シートの内容を確認してください。");
      return;
    }

    // 今日の日付に1日追加して "yyyy/mm/dd" 形式で取得（前日の通知）
    const tomorrow = formatDate(addDays(new Date(), 1));
    const returnBooks = getBooksDueOnDate(tomorrow);  // 返却期限が明日の本を取得

    if (returnBooks.length > 0) {  // 明日返却の本があれば通知
      const message = `以下の本の返却期限が明日です:\n${returnBooks.join('\n')}\n\nhttps://libweb.city.setagaya.tokyo.jp/rentallist`;
      sendMessageToLINE(config, message);
    }
  } catch (error) {
    logError(`checkReturnDatesAndNotify関数でエラー: ${error.message}`);
  }
};

/**
 * 設定シートからLINE APIに必要な情報を取得
 * @returns {Object} - URL、アクセストークン、送信先ユーザーID
 */
const getConfig = () => {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      logError("設定シートが見つかりません。");
      return null;
    }

    const data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();  // A列とB列の全データ取得
    const config = {};
    data.forEach(([key, value]) => {
      config[key] = value;
    });

    return config;
  } catch (error) {
    logError(`getConfig関数でエラー: ${error.message}`);
    return null;
  }
};

/**
 * 指定された日付が返却期限の本のリストを取得
 * @param {string} targetDate - 指定された日付 (yyyy/mm/dd)
 * @returns {Array} - 指定された日付が返却期限の本のタイトル一覧
 */
const getBooksDueOnDate = (targetDate) => {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
    if (!sheet) {
      logError("返却期限シートが見つかりません。");
      return [];
    }

    const data = sheet.getDataRange().getValues();  // 全データ取得
    const dueBooks = data.slice(1)  // 1行目のヘッダーをスキップ
      .filter(row => {
        const title = row[3]; // D列のタイトル
        const dueDate = row[6]; // G列の返却日
        const dueDateString = formatDate(dueDate); // 返却日を整形
        return dueDateString === targetDate; // 指定された日付と比較
      })
      .map(row => row[3]); // D列のタイトルを取得

    return dueBooks;
  } catch (error) {
    logError(`getBooksDueOnDate関数でエラー: ${error.message}`);
    return [];
  }
};

/**
 * 日付を "yyyy/mm/dd" 形式にフォーマットする
 * @param {Date} date - フォーマットする日付
 * @returns {string} - フォーマットされた日付文字列
 */
const formatDate = (date) => {
  if (!(date instanceof Date)) return date; // 日付でない場合はそのまま返す

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');  // 月を2桁に
  const day = String(date.getDate()).padStart(2, '0');  // 日を2桁に
  return `${year}/${month}/${day}`;
};

/**
 * 日付に指定した日数を加算
 * @param {Date} date - 元の日付
 * @param {number} days - 加算する日数
 * @returns {Date} - 日付を加算した結果
 */
const addDays = (date, days) => {
  const newDate = new Date(date);
  newDate.setDate(newDate.getDate() + days);
  return newDate;
};

/**
 * LINEにメッセージを送信する
 * @param {Object} config - 設定情報（URL、アクセストークン、送信先ID）
 * @param {string} message - 送信するメッセージ内容
 */
const sendMessageToLINE = (config, message) => {
  try {
    if (!config || !message) {
      logError("sendMessageToLINE関数に必要な引数が渡されていません。");
      return;
    }

    UrlFetchApp.fetch(config.URL, {
      method: 'post',
      headers: {
        "Content-Type": "application/json; charset=UTF-8",
        'Authorization': 'Bearer ' + config.ACCESS_TOKEN,
      },
      payload: JSON.stringify({
        to: config.TO_USER_ID,
        messages: [
          {
            type: 'text',
            text: message,
          }
        ]
      })
    });
  } catch (error) {
    logError(`sendMessageToLINE関数でエラー: ${error.message}`);
  }
};

/**
 * エラーログをスプレッドシートの「エラーログ」シートに記録
 * @param {string} errorMessage - 記録するエラーメッセージ
 */
const logError = (errorMessage) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERROR_LOG_SHEET_NAME) ||
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ERROR_LOG_SHEET_NAME);

  sheet.appendRow([new Date().toLocaleString(), errorMessage]);  // 日時とエラーメッセージを記録
};
