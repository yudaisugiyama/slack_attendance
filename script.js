// 書き込み処理
function updateSheet(d, sheet){
  const lastRow = sheet.getLastRow();
  
  const [date, time] = getTime(d);
  const name = getName(d);
  const status = d['parameter']['trigger_word'];

  if (status === '記録提出') {
      // 記録提出の場合は勤務時間を計算する
      postSheet(name);
  } else {
      // 出勤履歴のシートに書き込み
      const out = [[date, time, name, status]];
      sheet.getRange(lastRow + 1, 1, out.length, out[0].length).setValues(out);
  }
  return;
}

// 日付と時刻を取得する
function getTime(d) {
  // JavaScriptのDateオブジェクトに変換
  const ts = d['parameter']['timestamp'];
  var date = new Date(ts * 1000); 

  // UTC時間から日本時間に変換する
  date.setHours(date.getHours());

  // 年月日時分秒を取得する
  const month = ("0" + (date.getMonth() + 1)).slice(-2);
  const day = ("0" + date.getDate()).slice(-2);
  const hour = ("0" + date.getHours()).slice(-2);
  const minute = ("0" + date.getMinutes()).slice(-2);

  // フォーマットして出力する
  const formattedDate = `${month}/${day}`;
  const formattedTime = `${hour}:${minute}`;

  return [formattedDate, formattedTime];
}

// ユーザー名を取得する
function getName(d) {
  const user_name = d['parameter']['user_name'];
  const userNames = getUserNames();
  const name = userNames[user_name] || 'Unknown';
  return name;
}

// 従業員一覧シートからユーザー名情報を取得する
function getUserNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('従業員一覧');
  const range = sheet.getDataRange();
  const values = range.getValues();

  const userNames = {};
  for (const row of values) {
    const [user_name, name] = row;
    userNames[user_name] = name;
  }

  return userNames;
}

// 設定シートから年度と月を取得する
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const range = sheet.getDataRange();
  const values = range.getValues();

  const settings = {};
  for (const row of values) {
    const [key, value] = row;
    settings[key] = value;
  }

  return settings;  
}

// 個人のシートに書き込む
function postSheet(name) {
  // スプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 同じ名前のシートが存在する場合はそのシートを取得し、存在しない場合は新しいシートを作成
  var newSheet = ss.getSheetByName(name);
  if (!newSheet) {
      // 個人のシートを作成
      newSheet = ss.insertSheet();
      newSheet.setName(name);
  }

  // 全てのデータを取得
  var range = newSheet.getDataRange();
  
  // データを消去
  range.clearContent();

  // 個人のシートにデータを書き込む処理を記述
  const settings = getSettings();
  const setYear = settings['year'];
  const setMonth = settings['month'];
  out = calcTime(name, setYear, setMonth);
  
  // 個人のシートに書き込み
  const lastRow = newSheet.getLastRow();
  newSheet.getRange(lastRow + 1, 1, out.length, out[0].length).setValues(out);

  // // 現在スプレッドシートが格納されているフォルダをエクスポート先に指定するためにIDを取得
  // const parentFolder = DriveApp.getFileById(ssid).getParents();
  // const folderid = parentFolder.next().getId();
  
  // // pdfに変換して取得
  // const ssid = ss.getId();
  // const pdf = ss.getAs('application/pdf').setName(name);

  // // PDFをエクスポート
  // DriveApp.getFolderById(folderid).createFile(pdf);
  return;
}

function calcTime(name, setYear, setMonth) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('勤怠履歴');
  const data = sheet.getDataRange().getValues();
  var out = [];
  out.push(["出勤簿 令和", setYear.toString(), "年", setMonth.toString(), "月度", ""]);
  out.push(["株式会社", "会津コンピュータサイエンス研究所", "氏名:", name.toString(), "", ""]);
  out.push(["日", "開始時刻", "終了時刻", "除外時間", "労働時間", "備考"]);
  var startTime = null;
  var endTime = null;
  var totalHours = 0; // 勤務時間合計
  var totalBreakTime = 0; // 除外時間合計
  var breakStart1 = null; // 1階休憩開始時刻
  var breakEnd1 = null; // 1階休憩終了時刻
  var breakStart2 = null; // 2階休憩開始時刻
  var breakEnd2 = null; // 2階休憩終了時刻

  for (let i = 0; i < data.length; i++) {
    let [date, time, person, status] = data[i];
    date = date.toLocaleDateString();
    time = time.toLocaleTimeString();

    if (person === name) {
      if (status === "開始") {
        startTime = new Date(`${date} ${time}`);
      } else if (status === "終了") {
          endTime = new Date(`${date} ${time}`);
          totalHours += (endTime - startTime) / (1000 * 60 * 60);
          // その日の勤務時間を出力
          const month = ("0" + (startTime.getMonth() + 1)).slice(-2);
          const day = ("0" + startTime.getDate()).slice(-2);
          const startHour = ("0" + startTime.getHours()).slice(-2);
          const startMinute = ("0" + startTime.getMinutes()).slice(-2);
          const formattedDate = `${month}/${day}`;
          startTime = `${startHour}:${startMinute}`;
          const endHour = ("0" + endTime.getHours()).slice(-2);
          const endMinute = ("0" + endTime.getMinutes()).slice(-2);
          endTime = `${endHour}:${endMinute}`;
          out.push([formattedDate.toString(), startTime.toString(), endTime.toString(), totalBreakTime.toFixed(1).toString(), totalHours.toFixed(1).toString(), ""]);
          // 1階休憩開始/終了時刻をリセット
          breakStart1 = null;
          breakEnd1 = null;
          // 2階休憩開始/終了時刻をリセット
          breakStart2 = null;
          breakEnd2 = null;
          // 勤務時間をリセット
          totalHours = 0;
          // 除外時間をリセット
          totalBreakTime = 0;
      } else if (status === "休憩" || status === "再開") {
        let breakTime = new Date(`${date} ${time}`);
        if ((breakTime.getHours() <= 14)) {
          // 1階休憩
          if (breakStart1 === null) {
            // 1階休憩開始時刻を記録
            breakStart1 = breakTime;
          } else {
            // 1階休憩終了時刻を記録
            breakEnd1 = breakTime;
            // 休憩時間を計算してtotalHoursから差し引く
            totalBreakTime += (breakEnd1 - breakStart1) / (1000 * 60 * 60);
            totalHours -= (breakEnd1 - breakStart1) / (1000 * 60 * 60);
          }
        } else {
          // 2階休憩
          if (breakStart2 === null) {
            // 2階休憩開始時刻を記録
            breakStart2 = breakTime;
          } else {
            // 2階休憩終了時刻を記録
            breakEnd2 = breakTime;
            // 休憩時間を計算してtotalHoursから差し引く
            totalBreakTime += (breakEnd2 - breakStart2) / (1000 * 60 * 60);
            totalHours -= (breakEnd2 - breakStart2) / (1000 * 60 * 60);
          }
        }
      }
    }
  }

  return out;
}



// メイン関数
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('勤怠履歴');

  // Slackからのデータを処理
  var data = JSON.stringify(e);
  data = JSON.parse(data);
  
  // Outgoing Webhookで生成したトークンを入力
  const secret_token = 'bNO11I3pc0AE8JBaNzFzERLu';
  
  // トークンで認証
  const token = data['parameter']['token'];
  if (token === secret_token) {
      // 処理
      updateSheet(data, sheet);
  }
}