// POSTメソッドで受け取ったデータをシートに書き込む

function doPost(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('xxxx'); // シート名を「xxxx」に変更する

    // jsonを配列に変換
    var json_params = JSON.stringify(e);
    var arr_params = JSON.parse(json_params);
    var token = arr_params['parameter']['token'];
    var slack_token = 'xxxx'; // outgoing webhookで取得したトークン「xxxx」を入力する

    // Slackからのリクエストの場合の処理
    if (token == slack_token) {
        stamping_from_slack(arr_params, sheet);
        return;
    }
}

// 書き込み処理
function stamping_from_slack(arr_params, sheet){
    var row = sheet.getLastRow() + 1;
    var numColumns = sheet.getLastColumn();
    var next_row = sheet.getRange(row, 1, 1, numColumns);
    var data = next_row.getValues();

    var date = new Date();
    var hour = ('0' + date.getHours()).slice(-2);
    var min = ('0' + date.getMinutes()).slice(-2);
    var name = arr_params['parameter']['user_name'];
    var action = arr_params['parameter']['trigger_word'];

    data[0][0]  = date;
    data[0][1]  = hour;
    data[0][2]  = min;
    data[0][3]  = name;
    data[0][4]  = action;

    next_row.setValues(data);
    return;
}