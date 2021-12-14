function getBookStatus(isbn, title, appkey, systemid) {
    var res_finish = false;
    var no_registed = false;
    var res_json = '';
    while (true) {
        res_json = callCalilApi(appkey, isbn, systemid);
        if (isFinish(res_json)) {
            Utilities.sleep(1000);
            break;
        }
    }

    //var rj = JSON.parse(res_json); 
    no_registed = res_json.match(/"libkey": {}/);
    //Logger.log('rj = ' + res_json + 'no_registed = ' + no_registed );

    if (no_registed) {
        Logger.log('【未登録】ISBN:[' + isbn + ']の[' + title + ']は川越市の図書館に登録されていません。');
        return '-';
    } else {
        var msg = '【登録済】ISBN:[' + isbn + ']の[' + title + ']は川越市の図書館に登録されてます。'
        Logger.log(msg);
        return '蔵書あり';
        //notifyLine(msg)
    }
}

function isFinish(res_json) {
    return res_json.match(/"continue": 0/);
}

function callCalilApi(appkey, isbn, systemid) {
    var response = UrlFetchApp.fetch('http://api.calil.jp/check?appkey=' + appkey + '&isbn=' + isbn + '&systemid=' + systemid + '&format=json');
    res_json = response.toString().slice(9, -2);
    Logger.log(res_json);


    return res_json;
}

function scrapBooklog() {
    //スクレイピングしたいWebページのURLを変数で定義する
    let url = "https://booklog.jp/users/soralibrary";
    //URLに対しフェッチを行ってHTMLデータを取得する
    let html = UrlFetchApp.fetch(url).getContentText("UTF-8");

    //Parserライブラリを使用して条件を満たしたHTML要素を抽出する
    //本の大雑把な情報をリスト形式で取得
    let parser = Parser.data(html)
        .from('<div class="item-wrapper shelf-item item-init"')
        .to('<a href')
        .iterate();
    //.build();

    //スクレイピングで読み込んだ本の数
    //console.log('スクレイピングで読み込んだ本の数',parser.length)

    // スプレッドシートから読み込む
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // アクティブなシートを取得
    var sh = ss.getActiveSheet();
    Utilities.sleep(100);

    //カラムの場所を取得
    var checkcolumns = sh.getRange('A1:M1').getValues();
    for (var i = 0; i < 14; i++) {
        Utilities.sleep(100);
        colum_name = checkcolumns[0][i]
        if (colum_name == 'タイトル') {
            var idx_title = i
        } else if (colum_name == 'ISBN') {
            var idx_isbn = i
        } else if (colum_name == '蔵書有無') {
            var idx_check = i
        } else if (colum_name == '読書状況') {
            var idx_status = i
        } else if (colum_name == 'ブクログ順') {
            var idx_bukunum = i
        }
    }
    console.log('title', idx_title, 'isnb', idx_isbn, 'check', idx_check, 'status', idx_status)
    Utilities.sleep(1000)
    //getRangeで範囲を指定し、getValuesで値を取得
    var isbn_sheets = sh.getRange(2, idx_isbn + 1, 300)
    var isbn_sheet = isbn_sheets.getValues();
    console.log(isbn_sheet.flat());
    const bukunum = isbn_sheet.flat();
    var number = 0;
    //isbn_sheetのisbnの数を数える
    for (i in bukunum) {
        if (bukunum[i] !== '') {
            number += 1
        }
    }

    //本の数だけ実行
    for (var i = parser.length - 1; i >= 0; i--) {
        console.log('---------------------------------------------------------------')

        //読書ステータスの取得（読書状況）
        let status = parser[i].match(/status&quot;:&quot;\d&quot;/)
        if (status) {
            book_status = status[0].match(/\d/)[0]
            if (book_status == 0) { book_status = '未設定'; }
            else if (book_status == 1) { book_status = '読みたい'; }
            else if (book_status == 2) { book_status = '読んでる'; }
            else if (book_status == 3) { book_status = '読み終わった'; }
            else if (book_status == 4) { book_status = '積読' }
            else { book_status = '-' }

        } else {
            book_status = '-';
            console.log('読書状況なし');
        }

        //タイトルの取得
        let title = parser[i].match(/title=".*">/);
        if (title) {
            book_title = title[0].replace(/(title=")(.*)(">)/, '$2');

            console.log(book_title);
        } else {
            book_title = '-';
            console.log('タイトルなし');
        }

        //EANからISBNを抽出
        let isbn = parser[i].match(/EAN&quot;:&quot;\d{13}/);
        if (isbn) {
            let book_isbn = isbn[0].replace("EAN&quot;:&quot;", "");
            console.log(book_isbn, typeof (book_isbn));

            var isbn_sheet = sh.getRange(2, idx_isbn + 1, 300).getValues();

            //ISBNがシートにない場合新しくシート追加して詳細記入(文字と数字で検索)
            var sheet_index = isbn_sheet.flat().indexOf(book_isbn);
            if (sheet_index < 0) {
                sheet_index = isbn_sheet.flat().indexOf(+book_isbn);
            }
            console.log(sheet_index)
            if (sheet_index < 0) {
                //シート追加
                Utilities.sleep(1500);
                sh.insertRowBefore(2);
                //シート記入
                Utilities.sleep(1500);
                sh.getRange(2, idx_title + 1).setValue(book_title);
                Utilities.sleep(1500);
                sh.getRange(2, idx_isbn + 1).setValue(book_isbn);
                Utilities.sleep(1500);
                sh.getRange(2, idx_check + 1).setValue('-');
                Utilities.sleep(1500);
                sh.getRange(2, idx_status + 1).setValue(book_status);
                Utilities.sleep(1500);
                //ブクログナンバーの更新と記入
                number += 1
                sh.getRange(2, idx_bukunum + 1).setValue(number);
                SpreadsheetApp.flush();

                //シート記入結果の出力
                console.log(idx_title, idx_isbn, idx_check, idx_status)
                console.log('シート入力した値　　', book_title, book_isbn, '-', book_status)
            } else {
                Utilities.sleep(1000);
                sh.getRange(sheet_index + 2, idx_status + 1).setValue(book_status);
                console.log('読書状況のみ更新', book_status, sheet_index + 2, idx_status + 1)
                SpreadsheetApp.flush();
            }
        } else {
            console.log('ISBNなし');
        }
    }//forの終わり

}

function notifyLine(message) {
    var TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_API_TOKEN");
    var res = null;
    var options = {
        "method": "post",
        "payload": "message=" + message,
        "headers": { "Authorization": "Bearer " + TOKEN }
    };
    res = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
    return res;
}

function main() {
    scrapBooklog()

    // スプレッドシートから読み込む
    Utilities.sleep(1000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // アクティブなシートを取得
    var sh = ss.getActiveSheet();

    console.log('sh', sh.length)

    //カラムの場所を取得(A1~M1まで)
    var checkcolumns = sh.getRange('A1:M1').getValues();
    for (var i = 0; i < 14; i++) {
        Utilities.sleep(100);
        colum_name = checkcolumns[0][i]
        if (colum_name == 'タイトル') {
            var idx_title = i
        } else if (colum_name == 'ISBN') {
            var idx_isbn = i
        } else if (colum_name == '蔵書有無') {
            var idx_check = i
        } else if (colum_name == '読書状況') {
            var idx_status = i
        }
    }

    //getRangeで範囲を指定し、getValuesで値を取得
    var checkbooks = sh.getRange('A2:M300').getValues();

    // 固定
    var appkey = PropertiesService.getScriptProperties().getProperty("CALIL_APP_KEY");
    var systemid = 'Saitama_Kawagoe';

    //蔵書確認（checkbooksの範囲内）
    for (var i = 0; i < checkbooks.length; i++) {

        // '蔵書あり'または''または文字以外でcontinue以降をスキップ
        if (checkbooks[i][idx_check] === "蔵書あり" || isNaN(checkbooks[i][idx_isbn]) || checkbooks[i][idx_isbn] === '') {
            continue;
        }

        //蔵書確認をして本があった場合スプレッドシートに蔵書ありと書き込む
        if (getBookStatus(checkbooks[i][idx_isbn], checkbooks[i][idx_title], appkey, systemid) === "蔵書あり") {
            Utilities.sleep(1000);
            sh.getRange(i + 2, idx_check + 1).setValue("蔵書あり");
            SpreadsheetApp.flush();
        }
        console.log(i, '番目');
    }
}