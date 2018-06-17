/** @OnlyCurrentDoc */

// スプレッドシートの値変更をトリガーに実行する
function spreadBookInfo() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell()
  var currentRow = cell.getRow()
  var currentClolumn = cell.getColumn()

  if (currentClolumn !== 1) {
    return
  }
  var input = cell.getValue()
  if (!input) {
    return
  }
  var val = input
  if (!val.match(/^[0-9]*$/)) {
    val = val.replace(/[０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 65248);
    });
  }
  if (!val.match(/^[0-9]*$/)) {
    Browser.msgBox("1列目にはISBNを入力してください")
    return
  }
  var book = getBook(val);
  var bookArr = [book["isbn"], book["title"], book["publisher"], book["pubdate"], book["author"]]

  var bookInfoColumn = currentClolumn + 1
  for(var i = 0; i < bookArr.length; i++) {
    sheet.getRange(currentRow, bookInfoColumn + i).setValue(bookArr[i]);
  }
}

function getBook(isbn) {
  var Url = buildUrl(isbn)
  var response = JSON.parse(UrlFetchApp.fetch(Url).getContentText())[0];
  if(!response){
    Browser.msgBox("ISBN: " + isbn + "の本は見つかりませんでした")
  }
  var book = response["summary"];

  return book;
}

function buildUrl(isbn) {
  // データが一番豊富なのでAmazon Product Advertising APIを使いたかったが、
  // * Rate Limitが1req/secで厳しい（sleep入れれば対応可能）
  // * アフィリエイトタグとID、トークンの準備がやや面倒
  // * 同一リクエストがRate Limit関係なくこける
  // などあって見送った。
  // サンプルとして残すので今回採用したopenBDが厳しくなったら組み合わせるなりしてください。
  // https://github.com/toshi0607/Amazon-Product-Advertising-API-Client-for-GAS
  //
  // 今回採用したopenBDは
  // * リクエスト形式がシンプル（Amazon Product AdvertisingAPI比）
  // * レスポンスが高速
  // * キーやシークレット不要
  // が特徴。
  // https://openbd.jp/
  // ただ、データはAmazon Product Advertising API ほど豊富ではない。
  var ENDPOINT_URL = "https://api.openbd.jp/v1/get";
  var params = {
    "isbn": isbn,
  };
  var paramArr = [];
  for (var key in params) {
    paramArr.push(encodeURIComponent(key) + "=" + encodeURIComponent(params[key]));
  }
  var queryString = paramArr.join("&");
  var requestUrl = ENDPOINT_URL + "?" + queryString;

  return requestUrl;
}
