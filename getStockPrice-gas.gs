function getStockPriceGoogle(code) {
  let url = "https://www.google.com/finance/quote/" + code + ":TYO";
  let html = UrlFetchApp.fetch(url).getContentText();
  let stockPrice = Parser.data(html)
    .from('<div class="YMlKec fxKbKc">\xA5')
    .to("</div>")
    .build();
  return stockPrice;
}

function getPriceToshin(code) {
  let url = "https://www.rakuten-sec.co.jp/web/fund/detail/?ID=" + code;
  let html = UrlFetchApp.fetch(url).getContentText();
  let price = Parser.data(html)
    .from('<span class="value-01">')
    .to("</span>")
    .build();
  return price;
}

function getCellValue(cell) {
  //親スプレッドシートの先頭シートの指定セルの値を取得し返却する
  return SpreadsheetApp.getActive().getSheets()[0].getRange(cell).getValue();
}

function insertCellValue(cell, value) {
  //親スプレッドシートの先頭シートの指定セルに値を挿入する
  SpreadsheetApp.getActive().getSheets()[0].getRange(cell).setValue(value);
}

// 日本株の株価の取得
function updateStockPrices(inputCell, outputCell) {
  let stockCode = getCellValue(inputCell); //銘柄コードのセルから銘柄コード取得
  let stockPrice = getStockPriceGoogle(stockCode); //銘柄コードから株価取得
  insertCellValue(outputCell, stockPrice); //株価をセルに挿入
}

// 投資信託(日本)の基準価格の取得
function updateToshinPrices(inputCell, outputCell) {
  let toshinCode = getCellValue(inputCell); //銘柄コードのセルから銘柄コード取得
  let toshinPrice = getPriceToshin(toshinCode); //銘柄コードから投資信託の基準価格を取得
  insertCellValue(outputCell, toshinPrice); //基準価格をセルに挿入
}

// 日本株と投資信託を取得する(main)
function updatePrices() {
  // C列:証券コード
  // D列:株価
  //日本株
  updateStockPrices("C4", "D4");
  updateStockPrices("C5", "D5");
  //投資信託
  updateToshinPrices("C6", "D6");
  updateToshinPrices("C7", "D7");
}

function updateStockPriceList() {
  //株価シート：時価評価額セル
  const eDataCell = {
    GOOGLE: "I2",
    VOO: "I3",
    SP500_ETF: "I4",
    SOFTBANK: "I5",
    EMS_SP: "I6",
    EMS_ALL: "I7",
  };
  //表シート：出力カラム
  const eColumn = {
    DATE: 1,
    GOOGLE: 2,
    VOO: 3,
    SP500_ETF: 4,
    SOFTBANK: 5,
    EMS_SP: 6,
    EMS_ALL: 7,
  };
  //株価シートを取得
  let sheetStock = SpreadsheetApp.getActive().getSheetByName("株価");
  //表シート取得
  let sheet = SpreadsheetApp.getActive().getSheetByName("表");
  //表シート最終行の次の行を取得
  let row = sheet.getLastRow() + 1;

  //株価定期取得シートから表シートへデータをコピーする
  for (var key in eDataCell) {
    var eDataCell_val = eDataCell[key];
    var eColumn_val = eColumn[key];
    let value = sheetStock.getRange(eDataCell_val).getValue();
    sheet.getRange(row, eColumn_val).setValue(value);
  }

  //更新日時
  let date = Utilities.formatDate(
    new Date(),
    "Asia/Tokyo",
    "YYYY/MM/dd HH:mm:ss"
  );
  sheet.getRange(row, eColumn.DATE).setValue(date);
}
