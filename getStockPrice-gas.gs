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

// 日本株の株価の取得
function updateStockPrices(stockCode) {
  let stockPrice = getStockPriceGoogle(stockCode); //銘柄コードから株価取得
  return stockPrice;
}

// 投資信託(日本)の基準価格の取得
function updateToshinPrices(toshinCode) {
  let toshinPrice = getPriceToshin(toshinCode); //銘柄コードから投資信託の基準価格を取得
  return toshinPrice;
}

/**
 * 証券コードから価格を取得する。
 *
 * @param torihiki_code, shoken_code
 * @return 取得された価格です.
 * @customfunction
 *
 */
function STOCKPRICEJP(torihiki_code, shoken_code) {
  // shoken_code : 日本株(ETF含む):証券コード(4)、 投信:投信協会コード(8),isinコード(12)
  let param = torihiki_code;
  let error = "";
  try {
    if (
      param.length > 0 &&
      !(
        (param == "JP" && String(shoken_code).length == 4) ||
        (param == "TOSHIN" &&
          String(shoken_code).length == 12 &&
          String(shoken_code).startsWith("JP"))
      )
    ) {
      error = "取引、証券コードが正しくありません";
      throw error;
    }
  } catch {
    return error;
  }
  if ("JP" == param) {
    return updateStockPrices(shoken_code); // 日本株の価格を取得
  } else if ("TOSHIN" == param) {
    return updateToshinPrices(shoken_code); // 投信の価格を取得
  } else {
    return "取引コードが無し！";
  }
}

function updateStockPriceList() {
  //時価評価額セル
  const eDataCell = {
    GOOGLE: "I2",
    VOO: "I3",
    SP500_ETF: "I4",
    SOFTBANK: "I5",
    EMS_SP: "I6",
    EMS_ALL: "I7",
  };
  //表出力カラム
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
