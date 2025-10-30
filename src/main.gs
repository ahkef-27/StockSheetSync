function fetchDailyStockPrices() {
  const now = new Date();
  const hourJST = Utilities.formatDate(now, "Asia/Tokyo", "HH");

  //米国市場の時間帯（日本時間で22〜5時）
  const allowedHours = ["22","23", "00", "01", "02", "03", "04", "05"];
  if (!allowedHours.includes(hourJST)) {
  Logger.log("米国市場の時間外なのでスキップします");
    return;
  }

  const today = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価");

  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
    sheet.getRange(1, 1).setValue("取得時刻");
    sheet.getRange(1, 2, 1, 3).setValues([["AAPL", "META", "GOOGL"]]);
  }

  const lastDate = sheet.getRange(2, 1).getValue();
  const lastDateStr = Utilities.formatDate(new Date(lastDate), "Asia/Tokyo", "yyyy/MM/dd");

  // 日付が変わっていたら前日のデータを削除（行数チェック付き）
  const dataRowCount = sheet.getLastRow() - 1;
  if (lastDateStr !== today && lastDateStr !== "" && dataRowCount > 0) {
    sheet.deleteRows(2, dataRowCount);
  }

  const tickers = ["AAPL", "META", "GOOGL"];
  const row = sheet.getLastRow() + 1;

  // ✅ A列に「Date型」で現在時刻を記録！
  sheet.getRange(row, 1).setValue(now);

  tickers.forEach((ticker, i) => {
    let price;
    try {
      price = getPrice(ticker);
    } catch (e) {
      price = "取得失敗";
    }
    sheet.getRange(row, i + 2).setValue(price);
  });

  // ✅ 表示形式を「hh:mm」に設定（見た目だけ）
  sheet.getRange(2, 1, sheet.getLastRow() - 1).setNumberFormat("hh:mm");

  styleHeaderRows();
  insertTodayDateToAllSheets();
  colorizePricesHorizontal(sheet, row);
  createDailyChart(sheet);
  createSummaryCharts();
}
