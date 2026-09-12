function fetchDailyStockPrices() {
  const now = new Date();
  const hour = now.getHours();
  const minute = now.getMinutes();

  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
  const lastReset = props.getProperty("lastResetDate");

  // 22:00?22:29 の間に1回だけリセット
  if (hour === 22 && minute < 30) {
    if (lastReset !== today) {
      resetDailySheet();
      props.setProperty("lastResetDate", today);
    }
    return;
  }

  // 22:30?翌5:30 の間だけ株価取得 
  const isNightTime = 
    (hour === 22 && minute >= 30) || 
    (hour >= 23) || 
    (hour < 5) || 
    (hour === 5 && minute < 30);

  if (!isNightTime) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価");
  
  if (!sheet) {
    resetDailySheet();
    sheet = ss.getSheetByName("本日株価");
  }

  const symbols = ["AAPL", "MSFT", "GOOGL"];

  // 全銘柄をまとめて一括取得（数秒で完了）
  const prices = getBatchPricesFromGoogle(sheet, symbols);
  
  const row = [now, ...prices];
  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  colorizePricesHorizontal(sheet, lastRow);
  createDailyChart();
  styleHeaderRows();
}
