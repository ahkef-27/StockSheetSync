function resetDailySheet() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価"); 

  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
  } else {
    sheet.clear(); 
  }

  sheet.getRange(1, 1, 1, 4).setValues([
    ["時刻", "AAPL", "MSFT", "GOOGL"]
  ]);
}
