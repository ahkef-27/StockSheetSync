function getPrice(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet();
  tempSheet.getRange("A1").setFormula(`=GOOGLEFINANCE("${ticker}", "price")`);

  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  const price = tempSheet.getRange("A1").getValue();
  ss.deleteSheet(tempSheet);

  if (typeof price !== "number") {
    throw new Error(`価格取得失敗: ${ticker}`);
  }

  return price;
}
