function getPrice(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calcSheet = ss.getSheetByName("CalcSheet");

  if (!calcSheet) {
    calcSheet = ss.insertSheet("CalcSheet");
    calcSheet.hideSheet();
  }

  calcSheet.getRange("A1").clearContent();

  calcSheet.getRange("A1").setFormula(`=GOOGLEFINANCE("${ticker}", "price")`);
  SpreadsheetApp.flush();

  let price;
  for (let i = 0; i < 10; i++) {
    Utilities.sleep(800);
    price = calcSheet.getRange("A1").getValue();
    if (typeof price === "number") break;
  }

  calcSheet.getRange("A1").clearContent();

  if (typeof price !== "number") {
    throw new Error(`価格取得失敗: ${ticker}`);
  }

  return price;
}

