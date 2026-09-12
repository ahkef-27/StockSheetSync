function styleHeaderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ["本日株価", "株価グラフ"];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), 1);

    sheet.getRange(1, 1, lastRow, lastCol).setFontSize(13);
    sheet.getRange(1, 1, 1, lastCol)
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
  });
}
