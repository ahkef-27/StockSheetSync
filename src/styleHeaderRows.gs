function styleHeaderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ["本日株価", "株価グラフ"];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    // ✅ 最低でも1行1列を確保
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), 1);

    // ✅ フォントサイズを13に統一（空でも安全）
    sheet.getRange(1, 1, lastRow, lastCol).setFontSize(13);

    // ✅ 1行目のスタイル（列がなくても1列分確保）
    sheet.getRange(1, 1, 1, lastCol)
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
  });
}
