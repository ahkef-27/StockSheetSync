function styleHeaderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ["本日株価", "株価グラフ"];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    // データが存在する列数のみ（ヘッダー行のみを装飾対象にする）
    const lastCol = Math.max(sheet.getLastColumn(), 1);

    // ★ 1行目（ヘッダー）だけに範囲を絞り込んで処理
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontSize(13);
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
  });
}
