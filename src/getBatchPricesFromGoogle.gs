function getBatchPricesFromGoogle(sheet, symbols) {
  const startCol = 26; // Z列から横に展開
  const formulas = symbols.map(s => [`=GOOGLEFINANCE("${s}", "price")`]);
  
  // Z1:AB1 に一括で数式を設定
  const targetRange = sheet.getRange(1, startCol, 1, symbols.length);
  targetRange.setFormulas([formulas.map(f => f[0])]);
  
  SpreadsheetApp.flush();
  
  // 最大5秒だけ値の反映を確認
  let values = [];
  for (let i = 0; i < 5; i++) {
    values = targetRange.getValues()[0];
    const isAllReady = values.every(v => typeof v === 'number');
    if (isAllReady) break;
    Utilities.sleep(1000);
  }

  // 作業用セルのクリア
  targetRange.clearContent();

  // 数値化できていないものは null に置換
  return values.map(v => (typeof v === 'number' ? v : null));
}
