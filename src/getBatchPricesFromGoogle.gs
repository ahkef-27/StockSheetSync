function getBatchPricesFromGoogle(sheet, symbols) {
  const startCol = 26; // Z列から横に展開
  const formulas = symbols.map(s => [`=GOOGLEFINANCE("${s}", "price")`]);
  
  const targetRange = sheet.getRange(1, startCol, 1, symbols.length);
  targetRange.setFormulas([formulas.map(f => f[0])]);
  
  SpreadsheetApp.flush();
  
  // 待機時間を最大10秒まで拡張（GOOGLEFINANCEの遅延対策）
  let values = [];
  for (let i = 0; i < 10; i++) {
    values = targetRange.getValues()[0];
    const isAllReady = values.every(v => typeof v === 'number' && !isNaN(v));
    if (isAllReady) break;
    Utilities.sleep(1000);
  }

  // 作業用セルのクリア
  targetRange.clearContent();

  // 数値化できていないものは null に置換
  return values.map(v => (typeof v === 'number' && !isNaN(v) ? v : null));
}
