function getBatchPricesFromGoogle(sheet, symbols) {
  const startCol = 26; // Z列から横に展開
  const formulas = symbols.map(s => [`=GOOGLEFINANCE("${s}", "price")`]);
  
  const targetRange = sheet.getRange(1, startCol, 1, symbols.length);
  targetRange.setFormulas([formulas.map(f => f[0])]);
  
  SpreadsheetApp.flush();
  
  // 最初は少し長め（2秒）待ってからチェック（サーバー負荷軽減のため間隔を2秒に変更）
  Utilities.sleep(2000);

  let values = [];
  for (let i = 0; i < 5; i++) {
    values = targetRange.getValues()[0];
    const isAllReady = values.every(v => typeof v === 'number' && !isNaN(v));
    if (isAllReady) break;
    Utilities.sleep(2000); // 2秒待機
  }

  // 作業用セルのクリア
  targetRange.clearContent();

  // 数値化できていないものは null に置換
  return values.map(v => (typeof v === 'number' && !isNaN(v) ? v : null));
}
