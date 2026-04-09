function createDailyChart() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価");

  // シートがなければ作成（appendRow null 対策）
  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
    return; // 初回はデータがないので終了
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // データ不足なら終了

  // A列の幅を広げる（時刻が見やすくなる）
  sheet.setColumnWidth(1, 180);

  // A列〜D列のフォントサイズを12にする
  sheet.getRange(2, 1, lastRow, 4).setFontSize(12);

  // 1行目（ヘッダー）のフォントサイズも12にする 
  sheet.getRange(1, 1, 1, 4).setFontSize(12);

  const tickers = ["AAPL", "META", "GOOGL"];
  const titles = ["Apple", "Meta", "Google"];

  // 既存グラフ削除（安全）
  const charts = sheet.getCharts();
  charts.forEach(chart => {
    try { sheet.removeChart(chart); } catch (e) {}
  });

  // 削除直後に少し待つ（Google の内部処理待ち）
  Utilities.sleep(200);

  tickers.forEach((ticker, i) => {
    const col = i + 2;

    const timeRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const priceRange = sheet.getRange(2, col, lastRow - 1, 1);

    const prices = priceRange.getValues()
      .flat()
      .filter(v => typeof v === "number");

    if (prices.length === 0) return;

    const maxPrice = Math.max(...prices);
    const minPrice = Math.min(...prices);
    const pad = (maxPrice - minPrice) * 0.1;

    const chartBuilder = sheet.newChart();
    chartBuilder.addRange(timeRange);
    chartBuilder.addRange(priceRange);
    chartBuilder.setChartType(Charts.ChartType.LINE);
    chartBuilder.setOption("title", titles[i]);
    chartBuilder.setOption("legend", { position: "none" });
    chartBuilder.setOption("width", 800);
    chartBuilder.setOption("height", 140); // 見切れ対策
    chartBuilder.setOption("vAxis", {
      viewWindow: { min: minPrice - pad, max: maxPrice + pad }
    });

    // 行間隔を広げて見切れ防止
    chartBuilder.setPosition(1 + i * 7, 6, 0, 0);

    sheet.insertChart(chartBuilder.build());
  });
}
