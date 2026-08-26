function createDailyChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("本日株価");

  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; 

  // A列の幅調整
  sheet.setColumnWidth(1, 180);

  // データが存在するすべてのセルのフォントサイズを一瞬で12に変更
  sheet.getDataRange().setFontSize(12);

  // 既存グラフ削除
  const charts = sheet.getCharts();
  charts.forEach(chart => {
    try { sheet.removeChart(chart); } catch (e) {}
  });

  // 全データを1回で一括取得
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const timeRange = sheet.getRange(2, 1, lastRow - 1, 1);

  const tickers = ["AAPL", "META", "GOOGL"];
  const titles = ["Apple", "Meta", "Google"];

  tickers.forEach((ticker, i) => {
    const colIndex = i + 1;

    // メモリ内で価格データを抽出・計算
    const prices = data
      .map(row => row[colIndex])
      .filter(v => typeof v === "number" && !isNaN(v));

    if (prices.length === 0) return;

    const maxPrice = Math.max(...prices);
    const minPrice = Math.min(...prices);
    const pad = (maxPrice === minPrice) ? 1 : (maxPrice - minPrice) * 0.1;

    const priceRange = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);

    const chartBuilder = sheet.newChart()
      .addRange(timeRange)
      .addRange(priceRange)
      .setChartType(Charts.ChartType.LINE)
      .setOption("title", titles[i])
      .setOption("legend", { position: "none" })
      .setOption("width", 800)
      .setOption("height", 140)
      .setOption("vAxis", {
        viewWindow: { min: minPrice - pad, max: maxPrice + pad }
      })
      .setPosition(1 + i * 7, 6, 0, 0);

    sheet.insertChart(chartBuilder.build());
  });
}
