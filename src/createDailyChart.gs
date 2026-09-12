function createDailyChart() {
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName("本日株価");

      if (!sheet) {
        sheet = ss.insertSheet("本日株価");
        return;
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 3) return; 

      sheet.setColumnWidth(1, 180);
      sheet.getDataRange().setFontSize(12);

      const charts = sheet.getCharts();
      charts.forEach(chart => {
        try { sheet.removeChart(chart); } catch (e) {}
      });

      const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      const timeRange = sheet.getRange(2, 1, lastRow - 1, 1);

      const tickers = ["AAPL", "META", "GOOGL"];
      const titles = ["Apple", "Meta", "Google"];

      tickers.forEach((ticker, i) => {
        const colIndex = i + 1;

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
      break;

    } catch (e) {
      Logger.log(`createDailyChart 試行 ${attempt}/${maxRetries} でエラー: ${e.message}`);
      if (attempt === maxRetries) throw e;
      Utilities.sleep(2000);
    }
  }
}
