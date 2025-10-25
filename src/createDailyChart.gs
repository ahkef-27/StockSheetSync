function createDailyChart(sheet) {
  const lastRow = sheet.getLastRow();
  const tickers = ["AAPL", "META", "GOOGL"];
  const titles = ["Apple", "Meta", "Google"];
  sheet.getRange(2, 1, lastRow - 1).setNumberFormat("hh:mm");

  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));

  tickers.forEach((ticker, i) => {
    const col = i + 2;
    const timeRange = sheet.getRange(2, 1, lastRow - 1, 1); // A列（時刻）
    const priceRange = sheet.getRange(2, col, lastRow - 1, 1); // 株価列
    // 株価データを取得して最大・最小を計算
    const prices = priceRange.getValues().flat().filter(v => typeof v === "number");
    const maxPrice = Math.max(...prices);
    const minPrice = Math.min(...prices);

    // 少し余白を持たせる
    const rangePadding = (maxPrice - minPrice) * 0.1;
    const vMin = Math.floor(minPrice - rangePadding);
    const vMax = Math.ceil(maxPrice + rangePadding);

    const chartBuilder = sheet.newChart();
    chartBuilder.addRange(timeRange);
    chartBuilder.addRange(priceRange);
    chartBuilder.setChartType(Charts.ChartType.LINE);
    chartBuilder.setOption("title", titles[i]);
    chartBuilder.setOption("titleTextStyle", { fontSize: 12 }); // ← 追加
    chartBuilder.setOption("legend", { position: "none" });
    chartBuilder.setOption("width", 800);
    chartBuilder.setOption("height", 170);
    chartBuilder.setOption("hAxis", {
      format: "HH:mm:ss",
      textStyle: { fontSize: 10 },
      slantedText: true,
      slantedTextAngle: 45
    });

    chartBuilder.setOption("vAxis", {
      viewWindow: {
        min: vMin,
        max: vMax
    },
    textStyle: { fontSize: 13 }
  });

    const dataLength = prices.length;
    const rowOffset = Math.max(dataLength + 2, 8); // 最低8行確保
    chartBuilder.setPosition(1 + i * rowOffset, 6, 0, 0)

    sheet.insertChart(chartBuilder.build());
  });
}
