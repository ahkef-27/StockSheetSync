function createSummaryCharts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("ロック取得できず → 他の実行が動いているのでスキップ");
    return;
  }

  try {
    const maxRetries = 3;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const summarySheetName = "株価グラフ";
        const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
        const props = PropertiesService.getScriptProperties();
        const lastCreatedDate = props.getProperty("lastChartDate");

        if (lastCreatedDate === today) {
          Logger.log("本日のグラフはすでに作成済みです");
          return;
        }

        let summarySheet = ss.getSheetByName(summarySheetName);
        if (!summarySheet) {
          summarySheet = ss.insertSheet(summarySheetName);
        } else {
          summarySheet.getCharts().forEach(chart => {
            summarySheet.removeChart(chart);
          });
          summarySheet.clear();
        }

        summarySheet.getRange(1, 1, 40, 10).setFontSize(13);
        for (let r = 1; r <= 40; r++) summarySheet.setRowHeight(r, 21);

        summarySheet.getRange(1, 1, 1, 10)
          .setFontColor("black")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");

        const tickers = ["AAPL", "META", "GOOGL"];
        const now = new Date();
        const sevenDaysAgo = new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000);
        const startDate = Utilities.formatDate(sevenDaysAgo, "Asia/Tokyo", "yyyy-MM-dd");
        const endDate = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");

        let tempSheet = ss.getSheetByName("tempData");
        if (tempSheet) {
          ss.deleteSheet(tempSheet);
        }
        tempSheet = ss.insertSheet("tempData");

        let dates = [];
        let priceMap = {};
        let basePrices = {};

        tickers.forEach((ticker) => {
          tempSheet.clear();

          tempSheet.getRange(1, 1).setFormula(
            `=GOOGLEFINANCE("${ticker}", "price", DATEVALUE("${startDate}"), DATEVALUE("${endDate}"))`
          );
          SpreadsheetApp.flush();
          Utilities.sleep(1000);

          const data = tempSheet.getDataRange().getValues();
          if (data.length < 2) return;

          const rows = data.slice(1);

          if (dates.length === 0) {
            dates = rows.map(row => [row[0]]);
            summarySheet.getRange(1, 1).setValue("日付");
            summarySheet.getRange(2, 1, dates.length).setValues(dates);
            summarySheet.getRange(2, 1, dates.length).setNumberFormat("yyyy-MM-dd");
          }

          const prices = rows.map(row => [row[1]]);
          priceMap[ticker] = prices;
          basePrices[ticker] = prices[0][0];
        });

        try { ss.deleteSheet(tempSheet); } catch (e) {}

        tickers.forEach((ticker, i) => {
          const col = i + 2;
          const prices = priceMap[ticker];
          const base = basePrices[ticker];

          if (!prices || prices.length === 0) return;

          summarySheet.getRange(1, col).setValue(ticker);
          summarySheet.getRange(2, col, prices.length).setValues(prices);

          const bgColors = prices.map((p, idx) => {
            const value = p[0];
            if (typeof value !== "number") return ["#eeeeee"];
            if (idx === 0) return ["#ffffff"];

            const changeRate = (value - base) / base;
            if (changeRate >= 0.10) return ["#bbdefb"];
            if (changeRate >= 0.05) return ["#c8e6c9"];
            if (changeRate <= -0.10) return ["#ffcdd2"];
            if (changeRate <= -0.05) return ["#fff9c4"];
            return ["#ffffff"];
          });

          summarySheet.getRange(2, col, prices.length).setBackgrounds(bgColors);

          const chartBuilder = summarySheet.newChart();
          chartBuilder.addRange(summarySheet.getRange(1, 1, prices.length + 1, 1));
          chartBuilder.addRange(summarySheet.getRange(1, col, prices.length + 1, 1));
          chartBuilder.setChartType(Charts.ChartType.LINE);
          chartBuilder.setOption("title", `${ticker} 過去7日間の株価`);
          chartBuilder.setOption("legend", { position: "none" });
          chartBuilder.setOption("width", 800);
          chartBuilder.setOption("height", 140);
          chartBuilder.setOption("hAxis", {
            format: "MM/dd",
            textStyle: { fontSize: 10 }
          });
          chartBuilder.setOption("vAxis", {
            viewWindow: {
              min: base * 0.9,
              max: base * 1.1
            },
            textStyle: { fontSize: 10 }
          });

          chartBuilder.setPosition(1 + i * 8, 6, 0, 0);
          summarySheet.insertChart(chartBuilder.build());
        });

        props.setProperty("lastChartDate", today);
        break;

      } catch (e) {
        const oldTemp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tempData");
        if (oldTemp) {
          try { SpreadsheetApp.getActiveSpreadsheet().deleteSheet(oldTemp); } catch (err) {}
        }

        Logger.log(`createSummaryCharts 試行 ${attempt}/${maxRetries} でエラー: ${e.message}`);
        if (attempt === maxRetries) throw e;
        Utilities.sleep(2000);
      }
    }
  } finally {
    lock.releaseLock();
  }
}
