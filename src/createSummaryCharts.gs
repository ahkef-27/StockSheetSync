function createSummaryCharts() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet("tempData");
  const summarySheetName = "株価グラフ";
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const props = PropertiesService.getScriptProperties();
  const lastCreatedDate = props.getProperty("lastChartDate");

  // ✅ 不要な初期シート「シート1」を削除
  const defaultSheet = ss.getSheetByName("シート1");
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  // ✅ 本日分のグラフがすでに作成されていたらスキップ
  if (lastCreatedDate === today) {
    Logger.log("本日のグラフはすでに作成済みです");
    ss.deleteSheet(tempSheet);
    return;
  }

  let summarySheet = ss.getSheetByName(summarySheetName);

  if (summarySheet) {
    const lastRow = summarySheet.getLastRow();
    let dateValues = [];

    if (lastRow > 1) {
      dateValues = summarySheet.getRange(2, 1, lastRow - 1).getValues().flat();
    }

    const todayExists = dateValues.some(date => {
      const formatted = Utilities.formatDate(new Date(date), "Asia/Tokyo", "yyyy/MM/dd");
      return formatted === today;
    });

    if (todayExists) {
      Logger.log("本日のグラフはすでに存在しています");
      ss.deleteSheet(tempSheet);
      return;
    }

    ss.deleteSheet(summarySheet);
  }

  summarySheet = ss.insertSheet(summarySheetName);

  // ✅ グラフ描画前に全セルのフォントサイズを13に統一（仮に1000行×10列）
  summarySheet.getRange(1, 1, 1000, 10).setFontSize(13);

  // ✅ 全行の高さを固定（21px）
  for (let r = 1; r <= 1000; r++) {
    summarySheet.setRowHeight(r, 21);
  }

  // ✅ 1行目のスタイル
  summarySheet.getRange(1, 1, 1, 10)
              .setFontColor("black")
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");

  const tickers = ["AAPL", "META", "GOOGL"];
  const now = new Date();
  const sevenDaysAgo = new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000);
  const startDate = Utilities.formatDate(sevenDaysAgo, "Asia/Tokyo", "yyyy/MM/dd");
  const endDate = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");

  let dates = [];
  let priceMap = {};
  let basePrices = {};

  tickers.forEach((ticker) => {
    const cell = tempSheet.getRange(1, 1);
    cell.setFormula(`=GOOGLEFINANCE("${ticker}", "price", DATEVALUE("${startDate}"), DATEVALUE("${endDate}"))`);
    SpreadsheetApp.flush();
    Utilities.sleep(3000);

    const data = tempSheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log(`${ticker} の履歴データが取得できませんでした`);
      return;
    }

    const rows = data.slice(1);
    if (dates.length === 0) {
      dates = rows.map(row => [row[0]]);
      summarySheet.getRange(1, 1).setValue("日付");
      summarySheet.getRange(2, 1, dates.length).setValues(dates);
      summarySheet.getRange(2, 1, dates.length).setNumberFormat("yyyy/MM/dd");
    }

    const prices = rows.map(row => [row[1]]);
    if (prices.length === 0) return;

    priceMap[ticker] = prices;
    basePrices[ticker] = prices[0][0];
    tempSheet.clear();
  });

  tickers.forEach((ticker, i) => {
    const col = i + 2;
    const prices = priceMap[ticker];
    const base = basePrices[ticker];

    if (!prices || prices.length === 0) return;

    summarySheet.getRange(1, col).setValue(ticker);
    summarySheet.getRange(2, col, prices.length).setValues(prices);

    const range = summarySheet.getRange(2, col, prices.length);
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

    range.setBackgrounds(bgColors);

    const chartBuilder = summarySheet.newChart();
    chartBuilder.addRange(summarySheet.getRange(1, 1, prices.length + 1, 1));
    chartBuilder.addRange(summarySheet.getRange(1, col, prices.length + 1, 1));
    chartBuilder.setChartType(Charts.ChartType.LINE);
    chartBuilder.setOption("title", `${ticker} 過去7日間の株価`);
    chartBuilder.setOption("titleTextStyle", { fontSize: 13 });
    chartBuilder.setOption("legend", { position: "none" });
    chartBuilder.setOption("hAxis", {
      format: "MM/dd",
      textPosition: "out",
      textStyle: { fontSize: 10 }
    });
    chartBuilder.setOption("vAxis", {
      viewWindow: {
        min: base * 0.9,
        max: base * 1.1
      },
      textStyle: { fontSize: 10 }
    });
    chartBuilder.setOption("width", 800);
    chartBuilder.setOption("height", 160);
    chartBuilder.setPosition(1 + i * 8, 6, 0, 0);

    summarySheet.insertChart(chartBuilder.build());
  });

  // ✅ フォントサイズを13に統一（1000行×10列）
  summarySheet.getRange(1, 1, 1000, 10).setFontSize(13);
  
  const targetSheet = ss.getSheetByName("本日株価");
  if (targetSheet) {
    const maxRows = targetSheet.getMaxRows();
    for (let r = 1; r <= maxRows; r++) {
      targetSheet.setRowHeight(r, 21);
    }
  }

  ss.deleteSheet(tempSheet);
  props.setProperty("lastChartDate", today);
}
