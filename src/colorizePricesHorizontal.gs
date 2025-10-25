function colorizePricesHorizontal(sheet, row) {
  const tickers = ["AAPL", "META", "GOOGL"];
  tickers.forEach((ticker, i) => {
    const col = i + 2;
    const cell = sheet.getRange(row, col);
    const value = cell.getValue();

    let color = "#ffffff"; // デフォルトは白

    if (typeof value === "number") {
      const prevCell = sheet.getRange(row - 1, col);
      const prevValue = prevCell.getValue();

      if (typeof prevValue === "number") {
        const changeRate = (value - prevValue) / prevValue;

        if (changeRate >= 0.10) color = "#bbdefb"; // +10%以上 → 青
        else if (changeRate >= 0.05) color = "#c8e6c9"; // +5%以上 → 緑
        else if (changeRate <= -0.10) color = "#ffcdd2"; // -10%以上 → 赤
        else if (changeRate <= -0.05) color = "#fff9c4"; // -5%以上 → 黄
      }
    } else {
      color = "#eeeeee"; // 取得失敗
    }

    cell.setBackground(color);
  });
}
