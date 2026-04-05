function colorizePricesHorizontal(sheet, row) {
  const tickers = ["AAPL", "META", "GOOGL"];
  tickers.forEach((ticker, i) => {
    const col = i + 2;
    const cell = sheet.getRange(row, col);
    const value = cell.getValue();

    let color = "#ffffff";

    if (typeof value === "number") {
      const prevCell = sheet.getRange(row - 1, col);
      const prevValue = prevCell.getValue();

      if (typeof prevValue === "number") {
        const changeRate = (value - prevValue) / prevValue;

        if (changeRate >= 0.10) color = "#bbdefb";
        else if (changeRate >= 0.05) color = "#c8e6c9";
        else if (changeRate <= -0.10) color = "#ffcdd2";
        else if (changeRate <= -0.05) color = "#fff9c4";
      }
    } else {
      color = "#eeeeee";
    }

    cell.setBackground(color);
  });
}
