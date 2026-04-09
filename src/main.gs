function fetchDailyStockPrices() {
  const now = new Date();
  const hour = now.getHours();
  const minute = now.getMinutes();

  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
  const lastReset = props.getProperty("lastResetDate");

  // 22:00〜22:29 の間に1回だけリセット
  if (hour === 22 && minute < 30) {
    if (lastReset !== today) {
      resetDailySheet();
      props.setProperty("lastResetDate", today);
    }
    return;
  }

  // 22:30〜翌5:30 の間だけ株価取得 
  const isNightTime = 
    (hour === 22 && minute >= 30) || 
    (hour >= 23) || 
    (hour < 5) || 
    (hour === 5 && minute < 30);

  if (!isNightTime) {
    return;
  }

  // 株価取得処理
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価");

  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
    sheet.getRange(1, 1, 1, 4).setValues([
      ["時刻", "AAPL", "MSFT", "GOOGL"]
    ]);
  }

  const symbols = ["AAPL", "MSFT", "GOOGL"];
  const row = [now];

  symbols.forEach(symbol => {
    let price = getPriceWithRetryYahoo(symbol);
    row.push(price);
  });

  sheet.appendRow(row);
}

function getPriceWithRetryYahoo(symbol) {
  let price = getPriceFromYahoo(symbol);
  if (price !== null) return price;

  Utilities.sleep(2000);
  return getPriceFromYahoo(symbol);
}

function getPriceFromYahoo(symbol) {
  // Yahoo APIの代わりに、一時的な計算用セルを使ってGOOGLEFINANCE関数から価格を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.getSheetByName("本日株価"); // 既存のシートを利用
  
  // セルに一時的に式を書き込んで値を取得
  const tempCell = tempSheet.getRange("Z1"); // 邪魔にならない遠いセル
  tempCell.setFormula(`=GOOGLEFINANCE("${symbol}", "price")`);
  
  // スプレッドシートの計算が終わるまで少し待機
  SpreadsheetApp.flush();
  const price = tempCell.getValue();
  
  // 使用したセルをクリア
  tempCell.clear();
  
  // 取得した値が数値であれば返し、エラー（#N/Aなど）ならnullを返す
  if (price && typeof price === 'number') {
    return price;
  } else {
    Logger.log("価格取得失敗: " + symbol);
    return null;
  }
}

function resetDailySheet() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価"); 

  // deleteSheet() を使わず軽量化
  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  // ヘッダー再設定
  sheet.getRange(1, 1, 1, 4).setValues([
    ["時刻", "AAPL", "MSFT", "GOOGL"]
  ]);
}
