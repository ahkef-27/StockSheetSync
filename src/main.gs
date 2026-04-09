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

  if (!isNightTime) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価");
  if (!sheet) {
    resetDailySheet();
    sheet = ss.getSheetByName("本日株価");
  }

  const symbols = ["AAPL", "MSFT", "GOOGL"];
  const row = [now];

  // 他の関数をいじらないよう、名前をそのまま維持して呼び出す
  symbols.forEach(symbol => {
    let price = getPriceWithRetry(symbol); // ★名前をスッキリ
    row.push(price);
  });

  sheet.appendRow(row);
}

/**
 * リトライ処理（名前からYahooを外しました）
 */
function getPriceWithRetry(symbol) {
  let price = getPriceFromGoogle(symbol); 
  if (price !== null) return price;

  Utilities.sleep(2000); // 2秒待機
  return getPriceFromGoogle(symbol);
}

/**
 * GoogleFinanceから価格取得（名前を実態に合わせました）
 */
function getPriceFromGoogle(symbol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("本日株価");
  const tempCell = sheet.getRange("Z1"); 
  
  tempCell.setFormula(`=GOOGLEFINANCE("${symbol}", "price")`);
  SpreadsheetApp.flush();
  
  const price = tempCell.getValue();
  tempCell.clearContent(); // clear()より少し軽量
  
  return (typeof price === 'number') ? price : null;
}

function resetDailySheet() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("本日株価"); 

  if (!sheet) {
    sheet = ss.insertSheet("本日株価");
  } else {
    sheet.clear(); 
  }

  sheet.getRange(1, 1, 1, 4).setValues([
    ["時刻", "AAPL", "MSFT", "GOOGL"]
  ]);
}
