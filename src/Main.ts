interface Order {
  date: Date;
  title: string;
  price: number;
  count: number;
  url: string;
}

function main() {
  let spreadsheetId = PropertiesService.getScriptProperties().getProperty("SpreadsheetId");
  let scriptId = (<any>ScriptApp).getScriptId();
  Logger.log(scriptId);

  let labelNameProcessed = `gas-${scriptId}-processed`
  let labelProcessed = GmailApp.getUserLabelByName(labelNameProcessed);
  if (!labelProcessed) {
    labelProcessed = GmailApp.createLabel(labelNameProcessed);
  }

  let threads = GmailApp.search(`from:info@bookoffonline.jp ご注文ありがとうございます -label:${labelNameProcessed}`).reverse();
  for (let thread of threads) {
    let message = thread.getMessages()[0];
    let orders = extractOrders(message);
    orders.forEach((order) => addToSpreadsheet(spreadsheetId, order));
    labelProcessed.addToThread(thread);
  }
}

function extractOrders(message: GoogleAppsScript.Gmail.GmailMessage): Order[] {
  let date = message.getDate();
  let messageBody = message.getPlainBody();
  let orderId = /【ご注文番号：(\d+)】/.exec(messageBody)[1];
  let itemsText = /ご注文商品([^]+?)お届け先名：/.exec(messageBody)[1];
  let re = /(.+?) \(￥([0-9,]+)\)　ご注文点数：(\d+)点/g;
  let m;
  let orders: Order[] = [];
  while ((m = re.exec(itemsText))) {
    let [_, title, priceText, count] = m;
    title = title.replace(/^【中古】/, '');
    let price = parseInt(priceText.replace(/,/, ''));
    let order = {
      date,
      title,
      price,
      url: `https://www.bookoffonline.co.jp/member/CPmOrderHistoryDetail.jsp?ordno=${orderId}`,
      count: parseInt(count)
    };
    orders.push(order);
  }
  return orders;
}

function addToSpreadsheet(spreadsheetId: string, order: Order) {
  let sheet = SpreadsheetApp.openById(spreadsheetId);
  sheet.appendRow([order.date, order.title, order.price, order.count, order.url]);
}
