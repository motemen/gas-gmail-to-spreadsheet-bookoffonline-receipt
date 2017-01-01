function main() {
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty("SpreadsheetId");
    var scriptId = ScriptApp.getScriptId();
    Logger.log(scriptId);
    var labelNameProcessed = "gas-" + scriptId + "-processed";
    var labelProcessed = GmailApp.getUserLabelByName(labelNameProcessed);
    if (!labelProcessed) {
        labelProcessed = GmailApp.createLabel(labelNameProcessed);
    }
    var threads = GmailApp.search("from:info@bookoffonline.jp \u3054\u6CE8\u6587\u3042\u308A\u304C\u3068\u3046\u3054\u3056\u3044\u307E\u3059 -label:" + labelNameProcessed).reverse();
    for (var _i = 0, threads_1 = threads; _i < threads_1.length; _i++) {
        var thread = threads_1[_i];
        var message = thread.getMessages()[0];
        var orders = extractOrders(message);
        orders.forEach(function (order) { return addToSpreadsheet(spreadsheetId, order); });
        labelProcessed.addToThread(thread);
    }
}
function extractOrders(message) {
    var date = message.getDate();
    var messageBody = message.getPlainBody();
    var orderId = /【ご注文番号：(\d+)】/.exec(messageBody)[1];
    var itemsText = /ご注文商品([^]+?)お届け先名：/.exec(messageBody)[1];
    var re = /(.+?) \(￥([0-9,]+)\)　ご注文点数：(\d+)点/g;
    var m;
    var orders = [];
    while ((m = re.exec(itemsText))) {
        var _ = m[0], title = m[1], priceText = m[2], count = m[3];
        title = title.replace(/^【中古】/, '');
        var price = parseInt(priceText.replace(/,/, ''));
        var order = {
            date: date,
            title: title,
            price: price,
            url: "https://www.bookoffonline.co.jp/member/CPmOrderHistoryDetail.jsp?ordno=" + orderId,
            count: parseInt(count)
        };
        orders.push(order);
    }
    return orders;
}
function addToSpreadsheet(spreadsheetId, order) {
    var sheet = SpreadsheetApp.openById(spreadsheetId);
    sheet.appendRow([order.date, order.title, order.price, order.count, order.url]);
}
