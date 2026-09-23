/**
 * スプレッドシート → Shopify商品CSV の最小サンプル（Apps Script）
 *
 * 使い方:
 * 1. シート1行目にヘッダー: handle, title, body_html, vendor, type, tags, price, sku
 * 2. 2行目以降に商品を書く
 * 3. メニュー「Shopify CSV」→「書き出し」で CSV を生成してダウンロード
 *
 * 注意: 本番の一括取込前に、必ず少数行で Admin のインポートを試すこと。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Shopify CSV")
    .addItem("書き出し", "exportShopifyProductsCsv")
    .addToUi();
}

function exportShopifyProductsCsv() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("ヘッダーと1行以上のデータが必要です。");
    return;
  }

  const header = values[0].map(function (h) {
    return String(h).trim().toLowerCase();
  });
  const need = ["handle", "title", "price"];
  for (var i = 0; i < need.length; i++) {
    if (header.indexOf(need[i]) === -1) {
      SpreadsheetApp.getUi().alert("必須列がありません: " + need[i]);
      return;
    }
  }

  // Shopify product CSV の代表列（最小セット）
  const outHeader = [
    "Handle",
    "Title",
    "Body (HTML)",
    "Vendor",
    "Type",
    "Tags",
    "Published",
    "Variant SKU",
    "Variant Price",
    "Variant Inventory Tracker",
    "Variant Inventory Qty",
    "Variant Inventory Policy",
    "Variant Fulfillment Service",
    "Variant Requires Shipping",
    "Variant Taxable"
  ];

  const rows = [outHeader];
  for (var r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row.join("").trim()) continue;
    const get = function (name) {
      const idx = header.indexOf(name);
      return idx === -1 ? "" : String(row[idx] || "").trim();
    };
    const handle = get("handle") || slugify(get("title"));
    rows.push([
      handle,
      get("title"),
      get("body_html"),
      get("vendor"),
      get("type"),
      get("tags"),
      "TRUE",
      get("sku"),
      get("price"),
      "shopify",
      "0",
      "deny",
      "manual",
      "TRUE",
      "TRUE"
    ]);
  }

  const csv = rows
    .map(function (cols) {
      return cols
        .map(function (c) {
          const s = String(c).replace(/"/g, '""');
          return /[",\n]/.test(s) ? '"' + s + '"' : s;
        })
        .join(",");
    })
    .join("\n");

  const html =
    "<a download=\"shopify-products.csv\" href=\"data:text/csv;charset=utf-8," +
    encodeURIComponent(csv) +
    "\">CSVをダウンロード</a>";
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(320).setHeight(80),
    "Shopify CSV"
  );
}

function slugify(title) {
  return String(title || "")
    .toLowerCase()
    .replace(/[^a-z0-9\u3040-\u30ff\u4e00-\u9faf]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}
