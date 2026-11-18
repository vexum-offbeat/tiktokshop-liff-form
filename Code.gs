/**
 * TikToker商品サンプル発送依頼 & リマインド自動化システム
 * Google Apps Script (GAS) バックエンド
 * v2.0 - 承認制ワークフロー対応
 */

// ====================================================================
// スクリプトプロパティ一覧（GASの「プロジェクトの設定」から確認・編集できます）
//
// プロパティ名              説明                           デフォルト値
// SPREADSHEET_ID        スプレッドシートのID             （必須）
// LINE_CHANNEL_TOKEN    LINEチャネルアクセストークン      （必須）
// LIFF_ID               LINE DevelopersのLIFF ID         （必須）
// AFFILIATE_URL         アフィリエイトリンク              （任意）
// SHEET_NAME_PENDING    承認待ちシートのタブ名            承認待ちシート
// SHEET_NAME_CREATOR_DB クリエイターDBのタブ名            クリエイターDB
// SHEET_NAME_SHIPPING   配送シートのタブ名               配送シート
// SHEET_NAME_ADDRESS    住所一覧シートのタブ名            クリエイター・代理店サンプルリスト
// SHEET_NAME_SHOP_MASTER 店舗マスタのタブ名              店舗発送依頼先情報
// APPROVAL_COL          承認チェックボックスの列番号      13（M列）
// ====================================================================

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID:            props.getProperty("SPREADSHEET_ID")         || "10v-c76tg20i5sSzIM6T7-QQ7kTNLPiNOZhvF0A8NDVg",
    LINE_CHANNEL_ACCESS_TOKEN: props.getProperty("LINE_CHANNEL_TOKEN")     || "",
    SHEET_NAME_PENDING:        props.getProperty("SHEET_NAME_PENDING")     || "\u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8",
    SHEET_NAME_CREATOR_DB:     props.getProperty("SHEET_NAME_CREATOR_DB")  || "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB",
    SHEET_NAME_SHIPPING:       props.getProperty("SHEET_NAME_SHIPPING")    || "\u914d\u9001\u30b7\u30fc\u30c8",
    SHEET_NAME_ADDRESS:        props.getProperty("SHEET_NAME_ADDRESS")     || "\u30af\u30ea\u30a8\u30a4\u30bf\u30fc\u30fb\u4ee3\u7406\u5e97\u30b5\u30f3\u30d7\u30eb\u30ea\u30b9\u30c8",
    SHEET_NAME_SHOP_MASTER:    props.getProperty("SHEET_NAME_SHOP_MASTER") || "\u5e97\u8217\u767a\u9001\u4f9d\u983c\u5148\u60c5\u5831",
    AFFILIATE_INFO_TEXT:       props.getProperty("AFFILIATE_URL")          || "",
    APPROVAL_COL:              Number(props.getProperty("APPROVAL_COL"))   || 13
  };
}

/**
 * スクリプトプロパティに初期値を一括登録する
 * GASエディタで「initializeScriptProperties」を選んで1回実行するだけでOK
 * すでに設定済みのプロパティは上書きしません
 */
function initializeScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperties();

  const defaults = {
    "SPREADSHEET_ID":        "10v-c76tg20i5sSzIM6T7-QQ7kTNLPiNOZhvF0A8NDVg",
    "LINE_CHANNEL_TOKEN":    "",
    "LIFF_ID":               "",
    "AFFILIATE_URL":         "",
    "SHEET_NAME_PENDING":    "\u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8",
    "SHEET_NAME_CREATOR_DB": "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB",
    "SHEET_NAME_SHIPPING":   "\u914d\u9001\u30b7\u30fc\u30c8",
    "SHEET_NAME_ADDRESS":    "\u30af\u30ea\u30a8\u30a4\u30bf\u30fc\u30fb\u4ee3\u7406\u5e97\u30b5\u30f3\u30d7\u30eb\u30ea\u30b9\u30c8",
    "SHEET_NAME_SHOP_MASTER":"\u5e97\u8217\u767a\u9001\u4f9d\u983c\u5148\u60c5\u5831",
    "APPROVAL_COL":          "13"
  };

  const toSet = {};
  const skipped = [];
  const added = [];

  Object.keys(defaults).forEach(key => {
    if (current[key] !== undefined && current[key] !== "") {
      skipped.push(key + " = " + current[key]);
    } else {
      toSet[key] = defaults[key];
      added.push(key + " = " + defaults[key]);
    }
  });

  if (Object.keys(toSet).length > 0) {
    props.setProperties(toSet);
  }

  Logger.log("=== initializeScriptProperties ===");
  Logger.log("[登録済み・スキップ]\n" + (skipped.length ? skipped.join("\n") : "なし"));
  Logger.log("[新規追加]\n" + (added.length ? added.join("\n") : "なし"));
  Logger.log("完了！GASの「プロジェクトの設定 > スクリプトプロパティ」で内容を確認・編集できます。");
}

function getSpreadsheet() {
  const cfg = getConfig();
  if (cfg.SPREADSHEET_ID && cfg.SPREADSHEET_ID.trim()) {
    try {
      return SpreadsheetApp.openById(cfg.SPREADSHEET_ID.trim());
    } catch (e) {
      Logger.log("openById failed, fallback: " + e.toString());
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// =============================================
// LIFF\u30d5\u30a9\u30fc\u30e0\u8868\u793a\u30fb\u9001\u4fe1\u51e6\u7406
// =============================================

function doGet(e) {
  const liffId = PropertiesService.getScriptProperties().getProperty("LIFF_ID") || "";
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setTitle("\u5546\u54c1\u30b5\u30f3\u30d7\u30eb\u767a\u9001\u30fb\u7d39\u4ecb\u53d7\u4ed8\u30d5\u30a9\u30fc\u30e0")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.setContent(html.getContent().replace("__LIFF_ID__", liffId));
  return html;
}

function submitShippingRequest(data) {
  return processRequestData(data);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = processRequestData(data);
    return createJsonResponse(result);
  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

// =============================================
// \u30d5\u30a9\u30fc\u30e0\u9001\u4fe1\u51e6\u7406\uff08\u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB + \u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8\u306b\u66f8\u304f\u3060\u3051\u3067\u6b62\u307e\u308b\uff09
// =============================================

function processRequestData(data) {
  try {
    const ss = getSpreadsheet();

    const shopInfo = resolveShopInfo(ss, data.groupId);
    data.shopName          = shopInfo.shopName;
    data.shopNotifyGroupId = shopInfo.notifyGroupId;
    data.shopSheetUrl      = shopInfo.shopSheetUrl;

    const requestId = "REQ-" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd-HHmmss")
                      + "-" + Math.floor(Math.random() * 1000);
    data.requestId = requestId;

    // \u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB\u306b\u6c38\u4e45\u4fdd\u5b58\uff08\u30d0\u30c3\u30af\u30a2\u30c3\u30d7\uff09
    appendCreatorDB(ss, data);

    // \u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8\u306b\u66f8\u304d\u8fbc\u3093\u3067\u6b62\u307e\u308b
    appendPendingRecords(ss, data);

    return {
      success: true,
      shopName: data.shopName,
      action: "pending_approval",
      message: "\u7533\u8acb\u3092\u53d7\u3051\u4ed8\u3051\u307e\u3057\u305f\uff01\u62c5\u5f53\u8005\u304c\u78ba\u8a8d\u5f8c\u3001\u767a\u9001\u624b\u914d\u3044\u3044\u305f\u3057\u307e\u3059\u3002"
    };

  } catch (err) {
    Logger.log("Process Error: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

// =============================================
// \u30c1\u30a7\u30c3\u30af\u30dc\u30c3\u30af\u30b9ON\u306b\u3057\u305f\u77ac\u9593\u306b\u81ea\u52d5\u8d77\u52d5\uff08\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u53ef\u80fd\u30c8\u30ea\u30ac\u30fc\uff09
// =============================================

function onApprovalEdit(e) {
  try {
    const cfg   = getConfig();
    const range = e.range;
    const sheet = range.getSheet();

    if (sheet.getName() !== cfg.SHEET_NAME_PENDING) return;
    if (range.getColumn() !== cfg.APPROVAL_COL) return;
    if (range.getRow() <= 1) return;
    if (e.value !== "TRUE") return;

    const row = range.getRow();
    const ss  = getSpreadsheet();
    const requestId = sheet.getRange(row, 2).getValue();
    if (!requestId) return;

    processApprovedRequest(ss, sheet, requestId);

  } catch (err) {
    Logger.log("onApprovalEdit Error: " + err.toString());
  }
}

function processApprovedRequest(ss, pendingSheet, requestId) {
  const cfg = getConfig();

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log("Lock failed: " + e.toString());
    return;
  }

  try {
    const allData = pendingSheet.getDataRange().getValues();

    // \u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8\u306e\u5217\u5b9a\u7fa9\uff080\u59cb\u307e\u308a\uff09
    // A=0 B=1:\u7ba1\u7406\u756a\u53f7 C=2:\u5e97\u8217\u540d D=3:ID E=4:\u6c0f\u540d F=5:\u4f4f\u6240\u78ba\u8a8d
    // G=6:\u5546\u54c1\u540d H=7:\u5546\u54c1URL I=8:\u30d0\u30ea\u30a8 J=9:\u8a73\u7d30
    // K=10:USER_ID L=11:GROUP_ID M=12:\u2705\u627f\u8a8d N=13:\u627f\u8a8d\u65e5\u6642 O=14:\u30b9\u30c6\u30fc\u30bf\u30b9 P=15:\u767a\u9001\u30b7\u30fc\u30c8URL
    const COL = {
      requestId:1, shopName:2, creatorId:3, realName:4, addrCheck:5,
      prodName:6, prodUrl:7, prodVar:8, prodDetail:9,
      userId:10, groupId:11, approval:12, approvedAt:13, status:14, shopUrl:15
    };

    const targetRows = [];
    for (let i = 1; i < allData.length; i++) {
      const rowId = String(allData[i][COL.requestId] || "");
      const st    = String(allData[i][COL.status]    || "");
      if (rowId === requestId && st !== "\u8ee2\u8a18\u6e08" && st !== "\u51e6\u7406\u4e2d") {
        targetRows.push({ rowIndex: i + 1, rowData: allData[i] });
      }
    }

    if (targetRows.length === 0) {
      Logger.log("No target rows for: " + requestId);
      return;
    }

    targetRows.forEach(r => updatePendingStatus(pendingSheet, r.rowIndex, "\u51e6\u7406\u4e2d..."));

    const first = targetRows[0].rowData;
    const shopName     = String(first[COL.shopName]  || "");
    const creatorId    = String(first[COL.creatorId] || "");
    const realName     = String(first[COL.realName]  || "");
    const addrCheck    = String(first[COL.addrCheck] || "");
    const userId       = String(first[COL.userId]    || "");
    const groupId      = String(first[COL.groupId]   || "");
    const shopSheetUrl = String(first[COL.shopUrl]   || "");

    const products = targetRows.map(r => ({
      name:      String(r.rowData[COL.prodName]   || ""),
      url:       String(r.rowData[COL.prodUrl]    || ""),
      variation: String(r.rowData[COL.prodVar]    || ""),
      details:   String(r.rowData[COL.prodDetail] || "")
    }));

    const addressConfirmed = addrCheck.includes("\u5909\u66f4\u306a\u3057");
    const addressResult = lookupAddress(ss, creatorId, realName, addressConfirmed);

    if (!addressResult.isFound) {
      const errMsg = "\u300e\u8981\u5bfe\u5fdc\u300f" + addressResult.address;
      targetRows.forEach(r => {
        updatePendingStatus(pendingSheet, r.rowIndex, errMsg);
        pendingSheet.getRange(r.rowIndex, cfg.APPROVAL_COL).setValue(false);
      });
      Logger.log("Address not found: " + errMsg);
      return;
    }

    if (!shopSheetUrl || shopSheetUrl.trim() === "" || shopSheetUrl === "null") {
      const errMsg = "\u300e\u8981\u5bfe\u5fdc\u300f\u5e97\u8217\u767a\u9001\u30b7\u30fc\u30c8URL\u672a\u8a2d\u5b9a\uff08F\u5217\u78ba\u8a8d\uff09";
      targetRows.forEach(r => {
        updatePendingStatus(pendingSheet, r.rowIndex, errMsg);
        pendingSheet.getRange(r.rowIndex, cfg.APPROVAL_COL).setValue(false);
      });
      Logger.log("No shop sheet URL for: " + shopName);
      return;
    }

    const data = {
      shopName, shopNotifyGroupId: groupId, shopSheetUrl,
      creatorId, realName, address: addressResult.address,
      userId, requestId, products
    };

    const shopWriteRes = appendToShopExternalSheet(shopSheetUrl, data);
    if (!shopWriteRes || !shopWriteRes.success) {
      const errMsg = "\u300e\u8981\u5bfe\u5fdc\u300f\u5e97\u8217\u30b7\u30fc\u30c8\u66f8\u304d\u8fbc\u307f\u5931\u6557\uff08\u6a29\u9650/URL\u78ba\u8a8d\uff09";
      targetRows.forEach(r => {
        updatePendingStatus(pendingSheet, r.rowIndex, errMsg);
        pendingSheet.getRange(r.rowIndex, cfg.APPROVAL_COL).setValue(false);
      });
      Logger.log("Shop sheet write failed: " + (shopWriteRes ? shopWriteRes.error : "unknown"));
      return;
    }
    const shopRowText = shopWriteRes.rowText;

    const recordResult = appendShippingRecords(ss, data);
    const rowInfoText  = recordResult.rowText;

    notifyShopGroup(data, rowInfoText, shopRowText);

    const approvedAt = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    targetRows.forEach(r => {
      pendingSheet.getRange(r.rowIndex, 14).setValue(approvedAt);
      updatePendingStatus(pendingSheet, r.rowIndex, "\u8ee2\u8a18\u6e08");
    });

    Logger.log("Approved: " + requestId + " / shipping:" + rowInfoText + " / shop:" + shopRowText);

  } finally {
    lock.releaseLock();
  }
}

// =============================================
// \u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB\u30b7\u30fc\u30c8\u306b\u6c38\u4e45\u4fdd\u5b58
// =============================================

function appendCreatorDB(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_CREATOR_DB);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_CREATOR_DB);
    sheet.appendRow([
      "\u53d7\u4ed8\u65e5\u6642", "TikTok ID", "LINE_USER_ID", "LINE_GROUP_ID",
      "\u6c0f\u540d", "\u5546\u54c1\u540d", "\u30d6\u30e9\u30f3\u30c9\uff08\u5e97\u8217\u540d\uff09", "\u5099\u8003\uff08\u5546\u54c1\u30ea\u30f3\u30af\uff09"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 8)
      .setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
  }

  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{ name: data.productName || "\u6307\u5b9a\u306a\u3057", url: data.productUrl || "" }];
  }

  products.forEach(p => {
    sheet.appendRow([
      now,
      data.creatorId             || "",
      data.userId                || "",
      data.shopNotifyGroupId     || data.groupId || "",
      data.realName              || "",
      p.name                     || "",
      data.shopName              || "",
      p.url                      || ""
    ]);
  });
}

// =============================================
// \u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8\u306b\u66f8\u304d\u8fbc\u307f\uff081\u5546\u54c11\u884c\uff09
// =============================================

function appendPendingRecords(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_PENDING);
    sheet.appendRow([
      "\u53d7\u4ed8\u65e5\u6642", "\u7ba1\u7406\u756a\u53f7", "\u5e97\u8217\u540d", "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcID", "\u6c0f\u540d", "\u4f4f\u6240\u78ba\u8a8d",
      "\u5546\u54c1\u540d", "\u5546\u54c1\u30ea\u30f3\u30af", "\u5e0c\u671b\u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3", "\u5546\u54c1\u8a73\u7d30\u30fb\u3054\u8981\u671b",
      "LINE_USER_ID", "LINE_GROUP_ID", "\u2705\u627f\u8a8d", "\u627f\u8a8d\u65e5\u6642", "\u30b9\u30c6\u30fc\u30bf\u30b9", "\u767a\u9001\u30b7\u30fc\u30c8URL"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 16)
      .setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
    sheet.getRange("M2:M2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build()
    );
    sheet.setColumnWidth(13, 60);
    sheet.setColumnWidth(16, 300);
  }

  const now       = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const checkText = data.addressConfirmed ? "\u5909\u66f4\u306a\u3057" : "\u5909\u66f4\u3042\u308a/\u521d\u56de";

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{
      name:      data.productName      || "\u6307\u5b9a\u306a\u3057",
      url:       data.productUrl       || "",
      variation: data.productVariation || data.productColor || "-",
      details:   data.productDetails   || "-"
    }];
  }

  products.forEach(p => {
    sheet.appendRow([
      now,
      data.requestId,
      data.shopName,
      data.creatorId,
      data.realName,
      checkText,
      p.name,
      p.url       || "-",
      p.variation || "-",
      p.details   || "-",
      data.userId                           || "",
      data.shopNotifyGroupId || data.groupId || "",
      false,
      "",
      "\u53d7\u4ed8\u4e2d",
      data.shopSheetUrl                     || ""
    ]);
  });
}

function updatePendingStatus(sheet, rowIndex, statusText) {
  const statusCol = 15;
  const cell = sheet.getRange(rowIndex, statusCol);
  cell.setValue(statusText);
  if (statusText === "\u8ee2\u8a18\u6e08") {
    cell.setBackground("#dcfce7").setFontColor("#166534").setFontWeight("bold");
  } else if (statusText.includes("\u8981\u5bfe\u5fdc") || statusText.includes("\u5931\u6557")) {
    cell.setBackground("#fef08a").setFontColor("#854d0e").setFontWeight("bold");
  } else if (statusText === "\u51e6\u7406\u4e2d...") {
    cell.setBackground("#e0f2fe").setFontColor("#0369a1").setFontWeight("normal");
  } else {
    cell.setBackground("").setFontColor("").setFontWeight("normal");
  }
}

// =============================================
// \u30c8\u30ea\u30ac\u30fc\u767b\u9332\uff081\u56de\u3060\u3051\u5b9f\u884c\uff09
// =============================================

function setupApprovalTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onApprovalEdit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  const ss = getSpreadsheet();
  ScriptApp.newTrigger("onApprovalEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("setupApprovalTrigger: onApprovalEdit trigger registered successfully");
}

// =============================================
// \u4f4f\u6240\u30b7\u30fc\u30c8\u3092\u691c\u7d22\uff08\u52d5\u7684\u30d8\u30c3\u30c0\u30fc\u691c\u7d22\uff09
// =============================================

function lookupAddress(ss, creatorId, realName, addressConfirmed) {
  if (!addressConfirmed) {
    return { isFound: false, address: "\u4f4f\u6240\u5909\u66f4\u306e\u7533\u3057\u51fa\u3042\u308a\uff08\u500b\u5225\u30c1\u30e3\u30c3\u30c8\u5bfe\u5fdc\uff09" };
  }

  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_ADDRESS);
  if (!sheet) sheet = ss.getSheetByName("\u4f4f\u6240\u7ba1\u7406\u30b7\u30fc\u30c8");
  if (!sheet) return { isFound: false, address: "\u4f4f\u6240\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093" };

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { isFound: false, address: "\u4f4f\u6240\u30c7\u30fc\u30bf\u306a\u3057" };

  const headerRow = values[0];
  let idCol = -1, nameCol = -1, zipCol = -1, addrCol = -1;

  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] || "").replace(/[\s\u3000]/g, "").toLowerCase();
    if (idCol === -1   && (h === "id" || h.includes("id") || h.includes("\u30a2\u30ab\u30a6\u30f3\u30c8"))) idCol = c;
    else if (nameCol === -1 && (h === "\u540d\u524d" || h === "\u6c0f\u540d" || h.includes("\u30af\u30ea\u30a8\u30a4\u30bf\u30fc\u540d"))) nameCol = c;
    else if (zipCol === -1  && (h === "\u90f5\u4fbf\u756a\u53f7" || h.includes("\u3012") || h.includes("\u90f5\u4fbf"))) zipCol = c;
    else if (addrCol === -1 && (h === "\u4f4f\u6240" || h.includes("\u304a\u5c4a\u3051\u5148"))) addrCol = c;
  }

  if (idCol === -1)   idCol   = 3;
  if (nameCol === -1) nameCol = 5;
  if (zipCol === -1)  zipCol  = 6;
  if (addrCol === -1) addrCol = 7;

  const targetId   = String(creatorId || "").trim().toLowerCase().replace(/^@/, "");
  const targetName = String(realName  || "").trim().replace(/[\s\u3000]/g, "");

  for (let i = 1; i < values.length; i++) {
    const rowId   = String(values[i][idCol]   || "").trim().toLowerCase().replace(/^@/, "");
    const rowName = String(values[i][nameCol] || "").trim().replace(/[\s\u3000]/g, "");
    const isIdMatch   = targetId   && rowId   && (rowId   === targetId);
    const isNameMatch = targetName && rowName && (rowName === targetName);

    if (isIdMatch || isNameMatch) {
      const rawZip  = String(values[i][zipCol]  || "").trim();
      const rawAddr = String(values[i][addrCol] || "").trim();
      if (!rawAddr && !rawZip) return { isFound: false, address: "\u30b7\u30fc\u30c8\u306b\u767b\u9332\u306f\u3042\u308b\u304c\u4f4f\u6240\u304c\u7a7a\u6b04" };
      let fullAddress = rawAddr;
      if (rawZip && !rawAddr.includes(rawZip)) {
        const fmtZip = rawZip.startsWith("\u3012") ? rawZip : ("\u3012" + rawZip);
        fullAddress  = fmtZip + " " + rawAddr;
      }
      return { isFound: true, address: fullAddress.trim() };
    }
  }

  return { isFound: false, address: "\u30b7\u30fc\u30c8\u672a\u767b\u9332\uff08\u500b\u4eba\u30c1\u30e3\u30c3\u30c8\u3067\u78ba\u8a8d\uff09" };
}

// =============================================
// \u5e97\u8217\u30de\u30b9\u30bf\u304b\u3089\u5e97\u8217\u60c5\u5831\u3092\u53d6\u5f97
// =============================================

function resolveShopInfo(ss, groupId) {
  let shopName      = "\u5e97\u8217\uff08\u672a\u7279\u5b9a\uff09";
  let notifyGroupId = groupId || "";
  let shopSheetUrl  = null;

  const groupName = groupId ? fetchLineGroupName(groupId) : null;
  if (groupName) shopName = groupName;

  const cfg = getConfig();
  let masterSheet = ss.getSheetByName(cfg.SHEET_NAME_SHOP_MASTER);
  if (!masterSheet) masterSheet = ss.getSheetByName("\u5e97\u8217\u30de\u30b9\u30bf");

  if (masterSheet) {
    const lastRow = masterSheet.getLastRow();
    const lastCol = Math.max(masterSheet.getLastColumn(), 7);
    if (lastRow > 1) {
      const values = masterSheet.getRange(1, 1, lastRow, lastCol).getValues();
      for (let i = 1; i < values.length; i++) {
        const rowShop      = String(values[i][0] || "").trim();
        const rowGroupName = String(values[i][2] || "").trim();
        let isMatch = false;
        if (groupName && (
          rowGroupName === groupName ||
          groupName.includes(rowGroupName) ||
          (rowShop && groupName.includes(rowShop))
        )) {
          isMatch = true;
        } else if (values[i][0] === groupId) {
          isMatch = true;
        }
        if (isMatch) {
          if (rowShop) shopName = rowShop;
          const cell = masterSheet.getRange(i + 1, 6);
          shopSheetUrl = extractUrlFromCell(cell);
          break;
        }
      }
    }
  }

  return { shopName, notifyGroupId, shopSheetUrl };
}

function extractUrlFromCell(cellRange) {
  if (!cellRange) return null;
  try {
    const richText = cellRange.getRichTextValue();
    if (richText) { const link = richText.getLinkUrl(); if (link) return link; }
  } catch (e) {}
  try {
    const formula = cellRange.getFormula();
    if (formula) { const match = formula.match(/HYPERLINK\(\s*"([^"]+)"/i); if (match) return match[1]; }
  } catch (e) {}
  const val = String(cellRange.getValue() || "").trim();
  if (val.startsWith("http")) return val;
  return null;
}

// =============================================
// \u5e97\u8217\u5916\u90e8\u767a\u9001\u30b7\u30fc\u30c8\u3078\u8ee2\u8a18
// =============================================

function appendToShopExternalSheet(shopSheetUrl, data) {
  if (!shopSheetUrl) return null;
  try {
    const shopSS      = SpreadsheetApp.openByUrl(shopSheetUrl);
    const targetSheet = shopSS.getSheets()[0];
    if (!targetSheet) return null;

    const values = targetSheet.getDataRange().getValues();
    const now    = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

    let products = data.products;
    if (!products || !Array.isArray(products) || products.length === 0) {
      products = [{ name: data.productName || "\u6307\u5b9a\u306a\u3057", url: "", variation: "-", details: "-" }];
    }

    const startRow = targetSheet.getLastRow() + 1;
    let dateCol=-1, creatorCol=-1, nameCol=-1, addrCol=-1, prodCol=-1, urlCol=-1, varCol=-1, detailCol=-1;

    if (values.length >= 1) {
      const headers = values[0];
      for (let c = 0; c < headers.length; c++) {
        const h = String(headers[c] || "").replace(/[\s\u3000]/g, "").toLowerCase();
        if (dateCol    === -1 && (h.includes("\u65e5") || h.includes("date")))                               dateCol    = c;
        if (creatorCol === -1 && (h.includes("\u30af\u30ea\u30a8\u30a4\u30bf\u30fc") || h.includes("\u30a2\u30ab\u30a6\u30f3\u30c8") || h === "id")) creatorCol = c;
        if (nameCol    === -1 && (h.includes("\u6c0f\u540d") || h.includes("\u540d\u524d") || h.includes("\u5b9b\u540d")))       nameCol    = c;
        if (addrCol    === -1 && (h.includes("\u4f4f\u6240") || h.includes("\u304a\u5c4a\u3051\u5148")))                         addrCol    = c;
        if (prodCol    === -1 && (h.includes("\u5546\u54c1\u540d") || h.includes("\u5546\u54c1")))                           prodCol    = c;
        if (urlCol     === -1 && (h.includes("url") || h.includes("\u30ea\u30f3\u30af")))                            urlCol     = c;
        if (varCol     === -1 && (h.includes("\u30d0\u30ea\u30a8") || h.includes("\u30ab\u30e9\u30fc") || h.includes("\u30b5\u30a4\u30ba") || h.includes("\u898f\u683c"))) varCol = c;
        if (detailCol  === -1 && (h.includes("\u5099\u8003") || h.includes("\u8a73\u7d30") || h.includes("\u8981\u671b")))       detailCol  = c;
      }
    }

    const hasHeaderMap = (nameCol !== -1 || addrCol !== -1 || prodCol !== -1);

    products.forEach(p => {
      if (hasHeaderMap) {
        const rowData = new Array(Math.max(targetSheet.getLastColumn(), 10)).fill("");
        if (dateCol    !== -1) rowData[dateCol]    = now;
        if (creatorCol !== -1) rowData[creatorCol] = data.creatorId;
        if (nameCol    !== -1) rowData[nameCol]    = data.realName;
        if (addrCol    !== -1) rowData[addrCol]    = data.address;
        if (prodCol    !== -1) rowData[prodCol]    = p.name;
        if (urlCol     !== -1) rowData[urlCol]     = p.url       || "-";
        if (varCol     !== -1) rowData[varCol]     = p.variation || "-";
        if (detailCol  !== -1) rowData[detailCol]  = p.details   || "-";
        targetSheet.appendRow(rowData);
      } else {
        targetSheet.appendRow([now, data.creatorId, data.realName, data.address, p.name, p.url || "-", p.variation || "-", p.details || "-"]);
      }
    });

    const endRow  = targetSheet.getLastRow();
    const rowText = (startRow === endRow) ? (startRow + "\u884c\u76ee") : (startRow + "\uff5e" + endRow + "\u884c\u76ee");
    return { success: true, rowText, sheetTitle: targetSheet.getName() };

  } catch (err) {
    Logger.log("Shop sheet write error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

// =============================================
// \u81ea\u793e\u914d\u9001\u30b7\u30fc\u30c8\u3078\u8ee2\u8a18
// =============================================

function appendShippingRecords(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_SHIPPING);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_SHIPPING);
    sheet.appendRow([
      "\u8a18\u5165\u65e5", "\u5e97\u8217\u540d", "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcID", "\u6c0f\u540d", "\u304a\u5c4a\u3051\u5148\u4f4f\u6240",
      "\u5546\u54c1\u540d", "\u5546\u54c1\u30ea\u30f3\u30af", "\u5e0c\u671b\u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3", "\u5546\u54c1\u8a73\u7d30\u30fb\u3054\u8981\u671b",
      "\u767a\u9001\u30b9\u30c6\u30fc\u30bf\u30b9", "\u767a\u9001\u5b8c\u4e86\u65e5", "LINE_USER_ID", "5\u65e5\u5f8c\u30ea\u30de\u30a4\u30f3\u30c9"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13)
      .setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
  }

  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{ name: data.productName || "\u6307\u5b9a\u306a\u3057", url: "", variation: "-", details: "-" }];
  }

  const startRow = sheet.getLastRow() + 1;
  products.forEach(p => {
    sheet.appendRow([
      now, data.shopName, data.creatorId, data.realName, data.address,
      p.name, p.url || "-", p.variation || "-", p.details || "-",
      "\u672a\u767a\u9001", "", data.userId, "\u672a"
    ]);
  });

  const endRow  = sheet.getLastRow();
  const rowText = (startRow === endRow) ? (startRow + "\u884c\u76ee") : (startRow + "\uff5e" + endRow + "\u884c\u76ee");
  return { startRow, endRow, rowText };
}

// =============================================
// \u5e97\u8217LINE\u30b0\u30eb\u30fc\u30d7\u3078\u901a\u77e5
// =============================================

function notifyShopGroup(data, rowText, shopRowText) {
  const targetChatId = data.shopNotifyGroupId;
  if (!targetChatId) { Logger.log("No group ID, skip notification"); return; }

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{ name: data.productName || "\u6307\u5b9a\u306a\u3057", url: "", variation: "", details: "" }];
  }

  let productDetailsText = "";
  if (products.length === 1) {
    const p = products[0];
    productDetailsText = "\u25a0 \u5546\u54c1: " + p.name;
    if (p.url)       productDetailsText += "\n\u25a0 \u30ea\u30f3\u30af: " + p.url;
    if (p.variation) productDetailsText += "\n\u25a0 \u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3: " + p.variation;
    if (p.details)   productDetailsText += "\n\u25a0 \u8a73\u7d30\u30fb\u3054\u8981\u671b: " + p.details;
  } else {
    productDetailsText = "\u25a0 \u7533\u8acb\u5546\u54c1 (\u8a08" + products.length + "\u70b9):\n" + products.map((p, i) => {
      let item = "\u30fb\u3010\u5546\u54c1" + (i + 1) + "\u3011" + p.name;
      if (p.variation) item += "\n  \u2514 \u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3: " + p.variation;
      if (p.url)       item += "\n  \u2514 \u30ea\u30f3\u30af: " + p.url;
      if (p.details)   item += "\n  \u2514 \u8a73\u7d30\u30fb\u5099\u8003: " + p.details;
      return item;
    }).join("\n");
  }

  const shopSheetPart = shopRowText ? ("\n\u25a0 \u5e97\u8217\u69d8\u767a\u9001\u30b7\u30fc\u30c8: \u3010" + shopRowText + "\u3011\u306b\u81ea\u52d5\u8a18\u5165\u6e08") : "";

  const messageText =
    "\u304a\u75b2\u308c\u69d8\u3067\u3059\uff01\n" +
    "\u5546\u54c1\u30b5\u30f3\u30d7\u30eb\u306e\u767a\u9001\u4f9d\u983c\u3092\u304a\u9858\u3044\u3044\u305f\u3057\u307e\u3059\uff01\n\n" +
    "\u25a0 \u5e97\u8217\u69d8: " + data.shopName + "\n" +
    "\u25a0 \u30af\u30ea\u30a8\u30a4\u30bf\u30fc: " + data.creatorId + " (" + data.realName + "\u69d8)\n\n" +
    productDetailsText + "\n\n" +
    "\u25a0 \u304a\u5c4a\u3051\u5148\u60c5\u5831:\n" +
    "\u30fb\u304a\u540d\u524d: " + data.realName + " \u69d8\n" +
    "\u30fb\u3054\u4f4f\u6240: " + (data.address || "\u78ba\u8a8d\u4e2d") +
    shopSheetPart + "\n" +
    "\u25a0 \u914d\u9001\u30b7\u30fc\u30c8\u7ba1\u7406\u756a\u53f7: \u3010" + rowText + "\u3011\n\n" +
    "\u3054\u78ba\u8a8d\u306e\u4e0a\u3001\u767a\u9001\u624b\u914d\u3092\u304a\u9858\u3044\u3044\u305f\u3057\u307e\u3059\ud83d\ude47\u200d\u2642\ufe0f";

  pushLineMessage(targetChatId, messageText);
}

// =============================================
// \u767a\u9001\u5f8c5\u65e5\u30ea\u30de\u30a4\u30f3\u30c9\uff08\u6bce\u65e5\u5b9a\u671f\u30c8\u30ea\u30ac\u30fc\uff09
// =============================================

function checkAndSendReminders() {
  const cfg   = getConfig();
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(cfg.SHEET_NAME_SHIPPING);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const status       = data[i][9];
    const shippedDate  = data[i][10];
    const userId       = data[i][11];
    const remindStatus = data[i][12];
    const productName  = data[i][5];

    if (status === "\u767a\u9001\u6e08" && remindStatus === "\u672a" && shippedDate && userId) {
      const shipped = new Date(shippedDate);
      shipped.setHours(0, 0, 0, 0);
      const diffDays = Math.floor((today - shipped) / (1000 * 60 * 60 * 24));

      if (diffDays >= 5) {
        const remindText =
          "\u304a\u4e16\u8a71\u306b\u306a\u3063\u3066\u304a\u308a\u307e\u3059\uff01\n" +
          "\u5148\u65e5\u767a\u9001\u3044\u305f\u3057\u307e\u3057\u305f\u30b5\u30f3\u30d7\u30eb\u300c" + productName + "\u300d\u306f\u304a\u624b\u5143\u306b\u5c4a\u304d\u307e\u3057\u305f\u3067\u3057\u3087\u3046\u304b\uff1f\u2728\n\n" +
          "\u3010\u3054\u78ba\u8a8d\u306e\u304a\u9858\u3044\u3011\n" +
          "\u2460 TikTok\u30e9\u30a4\u30d6\u3067\u306e\u3054\u7d39\u4ecb\u65e5\u7a0b\u304c\u304a\u6c7a\u307e\u308a\u3067\u3057\u305f\u3089\u3001\u672c\u30c1\u30e3\u30c3\u30c8\u306b\u3066\u6559\u3048\u3066\u3044\u305f\u3060\u3051\u307e\u3059\u3068\u5e78\u3044\u3067\u3059\uff01\n" +
          "\u2461 \u3054\u7d39\u4ecb\u7528\u306e\u30a2\u30d5\u30a3\u30ea\u30a8\u30a4\u30c8\u30ea\u30f3\u30af\u306f\u4ee5\u4e0b\u3068\u306a\u308a\u307e\u3059\u3002\u5fc5\u305a\u3053\u3061\u3089\u3092\u3054\u6d3b\u7528\u304f\u3060\u3055\u3044\u307e\u305b\uff01\n" +
          "\ud83d\udd17 " + cfg.AFFILIATE_INFO_TEXT + "\n\n" +
          "\u5f15\u304d\u7d9a\u304d\u3088\u308d\u3057\u304f\u304a\u9858\u3044\u3044\u305f\u3057\u307e\u3059\uff01";

        pushLineMessage(userId, remindText);
        sheet.getRange(i + 1, 13).setValue(
          "\u9001\u4fe1\u6e08 (" + Utilities.formatDate(today, "Asia/Tokyo", "MM/dd") + ")"
        );
      }
    }
  }
}

// =============================================
// LINE Messaging API
// =============================================

function fetchLineGroupName(groupId) {
  const token = getConfig().LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return null;
  try {
    const res = UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/group/" + groupId + "/summary",
      { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).groupName;
  } catch (e) { Logger.log("Group name fetch failed: " + e.toString()); }
  return null;
}

function pushLineMessage(toId, text) {
  const token = getConfig().LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return;
  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
      payload: JSON.stringify({ to: toId, messages: [{ type: "text", text: text }] }),
      muteHttpExceptions: true
    });
  } catch (e) { Logger.log("Push error: " + e.toString()); }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
