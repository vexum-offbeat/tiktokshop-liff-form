/**
 * TikToker商品サンプル発送依頼 & リマインド自動化システム
 * Google Apps Script (GAS) バックエンド
 * v2.1 - 承認制ワークフロー / 配送シート廃止 / リマインド予定日を承認待ちシートで管理
 */

// ====================================================================
// スクリプトプロパティ一覧（GASの「プロジェクトの設定」から確認・編集）
//
// プロパティ名               説明                           デフォルト値
// SPREADSHEET_ID          スプレッドシートのID             （必須）
// LINE_CHANNEL_TOKEN      LINEチャネルアクセストークン      （必須）
// LIFF_ID                 LINE DevelopersのLIFF ID         （必須）
// AFFILIATE_URL           アフィリエイトリンク              （任意）
// SHEET_NAME_PENDING      承認待ちシートのタブ名            承認待ちシート
// SHEET_NAME_CREATOR_DB   クリエイターDBのタブ名            クリエイターDB
// SHEET_NAME_ADDRESS      住所一覧シートのタブ名            クリエイター・代理店サンプルリスト
// SHEET_NAME_SHOP_MASTER  店舗マスタのタブ名               店舗発送依頼先情報
// APPROVAL_COL            承認チェックボックスの列番号      13（M列）
// REMINDER_DAYS           発送依頼からリマインドまでの日数  5
// ====================================================================

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID:            props.getProperty("SPREADSHEET_ID")         || "10v-c76tg20i5sSzIM6T7-QQ7kTNLPiNOZhvF0A8NDVg",
    LINE_CHANNEL_ACCESS_TOKEN: props.getProperty("LINE_CHANNEL_TOKEN")     || "",
    SHEET_NAME_PENDING:        props.getProperty("SHEET_NAME_PENDING")     || "\u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8",
    SHEET_NAME_CREATOR_DB:     props.getProperty("SHEET_NAME_CREATOR_DB")  || "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB",
    SHEET_NAME_ADDRESS:        props.getProperty("SHEET_NAME_ADDRESS")     || "\u30af\u30ea\u30a8\u30a4\u30bf\u30fc\u30fb\u4ee3\u7406\u5e97\u30b5\u30f3\u30d7\u30eb\u30ea\u30b9\u30c8",
    SHEET_NAME_SHOP_MASTER:    props.getProperty("SHEET_NAME_SHOP_MASTER") || "\u5e97\u8217\u767a\u9001\u4f9d\u983c\u5148\u60c5\u5831",
    AFFILIATE_INFO_TEXT:       props.getProperty("AFFILIATE_URL")          || "",
    APPROVAL_COL:              Number(props.getProperty("APPROVAL_COL"))   || 13,
    REMINDER_DAYS:             Number(props.getProperty("REMINDER_DAYS"))  || 5
  };
}

/**
 * スクリプトプロパティに初期値を一括登録する
 * GASエディタで「initializeScriptProperties」を1回だけ実行すればOK
 * すでに設定済みのプロパティは上書きしません
 */
function initializeScriptProperties() {
  const props   = PropertiesService.getScriptProperties();
  const current = props.getProperties();

  const defaults = {
    "SPREADSHEET_ID":        "10v-c76tg20i5sSzIM6T7-QQ7kTNLPiNOZhvF0A8NDVg",
    "LINE_CHANNEL_TOKEN":    "",
    "LIFF_ID":               "",
    "AFFILIATE_URL":         "",
    "SHEET_NAME_PENDING":    "\u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8",
    "SHEET_NAME_CREATOR_DB": "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB",
    "SHEET_NAME_ADDRESS":    "\u30af\u30ea\u30a8\u30a4\u30bf\u30fc\u30fb\u4ee3\u7406\u5e97\u30b5\u30f3\u30d7\u30eb\u30ea\u30b9\u30c8",
    "SHEET_NAME_SHOP_MASTER":"\u5e97\u8217\u767a\u9001\u4f9d\u983c\u5148\u60c5\u5831",
    "APPROVAL_COL":          "13",
    "REMINDER_DAYS":         "5"
  };

  const toSet = {}, skipped = [], added = [];
  Object.keys(defaults).forEach(key => {
    if (current[key] !== undefined && current[key] !== "") {
      skipped.push(key + " = " + current[key]);
    } else {
      toSet[key] = defaults[key];
      added.push(key + " = " + defaults[key]);
    }
  });

  if (Object.keys(toSet).length > 0) props.setProperties(toSet);

  Logger.log("=== initializeScriptProperties ===");
  Logger.log("[\u767b\u9332\u6e08\u30fb\u30b9\u30ad\u30c3\u30d7]\n" + (skipped.length ? skipped.join("\n") : "\u306a\u3057"));
  Logger.log("[\u65b0\u898f\u8ffd\u52a0]\n" + (added.length ? added.join("\n") : "\u306a\u3057"));
  Logger.log("\u5b8c\u4e86\uff01GAS\u306e\u300c\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u306e\u8a2d\u5b9a > \u30b9\u30af\u30ea\u30d7\u30c8\u30d7\u30ed\u30d1\u30c6\u30a3\u300d\u3067\u5185\u5bb9\u3092\u78ba\u8a8d\u30fb\u7de8\u96c6\u3067\u304d\u307e\u3059\u3002");
}

function getSpreadsheet() {
  const cfg = getConfig();
  if (cfg.SPREADSHEET_ID && cfg.SPREADSHEET_ID.trim()) {
    try { return SpreadsheetApp.openById(cfg.SPREADSHEET_ID.trim()); }
    catch (e) { Logger.log("openById failed: " + e.toString()); }
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error(
    "\u30b9\u30d7\u30ec\u30c3\u30c9\u30b7\u30fc\u30c8\u3092\u958b\u3051\u307e\u305b\u3093\u3067\u3057\u305f\u3002" +
    "GAS\u306e\u300c\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u306e\u8a2d\u5b9a > \u30b9\u30af\u30ea\u30d7\u30c8\u30d7\u30ed\u30d1\u30c6\u30a3\u300d\u3067" +
    "SPREADSHEET_ID\u3092\u8a2d\u5b9a\u3057\u3066\u304f\u3060\u3055\u3044\u3002" +
    "\u73fe\u5728\u306eID: " + (getConfig().SPREADSHEET_ID || "\u672a\u8a2d\u5b9a")
  );
}

// =============================================
// LIFF\u30d5\u30a9\u30fc\u30e0\u8868\u793a\u30fb\u9001\u4fe1\u51e6\u7406
// =============================================

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "brands") {
    return createJsonResponse(getBrandList());
  }
  const liffId = PropertiesService.getScriptProperties().getProperty("LIFF_ID") || "";
  const brandList = getBrandList();
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setTitle("\u5546\u54c1\u30b5\u30f3\u30d7\u30eb\u767a\u9001\u30fb\u7d39\u4ecb\u53d7\u4ed8\u30d5\u30a9\u30fc\u30e0")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.setContent(html.getContent()
    .replace("__LIFF_ID__", liffId)
    .replace("__BRAND_LIST_JSON__", JSON.stringify(brandList))
  );
  return html;
}

function submitShippingRequest(data) { return processRequestData(data); }

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return createJsonResponse(processRequestData(data));
  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

// =============================================
// \u30d5\u30a9\u30fc\u30e0\u9001\u4fe1\u51e6\u7406\uff08\u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB + \u627f\u8a8d\u5f85\u3061\u30b7\u30fc\u30c8\u306b\u66f8\u304f\u3060\u3051\uff09
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

    // クリエイターDBに永久保存（バックアップ）
    appendCreatorDB(ss, data);

    // 承認待ちシートに書き込んで止まる
    appendPendingRecords(ss, data);

    // 住所未確認の場合、フォーム送信直後にクリエイターへLINE自動送信
    if (!data.addressConfirmed) {
      const targetId = data.userId || data.groupId;
      if (targetId) {
        const addrRequestMsg =
          "\u4f4f\u6240\u5909\u66f4\u306e\u3054\u9023\u7d61\u3042\u308a\u304c\u3068\u3046\u3054\u3056\u3044\u307e\u3059\u3002\n" +
          "\u6050\u308c\u5165\u308a\u307e\u3059\u304c\u3001\u304a\u5c4a\u3051\u5148\u306e\u3054\u4f4f\u6240\u3092\u3053\u3061\u3089\u306e\u30c1\u30e3\u30c3\u30c8\u306b\u3054\u8a18\u5165\u304f\u3060\u3055\u3044\u3002";
        pushLineMessage(targetId, addrRequestMsg);
        Logger.log("Address request sent to: " + targetId);
      } else {
        Logger.log("Address request skipped: No userId or groupId");
      }
    }

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
// \u627f\u8a8d\u30c1\u30a7\u30c3\u30af\u6642\u306b\u5373\u6642\u8d77\u52d5\uff08\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u53ef\u80fd\u30c8\u30ea\u30ac\u30fc\uff09
// =============================================

/**
 * onApprovalEdit は廃止しました。
 * 承認は「発送システム」メニュー → 「選択行を承認する」から行ってください。
 */
function onApprovalEdit(e) {
  // 旧チェックボックストリガー方式は廃止。何もしない。
}

/**
 * 承認待ちシートの列定義（1始まり）
 * A=1  B=2  C=3  D=4  E=5  F=6  G=7  H=8  I=9  J=10
 * K=11 L=12 M=13 N=14 O=15 P=16 Q=17 R=18
 *
 * A: 受付日時       B: 管理番号       C: 店舗名         D: クリエイターID  E: 氏名
 * F: 住所確認       G: 商品名         H: 商品リンク     I: 希望バリエ     J: 商品詳細
 * K: LINE_USER_ID  L: LINE_GROUP_ID  M: ✅承認         N: 承認日時       O: ステータス
 * P: 発送シートURL  Q: リマインド予定日  R: リマインド送信済
 */
var PENDING_COL = {
  recvAt:1, reqId:2, shopName:3, creatorId:4, realName:5, addrCheck:6,
  prodName:7, prodUrl:8, prodVar:9, prodDetail:10,
  userId:11, groupId:12, approval:13, approvedAt:14, status:15,
  shopUrl:16, remindDate:17, remindSent:18
};

/**
 * 承認処理の本体。住所が確定した後に呼ばれる。
 * 各行のC列（ブランド名）をもとに「店舗発送依頼先情報」タブのJ列を検索し、
 * F列（「発送シート」列）のURLを取得して外部スプレッドシートへ転記する。
 * @param {Spreadsheet} ss
 * @param {Sheet} pendingSheet
 * @param {string} requestId
 * @param {string} selectedAddress  確定した住所文字列
 * @returns {{ success: boolean, message: string }}
 */
function executeApproval(ss, pendingSheet, requestId, selectedAddress) {
  const cfg  = getConfig();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { success: false, message: "ロック取得に失敗しました。もう一度試してください。" }; }

  try {
    const allData = pendingSheet.getDataRange().getValues();

    // 同じ管理番号の行を収集（未転記のみ）
    const targetRows = [];
    for (let i = 1; i < allData.length; i++) {
      const rowId = String(allData[i][PENDING_COL.reqId - 1]  || "");
      const st    = String(allData[i][PENDING_COL.status - 1] || "");
      if (rowId === requestId && st !== "転記済" && st !== "処理中...") {
        targetRows.push({ rowIndex: i + 1, rowData: allData[i] });
      }
    }

    if (targetRows.length === 0) {
      return { success: false, message: "対象行が見つかりません（すでに転記済みの可能性があります）。" };
    }

    // 転記前に全行の発送シートURLを「店舗発送依頼先情報」タブ（J列→F列）から解決・検証
    const missingBrands = [];
    const resolvedItems = targetRows.map(r => {
      const brandName = String(r.rowData[PENDING_COL.shopName - 1] || "").trim();
      const shopInfo  = resolveShopInfoByName(ss, brandName);
      const url       = shopInfo && shopInfo.shopSheetUrl ? shopInfo.shopSheetUrl : "";
      if (!url) {
        missingBrands.push({ rowIndex: r.rowIndex, brandName: brandName || "（未入力）" });
      }
      return { row: r, brandName: brandName, shopInfo: shopInfo, shopSheetUrl: url };
    });

    if (missingBrands.length > 0) {
      const errMsg = "《要対応》発送シート未特定（店舗発送依頼先情報のJ列・F列確認）";
      missingBrands.forEach(m => updatePendingStatus(pendingSheet, m.rowIndex, errMsg));
      const names = [...new Set(missingBrands.map(m => m.brandName))].join(", ");
      return {
        success: false,
        message: "店舗の発送シートが見つかりませんでした。\n「" + cfg.SHEET_NAME_SHOP_MASTER + "」タブのJ列（ブランド名）とF列（発送シートリンク）を確認してください。\n\n対象ブランド: " + names
      };
    }

    // 二重実行防止
    targetRows.forEach(r => updatePendingStatus(pendingSheet, r.rowIndex, "処理中..."));

    // 同じ発送シートURLごとにグループ化して転記（1申請に複数ブランドが混在しても正しく振り分け）
    const groupsMap = {};
    resolvedItems.forEach(item => {
      const key = item.shopSheetUrl;
      if (!groupsMap[key]) groupsMap[key] = [];
      groupsMap[key].push(item);
    });

    const approvedAt      = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    const reminderDate    = new Date();
    reminderDate.setDate(reminderDate.getDate() + cfg.REMINDER_DAYS);
    const reminderDateStr = Utilities.formatDate(reminderDate, "Asia/Tokyo", "yyyy/MM/dd");

    const resultsSummary = [];
    const failedMessages = [];

    Object.keys(groupsMap).forEach(shopSheetUrl => {
      const groupItems = groupsMap[shopSheetUrl];
      const firstRow   = groupItems[0].row.rowData;
      const firstInfo  = groupItems[0].shopInfo;

      const shopName  = [...new Set(groupItems.map(it => it.brandName).filter(Boolean))].join(" / ");
      const creatorId = String(firstRow[PENDING_COL.creatorId - 1] || "");
      const realName  = String(firstRow[PENDING_COL.realName  - 1] || "");
      const userId    = String(firstRow[PENDING_COL.userId    - 1] || "");
      const groupId   = (firstInfo && firstInfo.notifyGroupId) || String(firstRow[PENDING_COL.groupId - 1] || "");

      const products = groupItems.map(it => ({
        name:      String(it.row.rowData[PENDING_COL.prodName   - 1] || ""),
        url:       String(it.row.rowData[PENDING_COL.prodUrl    - 1] || ""),
        variation: String(it.row.rowData[PENDING_COL.prodVar    - 1] || ""),
        details:   String(it.row.rowData[PENDING_COL.prodDetail - 1] || "")
      }));

      const data = {
        shopName, shopNotifyGroupId: groupId, shopSheetUrl,
        creatorId, realName, address: selectedAddress,
        userId, requestId, products,
        approvedAt: approvedAt  // 承認日時（yyyy/MM/dd HH:mm）
      };

      // 店舗外部発送シートへ転記
      const shopWriteRes = appendToShopExternalSheet(shopSheetUrl, data);
      if (!shopWriteRes || !shopWriteRes.success) {
        const errMsg = "《要対応》店舗シート書き込み失敗（権限/URL確認）";
        groupItems.forEach(it => updatePendingStatus(pendingSheet, it.row.rowIndex, errMsg));
        failedMessages.push("[" + shopName + "] " + (shopWriteRes ? shopWriteRes.error : "書き込み失敗"));
        return;
      }

      // 店舗LINEグループへ発送依頼通知
      notifyShopGroup(data, shopWriteRes.rowText);

      // 承認日時・リマインド予定日・取得した発送シートURLを記録して「転記済」に更新
      groupItems.forEach(it => {
        const idx = it.row.rowIndex;
        pendingSheet.getRange(idx, PENDING_COL.approvedAt).setValue(approvedAt);    // N: 承認日時
        pendingSheet.getRange(idx, PENDING_COL.shopUrl).setValue(shopSheetUrl);     // P: 発送シートURL（記録用）
        pendingSheet.getRange(idx, PENDING_COL.remindDate).setValue(reminderDateStr); // Q: リマインド予定日
        pendingSheet.getRange(idx, PENDING_COL.remindSent).setValue("未");           // R: リマインド送信済
        pendingSheet.getRange(idx, cfg.APPROVAL_COL).setValue(true);                // M: ✅承認（完了後にON）
        updatePendingStatus(pendingSheet, idx, "転記済");
      });

      resultsSummary.push("[" + shopName + "] " + shopWriteRes.rowText);
      Logger.log("executeApproval: " + requestId + " / brand: " + shopName + " / shop: " + shopWriteRes.rowText + " / remind: " + reminderDateStr);
    });

    if (failedMessages.length > 0) {
      return {
        success: false,
        message: "一部の店舗シートへの書き込みに失敗しました。権限やURLを確認してください。\n" +
                 failedMessages.join("\n") +
                 (resultsSummary.length > 0 ? "\n\n成功分:\n" + resultsSummary.join("\n") : "")
      };
    }

    return {
      success: true,
      message: "✅ 転記・通知が完了しました！\n" + resultsSummary.join("\n") + "\n住所: " + selectedAddress
    };

  } finally {
    lock.releaseLock();
  }
}

// processApprovedRequest は executeApproval に統合しました（後方互換のためエイリアスを残す）
function processApprovedRequest(ss, pendingSheet, requestId) {
  Logger.log("processApprovedRequest: このメソッドは廃止予定です。executeApproval を使ってください。");
}

// =============================================
// \u30af\u30ea\u30a8\u30a4\u30bf\u30fcDB\u30b7\u30fc\u30c8\u306b\u6c38\u4e45\u4fdd\u5b58
// =============================================

function appendCreatorDB(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_CREATOR_DB);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_CREATOR_DB);
    sheet.appendRow(["\u53d7\u4ed8\u65e5\u6642","TikTok ID","LINE_USER_ID","LINE_GROUP_ID","\u6c0f\u540d","\u5546\u54c1\u540d","\u30d6\u30e9\u30f3\u30c9\uff08\u5e97\u8217\u540d\uff09","\u5099\u8003\uff08\u5546\u54c1\u30ea\u30f3\u30af\uff09"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,8).setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
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
      p.brand                    || data.shopName || "",
      p.url                      || ""
    ]);
  });
}

// =============================================
// 承認待ちシートに書き込み（1商品1行）
// =============================================

/**
 * 承認待ちシートへの追記
 * 列: A:受付日時 B:管理番号 C:店舗名 D:クリエイターID E:氏名 F:住所確認
 *     G:商品名 H:商品リンク I:希望バリエーション J:商品詳細
 *     K:LINE_USER_ID L:LINE_GROUP_ID M:✅承認 N:承認日時 O:ステータス P:発送シートURL
 *     Q:リマインド予定日 R:リマインド送信済
 */
function appendPendingRecords(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  if (!sheet) {
    formatPendingSheet();
    sheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  }

  const now       = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const checkText = data.addressConfirmed ? "\u5909\u66f4\u306a\u3057" : "\u5909\u66f4\u3042\u308a/\u521d\u56de";

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{
      name:      data.productName      || "\u6307\u5b9a\u306a\u3057",
      url:       data.productUrl       || "",
      brand:     data.shopName         || "",
      variation: data.productVariation || data.productColor || "-",
      details:   data.productDetails   || "-"
    }];
  }

  products.forEach(p => {
    let rowShopName = p.brand || data.shopName;
    let rowShopUrl  = data.shopSheetUrl || "";
    let rowGroupId  = data.shopNotifyGroupId || data.groupId || "";

    if (p.brand) {
      rowShopName = p.brand; // クリエイターが選んだブランド名をそのままC列に残す
      const info = resolveShopInfoByName(ss, p.brand);
      if (info) {
        if (info.shopSheetUrl)  rowShopUrl = info.shopSheetUrl;
        if (info.notifyGroupId) rowGroupId = info.notifyGroupId;
      }
    }

    const rowData = [
      now,                                              // A: 受付日時
      data.requestId,                                   // B: 管理番号
      rowShopName,                                      // C: 店舗名（ブランド名）
      data.creatorId,                                   // D: クリエイターID
      data.realName,                                    // E: 氏名
      checkText,                                        // F: 住所確認
      p.name,                                           // G: 商品名
      p.url       || "-",                               // H: 商品リンク
      p.variation || "-",                               // I: 希望バリエーション
      p.details   || "-",                               // J: 商品詳細
      data.userId                           || "",      // K: LINE_USER_ID
      rowGroupId,                                       // L: LINE_GROUP_ID
      false,                                            // M: ✅承認
      "",                                               // N: 承認日時
      "\u53d7\u4ed8\u4e2d",                             // O: ステータス
      rowShopUrl,                                       // P: 発送シートURL
      "",                                               // Q: リマインド予定日
      ""                                                // R: リマインド送信済
    ];

    // A列（受付日時）が空いている最初の行を確実に探して書き込む
    const targetRow = findFirstEmptyRowInPending(sheet);
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    // 承認列のチェックボックスと書式を保証
    sheet.getRange(targetRow, PENDING_COL.approval).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build()
    );
  });
}

/**
 * 承認待ちシートで実際にデータが入っていない最初の行（A列が空の行）を探す
 */
function findFirstEmptyRowInPending(sheet) {
  const colA = sheet.getRange("A:A").getValues();
  for (let r = 1; r < colA.length; r++) {
    if (!colA[r][0] || String(colA[r][0]).trim() === "") {
      return r + 1; // 1始まりの行番号
    }
  }
  return colA.length + 1;
}

function updatePendingStatus(sheet, rowIndex, statusText) {
  const cell = sheet.getRange(rowIndex, PENDING_COL.status);
  cell.setValue(statusText);
  if      (statusText === "\u8ee2\u8a18\u6e08")                              cell.setBackground("#dcfce7").setFontColor("#166534").setFontWeight("bold");
  else if (statusText.includes("\u8981\u5bfe\u5fdc") || statusText.includes("\u5931\u6557")) cell.setBackground("#fef08a").setFontColor("#854d0e").setFontWeight("bold");
  else if (statusText === "\u51e6\u7406\u4e2d...")                           cell.setBackground("#e0f2fe").setFontColor("#0369a1").setFontWeight("normal");
  else                                                                        cell.setBackground("").setFontColor("").setFontWeight("normal");
}

/**
 * 承認待ちシートのデザイン・レイアウトを一括整形する関数
 * GASエディタで「formatPendingSheet」を実行すると、ヘッダー・列幅・交互の背景色などが即時適用されます
 */
function formatPendingSheet() {
  const cfg = getConfig();
  const ss  = getSpreadsheet();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_PENDING);
  }

  const headers = [
    "\u53d7\u4ed8\u65e5\u6642", "\u7ba1\u7406\u756a\u53f7", "\u5e97\u8217\u540d", "\u30af\u30ea\u30a8\u30a4\u30bf\u30fcID", "\u6c0f\u540d", "\u4f4f\u6240\u78ba\u8a8d",
    "\u5546\u54c1\u540d", "\u5546\u54c1\u30ea\u30f3\u30af", "\u5e0c\u671b\u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3", "\u5546\u54c1\u8a73\u7d30\u30fb\u3054\u8981\u671b",
    "LINE_USER_ID", "LINE_GROUP_ID", "\u2705\u627f\u8a8d", "\u627f\u8a8d\u65e5\u6642", "\u30b9\u30c6\u30fc\u30bf\u30b9", "\u767a\u9001\u30b7\u30fc\u30c8URL",
    "\u30ea\u30de\u30a4\u30f3\u30c9\u4e88\u5b9a\u65e5", "\u30ea\u30de\u30a4\u30f3\u30c9\u9001\u4fe1\u6e08"
  ];

  // 1. 既存の有効データを行番号に関係なく救出（下のほうに飛んだデータも救済）
  const allValues = sheet.getDataRange().getValues();
  const validDataRows = [];
  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];
    // A列(受付日時)またはB列(管理番号)が入っていれば有効データとみなす
    if ((row[0] && String(row[0]).trim()) || (row[1] && String(row[1]).trim())) {
      validDataRows.push(row.slice(0, headers.length));
    }
  }

  // 2. 2行目以降をクリア（余計な空白行や飛び地を初期化）
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).clearContent().clearDataValidations();
  }

  // 3. 救出した有効データを2行目から隙間なくきれいに配置！
  if (validDataRows.length > 0) {
    sheet.getRange(2, 1, validDataRows.length, headers.length).setValues(validDataRows);
  }

  // 4. ヘッダー設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 38);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // 5. 各列の幅（文字が見切れないよう最適化）
  const colWidths = {
    1: 140, // A: 受付日時
    2: 190, // B: 管理番号
    3: 130, // C: 店舗名
    4: 130, // D: クリエイターID
    5: 110, // E: 氏名
    6: 100, // F: 住所確認
    7: 190, // G: 商品名
    8: 160, // H: 商品リンク
    9: 130, // I: 希望バリエーション
    10: 180, // J: 商品詳細・ご要望
    11: 100, // K: LINE_USER_ID
    12: 100, // L: LINE_GROUP_ID
    13: 70,  // M: ✅承認
    14: 140, // N: 承認日時
    15: 110, // O: ステータス
    16: 150, // P: 発送シートURL
    17: 120, // Q: リマインド予定日
    18: 120  // R: リマインド送信済
  };
  Object.keys(colWidths).forEach(c => {
    sheet.setColumnWidth(Number(c), colWidths[c]);
  });

  // 6. データ行の高さと書式
  const targetRowCount = Math.max(validDataRows.length, 30);
  sheet.setRowHeights(2, targetRowCount, 32);
  sheet.getRange(2, 1, targetRowCount, headers.length).setVerticalAlignment("middle").setFontSize(10);

  // 中央揃え: A, B, F, K, L, M, N, O, Q, R
  [1, 2, 6, 11, 12, 13, 14, 15, 17, 18].forEach(col => {
    sheet.getRange(2, col, targetRowCount, 1).setHorizontalAlignment("center");
  });
  // 左揃え: C, D, E, G, H, I, J, P
  [3, 4, 5, 7, 8, 9, 10, 16].forEach(col => {
    sheet.getRange(2, col, targetRowCount, 1).setHorizontalAlignment("left");
  });

  // 7. 有効データが存在する行にチェックボックスを設定
  if (validDataRows.length > 0) {
    sheet.getRange(2, PENDING_COL.approval, validDataRows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build()
    );
  }

  // 8. 交互の背景色（ゼブラストライプ）の適用
  sheet.getBandings().forEach(b => b.remove());
  const dataRange = sheet.getRange(1, 1, Math.max(validDataRows.length + 1, 10), headers.length);
  const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor("#1e293b");
  banding.setFirstRowColor("#ffffff");
  banding.setSecondRowColor("#f8fafc");

  Logger.log("formatPendingSheet: 承認待ちシートの整形とデータ復元が完了しました！ 有効件数: " + validDataRows.length);
}

// =============================================
// \u30c8\u30ea\u30ac\u30fc\u767b\u9332\uff081\u56de\u3060\u3051\u5b9f\u884c\uff09
// =============================================

function setupApprovalTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onApprovalEdit") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("onApprovalEdit")
    .forSpreadsheet(getSpreadsheet())
    .onEdit()
    .create();

  Logger.log("setupApprovalTrigger: onApprovalEdit trigger registered successfully");
}

// =============================================
// \u767a\u9001\u4f9d\u983c\u304b\u30895\u65e5\u5f8c\u30ea\u30de\u30a4\u30f3\u30c9\uff08\u6bce\u65e5\u5b9a\u671f\u30c8\u30ea\u30ac\u30fc\uff09
// =============================================

/**
 * 承認待ちシートをスキャンし、リマインド予定日を過ぎた行のクリエイターへLINEリマインドを送信
 * GASのトリガーで毎日9〜10時に自動実行するよう設定する
 */
function checkAndSendReminders() {
  const cfg   = getConfig();
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  if (!sheet) { Logger.log("Pending sheet not found"); return; }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const status      = String(data[i][PENDING_COL.status     - 1] || "");
    const remindDate  =        data[i][PENDING_COL.remindDate - 1];
    const remindSent  = String(data[i][PENDING_COL.remindSent - 1] || "");
    const userId      = String(data[i][PENDING_COL.userId     - 1] || "");
    const productName = String(data[i][PENDING_COL.prodName   - 1] || "");

    if (status !== "\u8ee2\u8a18\u6e08")   continue;
    if (remindSent === "\u9001\u4fe1\u6e08" || remindSent.startsWith("\u9001\u4fe1\u6e08")) continue;
    if (!remindDate || !userId) continue;

    const rDate = new Date(remindDate);
    rDate.setHours(0, 0, 0, 0);
    if (today < rDate) continue;

    const remindText =
      "\u304a\u4e16\u8a71\u306b\u306a\u3063\u3066\u304a\u308a\u307e\u3059\uff01\n" +
      "\u5148\u65e5\u767a\u9001\u3044\u305f\u3057\u307e\u3057\u305f\u30b5\u30f3\u30d7\u30eb\u300c" + productName + "\u300d\u306f\u304a\u624b\u5143\u306b\u5c4a\u304d\u307e\u3057\u305f\u3067\u3057\u3087\u3046\u304b\uff1f\u2728\n\n" +
      "\u3010\u3054\u78ba\u8a8d\u306e\u304a\u9858\u3044\u3011\n" +
      "\u2460 TikTok\u30e9\u30a4\u30d6\u3067\u306e\u3054\u7d39\u4ecb\u65e5\u7a0b\u304c\u304a\u6c7a\u307e\u308a\u3067\u3057\u305f\u3089\u3001\u672c\u30c1\u30e3\u30c3\u30c8\u306b\u3066\u6559\u3048\u3066\u3044\u305f\u3060\u3051\u307e\u3059\u3068\u5e78\u3044\u3067\u3059\uff01\n" +
      "\u2461 \u3054\u7d39\u4ecb\u7528\u306e\u30a2\u30d5\u30a3\u30ea\u30a8\u30a4\u30c8\u30ea\u30f3\u30af\u306f\u4ee5\u4e0b\u3068\u306a\u308a\u307e\u3059\u3002\u5fc5\u305a\u3053\u3061\u3089\u3092\u3054\u6d3b\u7528\u304f\u3060\u3055\u3044\u307e\u305b\uff01\n" +
      "\ud83d\udd17 " + cfg.AFFILIATE_INFO_TEXT + "\n\n" +
      "\u5f15\u304d\u7d9a\u304d\u3088\u308d\u3057\u304f\u304a\u9858\u3044\u3044\u305f\u3057\u307e\u3059\uff01";

    pushLineMessage(userId, remindText);

    sheet.getRange(i + 1, PENDING_COL.remindSent).setValue(
      "\u9001\u4fe1\u6e08 (" + Utilities.formatDate(today, "Asia/Tokyo", "MM/dd") + ")"
    );

    Logger.log("Reminder sent: row " + (i + 1) + " / " + productName);
  }
}

// =============================================
// \u4f4f\u6240\u30b7\u30fc\u30c8\u3092\u691c\u7d22\uff08\u52d5\u7684\u30d8\u30c3\u30c0\u30fc\u691c\u7d22\uff09
// =============================================

/**
 * 住所シートから一致する住所を「全件」収集して返す
 * @returns {{ isFound: boolean, address?: string, multiple?: boolean, candidates?: {label:string, address:string}[] }}
 */
function lookupAddress(ss, creatorId, realName, addressConfirmed) {
  if (!addressConfirmed) {
    return { isFound: false, address: "住所変更の申し出あり（個別チャット対応）" };
  }

  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_ADDRESS);
  if (!sheet) sheet = ss.getSheetByName("住所管理シート");
  if (!sheet) return { isFound: false, address: "住所シートが見つかりません" };

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { isFound: false, address: "住所データなし" };

  const headerRow = values[0];
  let idCol = -1, nameCol = -1, zipCol = -1, addrCol = -1;

  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] || "").replace(/[\s\u3000]/g, "").toLowerCase();
    if      (idCol   === -1 && (h === "id" || h.includes("id") || h.includes("アカウント"))) idCol   = c;
    else if (nameCol === -1 && (h === "名前" || h === "氏名" || h.includes("クリエイター名"))) nameCol = c;
    else if (zipCol  === -1 && (h === "郵便番号" || h.includes("〒") || h.includes("郵便")))  zipCol  = c;
    else if (addrCol === -1 && (h === "住所" || h.includes("お届け先")))                       addrCol = c;
  }

  if (idCol   === -1) idCol   = 3;
  if (nameCol === -1) nameCol = 5;
  if (zipCol  === -1) zipCol  = 6;
  if (addrCol === -1) addrCol = 7;

  const targetId   = String(creatorId || "").trim().toLowerCase().replace(/^@/, "");
  const targetName = String(realName  || "").trim().replace(/[\s\u3000]/g, "");

  const candidates = [];

  for (let i = 1; i < values.length; i++) {
    const rowId   = String(values[i][idCol]   || "").trim().toLowerCase().replace(/^@/, "");
    const rowName = String(values[i][nameCol] || "").trim().replace(/[\s\u3000]/g, "");
    const matched = (targetId   && rowId   && rowId   === targetId) ||
                    (targetName && rowName && rowName === targetName);
    if (!matched) continue;

    const rawZip  = String(values[i][zipCol]  || "").trim();
    const rawAddr = String(values[i][addrCol] || "").trim();
    if (!rawAddr && !rawZip) continue; // 住所が空の行はスキップ

    let fullAddress = rawAddr;
    if (rawZip && !rawAddr.includes(rawZip)) {
      fullAddress = (rawZip.startsWith("〒") ? rawZip : "〒" + rawZip) + " " + rawAddr;
    }
    // ラベル（表示名）はシート行番号付きで区別できるようにする
    const displayName = String(values[i][nameCol] || "").trim() || creatorId || "不明";
    candidates.push({ label: displayName + "（" + (i + 1) + "行目）", address: fullAddress.trim() });
  }

  if (candidates.length === 0) {
    return { isFound: false, address: "シート未登録（個人チャットで確認）" };
  }
  if (candidates.length === 1) {
    return { isFound: true, address: candidates[0].address };
  }
  // 複数ヒット → 選択が必要
  return { isFound: true, multiple: true, candidates: candidates };
}

// =============================================
// 店舗マスタから店舗情報を取得
// =============================================

/**
 * ブランド名（承認待ちシートC列）から「店舗発送依頼先情報」タブのJ列を検索し、
 * 「発送シート」列（F列）にあるスプレッドシートURL等の店舗情報を取得する。
 * J列は "A\B\C" / "A\\B\\C" のようなバックスラッシュ区切りの複数ブランド表記に対応。
 */
function resolveShopInfoByName(ss, brandName) {
  if (!brandName || !String(brandName).trim()) return null;
  const target = String(brandName).trim();
  const cfg = getConfig();
  let masterSheet = ss.getSheetByName(cfg.SHEET_NAME_SHOP_MASTER) || ss.getSheetByName("\u5e97\u8217\u30de\u30b9\u30bf");
  if (!masterSheet) return null;

  const lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) return null;

  // 1行目のヘッダーから「発送シート」列を特定（見つからなければデフォルトF列=index 5）
  const maxCols = Math.max(masterSheet.getLastColumn(), 10);
  const values  = masterSheet.getRange(1, 1, lastRow, maxCols).getValues();
  const header  = values[0] || [];

  let urlColIndex = -1;
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] || "").replace(/[\s\u3000]/g, "");
    if (h === "発送シート") {
      urlColIndex = c;
      break;
    }
  }
  if (urlColIndex === -1) {
    for (let c = 0; c < header.length; c++) {
      const h = String(header[c] || "").replace(/[\s\u3000]/g, "");
      if (h.includes("発送シート")) {
        urlColIndex = c;
        break;
      }
    }
  }
  if (urlColIndex === -1) urlColIndex = 5; // デフォルト F列（0始まりで5）

  const brandColIndex = 9; // J列（0始まりで9）
  const normalize = s => String(s || "").replace(/[\s\u3000]/g, "").toLowerCase();
  const targetNorm = normalize(target);

  for (let i = 1; i < values.length; i++) {
    const rawBrands = String(values[i][brandColIndex] || "").trim();
    if (!rawBrands) continue;

    // "\" / "\\" / "￥" / "＼" で分割して比較
    const tokens = rawBrands.split(/[\\￥＼]+/).map(s => s.trim()).filter(Boolean);
    const isMatched = tokens.some(t => normalize(t) === targetNorm);

    if (isMatched) {
      const shopSheetUrl = extractUrlFromCell(masterSheet.getRange(i + 1, urlColIndex + 1));
      const rawGroup = String(values[i][1] || "").trim(); // B列: 依頼グループ
      const notifyGroupId = /^[CR][0-9a-f]{32}$/i.test(rawGroup) ? rawGroup : "";
      return {
        shopName: target,
        sheetShopName: String(values[i][0] || "").trim(),
        shopSheetUrl: shopSheetUrl,
        notifyGroupId: notifyGroupId
      };
    }
  }
  return null;
}

/**
 * 登録されているブランド（店舗名）一覧を取得
 */
function getBrandList() {
  const cleanFallback = [
    "\u6728\u6751\u6d77\u7523", "\u4e00\u67f3", "\u30a2\u30ec\u30b0\u30ec\u30c3\u30c8", "welleg",
    "\u6176\u6d0b\u30a8\u30f3\u30b8\u30cb\u30a2", "\u30e9\u30c7\u30a3\u30a2\u30f3\u30cc",
    "\u3061\u306c\u3084", "\u3057\u304a\u305d\u3046", "\u5e3d\u5b50\u5c4b", "vitaeflice"
  ];

  try {
    const ss = getSpreadsheet();
    const cfg = getConfig();
    const masterSheet = ss.getSheetByName(cfg.SHEET_NAME_SHOP_MASTER) || ss.getSheetByName("\u5e97\u8217\u30de\u30b9\u30bf");
    if (!masterSheet) return cleanFallback;

    const lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) return cleanFallback;

    const values = masterSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const list = [];
    values.forEach(r => {
      const name = String(r[0] || "").trim();
      if (name && !list.includes(name)) list.push(name);
    });

    return list.length > 0 ? list : cleanFallback;
  } catch (e) {
    Logger.log("getBrandList error: " + e.toString());
    return cleanFallback;
  }
}

function resolveShopInfo(ss, groupId) {
  let shopName = "\u5e97\u8217\uff08\u672a\u7279\u5b9a\uff09", notifyGroupId = groupId || "", shopSheetUrl = null;

  const groupName = groupId ? fetchLineGroupName(groupId) : null;
  if (groupName) shopName = groupName;

  const cfg = getConfig();
  let masterSheet = ss.getSheetByName(cfg.SHEET_NAME_SHOP_MASTER) || ss.getSheetByName("\u5e97\u8217\u30de\u30b9\u30bf");

  if (masterSheet) {
    const lastRow = masterSheet.getLastRow();
    if (lastRow > 1) {
      const values = masterSheet.getRange(1, 1, lastRow, Math.max(masterSheet.getLastColumn(), 7)).getValues();
      for (let i = 1; i < values.length; i++) {
        const rowShop      = String(values[i][0] || "").trim();
        const rowGroupName = String(values[i][2] || "").trim();
        const isMatch = groupName && (rowGroupName === groupName || groupName.includes(rowGroupName) || (rowShop && groupName.includes(rowShop)))
                        || values[i][0] === groupId;
        if (isMatch) {
          if (rowShop) shopName = rowShop;
          shopSheetUrl = extractUrlFromCell(masterSheet.getRange(i + 1, 6));
          break;
        }
      }
    }
  }

  return { shopName, notifyGroupId, shopSheetUrl };
}

function extractUrlFromCell(cellRange) {
  if (!cellRange) return null;
  try { const link = cellRange.getRichTextValue()?.getLinkUrl(); if (link) return link; } catch (e) {}
  try { const m = cellRange.getFormula().match(/HYPERLINK\(\s*"([^"]+)"/i); if (m) return m[1]; } catch (e) {}
  const val = String(cellRange.getValue() || "").trim();
  return val.startsWith("http") ? val : null;
}

// =============================================
// 店舗外部発送シートへ転記
// =============================================

/**
 * 日時（Date または "yyyy/MM/dd HH:mm(:ss)" 等）を "M/d"（例: "10/10"）に変換する。
 * 文字列に含まれる月・日を直接抽出し、タイムゾーンによる日付ずれを防ぐ。
 */
function formatShortDateMD(dateInput) {
  if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
    if (typeof Utilities !== "undefined" && Utilities.formatDate) {
      return Utilities.formatDate(dateInput, "Asia/Tokyo", "M/d");
    }
    return (dateInput.getMonth() + 1) + "/" + dateInput.getDate();
  }

  const str = String(dateInput || "").trim();
  if (str) {
    const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[\sT]|$)/);
    if (ymd) {
      return Number(ymd[2]) + "/" + Number(ymd[3]);
    }
    const md = str.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\sT]|$)/);
    if (md) {
      return Number(md[1]) + "/" + Number(md[2]);
    }
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      if (typeof Utilities !== "undefined" && Utilities.formatDate) {
        return Utilities.formatDate(parsed, "Asia/Tokyo", "M/d");
      }
      return (parsed.getMonth() + 1) + "/" + parsed.getDate();
    }
  }

  const fallback = new Date();
  if (typeof Utilities !== "undefined" && Utilities.formatDate) {
    return Utilities.formatDate(fallback, "Asia/Tokyo", "M/d");
  }
  return (fallback.getMonth() + 1) + "/" + fallback.getDate();
}

/**
 * 転記先シートの1行目ヘッダーから各項目に対応する列インデックス（0始まり）を判定する。
 * 判定順を明確にし、「商品URL」が「商品名」に、「お届け先氏名」が「住所」に誤マッチするのを防ぐ。
 */
function detectShopSheetColumns(headerRow) {
  const cols = {
    dateCol: -1, creatorCol: -1, nameCol: -1, addrCol: -1,
    prodCol: -1, urlCol: -1, varCol: -1, detailCol: -1
  };
  if (!headerRow || !headerRow.length) return cols;

  let fallbackDateCol = -1;
  headerRow.forEach((h, c) => {
    const hh = String(h || "").replace(/[\s\u3000]/g, "").toLowerCase();
    if (!hh) return;

    // 具体的な依頼・承認・受付系の日付列を優先し、単独の「日」は「発送日」等に奪われないよう後順位にする
    if (cols.dateCol === -1 && (
      hh.includes("依頼日") || hh.includes("申請日") || hh.includes("承認日") ||
      hh.includes("受付日") || hh.includes("日付") || hh.includes("date")
    )) {
      cols.dateCol = c;
      return;
    }
    if (cols.urlCol === -1 && (hh.includes("url") || hh.includes("リンク"))) {
      cols.urlCol = c;
      return;
    }
    if (cols.detailCol === -1 && (hh.includes("備考") || hh.includes("詳細") || hh.includes("要望"))) {
      cols.detailCol = c;
      return;
    }
    if (cols.varCol === -1 && (hh.includes("バリエ") || hh.includes("カラー") || hh.includes("サイズ") || hh.includes("規格"))) {
      cols.varCol = c;
      return;
    }
    if (cols.prodCol === -1 && (hh.includes("商品名") || hh.includes("商品"))) {
      cols.prodCol = c;
      return;
    }
    if (cols.creatorCol === -1 &&
        !hh.includes("氏名") && !hh.includes("名前") && !hh.includes("宛名") &&
        (hh.includes("クリエイター") || hh.includes("アカウント") || hh.includes("tiktok") || hh === "id")) {
      cols.creatorCol = c;
      return;
    }
    if (cols.nameCol === -1 && (hh.includes("氏名") || hh.includes("名前") || hh.includes("宛名"))) {
      cols.nameCol = c;
      return;
    }
    if (cols.addrCol === -1 && (hh.includes("住所") || hh.includes("お届け先") || hh.includes("届け先"))) {
      cols.addrCol = c;
      return;
    }
    if (fallbackDateCol === -1 && hh.includes("日")) {
      fallbackDateCol = c;
    }
  });

  if (cols.dateCol === -1 && fallbackDateCol !== -1) {
    cols.dateCol = fallbackDateCol;
  }
  return cols;
}

function appendToShopExternalSheet(shopSheetUrl, data) {
  if (!shopSheetUrl) return null;
  try {
    const shopSS = SpreadsheetApp.openByUrl(shopSheetUrl);
    const sheets = shopSS.getSheets();
    if (!sheets || sheets.length === 0) return null;

    // タブが複数ある場合、「発送依頼シート」を優先的に探す
    let targetSheet = null;
    if (sheets.length > 1) {
      // 1. 完全一致
      targetSheet = shopSS.getSheetByName("発送依頼シート");
      // 2. 部分一致（「発送依頼」や「発送」を含むタブ）
      if (!targetSheet) {
        for (let s = 0; s < sheets.length; s++) {
          const sName = sheets[s].getName().replace(/[\s\u3000]/g, "");
          if (sName.includes("発送依頼") || sName.includes("発送")) {
            targetSheet = sheets[s];
            break;
          }
        }
      }
    }
    // 見つからない場合やタブが1つだけの場合は先頭のタブを使用
    if (!targetSheet) {
      targetSheet = sheets[0];
    }

    const values   = targetSheet.getDataRange().getValues();
    const dateStr  = formatShortDateMD(data.approvedAt);
    const startRow = targetSheet.getLastRow() + 1;

    let products = data.products;
    if (!products || !Array.isArray(products) || products.length === 0) {
      products = [{ name: data.productName || "指定なし", url: "", variation: "-", details: "-" }];
    }

    const cols = values.length >= 1 ? detectShopSheetColumns(values[0]) : detectShopSheetColumns([]);
    const hasMap = Object.keys(cols).some(k => cols[k] !== -1);
    const maxMatchedCol = Math.max(
      cols.dateCol, cols.creatorCol, cols.nameCol, cols.addrCol,
      cols.prodCol, cols.urlCol, cols.varCol, cols.detailCol
    );
    const rowLen = Math.max(targetSheet.getLastColumn(), maxMatchedCol + 1, 1);

    products.forEach(p => {
      if (hasMap) {
        const row = new Array(rowLen).fill("");
        if (cols.dateCol    !== -1) row[cols.dateCol]    = dateStr;
        if (cols.creatorCol !== -1) row[cols.creatorCol] = data.creatorId || "";
        if (cols.nameCol    !== -1) row[cols.nameCol]    = data.realName  || "";
        if (cols.addrCol    !== -1) row[cols.addrCol]    = data.address   || "";
        if (cols.prodCol    !== -1) row[cols.prodCol]    = p.name         || "";
        if (cols.urlCol     !== -1) row[cols.urlCol]     = p.url          || "-";
        if (cols.varCol     !== -1) row[cols.varCol]     = p.variation    || "-";
        if (cols.detailCol  !== -1) row[cols.detailCol]  = p.details      || "-";
        targetSheet.appendRow(row);
      } else {
        targetSheet.appendRow([dateStr, data.creatorId || "", data.realName || "", data.address || "", p.name || "", p.url || "-", p.variation || "-", p.details || "-"]);
      }
    });

    const endRow = targetSheet.getLastRow();
    // 転記先の列が元々yyyy/MM/dd形式等にフォーマットされていても「10/10」形式で表示されるようにする
    const writtenDateCol = hasMap ? cols.dateCol : 0;
    if (writtenDateCol !== -1 && endRow >= startRow) {
      try {
        targetSheet.getRange(startRow, writtenDateCol + 1, endRow - startRow + 1, 1).setNumberFormat("M/d");
      } catch (fmtErr) {
        Logger.log("setNumberFormat M/d skip: " + fmtErr.toString());
      }
    }

    const rowText = startRow === endRow ? startRow + "行目" : startRow + "〜" + endRow + "行目";
    return { success: true, rowText, sheetTitle: targetSheet.getName() };

  } catch (err) {
    Logger.log("Shop sheet error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

// =============================================
// \u5e97\u8217LINE\u30b0\u30eb\u30fc\u30d7\u3078\u767a\u9001\u4f9d\u983c\u901a\u77e5
// =============================================

function notifyShopGroup(data, shopRowText) {
  const targetChatId = data.shopNotifyGroupId;
  if (!targetChatId) { Logger.log("No group ID, skip"); return; }

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{ name: data.productName || "\u6307\u5b9a\u306a\u3057", url: "", variation: "", details: "" }];
  }

  let productText = "";
  if (products.length === 1) {
    const p = products[0];
    productText = "\u25a0 \u5546\u54c1: " + p.name;
    if (p.url)       productText += "\n\u25a0 \u30ea\u30f3\u30af: " + p.url;
    if (p.variation) productText += "\n\u25a0 \u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3: " + p.variation;
    if (p.details)   productText += "\n\u25a0 \u8a73\u7d30\u30fb\u3054\u8981\u671b: " + p.details;
  } else {
    productText = "\u25a0 \u7533\u8acb\u5546\u54c1 (\u8a08" + products.length + "\u70b9):\n" + products.map((p, i) => {
      let s = "\u30fb\u3010\u5546\u54c1" + (i + 1) + "\u3011" + p.name;
      if (p.variation) s += "\n  \u2514 \u30d0\u30ea\u30a8\u30fc\u30b7\u30e7\u30f3: " + p.variation;
      if (p.url)       s += "\n  \u2514 \u30ea\u30f3\u30af: " + p.url;
      if (p.details)   s += "\n  \u2514 \u8a73\u7d30\u30fb\u5099\u8003: " + p.details;
      return s;
    }).join("\n");
  }

  const shopSheetPart = shopRowText ? ("\n\u25a0 \u5e97\u8217\u69d8\u767a\u9001\u30b7\u30fc\u30c8: \u3010" + shopRowText + "\u3011\u306b\u81ea\u52d5\u8a18\u5165\u6e08") : "";

  const msg =
    "\u304a\u75b2\u308c\u69d8\u3067\u3059\uff01\n" +
    "\u5546\u54c1\u30b5\u30f3\u30d7\u30eb\u306e\u767a\u9001\u4f9d\u983c\u3092\u304a\u9858\u3044\u3044\u305f\u3057\u307e\u3059\uff01\n\n" +
    "\u25a0 \u5e97\u8217\u69d8: " + data.shopName + "\n" +
    "\u25a0 \u30af\u30ea\u30a8\u30a4\u30bf\u30fc: " + data.creatorId + " (" + data.realName + "\u69d8)\n\n" +
    productText + "\n\n" +
    "\u25a0 \u304a\u5c4a\u3051\u5148\u60c5\u5831:\n" +
    "\u30fb\u304a\u540d\u524d: " + data.realName + " \u69d8\n" +
    "\u30fb\u3054\u4f4f\u6240: " + (data.address || "\u78ba\u8a8d\u4e2d") +
    shopSheetPart + "\n\n" +
    "\u3054\u78ba\u8a8d\u306e\u4e0a\u3001\u767a\u9001\u624b\u914d\u3092\u304a\u9858\u3044\u3044\u305f\u3057\u307e\u3059\ud83d\ude47\u200d\u2642\ufe0f";

  pushLineMessage(targetChatId, msg);
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
  } catch (e) { Logger.log("Group name error: " + e.toString()); }
  return null;
}

function pushLineMessage(toId, text) {
  const token = getConfig().LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) {
    Logger.log("pushLineMessage SKIPPED: LINE_CHANNEL_TOKEN is empty");
    return;
  }
  try {
    const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
      payload: JSON.stringify({ to: toId, messages: [{ type: "text", text: text }] }),
      muteHttpExceptions: true
    });
    Logger.log("pushLineMessage to " + toId + " -> status: " + res.getResponseCode() + " / body: " + res.getContentText());
  } catch (e) { Logger.log("Push error: " + e.toString()); }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// カスタムメニュー（スプレッドシートを開いたとき）
// =============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("発送システム")
    .addItem("✅ 選択行を承認する", "approveSelectedRow")
    .addSeparator()
    .addItem("承認待ちシートを整形", "formatPendingSheet")
    .addToUi();
}

// =============================================
// メニューから承認する（住所1件 or 複数で分岐）
// =============================================

/**
 * 「発送システム」→「選択行を承認する」から呼ばれる。
 * 承認待ちシートで選択中の行を対象に住所を検索し、
 * 1件なら確認ダイアログ → そのまま処理。
 * 複数件なら住所選択ポップアップを表示。
 */
function approveSelectedRow() {
  const ui  = SpreadsheetApp.getUi();
  const ss  = getSpreadsheet();
  const cfg = getConfig();

  // ① 承認待ちシートが開かれているか確認
  const activeSheet = SpreadsheetApp.getActiveSheet();
  if (activeSheet.getName() !== cfg.SHEET_NAME_PENDING) {
    ui.alert("エラー", "承認待ちシートを開いた状態で実行してください。", ui.ButtonSet.OK);
    return;
  }

  // ② 選択行のバリデーション
  const row = SpreadsheetApp.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert("エラー", "ヘッダー行は承認できません。データ行を選択してください。", ui.ButtonSet.OK);
    return;
  }

  const rowData   = activeSheet.getRange(row, 1, 1, 18).getValues()[0];
  const requestId = String(rowData[PENDING_COL.reqId    - 1] || "");
  const brandName = String(rowData[PENDING_COL.shopName - 1] || "").trim();
  const status    = String(rowData[PENDING_COL.status   - 1] || "");
  const creatorId = String(rowData[PENDING_COL.creatorId - 1] || "");
  const realName  = String(rowData[PENDING_COL.realName  - 1] || "");
  const addrCheck = String(rowData[PENDING_COL.addrCheck - 1] || "");

  if (!requestId) {
    ui.alert("エラー", "この行にはデータがありません。", ui.ButtonSet.OK);
    return;
  }
  if (status === "転記済") {
    ui.alert("確認", "この行はすでに転記済みです。", ui.ButtonSet.OK);
    return;
  }
  if (status === "処理中...") {
    ui.alert("確認", "この行は現在処理中です。しばらく待ってから再試行してください。", ui.ButtonSet.OK);
    return;
  }

  // ③ 「店舗発送依頼先情報」タブのJ列→F列（発送シート）から転記先が解決できるか事前確認
  const previewShop = resolveShopInfoByName(ss, brandName);
  if (!previewShop || !previewShop.shopSheetUrl) {
    ui.alert(
      "発送シートが見つかりません",
      "ブランド「" + (brandName || "未入力") + "」に対応する発送シートリンクが見つかりませんでした。\n\n" +
      "「" + cfg.SHEET_NAME_SHOP_MASTER + "」タブのJ列（ブランド名）とF列（発送シート）のリンクを確認してください。",
      ui.ButtonSet.OK
    );
    return;
  }

  // ④ 住所シートを検索
  const addressConfirmed = addrCheck.includes("変更なし");
  const addressResult    = lookupAddress(ss, creatorId, realName, addressConfirmed);

  if (!addressResult.isFound) {
    ui.alert("住所が見つかりません", addressResult.address + "\n\n住所シートを確認するか、個別チャットで住所を取得してください。", ui.ButtonSet.OK);
    return;
  }

  // ⑤-A 住所が1件 → 確認ダイアログを出してそのまま処理
  if (!addressResult.multiple) {
    const shopLabel = previewShop.sheetShopName
      ? brandName + "（発送先: " + previewShop.sheetShopName + "）"
      : brandName;
    const ans = ui.alert(
      "承認の確認",
      "以下の内容で転記・LINE通知を行います。よろしいですか？\n\n" +
      "■ ブランド: " + shopLabel + "\n" +
      "■ お届け先: " + addressResult.address,
      ui.ButtonSet.YES_NO
    );
    if (ans !== ui.Button.YES) return;

    const result = executeApproval(ss, activeSheet, requestId, addressResult.address);
    ui.alert(result.success ? "✅ 完了" : "❌ エラー", result.message, ui.ButtonSet.OK);
    return;
  }

  // ⑤-B 住所が複数 → 選択ポップアップを表示
  showAddressPickerDialog(requestId, addressResult.candidates);
}

// =============================================
// 住所選択ポップアップ（複数候補ヒット時）
// =============================================

/** HTMLエスケープ用ヘルパー */
function htmlEscape(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * 住所候補をラジオボタンで選ばせるモーダルダイアログを表示する。
 * 選択後にダイアログ内のJSが google.script.run.doFinalApproval() を呼ぶ。
 */
function showAddressPickerDialog(requestId, candidates) {
  // 候補データをHTMLに直接埋め込む（ScriptPropertiesなし）
  const candidatesJson = JSON.stringify(candidates);

  const optionsHtml = candidates.map((c, i) =>
    '<label style="display:block;margin:10px 0;cursor:pointer;padding:8px 10px;border:1px solid #e2e8f0;border-radius:8px;" ' +
    'onmouseover="this.style.background=\'#eff6ff\'" onmouseout="this.style.background=\'#fff\'">' +
    '<input type="radio" name="addr" value="' + i + '" style="margin-right:8px;">' +
    '<strong style="font-size:13px;">' + htmlEscape(c.label) + '</strong><br>' +
    '<span style="color:#555;font-size:12px;padding-left:20px;display:block;margin-top:2px;">' + htmlEscape(c.address) + '</span>' +
    '</label>'
  ).join("");

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:-apple-system,sans-serif;font-size:14px;padding:16px;margin:0;color:#1e293b;}' +
    'h3{margin:0 0 8px;font-size:15px;color:#1e3a8a;}' +
    'p{margin:0 0 12px;color:#475569;font-size:13px;}' +
    '.btns{margin-top:16px;display:flex;gap:8px;}' +
    'button{padding:9px 20px;border-radius:8px;border:none;cursor:pointer;font-size:14px;font-weight:600;}' +
    '.ok{background:#1d4ed8;color:#fff;} .ok:hover{background:#1e40af;}' +
    '.cancel{background:#e2e8f0;color:#334155;} .cancel:hover{background:#cbd5e1;}' +
    '</style></head><body>' +
    '<h3>🏠 発送先の住所を選択してください</h3>' +
    '<p>この申請者の住所が' + candidates.length + '件登録されています。<br>どちらに発送しますか？</p>' +
    '<div id="opts">' + optionsHtml + '</div>' +
    '<div class="btns">' +
    '<button class="ok" onclick="submit()">この住所で承認する</button>' +
    '<button class="cancel" onclick="google.script.host.close()">キャンセル</button>' +
    '</div>' +
    '<p id="msg" style="color:#ef4444;font-size:12px;margin-top:8px;"></p>' +
    '<script>' +
    'var CANDIDATES = ' + candidatesJson + ';' +
    'function submit() {' +
    '  var sel = document.querySelector("input[name=addr]:checked");' +
    '  if (!sel) { document.getElementById("msg").textContent = "住所を選択してください。"; return; }' +
    '  var idx = parseInt(sel.value);' +
    '  var addr = CANDIDATES[idx].address;' +
    '  document.querySelectorAll("button").forEach(function(b){b.disabled=true;});' +
    '  document.getElementById("msg").style.color = "#64748b";' +
    '  document.getElementById("msg").textContent = "処理中...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(res){' +
    '      alert(res.message);' +
    '      google.script.host.close();' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      alert("エラー: " + e.message);' +
    '      document.querySelectorAll("button").forEach(function(b){b.disabled=false;});' +
    '      document.getElementById("msg").textContent = "";' +
    '    })' +
    '    .doFinalApproval("' + requestId + '", addr);' +
    '}' +
    '<\/script>' +
    '</body></html>'
  ).setWidth(500).setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, "住所を選択して承認");
}

/**
 * 住所選択ポップアップから呼ばれる。選択した住所文字列を直接受け取って承認処理を実行する。
 * @param {string} requestId
 * @param {string} selectedAddress  ユーザーが選んだ住所の文字列
 * @returns {{ success: boolean, message: string }}
 */
function doFinalApproval(requestId, selectedAddress) {
  const ss          = getSpreadsheet();
  const cfg         = getConfig();
  const pendingSheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  if (!pendingSheet) return { success: false, message: "エラー: 承認待ちシートが見つかりません。" };
  return executeApproval(ss, pendingSheet, requestId, selectedAddress);
}

