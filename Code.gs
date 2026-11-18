/**
 * TikToker商品サンプル発送依頼 & リマインド自動化システム
 * Google Apps Script (GAS) バックエンド
 * v2.0 - 承認制ワークフロー対応
 */

// ====================================================================
// 設定はGASの「スクリプトプロパティ」に登録してください
// (プロジェクトの設定 > スクリプトプロパティ > 行を追加)
//
// プロパティ名             値
// SPREADSHEET_ID       スプレッドシートのID
// LIFF_ID              LINE DevelopersのLIFF ID
// LINE_CHANNEL_TOKEN   LINEチャネルアクセストークン（長期）
// AFFILIATE_URL        自社アフィリエイトリンク
// ====================================================================
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID:            props.getProperty("SPREADSHEET_ID") || "10v-c76tg20i5sSzIM6T7-QQ7kTNLPiNOZhvF0A8NDVg",
    LINE_CHANNEL_ACCESS_TOKEN: props.getProperty("LINE_CHANNEL_TOKEN") || "",
    SHEET_NAME_PENDING:        props.getProperty("SHEET_NAME_PENDING")     || "承認待ちシート",
    SHEET_NAME_CREATOR_DB:     props.getProperty("SHEET_NAME_CREATOR_DB")  || "クリエイターDB",
    SHEET_NAME_SHIPPING:       props.getProperty("SHEET_NAME_SHIPPING")    || "配送シート",
    SHEET_NAME_ADDRESS:        props.getProperty("SHEET_NAME_ADDRESS")     || "クリエイター・代理店サンプルリスト",
    SHEET_NAME_SHOP_MASTER:    props.getProperty("SHEET_NAME_SHOP_MASTER") || "店舗発送依頼先情報",
    AFFILIATE_INFO_TEXT:       props.getProperty("AFFILIATE_URL") || "",
    APPROVAL_COL: 13   // M列: チェックボックス承認列（1始まり）
  };
}

/**
 * 対象スプレッドシートを取得
 */
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
// LIFFフォーム表示・送信処理
// =============================================

/**
 * GETリクエスト: LIFFフォーム画面の表示
 */
function doGet(e) {
  const liffId = PropertiesService.getScriptProperties().getProperty("LIFF_ID") || "";
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setTitle("商品サンプル発送・紹介受付フォーム")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.setContent(html.getContent().replace("__LIFF_ID__", liffId));
  return html;
}

/**
 * LIFFフォームから直接呼び出す処理 (google.script.run 対応)
 */
function submitShippingRequest(data) {
  return processRequestData(data);
}

/**
 * POSTリクエスト処理 (外部fetchからの送信用)
 */
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
// フォーム送信処理（クリエイターDB + 承認待ちシートに書くだけで止まる）
// =============================================

/**
 * 発送リクエスト共通処理
 * フォーム送信 → クリエイターDB（永久保存）+ 承認待ちシートに書き込んで止まる
 * 転記処理は担当者の承認（チェックボックス）後に行う
 */
function processRequestData(data) {
  try {
    const ss = getSpreadsheet();

    // 1. 店舗情報を取得（グループIDから）
    const shopInfo = resolveShopInfo(ss, data.groupId);
    data.shopName          = shopInfo.shopName;
    data.shopNotifyGroupId = shopInfo.notifyGroupId;
    data.shopSheetUrl      = shopInfo.shopSheetUrl;

    // 管理番号を生成
    const requestId = "REQ-" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd-HHmmss")
                      + "-" + Math.floor(Math.random() * 1000);
    data.requestId = requestId;

    // 2. クリエイターDBシートに永久保存（システムが失敗してもデータが消えない）
    appendCreatorDB(ss, data);

    // 3. 承認待ちシートに書き込む（ここで止まる）
    appendPendingRecords(ss, data);

    return {
      success: true,
      shopName: data.shopName,
      action: "pending_approval",
      message: "申請を受け付けました！担当者が確認後、発送手配いたします。"
    };

  } catch (err) {
    Logger.log("Process Error: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

// =============================================
// 【インストール可能トリガー】
// チェックボックスをONにした瞬間に自動起動
// =============================================

/**
 * 承認チェックボックス編集時に即時起動するトリガー関数
 * setupApprovalTrigger() で1回だけ登録する
 */
function onApprovalEdit(e) {
  try {
    const cfg   = getConfig();
    const range = e.range;
    const sheet = range.getSheet();

    // 対象シートかチェック
    if (sheet.getName() !== cfg.SHEET_NAME_PENDING) return;

    // 承認チェックボックス列（M列=13）かチェック
    if (range.getColumn() !== cfg.APPROVAL_COL) return;

    // ヘッダー行は除外
    if (range.getRow() <= 1) return;

    // チェックがONになった場合のみ処理
    if (e.value !== "TRUE") return;

    const row = range.getRow();
    const ss  = getSpreadsheet();

    // 同じ管理番号の行をまとめて処理（同一申請の複数商品対応）
    const requestId = sheet.getRange(row, 2).getValue();
    if (!requestId) return;

    processApprovedRequest(ss, sheet, requestId);

  } catch (err) {
    Logger.log("onApprovalEdit Error: " + err.toString());
  }
}

/**
 * 承認された申請の転記・通知処理
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Sheet} pendingSheet - 承認待ちシート
 * @param {string} requestId - 管理番号
 */
function processApprovedRequest(ss, pendingSheet, requestId) {
  const cfg = getConfig();

  // LockServiceで二重実行防止
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log("ロック取得失敗（他の処理が実行中）: " + e.toString());
    return;
  }

  try {
    const allData = pendingSheet.getDataRange().getValues();

    // 承認待ちシートの列定義（0始まり）
    // A=0:受付日時 B=1:管理番号 C=2:店舗名 D=3:クリエイターID E=4:氏名 F=5:住所確認
    // G=6:商品名 H=7:商品リンク I=8:希望バリエ J=9:商品詳細
    // K=10:LINE_USER_ID L=11:LINE_GROUP_ID M=12:✅承認 N=13:承認日時 O=14:ステータス P=15:発送シートURL
    const COL = {
      requestId:    1,
      shopName:     2,
      creatorId:    3,
      realName:     4,
      addrCheck:    5,
      prodName:     6,
      prodUrl:      7,
      prodVar:      8,
      prodDetail:   9,
      userId:      10,
      groupId:     11,
      approval:    12,
      approvedAt:  13,
      status:      14,
      shopUrl:     15
    };

    // 同じ管理番号の行を収集（未転記のもののみ）
    const targetRows = [];
    for (let i = 1; i < allData.length; i++) {
      const rowId = String(allData[i][COL.requestId] || "");
      const st    = String(allData[i][COL.status]    || "");
      if (rowId === requestId && st !== "転記済" && st !== "処理中") {
        targetRows.push({ rowIndex: i + 1, rowData: allData[i] });
      }
    }

    if (targetRows.length === 0) {
      Logger.log("対象行なし（既に処理済みかも）: " + requestId);
      return;
    }

    // まず全行を「処理中」にする（二重処理防止）
    targetRows.forEach(r => updatePendingStatus(pendingSheet, r.rowIndex, "処理中..."));

    // 最初の行からメタ情報を取得（全行で共通）
    const first = targetRows[0].rowData;
    const shopName     = String(first[COL.shopName]  || "");
    const creatorId    = String(first[COL.creatorId] || "");
    const realName     = String(first[COL.realName]  || "");
    const addrCheck    = String(first[COL.addrCheck] || "");
    const userId       = String(first[COL.userId]    || "");
    const groupId      = String(first[COL.groupId]   || "");
    const shopSheetUrl = String(first[COL.shopUrl]   || "");

    // 商品情報を収集
    const products = targetRows.map(r => ({
      name:      String(r.rowData[COL.prodName]   || ""),
      url:       String(r.rowData[COL.prodUrl]    || ""),
      variation: String(r.rowData[COL.prodVar]    || ""),
      details:   String(r.rowData[COL.prodDetail] || "")
    }));

    // 住所確認フラグ（"変更なし" を含む場合はtrue）
    const addressConfirmed = addrCheck.includes("変更なし");

    // 住所検索
    const addressResult = lookupAddress(ss, creatorId, realName, addressConfirmed);
    if (!addressResult.isFound) {
      const errMsg = "【要対応】" + addressResult.address;
      targetRows.forEach(r => {
        updatePendingStatus(pendingSheet, r.rowIndex, errMsg);
        pendingSheet.getRange(r.rowIndex, cfg.APPROVAL_COL).setValue(false); // チェックを戻す
      });
      Logger.log("住所なし: " + errMsg);
      return;
    }

    // 店舗発送シートURLチェック（なければエラー）
    if (!shopSheetUrl || shopSheetUrl.trim() === "" || shopSheetUrl === "null") {
      const errMsg = "【要対応】店舗発送シートURL未設定（F列確認）";
      targetRows.forEach(r => {
        updatePendingStatus(pendingSheet, r.rowIndex, errMsg);
        pendingSheet.getRange(r.rowIndex, cfg.APPROVAL_COL).setValue(false);
      });
      Logger.log("店舗シートURLなし: " + shopName);
      return;
    }

    // 処理用データオブジェクトを組み立て
    const data = {
      shopName:          shopName,
      shopNotifyGroupId: groupId,
      shopSheetUrl:      shopSheetUrl,
      creatorId:         creatorId,
      realName:          realName,
      address:           addressResult.address,
      userId:            userId,
      requestId:         requestId,
      products:          products
    };

    // 店舗外部発送シートへ転記
    const shopWriteRes = appendToShopExternalSheet(shopSheetUrl, data);
    if (!shopWriteRes || !shopWriteRes.success) {
      const errMsg = "【要対応】店舗シート書き込み失敗（権限/URL確認）";
      targetRows.forEach(r => {
        updatePendingStatus(pendingSheet, r.rowIndex, errMsg);
        pendingSheet.getRange(r.rowIndex, cfg.APPROVAL_COL).setValue(false);
      });
      Logger.log("店舗シート書き込み失敗: " + (shopWriteRes ? shopWriteRes.error : "不明"));
      return;
    }
    const shopRowText = shopWriteRes.rowText;

    // 自社配送シートへ転記
    const recordResult = appendShippingRecords(ss, data);
    const rowInfoText  = recordResult.rowText;

    // 店舗LINEグループへ通知
    notifyShopGroup(data, rowInfoText, shopRowText);

    // 承認日時を記録し、ステータスを「転記済」に更新
    const approvedAt = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    targetRows.forEach(r => {
      pendingSheet.getRange(r.rowIndex, 14).setValue(approvedAt); // N列: 承認日時
      updatePendingStatus(pendingSheet, r.rowIndex, "転記済");
    });

    Logger.log("承認処理完了: " + requestId + " / 配送シート:" + rowInfoText + " / 店舗シート:" + shopRowText);

  } finally {
    lock.releaseLock();
  }
}

// =============================================
// クリエイターDBシートに永久保存（バックアップ）
// =============================================

/**
 * クリエイターDBシートへの追記
 * フォーム送信のたびに記録し、絶対に削除しない
 * 列: A:受付日時 B:TikTokID C:LINE_USER_ID D:LINE_GROUP_ID E:氏名 F:商品名 G:ブランド H:備考
 */
function appendCreatorDB(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_CREATOR_DB);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_CREATOR_DB);
    sheet.appendRow([
      "受付日時", "TikTok ID", "LINE_USER_ID", "LINE_GROUP_ID",
      "氏名", "商品名", "ブランド（店舗名）", "備考（商品リンク）"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 8)
      .setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
  }

  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{
      name: data.productName || "指定なし",
      url:  data.productUrl  || ""
    }];
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
// 承認待ちシートへの書き込み（1商品1行）
// =============================================

/**
 * 承認待ちシートへの追記
 * 列定義:
 * A:受付日時 B:管理番号 C:店舗名 D:クリエイターID E:氏名 F:住所確認
 * G:商品名 H:商品リンク I:希望バリエーション J:商品詳細
 * K:LINE_USER_ID L:LINE_GROUP_ID M:✅承認 N:承認日時 O:ステータス P:発送シートURL
 */
function appendPendingRecords(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_PENDING);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_PENDING);
    sheet.appendRow([
      "受付日時", "管理番号", "店舗名", "クリエイターID", "氏名", "住所確認",
      "商品名", "商品リンク", "希望バリエーション", "商品詳細・ご要望",
      "LINE_USER_ID", "LINE_GROUP_ID", "✅承認", "承認日時", "ステータス", "発送シートURL"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 16)
      .setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
    // M列（承認チェックボックス）に範囲で設定（新規行も自動でチェックボックスになる）
    sheet.getRange("M2:M2000").setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build()
    );
    sheet.setColumnWidth(13, 60);  // M列を少し狭く
    sheet.setColumnWidth(16, 300); // P列（URL）を広く
  }

  const now       = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const checkText = data.addressConfirmed ? "変更なし" : "変更あり/初回";

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{
      name:      data.productName      || "指定なし",
      url:       data.productUrl       || "",
      variation: data.productVariation || data.productColor || "-",
      details:   data.productDetails   || "-"
    }];
  }

  products.forEach(p => {
    sheet.appendRow([
      now,                                              // A: 受付日時
      data.requestId,                                   // B: 管理番号
      data.shopName,                                    // C: 店舗名
      data.creatorId,                                   // D: クリエイターID
      data.realName,                                    // E: 氏名
      checkText,                                        // F: 住所確認
      p.name,                                           // G: 商品名
      p.url       || "-",                               // H: 商品リンク
      p.variation || "-",                               // I: 希望バリエーション
      p.details   || "-",                               // J: 商品詳細
      data.userId                           || "",      // K: LINE_USER_ID
      data.shopNotifyGroupId || data.groupId || "",     // L: LINE_GROUP_ID
      false,                                            // M: ✅承認（チェックボックス）
      "",                                               // N: 承認日時
      "受付中",                                         // O: ステータス
      data.shopSheetUrl                     || ""       // P: 発送シートURL
    ]);
  });
}

/**
 * 承認待ちシートのステータスセルを更新してハイライト
 */
function updatePendingStatus(sheet, rowIndex, statusText) {
  const statusCol = 15; // O列（1始まり）
  const cell = sheet.getRange(rowIndex, statusCol);
  cell.setValue(statusText);

  if (statusText === "転記済") {
    cell.setBackground("#dcfce7").setFontColor("#166534").setFontWeight("bold");
  } else if (statusText.includes("要対応") || statusText.includes("失敗")) {
    cell.setBackground("#fef08a").setFontColor("#854d0e").setFontWeight("bold");
  } else if (statusText === "処理中...") {
    cell.setBackground("#e0f2fe").setFontColor("#0369a1").setFontWeight("normal");
  } else {
    cell.setBackground("").setFontColor("").setFontWeight("normal");
  }
}

// =============================================
// インストール可能トリガーのセットアップ（1回だけ実行すればOK）
// =============================================

/**
 * 承認チェックボックス用トリガーを登録する
 * GASエディタの「実行」メニューから setupApprovalTrigger を1回だけ実行すること
 */
function setupApprovalTrigger() {
  // 既存の onApprovalEdit トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onApprovalEdit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 承認待ちシートの編集トリガーを設定
  const ss = getSpreadsheet();
  ScriptApp.newTrigger("onApprovalEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("トリガー設定完了: onApprovalEdit");
  SpreadsheetApp.getUi().alert(
    "✅ 承認トリガーの設定が完了しました！\n" +
    "これでチェックボックスをONにした瞬間に処理が自動で始まります。"
  );
}

// =============================================
// 住所シートを検索（動的ヘッダー検索）
// =============================================

/**
 * 住所管理シートをクリエイターIDまたは氏名で検索
 * 1行目のヘッダーから列を動的に検出（列位置変更に強い設計）
 */
function lookupAddress(ss, creatorId, realName, addressConfirmed) {
  if (!addressConfirmed) {
    return {
      isFound: false,
      address: "住所変更の申し出あり（個別チャット対応）"
    };
  }

  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_ADDRESS);
  if (!sheet) sheet = ss.getSheetByName("住所管理シート");
  if (!sheet) {
    return { isFound: false, address: "住所シート（" + cfg.SHEET_NAME_ADDRESS + "）が見つかりません" };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { isFound: false, address: "住所データなし" };
  }

  // ヘッダー行から列位置を動的に検出
  const headerRow = values[0];
  let idCol = -1, nameCol = -1, zipCol = -1, addrCol = -1;

  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] || "").replace(/[\s　]/g, "").toLowerCase();
    if (idCol === -1   && (h === "id" || h.includes("クリエイターid") || h.includes("アカウント"))) idCol = c;
    else if (nameCol === -1 && (h === "名前" || h === "氏名" || h.includes("クリエイター名"))) nameCol = c;
    else if (zipCol === -1  && (h === "郵便番号" || h.includes("〒") || h.includes("郵便"))) zipCol = c;
    else if (addrCol === -1 && (h === "住所" || h.includes("お届け先"))) addrCol = c;
  }

  // ヘッダーが見つからない場合のフォールバック（既知の列位置）
  if (idCol === -1)   idCol   = 3; // D列
  if (nameCol === -1) nameCol = 5; // F列
  if (zipCol === -1)  zipCol  = 6; // G列
  if (addrCol === -1) addrCol = 7; // H列

  const targetId   = String(creatorId || "").trim().toLowerCase().replace(/^@/, "");
  const targetName = String(realName  || "").trim().replace(/[\s　]/g, "");

  for (let i = 1; i < values.length; i++) {
    const rowId   = String(values[i][idCol]   || "").trim().toLowerCase().replace(/^@/, "");
    const rowName = String(values[i][nameCol] || "").trim().replace(/[\s　]/g, "");
    const isIdMatch   = targetId   && rowId   && (rowId   === targetId);
    const isNameMatch = targetName && rowName && (rowName === targetName);

    if (isIdMatch || isNameMatch) {
      const rawZip  = String(values[i][zipCol]  || "").trim();
      const rawAddr = String(values[i][addrCol] || "").trim();

      if (!rawAddr && !rawZip) {
        return { isFound: false, address: "シートに登録はあるが住所が空欄" };
      }

      let fullAddress = rawAddr;
      if (rawZip && !rawAddr.includes(rawZip)) {
        const fmtZip = rawZip.startsWith("〒") ? rawZip : ("〒" + rawZip);
        fullAddress  = fmtZip + " " + rawAddr;
      }
      return { isFound: true, address: fullAddress.trim() };
    }
  }

  return { isFound: false, address: "シート未登録（個人チャットで確認）" };
}

// =============================================
// 店舗マスタから店舗情報を取得
// =============================================

/**
 * 店舗発送依頼先情報シートから店舗名・通知グループID・発送シートURLを取得
 * 列: A:ショップ B:依頼グループ C:連絡グループ名 D:依頼方法 E:メンション先 F:発送シート G:備考
 */
function resolveShopInfo(ss, groupId) {
  let shopName      = "店舗（未特定）";
  let notifyGroupId = groupId || "";
  let shopSheetUrl  = null;

  const groupName = groupId ? fetchLineGroupName(groupId) : null;
  if (groupName) shopName = groupName;

  const cfg = getConfig();
  let masterSheet = ss.getSheetByName(cfg.SHEET_NAME_SHOP_MASTER);
  if (!masterSheet) masterSheet = ss.getSheetByName("店舗マスタ");

  if (masterSheet) {
    const lastRow = masterSheet.getLastRow();
    const lastCol = Math.max(masterSheet.getLastColumn(), 7);
    if (lastRow > 1) {
      const values = masterSheet.getRange(1, 1, lastRow, lastCol).getValues();
      for (let i = 1; i < values.length; i++) {
        const rowShop      = String(values[i][0] || "").trim(); // A: ショップ名
        const rowGroupName = String(values[i][2] || "").trim(); // C: 連絡グループ名

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
          // F列（列インデックス5）からURLを抽出（埋め込みリンク対応）
          const cell = masterSheet.getRange(i + 1, 6);
          shopSheetUrl = extractUrlFromCell(cell);
          break;
        }
      }
    }
  }

  return { shopName, notifyGroupId, shopSheetUrl };
}

/**
 * セルからURLを抽出（RichText埋め込みリンク / HYPERLINK数式 / 通常テキスト 対応）
 */
function extractUrlFromCell(cellRange) {
  if (!cellRange) return null;

  // 1. RichTextからリンク取得（画像のような埋め込みリンク）
  try {
    const richText = cellRange.getRichTextValue();
    if (richText) {
      const link = richText.getLinkUrl();
      if (link) return link;
    }
  } catch (e) {}

  // 2. =HYPERLINK() 数式から抽出
  try {
    const formula = cellRange.getFormula();
    if (formula) {
      const match = formula.match(/HYPERLINK\(\s*"([^"]+)"/i);
      if (match) return match[1];
    }
  } catch (e) {}

  // 3. セル値がそのままURL
  const val = String(cellRange.getValue() || "").trim();
  if (val.startsWith("http")) return val;

  return null;
}

// =============================================
// 店舗外部発送シートへ転記
// =============================================

/**
 * 店舗ごとの発送シート（F列リンク先）を開いて最新行に追記
 * ヘッダーがあれば動的マッピング、なければ標準並び順で書き込む
 */
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
      products = [{
        name:      data.productName      || "指定なし",
        url:       data.productUrl       || "",
        variation: data.productVariation || "-",
        details:   data.productDetails   || "-"
      }];
    }

    const startRow = targetSheet.getLastRow() + 1;

    // ヘッダーから列を動的に検出
    let dateCol=-1, creatorCol=-1, nameCol=-1, addrCol=-1, prodCol=-1, urlCol=-1, varCol=-1, detailCol=-1;
    if (values.length >= 1) {
      const headers = values[0];
      for (let c = 0; c < headers.length; c++) {
        const h = String(headers[c] || "").replace(/[\s　]/g, "").toLowerCase();
        if (dateCol    === -1 && (h.includes("日") || h.includes("date")))                               dateCol    = c;
        if (creatorCol === -1 && (h.includes("クリエイター") || h.includes("アカウント") || h === "id")) creatorCol = c;
        if (nameCol    === -1 && (h.includes("氏名") || h.includes("名前") || h.includes("宛名")))       nameCol    = c;
        if (addrCol    === -1 && (h.includes("住所") || h.includes("お届け先")))                         addrCol    = c;
        if (prodCol    === -1 && (h.includes("商品名") || h.includes("商品")))                           prodCol    = c;
        if (urlCol     === -1 && (h.includes("url") || h.includes("リンク")))                            urlCol     = c;
        if (varCol     === -1 && (h.includes("バリエ") || h.includes("カラー") || h.includes("サイズ") || h.includes("規格"))) varCol = c;
        if (detailCol  === -1 && (h.includes("備考") || h.includes("詳細") || h.includes("要望")))       detailCol  = c;
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
        // 標準並び順
        targetSheet.appendRow([
          now, data.creatorId, data.realName, data.address,
          p.name, p.url || "-", p.variation || "-", p.details || "-"
        ]);
      }
    });

    const endRow  = targetSheet.getLastRow();
    const rowText = (startRow === endRow) ? (startRow + "行目") : (startRow + "〜" + endRow + "行目");
    return { success: true, rowText, sheetTitle: targetSheet.getName() };

  } catch (err) {
    Logger.log("店舗発送シート書き込みエラー: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

// =============================================
// 自社配送シートへ転記
// =============================================

/**
 * 自社の配送シートに転記（1商品1行）
 * 列: A:記入日 B:店舗名 C:クリエイターID D:氏名 E:住所 F:商品名 G:商品リンク
 *     H:希望バリエーション I:商品詳細 J:発送ステータス K:発送完了日 L:LINE_USER_ID M:リマインド
 */
function appendShippingRecords(ss, data) {
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEET_NAME_SHIPPING);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEET_NAME_SHIPPING);
    sheet.appendRow([
      "記入日", "店舗名", "クリエイターID", "氏名", "お届け先住所",
      "商品名", "商品リンク", "希望バリエーション", "商品詳細・ご要望",
      "発送ステータス", "発送完了日", "LINE_USER_ID", "5日後リマインド"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13)
      .setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
  }

  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{
      name:      data.productName      || "指定なし",
      url:       data.productUrl       || "",
      variation: data.productVariation || "-",
      details:   data.productDetails   || "-"
    }];
  }

  const startRow = sheet.getLastRow() + 1;

  products.forEach(p => {
    sheet.appendRow([
      now,
      data.shopName,
      data.creatorId,
      data.realName,
      data.address,
      p.name,
      p.url       || "-",
      p.variation || "-",
      p.details   || "-",
      "未発送",
      "",
      data.userId,
      "未"
    ]);
  });

  const endRow  = sheet.getLastRow();
  const rowText = (startRow === endRow) ? (startRow + "行目") : (startRow + "〜" + endRow + "行目");
  return { startRow, endRow, rowText };
}

// =============================================
// 店舗LINEグループへ通知
// =============================================

/**
 * 承認後、店舗LINEグループへ転記完了と発送依頼を通知
 */
function notifyShopGroup(data, rowText, shopRowText) {
  const targetChatId = data.shopNotifyGroupId;
  if (!targetChatId) {
    Logger.log("通知先グループIDがないためスキップ");
    return;
  }

  let products = data.products;
  if (!products || !Array.isArray(products) || products.length === 0) {
    products = [{ name: data.productName || "指定なし", url: "", variation: "", details: "" }];
  }

  let productDetailsText = "";
  if (products.length === 1) {
    const p = products[0];
    productDetailsText = "■ 商品: " + p.name;
    if (p.url)       productDetailsText += "\n■ リンク: " + p.url;
    if (p.variation) productDetailsText += "\n■ バリエーション: " + p.variation;
    if (p.details)   productDetailsText += "\n■ 詳細・ご要望: " + p.details;
  } else {
    productDetailsText = "■ 申請商品 (計" + products.length + "点):\n" + products.map((p, i) => {
      let item = "・【商品" + (i + 1) + "】" + p.name;
      if (p.variation) item += "\n  └ バリエーション: " + p.variation;
      if (p.url)       item += "\n  └ リンク: " + p.url;
      if (p.details)   item += "\n  └ 詳細・備考: " + p.details;
      return item;
    }).join("\n");
  }

  const shopSheetPart = shopRowText
    ? ("\n■ 店舗様発送シート: 【" + shopRowText + "】に自動記入済")
    : "";

  const addressInfoText =
    "■ お届け先情報:\n" +
    "・お名前: " + data.realName + " 様\n" +
    "・ご住所: " + (data.address || "確認中");

  const messageText =
    "お疲れ様です！\n" +
    "商品サンプルの発送依頼をお願いいたします！\n\n" +
    "■ 店舗様: " + data.shopName + "\n" +
    "■ クリエイター: " + data.creatorId + " (" + data.realName + "様)\n\n" +
    productDetailsText + "\n\n" +
    addressInfoText +
    shopSheetPart + "\n" +
    "■ 配送シート管理番号: 【" + rowText + "】\n\n" +
    "ご確認の上、発送手配をお願いいたします🙇‍♂️";

  pushLineMessage(targetChatId, messageText);
}

// =============================================
// 発送後5日リマインド（定期トリガー: 毎日9〜10時に実行）
// =============================================

/**
 * 配送シートを走査し、発送済みから5日経過した行のクリエイターへLINEリマインドを送信
 */
function checkAndSendReminders() {
  const cfg   = getConfig();
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(cfg.SHEET_NAME_SHIPPING);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 配送シートの列定義（0始まり）
  // J=9:発送ステータス K=10:発送完了日 L=11:LINE_USER_ID M=12:リマインド F=5:商品名
  for (let i = 1; i < data.length; i++) {
    const status       = data[i][9];
    const shippedDate  = data[i][10];
    const userId       = data[i][11];
    const remindStatus = data[i][12];
    const productName  = data[i][5];

    if (status === "発送済" && remindStatus === "未" && shippedDate && userId) {
      const shipped = new Date(shippedDate);
      shipped.setHours(0, 0, 0, 0);
      const diffDays = Math.floor((today - shipped) / (1000 * 60 * 60 * 24));

      if (diffDays >= 5) {
        const remindText =
          "お世話になっております！\n" +
          "先日発送いたしましたサンプル「" + productName + "」はお手元に届きましたでしょうか？✨\n\n" +
          "【ご確認のお願い】\n" +
          "① TikTokライブでのご紹介日程がお決まりでしたら、本チャットにて教えていただけますと幸いです！\n" +
          "② ご紹介用のアフィリエイトリンクは以下となります。必ずこちらをご活用くださいませ！\n" +
          "🔗 " + cfg.AFFILIATE_INFO_TEXT + "\n\n" +
          "引き続きよろしくお願いいたします！";

        pushLineMessage(userId, remindText);
        sheet.getRange(i + 1, 13).setValue(
          "送信済 (" + Utilities.formatDate(today, "Asia/Tokyo", "MM/dd") + ")"
        );
      }
    }
  }
}

// =============================================
// LINE Messaging API
// =============================================

/**
 * LINE APIでグループ名を取得（店舗マスタのマッチングに使用）
 */
function fetchLineGroupName(groupId) {
  const token = getConfig().LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return null;

  try {
    const res = UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/group/" + groupId + "/summary",
      {
        method: "get",
        headers: { "Authorization": "Bearer " + token },
        muteHttpExceptions: true
      }
    );
    if (res.getResponseCode() === 200) {
      return JSON.parse(res.getContentText()).groupName;
    }
  } catch (e) {
    Logger.log("グループ名取得失敗: " + e.toString());
  }
  return null;
}

/**
 * LINE Messaging API Push送信
 */
function pushLineMessage(toId, text) {
  const token = getConfig().LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return;

  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      headers: {
        "Content-Type":  "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify({
        to:       toId,
        messages: [{ type: "text", text: text }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("Push送信エラー: " + e.toString());
  }
}

/**
 * JSONレスポンス生成
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
