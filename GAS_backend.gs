var SPREADSHEET_ID = '1CLYVwTISKxHFc583wFNCIIQ_unLFwqjHS-SxOPTYXgE';
var CUSTOMER_DB_SHEET = '顧客DB';
var FORM_SHEETS_CONFIG = [
  { sheetName: '西船橋店オンダリフト', shopCode: 'ONDARI_NISHIFUNA', shopLabel: '西船橋 オンダリフト', lineChannel: 'LINE_NISHIFUNA' },
  { sheetName: '西船橋店ピーリング', shopCode: 'PEELING_NISHIFUNA', shopLabel: '西船橋 ハーブピーリング', lineChannel: 'LINE_NISHIFUNA' }
];

// ============================================================
// 列インデックス定数（0始まり）
// ============================================================
var COL_ID       = 0;  // A: 顧客ID
var COL_NAME     = 1;  // B: 氏名
var COL_KANA     = 2;  // C: よみがな
var COL_PHONE    = 3;  // D: 電話番号
var COL_EMAIL    = 4;  // E: メールアドレス
var COL_BIRTH    = 5;  // F: 生年月日
var COL_SKIN     = 6;  // G: 肌タイプ
var COL_CONCERN  = 7;  // H: お悩み
var COL_SHOP     = 8;  // I: 店舗コード
var COL_LINE_ID  = 9;  // J: LINE_userId
var COL_LINE_DT  = 10; // K: LINE流入日時
var COL_STATUS   = 11; // L: ステータス
var COL_MEMO     = 12; // M: メモ
var COL_REG_DATE = 13; // N: 登録日時
var COL_UPDATED  = 14; // O: 最終更新

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'getCustomers')         return getCustomers(e);
  if (action === 'getCustomerDetail')    return getCustomerDetail(e);
  if (action === 'getContracts')         return getContracts(e);
  if (action === 'getConfig')            return getConfig(e);
  if (action === 'syncAllFormResponses') return syncAllFormResponses();
  if (action === 'updateCustomer')       return updateCustomerFields(e.parameter);
  if (action === 'cleanDuplicates')      return cleanDuplicates();
  if (action === 'getFormHeaders')       return getFormHeaders();
  if (action === 'getRawFormData')       return getRawFormData();
  if (action === 'linkLineId')           return linkLineIdByPhone(e.parameter);
  if (action === 'getLineUsers')         return getLineUsers();
  if (action === 'sendLine')             return sendLineFromDashboard(e.parameter);
  if (action === 'cleanCustomerDB')      return cleanCustomerDB();
  if (action === 'initDB')               return initDB();
  if (action === 'rebuildFromForms')     return rebuildFromForms();
  return jsonResponse({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.events) {
      var shopCode = (e.parameter && e.parameter.shop) ? e.parameter.shop : 'ONDARI_NISHIFUNA';
      for (var i = 0; i < body.events.length; i++) {
        var ev = body.events[i];
        if (ev.type === 'follow')  handleFollowEvent(ev, shopCode);
        if (ev.type === 'message') handleMessageEvent(ev);
      }
      return ok200();
    }
    var action = body.action;
    if (action === 'addTreatment')        return addTreatment(body);
    if (action === 'addContract')         return addContract(body);
    if (action === 'saveClosingResult')   return saveClosingResult(body);
    if (action === 'addCustomerManual')   return addCustomerManual(body);
    if (action === 'initializeSheets')    return initializeSheets();
    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch(err) {
    console.error('doPost error: ' + err.message);
    return ok200();
  }
}

function ok200() {
  return ContentService.createTextOutput('{"status":"ok"}').setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LINE Webhook ハンドラー
// ============================================================
function handleFollowEvent(event, shopCode) {
  try {
    var userId = event.source.userId;
    if (!userId) return;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
    if (!sheet) return;
    if (findCustomerByLineId(userId)) {
      console.log('既存LINE顧客: ' + userId);
      return;
    }
    var now = new Date();
    var newId = generateId();
    sheet.appendRow([
      newId, 'LINE新規', '', '', '', '', '', '', shopCode,
      userId, now, 'LINE新規', '', now, ''
    ]);
    console.log('LINE新規登録: ' + newId);
    var token = getToken();
    if (!token) return;
    var config = getShopConfig(shopCode);
    var label = config ? config.shopLabel : '当サロン';
    pushLine(token, userId, label + 'にご登録ありがとうございます！\nカウンセリングシートのご記入をお願いいたします。');
  } catch(err) { console.error('handleFollowEvent: ' + err.message); }
}

function handleMessageEvent(event) {
  try {
    var userId = event.source.userId;
    var replyToken = event.replyToken;
    var text = (event.message && event.message.type === 'text') ? String(event.message.text).trim() : '';
    if (!userId || !text) return;
    var token = getToken();
    if (!token) return;
    var phoneNorm = text.replace(/[-\s\(\)]/g, '');
    var isPhone = /^(0\d{9,10}|\d{10,11})$/.test(phoneNorm);
    if (isPhone) {
      var matched = findCustomerByPhone(phoneNorm);
      if (matched) {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][COL_ID]) === String(matched.customerId)) {
            var existingLineId = String(data[i][COL_LINE_ID] || '');
            if (!existingLineId) {
              sheet.getRange(i+1, COL_LINE_ID+1).setValue(userId);
              sheet.getRange(i+1, COL_LINE_DT+1).setValue(new Date());
              sheet.getRange(i+1, COL_UPDATED+1).setValue(new Date());
              console.log('電話番号でLINE紐づけ: ' + matched.customerId);
            }
            break;
          }
        }
        var existing = findCustomerByLineId(userId);
        if (existing && existing.customerId !== matched.customerId) {
          for (var j = 1; j < data.length; j++) {
            if (String(data[j][COL_ID]) === String(existing.customerId) && data[j][COL_NAME] === 'LINE新規') {
              sheet.deleteRow(j + 1);
              break;
            }
          }
        }
        replyLine(token, replyToken, matched.customerName + '様、LINE連携が完了しました！\nスタッフよりご連絡いたします。');
        return;
      } else {
        replyLine(token, replyToken, '電話番号が見つかりませんでした。\nご来店時にスタッフにお申し付けください。');
        return;
      }
    }
    var customer = findCustomerByLineId(userId);
    if (customer && customer.customerName !== 'LINE新規') {
      replyLine(token, replyToken, customer.customerName + '様、メッセージありがとうございます。\n担当スタッフより折り返しご連絡いたします。');
    } else {
      replyLine(token, replyToken, 'メッセージありがとうございます。\nご登録の電話番号を送っていただくとスムーズにご連絡できます。\n例：09012345678');
    }
  } catch(err) { console.error('handleMessageEvent: ' + err.message); }
}

// ============================================================
// 顧客DB CRUD
// ============================================================
function getCustomers(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse([]);
  var treatSheet = ss.getSheetByName('施術履歴');
  var treatData = treatSheet ? treatSheet.getDataRange().getValues() : [];
  var customers = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL_ID]) continue;
    var name = String(row[COL_NAME] || '');
    if (name === 'LINE新規') continue;
    var age = calcAge(row[COL_BIRTH]);
    var customerId = String(row[COL_ID]);
    var treatmentCount = 0;
    var lastVisit = '';
    for (var t = 1; t < treatData.length; t++) {
      if (String(treatData[t][0]) === customerId) {
        treatmentCount++;
        var tDate = treatData[t][1] ? Utilities.formatDate(new Date(treatData[t][1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '';
        if (tDate > lastVisit) lastVisit = tDate;
      }
    }
    customers.push({
      id:             customerId,
      name:           name,
      furigana:       String(row[COL_KANA] || ''),
      phone:          String(row[COL_PHONE] || ''),
      email:          String(row[COL_EMAIL] || ''),
      birthdate:      row[COL_BIRTH] ? formatBirthDate(row[COL_BIRTH]) : '',
      age:            age,
      skinType:       String(row[COL_SKIN] || ''),
      concerns:       row[COL_CONCERN] ? String(row[COL_CONCERN]).split(/[\/、,]/).map(function(s){return s.trim();}).filter(Boolean) : [],
      shop:           String(row[COL_SHOP] || ''),
      lineUserId:     String(row[COL_LINE_ID] || ''),
      lineInflowDate: row[COL_LINE_DT] ? Utilities.formatDate(new Date(row[COL_LINE_DT]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '',
      status:         String(row[COL_STATUS] || '新規'),
      memo:           String(row[COL_MEMO] || '').replace(/_ts:.*$/g, '').trim(),
      registeredDate: row[COL_REG_DATE] ? Utilities.formatDate(new Date(row[COL_REG_DATE]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
      lastVisit:      lastVisit,
      treatmentCount: treatmentCount,
      staff:          '',
      contractId:     ''
    });
  }
  var keyword = (e && e.parameter && e.parameter.keyword) ? e.parameter.keyword : '';
  if (keyword) {
    customers = customers.filter(function(c) {
      return c.name.indexOf(keyword) >= 0 || c.phone.indexOf(keyword) >= 0 || c.id.indexOf(keyword) >= 0;
    });
  }
  return jsonResponse(customers);
}

function getCustomerDetail(e) {
  var customerId = e.parameter.customerId;
  if (!customerId) return jsonResponse({ error: 'customerId required' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID]) === customerId) {
      var row = data[i];
      var customer = {
        customerId:       String(row[COL_ID]),
        customerName:     String(row[COL_NAME] || ''),
        furigana:         String(row[COL_KANA] || ''),
        phone:            String(row[COL_PHONE] || ''),
        email:            String(row[COL_EMAIL] || ''),
        birthDate:        String(row[COL_BIRTH] || ''),
        skinType:         String(row[COL_SKIN] || ''),
        concerns:         String(row[COL_CONCERN] || ''),
        shopCode:         String(row[COL_SHOP] || ''),
        lineId:           String(row[COL_LINE_ID] || ''),
        lineInflowDate:   row[COL_LINE_DT] ? Utilities.formatDate(new Date(row[COL_LINE_DT]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '',
        status:           String(row[COL_STATUS] || '新規'),
        memo:             String(row[COL_MEMO] || '').replace(/_ts:.*$/g, '').trim(),
        registrationDate: row[COL_REG_DATE] ? Utilities.formatDate(new Date(row[COL_REG_DATE]), 'Asia/Tokyo', 'yyyy-MM-dd') : ''
      };
      var treatments = getTreatments(ss, customerId);
      return jsonResponse({ customer: customer, treatments: treatments });
    }
  }
  return jsonResponse({ error: 'not found' });
}

function addNewCustomer(mapped) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var newId = generateId();
  var now = new Date();
  sheet.appendRow([
    newId,                         // A: 顧客ID
    mapped.customerName || '',     // B: 氏名
    mapped.furigana || '',         // C: よみがな
    mapped.phone || '',            // D: 電話番号
    mapped.email || '',            // E: メールアドレス
    mapped.birthDate || '',        // F: 生年月日
    mapped.skinType || '',         // G: 肌タイプ
    mapped.concerns || '',         // H: お悩み
    mapped.shopCode || '',         // I: 店舗コード
    '',                            // J: LINE_userId（空）
    '',                            // K: LINE流入日時（空）
    '新規',                         // L: ステータス
    mapped.memo || '',             // M: メモ
    now,                           // N: 登録日時
    now                            // O: 最終更新
  ]);
  return newId;
}

function updateCustomerFields(params) {
  var customerId = params.customerId;
  if (!customerId) return jsonResponse({ error: 'customerId required' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID]) === String(customerId)) {
      var row = i + 1;
      if (params.name)      sheet.getRange(row, COL_NAME+1).setValue(params.name);
      if (params.phone)     sheet.getRange(row, COL_PHONE+1).setValue(params.phone);
      if (params.email)     sheet.getRange(row, COL_EMAIL+1).setValue(params.email);
      if (params.birthDate) sheet.getRange(row, COL_BIRTH+1).setValue(params.birthDate);
      if (params.skinType !== undefined) sheet.getRange(row, COL_SKIN+1).setValue(params.skinType);
      if (params.memo !== undefined)     sheet.getRange(row, COL_MEMO+1).setValue(params.memo);
      if (params.status)    sheet.getRange(row, COL_STATUS+1).setValue(params.status);
      sheet.getRange(row, COL_UPDATED+1).setValue(new Date());
      return jsonResponse({ ok: true, customerId: customerId });
    }
  }
  return jsonResponse({ error: 'not found' });
}

function updateCustomerFromForm(rowIndex, mapped) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  if (mapped.customerName) sheet.getRange(rowIndex, COL_NAME+1).setValue(mapped.customerName);
  if (mapped.furigana)     sheet.getRange(rowIndex, COL_KANA+1).setValue(mapped.furigana);
  if (mapped.email)        sheet.getRange(rowIndex, COL_EMAIL+1).setValue(mapped.email);
  if (mapped.birthDate)    sheet.getRange(rowIndex, COL_BIRTH+1).setValue(mapped.birthDate);
  if (mapped.skinType)     sheet.getRange(rowIndex, COL_SKIN+1).setValue(mapped.skinType);
  if (mapped.concerns)     sheet.getRange(rowIndex, COL_CONCERN+1).setValue(mapped.concerns);
  if (mapped.shopCode)     sheet.getRange(rowIndex, COL_SHOP+1).setValue(mapped.shopCode);
  if (mapped.memo)         sheet.getRange(rowIndex, COL_MEMO+1).setValue(mapped.memo);
  sheet.getRange(rowIndex, COL_UPDATED+1).setValue(new Date());
  // LINE IDは絶対に上書きしない
  var existingLineId = sheet.getRange(rowIndex, COL_LINE_ID+1).getValue();
  if (!existingLineId) {
    tryAutoLinkLine(rowIndex);
  }
}

function tryAutoLinkLine(targetRowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
    var data = sheet.getDataRange().getValues();
    var lineNewRows = [];
    for (var i = 1; i < data.length; i++) {
      if (i + 1 === targetRowIndex) continue;
      if (data[i][COL_NAME] === 'LINE新規' && data[i][COL_LINE_ID] && !data[i][COL_PHONE]) {
        lineNewRows.push({ rowIndex: i+1, lineId: String(data[i][COL_LINE_ID]) });
      }
    }
    if (lineNewRows.length === 1) {
      sheet.getRange(targetRowIndex, COL_LINE_ID+1).setValue(lineNewRows[0].lineId);
      sheet.getRange(targetRowIndex, COL_LINE_DT+1).setValue(new Date());
      sheet.deleteRow(lineNewRows[0].rowIndex);
      console.log('LINE自動紐づけ完了: ' + lineNewRows[0].lineId);
      var token = getToken();
      var data2 = sheet.getDataRange().getValues();
      for (var j = 1; j < data2.length; j++) {
        if (j + 1 === targetRowIndex) {
          if (token && data2[j][COL_NAME]) {
            pushLine(token, lineNewRows[0].lineId, data2[j][COL_NAME] + '様、カウンセリングシートのご記入ありがとうございます！\n担当スタッフよりご連絡いたします。');
          }
          break;
        }
      }
    }
  } catch(err) { console.error('tryAutoLinkLine: ' + err.message); }
}

// ============================================================
// フォーム同期
// ============================================================
function syncAllFormResponses() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var total = 0;
  for (var f = 0; f < FORM_SHEETS_CONFIG.length; f++) {
    var config = FORM_SHEETS_CONFIG[f];
    var formSheet = ss.getSheetByName(config.sheetName);
    if (!formSheet || formSheet.getLastRow() < 2) continue;
    var rows = formSheet.getDataRange().getValues();
    var headers = rows[0];
    for (var i = 1; i < rows.length; i++) {
      var mapped = mapFormRow(rows[i], headers, config.shopCode);
      if (!mapped.customerName && !mapped.phone) continue;
      // タイムスタンプをキーにして重複チェック
      var timestamp = rows[i][0] ? String(rows[i][0]) : '';
      var existing = findCustomerByTimestamp(timestamp);
      if (existing) {
        updateCustomerFromForm(existing.rowIndex, mapped);
      } else {
        var byPhone = mapped.phone ? findCustomerByPhone(mapped.phone) : null;
        var byName = (!byPhone && mapped.customerName) ? findCustomerByName(mapped.customerName) : null;
        // 同名・同電話でも別タイムスタンプなら別人として登録
        if (byPhone && !byPhone.timestamp) {
          // 既存顧客にタイムスタンプがない場合のみ更新
          updateCustomerFromForm(byPhone.rowIndex, mapped);
          saveTimestamp(byPhone.rowIndex, timestamp);
        } else {
          addNewCustomerWithTimestamp(mapped, timestamp);
        }
      }
      total++;
    }
  }
  return jsonResponse({ ok: true, totalSynced: total });
}

function onFormSubmit(e) {
  try {
    var sheetName = e.source.getActiveSheet().getName();
    var shopCode = '';
    for (var f = 0; f < FORM_SHEETS_CONFIG.length; f++) {
      if (FORM_SHEETS_CONFIG[f].sheetName === sheetName) {
        shopCode = FORM_SHEETS_CONFIG[f].shopCode;
        break;
      }
    }
    if (!shopCode) return;

    var sheet = e.source.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = e.values || sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapped = mapFormRow(row, headers, shopCode);
    if (!mapped.customerName && !mapped.phone) return;

    var existing = mapped.phone ? findCustomerByPhone(mapped.phone) : null;
    if (!existing && mapped.customerName) existing = findCustomerByName(mapped.customerName);

    var customerId = '';
    var rowIndex = -1;
    if (existing) {
      updateCustomerFromForm(existing.rowIndex, mapped);
      customerId = existing.customerId;
      rowIndex = existing.rowIndex;
    } else {
      customerId = addNewCustomer(mapped);
      var ss2 = SpreadsheetApp.openById(SPREADSHEET_ID);
      var dbSheet2 = ss2.getSheetByName(CUSTOMER_DB_SHEET);
      var data2 = dbSheet2.getDataRange().getValues();
      for (var k = 1; k < data2.length; k++) {
        if (String(data2[k][COL_ID]) === String(customerId)) {
          rowIndex = k + 1;
          break;
        }
      }
    }

    if (!customerId || rowIndex < 0) return;

    // LINE新規レコードを全件取得
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dbSheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
    var allData = dbSheet.getDataRange().getValues();
    var lineNewRecords = [];
    for (var i = 1; i < allData.length; i++) {
      if (i + 1 === rowIndex) continue;
      if (String(allData[i][COL_NAME]) === 'LINE新規' && allData[i][COL_LINE_ID]) {
        lineNewRecords.push({
          rowIndex: i + 1,
          lineId: String(allData[i][COL_LINE_ID]),
          shopCode: String(allData[i][COL_SHOP] || '')
        });
      }
    }

    var currentLineId = dbSheet.getRange(rowIndex, COL_LINE_ID + 1).getValue();
    if (currentLineId) {
      console.log('LINE ID既に設定済み: ' + currentLineId);
      return;
    }

    if (lineNewRecords.length === 1) {
      // LINE新規が1件だけ → 確実にその人なので紐づけ
      var rec = lineNewRecords[0];
      dbSheet.getRange(rowIndex, COL_LINE_ID + 1).setValue(rec.lineId);
      dbSheet.getRange(rowIndex, COL_LINE_DT + 1).setValue(new Date());
      dbSheet.getRange(rowIndex, COL_UPDATED + 1).setValue(new Date());
      dbSheet.deleteRow(rec.rowIndex);
      console.log('フォーム回答でLINE自動紐づけ完了: ' + customerId + ' lineId=' + rec.lineId);

      var token = getToken();
      if (token && mapped.customerName) {
        pushLine(token, rec.lineId,
          mapped.customerName + '様、カウンセリングシートのご記入ありがとうございます！\n担当スタッフより改めてご連絡いたします。');
      }
    } else if (lineNewRecords.length === 0) {
      console.log('LINE新規レコードなし。LINE友達追加後に電話番号送信で紐づき可能。');
    } else {
      console.log('LINE新規が複数件(' + lineNewRecords.length + '件)のため自動紐づけスキップ。');
    }
  } catch(err) {
    console.error('onFormSubmit error: ' + err.message);
  }
}

function mapFormRow(row, headers, shopCode) {
  var result = { customerName:'', furigana:'', phone:'', email:'', birthDate:'', skinType:'', concerns:'', memo:'', shopCode:shopCode };
  var lastName='', firstName='', lastKana='', firstKana='';
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    var v = row[i] !== undefined ? String(row[i]).trim() : '';
    if (h === 'お名前（姓）' || h === '姓') lastName = v;
    else if (h === 'お名前（名）' || h === '名') firstName = v;
    else if ((h.indexOf('名前') >= 0 || h.indexOf('氏名') >= 0) && !lastName) result.customerName = v;
    else if (h === 'フリガナ（セイ）' || h === 'セイ') lastKana = v;
    else if (h === 'フリガナ（メイ）' || h === 'メイ') firstKana = v;
    else if ((h.indexOf('フリガナ') >= 0 || h.indexOf('ふりがな') >= 0) && !lastKana) result.furigana = v;
    else if (h.indexOf('電話') >= 0 || h === 'TEL') result.phone = v.replace(/[-\s]/g, '');
    else if (h.indexOf('メール') >= 0 || h.indexOf('mail') >= 0 || h === 'Email') result.email = v;
    else if (h.indexOf('生年月日') >= 0 || h.indexOf('誕生日') >= 0) result.birthDate = v;
    else if (h === '肌タイプ' || h === '肌質' || h === 'スキンタイプ') result.skinType = v;
    else if (h.indexOf('アレルギー') >= 0 || h.indexOf('ケロイド') >= 0) {
      if (v && v !== 'いいえ' && v !== 'なし') result.concerns = result.concerns ? result.concerns + ' / ' + v : v;
    }
    else if (h.indexOf('お悩み') >= 0 || h.indexOf('悩み') >= 0) result.concerns = result.concerns ? result.concerns + ' / ' + v : v;
    else if (h.indexOf('目的') >= 0 || h.indexOf('ご来店') >= 0) result.memo = result.memo ? result.memo + ' / ' + v : v;
    else if (h.indexOf('備考') >= 0) result.memo = result.memo ? result.memo + ' ' + v : v;
  }
  if (lastName || firstName) result.customerName = (lastName + ' ' + firstName).trim();
  if (lastKana || firstKana) result.furigana = (lastKana + ' ' + firstKana).trim();
  if (result.birthDate) {
    var bd = result.birthDate.replace(/[-\/]/g, '');
    if (/^\d{8}$/.test(bd)) result.birthDate = bd.slice(0,4)+'/'+bd.slice(4,6)+'/'+bd.slice(6,8);
  }
  return result;
}

// ============================================================
// LINE送信
// ============================================================
function sendLineFromDashboard(params) {
  var customerId = params.customerId || '';
  var message = params.message || '';
  if (!message) return jsonResponse({ error: 'message required' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var targetLineId = '';
  var customerName = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID]) === String(customerId)) {
      targetLineId = String(data[i][COL_LINE_ID] || '');
      customerName = String(data[i][COL_NAME] || '');
      break;
    }
  }
  if (!targetLineId) return jsonResponse({ error: 'LINE IDが設定されていません' });
  var token = getToken();
  if (!token) return jsonResponse({ error: 'LINE_CHANNEL_ACCESS_TOKEN未設定' });
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({ to: targetLineId, messages: [{ type:'text', text:message }] }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  console.log('LINE push: ' + code + ' ' + res.getContentText());
  if (code === 200) return jsonResponse({ ok: true, sentTo: customerName });
  return jsonResponse({ error: 'LINE API error: ' + res.getContentText() });
}

// ============================================================
// ユーティリティ
// ============================================================
function findCustomerByPhone(phone) {
  if (!phone) return null;
  var norm = phone.replace(/[-\s]/g, '');
  var norm10 = norm.replace(/^0/, '');
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var cell = String(data[i][COL_PHONE] || '').replace(/[-\s]/g, '');
    var cell10 = cell.replace(/^0/, '');
    if (cell && (cell === norm || cell10 === norm10)) {
      return {
        rowIndex: i+1,
        customerId: String(data[i][COL_ID]),
        customerName: String(data[i][COL_NAME]),
        phone: String(data[i][COL_PHONE]),
        lineId: String(data[i][COL_LINE_ID] || '')
      };
    }
  }
  return null;
}

function findCustomerByLineId(lineId) {
  if (!lineId) return null;
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_LINE_ID]) === lineId) {
      return { rowIndex:i+1, customerId:String(data[i][COL_ID]), customerName:String(data[i][COL_NAME]), phone:String(data[i][COL_PHONE]||''), lineId:lineId };
    }
  }
  return null;
}

function findCustomerByName(name) {
  if (!name) return null;
  var norm = name.replace(/\s/g,'');
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_NAME]||'').replace(/\s/g,'') === norm) {
      return { rowIndex:i+1, customerId:String(data[i][COL_ID]), customerName:String(data[i][COL_NAME]), phone:String(data[i][COL_PHONE]||''), lineId:String(data[i][COL_LINE_ID]||'') };
    }
  }
  return null;
}

function findCustomerByTimestamp(timestamp) {
  return null; // タイムスタンプ方式を廃止
}

function saveTimestamp(rowIndex, timestamp) {
  // タイムスタンプ方式を廃止 — メモ欄には書き込まない
}

function addNewCustomerWithTimestamp(mapped, timestamp) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var newId = generateId();
  var now = new Date();
  var memo = mapped.memo || '';
  // タイムスタンプはメモに含めず、顧客IDとして管理
  sheet.appendRow([
    newId,
    mapped.customerName || '',
    mapped.furigana || '',
    mapped.phone || '',
    mapped.email || '',
    mapped.birthDate || '',
    mapped.skinType || '',
    mapped.concerns || '',
    mapped.shopCode || '',
    '',
    '',
    '新規',
    memo,
    now,
    now
  ]);
  return newId;
}

function generateId() {
  var now = new Date();
  return 'C' + now.getFullYear() + pad(now.getMonth()+1) + pad(now.getDate()) + pad(now.getHours()) + pad(now.getMinutes()) + pad(now.getSeconds()) + ('00'+now.getMilliseconds()).slice(-3);
}

function pad(n) { return ('0'+n).slice(-2); }

function formatBirthDate(val) {
  if (!val) return '';
  try {
    var d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch(e) { return String(val); }
}

function calcAge(birthVal) {
  if (!birthVal) return '';
  try {
    var birth;
    if (birthVal instanceof Date) {
      birth = birthVal;
    } else {
      var s = String(birthVal);
      if (s.indexOf('/') !== -1) {
        var p = s.split('/');
        birth = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
      } else if (s.indexOf('-') !== -1) {
        var p = s.split('-');
        birth = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
      } else {
        birth = new Date(s);
      }
    }
    if (isNaN(birth.getTime())) return '';
    var today = new Date();
    var age = today.getFullYear() - birth.getFullYear();
    if (today.getMonth() < birth.getMonth() || (today.getMonth()===birth.getMonth() && today.getDate()<birth.getDate())) age--;
    return (age>=0 && age<120) ? age : '';
  } catch(e) { return ''; }
}

function getToken() {
  return PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
}

function getShopConfig(shopCode) {
  for (var i = 0; i < FORM_SHEETS_CONFIG.length; i++) {
    if (FORM_SHEETS_CONFIG[i].shopCode === shopCode) return FORM_SHEETS_CONFIG[i];
  }
  return null;
}

function pushLine(token, userId, message) {
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method:'post', contentType:'application/json',
    headers:{'Authorization':'Bearer '+token},
    payload:JSON.stringify({to:userId, messages:[{type:'text',text:message}]}),
    muteHttpExceptions:true
  });
}

function replyLine(token, replyToken, message) {
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method:'post', contentType:'application/json',
    headers:{'Authorization':'Bearer '+token},
    payload:JSON.stringify({replyToken:replyToken, messages:[{type:'text',text:message}]}),
    muteHttpExceptions:true
  });
}

function getTreatments(ss, customerId) {
  var sheet = ss.getSheetByName('施術履歴');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === customerId) {
      result.push({ date:data[i][1]?Utilities.formatDate(new Date(data[i][1]),'Asia/Tokyo','yyyy-MM-dd'):'', shopCode:String(data[i][2]||''), menuName:String(data[i][3]||''), staff:String(data[i][4]||''), note:String(data[i][5]||'') });
    }
  }
  return result;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify({ok:true, data:obj})).setMimeType(ContentService.MimeType.JSON);
}

function getConfig(e) {
  return jsonResponse({ formSheets:FORM_SHEETS_CONFIG, customerDbSheet:CUSTOMER_DB_SHEET, spreadsheetId:SPREADSHEET_ID });
}

function getFormHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {};
  for (var i = 0; i < FORM_SHEETS_CONFIG.length; i++) {
    var config = FORM_SHEETS_CONFIG[i];
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) continue;
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].filter(function(h){return h!=='';});
    result[config.shopCode] = { shopLabel:config.shopLabel, headers:headers };
  }
  return jsonResponse(result);
}

function getRawFormData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allRows = [];
  for (var i = 0; i < FORM_SHEETS_CONFIG.length; i++) {
    var config = FORM_SHEETS_CONFIG[i];
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet || sheet.getLastRow() < 2) continue;
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    for (var r = 0; r < rows.length; r++) {
      var obj = { _shopCode:config.shopCode, _shopLabel:config.shopLabel, _lineChannel:config.lineChannel };
      for (var c = 0; c < headers.length; c++) {
        var key = String(headers[c]).trim();
        if (key) obj[key] = rows[r][c] !== undefined ? String(rows[r][c]) : '';
      }
      allRows.push(obj);
    }
  }
  return jsonResponse(allRows);
}

function getContracts(e) {
  var customerId = e.parameter.customerId;
  if (!customerId) return jsonResponse({ error: 'customerId required' });
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('契約情報');
  if (!sheet) return jsonResponse([]);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === customerId) {
      result.push({ customerId:String(data[i][0]), contractDate:data[i][1]?Utilities.formatDate(new Date(data[i][1]),'Asia/Tokyo','yyyy-MM-dd'):'', shopCode:String(data[i][2]||''), courseName:String(data[i][3]||''), sessions:data[i][4]||0, amount:data[i][5]||0, paymentMethod:String(data[i][6]||''), status:String(data[i][7]||''), note:String(data[i][8]||'') });
    }
  }
  return jsonResponse(result);
}

function addTreatment(data) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('施術履歴');
  if (!sheet) { sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet('施術履歴'); sheet.appendRow(['顧客ID','施術日','店舗コード','メニュー名','担当スタッフ','備考','Before写真','After写真']); }
  sheet.appendRow([data.customerId||'', data.date||new Date(), data.shopCode||'', data.menuName||'', data.staff||'', data.note||'', data.beforePhoto||'', data.afterPhoto||'']);
  return jsonResponse({ ok:true });
}

function addContract(data) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('契約情報');
  if (!sheet) { sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet('契約情報'); sheet.appendRow(['顧客ID','契約日','店舗コード','コース名','回数','金額','支払方法','ステータス','備考']); }
  sheet.appendRow([data.customerId||'', data.contractDate||new Date(), data.shopCode||'', data.courseName||'', data.sessions||0, data.amount||0, data.paymentMethod||'', data.status||'有効', data.note||'']);
  return jsonResponse({ ok:true });
}

function saveClosingResult(data) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('クロージング結果');
  if (!sheet) { sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet('クロージング結果'); sheet.appendRow(['顧客ID','日付','店舗コード','提案コース','結果','金額','備考']); }
  sheet.appendRow([data.customerId||'', data.date||new Date(), data.shopCode||'', data.courseName||'', data.result||'', data.amount||0, data.note||'']);
  return jsonResponse({ ok:true });
}

function addCustomerManual(data) {
  var newId = addNewCustomer({ customerName:data.customerName||'', furigana:data.furigana||'', phone:data.phone||'', email:data.email||'', birthDate:data.birthDate||'', skinType:data.skinType||'', concerns:data.concerns||'', memo:data.memo||'', shopCode:data.shopCode||'' });
  return jsonResponse({ ok:true, customerId:newId });
}

function initializeSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(CUSTOMER_DB_SHEET);
    dbSheet.appendRow(['顧客ID','氏名','よみがな','電話番号','メールアドレス','生年月日','肌タイプ','お悩み','店舗コード','LINE_userId','LINE流入日時','ステータス','メモ','登録日時','最終更新']);
  }
  return jsonResponse({ ok:true });
}

function cleanDuplicates() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var seen = {}, toDelete = [];
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][COL_PHONE]||'').replace(/[-\s]/g,'') || String(data[i][COL_NAME]||'').replace(/\s/g,'');
    if (!key) continue;
    if (seen[key]) toDelete.push(i+1); else seen[key] = true;
  }
  for (var j = toDelete.length-1; j >= 0; j--) sheet.deleteRow(toDelete[j]);
  return jsonResponse({ deleted:toDelete.length });
}

function cleanCustomerDB() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var keepNames = ['林 治希','新里 愛','髙橋 将太','儀間 理生','氏家 大樹'];
  var toDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (!keepNames.includes(String(data[i][COL_NAME]||'').trim())) toDelete.push(i+1);
  }
  for (var j = toDelete.length-1; j >= 0; j--) sheet.deleteRow(toDelete[j]);
  return jsonResponse({ deleted:toDelete.length });
}

function initDB() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CUSTOMER_DB_SHEET);
  }
  var headers = sheet.getRange(1, 1, 1, 15).getValues()[0];
  var expected = ['顧客ID','氏名','よみがな','電話番号','メールアドレス','生年月日','肌タイプ','お悩み','店舗コード','LINE_userId','LINE流入日時','ステータス','メモ','登録日時','最終更新'];
  var needsHeader = !headers[0] || headers[0] !== '顧客ID';
  if (needsHeader) {
    sheet.getRange(1, 1, 1, 15).setValues([expected]);
    console.log('ヘッダー初期化完了');
  }
  return jsonResponse({ ok: true, message: 'DB初期化完了' });
}

function rebuildFromForms() {
  initDB();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var lastRow = dbSheet.getLastRow();
  if (lastRow > 1) {
    dbSheet.deleteRows(2, lastRow - 1);
  }
  var total = 0;
  for (var f = 0; f < FORM_SHEETS_CONFIG.length; f++) {
    var config = FORM_SHEETS_CONFIG[f];
    var formSheet = ss.getSheetByName(config.sheetName);
    if (!formSheet || formSheet.getLastRow() < 2) continue;
    var rows = formSheet.getDataRange().getValues();
    var headers = rows[0];
    for (var i = 1; i < rows.length; i++) {
      var mapped = mapFormRow(rows[i], headers, config.shopCode);
      if (!mapped.customerName && !mapped.phone) continue;
      var timestamp = rows[i][0] ? String(rows[i][0]) : '';
      addNewCustomerWithTimestamp(mapped, timestamp);
      total++;
    }
  }
  return jsonResponse({ ok: true, totalRebuilt: total });
}

function linkLineIdByPhone(params) {
  var phone = params.phone || '';
  var lineUserId = params.lineUserId || '';
  if (!phone || !lineUserId) return jsonResponse({ error: 'phone and lineUserId required' });
  var customer = findCustomerByPhone(phone.replace(/[-\s]/g,''));
  if (!customer) return jsonResponse({ error: '電話番号が見つかりません' });
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  sheet.getRange(customer.rowIndex, COL_LINE_ID+1).setValue(lineUserId);
  sheet.getRange(customer.rowIndex, COL_LINE_DT+1).setValue(new Date());
  return jsonResponse({ ok:true, customerId:customer.customerId, name:customer.customerName });
}

function getLineUsers() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][COL_LINE_ID]) {
      result.push({ customerId:String(data[i][COL_ID]), name:String(data[i][COL_NAME]), phone:String(data[i][COL_PHONE]||''), lineUserId:String(data[i][COL_LINE_ID]), status:String(data[i][COL_STATUS]||'') });
    }
  }
  return jsonResponse(result);
}

function setupProperties() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('LINE_CHANNEL_ACCESS_TOKEN', 'lEFEcdsU7W00c0nexEy0q5bVgzwa6PSknzbieVxTz16xx6UZ9hJ4fNssaNv32mrTRayAeHqKL6lrV1XCdr26vy8kgvwvoaKqb5do/QIlV7c5pEzMJFRKbEhaA6gZkBIckhTnKXkEb1xkJ6Oaf3aepAdB04t89/1O/w1cDnyilFU=');
  props.setProperty('LINE_CHANNEL_SECRET', '6ab448d0c63c2635f3ca8e602e4afd90');
  console.log('プロパティ設定完了');
  return { ok:true };
}

function diagToken() {
  var token = getToken();
  console.log('TOKEN存在: ' + (token ? 'あり（長さ:' + token.length + '）' : 'なし'));
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  console.log('スプレッドシート接続: ' + (sheet ? 'OK' : 'NG'));
  console.log('顧客DB行数: ' + sheet.getLastRow());
}
