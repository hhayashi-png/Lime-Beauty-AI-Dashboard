// ============================================================
// 設定（店舗追加はここだけ変更）
// ============================================================
var SPREADSHEET_ID = '1CLYVwTISKxHFc583wFNCIIQ_unLFwqjHS-SxOPTYXgE';
var CUSTOMER_DB_SHEET = '顧客DB';

var SHOPS = [
  {
    code:      'ONDARI_NISHIFUNA',
    label:     '西船橋 オンダリフト',
    sheetName: '西船橋店オンダリフト',
    aliases:   ['NISHIFUNA_ONDARI', 'NISHIFUNA', 'ONDARI']
  },
  {
    code:      'PEELING_NISHIFUNA',
    label:     '西船橋 ハーブピーリング',
    sheetName: '西船橋店ピーリング',
    aliases:   ['NISHIFUNA_PEELING', 'PEELING']
  }
];

// ============================================================
// 列定数（絶対に変更しない・追加は右端のみ）
// ============================================================
var COL = {
  ID:         0,  // A: 顧客ID
  NAME:       1,  // B: 氏名
  KANA:       2,  // C: よみがな
  PHONE:      3,  // D: 電話番号
  EMAIL:      4,  // E: メールアドレス
  BIRTH:      5,  // F: 生年月日
  SKIN:       6,  // G: 肌タイプ
  CONCERN:    7,  // H: お悩み
  SHOP:       8,  // I: 店舗コード
  LINE_ID:    9,  // J: LINE_userId ※絶対上書き禁止
  LINE_DT:    10, // K: LINE流入日時
  STATUS:     11, // L: ステータス
  MEMO:       12, // M: メモ（スタッフ入力のみ）
  REG_DATE:   13, // N: 登録日時
  UPDATED:    14, // O: 最終更新
  FORM_TS:    15  // P: フォームタイムスタンプ（重複防止キー）
};

var DB_HEADERS = [
  '顧客ID','氏名','よみがな','電話番号','メールアドレス',
  '生年月日','肌タイプ','お悩み','店舗コード','LINE_userId',
  'LINE流入日時','ステータス','メモ','登録日時','最終更新','フォームTS'
];

// ============================================================
// ルーティング
// ============================================================
function doGet(e) {
  var action = e.parameter.action;
  try {
    if (action === 'getCustomers')         return getCustomers(e);
    if (action === 'getCustomerDetail')    return getCustomerDetail(e);
    if (action === 'getContracts')         return getContracts(e);
    if (action === 'getConfig')            return getConfig();
    if (action === 'syncAllFormResponses') return syncAllFormResponses();
    if (action === 'updateCustomer')       return updateCustomerFields(e.parameter);
    if (action === 'getFormHeaders')       return getFormHeaders();
    if (action === 'getRawFormData')       return getRawFormData();
    if (action === 'getLineUsers')         return getLineUsers();
    if (action === 'sendLine')             return sendLineFromDashboard(e.parameter);
    if (action === 'initDB')               return initDB();
    if (action === 'rebuildFromForms')     return rebuildFromForms();
    if (action === 'addContract')          return addContractFromGet(e.parameter);
    return jsonError('Unknown action: ' + action);
  } catch(err) {
    console.error('doGet error: ' + err.message);
    return jsonError(err.message);
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.events) {
      var rawShop = (e.parameter && e.parameter.shop) ? e.parameter.shop : '';
      var resolvedShop = getShopByCode(rawShop);
      var shopCode = resolvedShop ? resolvedShop.code : SHOPS[0].code;
      body.events.forEach(function(ev) {
        if (ev.type === 'follow')  handleFollowEvent(ev, shopCode);
        if (ev.type === 'message') handleMessageEvent(ev);
      });
      return ok200();
    }
    var action = body.action;
    if (action === 'addContract') return addContract(body);
    return jsonError('Unknown action: ' + action);
  } catch(err) {
    console.error('doPost error: ' + err.message);
    return ok200();
  }
}

function ok200() {
  return ContentService.createTextOutput('{"status":"ok"}').setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// DB初期化（ヘッダー保護）
// ============================================================
function initDB() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CUSTOMER_DB_SHEET);
  }
  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== '顧客ID') {
    sheet.getRange(1, 1, 1, DB_HEADERS.length).setValues([DB_HEADERS]);
    console.log('DB初期化: ヘッダー設定完了');
  }
  return jsonResponse({ ok: true, message: 'DB初期化完了' });
}

// ============================================================
// フォーム同期（タイムスタンプをキーに重複防止）
// ============================================================
function syncAllFormResponses() {
  initDB();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var total = 0;
  SHOPS.forEach(function(shop) {
    var formSheet = ss.getSheetByName(shop.sheetName);
    if (!formSheet || formSheet.getLastRow() < 2) return;
    var rows = formSheet.getDataRange().getValues();
    var headers = rows[0];
    for (var i = 1; i < rows.length; i++) {
      var ts = rows[i][0] ? String(rows[i][0]) : '';
      var mapped = mapFormRow(rows[i], headers, shop.code);
      if (!mapped.customerName && !mapped.phone) continue;
      var existing = ts ? findCustomerByFormTS(ts) : null;
      if (existing) {
        updateCustomerFromForm(existing.rowIndex, mapped);
      } else {
        addNewCustomer(mapped, ts);
      }
      total++;
    }
  });
  return jsonResponse({ ok: true, totalSynced: total });
}

function rebuildFromForms() {
  initDB();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(CUSTOMER_DB_SHEET);

  // 既存のLINE ID情報を退避（消えないように保護）
  var existingData = dbSheet.getDataRange().getValues();
  var lineIdMap = {}; // phone -> {lineId, lineDt}
  var tsLineIdMap = {}; // formTS -> {lineId, lineDt}
  for (var i = 1; i < existingData.length; i++) {
    var row = existingData[i];
    var lineId = String(row[COL.LINE_ID] || '');
    if (!lineId) continue;
    var phone = String(row[COL.PHONE] || '').replace(/[-\s]/g,'');
    var ts = String(row[COL.FORM_TS] || '');
    var lineDt = row[COL.LINE_DT] || '';
    if (phone) lineIdMap[phone] = { lineId: lineId, lineDt: lineDt };
    if (ts) tsLineIdMap[ts] = { lineId: lineId, lineDt: lineDt };
  }
  console.log('LINE ID退避完了: ' + Object.keys(lineIdMap).length + '件');

  // DBをクリア（ヘッダーは残す）
  var lastRow = dbSheet.getLastRow();
  if (lastRow > 1) dbSheet.deleteRows(2, lastRow - 1);

  // フォームから再構築
  var total = 0;
  SHOPS.forEach(function(shop) {
    var formSheet = ss.getSheetByName(shop.sheetName);
    if (!formSheet || formSheet.getLastRow() < 2) return;
    var rows = formSheet.getDataRange().getValues();
    var headers = rows[0];
    for (var i = 1; i < rows.length; i++) {
      var ts = rows[i][0] ? String(rows[i][0]) : '';
      var mapped = mapFormRow(rows[i], headers, shop.code);
      if (!mapped.customerName && !mapped.phone) continue;

      // LINE IDを復元（フォームTSまたは電話番号で照合）
      var savedLine = null;
      if (ts && tsLineIdMap[ts]) {
        savedLine = tsLineIdMap[ts];
      } else if (mapped.phone) {
        var normPhone = mapped.phone.replace(/[-\s]/g,'');
        var norm10 = normPhone.replace(/^0/,'');
        for (var p in lineIdMap) {
          if (p === normPhone || p.replace(/^0/,'') === norm10) {
            savedLine = lineIdMap[p];
            break;
          }
        }
      }

      var newId = addNewCustomer(mapped, ts);
      if (savedLine && savedLine.lineId) {
        var data2 = dbSheet.getDataRange().getValues();
        for (var k = 1; k < data2.length; k++) {
          if (String(data2[k][COL.ID]) === String(newId)) {
            dbSheet.getRange(k+1, COL.LINE_ID+1).setValue(savedLine.lineId);
            if (savedLine.lineDt) dbSheet.getRange(k+1, COL.LINE_DT+1).setValue(savedLine.lineDt);
            console.log('LINE ID復元: ' + newId + ' -> ' + savedLine.lineId);
            break;
          }
        }
      }
      total++;
    }
  });

  console.log('rebuildFromForms完了: ' + total + '件（LINE ID保持）');
  return jsonResponse({ ok: true, totalRebuilt: total });
}

function onFormSubmit(e) {
  try {
    var sheetName = e.source.getActiveSheet().getName();
    var shop = getShopBySheetName(sheetName);
    if (!shop) return;
    var sheet = e.source.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = e.values || sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    var ts = row[0] ? String(row[0]) : '';
    var mapped = mapFormRow(row, headers, shop.code);
    if (!mapped.customerName && !mapped.phone) return;
    var existing = ts ? findCustomerByFormTS(ts) : null;
    var customerId, rowIndex;
    if (existing) {
      updateCustomerFromForm(existing.rowIndex, mapped);
      customerId = existing.customerId;
      rowIndex = existing.rowIndex;
    } else {
      customerId = addNewCustomer(mapped, ts);
      var dbSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
      var data = dbSheet.getDataRange().getValues();
      for (var k = 1; k < data.length; k++) {
        if (String(data[k][COL.ID]) === String(customerId)) { rowIndex = k + 1; break; }
      }
    }
    if (!customerId || !rowIndex) return;
    tryAutoLinkLine(rowIndex, customerId, mapped.customerName);
  } catch(err) { console.error('onFormSubmit: ' + err.message); }
}

// ============================================================
// LINE連携（1つの経路のみ・LINE IDは絶対上書き禁止）
// ============================================================
function handleFollowEvent(event, shopCode) {
  try {
    var userId = event.source.userId;
    if (!userId) return;
    if (findCustomerByLineId(userId)) return;
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
    var now = new Date();
    var newId = generateId();
    var row = new Array(DB_HEADERS.length).fill('');
    row[COL.ID]      = newId;
    row[COL.NAME]    = 'LINE新規';
    row[COL.SHOP]    = shopCode;
    row[COL.LINE_ID] = userId;
    row[COL.LINE_DT] = now;
    row[COL.STATUS]  = 'LINE新規';
    row[COL.REG_DATE]= now;
    row[COL.UPDATED] = now;
    sheet.appendRow(row);
    console.log('LINE新規登録: ' + newId);
    var token = getToken();
    if (!token) return;
    var shop = getShopByCode(shopCode);
    var label = shop ? shop.label : '当サロン';
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
    if (/^(0\d{9,10}|\d{10,11})$/.test(phoneNorm)) {
      var matched = findCustomerByPhone(phoneNorm);
      if (matched) {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][COL.ID]) === String(matched.customerId)) {
            if (!data[i][COL.LINE_ID]) {
              sheet.getRange(i+1, COL.LINE_ID+1).setValue(userId);
              sheet.getRange(i+1, COL.LINE_DT+1).setValue(new Date());
              sheet.getRange(i+1, COL.UPDATED+1).setValue(new Date());
            }
            break;
          }
        }
        var lineNew = findCustomerByLineId(userId);
        if (lineNew && lineNew.customerId !== matched.customerId) {
          var data2 = sheet.getDataRange().getValues();
          for (var j = 1; j < data2.length; j++) {
            if (String(data2[j][COL.ID]) === String(lineNew.customerId) && data2[j][COL.NAME] === 'LINE新規') {
              sheet.deleteRow(j + 1); break;
            }
          }
        }
        replyLine(token, replyToken, matched.customerName + '様、LINE連携が完了しました！\nスタッフよりご連絡いたします。');
        return;
      }
      replyLine(token, replyToken, '電話番号が見つかりませんでした。\nご来店時にスタッフにお申し付けください。');
      return;
    }
    var customer = findCustomerByLineId(userId);
    if (customer && customer.customerName !== 'LINE新規') {
      replyLine(token, replyToken, customer.customerName + '様、メッセージありがとうございます。\n担当スタッフより折り返しご連絡いたします。');
    } else {
      replyLine(token, replyToken, 'メッセージありがとうございます。\nご登録の電話番号を送っていただくとスムーズにご連絡できます。\n例：09012345678');
    }
  } catch(err) { console.error('handleMessageEvent: ' + err.message); }
}

function tryAutoLinkLine(targetRowIndex, customerId, customerName) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
    var currentLineId = sheet.getRange(targetRowIndex, COL.LINE_ID+1).getValue();
    if (currentLineId) return;
    var data = sheet.getDataRange().getValues();
    var lineNewRows = [];
    for (var i = 1; i < data.length; i++) {
      if (i + 1 === targetRowIndex) continue;
      if (data[i][COL.NAME] === 'LINE新規' && data[i][COL.LINE_ID] && !data[i][COL.PHONE]) {
        lineNewRows.push({ rowIndex: i+1, lineId: String(data[i][COL.LINE_ID]) });
      }
    }
    if (lineNewRows.length !== 1) return;
    var rec = lineNewRows[0];
    sheet.getRange(targetRowIndex, COL.LINE_ID+1).setValue(rec.lineId);
    sheet.getRange(targetRowIndex, COL.LINE_DT+1).setValue(new Date());
    sheet.getRange(targetRowIndex, COL.UPDATED+1).setValue(new Date());
    sheet.deleteRow(rec.rowIndex);
    console.log('LINE自動紐づけ完了: ' + customerId + ' lineId=' + rec.lineId);
    var token = getToken();
    if (token && customerName) {
      pushLine(token, rec.lineId, customerName + '様、カウンセリングシートのご記入ありがとうございます！\n担当スタッフより改めてご連絡いたします。');
    }
  } catch(err) { console.error('tryAutoLinkLine: ' + err.message); }
}

// ============================================================
// 顧客CRUD
// ============================================================
function getCustomers(e) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse([]);
  var customers = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL.ID] || row[COL.NAME] === 'LINE新規') continue;
    customers.push({
      id:             String(row[COL.ID]),
      name:           String(row[COL.NAME] || ''),
      furigana:       String(row[COL.KANA] || ''),
      phone:          String(row[COL.PHONE] || ''),
      email:          String(row[COL.EMAIL] || ''),
      birthdate:      row[COL.BIRTH] ? formatDate(row[COL.BIRTH]) : '',
      age:            calcAge(row[COL.BIRTH]),
      skinType:       String(row[COL.SKIN] || ''),
      concerns:       row[COL.CONCERN] ? String(row[COL.CONCERN]).split(/[\/、,]/).map(function(s){return s.trim();}).filter(Boolean) : [],
      shop:           String(row[COL.SHOP] || ''),
      lineUserId:     String(row[COL.LINE_ID] || ''),
      lineInflowDate: row[COL.LINE_DT] ? Utilities.formatDate(new Date(row[COL.LINE_DT]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '',
      status:         String(row[COL.STATUS] || '新規'),
      memo:           String(row[COL.MEMO] || ''),
      registeredDate: row[COL.REG_DATE] ? Utilities.formatDate(new Date(row[COL.REG_DATE]), 'Asia/Tokyo', 'yyyy-MM-dd') : ''
    });
  }
  var keyword = (e && e.parameter && e.parameter.keyword) ? e.parameter.keyword : '';
  if (keyword) {
    customers = customers.filter(function(c) {
      return c.name.indexOf(keyword) >= 0 || c.phone.indexOf(keyword) >= 0;
    });
  }
  return jsonResponse(customers);
}

function getCustomerDetail(e) {
  var customerId = e.parameter.customerId;
  if (!customerId) return jsonError('customerId required');
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.ID]) === customerId) {
      var row = data[i];
      return jsonResponse({
        customerId:   String(row[COL.ID]),
        customerName: String(row[COL.NAME] || ''),
        furigana:     String(row[COL.KANA] || ''),
        phone:        String(row[COL.PHONE] || ''),
        email:        String(row[COL.EMAIL] || ''),
        birthDate:    row[COL.BIRTH] ? formatDate(row[COL.BIRTH]) : '',
        skinType:     String(row[COL.SKIN] || ''),
        concerns:     String(row[COL.CONCERN] || ''),
        shopCode:     String(row[COL.SHOP] || ''),
        lineId:       String(row[COL.LINE_ID] || ''),
        status:       String(row[COL.STATUS] || '新規'),
        memo:         String(row[COL.MEMO] || ''),
        registrationDate: row[COL.REG_DATE] ? Utilities.formatDate(new Date(row[COL.REG_DATE]), 'Asia/Tokyo', 'yyyy-MM-dd') : ''
      });
    }
  }
  return jsonError('not found');
}

function updateCustomerFields(params) {
  var customerId = params.customerId;
  if (!customerId) return jsonError('customerId required');
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.ID]) === String(customerId)) {
      var r = i + 1;
      if (params.name)     sheet.getRange(r, COL.NAME+1).setValue(params.name);
      if (params.phone)    sheet.getRange(r, COL.PHONE+1).setValue(params.phone);
      if (params.email)    sheet.getRange(r, COL.EMAIL+1).setValue(params.email);
      if (params.birthDate)sheet.getRange(r, COL.BIRTH+1).setValue(params.birthDate);
      if (params.skinType !== undefined) sheet.getRange(r, COL.SKIN+1).setValue(params.skinType);
      if (params.memo !== undefined)     sheet.getRange(r, COL.MEMO+1).setValue(params.memo);
      if (params.status)   sheet.getRange(r, COL.STATUS+1).setValue(params.status);
      sheet.getRange(r, COL.UPDATED+1).setValue(new Date());
      return jsonResponse({ ok: true });
    }
  }
  return jsonError('not found');
}

function addNewCustomer(mapped, formTS) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var newId = generateId();
  var now = new Date();
  var row = new Array(DB_HEADERS.length).fill('');
  row[COL.ID]       = newId;
  row[COL.NAME]     = mapped.customerName || '';
  row[COL.KANA]     = mapped.furigana || '';
  row[COL.PHONE]    = mapped.phone || '';
  row[COL.EMAIL]    = mapped.email || '';
  row[COL.BIRTH]    = mapped.birthDate || '';
  row[COL.SKIN]     = mapped.skinType || '';
  row[COL.CONCERN]  = mapped.concerns || '';
  row[COL.SHOP]     = mapped.shopCode || '';
  row[COL.LINE_ID]  = '';
  row[COL.STATUS]   = '新規';
  row[COL.MEMO]     = mapped.memo || '';
  row[COL.REG_DATE] = now;
  row[COL.UPDATED]  = now;
  row[COL.FORM_TS]  = formTS || '';
  sheet.appendRow(row);
  return newId;
}

function updateCustomerFromForm(rowIndex, mapped) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  if (mapped.customerName) sheet.getRange(rowIndex, COL.NAME+1).setValue(mapped.customerName);
  if (mapped.furigana)     sheet.getRange(rowIndex, COL.KANA+1).setValue(mapped.furigana);
  if (mapped.email)        sheet.getRange(rowIndex, COL.EMAIL+1).setValue(mapped.email);
  if (mapped.birthDate)    sheet.getRange(rowIndex, COL.BIRTH+1).setValue(mapped.birthDate);
  if (mapped.skinType)     sheet.getRange(rowIndex, COL.SKIN+1).setValue(mapped.skinType);
  if (mapped.concerns)     sheet.getRange(rowIndex, COL.CONCERN+1).setValue(mapped.concerns);
  if (mapped.shopCode)     sheet.getRange(rowIndex, COL.SHOP+1).setValue(mapped.shopCode);
  if (mapped.memo)         sheet.getRange(rowIndex, COL.MEMO+1).setValue(mapped.memo);
  sheet.getRange(rowIndex, COL.UPDATED+1).setValue(new Date());
}

// ============================================================
// フォームマッピング
// ============================================================
function mapFormRow(row, headers, shopCode) {
  var result = { customerName:'', furigana:'', phone:'', email:'', birthDate:'', skinType:'', concerns:'', memo:'', shopCode:shopCode };
  var lastName='', firstName='', lastKana='', firstKana='';
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    var v = row[i] !== undefined ? String(row[i]).trim() : '';
    if (!v) continue;
    if (h === 'お名前（姓）' || h === '姓') lastName = v;
    else if (h === 'お名前（名）' || h === '名') firstName = v;
    else if ((h.indexOf('名前') >= 0 || h.indexOf('氏名') >= 0) && !lastName) result.customerName = v;
    else if (h === 'フリガナ（セイ）' || h === 'セイ') lastKana = v;
    else if (h === 'フリガナ（メイ）' || h === 'メイ') firstKana = v;
    else if ((h.indexOf('フリガナ') >= 0 || h.indexOf('ふりがな') >= 0) && !lastKana) result.furigana = v;
    else if (h.indexOf('電話') >= 0 || h === 'TEL') result.phone = v.replace(/[-\s\(\)]/g, '');
    else if (h.indexOf('メール') >= 0 || h.indexOf('mail') >= 0 || h === 'Email') result.email = v;
    else if (h.indexOf('生年月日') >= 0 || h.indexOf('誕生日') >= 0) result.birthDate = v;
    else if (h === '肌タイプ' || h === '肌質' || h === 'スキンタイプ' || h === '肌のタイプ') result.skinType = v;
    else if (h.indexOf('アレルギー') >= 0 || h.indexOf('ケロイド') >= 0) {
      if (v !== 'いいえ' && v !== 'なし') result.concerns = result.concerns ? result.concerns + ' / ' + v : v;
    }
    else if (h.indexOf('お悩み') >= 0 || h.indexOf('悩み') >= 0 || h.indexOf('気になる') >= 0) {
      result.concerns = result.concerns ? result.concerns + ' / ' + v : v;
    }
    else if (h.indexOf('目的') >= 0 || h.indexOf('ご来店') >= 0 || h.indexOf('きっかけ') >= 0) {
      result.memo = result.memo ? result.memo + ' / ' + v : v;
    }
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
  if (!message) return jsonError('message required');
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var targetLineId = '', customerName = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.ID]) === String(customerId)) {
      targetLineId = String(data[i][COL.LINE_ID] || '');
      customerName = String(data[i][COL.NAME] || '');
      break;
    }
  }
  if (!targetLineId) return jsonError('LINE IDが設定されていません');
  var token = getToken();
  if (!token) return jsonError('LINE_CHANNEL_ACCESS_TOKEN未設定');
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({ to: targetLineId, messages: [{ type:'text', text:message }] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) return jsonResponse({ ok: true, sentTo: customerName });
  return jsonError('LINE API error: ' + res.getContentText());
}

// ============================================================
// 契約管理
// ============================================================
function getContracts(e) {
  var customerId = e.parameter.customerId;
  if (!customerId) return jsonError('customerId required');
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('契約情報');
  if (!sheet) return jsonResponse([]);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === customerId) {
      result.push({
        contractId:    String(data[i][0]),
        contractDate:  data[i][1] ? Utilities.formatDate(new Date(data[i][1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
        shopCode:      String(data[i][2] || ''),
        courseName:    String(data[i][3] || ''),
        sessions:      Number(data[i][4] || 0),
        unitPrice:     Number(data[i][5] || 0),
        totalAmount:   Number(data[i][6] || 0),
        paymentMethod: String(data[i][7] || ''),
        status:        String(data[i][8] || '有効'),
        note:          String(data[i][9] || '')
      });
    }
  }
  return jsonResponse(result);
}

function addContractFromGet(params) {
  return addContract(params);
}

function addContract(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('契約情報');
  if (!sheet) {
    sheet = ss.insertSheet('契約情報');
    sheet.appendRow(['顧客ID','契約日','店舗コード','コース名','回数','単価','合計金額','支払方法','ステータス','備考']);
  }
  var sessions = Number(data.sessions || 0);
  var unitPrice = Number(data.unitPrice || 0);
  sheet.appendRow([
    data.customerId || '',
    data.contractDate || new Date(),
    data.shopCode || '',
    data.courseName || '',
    sessions,
    unitPrice,
    sessions * unitPrice,
    data.paymentMethod || '',
    data.status || '有効',
    data.note || ''
  ]);
  return jsonResponse({ ok: true });
}

// ============================================================
// 検索ユーティリティ
// ============================================================
function findCustomerByPhone(phone) {
  if (!phone) return null;
  var norm = phone.replace(/[-\s]/g, '');
  var norm10 = norm.replace(/^0/, '');
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var cell = String(data[i][COL.PHONE] || '').replace(/[-\s]/g, '');
    if (cell && (cell === norm || cell.replace(/^0/,'') === norm10)) {
      return { rowIndex:i+1, customerId:String(data[i][COL.ID]), customerName:String(data[i][COL.NAME]), phone:String(data[i][COL.PHONE]) };
    }
  }
  return null;
}

function findCustomerByLineId(lineId) {
  if (!lineId) return null;
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.LINE_ID]) === lineId) {
      return { rowIndex:i+1, customerId:String(data[i][COL.ID]), customerName:String(data[i][COL.NAME]) };
    }
  }
  return null;
}

function findCustomerByFormTS(ts) {
  if (!ts) return null;
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.FORM_TS]) === ts) {
      return { rowIndex:i+1, customerId:String(data[i][COL.ID]) };
    }
  }
  return null;
}

// ============================================================
// ショップ設定ユーティリティ
// ============================================================
function getShopByCode(code) {
  if (!code) return SHOPS[0]; // デフォルトは1番目の店舗
  for (var i = 0; i < SHOPS.length; i++) {
    if (SHOPS[i].code === code) return SHOPS[i];
    if (SHOPS[i].aliases && SHOPS[i].aliases.indexOf(code) >= 0) return SHOPS[i];
  }
  return SHOPS[0]; // 見つからない場合もデフォルト（エラーにしない）
}

function getShopBySheetName(name) {
  return SHOPS.filter(function(s){ return s.sheetName === name; })[0] || null;
}

// ============================================================
// フォームデータ取得
// ============================================================
function getRawFormData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allRows = [];
  SHOPS.forEach(function(shop) {
    var sheet = ss.getSheetByName(shop.sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    rows.forEach(function(row) {
      var obj = { _shopCode:shop.code, _shopLabel:shop.label };
      headers.forEach(function(h, c) {
        var key = String(h).trim();
        if (key) obj[key] = row[c] !== undefined ? String(row[c]) : '';
      });
      allRows.push(obj);
    });
  });
  return jsonResponse(allRows);
}

function getFormHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {};
  SHOPS.forEach(function(shop) {
    var sheet = ss.getSheetByName(shop.sheetName);
    if (!sheet) return;
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].filter(function(h){return h!=='';});
    result[shop.code] = { shopLabel:shop.label, headers:headers };
  });
  return jsonResponse(result);
}

function getLineUsers() {
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_DB_SHEET).getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][COL.LINE_ID]) {
      result.push({ customerId:String(data[i][COL.ID]), name:String(data[i][COL.NAME]), phone:String(data[i][COL.PHONE]||''), lineUserId:String(data[i][COL.LINE_ID]) });
    }
  }
  return jsonResponse(result);
}

function getConfig() {
  return jsonResponse({ shops: SHOPS, customerDbSheet: CUSTOMER_DB_SHEET });
}

// ============================================================
// 日付・年齢ユーティリティ
// ============================================================
function formatDate(val) {
  if (!val) return '';
  try {
    var d = val instanceof Date ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch(e) { return String(val); }
}

function calcAge(birthVal) {
  if (!birthVal) return '';
  try {
    var birth = birthVal instanceof Date ? birthVal : new Date(String(birthVal).replace(/\//g,'-'));
    if (isNaN(birth.getTime())) return '';
    var today = new Date();
    var age = today.getFullYear() - birth.getFullYear();
    if (today.getMonth() < birth.getMonth() || (today.getMonth()===birth.getMonth() && today.getDate()<birth.getDate())) age--;
    return (age>=0 && age<120) ? age : '';
  } catch(e) { return ''; }
}

// ============================================================
// ID生成
// ============================================================
function generateId() {
  var now = new Date();
  var pad = function(n,l){ return ('000'+n).slice(-(l||2)); };
  return 'C' + now.getFullYear() + pad(now.getMonth()+1) + pad(now.getDate()) + pad(now.getHours()) + pad(now.getMinutes()) + pad(now.getSeconds()) + pad(now.getMilliseconds(),3);
}

// ============================================================
// LINE APIユーティリティ
// ============================================================
function getToken() {
  return PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
}

function pushLine(token, userId, message) {
  if (!token || !userId) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method:'post', contentType:'application/json',
    headers:{'Authorization':'Bearer '+token},
    payload:JSON.stringify({to:userId, messages:[{type:'text',text:message}]}),
    muteHttpExceptions:true
  });
}

function replyLine(token, replyToken, message) {
  if (!token || !replyToken) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method:'post', contentType:'application/json',
    headers:{'Authorization':'Bearer '+token},
    payload:JSON.stringify({replyToken:replyToken, messages:[{type:'text',text:message}]}),
    muteHttpExceptions:true
  });
}

// ============================================================
// レスポンスユーティリティ
// ============================================================
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify({ok:true, data:obj})).setMimeType(ContentService.MimeType.JSON);
}

function jsonError(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false, error:msg})).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 初期設定（初回のみ実行）
// ============================================================
function setupProperties() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('LINE_CHANNEL_ACCESS_TOKEN', 'lEFEcdsU7W00c0nexEy0q5bVgzwa6PSknzbieVxTz16xx6UZ9hJ4fNssaNv32mrTRayAeHqKL6lrV1XCdr26vy8kgvwvoaKqb5do/QIlV7c5pEzMJFRKbEhaA6gZkBIckhTnKXkEb1xkJ6Oaf3aepAdB04t89/1O/w1cDnyilFU=');
  props.setProperty('LINE_CHANNEL_SECRET', '6ab448d0c63c2635f3ca8e602e4afd90');
  console.log('プロパティ設定完了');
  return { ok:true };
}
