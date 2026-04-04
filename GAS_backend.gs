var SPREADSHEET_ID = '1CLYVwTISKxHFc583wFNCIIQ_unLFwqjHS-SxOPTYXgE';

var FORM_SHEETS_CONFIG = [
  {
    sheetName: '【西船橋店】オンダリフト',
    shopCode: 'ONDARI_NISHIFUNA',
    shopLabel: '西船橋店 オンダリフト',
    lineChannel: 'LINE_NISHIFUNA'
  },
  {
    sheetName: '【西船橋店】ピーリング',
    shopCode: 'PEELING_NISHIFUNA',
    shopLabel: '西船橋店 ピーリング',
    lineChannel: 'LINE_NISHIFUNA'
  }
];

var CUSTOMER_DB_SHEET = '顧客DB';

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'getCustomers') return getCustomers(e);
  if (action === 'getCustomerDetail') return getCustomerDetail(e);
  if (action === 'getContracts') return getContracts(e);
  if (action === 'getConfig') return getConfig(e);
  if (action === 'syncAllFormResponses') return syncAllFormResponses();
  if (action === 'updateCustomer') return updateExistingCustomer(e.parameter);
  if (action === 'cleanDuplicates') return cleanDuplicates();
  if (action === 'getFormHeaders') return getFormHeaders();
  if (action === 'getRawFormData') return getRawFormData();
  if (action === 'linkLineId') return linkLineIdByPhone(e.parameter);
  if (action === 'getLineUsers') return getLineUsers();
  if (action === 'sendLine') return sendLineFromDashboard(e.parameter);
  if (action === 'cleanCustomerDB') return cleanCustomerDB();
  return jsonResponse({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.events) {
      var shopCode = (e.parameter && e.parameter.shop) ? e.parameter.shop : 'NISHIFUNA';
      var events = body.events;
      for (var i = 0; i < events.length; i++) {
        var event = events[i];
        if (event.type === 'follow') handleFollowEvent(event, shopCode);
        if (event.type === 'message') handleMessageEvent(event);
      }
      return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var action = body.action;
    if (action === 'handleFollow') return handleFollow(body);
    if (action === 'handleAutoReply') return handleAutoReply(body);
    if (action === 'addTreatment') return addTreatment(body);
    if (action === 'addContract') return addContract(body);
    if (action === 'saveClosingResult') return saveClosingResult(body);
    if (action === 'addCustomerManual') return addCustomerManual(body);
    if (action === 'sendLineFromDashboard') return sendLineFromDashboard(body);
    if (action === 'updateExistingCustomer') return updateExistingCustomer(body);
    if (action === 'initializeSheets') return initializeSheets();
    return jsonResponse({error: 'Unknown action: ' + action});

  } catch(err) {
    console.error('doPost error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleFollowEvent(event, shopCode) {
  try {
    var userId = event.source.userId;
    if (!userId) return;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
    if (!sheet) return;

    var existing = findCustomerByLineId(userId);
    if (existing) {
      console.log('既存顧客のLINE再連携: ' + existing.customerId);
    } else {
      var newId = generateCustomerId();
      var now = new Date();
      // 列順: A:ID, B:氏名, C:よみがな, D:電話, E:メール, F:生年月日, G:肌タイプ, H:お悩み, I:店舗, J:LINE_userId, K:LINE流入日時, L:ステータス, M:メモ, N:登録日時, O:最終更新
      sheet.appendRow([
        newId,              // A: 顧客ID
        'LINE新規',          // B: 氏名
        '',                 // C: よみがな
        '',                 // D: 電話番号
        '',                 // E: メールアドレス
        '',                 // F: 生年月日
        '',                 // G: 肌タイプ
        '',                 // H: お悩み
        shopCode || '',     // I: 店舗コード
        userId,             // J: LINE_userId
        now,                // K: LINE流入日時
        'LINE新規',          // L: ステータス
        '',                 // M: メモ
        now,                // N: 登録日時
        ''                  // O: 最終更新
      ]);
      console.log('新規LINE顧客登録: ' + newId + ' userId=' + userId + ' shop=' + shopCode);
    }

    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (!token) return;

    var shopConfig = null;
    for (var i = 0; i < FORM_SHEETS_CONFIG.length; i++) {
      if (FORM_SHEETS_CONFIG[i].shopCode === shopCode) {
        shopConfig = FORM_SHEETS_CONFIG[i];
        break;
      }
    }
    var shopLabel = shopConfig ? shopConfig.shopLabel : '当サロン';

    pushLineMessage(token, userId, shopLabel + 'にご登録ありがとうございます！\n\nカウンセリングシートのご記入をお願いいたします。\nご記入後、こちらにご登録の電話番号を送っていただくと担当スタッフとスムーズに連絡が取れます。\n例：09012345678');
  } catch(err) {
    console.error('handleFollowEvent error: ' + err.message);
  }
}

function handleMessageEvent(event) {
  try {
    var userId = event.source.userId;
    var replyToken = event.replyToken;
    var messageText = '';
    if (event.message && event.message.type === 'text') {
      messageText = String(event.message.text).trim();
    }
    if (!userId || !messageText) return;

    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (!token) return;

    var existing = findCustomerByLineId(userId);

    var phoneNorm = messageText.replace(/[-\s\(\)]/g, '');
    var isPhone = /^(0\d{9,10}|\d{10,11})$/.test(phoneNorm);

    if (isPhone) {
      var matched = findCustomerByPhone(phoneNorm);
      if (matched) {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(matched.customerId)) {
            sheet.getRange(i + 1, 10).setValue(userId);  // J列=LINE_userId
            break;
          }
        }

        if (existing && existing.customerId !== matched.customerId) {
          for (var j = 1; j < data.length; j++) {
            if (String(data[j][0]) === String(existing.customerId)) {
              sheet.deleteRow(j + 1);
              break;
            }
          }
        }

        replyLineMessage(token, replyToken, [{
          type: 'text',
          text: matched.customerName + '様、LINE連携が完了しました！\nスタッフからのご連絡をお待ちください。'
        }]);
        console.log('LINE ID紐付け完了: ' + matched.customerId + ' userId=' + userId);
        return;
      } else {
        replyLineMessage(token, replyToken, [{
          type: 'text',
          text: '電話番号が見つかりませんでした。\nご来店時にスタッフにお申し付けください。'
        }]);
        return;
      }
    }

    if (!existing || !existing.customerName || existing.customerName === 'LINE新規') {
      replyLineMessage(token, replyToken, [{
        type: 'text',
        text: 'メッセージありがとうございます。\nご登録のお客様の電話番号を送っていただくと、担当スタッフとスムーズに連絡が取れます。\n例：09012345678'
      }]);
    } else {
      replyLineMessage(token, replyToken, [{
        type: 'text',
        text: existing.customerName + '様、メッセージありがとうございます。\n担当スタッフより折り返しご連絡いたします。'
      }]);
    }
  } catch(err) {
    console.error('handleMessageEvent error: ' + err.message);
  }
}

function cleanUpLineNewRecord(lineUserId, keepRowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (i + 1 === keepRowIndex) continue;
      if (data[i][9] === lineUserId && data[i][1] === 'LINE新規') {
        sheet.deleteRow(i + 1);
        console.log('LINE新規レコード削除: row=' + (i + 1));
      }
    }
  } catch(err) {
    console.error('cleanUpLineNewRecord error: ' + err.message);
  }
}

function handleFollow(data) {
  var lineId = data.lineUserId || '';
  var displayName = data.displayName || '';
  if (!lineId) return jsonResponse({ error: 'lineUserId is required' });
  var existing = findCustomerByLineId(lineId);
  if (existing) {
    return jsonResponse({ status: 'already_exists', customerId: existing.customerId });
  }
  var newId = generateCustomerId();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  // 列順: A:ID, B:氏名, C:よみがな, D:電話, E:メール, F:生年月日, G:肌タイプ, H:お悩み, I:店舗, J:LINE_userId, K:LINE流入日時, L:ステータス, M:メモ, N:登録日時, O:最終更新
  sheet.appendRow([
    newId,              // A: 顧客ID
    displayName,        // B: 氏名
    '',                 // C: よみがな
    '',                 // D: 電話番号
    '',                 // E: メールアドレス
    '',                 // F: 生年月日
    '',                 // G: 肌タイプ
    '',                 // H: お悩み
    '',                 // I: 店舗コード
    lineId,             // J: LINE_userId
    new Date(),         // K: LINE流入日時
    'LINE友だち追加',     // L: ステータス
    '',                 // M: メモ
    new Date(),         // N: 登録日時
    ''                  // O: 最終更新
  ]);
  return jsonResponse({ status: 'new_customer', customerId: newId });
}

function handleAutoReply(data) {
  var lineId = data.lineUserId || '';
  var message = data.message || '';
  var customer = findCustomerByLineId(lineId);
  var replyText = '';
  if (!customer) {
    replyText = 'ご登録がありません。お名前とお電話番号をお知らせください。';
  } else {
    replyText = customer.customerName + '様、メッセージありがとうございます。担当者より折り返しご連絡いたします。';
  }
  if (data.replyToken) {
    replyLineMessage(data.replyToken, replyText);
  }
  return jsonResponse({ status: 'ok', reply: replyText });
}

function getCustomers(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse([]);
  var customers = [];

  var treatSheet = ss.getSheetByName('施術履歴');
  var treatData = treatSheet ? treatSheet.getDataRange().getValues() : [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var age = '';
    if (row[5]) {
      try {
        var birthStr = String(row[5]);
        var birth;
        if (birthStr.indexOf('/') !== -1) {
          var parts = birthStr.split('/');
          birth = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        } else if (birthStr.indexOf('-') !== -1) {
          var parts = birthStr.split('-');
          birth = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        } else {
          birth = new Date(row[5]);
        }
        var today = new Date();
        var a = today.getFullYear() - birth.getFullYear();
        var m = today.getMonth() - birth.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birth.getDate())) a--;
        age = (a >= 0 && a < 120) ? a : '';
      } catch(err) {
        age = '';
      }
    }

    var customerId = row[0] || '';
    var treatmentCount = 0;
    var lastVisit = '';
    for (var t = 1; t < treatData.length; t++) {
      if (String(treatData[t][0]) === String(customerId)) {
        treatmentCount++;
        var tDate = treatData[t][1] ? Utilities.formatDate(new Date(treatData[t][1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '';
        if (tDate > lastVisit) lastVisit = tDate;
      }
    }

    customers.push({
      id: customerId,
      name: row[1] || '',
      furigana: row[2] || '',
      phone: row[3] || '',
      email: row[4] || '',
      age: age,
      skinType: row[6] || '',
      allergies: row[7] || '',
      shop: row[8] || '',
      lineUserId: row[9] || '',
      lineInflowDate: row[10] ? Utilities.formatDate(new Date(row[10]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '',
      status: row[11] || '',
      memo: row[12] || '',
      registeredDate: row[13] ? Utilities.formatDate(new Date(row[13]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
      lastVisit: lastVisit,
      treatmentCount: treatmentCount,
      concerns: [],
      staff: '',
      contractId: ''
    });
  }

  var keyword = (e && e.parameter && e.parameter.keyword) ? e.parameter.keyword : '';
  if (keyword) {
    customers = customers.filter(function(c) {
      return (c.name && c.name.indexOf(keyword) >= 0) ||
             (c.phone && c.phone.indexOf(keyword) >= 0) ||
             (c.id && c.id.indexOf(keyword) >= 0);
    });
  }
  return jsonResponse(customers);
}

function getCustomerDetail(e) {
  var customerId = e.parameter.customerId;
  if (!customerId) return jsonResponse({ error: 'customerId is required' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === customerId) {
      var row = data[i];
      var customer = {
        customerId: row[0] || '',
        customerName: row[1] || '',
        furigana: row[2] || '',
        phone: row[3] || '',
        email: row[4] || '',
        birthDate: row[5] ? Utilities.formatDate(new Date(row[5]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
        skinType: row[6] || '',
        allergies: row[7] || '',
        shopCode: row[8] || '',
        lineId: row[9] || '',
        lineInflowDate: row[10] ? Utilities.formatDate(new Date(row[10]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '',
        status: row[11] || '',
        memo: row[12] || '',
        registrationDate: row[13] ? Utilities.formatDate(new Date(row[13]), 'Asia/Tokyo', 'yyyy-MM-dd') : ''
      };
      var treatments = [];
      var treatSheet = ss.getSheetByName('施術履歴');
      if (treatSheet) {
        var tData = treatSheet.getDataRange().getValues();
        for (var j = 1; j < tData.length; j++) {
          if (tData[j][0] === customerId) {
            treatments.push({
              date: tData[j][1] ? Utilities.formatDate(new Date(tData[j][1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
              shopCode: tData[j][2] || '',
              menuName: tData[j][3] || '',
              staff: tData[j][4] || '',
              note: tData[j][5] || '',
              beforePhoto: tData[j][6] || '',
              afterPhoto: tData[j][7] || ''
            });
          }
        }
      }
      return jsonResponse({ customer: customer, treatments: treatments });
    }
  }
  return jsonResponse({ error: 'Customer not found' });
}

function syncAllFormResponses() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var totalSynced = 0;
  for (var f = 0; f < FORM_SHEETS_CONFIG.length; f++) {
    var config = FORM_SHEETS_CONFIG[f];
    var formSheet = ss.getSheetByName(config.sheetName);
    if (!formSheet) continue;
    var formData = formSheet.getDataRange().getValues();
    if (formData.length <= 1) continue;
    for (var i = 1; i < formData.length; i++) {
      var mapped = mapFormRow(formData[i], formData[0], config.shopCode);
      if (!mapped.customerName && !mapped.phone) continue;
      var existing = null;
      if (mapped.phone) {
        existing = findCustomerByPhone(mapped.phone);
      }
      if (!existing && mapped.customerName) {
        existing = findCustomerByName(mapped.customerName);
      }
      if (existing) {
        updateExistingCustomerFromForm(existing.rowIndex, mapped);
      } else {
        addNewCustomer(mapped);
      }
      totalSynced++;
    }
  }
  return jsonResponse({ status: 'ok', totalSynced: totalSynced });
}

function onFormSubmit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var sheetName = sheet.getName();
    var shopCode = '';
    for (var f = 0; f < FORM_SHEETS_CONFIG.length; f++) {
      if (FORM_SHEETS_CONFIG[f].sheetName === sheetName) {
        shopCode = FORM_SHEETS_CONFIG[f].shopCode;
        break;
      }
    }
    if (!shopCode) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = e.values || sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapped = mapFormRow(row, headers, shopCode);
    if (!mapped || (!mapped.phone && !mapped.customerName)) return;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dbSheet = ss.getSheetByName(CUSTOMER_DB_SHEET);

    var existing = null;
    if (mapped.phone) existing = findCustomerByPhone(mapped.phone);
    if (!existing && mapped.customerName) existing = findCustomerByName(mapped.customerName);

    var customerId = '';
    if (existing) {
      updateExistingCustomerFromForm(existing.rowIndex, mapped);
      customerId = existing.customerId;
    } else {
      customerId = addNewCustomer(mapped);
    }

    if (!customerId) return;

    var lineNewRecords = [];
    var allData = dbSheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][11]) === 'LINE友だち追加' &&
          (!allData[i][1] || allData[i][1] === 'LINE新規') &&
          allData[i][9]) {
        lineNewRecords.push({
          rowIndex: i + 1,
          lineUserId: allData[i][9],
          customerId: allData[i][0]
        });
      }
    }

    if (lineNewRecords.length === 1) {
      var lineRecord = lineNewRecords[0];
      var allData2 = dbSheet.getDataRange().getValues();
      for (var j = 1; j < allData2.length; j++) {
        if (String(allData2[j][0]) === String(customerId)) {
          dbSheet.getRange(j + 1, 10).setValue(lineRecord.lineUserId);  // J列=LINE_userId
          console.log('フォーム回答でLINE自動紐づけ完了: customerId=' + customerId + ' lineUserId=' + lineRecord.lineUserId);
          break;
        }
      }
      dbSheet.deleteRow(lineRecord.rowIndex);

      var props = PropertiesService.getScriptProperties();
      var token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
      if (token && mapped.customerName) {
        pushLineMessage(token, lineRecord.lineUserId,
          mapped.customerName + '様、カウンセリングシートのご記入ありがとうございます！\n担当スタッフより改めてご連絡いたします。');
      }
    } else if (lineNewRecords.length > 1) {
      console.log('LINE新規が複数件あるため自動紐づけをスキップ: ' + lineNewRecords.length + '件');
    }
  } catch(err) {
    console.error('onFormSubmit error: ' + err.message);
  }
}

function mapFormRow(row, headers, shopCode) {
  var result = {
    customerName: '',
    furigana: '',
    phone: '',
    email: '',
    birthDate: '',
    gender: '',
    address: '',
    skinType: '',
    allergies: '',
    memo: '',
    shopCode: shopCode,
    source: 'フォーム回答',
    registrationDate: new Date()
  };
  var lastName = '';
  var firstName = '';
  var lastKana = '';
  var firstKana = '';
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    var v = row[i] !== undefined ? String(row[i]).trim() : '';
    if (h === 'お名前（姓）' || h === '姓' || h === '名字') lastName = v;
    else if (h === 'お名前（名）' || h === '名' || h === '下の名前') firstName = v;
    else if (h.indexOf('名前') >= 0 || h.indexOf('氏名') >= 0 || h === 'お名前') {
      if (!lastName) result.customerName = v;
    }
    else if (h === 'フリガナ（セイ）' || h === 'セイ') lastKana = v;
    else if (h === 'フリガナ（メイ）' || h === 'メイ') firstKana = v;
    else if (h.indexOf('フリガナ') >= 0 || h.indexOf('ふりがな') >= 0 || h.indexOf('よみがな') >= 0) {
      if (!lastKana) result.furigana = v;
    }
    else if (h.indexOf('電話') >= 0 || h === 'TEL' || h === 'tel' || h.indexOf('携帯') >= 0) {
      result.phone = v.replace(/[-\s]/g, '');
    }
    else if (h.indexOf('メール') >= 0 || h.indexOf('mail') >= 0 || h.indexOf('Mail') >= 0 || h === 'Email') {
      result.email = v;
    }
    else if (h.indexOf('生年月日') >= 0 || h.indexOf('誕生日') >= 0) {
      result.birthDate = v;
    }
    else if (h.indexOf('年齢') >= 0) {
      if (!result.birthDate) result.ageRaw = v;
    }
    else if (h.indexOf('性別') >= 0) result.gender = v;
    else if (h.indexOf('住所') >= 0 || h.indexOf('郵便') >= 0) result.address = v;
    // 肌タイプ：「肌タイプ」「肌質」など短い項目名のみ対象（「肌が敏感〜症状はありますか？」は対象外）
    else if (h === '肌タイプ' || h === '肌質' || h === 'スキンタイプ' ||
             (h.indexOf('スキン') >= 0 && h.indexOf('タイプ') >= 0)) {
      result.skinType = v;
    }
    // アレルギー・ケロイド・肌症状・敏感肌はアレルギー欄へ（「いいえ」「なし」は除外）
    else if (h.indexOf('アレルギー') >= 0 || h.indexOf('ケロイド') >= 0 ||
             (h.indexOf('肌') >= 0 && h.indexOf('症状') >= 0) ||
             (h.indexOf('肌') >= 0 && h.indexOf('敏感') >= 0)) {
      if (v && v !== 'いいえ' && v !== 'なし') {
        result.allergies = result.allergies ? result.allergies + ' / ' + v : v;
      }
    }
    // 症状の詳細（「はいとお答えの方のみ症状を教えてください」など）
    else if (h.indexOf('症状') >= 0 && h.indexOf('教えて') >= 0) {
      if (v && v !== 'いいえ' && v !== 'なし') {
        result.allergies = result.allergies ? result.allergies + '（' + v + '）' : v;
      }
    }
    else if (h.indexOf('お悩み') >= 0 || h.indexOf('悩み') >= 0 || h.indexOf('ご要望') >= 0) {
      result.memo = result.memo ? result.memo + ' / ' + v : v;
    }
    else if (h.indexOf('目的') >= 0 || h.indexOf('ご来店') >= 0) {
      result.memo = result.memo ? result.memo + ' / ' + v : v;
    }
    else if (h.indexOf('備考') >= 0) result.memo = (result.memo ? result.memo + ' ' : '') + v;
  }
  if (lastName || firstName) {
    result.customerName = (lastName + ' ' + firstName).trim();
  }
  if (lastKana || firstKana) {
    result.furigana = (lastKana + ' ' + firstKana).trim();
  }
  if (result.birthDate) {
    var bd = result.birthDate.replace(/[-\/]/g, '');
    if (/^\d{8}$/.test(bd)) {
      result.birthDate = bd.slice(0,4) + '/' + bd.slice(4,6) + '/' + bd.slice(6,8);
    }
  }
  return result;
}

function findCustomerByPhone(phone) {
  if (!phone) return null;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var normalizedPhone = phone.replace(/[-\s]/g, '');
  for (var i = 1; i < data.length; i++) {
    var cellPhone = String(data[i][3] || '').replace(/[-\s]/g, '');
    if (cellPhone && cellPhone === normalizedPhone) {
      return {
        rowIndex: i + 1,
        customerId: data[i][0],
        customerName: data[i][1],
        phone: data[i][3],
        lineId: data[i][9] || ''
      };
    }
  }
  return null;
}

function findCustomerByLineId(lineId) {
  if (!lineId) return null;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === lineId) {
      return {
        rowIndex: i + 1,
        customerId: data[i][0],
        customerName: data[i][1],
        phone: data[i][3],
        lineId: data[i][9]
      };
    }
  }
  return null;
}

function generateCustomerId() {
  var now = new Date();
  var y = now.getFullYear();
  var m = ('0' + (now.getMonth() + 1)).slice(-2);
  var d = ('0' + now.getDate()).slice(-2);
  var h = ('0' + now.getHours()).slice(-2);
  var min = ('0' + now.getMinutes()).slice(-2);
  var s = ('0' + now.getSeconds()).slice(-2);
  var ms = ('00' + now.getMilliseconds()).slice(-3);
  return 'C' + y + m + d + h + min + s + ms;
}

function addNewCustomer(mapped) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var newId = generateCustomerId();
  // 列順: A:ID, B:氏名, C:よみがな, D:電話, E:メール, F:生年月日, G:肌タイプ, H:お悩み, I:店舗, J:LINE_userId, K:LINE流入日時, L:ステータス, M:メモ, N:登録日時, O:最終更新
  sheet.appendRow([
    newId,                          // A: 顧客ID
    mapped.customerName || '',      // B: 氏名
    mapped.furigana || '',          // C: よみがな
    mapped.phone || '',             // D: 電話番号
    mapped.email || '',             // E: メールアドレス
    mapped.birthDate || '',         // F: 生年月日
    mapped.skinType || '',          // G: 肌タイプ
    mapped.allergies || '',         // H: お悩み
    mapped.shopCode || '',          // I: 店舗コード
    mapped.lineId || '',            // J: LINE_userId
    '',                             // K: LINE流入日時（空）
    '新規',                          // L: ステータス（デフォルト新規）
    mapped.memo || '',              // M: メモ
    mapped.registrationDate || new Date(), // N: 登録日時
    new Date()                      // O: 最終更新
  ]);
  return newId;
}

function updateExistingCustomer(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var allData = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(data.customerId)) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex < 0) return jsonResponse({ error: 'Customer not found' });
  // B列=2:氏名, D列=4:電話, E列=5:メール, F列=6:生年月日, G列=7:肌タイプ, M列=13:メモ, O列=15:最終更新
  if (data.name && data.name !== '') sheet.getRange(rowIndex, 2).setValue(data.name);
  if (data.phone && data.phone !== '') sheet.getRange(rowIndex, 4).setValue(data.phone);
  if (data.email && data.email !== '') sheet.getRange(rowIndex, 5).setValue(data.email);
  if (data.birthDate && data.birthDate !== '') sheet.getRange(rowIndex, 6).setValue(data.birthDate);
  if (data.skinType !== undefined) sheet.getRange(rowIndex, 7).setValue(data.skinType || '');
  if (data.memo !== undefined) sheet.getRange(rowIndex, 13).setValue(data.memo || '');
  sheet.getRange(rowIndex, 15).setValue(new Date());
  return jsonResponse({ status: 'updated', customerId: data.customerId });
}

function updateExistingCustomerFromForm(rowIndex, mapped) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  // B列=2:氏名, E列=5:メール, F列=6:生年月日, G列=7:肌タイプ, H列=8:お悩み, I列=9:店舗, M列=13:メモ, O列=15:最終更新
  if (mapped.customerName) sheet.getRange(rowIndex, 2).setValue(mapped.customerName);
  if (mapped.furigana) sheet.getRange(rowIndex, 3).setValue(mapped.furigana);
  if (mapped.email) sheet.getRange(rowIndex, 5).setValue(mapped.email);
  if (mapped.birthDate) sheet.getRange(rowIndex, 6).setValue(mapped.birthDate);
  if (mapped.skinType) sheet.getRange(rowIndex, 7).setValue(mapped.skinType);
  if (mapped.allergies) sheet.getRange(rowIndex, 8).setValue(mapped.allergies);
  if (mapped.shopCode) sheet.getRange(rowIndex, 9).setValue(mapped.shopCode);
  if (mapped.memo) sheet.getRange(rowIndex, 13).setValue(mapped.memo);
  sheet.getRange(rowIndex, 15).setValue(new Date());

  // LINE ID（J列=10）は絶対に上書きしない — 既存値を保護
  var currentLineId = sheet.getRange(rowIndex, 10).getValue();  // J列=LINE_userId
  if (currentLineId) {
    console.log('LINE ID保護: 既存LINE ID維持 rowIndex=' + rowIndex + ' lineId=' + currentLineId);
  } else if (mapped.phone) {
    // LINE ID未紐付けの場合のみ、LINE新規レコードから自動マッチング
    autoLinkLineIdFromNewRecord(rowIndex, mapped.phone);
  }
}

function autoLinkLineIdFromNewRecord(targetRowIndex, phone) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
    var data = sheet.getDataRange().getValues();
    // LINE新規レコード（name=LINE新規、lineUserId有、phone無）を探す
    // ※電話番号では直接マッチできないので、LINE新規が1件だけの場合に限り紐付け
    var lineNewRows = [];
    for (var i = 1; i < data.length; i++) {
      if (i + 1 === targetRowIndex) continue;
      if (data[i][1] === 'LINE新規' && data[i][9] && !data[i][3]) {
        lineNewRows.push({ rowIndex: i + 1, lineId: data[i][9] });
      }
    }
    if (lineNewRows.length === 1) {
      // LINE新規が1件だけなら確実にその人なので自動紐付け
      sheet.getRange(targetRowIndex, 10).setValue(lineNewRows[0].lineId);  // J列=LINE_userId
      sheet.deleteRow(lineNewRows[0].rowIndex);
      console.log('フォーム→LINE自動紐付け完了: lineId=' + lineNewRows[0].lineId);
    }
  } catch(err) {
    console.error('autoLinkLineIdFromNewRecord error: ' + err.message);
  }
}

function addTreatment(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('施術履歴');
  if (!sheet) {
    sheet = ss.insertSheet('施術履歴');
    sheet.appendRow(['顧客ID', '施術日', '店舗コード', 'メニュー名', '担当スタッフ', '備考', 'Before写真', 'After写真']);
  }
  sheet.appendRow([
    data.customerId || '',
    data.date || new Date(),
    data.shopCode || '',
    data.menuName || '',
    data.staff || '',
    data.note || '',
    data.beforePhoto || '',
    data.afterPhoto || ''
  ]);
  return jsonResponse({ status: 'ok' });
}

function addContract(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('契約情報');
  if (!sheet) {
    sheet = ss.insertSheet('契約情報');
    sheet.appendRow(['顧客ID', '契約日', '店舗コード', 'コース名', '回数', '金額', '支払方法', 'ステータス', '備考']);
  }
  sheet.appendRow([
    data.customerId || '',
    data.contractDate || new Date(),
    data.shopCode || '',
    data.courseName || '',
    data.sessions || '',
    data.amount || '',
    data.paymentMethod || '',
    data.status || '有効',
    data.note || ''
  ]);
  return jsonResponse({ status: 'ok' });
}

function getContracts(e) {
  var customerId = e.parameter.customerId;
  if (!customerId) return jsonResponse({ error: 'customerId is required' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('契約情報');
  if (!sheet) return jsonResponse({ contracts: [] });
  var data = sheet.getDataRange().getValues();
  var contracts = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === customerId) {
      contracts.push({
        customerId: data[i][0],
        contractDate: data[i][1] ? Utilities.formatDate(new Date(data[i][1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
        shopCode: data[i][2] || '',
        courseName: data[i][3] || '',
        sessions: data[i][4] || '',
        amount: data[i][5] || '',
        paymentMethod: data[i][6] || '',
        status: data[i][7] || '',
        note: data[i][8] || ''
      });
    }
  }
  return jsonResponse({ contracts: contracts });
}

function saveClosingResult(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('クロージング結果');
  if (!sheet) {
    sheet = ss.insertSheet('クロージング結果');
    sheet.appendRow(['顧客ID', '日付', '店舗コード', '提案コース', '結果', '金額', '備考']);
  }
  sheet.appendRow([
    data.customerId || '',
    data.date || new Date(),
    data.shopCode || '',
    data.courseName || '',
    data.result || '',
    data.amount || '',
    data.note || ''
  ]);
  return jsonResponse({ status: 'ok' });
}

function addCustomerManual(data) {
  var mapped = {
    customerName: data.customerName || '',
    phone: data.phone || '',
    email: data.email || '',
    birthDate: data.birthDate || '',
    gender: data.gender || '',
    address: data.address || '',
    skinType: data.skinType || '',
    allergies: data.allergies || '',
    memo: data.memo || '',
    lineId: data.lineId || '',
    shopCode: data.shopCode || '',
    source: data.source || '手動登録',
    registrationDate: new Date()
  };
  var newId = addNewCustomer(mapped);
  return jsonResponse({ status: 'ok', customerId: newId });
}

function sendLineFromDashboard(data) {
  var customerId = data.customerId || data.userId || '';
  var message = data.message || data.messages || '';
  var lineUserId = data.lineUserId || data.userId || '';

  if (!message) return jsonResponse({ error: 'message is required' });

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var allData = sheet.getDataRange().getValues();

  var targetLineId = lineUserId;
  var customerName = '';

  if (!targetLineId && customerId) {
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(customerId)) {
        targetLineId = allData[i][9] || '';
        customerName = allData[i][1] || '';
        break;
      }
    }
  }

  if (!targetLineId) return jsonResponse({ error: 'LINE IDが見つかりません。顧客詳細からLINE IDを設定してください。' });

  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return jsonResponse({ error: 'LINE_CHANNEL_ACCESS_TOKENが設定されていません' });

  var payload = {
    to: targetLineId,
    messages: [{ type: 'text', text: message }]
  };

  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var responseCode = res.getResponseCode();
  var responseText = res.getContentText();
  console.log('LINE API response: ' + responseCode + ' ' + responseText);

  if (responseCode === 200) {
    return jsonResponse({ ok: true, sentTo: customerName || targetLineId });
  } else {
    return jsonResponse({ error: 'LINE API エラー: ' + responseText });
  }
}

function pushLineMessage(tokenOrLineId, lineIdOrMessage, message) {
  var token, lineId, text;
  if (message !== undefined) {
    // 3引数形式: pushLineMessage(token, lineId, message)
    token = tokenOrLineId;
    lineId = lineIdOrMessage;
    text = message;
  } else {
    // 2引数形式: pushLineMessage(lineId, message)
    var props = PropertiesService.getScriptProperties();
    token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    lineId = tokenOrLineId;
    text = lineIdOrMessage;
  }
  if (!token) return;
  var payload = {
    to: lineId,
    messages: [{ type: 'text', text: text }]
  };
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function replyLineMessage(tokenOrReplyToken, replyTokenOrMessage, messages) {
  var token, replyToken, payload;
  if (messages !== undefined) {
    // 3引数形式: replyLineMessage(token, replyToken, messages[])
    token = tokenOrReplyToken;
    replyToken = replyTokenOrMessage;
    payload = { replyToken: replyToken, messages: messages };
  } else {
    // 2引数形式: replyLineMessage(replyToken, message string)
    var props = PropertiesService.getScriptProperties();
    token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    replyToken = tokenOrReplyToken;
    payload = { replyToken: replyToken, messages: [{ type: 'text', text: replyTokenOrMessage }] };
  }
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function getConfig(e) {
  return jsonResponse({
    formSheets: FORM_SHEETS_CONFIG,
    customerDbSheet: CUSTOMER_DB_SHEET,
    spreadsheetId: SPREADSHEET_ID
  });
}

function getFormHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {};
  for (var i = 0; i < FORM_SHEETS_CONFIG.length; i++) {
    var config = FORM_SHEETS_CONFIG[i];
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) continue;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    result[config.shopCode] = {
      shopLabel: config.shopLabel,
      headers: headers.filter(function(h) { return h !== ''; })
    };
  }
  return jsonResponse(result);
}

function getRawFormData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allRows = [];
  for (var i = 0; i < FORM_SHEETS_CONFIG.length; i++) {
    var config = FORM_SHEETS_CONFIG[i];
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) continue;
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var r = 0; r < rows.length; r++) {
      var obj = {
        _shopCode: config.shopCode,
        _shopLabel: config.shopLabel,
        _lineChannel: config.lineChannel
      };
      for (var c = 0; c < headers.length; c++) {
        var key = String(headers[c]).trim();
        if (key) obj[key] = rows[r][c] !== undefined ? String(rows[r][c]) : '';
      }
      allRows.push(obj);
    }
  }
  return jsonResponse(allRows);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data: obj }))
    .setMimeType(ContentService.MimeType.JSON);
}

function initializeSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(CUSTOMER_DB_SHEET);
    dbSheet.appendRow([
      '顧客ID', '顧客名', '電話番号', 'メールアドレス', '生年月日',
      '性別', '住所', '肌タイプ', 'アレルギー', 'メモ',
      'LINE ID', '登録日', '店舗コード', '流入元'
    ]);
  }
  var treatSheet = ss.getSheetByName('施術履歴');
  if (!treatSheet) {
    treatSheet = ss.insertSheet('施術履歴');
    treatSheet.appendRow(['顧客ID', '施術日', '店舗コード', 'メニュー名', '担当スタッフ', '備考', 'Before写真', 'After写真']);
  }
  var contractSheet = ss.getSheetByName('契約情報');
  if (!contractSheet) {
    contractSheet = ss.insertSheet('契約情報');
    contractSheet.appendRow(['顧客ID', '契約日', '店舗コード', 'コース名', '回数', '金額', '支払方法', 'ステータス', '備考']);
  }
  var closingSheet = ss.getSheetByName('クロージング結果');
  if (!closingSheet) {
    closingSheet = ss.insertSheet('クロージング結果');
    closingSheet.appendRow(['顧客ID', '日付', '店舗コード', '提案コース', '結果', '金額', '備考']);
  }
  return jsonResponse({ status: 'initialized' });
}

function cleanDuplicates() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    var phone = String(data[i][3] || '').replace(/[-\s]/g, '');
    var name = String(data[i][1] || '').replace(/\s/g, '');
    var key = phone || name;
    if (!key) continue;
    if (seen[key]) {
      rowsToDelete.push(i + 1);
    } else {
      seen[key] = true;
    }
  }
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  return jsonResponse({ deleted: rowsToDelete.length });
}

function findCustomerByName(name) {
  if (!name) return null;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var normalized = name.replace(/\s/g, '');
  for (var i = 1; i < data.length; i++) {
    var cellName = String(data[i][1] || '').replace(/\s/g, '');
    if (cellName && cellName === normalized) {
      return {
        rowIndex: i + 1,
        customerId: data[i][0],
        customerName: data[i][1],
        phone: data[i][3],
        lineId: data[i][9] || ''
      };
    }
  }
  return null;
}

function linkLineIdByPhone(params) {
  var phone = params.phone || '';
  var lineUserId = params.lineUserId || '';
  if (!phone || !lineUserId) return jsonResponse({ error: 'phone and lineUserId are required' });

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var normalizedPhone = phone.replace(/[-\s]/g, '');

  for (var i = 1; i < data.length; i++) {
    var cellPhone = String(data[i][3] || '').replace(/[-\s]/g, '');
    if (cellPhone && cellPhone === normalizedPhone) {
      sheet.getRange(i + 1, 10).setValue(lineUserId);  // J列=LINE_userId
      console.log('LINE ID紐付け完了: row=' + (i+1) + ' phone=' + phone + ' lineId=' + lineUserId);
      return jsonResponse({ ok: true, customerId: data[i][0], name: data[i][1] });
    }
  }
  return jsonResponse({ error: '電話番号が見つかりません: ' + phone });
}

function getLineUsers() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][9]) {
      result.push({
        customerId: data[i][0],
        name: data[i][1],
        phone: data[i][3],
        lineUserId: data[i][9],
        status: data[i][11]
      });
    }
  }
  return jsonResponse(result);
}

function cleanCustomerDB() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CUSTOMER_DB_SHEET);
  var data = sheet.getDataRange().getValues();

  var keepNames = ['林 治希', '新里 愛', '髙橋 将太', '儀間 理生', '氏家 大樹'];
  var rowsToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    if (!keepNames.includes(name)) {
      rowsToDelete.push(i + 1);
    }
  }

  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  var data2 = sheet.getDataRange().getValues();
  for (var k = 1; k < data2.length; k++) {
    var lineId = String(data2[k][9] || '');
    if (lineId && !lineId.startsWith('U')) {
      sheet.getRange(k + 1, 10).setValue('');
    }
    var status = String(data2[k][11] || '');
    if (status === 'フォーム回答' || status === 'LINE友だち追加' || status === '') {
      sheet.getRange(k + 1, 12).setValue('新規');
    }
  }

  console.log('クリーニング完了: ' + rowsToDelete.length + '行削除');
  return jsonResponse({ deleted: rowsToDelete.length });
}

function setupProperties() {
  var props = PropertiesService.getScriptProperties();
  var token = 'lEFEcdsU7W00c0nexEy0q5bVgzwa6PSknzbieVxTz16xx6UZ9hJ4fNssaNv32mrTRayAeHqKL6lrV1XCdr26vy8kgvwvoaKqb5do/QIlV7c5pEzMJFRKbEhaA6gZkBIckhTnKXkEb1xkJ6Oaf3aepAdB04t89/1O/w1cDnyilFU=';
  var secret = '6ab448d0c63c2635f3ca8e602e4afd90';
  props.setProperty('LINE_CHANNEL_ACCESS_TOKEN', token);
  props.setProperty('LINE_CHANNEL_SECRET', secret);
  console.log('プロパティ設定完了');
  return { ok: true };
}
