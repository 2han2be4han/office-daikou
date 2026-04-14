/**
 * Office DAIKOU - 人工出し請求書自動生成ツール
 * Google Apps Script 本体
 *
 * スクリプトプロパティに以下を設定してください：
 *   SHEET_ID            : GoogleスプレッドシートのID
 *   FORM_PASSWORD       : 応援業者用パスワード（index.html）
 *   FORM_PASSWORD_SELF  : 自社用パスワード（invoice.html）
 *   PARENT_COMPANY_EMAIL: 親会社担当者のメールアドレス
 */

// ===== 定数 =====
const SHEET_NAME_DATA_OWN = '自社の現場';
const SHEET_NAME_DATA_SUPPORT = '応援現場';
const SHEET_NAME_INVOICE = '請求書';
const SHEET_NAME_EMPLOYEES = '従業員';
const SHEET_NAME_TOKENS = '認証トークン';
const SHEET_NAME_CLIENT_SETTINGS = '取引先設定';
const SHEET_NAME_DATA = '入力データ'; // 旧互換

// ===== ヘルパー =====

function getProps() {
  return PropertiesService.getScriptProperties();
}

function getSpreadsheet() {
  const id = getProps().getProperty('SHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAME_DATA_OWN) {
      sheet.appendRow([
        'タイムスタンプ', '年', '月', '日', '現場名',
        '数量', '単価', '車台数', '高速代等', '小計', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_DATA_SUPPORT) {
      sheet.appendRow([
        'タイムスタンプ', '年', '月', '日', '送信者', 'あなたの会社名(元)', '応援先会社名(先)',
        '現場名', '人数', '作業員名', '単価', '車台数', '高速代等', '小計', 'コメント', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_INVOICE) {
      sheet.appendRow([
        '年', '月', '企業名', '合計金額', '請求書URL', '送信日時', '承認日時', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_EMPLOYEES) {
      sheet.appendRow(['名前']);
      const samples = ['Aさん', 'Bさん', 'Cさん', 'Dさん', 'Eさん', 'Fさん', 'Gさん', 'Hさん', 'Iさん', 'Jさん'];
      samples.forEach(n => sheet.appendRow([n]));
    }
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== Web App エンドポイント =====

function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  switch (action) {
    case 'auth':
      return handleAuth(e);
    case 'requestmagiclink':
      return handleRequestMagicLink(e);
    case 'verifytoken':
      return handleVerifyToken(e);
    case 'invoice':
      return handleGetInvoice(e);
    case 'invoicelist':
      return handleGetInvoiceList(e);
    case 'employees':
      return handleGetEmployees(e);
    case 'getclientemail':
      return handleGetClientEmail(e);
    default:
      return jsonResponse({ success: false, message: '不明なアクションです' });
  }
}

function doPost(e) {
  const action = (e.parameter && e.parameter.action || '').toLowerCase();

  if (action === 'approve') {
    return handleApprove(e);
  }
  if (action === 'saveclientemail') {
    const data = JSON.parse(e.postData.contents);
    return handleSaveClientEmail(data.company, data.email);
  }

  // デフォルト：フォームデータ受信
  return handleReceiveFormData(e);
}

// ===== 機能1: パスワード認証 =====

function handleAuth(e) {
  const password = e.parameter.password || '';
  const formType = e.parameter.formType || 'support';
  const props = getProps();

  const correctPassword = formType === 'self'
    ? props.getProperty('FORM_PASSWORD_SELF')
    : props.getProperty('FORM_PASSWORD');

  if (password === correctPassword) {
    return jsonResponse({ success: true });
  }
  return jsonResponse({ success: false, message: 'パスワードが違います' });
}

function handleRequestMagicLink(e) {
  const type = (e.parameter.type || 'admin').toLowerCase();
  const props = getProps();
  const adminEmail = props.getProperty('PARENT_COMPANY_EMAIL');
  const dilettoEmail = '2han2be4han@gmail.com';
  
  const targetEmail = (type === 'diletto') ? dilettoEmail : adminEmail;
  if (!targetEmail) return jsonResponse({ success: false, message: '送信先メールアドレスが設定されていません' });

  const token = Utilities.getUuid();
  const expires = new Date();
  expires.setMinutes(expires.getMinutes() + 5); // 5分間有効

  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME_TOKENS);
  sheet.appendRow([token, expires, false]);

  const subject = `【Office DAIKOU】管理画面ログインリンク (${type === 'diletto' ? 'diletto' : '大晃工業'})`;
  const plainBody = `お疲れ様です。\n管理画面へのログイン用トークンを発行しました。\n\nトークン: ${token}\n有効期限: 5分 (${expires.toLocaleString()})`;
  const htmlBody = getMagicLinkHtml_(type, token, expires);
  
  try {
    MailApp.sendEmail({
      to: targetEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    return jsonResponse({ success: true, message: `${type === 'diletto' ? 'diletto' : '大晃工業'}宛にメールを送信しました` });
  } catch (err) {
    return jsonResponse({ success: false, message: 'メール送信に失敗しました: ' + err.message });
  }
}

function handleVerifyToken(e) {
  const token = e.parameter.token || '';
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME_TOKENS);
  const rows = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === token && !rows[i][2]) {
      const expires = new Date(rows[i][1]);
      if (now < expires) {
        // 使用済みにする
        sheet.getRange(i + 1, 3).setValue(true);
        return jsonResponse({ success: true });
      }
    }
  }
  return jsonResponse({ success: false, message: 'トークンが無効または期限切れです' });
}

// ===== 取引先メール管理 =====

function handleGetClientEmail(e) {
  const company = e.parameter.company || '';
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME_CLIENT_SETTINGS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === company) return jsonResponse({ success: true, email: rows[i][1] });
  }
  return jsonResponse({ success: true, email: '' });
}

function handleSaveClientEmail(company, email) {
  if (!company) return jsonResponse({ success: false, message: '会社名が指定されていません' });
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME_CLIENT_SETTINGS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === company) {
      sheet.getRange(i + 1, 2).setValue(email);
      return jsonResponse({ success: true });
    }
  }
  sheet.appendRow([company, email]);
  return jsonResponse({ success: true });
}

// ===== 機能2: 従業員リスト取得 =====

function handleGetEmployees(e) {
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME_EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  const employees = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) employees.push(rows[i][0]);
  }
  return jsonResponse({ success: true, employees: employees });
}

// ===== 機能3: フォームデータ受信・スプレッドシート追記 =====

function handleReceiveFormData(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();

    // 自社受け持ち現場（invoice.html タブ1）
    if (data.formType === 'selfSite') {
      const sheet = getOrCreateSheet(ss, SHEET_NAME_DATA_OWN);
      const subtotal = (data.amount * data.rate) + data.highway;
      sheet.appendRow([
        new Date(), data.year, data.month, data.day, data.primeCompany || '', data.site,
        data.amount, data.rate, data.carCount, data.highway, subtotal, '未承認'
      ]);
      return jsonResponse({ success: true, subtotal: subtotal });
    }
    
    // 自社から応援（invoice.html タブ2）
    if (data.formType === 'supportSite') {
      const sheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
      const employees = (data.employees || []).join(', ');
      const workerStr = data.otherEmployee ? (employees ? employees + ', ' + data.otherEmployee : data.otherEmployee) : employees;
      const count = (data.employees || []).length + (data.otherEmployee ? 1 : 0);
      const workersCount = count > 0 ? count : 1;
      const subtotal = (workersCount * data.rate) + data.highway;
      
      sheet.appendRow([
        new Date(), data.year, data.month, data.day, '自社', '大晃工業', data.targetCompany,
        data.site, workersCount, workerStr, data.rate || 0, data.carCount, data.highway || 0, subtotal, '', '未承認'
      ]);
      return jsonResponse({ success: true, subtotal: data.highway });
    }

    // 他社から応援（index.html）
    if (data.formType === 'support') {
      const sheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
      const workerNames = (data.workerNames || []).join(', ');
      const highwayAndParking = (data.parking || 0) + (data.highway || 0);
      // index.htmlからの入力は単価がないので0、小計も0(あとで補完)
      sheet.appendRow([
        new Date(), data.year, data.month, data.day, '他社', data.myCompany, data.targetCompany,
        data.site, data.workers, workerNames, 0, data.carCount, highwayAndParking, 0, data.comment || '', '未承認'
      ]);
      return jsonResponse({ success: true, subtotal: highwayAndParking });
    }

    // 旧フォーマット（後方互換）
    return handleLegacyFormData(ss, data);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function handleLegacyFormData(ss, data) {
  const sheet = getOrCreateSheet(ss, SHEET_NAME_DATA);
  const timestamp = new Date();
  const subtotal = data.workers * data.rate + data.highway;

  sheet.appendRow([
    timestamp,        // A: タイムスタンプ
    data.year,        // B: 年
    data.month,       // C: 月
    data.day,         // D: 日
    data.company,     // E: 企業名
    data.site,        // F: 現場名
    data.workers,     // G: 人数
    data.rate,        // H: 駐車場代
    data.highway,     // I: 高速代
    subtotal,         // J: 小計
    '未承認'           // K: 承認ステータス
  ]);

  return jsonResponse({ success: true, subtotal: subtotal });
}

// ===== 機能4: 月初自動請求書URL生成＆メール送信 =====

function generateAndSendInvoiceUrls() {
  const ss = getSpreadsheet();
  const dataSheet = getOrCreateSheet(ss, SHEET_NAME_DATA);
  const invoiceSheet = getOrCreateSheet(ss, SHEET_NAME_INVOICE);
  const props = getProps();

  // 前月の年月を取得
  const now = new Date();
  const prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const targetYear = prevMonth.getFullYear();
  const targetMonth = prevMonth.getMonth() + 1;

  // 入力データ取得（ヘッダー除く）
  const dataRows = dataSheet.getDataRange().getValues();
  if (dataRows.length <= 1) return;

  // 企業ごとに集計
  const companyMap = {};
  for (let i = 1; i < dataRows.length; i++) {
    const row = dataRows[i];
    const year = row[1];
    const month = row[2];
    const company = row[4];

    if (year === targetYear && month === targetMonth) {
      if (!companyMap[company]) {
        companyMap[company] = [];
      }
      companyMap[company].push(row);
    }
  }

  // 既に送信済みの企業を確認
  const invoiceRows = invoiceSheet.getDataRange().getValues();
  const sentCompanies = new Set();
  for (let i = 1; i < invoiceRows.length; i++) {
    if (invoiceRows[i][0] === targetYear && invoiceRows[i][1] === targetMonth) {
      sentCompanies.add(invoiceRows[i][2]);
    }
  }

  // 請求書一覧ページのURL
  const listPageUrl = 'https://2han2be4han.github.io/office-daikou/invoice.html?year=' + targetYear + '&month=' + targetMonth;
  const emailTo = props.getProperty('PARENT_COMPANY_EMAIL');
  const companySummary = [];
  let grandTotal = 0;
  let hasNew = false;

  for (const company in companyMap) {
    // 重複送信防止
    if (sentCompanies.has(company)) continue;
    hasNew = true;

    const items = companyMap[company];
    let total = 0;
    items.forEach(row => { total += row[9]; }); // J列: 小計
    grandTotal += total;

    // 請求書シートに記録
    invoiceSheet.appendRow([
      targetYear,
      targetMonth,
      company,
      total,
      listPageUrl,
      new Date(),  // 送信日時
      '',          // 承認日時（未承認）
      '未承認'
    ]);

    companySummary.push(`  ・${company}：¥${total.toLocaleString()}`);
  }

  // メール送信（一覧URL1つだけ）
  if (emailTo && hasNew) {
    const subject = `【Office DAIKOU】${targetYear}年${targetMonth}月分 請求書のご確認`;
    const plainBody = `お疲れ様です。\n${targetYear}年${targetMonth}月分の請求書をお送りします。\n合計：¥${grandTotal.toLocaleString()}\n${listPageUrl}`;
    const htmlBody = getInvoiceNotificationHtml_(targetYear, targetMonth, companySummary, grandTotal, listPageUrl);

    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
    });
  }
}

// ===== 機能5: 請求書データ取得 =====

function handleGetInvoice(e) {
  try {
    const company = e.parameter.company || '';
    const year = parseInt(e.parameter.year);
    const month = parseInt(e.parameter.month);

    const ss = getSpreadsheet();
    const ownSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_OWN);
    const supportSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
    const invoiceSheet = getOrCreateSheet(ss, SHEET_NAME_INVOICE);

    const items = [];
    let payableTotal = 0;

    // 自社の単価辞書を作成 (キー: "日付_現場名", 値: 単価)
    const rateMap = {};
    const ownRows = ownSheet.getDataRange().getValues();
    const supportRows = supportSheet.getDataRange().getValues();
    for (let i = 1; i < ownRows.length; i++) {
      if (ownRows[i][1] === year && ownRows[i][2] === month) rateMap[`${ownRows[i][3]}_${ownRows[i][4]}`] = ownRows[i][6] || 0;
    }
    for (let i = 1; i < supportRows.length; i++) {
      if (supportRows[i][1] === year && supportRows[i][2] === month && supportRows[i][4] === '自社') {
        rateMap[`${supportRows[i][3]}_${supportRows[i][7]}`] = supportRows[i][10] || 0;
      }
    }

    if (company === '【自社現場】') {
      for (let i = 1; i < ownRows.length; i++) {
        const row = ownRows[i];
        if (row[1] === year && row[2] === month) {
          items.push({
            date: row[2] + '/' + row[3],
            site: (row[5] || '') + (row[4] ? ` (${row[4]})` : ''), // 現場名 (元受け)
            workers: row[6],
            rate: row[7],
            highway: row[9],
            subtotal: row[10],
          });
        }
      }
    } else {
      for (let i = 1; i < supportRows.length; i++) {
        const row = supportRows[i];
        if (row[1] === year && row[2] === month) {
          if (row[4] === '自社' && row[6] === company) {
            items.push({
              date: row[2] + '/' + row[3],
              site: row[7],
              workers: row[8],
              rate: row[10],
              highway: row[12],
              subtotal: row[13],
            });
          } else if (row[4] === '他社' && row[5] === company) {
            const matchedRate = rateMap[`${row[3]}_${row[7]}`] || 0;
            const workers = row[8] || 1;
            const hw = row[12] || 0;
            payableTotal += (workers * matchedRate) + hw;
          }
        }
      }
    }

    // 承認状態確認
    let approved = false;
    const invoiceRows = invoiceSheet.getDataRange().getValues();
    for (let i = 1; i < invoiceRows.length; i++) {
      if (invoiceRows[i][0] === year && invoiceRows[i][1] === month && invoiceRows[i][2] === company) {
        approved = invoiceRows[i][7] === '承認済み';
        break;
      }
    }

    return jsonResponse({
      success: true,
      company: company,
      year: year,
      month: month,
      items: items,
      payable: payableTotal,
      approved: approved,
    });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

// ===== 機能5b: 請求書一覧データ取得 =====

function handleGetInvoiceList(e) {
  try {
    const year = parseInt(e.parameter.year);
    const month = parseInt(e.parameter.month);

    const ss = getSpreadsheet();
    const ownSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_OWN);
    const supportSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
    const invoiceSheet = getOrCreateSheet(ss, SHEET_NAME_INVOICE);

    const companyMap = {};

    // 自社の単価辞書を作成
    const rateMap = {};
    const ownRows = ownSheet.getDataRange().getValues();
    const supportRows = supportSheet.getDataRange().getValues();
    for (let i = 1; i < ownRows.length; i++) {
      if (ownRows[i][1] === year && ownRows[i][2] === month) rateMap[`${ownRows[i][3]}_${ownRows[i][5]}`] = ownRows[i][7] || 0;
    }
    for (let i = 1; i < supportRows.length; i++) {
      if (supportRows[i][1] === year && supportRows[i][2] === month && supportRows[i][4] === '自社') {
        rateMap[`${supportRows[i][3]}_${supportRows[i][7]}`] = supportRows[i][10] || 0;
      }
    }

    // 1. 自社現場
    let ownTotal = 0;
    let ownCount = 0;
    for (let i = 1; i < ownRows.length; i++) {
      if (ownRows[i][1] === year && ownRows[i][2] === month) {
        ownTotal += ownRows[i][10] || 0; // 小計の列（インデックス10）
        ownCount++;
      }
    }
    if (ownCount > 0) {
      companyMap['【自社現場】'] = { total: ownTotal, count: ownCount, payable: 0 };
    }

    // 2. 応援現場
    for (let i = 1; i < supportRows.length; i++) {
      const row = supportRows[i];
      if (row[1] === year && row[2] === month) {
        if (row[4] === '自社') {
          const company = row[6];
          if (!companyMap[company]) companyMap[company] = { total: 0, count: 0, payable: 0 };
          companyMap[company].total += row[13] || 0;
          companyMap[company].count++;
        } else if (row[4] === '他社') {
          const company = row[5];
          if (!companyMap[company]) companyMap[company] = { total: 0, count: 0, payable: 0 };
          const matchedRate = rateMap[`${row[3]}_${row[7]}`] || 0;
          const workers = row[8] || 1;
          const hw = row[12] || 0;
          companyMap[company].payable += (workers * matchedRate) + hw;
        }
      }
    }

    // 承認状態を確認
    const invoiceRows = invoiceSheet.getDataRange().getValues();
    const approvedCompanies = new Set();
    for (let i = 1; i < invoiceRows.length; i++) {
      if (invoiceRows[i][0] === year && invoiceRows[i][1] === month && invoiceRows[i][7] === '承認済み') {
        approvedCompanies.add(invoiceRows[i][2]);
      }
    }

    const companies = [];
    for (const company in companyMap) {
      companies.push({
        company: company,
        total: companyMap[company].total,
        count: companyMap[company].count,
        payable: companyMap[company].payable || 0,
        approved: approvedCompanies.has(company),
      });
    }

    companies.sort((a, b) => a.company.localeCompare(b.company, 'ja'));

    return jsonResponse({ success: true, companies: companies });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

// ===== 機能6: 承認処理 & PDF生成・メール送信 =====

function handleApprove(e) {
  try {
    const company = e.parameter.company || '';
    const year = parseInt(e.parameter.year);
    const month = parseInt(e.parameter.month);

    const ss = getSpreadsheet();
    const invoiceSheet = getOrCreateSheet(ss, SHEET_NAME_INVOICE);
    const dataSheet = getOrCreateSheet(ss, SHEET_NAME_DATA);

    // 請求書シートで対象行を探して承認更新
    const invoiceRows = invoiceSheet.getDataRange().getValues();
    let invoiceRowIdx = -1;
    for (let i = 1; i < invoiceRows.length; i++) {
      if (invoiceRows[i][0] === year && invoiceRows[i][1] === month && invoiceRows[i][2] === company) {
        if (invoiceRows[i][7] === '承認済み') {
          return jsonResponse({ success: false, message: 'この請求書は既に承認済みです' });
        }
        invoiceRowIdx = i + 1; // シート行番号（1始まり）
        break;
      }
    }

    if (invoiceRowIdx === -1) {
      return jsonResponse({ success: false, message: '対象の請求書が見つかりません' });
    }

    // 承認ステータス更新
    invoiceSheet.getRange(invoiceRowIdx, 7).setValue(new Date()); // 承認日時
    invoiceSheet.getRange(invoiceRowIdx, 8).setValue('承認済み');

    // 入力データの承認ステータスも更新
    if (company === '【自社現場】') {
      const ownSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_OWN);
      const ownRows = ownSheet.getDataRange().getValues();
      for (let i = 1; i < ownRows.length; i++) {
        if (ownRows[i][1] === year && ownRows[i][2] === month) ownSheet.getRange(i + 1, 10).setValue('承認済み'); // J列
      }
    } else {
      const supportSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
      const supportRows = supportSheet.getDataRange().getValues();
      for (let i = 1; i < supportRows.length; i++) {
        if (supportRows[i][1] === year && supportRows[i][2] === month && supportRows[i][6] === company && supportRows[i][4] === '自社') {
          supportSheet.getRange(i + 1, 16).setValue('承認済み'); // P列
        }
      }
    }

    // PDF生成・メール送信
    generateAndSendPdf(company, year, month);

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function generateAndSendPdf(company, year, month) {
  const ss = getSpreadsheet();
  const ownSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_OWN);
  const supportSheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
  const props = getProps();

  const items = [];
  let total = 0;

  if (company === '【自社現場】') {
    const ownRows = ownSheet.getDataRange().getValues();
    for (let i = 1; i < ownRows.length; i++) {
      const row = ownRows[i];
      if (row[1] === year && row[2] === month) {
        items.push({ date: row[2] + '/' + row[3], site: row[4], workers: row[5], rate: row[6], highway: row[7], subtotal: row[8] });
        total += row[8] || 0;
      }
    }
  } else {
    const supportRows = supportSheet.getDataRange().getValues();
    for (let i = 1; i < supportRows.length; i++) {
      const row = supportRows[i];
      if (row[1] === year && row[2] === month && row[4] === '自社' && row[6] === company) {
        items.push({ date: row[2] + '/' + row[3], site: row[7], workers: row[8], rate: row[10], highway: row[12], subtotal: row[13] });
        total += row[13] || 0;
      }
    }
  }

  // PDF用HTML生成
  const html = `
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
        body { font-family: 'Noto Sans JP', sans-serif; padding: 40px; color: #333; background: #fff; }
        .header { border-bottom: 3px solid #111; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: flex-end; }
        .logo { font-size: 28px; font-weight: 900; letter-spacing: -0.02em; color: #111; }
        .logo span { color: #3b82f6; }
        .title-box { text-align: right; }
        .title { font-size: 32px; font-weight: 900; margin: 0; color: #111; }
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 40px; margin-bottom: 40px; }
        .bill-to { font-size: 18px; font-weight: 700; border-bottom: 1px solid #111; padding-bottom: 8px; }
        .company-info { font-size: 13px; line-height: 1.6; text-align: right; }
        table { width: 100%; border-collapse: collapse; margin-block: 30px; }
        th { background: #111; color: #fff; padding: 12px 10px; font-size: 12px; font-weight: 700; text-align: left; }
        td { padding: 12px 10px; border-bottom: 1px solid #eee; font-size: 13px; }
        .r { text-align: right; }
        .total-box { margin-left: auto; width: 300px; padding: 20px; background: #f8f7f5; border-radius: 8px; text-align: right; }
        .total-label { font-size: 14px; color: #666; font-weight: 700; }
        .total-val { font-size: 28px; font-weight: 900; color: #111; margin-top: 4px; }
        .footer-note { margin-top: 60px; font-size: 11px; color: #999; text-align: center; }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="logo">di<span>let</span>to</div>
        <div class="title-box">
          <p style="font-size: 12px; color: #666; margin:0;">${year}年${month}月分</p>
          <h1 class="title">請求書</h1>
        </div>
      </div>

      <div class="info-grid">
        <div class="bill-to">
          ${escapeHtml_(company)} 御中
        </div>
      </div>

      <table>
        <thead>
          <tr>
            <th>日付</th>
            <th>現場名</th>
            <th class="r">人数/数量</th>
            <th class="r">単価</th>
            <th class="r">高速等</th>
            <th class="r">小計</th>
          </tr>
        </thead>
        <tbody>
          ${items.map(item => `
            <tr>
              <td>${item.date}</td>
              <td style="font-weight:700;">${escapeHtml_(item.site)}</td>
              <td class="r">${item.workers}</td>
              <td class="r">¥${Number(item.rate).toLocaleString()}</td>
              <td class="r">¥${Number(item.highway).toLocaleString()}</td>
              <td class="r" style="font-weight:700;">¥${Number(item.subtotal).toLocaleString()}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>

      <div class="total-box">
        <p class="total-label">合計金額 (税込)</p>
        <p class="total-val">¥${total.toLocaleString()}</p>
      </div>

      <div class="footer-note">
        Office DAIKOU by 大晃工業合同会社 / diletto
      </div>
    </body>
    </html>
  `;

  // HTMLをBlobとしてPDFに変換
  const blob = Utilities.newBlob(html, 'text/html', `請求書_${company}_${year}年${month}月.html`)
    .getAs('application/pdf')
    .setName(`請求書_${company}_${year}年${month}月.pdf`);

  // メール送信
  let emailTo = props.getProperty('PARENT_COMPANY_EMAIL');
  
  // 取引先設定から個別メールアドレスを取得
  const clientSheet = getOrCreateSheet(ss, SHEET_NAME_CLIENT_SETTINGS);
  const clientRows = clientSheet.getDataRange().getValues();
  for (let i = 1; i < clientRows.length; i++) {
    if (clientRows[i][0] === company && clientRows[i][1]) {
      emailTo = clientRows[i][1];
      break;
    }
  }

  if (emailTo) {
    const plainBody = `${company}の${year}年${month}月分の請求書が承認されました。\nPDFを添付しておりますのでご確認ください。`;
    const htmlBody = getInvoiceApprovedHtml_(company, year, month, total, items);

    MailApp.sendEmail({
      to: emailTo,
      subject: `【Office DAIKOU】${company} ${year}年${month}月分 請求書（承認済み）`,
      body: plainBody,
      htmlBody: htmlBody,
      attachments: [blob],
    });
  }
}

// HTMLエスケープ
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ===== メールHTMLテンプレート =====

function getInvoiceNotificationHtml_(year, month, companySummary, grandTotal, listPageUrl) {
  const companyRows = companySummary.map(s => {
    return `<tr><td style="padding:12px 14px; font-size:15px; color:#333; border-bottom:1px solid #f0f0f0; font-weight:500;">${escapeHtml_(s.replace(/^\s+・/, ''))}</td></tr>`;
  }).join('');

  return `
  <div style="font-family:'Helvetica Neue', Arial, 'Noto Sans JP', sans-serif; background-color:#f8f7f5; padding:50px 20px; color:#111;">
    <div style="max-width:600px; margin:0 auto; background-color:#fff; border-radius:16px; overflow:hidden; box-shadow:0 20px 40px rgba(0,0,0,0.06);">
      <div style="background-color:#111; padding:40px; text-align:center;">
        <span style="color:#fff; font-size:32px; font-weight:900; letter-spacing:-0.02em;">di<span style="color:#3b82f6;">let</span>to</span>
      </div>
      <div style="padding:50px 40px;">
        <h1 style="font-size:24px; font-weight:900; color:#111; margin-bottom:8px;">請求書のご確認</h1>
        <p style="font-size:15px; color:#888; margin-bottom:32px;">${year}年${month}月分</p>
        
        <p style="font-size:16px; line-height:1.7; color:#333; margin-bottom:32px;">
          お疲れ様です。<br>
          以下の通り請求明細が届いています。内容を確認し、承認をお願いいたします。
        </p>

        <div style="background:#f8f7f5; border-radius:12px; padding:24px; margin-bottom:40px;">
          <p style="font-size:12px; color:#aaa; font-weight:800; text-transform:uppercase; letter-spacing:0.1em; margin:0 0 16px;">Summary</p>
          <table style="width:100%; border-collapse:collapse; background:#fff; border-radius:8px; overflow:hidden;">
            ${companyRows}
          </table>
          <div style="margin-top:24px; text-align:right;">
            <p style="font-size:12px; color:#aaa; margin:0;">Grand Total</p>
            <p style="font-size:28px; font-weight:900; color:#111; margin:0;">¥${grandTotal.toLocaleString()}</p>
          </div>
        </div>

        <div style="text-align:center; margin-bottom:48px;">
          <a href="${listPageUrl}" style="display:inline-block; background-color:#111; color:#fff; padding:20px 40px; border-radius:10px; font-size:16px; font-weight:800; text-decoration:none; box-shadow:0 10px 20px rgba(0,0,0,0.15);">
            一覧ページを開く
          </a>
        </div>

        <p style="font-size:14px; color:#bbb; line-height:1.6; text-align:center;">
          Office DAIKOU by 大晃工業合同会社 / diletto
        </p>
      </div>
    </div>
  </div>`;
}

function getInvoiceApprovedHtml_(company, year, month, total, items) {
  const tableRows = items.map(item => {
    return `<tr>
      <td style="padding:12px 10px; border-bottom:1px solid #f0f0f0; font-size:13px; color:#333;">${escapeHtml_(item.date)}</td>
      <td style="padding:12px 10px; border-bottom:1px solid #f0f0f0; font-size:13px; color:#333; font-weight:700;">${escapeHtml_(item.site)}</td>
      <td style="padding:12px 10px; border-bottom:1px solid #f0f0f0; font-size:13px; color:#333; text-align:right;">${item.workers}</td>
      <td style="padding:12px 10px; border-bottom:1px solid #f0f0f0; font-size:13px; color:#333; text-align:right;">&yen;${Number(item.rate).toLocaleString()}</td>
      <td style="padding:12px 10px; border-bottom:1px solid #f0f0f0; font-size:13px; color:#333; text-align:right; font-weight:700;">&yen;${Number(item.subtotal).toLocaleString()}</td>
    </tr>`;
  }).join('');

  return `
  <div style="font-family:'Helvetica Neue', Arial, 'Noto Sans JP', sans-serif; background-color:#f8f7f5; padding:50px 20px; color:#111;">
    <div style="max-width:640px; margin:0 auto; background-color:#fff; border-radius:16px; overflow:hidden; box-shadow:0 20px 40px rgba(0,0,0,0.06);">
      <div style="background-color:#111; padding:40px; text-align:center;">
        <span style="color:#fff; font-size:32px; font-weight:900; letter-spacing:-0.02em;">di<span style="color:#3b82f6;">let</span>to</span>
      </div>
      <div style="padding:50px 40px;">
        <h2 style="font-size:22px; font-weight:900; color:#111; margin-bottom:8px;">${escapeHtml_(company)} 様</h2>
        <p style="font-size:15px; color:#888; margin-bottom:32px;">${year}年${month}月分 請求明細</p>

        <p style="font-size:16px; line-height:1.7; color:#333; margin-bottom:32px;">
          いつも大変お世話になっております。<br>
          本日、請求書の承認が完了いたしました。
        </p>

        <div style="margin-bottom:40px; border:1px solid #eee; border-radius:12px; overflow:hidden;">
          <table style="width:100%; border-collapse:collapse;">
            <thead>
              <tr>
                <th style="background:#f8f7f5; color:#111; padding:12px 10px; font-size:11px; font-weight:900; text-align:left; border-bottom:1px solid #eee;">日付</th>
                <th style="background:#f8f7f5; color:#111; padding:12px 10px; font-size:11px; font-weight:900; text-align:left; border-bottom:1px solid #eee;">現場名</th>
                <th style="background:#f8f7f5; color:#111; padding:12px 10px; font-size:11px; font-weight:900; text-align:right; border-bottom:1px solid #eee;">人数/数量</th>
                <th style="background:#f8f7f5; color:#111; padding:12px 10px; font-size:11px; font-weight:900; text-align:right; border-bottom:1px solid #eee;">単価値</th>
                <th style="background:#f8f7f5; color:#111; padding:12px 10px; font-size:11px; font-weight:900; text-align:right; border-bottom:1px solid #eee;">小計</th>
              </tr>
            </thead>
            <tbody>
              ${tableRows}
            </tbody>
          </table>
          <div style="padding:24px; background:#fff; text-align:right; border-top:2px solid #111;">
            <p style="font-size:12px; color:#aaa; margin:0 0 4px; font-weight:700;">Total Amount</p>
            <p style="font-size:32px; font-weight:900; color:#111; margin:0;">&yen;${total.toLocaleString()}</p>
          </div>
        </div>

        <p style="font-size:15px; line-height:1.7; color:#333; margin-bottom:48px; background:#f8f7f5; padding:20px; border-radius:8px;">
          正式な請求書（PDF）を本メールに添付いたしました。<br>
          お手数ですが、ご確認いただけますようお願い申し上げます。
        </p>

        <p style="font-size:13px; color:#bbb; text-align:center;">
          Office DAIKOU by 大晃工業合同会社 / diletto
        </p>
      </div>
    </div>
  </div>`;
}

function getMagicLinkHtml_(type, token, expires) {
  const userType = type === 'diletto' ? 'diletto' : '大晃工業';
  const loginUrl = `https://2han2be4han.github.io/office-daikou/admin-login.html?token=${token}`;
  
  return `
  <div style="font-family:'Helvetica Neue', Arial, 'Noto Sans JP', sans-serif; background-color:#f8f7f5; padding:50px 20px; color:#111;">
    <div style="max-width:500px; margin:0 auto; background-color:#fff; border-radius:16px; overflow:hidden; box-shadow:0 20px 40px rgba(0,0,0,0.06);">
      <div style="background-color:#111; padding:40px; text-align:center;">
        <span style="color:#fff; font-size:32px; font-weight:900; letter-spacing:-0.02em;">di<span style="color:#3b82f6;">let</span>to</span>
      </div>
      <div style="padding:50px 40px;">
        <h2 style="font-size:22px; font-weight:900; color:#111; margin-bottom:8px;">認証コード</h2>
        <p style="font-size:15px; color:#888; margin-bottom:32px;">管理画面ログイン (${userType})</p>

        <p style="font-size:16px; line-height:1.7; color:#333; margin-bottom:32px;">
          お疲れ様です。<br>
          以下のボタンをクリックしてログインを完了してください。
        </p>

        <div style="text-align:center; margin-bottom:32px;">
          <a href="${loginUrl}" style="display:inline-block; background-color:#111; color:#fff; padding:18px 32px; border-radius:10px; font-size:16px; font-weight:800; text-decoration:none; box-shadow:0 10px 20px rgba(0,0,0,0.15);">
            管理画面にログインする
          </a>
        </div>

        <div style="background:#f8f7f5; border-radius:12px; padding:24px; text-align:center; margin-bottom:32px;">
          <p style="font-size:12px; color:#aaa; font-weight:800; text-transform:uppercase; letter-spacing:0.1em; margin:0 0 12px;">Token (手動入力用)</p>
          <p style="font-size:13px; font-family:monospace; font-weight:700; color:#111; margin:0; word-break:break-all;">${token}</p>
        </div>

        <p style="font-size:14px; color:#999; text-align:center; margin-bottom:48px;">
          有効期限: ${expires.toLocaleString()}<br>
          <span style="font-size:12px;">※発行から5分間のみ有効です</span>
        </p>

        <p style="font-size:13px; color:#bbb; text-align:center;">
          Office DAIKOU by 大晃工業合同会社 / diletto
        </p>
      </div>
    </div>
  </div>`;
}

// ===== トリガー設定用関数 =====
// GASエディタから手動で1回だけ実行してください

function setupMonthlyTrigger() {
  // 既存トリガーの削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'generateAndSendInvoiceUrls') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎月1日 0:00〜1:00 にトリガー設定
  ScriptApp.newTrigger('generateAndSendInvoiceUrls')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
}

    .atHour(0)
    .create();
}
