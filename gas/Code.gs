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
const SHEET_NAME_DATA = '入力データ';
const SHEET_NAME_DATA_SUPPORT = '入力データ（応援）';
const SHEET_NAME_DATA_SELF = '入力データ（自社）';
const SHEET_NAME_INVOICE = '請求書';
const SHEET_NAME_EMPLOYEES = '従業員';

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
    if (name === SHEET_NAME_DATA) {
      sheet.appendRow([
        'タイムスタンプ', '年', '月', '日', '企業名',
        '現場名', '人数', '駐車場代', '高速代', '小計', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_DATA_SUPPORT) {
      sheet.appendRow([
        'タイムスタンプ', '年', '月', '日', 'あなたの会社名',
        '応援先会社名', '現場名', '人数', '作業員名',
        '車台数', '駐車場代', '高速代', '小計', 'コメント', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_DATA_SELF) {
      sheet.appendRow([
        'タイムスタンプ', '年', '月', '日', '応援先会社名',
        '現場名', '従業員名', '車台数', '駐車場代', '高速代', '小計', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_INVOICE) {
      sheet.appendRow([
        '年', '月', '企業名', '合計金額', '請求書URL', '送信日時', '承認日時', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_EMPLOYEES) {
      sheet.appendRow(['名前']);
      // サンプルデータ
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
    case 'invoice':
      return handleGetInvoice(e);
    case 'invoicelist':
      return handleGetInvoiceList(e);
    case 'employees':
      return handleGetEmployees(e);
    default:
      return jsonResponse({ success: false, message: '不明なアクションです' });
  }
}

function doPost(e) {
  const action = (e.parameter && e.parameter.action || '').toLowerCase();

  if (action === 'approve') {
    return handleApprove(e);
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

    if (data.formType === 'support') {
      return handleSupportFormData(ss, data);
    } else if (data.formType === 'self') {
      return handleSelfFormData(ss, data);
    }

    // 旧フォーマット（後方互換）
    return handleLegacyFormData(ss, data);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function handleSupportFormData(ss, data) {
  const sheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SUPPORT);
  const timestamp = new Date();
  const subtotal = (data.parking || 0) + (data.highway || 0);
  const workerNames = (data.workerNames || []).join(', ');

  sheet.appendRow([
    timestamp,           // A: タイムスタンプ
    data.year,           // B: 年
    data.month,          // C: 月
    data.day,            // D: 日
    data.myCompany,      // E: あなたの会社名
    data.targetCompany,  // F: 応援先会社名
    data.site,           // G: 現場名
    data.workers,        // H: 人数
    workerNames,         // I: 作業員名（カンマ区切り）
    data.carCount,       // J: 車台数
    data.parking || 0,   // K: 駐車場代
    data.highway || 0,   // L: 高速代
    subtotal,            // M: 小計
    data.comment || '',  // N: コメント
    '未承認'              // O: 承認ステータス
  ]);

  return jsonResponse({ success: true, subtotal: subtotal });
}

function handleSelfFormData(ss, data) {
  const sheet = getOrCreateSheet(ss, SHEET_NAME_DATA_SELF);
  const timestamp = new Date();
  const subtotal = (data.parking || 0) + (data.highway || 0);
  const employees = (data.employees || []).join(', ');
  if (data.otherEmployee) {
    const combined = employees ? employees + ', ' + data.otherEmployee : data.otherEmployee;
  }

  sheet.appendRow([
    timestamp,           // A: タイムスタンプ
    data.year,           // B: 年
    data.month,          // C: 月
    data.day,            // D: 日
    data.targetCompany,  // E: 応援先会社名
    data.site,           // F: 現場名
    data.otherEmployee ? (employees ? employees + ', ' + data.otherEmployee : data.otherEmployee) : employees, // G: 従業員名
    data.carCount,       // H: 車台数
    data.parking || 0,   // I: 駐車場代
    data.highway || 0,   // J: 高速代
    subtotal,            // K: 小計
    '未承認'              // L: 承認ステータス
  ]);

  return jsonResponse({ success: true, subtotal: subtotal });
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
    const dataSheet = getOrCreateSheet(ss, SHEET_NAME_DATA);
    const invoiceSheet = getOrCreateSheet(ss, SHEET_NAME_INVOICE);

    // 入力データからフィルタ
    const dataRows = dataSheet.getDataRange().getValues();
    const items = [];

    for (let i = 1; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (row[1] === year && row[2] === month && row[4] === company) {
        items.push({
          month: row[2],
          day: row[3],
          site: row[5],
          workers: row[6],
          rate: row[7],
          highway: row[8],
        });
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
    const dataSheet = getOrCreateSheet(ss, SHEET_NAME_DATA);
    const invoiceSheet = getOrCreateSheet(ss, SHEET_NAME_INVOICE);

    // 入力データから企業ごとに集計
    const dataRows = dataSheet.getDataRange().getValues();
    const companyMap = {};

    for (let i = 1; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (row[1] === year && row[2] === month) {
        const company = row[4];
        if (!companyMap[company]) {
          companyMap[company] = { total: 0, count: 0 };
        }
        companyMap[company].total += row[9]; // 小計
        companyMap[company].count += 1;
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
        approved: approvedCompanies.has(company),
      });
    }

    // 企業名でソート
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
    const dataRows = dataSheet.getDataRange().getValues();
    for (let i = 1; i < dataRows.length; i++) {
      if (dataRows[i][1] === year && dataRows[i][2] === month && dataRows[i][4] === company) {
        dataSheet.getRange(i + 1, 11).setValue('承認済み'); // K列
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
  const dataSheet = getOrCreateSheet(ss, SHEET_NAME_DATA);
  const props = getProps();

  // 対象データ取得
  const dataRows = dataSheet.getDataRange().getValues();
  const items = [];
  let total = 0;

  for (let i = 1; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (row[1] === year && row[2] === month && row[4] === company) {
      const subtotal = row[9];
      total += subtotal;
      items.push({
        date: `${row[2]}/${row[3]}`,
        site: row[5],
        workers: row[6],
        rate: row[7],
        highway: row[8],
        subtotal: subtotal,
      });
    }
  }

  // PDF用HTML生成
  const html = `
    <html>
    <head>
      <style>
        body { font-family: 'Noto Sans JP', sans-serif; padding: 40px; color: #1e293b; }
        h1 { font-size: 24px; text-align: center; margin-bottom: 8px; color: #1a2744; }
        .meta { text-align: center; color: #64748b; margin-bottom: 32px; font-size: 14px; }
        .company { font-size: 18px; font-weight: bold; margin-bottom: 24px; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { background: #1a2744; color: white; padding: 8px 10px; text-align: left; }
        th:last-child { text-align: right; }
        td { padding: 8px 10px; border-bottom: 1px solid #e2e8f0; }
        .right { text-align: right; }
        .total { text-align: right; font-size: 20px; font-weight: bold; margin-top: 16px; padding-top: 16px; border-top: 2px solid #1a2744; }
        .footer { margin-top: 40px; text-align: center; font-size: 12px; color: #94a3b8; }
      </style>
    </head>
    <body>
      <h1>請求書</h1>
      <p class="meta">${year}年${month}月分</p>
      <p class="company">${escapeHtml_(company)} 御中</p>
      <table>
        <thead>
          <tr><th>日付</th><th>現場名</th><th>人数</th><th>駐車場代</th><th>高速代</th><th style="text-align:right">小計</th></tr>
        </thead>
        <tbody>
          ${items.map(item => `
            <tr>
              <td>${item.date}</td>
              <td>${escapeHtml_(item.site)}</td>
              <td class="right">${item.workers}</td>
              <td class="right">¥${Number(item.rate).toLocaleString()}</td>
              <td class="right">¥${Number(item.highway).toLocaleString()}</td>
              <td class="right">¥${Number(item.subtotal).toLocaleString()}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <p class="total">合計金額: ¥${total.toLocaleString()}</p>
      <p class="footer">Office DAIKOU（大晃工業合同会社）</p>
    </body>
    </html>
  `;

  // HTMLをBlobとしてPDFに変換
  const blob = Utilities.newBlob(html, 'text/html', `請求書_${company}_${year}年${month}月.html`)
    .getAs('application/pdf')
    .setName(`請求書_${company}_${year}年${month}月.pdf`);

  // メール送信
  const emailTo = props.getProperty('PARENT_COMPANY_EMAIL');
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
    return `<tr><td style="padding:8px 12px; font-size:14px; color:#444; border-bottom:1px solid #eee;">${escapeHtml_(s.replace(/^\s+・/, ''))}</td></tr>`;
  }).join('');

  return `
  <div style="font-family:'Helvetica Neue', Arial, 'Noto Sans JP', sans-serif; background-color:#f5f0ea; padding:40px 20px; color:#111;">
    <div style="max-width:600px; margin:0 auto; background-color:#fff; border-radius:12px; overflow:hidden; box-shadow:0 10px 30px rgba(0,0,0,0.05); border:1px solid #e0dbd4;">
      <div style="background-color:#222; padding:30px; text-align:center;">
        <span style="color:#fff; font-size:24px; font-weight:800; letter-spacing:0.1em;"><em style="font-style:normal; color:#e07830;">O</em>ffice DAIKOU</span>
      </div>
      <div style="padding:40px 30px;">
        <h1 style="font-size:20px; font-weight:700; color:#111; margin-bottom:24px;">請求書のご確認</h1>
        <p style="font-size:16px; line-height:1.7; color:#444; margin-bottom:24px;">
          お疲れ様です。<br>
          <strong>${year}年${month}月分</strong>の請求書をお送りします。
        </p>
        <div style="background:#f5f0ea; border-radius:8px; padding:20px; margin-bottom:24px;">
          <p style="font-size:12px; color:#888; font-weight:600; letter-spacing:1px; margin:0 0 12px;">対象企業（${companySummary.length}社）</p>
          <table style="width:100%; border-collapse:collapse;">
            ${companyRows}
          </table>
          <div style="margin-top:16px; padding-top:12px; border-top:2px solid #e0dbd4; text-align:right;">
            <span style="font-size:12px; color:#888; margin-right:8px;">合計</span>
            <span style="font-size:22px; font-weight:900; color:#222;">¥${grandTotal.toLocaleString()}</span>
          </div>
        </div>
        <p style="font-size:16px; line-height:1.7; color:#444; margin-bottom:32px;">
          下記ボタンから各社の請求内容を確認し、承認をお願いいたします。
        </p>
        <div style="text-align:center; margin-bottom:40px;">
          <a href="${listPageUrl}" style="display:inline-block; background-color:#e07830; color:#fff; padding:18px 36px; border-radius:8px; font-size:16px; font-weight:700; text-decoration:none; box-shadow:0 8px 20px rgba(224,120,48,0.25);">
            請求書一覧を確認する
          </a>
        </div>
        <hr style="border:none; border-top:1px solid #e0dbd4; margin-bottom:32px;">
        <p style="font-size:14px; color:#888; line-height:1.6;">
          ※このメールは送信専用です。<br>
          ご不明点がございましたら担当者までお問い合わせください。
        </p>
      </div>
      <div style="background-color:#fafaf8; padding:30px; text-align:center; border-top:1px solid #e0dbd4;">
        <p style="font-size:12px; color:#aaa; margin:0;">
          &copy; 2026 diletto by AI Skill Exchange. All rights reserved.<br>
          Office DAIKOU（大晃工業合同会社）
        </p>
      </div>
    </div>
  </div>`;
}

function getInvoiceApprovedHtml_(company, year, month, total, items) {
  const tableRows = items.map(item => {
    return `<tr>
      <td style="padding:8px 10px; border-bottom:1px solid #eee; font-size:13px; color:#444;">${escapeHtml_(item.date)}</td>
      <td style="padding:8px 10px; border-bottom:1px solid #eee; font-size:13px; color:#444;">${escapeHtml_(item.site)}</td>
      <td style="padding:8px 10px; border-bottom:1px solid #eee; font-size:13px; color:#444; text-align:right;">${item.workers}</td>
      <td style="padding:8px 10px; border-bottom:1px solid #eee; font-size:13px; color:#444; text-align:right;">&yen;${Number(item.rate).toLocaleString()}</td>
      <td style="padding:8px 10px; border-bottom:1px solid #eee; font-size:13px; color:#444; text-align:right;">&yen;${Number(item.highway).toLocaleString()}</td>
      <td style="padding:8px 10px; border-bottom:1px solid #eee; font-size:13px; color:#444; text-align:right; font-weight:600;">&yen;${Number(item.subtotal).toLocaleString()}</td>
    </tr>`;
  }).join('');

  return `
  <div style="font-family:'Helvetica Neue', Arial, 'Noto Sans JP', sans-serif; background-color:#f5f0ea; padding:40px 20px; color:#111;">
    <div style="max-width:600px; margin:0 auto; background-color:#fff; border-radius:12px; overflow:hidden; box-shadow:0 10px 30px rgba(0,0,0,0.05); border:1px solid #e0dbd4;">
      <div style="background-color:#222; padding:30px; text-align:center;">
        <span style="color:#fff; font-size:24px; font-weight:800; letter-spacing:0.1em;"><em style="font-style:normal; color:#e07830;">O</em>ffice DAIKOU</span>
      </div>
      <div style="padding:40px 30px;">
        <div style="text-align:center; margin-bottom:24px;">
          <div style="width:56px; height:56px; border-radius:50%; background:#e8f5ee; border:2px solid #b0d4be; display:inline-flex; align-items:center; justify-content:center;">
            <span style="font-size:28px;">&#10003;</span>
          </div>
        </div>
        <h1 style="font-size:20px; font-weight:700; color:#111; margin-bottom:8px; text-align:center;">承認が完了しました</h1>
        <p style="font-size:14px; color:#888; text-align:center; margin-bottom:24px;">${year}年${month}月分</p>

        <div style="background:#f5f0ea; border-radius:8px; padding:20px; margin-bottom:24px;">
          <p style="font-size:16px; font-weight:700; color:#222; margin:0 0 4px;">${escapeHtml_(company)} 御中</p>
          <p style="font-size:14px; color:#888; margin:0;">${items.length}件の明細</p>
        </div>

        <table style="width:100%; border-collapse:collapse; margin-bottom:16px;">
          <thead>
            <tr>
              <th style="background:#222; color:#fff; padding:8px 10px; font-size:11px; font-weight:500; text-align:left;">日付</th>
              <th style="background:#222; color:#fff; padding:8px 10px; font-size:11px; font-weight:500; text-align:left;">現場名</th>
              <th style="background:#222; color:#fff; padding:8px 10px; font-size:11px; font-weight:500; text-align:right;">人数</th>
              <th style="background:#222; color:#fff; padding:8px 10px; font-size:11px; font-weight:500; text-align:right;">駐車場代</th>
              <th style="background:#222; color:#fff; padding:8px 10px; font-size:11px; font-weight:500; text-align:right;">高速代</th>
              <th style="background:#222; color:#fff; padding:8px 10px; font-size:11px; font-weight:500; text-align:right;">小計</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>

        <div style="text-align:right; padding-top:12px; border-top:2px solid #222;">
          <span style="font-size:12px; color:#888; margin-right:8px;">合計金額</span>
          <span style="font-size:24px; font-weight:900; color:#222;">&yen;${total.toLocaleString()}</span>
        </div>

        <p style="font-size:16px; line-height:1.7; color:#444; margin:24px 0 32px;">
          PDFファイルをこのメールに添付しております。<br>
          ご確認をお願いいたします。
        </p>

        <hr style="border:none; border-top:1px solid #e0dbd4; margin-bottom:32px;">
        <p style="font-size:14px; color:#888; line-height:1.6;">
          ※このメールは送信専用です。
        </p>
      </div>
      <div style="background-color:#fafaf8; padding:30px; text-align:center; border-top:1px solid #e0dbd4;">
        <p style="font-size:12px; color:#aaa; margin:0;">
          &copy; 2026 diletto by AI Skill Exchange. All rights reserved.<br>
          Office DAIKOU（大晃工業合同会社）
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
