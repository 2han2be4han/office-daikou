/**
 * Office DAIKOU - 人工出し請求書自動生成ツール
 * Google Apps Script 本体
 *
 * スクリプトプロパティに以下を設定してください：
 *   SHEET_ID            : GoogleスプレッドシートのID
 *   FORM_PASSWORD       : フォーム認証用パスワード
 *   PARENT_COMPANY_EMAIL: 親会社担当者のメールアドレス
 */

// ===== 定数 =====
const SHEET_NAME_DATA = '入力データ';
const SHEET_NAME_INVOICE = '請求書';

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
        '現場名', '人数', '単価', '高速代', '小計', '承認ステータス'
      ]);
    }
    if (name === SHEET_NAME_INVOICE) {
      sheet.appendRow([
        '年', '月', '企業名', '合計金額', '請求書URL', '送信日時', '承認日時', '承認ステータス'
      ]);
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
  const correctPassword = getProps().getProperty('FORM_PASSWORD');

  if (password === correctPassword) {
    return jsonResponse({ success: true });
  }
  return jsonResponse({ success: false, message: 'パスワードが違います' });
}

// ===== 機能3: フォームデータ受信・スプレッドシート追記 =====

function handleReceiveFormData(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();
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
      data.rate,        // H: 単価
      data.highway,     // I: 高速代
      subtotal,         // J: 小計
      '未承認'           // K: 承認ステータス
    ]);

    return jsonResponse({ success: true, subtotal: subtotal });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
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

  // WebアプリのURL
  const webAppUrl = ScriptApp.getService().getUrl();
  const emailTo = props.getProperty('PARENT_COMPANY_EMAIL');
  const emailBody = [];

  for (const company in companyMap) {
    // 重複送信防止
    if (sentCompanies.has(company)) continue;

    const items = companyMap[company];
    let total = 0;
    items.forEach(row => { total += row[9]; }); // J列: 小計

    // 請求書URL生成
    const invoiceUrl = webAppUrl +
      '?action=invoice' +
      '&company=' + encodeURIComponent(company) +
      '&year=' + targetYear +
      '&month=' + targetMonth;

    // 請求書シートに記録
    invoiceSheet.appendRow([
      targetYear,
      targetMonth,
      company,
      total,
      invoiceUrl,
      new Date(),  // 送信日時
      '',          // 承認日時（未承認）
      '未承認'
    ]);

    emailBody.push(`■ ${company}\n  合計金額: ¥${total.toLocaleString()}\n  請求書URL: ${invoiceUrl}\n`);
  }

  // メール送信
  if (emailTo && emailBody.length > 0) {
    const subject = `【Office DAIKOU】${targetYear}年${targetMonth}月分 請求書のご確認`;
    const body = `お疲れ様です。\n\n${targetYear}年${targetMonth}月分の請求書URLをお送りします。\n\n` +
      emailBody.join('\n') +
      `\n各URLをクリックして内容をご確認の上、承認をお願いいたします。\n\n` +
      `---\nOffice DAIKOU（大晃工業合同会社）`;

    MailApp.sendEmail(emailTo, subject, body);
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
          <tr><th>日付</th><th>現場名</th><th>人数</th><th>単価</th><th>高速代</th><th style="text-align:right">小計</th></tr>
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
    MailApp.sendEmail({
      to: emailTo,
      subject: `【Office DAIKOU】${company} ${year}年${month}月分 請求書（承認済み）`,
      body: `${company}の${year}年${month}月分の請求書が承認されました。\nPDFを添付しておりますのでご確認ください。\n\n---\nOffice DAIKOU（大晃工業合同会社）`,
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
