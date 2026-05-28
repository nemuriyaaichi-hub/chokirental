const NOTIFICATION_EMAIL = "nemuriya.aichi@gmail.com";
const SHEET_NAME = "お問い合わせ";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(spreadsheet);
    const data = e.parameter || {};
    const headers = [
      "受信日時",
      "ご予約・お見積もり・ご相談",
      "メールアドレス",
      "レンタルプラン",
      "お届け・相談希望日",
      "お届け・相談希望時間",
      "返却希望日",
      "レンタル希望組数",
      "お名前",
      "ご住所",
      "電話番号",
      "その他質問など"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    const row = [
      data.submitted_at || new Date(),
      data.request_type || "",
      data.email || "",
      data.plan || "",
      data.delivery_date || "",
      data.delivery_time || "",
      data.return_date || "",
      data.quantity || "",
      data.name || "",
      data.address || "",
      data.phone || "",
      data.message || ""
    ];
    sheet.appendRow(row);

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      replyTo: data.email || NOTIFICATION_EMAIL,
      subject: "ホームページからお問い合わせがありました",
      body: headers.map(function (header, index) {
        return header + ": " + row[index];
      }).join("\n")
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService
    .createTextOutput("contact form endpoint")
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet_(spreadsheet) {
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}
