/*

* Gmail Bounce Report
* Written by Shivanand Nalgire

* email: 30004693@iitb.ac.in

*/

const toast_ = e => SpreadsheetApp.getActiveSpreadsheet().toast(e);

const parseMessage_ = messageId => {
  const message = GmailApp.getMessageById(messageId);
  const body = message.getPlainBody();
  const [, failAction] = body.match(/^Action:\s*(.+)/m) || [];
  if (failAction === "failed") {
    const emailAddress = message.getHeader("X-Failed-Recipients");
    const [, errorStatus] = body.match(/^Status:\s*([.\d]+)/m) || [];
    const [, , bounceReason] =
      body.match(/^Diagnostic-Code:\s*(.+)\s*;\s*(.+)/m) || [];
    if (errorStatus && bounceReason) {
     return [
      message.getDate(),
      emailAddress,
      errorStatus,
      bounceReason.replace(/\s*(Please|Learn|See).+$/, ""),
      `=HYPERLINK("${message.getThread().getPermalink()}";"View")`
     ];
    }
  }
  return false;
};

const writeBouncedEmails_ = (data = []) => {
  if (data.length > 0) {
    toast_("Writing data to Google Sheet..");
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet
      .getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn())
      .clearContent();
    sheet.getRange(3, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
    toast_("Bounce Report is ready!");
    showCredit_();
  }
};

const findBouncedEmails_ = () => {
  try {
    const rows = [];
    const { messages = [] } = Gmail.Users.Messages.list("me", {
      q: "from:mailer-daemon",
      maxResults: 200
    });
    if (messages.length === 0) {
      toast_("No bounced emails found in your mailbox!");
      return;
    }
    toast_("Working..");
    for (let m = 0; m < messages.length; m += 1) {
      const bounceData = parseMessage_(messages[m].id);
      if (bounceData) {
        rows.push(bounceData);
      }
      if (rows.length % 10 === 0) {
        toast_(`${rows.length} bounced emails found so far..`);
      }
    }
    writeBouncedEmails_(rows);
  } catch (error) {
    toast_(error.toString());
  }
};



const onOpen = e => {
  SpreadsheetApp.getUi()
    .createMenu("🕵🏻‍♂️ Bounced Emails")
    .addItem("Run Report", "findBouncedEmails_")
    .addSeparator()
    .addItem("Credits", "showCredit_")
    .addToUi();
};


const showCredit_ = () => {
  const template = HtmlService.createHtmlOutputFromFile("help");
  const html = template
    .setTitle("Bounce Report for Gmail")
    .setWidth(460)
    .setHeight(225);
  SpreadsheetApp.getActiveSpreadsheet().show(html);
}
