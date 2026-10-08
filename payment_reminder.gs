/**
 * LINE Payment Reminder Bot
 * ---------------------------------------------------------
 * Setup properties:
 *   LINE_CHANNEL_ACCESS_TOKEN  
 *   LINE_USER_ID               -> LINE userId
 *   SHEET_NAME                 -> tab in spreadsheet
 *   SHEET_ID                   -> Spreadsheet ID
 *
 * Functions:
 *   sendDueReminders()  -> run daily via a time-driven trigger
 *   doPost(e)           -> function when webhook is triggered
 */

const PROPS = PropertiesService.getScriptProperties();
const REMIND_DAYS_BEFORE = 7; // reminders start this many days before the due date

function getConfig_() {
  return {
    token: PROPS.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    userId: PROPS.getProperty('LINE_USER_ID'),
    sheetName: PROPS.getProperty('SHEET_NAME') || 'Payments',
    sheetId: PROPS.getProperty('SHEET_ID')
  };
}

function getSheet_() {
  const { sheetId, sheetName } = getConfig_();
  Logger.log('sheetId: ' + sheetId);
  return SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
}
 
/* ============================================================
 *  DATE HELPERS
 * ============================================================ */
 
function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
 
// Get due date in a given month (month starts at 0)
// month could be -1 or 12 etc. and JS Date rolls the year over automatically
function dueDateIn_(year, month, dueDay) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  // Make sure that the date won't exceed number of days in specific month
  return new Date(year, month, Math.min(Number(dueDay), daysInMonth));
}
 
function daysBetween_(from, to) {
  return Math.round((to - from) / 86400000);
}
 
/* ============================================================
 *  DAILY REMINDER
 * ============================================================ */
 
function sendDueReminders() {
  const { token, userId } = getConfig_();
  const data = getSheet_().getDataRange().getValues();
  const today = startOfDay_(new Date());
  const tz = Session.getScriptTimeZone();
  const bubbles = [];
 
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const [name, amount, dueDay, lastPaid] = data[i];
    if (!name || !dueDay) continue;
 
    const due = findDueCycle_(today, dueDay, lastPaid);
    if (!due) continue;
 
    const daysLeft = daysBetween_(today, due);
    const dueLabel = Utilities.formatDate(due, tz, 'd MMM yyyy');
    // different bubble will be created based on availability of the amount
    if (amount){
      bubbles.push(buildBubbleWithAmount_(name, amount, dueDay, row));
    }
    else {
      bubbles.push(buildBubbleWithoutAmount_(name, dueDay, row));
    }
  }
 
  Logger.log('Cards to send: ' + bubbles.length);
  if (bubbles.length === 0) return;
 
  pushMessage_(token, userId, [{
    type: 'flex',
    altText: `You have ${bubbles.length} payment(s) to pay`,
    contents: { type: 'carousel', contents: bubbles }
  }]);
}
 
/**
 * Returns the due Date of the earliest cycle that is unpaid and whose
 * reminder window is open (previous, this, next month are checked so
 * overdue items keep nagging and early-month due dates get reminded
 * in the previous month). Returns null if nothing to remind.
 */
function findDueCycle_(today, dueDay, lastPaid) {
  const y = today.getFullYear();
  const m = today.getMonth();
  const paidOn = lastPaid instanceof Date ? startOfDay_(lastPaid) : null;

  // Offset -> -1 is last month, 0 is current month, 1 is next month
  for (const offset of [-1, 0, 1]) {
    const due = dueDateIn_(y, m + offset, dueDay);
    const windowStart = new Date(due.getTime() - REMIND_DAYS_BEFORE * 86400000);
    const prevDue = dueDateIn_(y, m + offset - 1, dueDay); // last cycle's due date

    if (today < windowStart) return null;
    if (paidOn && paidOn > prevDue) continue; // paid any time after the last due date covers this cycle
    return due;
  }
  return null;
}

function buildBubbleWithAmount_(name, amount, dueDay, row) {
  return {
    type: 'bubble',
    size: 'kilo',
    body: {
      type: 'box',
      layout: 'vertical',
      spacing: 'sm',
      contents: [
        { type: 'text', text: name, weight: 'bold', size: 'lg' },
        { type: 'text', text: `฿${amount}`, size: 'xl', color: '#1DB446' },
        { type: 'text', text: `Due day: ${dueDay} of this month`, size: 'sm', color: '#888888' }
      ]
    },
    footer: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'button',
          style: 'primary',
          color: '#1DB446',
          action: {
            type: 'postback',
            label: 'Mark as Paid',
            data: `action=paid&row=${row}`,
            displayText: `Marking "${name}" as paid...`
          }
        }
      ]
    }
  };
}

function buildBubbleWithoutAmount_(name, dueDay, row) {
  return {
    type: 'bubble',
    size: 'kilo',
    body: {
      type: 'box',
      layout: 'vertical',
      spacing: 'sm',
      contents: [
        { type: 'text', text: name, weight: 'bold', size: 'lg' },
        { type: 'text', text: 'check the amount in their own app', size: 'md', color: '#aaaaaa', style: 'italic', wrap: true},
        { type: 'text', text: `Due day: ${dueDay} of this month`, size: 'sm', color: '#888888' }
      ]
    },
    footer: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'button',
          style: 'primary',
          color: '#1DB446',
          action: {
            type: 'postback',
            label: 'Mark as Paid',
            data: `action=paid&row=${row}`,
            displayText: `Marking "${name}" as paid...`
          }
        }
      ]
    }
  };
}

/* ============================================================
 *  WEBHOOK — when user clicks "Mark as paid"
 * ============================================================ */

function doPost(e) {
  const { token } = getConfig_();
  const body = JSON.parse(e.postData.contents);
  const events = body.events || [];

  events.forEach(event => {
    const userId = event.source && event.source.userId;

    if (event.type === 'postback') {
      const params = parseQueryString_(event.postback.data);
      if (params.action === 'paid' && params.row) {
        const name = markRowPaid_(Number(params.row));
        if (event.replyToken) {
          replyMessage_(token, event.replyToken, [
            { type: 'text', text: `✅ ${name} marked as paid.` }
          ]);
        }
      }
    } 
    // Get user ID
    // else if (event.type === 'message' && event.replyToken) {
    //   // Temporary: reply back with your own userId so you can copy it
    //   replyMessage_(token, event.replyToken, [
    //     { type: 'text', text: `Your userId is:\n${userId}` }
    //   ]);
    // }
  });

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function markRowPaid_(row) {
  const sheet = getSheet_();
  const name = sheet.getRange(row, 1).getValue();
  sheet.getRange(row, 4).setValue(new Date()); // column D: Last Paid
  return name;
}

function parseQueryString_(str) {
  const result = {};
  str.split('&').forEach(pair => {
    const [k, v] = pair.split('=');
    result[decodeURIComponent(k)] = decodeURIComponent(v || '');
  });
  return result;
}

/* ============================================================
 *  LINE API (push and reply API)
 * ============================================================ */

function pushMessage_(token, to, messages) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to, messages }),
    muteHttpExceptions: true
  });
}

function replyMessage_(token, replyToken, messages) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken, messages }),
    muteHttpExceptions: true
  });
}
