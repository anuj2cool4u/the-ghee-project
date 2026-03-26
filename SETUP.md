# Website Form Setup Guide — The Ghee Project

Connect the website's order and enquiry forms to Google Sheets + email notifications.

**Time required: ~10 minutes.**
**Cost: zero.** Google Sheets, Apps Script, and Gmail are all free.

---

## How it works

1. Visitor fills out a form and clicks Submit
2. The page sends the data to a Google Apps Script "Web App" you own
3. The script does two things simultaneously:
   - Appends a new row to a Google Sheet you control
   - Sends you an email notification with the enquiry details
4. You action the enquiry directly from your inbox or the sheet

---

## Step 1 — Create your Google Sheet

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it: **The Ghee Project — Enquiries**
3. Right-click the tab at the bottom → **Rename** → call it `enquiries`
4. Leave it blank — the script creates the column headers automatically on first submission

---

## Step 2 — Open Apps Script

1. In the Google Sheet, click **Extensions → Apps Script**
2. A new Apps Script project opens in a new browser tab
3. **Delete all existing code** in the editor
4. **Paste the entire script below** into the editor

---

## The complete Apps Script (paste this entire block)

```javascript
/**
 * The Ghee Project — Form Handler v2
 * ─────────────────────────────────────────────────────────────────────────────
 * Handles POST requests from the website forms (index.html, wholesale.html,
 * shop.html). For each submission:
 *   1. Appends a row to the "enquiries" Google Sheet
 *   2. Sends an email notification to NOTIFICATION_EMAIL
 *
 * IMPORTANT — Content-Type compatibility:
 *   The website sends requests with Content-Type: text/plain (required for
 *   browser no-cors mode). The body is still valid JSON and is parsed below.
 *
 * SETUP: Update NOTIFICATION_EMAIL to the address that should receive alerts.
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── CONFIG ───────────────────────────────────────────────────────────────────

const NOTIFICATION_EMAIL = 'anujchopra30@gmail.com'; // ← change to your email
const SHEET_NAME         = 'enquiries';

// ── MAIN POST HANDLER ────────────────────────────────────────────────────────

function doPost(e) {
  let data = {};

  try {
    // Google Apps Script redirects POST requests, which can drop the raw body.
    // The website sends data as: payload=<urlencoded-JSON>
    // So we read from e.parameter.payload (survives redirects), with a fallback
    // to e.postData.contents for direct (non-redirected) requests.
    let raw = '';

    if (e && e.parameter && e.parameter.payload) {
      // Primary path: form-encoded parameter — always works after redirect
      raw = e.parameter.payload;
    } else if (e && e.postData && e.postData.contents) {
      // Fallback: raw body — works only when there's no redirect
      raw = e.postData.contents.trim();
    }

    if (!raw) {
      throw new Error(
        'No data received. e.parameter: ' + JSON.stringify(e ? e.parameter : null) +
        ' | e.postData: ' + (e && e.postData ? e.postData.contents : 'null')
      );
    }

    data = JSON.parse(raw);

  } catch (parseErr) {
    logErrorToSheet('PARSE ERROR: ' + parseErr.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', step: 'parse', message: parseErr.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Step 1: log to sheet (always, even if email fails)
  try {
    logToSheet(data);
  } catch (sheetErr) {
    logErrorToSheet('SHEET ERROR: ' + sheetErr.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', step: 'sheet', message: sheetErr.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Step 2: send email notification to founders (non-fatal)
  try {
    sendNotification(data);
  } catch (mailErr) {
    logErrorToSheet('MAIL ERROR (row was saved): ' + mailErr.message);
  }

  // Step 3: send auto-reply confirmation to the enquirer (non-fatal)
  try {
    sendAutoReply(data);
  } catch (replyErr) {
    logErrorToSheet('AUTO-REPLY ERROR (row was saved): ' + replyErr.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET LOGGING ────────────────────────────────────────────────────────────

function logToSheet(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Tab named "' + SHEET_NAME + '" not found. ' +
      'Right-click a sheet tab → Rename → type exactly: enquiries'
    );
  }

  // Check whether row 1 already contains the 'Timestamp' header.
  // We check the specific cell rather than getLastRow() === 0 because
  // previous error rows may already exist in the sheet.
  const firstCell = sheet.getRange(1, 1).getValue();
  const hasHeaders = (String(firstCell).trim().toLowerCase() === 'timestamp');

  if (!hasHeaders) {
    // Insert a proper header row at the top, above any existing content
    const headers = buildHeaders(data);
    sheet.insertRowBefore(1);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold')
               .setBackground('#C8882A')
               .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  // Map each header column to the correct value
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row     = headers.map(h => {
    const key = headerToKey(h);
    return (data[key] !== undefined && data[key] !== null) ? String(data[key]) : '';
  });

  sheet.appendRow(row);

  // Auto-resize columns periodically
  if (sheet.getLastRow() % 25 === 0) {
    sheet.autoResizeColumns(1, lastCol);
  }
}

/**
 * Write a plain error message row so failures are visible in the sheet.
 */
function logErrorToSheet(message) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    sheet.appendRow([new Date().toISOString(), 'ERROR', message]);
  } catch (e) {
    // If even this fails, nothing we can do — check Apps Script execution logs
  }
}

// ── EMAIL NOTIFICATION ───────────────────────────────────────────────────────

function sendNotification(data) {
  const formType = data['form_type'] || 'unknown';

  const subjectMap = {
    'wholesale_enquiry': 'New wholesale enquiry — The Ghee Project',
    'shop_order':        'New shop order enquiry — The Ghee Project',
    'contact':           'New contact message — The Ghee Project',
  };
  const subject = subjectMap[formType] || 'New form submission — The Ghee Project';

  const labelMap = {
    '_timestamp':       'Submitted at',
    'form_type':        'Form type',
    '_page':            'From page',
    'name':             'Name',
    'email':            'Email',
    'phone':            'Phone',
    'business':         'Business name',
    'business_type':    'Business type',
    'subject':          'Subject',
    'product':          'Product interest',
    'products':         'Products of interest',
    'quantity':         'Quantity',
    'estimated_volume': 'Est. monthly volume',
    'suburb':           'Suburb / postcode',
    'message':          'Message',
    'notes':            'Notes',
  };

  const displayOrder = [
    '_timestamp', 'form_type', '_page',
    'name', 'email', 'phone',
    'business', 'business_type',
    'subject', 'product', 'products', 'quantity', 'estimated_volume', 'suburb',
    'message', 'notes',
  ];

  // ── Plain text version ──
  const lines = [
    'New enquiry from thegeeproject.com.au',
    '',
    '─────────────────────────────────────',
  ];
  displayOrder.forEach(key => {
    if (data[key] && String(data[key]).trim()) {
      lines.push((labelMap[key] || key) + ': ' + data[key]);
    }
  });
  Object.keys(data).forEach(key => {
    if (!displayOrder.includes(key) && data[key] && String(data[key]).trim()) {
      lines.push((labelMap[key] || key) + ': ' + data[key]);
    }
  });
  lines.push('─────────────────────────────────────');
  lines.push('Reply to this email to respond to ' + (data['name'] || 'the enquirer') + '.');

  // ── HTML version ──
  const tableRows = displayOrder
    .filter(key => data[key] && String(data[key]).trim())
    .map(key => {
      const label = labelMap[key] || key;
      const val   = String(data[key]).replace(/</g,'&lt;').replace(/>/g,'&gt;');
      return '<tr>' +
        '<td style="padding:9px 14px;font-weight:600;color:#7B3F00;white-space:nowrap;vertical-align:top;font-size:13px">' + label + '</td>' +
        '<td style="padding:9px 14px;color:#1A0900;font-size:13px">' + val + '</td>' +
        '</tr>';
    });

  // Add any extra fields
  Object.keys(data).forEach(key => {
    if (!displayOrder.includes(key) && data[key] && String(data[key]).trim()) {
      const label = labelMap[key] || key;
      const val   = String(data[key]).replace(/</g,'&lt;').replace(/>/g,'&gt;');
      tableRows.push(
        '<tr>' +
        '<td style="padding:9px 14px;font-weight:600;color:#7B3F00;white-space:nowrap;vertical-align:top;font-size:13px">' + label + '</td>' +
        '<td style="padding:9px 14px;color:#1A0900;font-size:13px">' + val + '</td>' +
        '</tr>'
      );
    }
  });

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">' +
      '<div style="background:#1A0900;padding:20px 28px">' +
        '<p style="font-family:Georgia,serif;font-size:20px;font-weight:700;color:#E5AB50;margin:0">The Ghee Project</p>' +
        '<p style="font-size:11px;color:rgba(254,252,248,0.5);margin:4px 0 0;letter-spacing:0.12em;text-transform:uppercase">New Form Submission</p>' +
      '</div>' +
      '<div style="background:#FAF3E4;padding:20px 28px">' +
        '<h2 style="font-family:Georgia,serif;font-size:17px;color:#1A0900;margin:0 0 14px">' + subject + '</h2>' +
        '<table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #F0E4C7">' +
          tableRows.join('') +
        '</table>' +
        '<p style="margin:16px 0 0;font-size:12px;color:#8A6840">' +
          'Reply to this email to respond directly. All enquiries are also logged in your Google Sheet.' +
        '</p>' +
      '</div>' +
    '</div>';

  const mailOptions = {
    name:     'The Ghee Project Website',
    htmlBody: htmlBody,
  };
  if (data['email'] && data['email'].includes('@')) {
    mailOptions.replyTo = data['email'];
  }

  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, lines.join('\n'), mailOptions);
}

// ── AUTO-REPLY TO ENQUIRER ───────────────────────────────────────────────────

function sendAutoReply(data) {
  // Only send if we have a valid email address
  if (!data['email'] || !data['email'].includes('@')) return;

  const firstName = (data['name'] || 'there').split(' ')[0];
  const formType  = data['form_type'] || 'contact';

  // Tailor subject + body copy by form type
  const config = {
    wholesale_enquiry: {
      subject:     'We got your wholesale enquiry — The Ghee Project',
      heading:     'Thanks for reaching out, ' + firstName + '.',
      body:        'We\'ve received your trade enquiry and one of our team will be in touch within 1–2 business days to discuss your requirements.',
      footer:      'In the meantime, feel free to reply to this email with any questions.',
    },
    shop_order: {
      subject:     'We got your order enquiry — The Ghee Project',
      heading:     'Thanks for your interest, ' + firstName + '.',
      body:        'We\'ve received your order enquiry and will be in touch shortly to confirm availability and next steps.',
      footer:      'If you need anything urgently, just reply to this email.',
    },
    contact: {
      subject:     'We got your message — The Ghee Project',
      heading:     'Thanks for getting in touch, ' + firstName + '.',
      body:        'We\'ve received your message and will get back to you within 1–2 business days.',
      footer:      'In the meantime, feel free to reply if you have anything to add.',
    },
  };

  const c = config[formType] || config['contact'];

  // ── Plain text ──
  const plainText = [
    c.heading,
    '',
    c.body,
    '',
    c.footer,
    '',
    '────────────────────────────',
    'The Ghee Project',
    'Small Batch. Pure. Local.',
    'hello@thegeeproject.com.au',
    'thegeeproject.com.au',
    '────────────────────────────',
    '',
    'This is an automated confirmation — your message has been logged and a real person will follow up.',
  ].join('\n');

  // ── HTML ──
  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#FAF3E4">' +
      '<div style="background:#1A0900;padding:24px 32px">' +
        '<p style="font-family:Georgia,serif;font-size:22px;font-weight:700;color:#E5AB50;margin:0">The Ghee Project</p>' +
        '<p style="font-size:11px;color:rgba(254,252,248,0.45);margin:5px 0 0;letter-spacing:0.14em;text-transform:uppercase">Small Batch. Pure. Local.</p>' +
      '</div>' +
      '<div style="padding:32px">' +
        '<h2 style="font-family:Georgia,serif;font-size:19px;color:#1A0900;margin:0 0 12px">' + c.heading + '</h2>' +
        '<p style="font-size:15px;color:#3D1A00;line-height:1.7;margin:0 0 12px">' + c.body + '</p>' +
        '<p style="font-size:15px;color:#3D1A00;line-height:1.7;margin:0 0 28px">' + c.footer + '</p>' +
        '<div style="border-top:1px solid #F0E4C7;padding-top:20px">' +
          '<p style="font-size:13px;color:#8A6840;margin:0;line-height:1.6">' +
            '<strong style="color:#1A0900">The Ghee Project</strong><br>' +
            'Sydney, NSW, Australia<br>' +
            '<a href="mailto:hello@thegeeproject.com.au" style="color:#C8882A;text-decoration:none">hello@thegeeproject.com.au</a>' +
          '</p>' +
        '</div>' +
      '</div>' +
      '<div style="background:#F0E4C7;padding:12px 32px">' +
        '<p style="font-size:11px;color:#8A6840;margin:0">This is an automated confirmation. A real person will follow up shortly.</p>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail(data['email'], c.subject, plainText, {
    name:     'The Ghee Project',
    htmlBody: htmlBody,
    replyTo:  NOTIFICATION_EMAIL,
  });
}

// ── HELPERS ──────────────────────────────────────────────────────────────────

function buildHeaders(data) {
  const labelMap = {
    '_timestamp':       'Timestamp',
    'form_type':        'Form type',
    '_page':            'Source page',
    'name':             'Name',
    'email':            'Email',
    'phone':            'Phone',
    'business':         'Business name',
    'business_type':    'Business type',
    'subject':          'Subject',
    'product':          'Product interest',
    'products':         'Products interest',
    'quantity':         'Quantity',
    'estimated_volume': 'Est. monthly volume',
    'suburb':           'Suburb / postcode',
    'message':          'Message',
    'notes':            'Notes',
  };

  const priorityKeys = [
    '_timestamp', 'form_type', '_page',
    'name', 'email', 'phone',
    'business', 'business_type',
    'subject', 'product', 'products', 'quantity', 'estimated_volume', 'suburb',
    'message', 'notes',
  ];

  const headers = [];
  priorityKeys.forEach(k => {
    if (data[k] !== undefined) headers.push(labelMap[k] || k);
  });
  Object.keys(data).forEach(k => {
    const label = labelMap[k] || k;
    if (!headers.includes(label)) headers.push(label);
  });
  return headers;
}

function headerToKey(header) {
  const map = {
    'Timestamp':           '_timestamp',
    'Form type':           'form_type',
    'Source page':         '_page',
    'Name':                'name',
    'Email':               'email',
    'Phone':               'phone',
    'Business name':       'business',
    'Business type':       'business_type',
    'Subject':             'subject',
    'Product interest':    'product',
    'Products interest':   'products',
    'Quantity':            'quantity',
    'Est. monthly volume': 'estimated_volume',
    'Suburb / postcode':   'suburb',
    'Message':             'message',
    'Notes':               'notes',
  };
  return map[header] || header.toLowerCase().replace(/[^a-z0-9]/g, '_');
}

// ── GET (health check) ───────────────────────────────────────────────────────

function doGet() {
  return ContentService
    .createTextOutput('The Ghee Project form handler is running OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── TEST FUNCTION (run manually from Apps Script editor) ─────────────────────
// Select 'testSubmission' in the function dropdown and click Run (▶).
// This simulates a real form POST. Check your sheet and inbox afterward.

function testSubmission() {
  // Simulate exactly what the website sends: a form-encoded payload parameter
  const fakePayload = {
    form_type:        'wholesale_enquiry',
    _timestamp:       new Date().toISOString(),
    _page:            '/wholesale.html',
    name:             'Test User',
    email:            NOTIFICATION_EMAIL,
    phone:            '0400 000 000',
    business:         'Test Restaurant',
    business_type:    'restaurant',
    products:         'Original Pure Ghee',
    estimated_volume: '5–20kg/month',
    message:          'This is a test submission from the Apps Script editor.',
  };

  // Mimic what the browser sends: e.parameter.payload = JSON string
  const fakeEvent = {
    parameter: { payload: JSON.stringify(fakePayload) },
    postData:  null,
  };

  doPost(fakeEvent);
  Logger.log(
    'testSubmission complete.\n' +
    '1. Check the "enquiries" sheet for a new row.\n' +
    '2. Check ' + NOTIFICATION_EMAIL + ' for the founder notification email.\n' +
    '3. Check the same inbox for the auto-reply confirmation (it sends to the email field, which is also NOTIFICATION_EMAIL in this test).'
  );
}
```

---

## Step 3 — Set your notification email

At the top of the script, change this line:

```javascript
const NOTIFICATION_EMAIL = 'hello@thegeeproject.com.au';
```

Replace with whichever email address you want the notifications sent to.
This **must** be a Google account email or an alias connected to your Google account
(Google Apps Script can only send from/to Google-authenticated addresses).

---

## Step 4 — Deploy the Web App

1. Click **Deploy → New deployment** (top-right in Apps Script editor)
2. Click the **gear icon** ⚙ next to "Type" → select **Web app**
3. Configure it:
   - Description: `Ghee Project form handler`
   - Execute as: **Me**
   - Who has access: **Anyone** ← this is required so the website can reach it
4. Click **Deploy**
5. When prompted, click **Authorise access**, then choose your Google account and click **Allow**
6. After deployment, you'll see a **Web app URL** like:
   ```
   https://script.google.com/macros/s/AKfycbXXXXXXXX/exec
   ```
7. **Copy that URL** — you'll need it in Step 5

> **Note:** If Google shows a warning saying "This app isn't verified", click
> **Advanced → Go to [project name] (unsafe)**. This is expected for personal
> scripts you wrote yourself.

---

## Step 5 — Add the URL to the three HTML files

Open each HTML file and replace the placeholder `YOUR_SCRIPT_URL` with the URL you copied.

There is exactly **one instance** in each file — find it with Ctrl+F.

### index.html
```html
<!-- Find: -->
data-script-url="YOUR_SCRIPT_URL"

<!-- Replace with: -->
data-script-url="https://script.google.com/macros/s/YOUR_ACTUAL_ID/exec"
```

Do the same in **wholesale.html** and **shop.html**.

---

## Step 6 — Test from the Apps Script editor first (recommended)

Before testing from the website, run the built-in test function to confirm the script itself works:

1. In the Apps Script editor, click the function dropdown (it usually shows `doPost` or `doGet`) and select **`testSubmission`**
2. Click the **Run** button (▶)
3. If prompted to grant permissions, click **Review permissions → Allow**
4. Check:
   - Your **Google Sheet** — a new row should appear with test data
   - Your **inbox** at `NOTIFICATION_EMAIL` — a test email should arrive

If that works, the script is fine. If it throws an error, the error message in the Apps Script execution log will tell you exactly what's wrong (usually a wrong sheet name or email address).

## Step 7 — Test from the website

1. Open `wholesale.html` in your browser
2. Fill out the trade enquiry form and click submit
3. Within a few seconds you should see "Message sent" on the form, and:
   - A new row in the Google Sheet
   - An email notification in your inbox

**If the row still doesn't appear**, check these in order:
- **Sheet tab name:** make sure the tab is named exactly `enquiries` (lowercase, no spaces)
- **URL correctness:** open your Apps Script deployment URL directly in a new browser tab — you should see: `The Ghee Project form handler is running OK`
- **Access setting:** the deployment must be **Anyone** (not "Anyone with Google account")
- **Re-deployed after pasting the new script?** Editing the script code requires a new deployment — see the section below
- **Execution log:** in Apps Script, go to **Executions** (left sidebar icon) to see recent runs and any error messages

---

## What the email looks like

You'll receive a nicely formatted HTML email with:
- Subject line indicating the form type (e.g. "🛒 New wholesale enquiry")
- All submitted fields in a readable table
- **Reply-To set to the enquirer's email** — so you can reply directly from Gmail without copying their address

---

## Managing enquiries in the Sheet

All submissions land in the `enquiries` tab. Suggested workflow:
- Add a **Status** column manually (column after the last data column) and use it to track: `New`, `Replied`, `Done`
- Use **Data → Create a filter view** to see only wholesale enquiries, or only new ones
- Turn on **Tools → Notification settings → When any changes are made** for a second email alert if needed

---

## Re-deploying after script changes

If you edit the script code, you must **create a new deployment** to push the changes live:
- Go to **Deploy → New deployment**
- Use the same settings as before
- Copy the new URL and replace it in all three HTML files

Updating an existing deployment does **not** push code changes to the live URL.

---

## File structure reference

```
website/
├── index.html           ← Homepage — contact/order form
├── wholesale.html       ← Wholesale page — trade enquiry form
├── shop.html            ← Shop page — order enquiry form
├── SETUP.md             ← This file
└── assets/
    ├── css/main.css     ← Shared styles
    └── js/main.js       ← Shared scripts (form handler, animations)
```

---

## Going live (when you have your domain)

When `thegeeproject.com.au` is registered and hosted:
1. Upload all files in the `website/` folder to your server root
2. Update the `<link rel="canonical">` in each HTML file to the real URL
3. Submit to Google Search Console
4. No build step needed — plain HTML files

---

*Last updated: March 2026 · The Ghee Project*
