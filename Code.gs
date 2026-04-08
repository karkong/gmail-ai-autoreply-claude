// ============================================================
// Gmail AI Auto-Reply + Cold Outreach Bot
// Google Apps Script + Claude API
// ============================================================
//
// HOW TO USE:
//   1. Open script.google.com — paste this entire file
//   2. Fill in YOUR details in the CONFIG block below
//   3. Edit PRODUCT_CONTEXT with your product/service details
//   4. Edit the BUYERS array with your target contacts
//   5. Run setupTrigger() ONCE to activate the 5-minute schedule
//   6. Authorize permissions when Google prompts you
//   7. Run sendColdOutreach() in PREVIEW mode first, then live
//
// DAILY ROUTINE:
//   - Check your Gmail label "AI-Replied" each morning
//   - Review the log sheet in Google Sheets
//   - Manually respond to anything needing negotiation or decisions
// ============================================================


// ─── CONFIGURATION — Fill in your own details ───────────────
const CONFIG = {

  // Your Claude API key from console.anthropic.com → API Keys
  CLAUDE_API_KEY: 'YOUR_CLAUDE_API_KEY_HERE',

  // Your Gmail address (used to filter out emails you sent)
  GMAIL_ADDRESS: 'YOUR_GMAIL_ADDRESS@gmail.com',

  // Gmail labels — created automatically if they don't exist
  LABEL_REPLIED: 'AI-Replied',  // Applied after bot sends a reply
  LABEL_SKIP:    'AI-Skip',     // Apply manually to threads you want the bot to ignore

  // Google Sheet name for the reply log
  // After first run, replace YOUR_SPREADSHEET_ID_HERE with your actual sheet ID
  // (found in the sheet URL: docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit)
  LOG_SHEET_NAME: 'AI Reply Log',

  // Truncate long emails before sending to Claude (saves API cost)
  MAX_BODY_CHARS: 800,

  // Claude model — Haiku is fast and cheap (~$0.001 per reply)
  REPLY_MODEL: 'claude-haiku-4-5-20251001',

};


// ─── PRODUCT CONTEXT — Edit with your own product/service ────
// This is injected into every auto-reply Claude generates.
// Replace all placeholder values with your actual details.

const PRODUCT_CONTEXT = `
You are a professional sales assistant handling email inquiries on behalf of [SELLER NAME] ([SELLER TITLE]) for a [PRODUCT TYPE] sale based in [YOUR CITY], [YOUR STATE/COUNTRY].

PRODUCT DETAILS:
- Product: [YOUR PRODUCT NAME AND DESCRIPTION]
- Quantity available: [YOUR QUANTITY]
- Condition: [PRODUCT CONDITION — e.g. new, surplus, as-is]
- Certification: [ANY CERTIFICATIONS OR NOTE IF NONE]
- Pricing: [YOUR PRICING — e.g. $X per kg, negotiable for quantities over Y kg]
- Location: [YOUR LOCATION]. [FREIGHT/PICKUP OPTIONS]
- Transport: [ANY TRANSPORT CLASSIFICATION INFO]

CRITICAL RULES:
1. Do NOT claim any certifications or specs you cannot verify.
2. Do NOT quote prices outside the range specified above.
3. Be honest about the product condition — do not oversell.
4. Sound human, professional, and direct — not like a bot.
5. [ADD YOUR OWN RULES HERE]

RESPONSE GUIDE BY INQUIRY TYPE:
- Price inquiry → State pricing is [YOUR PRICE], negotiable for large quantities, invite them to contact for best price.
- Sample request → Confirm samples can be arranged, ask them to call or reply to coordinate.
- Certification request → [YOUR RESPONSE TO CERT QUESTIONS]
- Inspection request → Welcome the request, ask them to call to arrange a time at [YOUR LOCATION].
- Purchase / order inquiry → Express strong interest, encourage them to call directly for fast resolution.
- General inquiry → Give a brief honest product overview, invite next steps.

REPLY FORMAT:
- 100 to 150 words maximum
- Professional but direct tone
- End with a clear call to action (call, reply, or arrange inspection)
- Sign off exactly as follows:

Kind regards,
[YOUR FULL NAME]
[YOUR TITLE] | [YOUR BUSINESS NAME]
[YOUR CITY], [YOUR STATE]
`;


// ─── MAIN FUNCTION — Runs on timer every 5 minutes ───────────

function checkAndReply() {
  const repliedLabel = getOrCreateLabel(CONFIG.LABEL_REPLIED);
  const skipLabel    = getOrCreateLabel(CONFIG.LABEL_SKIP);
  const sheet        = getOrCreateLogSheet();

  // Search inbox — only matching subjects, exclude auto-replies and automated senders
  // IMPORTANT: Replace "Your Keyword" with a word always in your buyers' subject lines
  const query   = 'in:inbox is:unread -from:' + CONFIG.GMAIL_ADDRESS +
                  ' subject:"YOUR KEYWORD"' +          // ← e.g. subject:"Molecular Sieve"
                  ' -subject:"automatic reply"' +
                  ' -subject:"auto reply"' +
                  ' -subject:"auto-reply"' +
                  ' -subject:"auto response"' +
                  ' -subject:"out of office"' +
                  ' -subject:"autoreply"' +
                  ' -subject:"do not reply"';
  const threads = GmailApp.search(query, 0, 20);

  Logger.log('Run started — found ' + threads.length + ' candidate thread(s).');

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];

    try {
      // Skip threads already labelled
      const threadLabels = thread.getLabels().map(function(l) { return l.getName(); });
      if (threadLabels.indexOf(CONFIG.LABEL_REPLIED) !== -1) {
        Logger.log('Skip (already replied): ' + thread.getFirstMessageSubject());
        continue;
      }
      if (threadLabels.indexOf(CONFIG.LABEL_SKIP) !== -1) {
        Logger.log('Skip (manual skip label): ' + thread.getFirstMessageSubject());
        continue;
      }

      // Get the most recent message in the thread
      const messages = thread.getMessages();
      const lastMsg  = messages[messages.length - 1];
      const sender   = lastMsg.getFrom();
      const subject  = lastMsg.getSubject();
      const body     = lastMsg.getPlainBody().substring(0, CONFIG.MAX_BODY_CHARS).trim();

      // Block auto-reply subjects (second filter in case Gmail search misses them)
      const subjectLower = subject.toLowerCase();
      const autoReplyKeywords = ['automatic reply', 'auto reply', 'auto-reply', 'autoreply',
                                  'auto response', 'out of office', 'do not reply',
                                  'thank you for contacting', 'we have received your email',
                                  'we\'ve received your', 'this is an automated'];
      const isAutoReply = autoReplyKeywords.some(function(kw) { return subjectLower.indexOf(kw) !== -1; });
      if (isAutoReply) {
        Logger.log('Skip (auto-reply subject): ' + subject);
        thread.addLabel(getOrCreateLabel(CONFIG.LABEL_SKIP));
        thread.markRead();
        continue;
      }

      // Block automated/no-reply senders
      const senderLower = sender.toLowerCase();
      const blockedSenders = ['no-reply', 'noreply', 'do-not-reply', 'donotreply',
                              'mailer-daemon', 'postmaster', 'notifications@', 'newsletter',
                              'welcome@', 'rewards@', 'marketing@', 'info@email.', 'bounce'];
      const isBlocked = blockedSenders.some(function(term) { return senderLower.indexOf(term) !== -1; });
      if (isBlocked) {
        Logger.log('Skip (automated sender): ' + sender);
        thread.addLabel(getOrCreateLabel(CONFIG.LABEL_SKIP));
        continue;
      }

      Logger.log('Processing: "' + subject + '" from ' + sender);

      // Call Claude API to generate reply
      const replyText = callClaudeAPI(sender, subject, body);

      if (!replyText || replyText.length < 10) {
        Logger.log('Claude returned empty or very short response — skipping.');
        sheet.appendRow([new Date(), sender, subject, body.substring(0, 200), '(empty response)', 'SKIPPED']);
        continue;
      }

      // Send reply in the same thread
      thread.reply(replyText);

      // Apply "AI-Replied" label
      thread.addLabel(repliedLabel);

      // Mark thread as read
      thread.markRead();

      // Log to Google Sheet
      sheet.appendRow([
        new Date(),
        sender,
        subject,
        body.substring(0, 250),
        replyText.substring(0, 350),
        'SENT'
      ]);

      Logger.log('Reply sent to: ' + sender);

      // Small pause to avoid Gmail rate limits
      Utilities.sleep(2000);

    } catch (err) {
      Logger.log('ERROR on thread "' + thread.getFirstMessageSubject() + '": ' + err.toString());
      sheet.appendRow([new Date(), '', thread.getFirstMessageSubject(), '', '', 'ERROR: ' + err.toString()]);
    }
  }

  Logger.log('Run complete.');
}


// ─── CLAUDE API CALL ─────────────────────────────────────────

function callClaudeAPI(senderEmail, subject, emailBody) {

  const userPrompt =
    PRODUCT_CONTEXT +
    '\n---\nINCOMING EMAIL\n' +
    'From: ' + senderEmail + '\n' +
    'Subject: ' + subject + '\n' +
    'Body:\n' + emailBody +
    '\n---\n\n' +
    'Write a professional reply email body only (no subject line, no "Subject:" prefix). ' +
    'Follow all instructions and rules above exactly.';

  const payload = {
    model:      CONFIG.REPLY_MODEL,
    max_tokens: 400,
    messages: [{ role: 'user', content: userPrompt }]
  };

  const options = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         CONFIG.CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response     = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('Claude API returned HTTP ' + responseCode + ': ' + responseText);
  }

  const json = JSON.parse(responseText);

  if (!json.content || json.content.length === 0) {
    throw new Error('Claude API returned no content in response.');
  }

  return json.content[0].text.trim();
}


// ─── HELPER: Get or create a Gmail label ─────────────────────

function getOrCreateLabel(name) {
  const existing = GmailApp.getUserLabelByName(name);
  if (existing) return existing;
  Logger.log('Creating label: ' + name);
  return GmailApp.createLabel(name);
}


// ─── HELPER: Get or create the log Google Sheet ──────────────

function getOrCreateLogSheet() {
  // Paste your Google Sheet ID here after first run to avoid duplicate sheets
  // Found in sheet URL: docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit
  var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

  var ss;
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    // First run — create a new spreadsheet automatically
    ss = SpreadsheetApp.create('AI Reply Log');
    Logger.log('Created new log spreadsheet. Copy this ID into SPREADSHEET_ID: ' + ss.getId());
  } else {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'From', 'Subject', 'Email Preview (250 chars)', 'Reply Preview (350 chars)', 'Status']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 280);
    sheet.setColumnWidth(5, 320);
    sheet.setColumnWidth(6, 100);
    Logger.log('Log sheet created.');
  }

  return sheet;
}


// ─── SETUP: Run ONCE to activate the 5-minute trigger ────────

function setupTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('checkAndReply')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Trigger activated: checkAndReply will run every 5 minutes.');
  Logger.log('Setup complete. Bot is now live.');
}


// ─── UTILITY: Test the bot manually ──────────────────────────

function testBotManually() {
  Logger.log('=== MANUAL TEST RUN ===');
  checkAndReply();
  Logger.log('=== TEST RUN COMPLETE — check Gmail and log sheet ===');
}


// ─── UTILITY: Clear log sheets (keeps headers) ───────────────

function clearLogs() {
  var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var replySheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (replySheet && replySheet.getLastRow() > 1) {
    replySheet.deleteRows(2, replySheet.getLastRow() - 1);
    Logger.log('Cleared: ' + CONFIG.LOG_SHEET_NAME);
  } else {
    Logger.log('AI Reply Log already empty.');
  }

  var outreachSheet = ss.getSheetByName('Outreach Status');
  if (outreachSheet && outreachSheet.getLastRow() > 1) {
    outreachSheet.deleteRows(2, outreachSheet.getLastRow() - 1);
    Logger.log('Cleared: Outreach Status');
  } else {
    Logger.log('Outreach Status already empty.');
  }

  Logger.log('Both logs cleared. Fresh start — headers preserved.');
}


// ─── UTILITY: Pause or resume the bot ────────────────────────

function pauseBot() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  Logger.log('Bot paused — all triggers removed. Run setupTrigger() to restart.');
}


// ============================================================
// COLD OUTREACH SENDER
// Sends tailored cold emails to your buyer list via Claude AI.
//
//   STEP 1: Set PREVIEW_ONLY = true (default) — review first
//   STEP 2: Run sendColdOutreach() — emails logged to sheet
//   STEP 3: Check "Outreach Status" tab — review all drafts
//   STEP 4: Set PREVIEW_ONLY = false
//   STEP 5: Run sendColdOutreach() again — sends for real
//
// Already-sent emails are skipped automatically (no duplicates).
// ============================================================

var PREVIEW_ONLY = true; // ← Change to false only when ready to send

// ─── BUYER LIST — Replace with your own contacts ─────────────
// Add one entry per target company.
// Categories must match keys in CATEGORY_CONTEXT below.

var BUYERS = [

  // ── CATEGORY A ───────────────────────────────────────────────
  { company:  'Customer 1 Company Name',
    contact:  'Purchasing Manager',
    email:    'contact@customer1.com',
    category: 'category_a',
    priority: '★★★ HIGH',
    notes:    'Describe what makes this customer a strong lead and what angle to use in the email.' },

  { company:  'Customer 2 Company Name',
    contact:  'Sales Team',
    email:    'contact@customer2.com',
    category: 'category_a',
    priority: '★★ MEDIUM',
    notes:    'Describe what makes this customer a strong lead and what angle to use in the email.' },

  // ── CATEGORY B ───────────────────────────────────────────────
  { company:  'Customer 3 Company Name',
    contact:  'Technical Manager',
    email:    'contact@customer3.com',
    category: 'category_b',
    priority: '★★★ HIGH',
    notes:    'Describe what makes this customer a strong lead and what angle to use in the email.' },

  { company:  'Customer 4 Company Name',
    contact:  'Purchasing Team',
    email:    'contact@customer4.com',
    category: 'category_b',
    priority: '★★ MEDIUM',
    notes:    'Describe what makes this customer a strong lead and what angle to use in the email.' },

  // ── CATEGORY C ───────────────────────────────────────────────
  { company:  'Customer 5 Company Name',
    contact:  'Operations Manager',
    email:    'contact@customer5.com',
    category: 'category_c',
    priority: '★★★ HIGH',
    notes:    'Describe what makes this customer a strong lead and what angle to use in the email.' },

  // Add more buyers as needed — copy and paste a block above
];

// ─── CATEGORY CONTEXT — Describe each buyer type ─────────────
// Claude uses this to tailor each email to the buyer's industry.
// Add/rename categories to match your BUYERS list above.

var CATEGORY_CONTEXT = {

  category_a:
    'Describe Category A buyers: who they are, what they use your product for, ' +
    'what they care about most (price, local supply, volume, margin, etc.), ' +
    'and what angle to emphasise in the cold email.',

  category_b:
    'Describe Category B buyers: who they are, what they use your product for, ' +
    'what they care about most, and what angle to emphasise in the cold email.',

  category_c:
    'Describe Category C buyers: who they are, what they use your product for, ' +
    'what they care about most, and what angle to emphasise in the cold email.',
};


// ─── MAIN OUTREACH FUNCTION ───────────────────────────────────

function sendColdOutreach() {
  var sheet    = getOrCreateOutreachSheet();
  var sentList = getSentEmails(sheet);

  Logger.log('=== COLD OUTREACH RUN | PREVIEW_ONLY: ' + PREVIEW_ONLY + ' ===');
  Logger.log('Total buyers: ' + BUYERS.length + ' | Already sent: ' + sentList.length);

  for (var i = 0; i < BUYERS.length; i++) {
    var buyer = BUYERS[i];

    if (sentList.indexOf(buyer.email) !== -1) {
      Logger.log('Skip (already sent): ' + buyer.company);
      continue;
    }

    try {
      Logger.log('Generating email for: ' + buyer.company + ' [' + buyer.category + ']');

      var emailContent = generateOutreachEmail(buyer);

      if (!emailContent || !emailContent.subject || !emailContent.body) {
        Logger.log('ERROR: Claude returned incomplete response for ' + buyer.company);
        sheet.appendRow([new Date(), buyer.company, buyer.email, buyer.category,
                         buyer.priority, '', '', 'ERROR: incomplete response']);
        continue;
      }

      Logger.log('Subject: ' + emailContent.subject);
      Logger.log('Preview: ' + emailContent.body.substring(0, 120) + '...');

      if (!PREVIEW_ONLY) {
        GmailApp.sendEmail(
          buyer.email,
          emailContent.subject,
          emailContent.body,
          { name: '[YOUR NAME] | [YOUR BUSINESS]' }  // ← Replace with your name
        );
        Logger.log('SENT → ' + buyer.email);
        sheet.appendRow([new Date(), buyer.company, buyer.email, buyer.category,
                         buyer.priority, emailContent.subject, emailContent.body, 'SENT']);
      } else {
        Logger.log('PREVIEW (not sent): ' + buyer.company);
        sheet.appendRow([new Date(), buyer.company, buyer.email, buyer.category,
                         buyer.priority, emailContent.subject, emailContent.body, 'PREVIEW']);
      }

      Utilities.sleep(3000);

    } catch (err) {
      Logger.log('ERROR for ' + buyer.company + ': ' + err.toString());
      sheet.appendRow([new Date(), buyer.company, buyer.email, buyer.category,
                       buyer.priority, '', '', 'ERROR: ' + err.toString()]);
    }
  }

  Logger.log('=== OUTREACH RUN COMPLETE ===');
  if (PREVIEW_ONLY) {
    Logger.log('PREVIEW MODE — no emails sent. Check "Outreach Status" sheet.');
    Logger.log('Set PREVIEW_ONLY = false and run again to send for real.');
  } else {
    Logger.log('LIVE MODE — emails sent. Check "Outreach Status" sheet for confirmation.');
  }
}


// ─── CLAUDE EMAIL GENERATOR ───────────────────────────────────

function generateOutreachEmail(buyer) {

  var prompt =
    'You are writing a cold outreach email on behalf of [YOUR NAME] ([YOUR TITLE]) ' +
    'for a [YOUR PRODUCT/SERVICE] sale in [YOUR CITY], [YOUR COUNTRY].\n\n' +

    'SELLER: [YOUR NAME] | [YOUR TITLE] | [YOUR BUSINESS] | [YOUR CITY]\n' +
    'PRODUCT: [YOUR PRODUCT NAME AND KEY SPECS]\n' +
    'QUANTITY: [YOUR AVAILABLE QUANTITY]\n' +
    'CONDITION: [PRODUCT CONDITION]\n' +
    'CERTIFICATION: [CERTIFICATION STATUS]\n' +
    'PRICING: [YOUR PRICING STRUCTURE]\n' +
    'LOCATION: [YOUR LOCATION]. [FREIGHT OPTIONS]\n\n' +

    'MANDATORY — include all of the following in every email:\n' +
    '1. [YOUR MANDATORY DISCLOSURE 1]\n' +
    '2. [YOUR MANDATORY DISCLOSURE 2]\n' +
    '3. Pricing: [YOUR PRICE SUMMARY]\n' +
    '4. Location: [YOUR LOCATION]\n' +
    '5. [ANY BRAND OR CLAIM RESTRICTIONS]\n\n' +

    'RECIPIENT:\n' +
    'Company: ' + buyer.company + '\n' +
    'Contact: ' + buyer.contact + '\n' +
    'Buyer type context: ' + CATEGORY_CONTEXT[buyer.category] + '\n' +
    'Specific notes about this company: ' + buyer.notes + '\n\n' +

    'INSTRUCTIONS:\n' +
    '- Write a tailored cold outreach email for this specific company and their use case\n' +
    '- 150 to 200 words maximum (body only, excluding sign-off)\n' +
    '- Direct, factual, professional tone\n' +
    '- Weave the mandatory disclosures in naturally\n' +
    '- End with a single clear call to action\n\n' +

    'OUTPUT: Respond with ONLY this JSON — no extra text before or after:\n' +
    '{\n' +
    '  "subject": "the subject line here",\n' +
    '  "body": "full email body including greeting and sign-off"\n' +
    '}\n\n' +

    'Sign-off to use exactly:\n' +
    'Kind regards,\n' +
    '[YOUR FULL NAME]\n' +
    '[YOUR TITLE] | [YOUR BUSINESS]\n' +
    '[YOUR CITY], [YOUR STATE]';

  var payload = {
    model:      CONFIG.REPLY_MODEL,
    max_tokens: 700,
    messages:   [{ role: 'user', content: prompt }]
  };

  var options = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         CONFIG.CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  var code     = response.getResponseCode();
  var text     = response.getContentText();

  if (code !== 200) {
    throw new Error('Claude API error ' + code + ': ' + text.substring(0, 300));
  }

  var json    = JSON.parse(text);
  var rawText = json.content[0].text.trim();

  var jsonMatch = rawText.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('Claude did not return valid JSON. Raw: ' + rawText.substring(0, 200));
  }

  return JSON.parse(jsonMatch[0]);
}


// ─── OUTREACH SHEET HELPERS ───────────────────────────────────

function getOrCreateOutreachSheet() {
  var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

  var ss;
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    ss = SpreadsheetApp.create('AI Reply Log');
    Logger.log('Created new spreadsheet. Copy this ID into SPREADSHEET_ID: ' + ss.getId());
  } else {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  var sheet = ss.getSheetByName('Outreach Status');
  if (!sheet) {
    sheet = ss.insertSheet('Outreach Status');
    sheet.appendRow(['Timestamp', 'Company', 'Email', 'Category',
                     'Priority', 'Subject Sent', 'Email Body', 'Status']);
    sheet.getRange(1, 1, 1, 8)
         .setFontWeight('bold')
         .setBackground('#1a73e8')
         .setFontColor('#ffffff');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 230);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 300);
    sheet.setColumnWidth(7, 420);
    sheet.setColumnWidth(8, 90);
    Logger.log('Outreach Status sheet created.');
  }
  return sheet;
}

function getSentEmails(sheet) {
  var data = sheet.getDataRange().getValues();
  var sent = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'SENT') {
      sent.push(data[i][2]);
    }
  }
  return sent;
}
