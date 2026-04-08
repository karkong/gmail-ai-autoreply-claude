# Gmail AI Auto-Reply Bot
### Free automated cold outreach + AI reply system using Google Apps Script + Claude API

No Zapier. No Make. No monthly subscriptions. Runs entirely inside your Google account.

---

## What It Does

**Cold Outreach Sender**
- Reads a hardcoded buyer list with company-specific context
- Uses Claude AI to write a tailored email for each buyer
- Sends from your Gmail with your signature
- Logs every email to Google Sheets
- Preview mode — review all emails before sending

**Auto-Reply Bot**
- Scans inbox every 5 minutes for new unread emails
- Filters by subject keyword (only your target topic)
- Blocks no-reply senders, auto-acknowledgements, and out-of-office replies
- Passes email to Claude API — generates a professional reply
- Sends reply in the same Gmail thread
- Labels thread "AI-Replied" for your daily review
- Logs everything to Google Sheets

---

## Stack

| Component | Tool | Cost |
|-----------|------|------|
| Email send/receive | Gmail | Free |
| Automation platform | Google Apps Script | Free |
| AI engine | Claude API (Haiku model) | ~$0.001/reply |
| Reply log | Google Sheets | Free |

**Estimated total cost for ~50 replies/month: under $0.10**

---

## Setup (30 minutes)

### Prerequisites
- A Gmail account
- A Claude API key from [console.anthropic.com](https://console.anthropic.com)

### Step 1 — Open Google Apps Script
Go to [script.google.com](https://script.google.com) → **New project**

### Step 2 — Paste the script
Delete the default `myFunction()` content and paste the full contents of `Code.gs` from this repo.

> ⚠️ Common mistake: Do NOT paste inside the existing `function myFunction() {}` — delete that first, then paste.

### Step 3 — Configure your details
Find the `CONFIG` block at the top and fill in:
```javascript
const CONFIG = {
  CLAUDE_API_KEY: 'sk-ant-YOUR-KEY-HERE',
  GMAIL_ADDRESS:  'your@gmail.com',
  ...
};
```

### Step 4 — Edit the buyer list
Find the `BUYERS` array and replace with your own contacts:
```javascript
var BUYERS = [
  { company: 'Company Name',
    contact: 'Purchasing Manager',
    email:   'contact@company.com',
    category: 'distributor',   // distributor | igu | compressed_air (or add your own)
    priority: '★★★',
    notes: 'Specific notes about this company and what angle to use' },
  // ... add more buyers
];
```

### Step 5 — Edit product context
Find `PRODUCT_CONTEXT` and `CATEGORY_CONTEXT` — replace with your own product details, pricing, and location.

### Step 6 — Activate the bot
1. Select `setupTrigger` from the function dropdown → ▶ Run
2. Approve permissions when Google prompts (click Advanced → proceed)
3. Bot is now scanning every 5 minutes automatically

### Step 7 — Send cold outreach (preview first)
1. Confirm `var PREVIEW_ONLY = true;`
2. Select `sendColdOutreach` → ▶ Run
3. Open Google Sheets → **Outreach Status** tab — review all generated emails
4. If happy, change to `var PREVIEW_ONLY = false;`
5. Run `sendColdOutreach` again — emails go out for real

---

## Functions Reference

| Function | What it does | When to run |
|----------|-------------|-------------|
| `setupTrigger()` | Activates 5-minute auto-reply schedule | Once, at setup |
| `sendColdOutreach()` | Sends tailored cold emails to buyer list | Once (preview first) |
| `testBotManually()` | Runs the auto-reply check immediately | Anytime for testing |
| `clearLogs()` | Clears old data from both sheet tabs | When you want a fresh log |
| `pauseBot()` | Stops the 5-minute trigger | When you want to pause |

---

## Filters & Safeguards

The bot has 4 layers of protection against replying to wrong emails:

1. **Gmail search query** — only fetches emails with your subject keyword
2. **Subject auto-reply check** — blocks "automatic reply", "out of office", "auto response" etc.
3. **Sender blocklist** — blocks no-reply, noreply, mailer-daemon, newsletter, welcome@, etc.
4. **AI-Replied label** — already-replied threads are skipped permanently

---

## Daily Routine

1. Open Gmail → check **"AI-Replied"** label (review what bot sent)
2. Open Google Sheets → check **"MOLSIV 4A — AI Reply Log"** tab
3. Step in manually for: price negotiation, purchase orders, inspection bookings

---

## Customising the Claude Prompt

The bot's behaviour is controlled by two text blocks in the script:

**`PRODUCT_CONTEXT`** — injected into every auto-reply prompt. Edit this to change:
- What the product is
- Pricing
- Location
- Rules (what NOT to say)
- How to handle different inquiry types

**`CATEGORY_CONTEXT`** — used in cold outreach. One entry per buyer category. Tells Claude what each type of buyer cares about.

---

## Common Issues

| Problem | Fix |
|---------|-----|
| Functions not in dropdown | Script pasted inside `myFunction()` — delete line 1 and last `}` |
| "Unexpected end of input" | Mismatched braces — check edit was applied cleanly |
| Bot replying to wrong emails | Tighten subject keyword in Gmail search query |
| Duplicate spreadsheets created | Hardcode spreadsheet ID using `SpreadsheetApp.openById()` |
| Status shows PREVIEW, nothing sent | Change `PREVIEW_ONLY = false` and run again |
| Claude API error 401 | API key wrong or expired |

---

## Cost Estimate

- Claude Haiku: ~$0.001 per reply (input + output tokens)
- 50 replies/month = ~$0.05
- 500 replies/month = ~$0.50

Google Apps Script is completely free up to 6 mins execution time/day — more than enough.

---

## License
MIT — free to use, modify, and distribute.

---

## Credits
Built with [Anthropic Claude API](https://console.anthropic.com) + Google Apps Script.
