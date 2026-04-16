# Gmail Archiver

A Google Apps Script that archives your Gmail into a searchable Google Sheet. Works with any Gmail account — personal, Google Workspace, or enterprise.

## Why

Gmail has no built-in way to export your email to a spreadsheet. Google Takeout gives you an MBOX file, which isn't searchable without extra tools (and isn't available on all Workspace accounts). This script creates a Google Sheet you can search, filter, sort, and share — directly in your Drive.

## What You Get

A Google Sheet with one row per message, containing:

| Column | Description |
|---|---|
| Date | Send date/time |
| From | Sender name |
| From Email | Sender email address |
| To | Recipients |
| CC | CC recipients |
| Subject | Email subject line |
| Snippet | First ~200 chars of the message |
| Body | Full plain-text body (optional, configurable) |
| Labels | Gmail labels on the thread |
| Starred | Whether the message is starred |
| Unread | Whether the message is unread || Has Attachments | Whether the message has attachments |
| Attachment Names | Names of attached files |
| Thread ID | Gmail thread ID |
| Message ID | Gmail message ID |

The sheet is auto-formatted with frozen headers, column filters, and date sorting.

## Quick Start

1. Go to [script.google.com](https://script.google.com) (log in with the Gmail account you want to archive)
2. Click **New Project**
3. Delete the placeholder code and paste the contents of [`GmailArchiver.gs`](GmailArchiver.gs)
4. Save (Ctrl+S / Cmd+S)
5. In the function dropdown at the top, select **`archiveInbox`**
6. Click **Run** (▶)
7. Google will ask you to authorize — click "Review Permissions" → choose your account → "Allow"
8. The script runs and creates a sheet called "Gmail Archive - Inbox" in your Drive

That's it. Open the sheet and use Ctrl+F or the column filter dropdowns to search.

## Prebuilt Functions

Select any of these from the function dropdown and run them directly:

| Function | What it archives |
|---|---|
| `archiveInbox` | All Inbox messages |
| `archiveSent` | All Sent messages |
| `archiveAllMail` | Everything (inbox, sent, archived, drafts) |
| `archiveUnread` | Unread messages only |
| `archiveStarred` | Starred messages only |
| `archiveWithAttachments` | Messages that have attachments || `archiveLabel` | Messages with a specific label (edit the label name in the function) |
| `archiveCustomQuery` | Any custom Gmail search query (edit the query in the function) |

## Things to Try

1. **Paste `GmailArchiver.gs` into a new Apps Script project, select `archiveInbox`, and click Run** — after authorizing, a formatted Google Sheet appears with one row per message, sortable columns, and frozen headers.
2. **Select `archiveSent` and run it** — a second pass adds your sent messages; compare the "From" column to see both sides of a thread.
3. **Edit `CONFIG.QUERY` to `from:boss@company.com after:2024/01/01`** and run `archiveCustomQuery` — only matching messages appear; any valid Gmail search syntax works here.
4. **Let it run on a large mailbox (1,000+ messages)** — when the 6-minute Apps Script limit approaches, it saves progress to a `_Meta` sheet and stops cleanly; run it again to resume exactly where it left off.
5. **Add a column filter on "Has Attachments" in the sheet, then filter to TRUE** — instantly see every email with an attachment alongside the file names in the "Attachment Names" column.

## Custom Queries

The `CONFIG.QUERY` field accepts any [Gmail search syntax](https://support.google.com/mail/answer/7190). Some examples:

```
from:boss@company.com                    → emails from a specific person
after:2024/01/01 before:2025/01/01       → date range
has:attachment filename:pdf               → PDFs only
subject:"quarterly report"               → subject line search
from:me to:client@example.com            → emails you sent to someone
is:unread label:important                → unread important messages
larger:5M                                → messages over 5MB
```

Edit the `archiveCustomQuery` function, or modify `CONFIG.QUERY` directly at the top of the script.

## Large Mailboxes

Google Apps Script enforces a **6-minute execution limit** per run. This script handles it automatically:

1. Before hitting the limit, it saves progress to a hidden `_Meta` sheet
2. It schedules a continuation trigger to run ~1 minute later
3. The next run picks up exactly where it left off (no duplicates)
4. When all messages are processed, it cleans up and formats the sheet

A mailbox with 10,000 messages typically completes in 3–5 automatic runs. You don't need to do anything — just let it work.
## Configuration

All options are in the `CONFIG` object at the top of the script:

```javascript
const CONFIG = {
  QUERY: 'in:inbox',          // Gmail search query
  SHEET_NAME: 'Gmail Archive', // Name of the output spreadsheet
  BATCH_SIZE: 100,             // Threads per API call (max 500)
  MAX_RUNTIME_MS: 5 * 60000,  // Stop before 6-min limit
  INCLUDE_BODY: true,          // Include full message body (slower if true)
  BODY_MAX_LENGTH: 5000,       // Truncate body text per message
  INCLUDE_BCC: false,          // Include BCC column
  AUTO_CONTINUE: true,         // Auto-schedule continuation runs
  CONTINUE_DELAY_MS: 60000,    // Delay between continuation runs
};
```

Setting `INCLUDE_BODY: false` significantly speeds up the archive if you only need metadata for searching.

## Utility Functions

| Function | What it does |
|---|---|
| `checkProgress` | Logs the current status and message count of a running archive |
| `stopArchive` | Cancels scheduled continuation triggers without deleting the sheet |
| `resetArchive` | Trashes the existing archive sheet and removes triggers so you can start fresh |

## Resuming After Interruption

If the script is interrupted (lost internet, closed the tab, hit an error), just run it again. It reads the existing sheet, identifies which messages are already archived by Message ID, and continues from where it stopped. No duplicates.
## Permissions

The script requests these Google API scopes:

- **Gmail (read-only)** — to read your messages
- **Google Sheets** — to create and write to the archive spreadsheet
- **Google Drive** — to check if an archive sheet already exists

It does **not** modify, delete, or send any emails. It only reads.

## Limitations

- **6-minute execution limit**: handled automatically via continuation triggers, but very large mailboxes (50k+) will take many runs
- **Google Sheets cell limit**: Sheets has a 10 million cell limit. At ~15 columns per row, that's ~660k messages before hitting the cap. For mailboxes larger than that, consider setting `INCLUDE_BODY: false` or splitting into multiple archives by date range
- **Body text only**: the script extracts plain-text bodies, not HTML. Rich formatting, inline images, and embedded content are not preserved
- **Attachments not downloaded**: attachment names are listed, but the files themselves are not saved. You could extend the script to save them to Drive if needed
- **Workspace admin restrictions**: some Google Workspace admins disable Apps Script or restrict OAuth scopes. If you can't authorize the script, check with your IT admin

## License

MIT — do whatever you want with it.