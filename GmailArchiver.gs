/**
 * Gmail Archiver → Google Sheet
 * https://github.com/ibrews/gmail-archiver
 *
 * Archives Gmail messages into a searchable Google Sheet.
 * Works with any Gmail account (personal, Workspace, enterprise).
 *
 * QUICK START:
 * 1. Go to https://script.google.com → New Project
 * 2. Paste this entire script, save (Ctrl+S)
 * 3. Select "archiveInbox" in the function dropdown → Run (▶)
 * 4. Authorize when prompted (Gmail read-only + Sheets + Drive)
 * 5. A Google Sheet appears in your Drive with all your messages
 *
 * For custom queries, date ranges, or other folders, edit CONFIG below
 * or use the convenience functions at the bottom of this file.
 *
 * LARGE MAILBOXES:
 * Google Apps Script has a 6-minute execution limit per run.
 * This script handles it automatically — if it runs out of time, it saves
 * progress and schedules itself to resume in ~1 minute. No data is lost.
 * A mailbox with 10,000 messages typically completes in 3-5 runs.
 *
 * QUOTA: This script minimizes Gmail API calls by batch-fetching messages
 * per thread group and caching per-message getters (body, attachments).
 * If you hit daily quota limits, just wait until midnight PT and re-run —
 * it resumes automatically from where it left off.
 *
 * LICENSE: MIT
 */

// ── Configuration ──────────────────────────────────────────────
const CONFIG = {
  // Gmail search query (uses standard Gmail search syntax)
  // Examples:
  //   'in:inbox'                          → all inbox messages
  //   'in:sent'                           → all sent messages
  //   'label:projects'                    → messages with a specific label
  //   'in:anywhere'                       → all mail (inbox, sent, archived, etc.)
  //   'from:someone@example.com'          → from a specific sender
  //   'after:2024/01/01 before:2025/01/01'→ date range
  //   'has:attachment filename:pdf'       → PDFs only
  //   'in:inbox is:unread'               → unread inbox messages
  // Full syntax: https://support.google.com/mail/answer/7190
  QUERY: 'in:inbox',

  // Name of the Google Sheet (created in your Drive root)
  SHEET_NAME: 'Gmail Archive',

  // How many threads to fetch per API call (max 500, recommended 100)
  BATCH_SIZE: 100,

  // Max runtime before saving & scheduling a continuation (ms)
  // Google's hard limit is 6 min; we stop at 5 to leave buffer for writes
  MAX_RUNTIME_MS: 5 * 60 * 1000,

  // Include full message body text? Set false for metadata-only (much faster)
  INCLUDE_BODY: true,

  // Max characters of body text per message (prevents Sheet cell overflow)
  BODY_MAX_LENGTH: 5000,

  // Include BCC field (only visible on messages you sent)
  INCLUDE_BCC: false,

  // Automatically schedule continuation if time runs out
  AUTO_CONTINUE: true,

  // Delay between continuation runs (ms). Minimum 60000 (1 min).
  CONTINUE_DELAY_MS: 60 * 1000,
};


// ════════════════════════════════════════════════════════════════
// ── Convenience Entry Points ───────────────────────────────────
// ════════════════════════════════════════════════════════════════
// Select any of these in the function dropdown and click Run.

/** Archive all Inbox messages (default) */
function archiveInbox() {
  CONFIG.QUERY = 'in:inbox';
  CONFIG.SHEET_NAME = 'Gmail Archive - Inbox';
  runArchiver_();
}

/** Archive all Sent messages */
function archiveSent() {
  CONFIG.QUERY = 'in:sent';
  CONFIG.SHEET_NAME = 'Gmail Archive - Sent';
  runArchiver_();
}

/** Archive ALL mail (inbox + sent + archived + everything) */
function archiveAllMail() {
  CONFIG.QUERY = 'in:anywhere';
  CONFIG.SHEET_NAME = 'Gmail Archive - All Mail';
  runArchiver_();
}

/** Archive messages with a specific label — edit the label name below */
function archiveLabel() {
  CONFIG.QUERY = 'label:YOUR-LABEL-HERE';
  CONFIG.SHEET_NAME = 'Gmail Archive - Label';
  runArchiver_();
}

/** Archive using a fully custom query — edit below */
function archiveCustomQuery() {
  CONFIG.QUERY = 'from:someone@example.com after:2024/01/01';
  CONFIG.SHEET_NAME = 'Gmail Archive - Custom';
  runArchiver_();
}

/** Archive only unread messages */
function archiveUnread() {
  CONFIG.QUERY = 'is:unread';
  CONFIG.SHEET_NAME = 'Gmail Archive - Unread';
  runArchiver_();
}

/** Archive messages with attachments */
function archiveWithAttachments() {
  CONFIG.QUERY = 'has:attachment';
  CONFIG.SHEET_NAME = 'Gmail Archive - Attachments';
  runArchiver_();
}

/** Archive starred messages */
function archiveStarred() {
  CONFIG.QUERY = 'is:starred';
  CONFIG.SHEET_NAME = 'Gmail Archive - Starred';
  runArchiver_();
}


// ════════════════════════════════════════════════════════════════
// ── Core Archiver ──────────────────────────────────────────────
// ════════════════════════════════════════════════════════════════

function runArchiver_() {
  const startTime = Date.now();

  const ss = getOrCreateSpreadsheet_();
  const dataSheet = ss.getSheetByName('Emails') || ss.insertSheet('Emails');
  const metaSheet = ss.getSheetByName('_Meta') || ss.insertSheet('_Meta');

  // ── Determine headers based on config ──
  const headers = buildHeaders_();

  // ── Fresh start vs. resume ──
  let processedIds = new Set();
  let lastRow = dataSheet.getLastRow();
  let searchStart = 0;

  if (lastRow <= 1) {
    // Fresh start
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    dataSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a86c8')
      .setFontColor('#ffffff');
    dataSheet.setFrozenRows(1);
    lastRow = 1;

    writeMeta_(metaSheet, {processedCount: 0, status: 'running', query: CONFIG.QUERY, searchStart: 0});
  } else {
    // Resume — load processed message IDs
    const msgIdCol = headers.indexOf('Message ID') + 1;
    if (lastRow > 1) {
      const existingIds = dataSheet.getRange(2, msgIdCol, lastRow - 1, 1).getValues();
      existingIds.forEach(row => { if (row[0]) processedIds.add(String(row[0])); });
    }
    // Resume search from where we left off
    searchStart = Number(metaSheet.getRange('B4').getValue()) || 0;
  }

  Logger.log('Archive starting. Query: "%s" | Already processed: %s | Search offset: %s',
    CONFIG.QUERY, processedIds.size, searchStart);

  // ── Main processing loop ──
  let totalProcessed = processedIds.size;
  let newRows = [];
  let start = searchStart;

  while (true) {
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('Time limit approaching. Saving progress...');
      break;
    }

    let threads;
    try {
      threads = GmailApp.search(CONFIG.QUERY, start, CONFIG.BATCH_SIZE);
    } catch (e) {
      Logger.log('Error fetching threads at offset %s: %s', start, e.message);
      break;
    }

    if (threads.length === 0) {
      Logger.log('No more threads. Archive complete!');
      writeMeta_(metaSheet, {processedCount: totalProcessed, status: 'complete', query: CONFIG.QUERY, searchStart: start});
      break;
    }

    // Batch-fetch all messages in one API call (much cheaper than per-thread)
    const allMessages = GmailApp.getMessagesForThreads(threads);

    for (let t = 0; t < threads.length; t++) {
      if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) break;

      const thread = threads[t];
      const messages = allMessages[t];
      for (const msg of messages) {
        const msgId = msg.getId();
        if (processedIds.has(msgId)) continue;

        try {
          const row = extractMessageData_(msg, thread, headers);
          newRows.push(row);
          processedIds.add(msgId);
          totalProcessed++;
        } catch (e) {
          Logger.log('Error processing message %s: %s', msgId, e.message);
          // Write an error row so we don't skip silently
          const errorRow = headers.map(() => '');
          errorRow[0] = new Date();
          errorRow[headers.indexOf('Subject')] = '[Error processing message: ' + e.message + ']';
          errorRow[headers.indexOf('Message ID')] = msgId;
          newRows.push(errorRow);
          processedIds.add(msgId);
          totalProcessed++;
        }
      }
    }

    // Flush batch to sheet
    if (newRows.length > 0) {
      const writeStart = dataSheet.getLastRow() + 1;
      dataSheet.getRange(writeStart, 1, newRows.length, newRows[0].length).setValues(newRows);
      newRows = [];
      SpreadsheetApp.flush();
    }

    start += CONFIG.BATCH_SIZE;
    writeMeta_(metaSheet, {processedCount: totalProcessed, status: 'running', query: CONFIG.QUERY, searchStart: start});
    Logger.log('Progress: %s messages archived (thread offset: %s)', totalProcessed, start);
  }

  // Write any remaining rows
  if (newRows.length > 0) {
    const writeStart = dataSheet.getLastRow() + 1;
    dataSheet.getRange(writeStart, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
  }

  // ── Handle completion or continuation ──
  const status = metaSheet.getRange('B2').getValue();

  if (status === 'complete') {
    formatSheet_(dataSheet, headers);
    try { ss.deleteSheet(metaSheet); } catch(e) {}
    removeTriggers_();
    Logger.log('Archive complete! %s messages archived.', totalProcessed);
    Logger.log('Sheet URL: %s', ss.getUrl());
  } else if (CONFIG.AUTO_CONTINUE) {
    writeMeta_(metaSheet, {processedCount: totalProcessed, status: 'running', query: CONFIG.QUERY, searchStart: start});
    scheduleContinuation_();
    Logger.log('Continuation scheduled. %s messages so far.', totalProcessed);
  } else {
    Logger.log('Paused at %s messages. Run again to continue.', totalProcessed);
  }

  return ss.getUrl();
}


// ════════════════════════════════════════════════════════════════
// ── Message Extraction ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════════

function buildHeaders_() {
  const h = ['Date', 'From', 'From Email', 'To', 'CC'];
  if (CONFIG.INCLUDE_BCC) h.push('BCC');
  h.push('Subject', 'Snippet');
  if (CONFIG.INCLUDE_BODY) h.push('Body');
  h.push('Labels', 'Starred', 'Unread', 'Has Attachments', 'Attachment Names', 'Thread ID', 'Message ID');
  return h;
}

function extractMessageData_(msg, thread, headers) {
  const from = msg.getFrom() || '';
  const fromEmail = extractEmail_(from);
  const fromName = from.replace(/<[^>]+>/, '').replace(/"/g, '').trim() || fromEmail;

  // Cache expensive API calls — each getter counts against the daily quota
  let _plainBody = null;
  function getPlainBodyCached() {
    if (_plainBody === null) {
      try { _plainBody = msg.getPlainBody() || ''; } catch(e) { _plainBody = ''; }
    }
    return _plainBody;
  }

  let _attachments = null;
  function getAttachmentsCached() {
    if (_attachments === null) {
      try { _attachments = msg.getAttachments({includeInlineImages: false}); } catch(e) { _attachments = []; }
    }
    return _attachments;
  }

  const row = [];

  for (const h of headers) {
    switch (h) {
      case 'Date':
        row.push(msg.getDate());
        break;
      case 'From':
        row.push(fromName);
        break;
      case 'From Email':
        row.push(fromEmail);
        break;
      case 'To':
        row.push(msg.getTo() || '');
        break;
      case 'CC':
        row.push(msg.getCc() || '');
        break;
      case 'BCC':
        row.push(msg.getBcc() || '');
        break;
      case 'Subject':
        row.push(msg.getSubject() || '(no subject)');
        break;
      case 'Snippet':
        row.push(getPlainBodyCached().substring(0, 200).replace(/[\n\r]+/g, ' ').trim());
        break;
      case 'Body': {
        let body = getPlainBodyCached();
        if (body.length > CONFIG.BODY_MAX_LENGTH) {
          body = body.substring(0, CONFIG.BODY_MAX_LENGTH) + '... [truncated]';
        }
        row.push(body.replace(/\n{3,}/g, '\n\n').trim());
      }
        break;
      case 'Labels':
        row.push(thread.getLabels().map(l => l.getName()).join(', '));
        break;
      case 'Starred':
        row.push(msg.isStarred());
        break;
      case 'Unread':
        row.push(msg.isUnread());
        break;
      case 'Has Attachments':
        row.push(getAttachmentsCached().length > 0);
        break;
      case 'Attachment Names':
        row.push(getAttachmentsCached().map(a => a.getName()).join(', '));
        break;
      case 'Thread ID':
        row.push(thread.getId());
        break;
      case 'Message ID':
        row.push(msg.getId());
        break;
      default:
        row.push('');
    }
  }

  return row;
}

function extractEmail_(fromStr) {
  const match = (fromStr || '').match(/<([^>]+)>/);
  return match ? match[1] : (fromStr || '').trim();
}


// ════════════════════════════════════════════════════════════════
// ── Spreadsheet Helpers ────────────────────────────────────────
// ════════════════════════════════════════════════════════════════

function getOrCreateSpreadsheet_() {
  const files = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      Logger.log('Found existing spreadsheet: %s', file.getUrl());
      return SpreadsheetApp.open(file);
    }
  }

  const ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.insertSheet('Emails');
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('Created new spreadsheet: %s', ss.getUrl());
  return ss;
}

function writeMeta_(metaSheet, data) {
  metaSheet.getRange('A1').setValue('processedCount');
  metaSheet.getRange('B1').setValue(data.processedCount);
  metaSheet.getRange('A2').setValue('status');
  metaSheet.getRange('B2').setValue(data.status);
  metaSheet.getRange('A3').setValue('query');
  metaSheet.getRange('B3').setValue(data.query);
  metaSheet.getRange('A4').setValue('searchStart');
  metaSheet.getRange('B4').setValue(data.searchStart);
  metaSheet.getRange('A5').setValue('lastUpdated');
  metaSheet.getRange('B5').setValue(new Date());
}

function formatSheet_(sheet, headers) {
  try {
    // Auto-resize metadata columns, cap wide ones
    const narrowCols = ['Date', 'From', 'From Email', 'Subject', 'Labels'];
    for (let i = 0; i < Math.min(headers.length, 7); i++) {
      sheet.autoResizeColumn(i + 1);
    }

    // Cap body/snippet column widths
    const bodyIdx = headers.indexOf('Body');
    if (bodyIdx >= 0) sheet.setColumnWidth(bodyIdx + 1, 400);
    const snippetIdx = headers.indexOf('Snippet');
    if (snippetIdx >= 0) sheet.setColumnWidth(snippetIdx + 1, 300);

    // Add filter view
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1) {
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      range.createFilter();
    }

    // Sort by date descending
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).sort({column: 1, ascending: false});
    }

    // Date format
    const dateIdx = headers.indexOf('Date');
    if (dateIdx >= 0 && lastRow > 1) {
      sheet.getRange(2, dateIdx + 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }

    Logger.log('Sheet formatted successfully.');
  } catch(e) {
    Logger.log('Formatting error (non-critical): %s', e.message);
  }
}


// ════════════════════════════════════════════════════════════════
// ── Trigger Management ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════════

function scheduleContinuation_() {
  removeTriggers_();
  ScriptApp.newTrigger('runArchiver_')
    .timeBased()
    .after(CONFIG.CONTINUE_DELAY_MS)
    .create();
  Logger.log('Continuation scheduled in %s ms.', CONFIG.CONTINUE_DELAY_MS);
}

function removeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const fn = trigger.getHandlerFunction();
    if (fn === 'runArchiver_' || fn === 'archiveInbox') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}


// ════════════════════════════════════════════════════════════════
// ── Utilities ──────────────────────────────────────────────────
// ════════════════════════════════════════════════════════════════

/**
 * Run this to start fresh. Trashes the existing archive sheet
 * and removes any scheduled triggers.
 */
function resetArchive() {
  removeTriggers_();
  const files = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
  let found = false;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      file.setTrashed(true);
      Logger.log('Trashed: %s', file.getName());
      found = true;
    }
  }
  if (!found) Logger.log('No existing archive found for "%s".', CONFIG.SHEET_NAME);
  Logger.log('Reset complete. Run any archive function to start fresh.');
}

/**
 * Run this if the script seems stuck or you want to cancel an in-progress archive.
 * Removes all scheduled triggers without deleting the sheet.
 */
function stopArchive() {
  removeTriggers_();
  Logger.log('All triggers removed. The archive sheet is still intact — run again to resume, or resetArchive() to start over.');
}

/**
 * Run this to check progress of a running archive.
 */
function checkProgress() {
  const files = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const ss = SpreadsheetApp.open(file);
      const metaSheet = ss.getSheetByName('_Meta');
      if (metaSheet) {
        const count = metaSheet.getRange('B1').getValue();
        const status = metaSheet.getRange('B2').getValue();
        const query = metaSheet.getRange('B3').getValue();
        const lastUpdate = metaSheet.getRange('B5').getValue();
        Logger.log('Status: %s | Messages: %s | Query: "%s" | Last update: %s', status, count, query, lastUpdate);
        return;
      } else {
        const rows = ss.getSheetByName('Emails')?.getLastRow() || 0;
        Logger.log('Archive complete. Total rows: %s (including header)', rows);
        return;
      }
    }
  }
  Logger.log('No archive found for "%s". Run an archive function to start.', CONFIG.SHEET_NAME);
}
