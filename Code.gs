/**
 * ==============================================================================
 * Project: Contractor Invoice Importer & Monthly Email Drafter
 * Version: v1.1.1
 * Environment: Google Sheets Apps Script (Single-File Architecture)
 * Scope:
 *   1. Imports Excel invoices, checks global duplicates, matches visible
 *      contractor tabs, and appends strictly to Columns A:F.
 *   2. Scans visible contractor tabs (Email in Q1, Month in K16, Breakdown in
 *      H18:K) and creates formatted HTML Gmail drafts for manual review.
 *   3. Excludes all hidden tabs from both workflows.
 * ==============================================================================
 */

/**
 * Adds custom menu to Google Sheets upon workbook launch.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Invoices')
    .addItem('Import Invoices from Excel...', 'showImportDialog')
    .addSeparator()
    .addItem('Draft Monthly Contractor Emails...', 'showEmailDraftDialog')
    .addToUi();
}

/* ============================================================================
 * FEATURE 1: EXCEL INVOICE IMPORTER
 * ============================================================================ */

/**
 * Opens the file upload modal dialog directly from embedded HTML.
 */
function showImportDialog() {
  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 15px; color: #333; font-size: 13px; }
    .field-row { display: flex; gap: 15px; margin-bottom: 15px; align-items: center; }
    label { font-weight: 600; margin-bottom: 4px; display: block; }
    input[type="date"], input[type="file"] { padding: 6px; border: 1px solid #ccc; border-radius: 4px; }
    button { padding: 8px 16px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 500; }
    button:disabled { background-color: #ccc; cursor: not-allowed; }
    .summary-box { display: flex; gap: 10px; margin: 15px 0; }
    .card { flex: 1; padding: 10px; border-radius: 4px; background: #f8f9fa; border-left: 4px solid #ccc; }
    .card.green { border-left-color: #34a853; }
    .card.yellow { border-left-color: #fbbc05; }
    .card.red { border-left-color: #ea4335; }
    .card-val { font-size: 18px; font-weight: bold; }
    .table-container { max-height: 240px; overflow-y: auto; border: 1px solid #e0e0e0; margin-top: 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th, td { border: 1px solid #eee; padding: 6px 8px; text-align: left; }
    th { background-color: #f1f3f4; position: sticky; top: 0; }
    .notice { padding: 8px; margin-bottom: 10px; border-radius: 4px; font-size: 12px; }
    .warning { background: #fff3cd; color: #856404; }
    .version-tag { font-size: 11px; color: #666; margin-left: 6px; }
  </style>
</head>
<body>
  <div class="field-row">
    <div>
      <label>1. Select Excel File (.xlsx): <span class="version-tag">v1.1.1</span></label>
      <input type="file" id="excelFile" accept=".xlsx, .xls" />
    </div>
    <div>
      <label>Start Date (Optional):</label>
      <input type="date" id="startDate" />
    </div>
    <div>
      <label>End Date (Optional):</label>
      <input type="date" id="endDate" />
    </div>
  </div>

  <div id="statusNotice" class="notice warning" style="display:none;"></div>

  <div class="summary-box">
    <div class="card green">
      <div>Ready to Insert</div>
      <div id="readyCount" class="card-val">0</div>
    </div>
    <div class="card yellow">
      <div>Duplicates Skipped</div>
      <div id="dupCount" class="card-val">0</div>
    </div>
    <div class="card red">
      <div>No Tab Found</div>
      <div id="unmatchedCount" class="card-val">0</div>
    </div>
  </div>

  <div class="table-container">
    <table id="previewTable">
      <thead>
        <tr>
          <th>Target Tab</th>
          <th>Invoice#</th>
          <th>Staff</th>
          <th>Date (MM/DD/YYYY)</th>
          <th>Paid</th>
          <th>Total</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        <tr><td colspan="7" style="text-align:center; color:#888;">Select an Excel file to preview data.</td></tr>
      </tbody>
    </table>
  </div>

  <div style="margin-top: 15px; display:flex; justify-content: flex-end; gap: 10px;">
    <button id="confirmBtn" onclick="confirmAndPush()" disabled>🚀 Confirm & Push to Google Sheets</button>
  </div>

  <script>
    let sheetMetadata = null;
    let stagedRowsByTab = {};
    let totalReadyRows = 0;

    google.script.run
      .withSuccessHandler(data => { sheetMetadata = data; })
      .getSheetMetadata();

    document.getElementById('excelFile').addEventListener('change', handleFile);
    document.getElementById('startDate').addEventListener('change', handleFile);
    document.getElementById('endDate').addEventListener('change', handleFile);

    function normalizeTokens(str) {
      return (String(str || '').toLowerCase().match(/\\w+/g) || []);
    }

    function matchContractorToTab(staffName, availableTabs) {
      if (!staffName) return null;
      const cleanStaff = String(staffName).trim().toLowerCase();
      const staffTokens = normalizeTokens(cleanStaff);

      for (const tab of availableTabs) {
        if (tab.trim().toLowerCase() === cleanStaff) return tab;
      }
      if (staffTokens.length > 0) {
        const firstName = staffTokens[0];
        for (const tab of availableTabs) {
          if (tab.trim().toLowerCase() === firstName) return tab;
        }
      }
      for (const tab of availableTabs) {
        const tabLower = tab.trim().toLowerCase();
        if (tabLower.includes(cleanStaff) || cleanStaff.includes(tabLower)) return tab;
        const tabTokens = normalizeTokens(tab);
        const overlap = staffTokens.filter(t => tabTokens.includes(t));
        if (overlap.length > 0) return tab;
      }
      return null;
    }

    function handleFile() {
      const fileInput = document.getElementById('excelFile');
      if (!fileInput.files.length || !sheetMetadata) return;

      const file = fileInput.files[0];
      const reader = new FileReader();

      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const rawJson = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
        processRecords(rawJson);
      };
      reader.readAsArrayBuffer(file);
    }

    function parseDateVal(val) {
      if (!val) return null;
      if (val instanceof Date) return val;
      const d = new Date(val);
      return isNaN(d) ? null : d;
    }

    function formatDateMMDDYYYY(dateObj) {
      if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
      const mm = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(dateObj.getUTCDate()).padStart(2, '0');
      const yyyy = dateObj.getUTCFullYear();
      return mm + '/' + dd + '/' + yyyy;
    }

    function getVal(row, possibleKeys) {
      for (const k of possibleKeys) {
        if (row[k] !== undefined) return row[k];
      }
      const rowKeys = Object.keys(row);
      for (const target of possibleKeys) {
        const targetClean = target.toLowerCase().replace(/\\s+/g, '');
        for (const rk of rowKeys) {
          if (rk.toLowerCase().replace(/\\s+/g, '') === targetClean) {
            return row[rk];
          }
        }
      }
      return '';
    }

    function processRecords(records) {
      const startStr = document.getElementById('startDate').value;
      const endStr = document.getElementById('endDate').value;
      const startFilter = startStr ? new Date(startStr + 'T00:00:00') : null;
      const endFilter = endStr ? new Date(endStr + 'T23:59:59') : null;

      stagedRowsByTab = {};
      totalReadyRows = 0;
      let dupCount = 0;
      let unmatchedCount = 0;
      let unmatchedNames = new Set();
      const previewRows = [];

      records.forEach(row => {
        const rawInv = getVal(row, ['Invoice#', 'Invoice #', 'Invoice Number', 'Inv #']);
        const invNo = String(rawInv || '').replace(/\\.0$/, '').trim();
        const staff = String(getVal(row, ['Staff', 'Contractor', 'Name'])).trim();
        const rawDate = getVal(row, ['Invoice Date', 'Date']);
        const dateObj = parseDateVal(rawDate);

        if (dateObj) {
          if (startFilter && dateObj < startFilter) return;
          if (endFilter && dateObj > endFilter) return;
        }

        if (sheetMetadata.existingInvoices[invNo]) {
          dupCount++;
          return;
        }

        const matchedTab = matchContractorToTab(staff, sheetMetadata.tabNames);
        if (!matchedTab) {
          unmatchedCount++;
          unmatchedNames.add(staff);
          return;
        }

        let formattedDate = '';
        if (dateObj) {
          formattedDate = formatDateMMDDYYYY(dateObj);
        } else if (rawDate) {
          formattedDate = String(rawDate);
        }

        const paid = getVal(row, ['Paid']);
        const total = getVal(row, ['Total', 'Amount']);
        const status = getVal(row, ['Status']);

        const rowValues = [invNo, staff, formattedDate, paid, total, status];

        if (!stagedRowsByTab[matchedTab]) {
          stagedRowsByTab[matchedTab] = [];
        }
        stagedRowsByTab[matchedTab].push(rowValues);
        totalReadyRows++;

        previewRows.push({ tab: matchedTab, values: rowValues });
      });

      document.getElementById('readyCount').innerText = totalReadyRows;
      document.getElementById('dupCount').innerText = dupCount;
      document.getElementById('unmatchedCount').innerText = unmatchedCount;

      const noticeBox = document.getElementById('statusNotice');
      if (unmatchedNames.size > 0) {
        noticeBox.innerText = '⚠️ Skipped ' + unmatchedCount + ' rows: No matching visible tab found for: ' + Array.from(unmatchedNames).join(', ');
        noticeBox.style.display = 'block';
      } else {
        noticeBox.style.display = 'none';
      }

      const tbody = document.querySelector('#previewTable tbody');
      tbody.innerHTML = '';
      if (previewRows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; color:#888;">No new invoices to import matching criteria.</td></tr>';
        document.getElementById('confirmBtn').disabled = true;
      } else {
        previewRows.forEach(item => {
          const tr = document.createElement('tr');
          tr.innerHTML = '<td><strong>' + item.tab + '</strong></td>' + item.values.map(v => '<td>' + v + '</td>').join('');
          tbody.appendChild(tr);
        });
        document.getElementById('confirmBtn').disabled = false;
      }
    }

    function confirmAndPush() {
      const btn = document.getElementById('confirmBtn');
      btn.disabled = true;
      btn.innerText = 'Writing to Google Sheets...';

      google.script.run
        .withSuccessHandler(summary => {
          let msg = 'Successfully inserted ' + totalReadyRows + ' invoices!\\n\\n';
          for (const tab in summary) {
            msg += '• ' + tab + ': +' + summary[tab] + ' rows\\n';
          }
          alert(msg);
          google.script.host.close();
        })
        .withFailureHandler(err => {
          alert('Error inserting rows: ' + err.message);
          btn.disabled = false;
          btn.innerText = '🚀 Confirm & Push to Google Sheets';
        })
        .appendBatchInvoices(stagedRowsByTab);
    }
  </script>
</body>
</html>
  `;

  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(750)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Contractor Invoices (v1.1.1)');
}

/**
 * Scans columns A to F to find the true first blank row.
 */
function getNextRowColumnsAtoF(sheet) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(1, 1, maxRows, 6).getValues();

  let lastOccupiedRow = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    const isRowOccupied = values[i].some(cell => String(cell).trim() !== '');
    if (isRowOccupied) {
      lastOccupiedRow = i + 1;
      break;
    }
  }

  return Math.max(lastOccupiedRow + 1, 2);
}

/**
 * Scans columns A to F across all visible tabs to build a set of existing invoice numbers.
 * Hidden tabs are excluded.
 */
function getSheetMetadata() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const tabNames = [];
  const existingInvoices = {};

  sheets.forEach(sheet => {
    if (sheet.isSheetHidden()) return;

    const title = sheet.getName();
    tabNames.push(title);

    const maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      const sampleRows = Math.min(maxRows, 500);
      const data = sheet.getRange(1, 1, sampleRows, 6).getValues();
      
      const headerRow = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
      let invIdx = headerRow.indexOf('invoice#');
      if (invIdx === -1) {
        invIdx = headerRow.indexOf('invoicenumber');
      }

      if (invIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const invVal = String(data[i][invIdx]).replace(/\.0$/, '').trim();
          if (invVal) {
            existingInvoices[invVal] = true;
          }
        }
      }
    }
  });

  return {
    tabNames: tabNames,
    existingInvoices: existingInvoices
  };
}

/**
 * Batch-appends staged rows strictly to columns A through F at the bottom of each tab.
 */
function appendBatchInvoices(rowsByTab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = {};

  for (const tabName in rowsByTab) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.isSheetHidden()) continue;

    const rows = rowsByTab[tabName];
    if (rows && rows.length > 0) {
      const startRow = getNextRowColumnsAtoF(sheet);
      const numRows = rows.length;
      const numCols = 6;

      const requiredRows = startRow + numRows - 1;
      if (sheet.getMaxRows() < requiredRows) {
        sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
      }

      sheet.getRange(startRow, 1, numRows, numCols).setValues(rows);
      summary[tabName] = numRows;
    }
  }

  return summary;
}


/* ============================================================================
 * FEATURE 2: MONTHLY PAY & BREAKDOWN EMAIL DRAFTER
 * ============================================================================ */

/**
 * Opens the Monthly Contractor Email Drafter preview dialog.
 */
function showEmailDraftDialog() {
  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 15px; color: #333; font-size: 13px; }
    .header-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
    .header-bar h3 { margin: 0; font-size: 16px; }
    .sub-note { font-size: 12px; color: #555; margin-bottom: 12px; }
    button { padding: 8px 16px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 500; }
    button:disabled { background-color: #ccc; cursor: not-allowed; }
    .table-container { max-height: 360px; overflow-y: auto; border: 1px solid #e0e0e0; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th, td { border: 1px solid #eee; padding: 7px 8px; text-align: left; vertical-align: top; }
    th { background-color: #f1f3f4; position: sticky; top: 0; z-index: 1; }
    .badge { display: inline-block; padding: 2px 6px; border-radius: 3px; font-size: 11px; font-weight: 600; }
    .badge.ok { background: #e6f4ea; color: #137333; }
    .badge.skip { background: #fce8e6; color: #c5221f; }
    .preview-mini { font-size: 11px; color: #555; margin-top: 4px; max-height: 75px; overflow-y: auto; background: #fafafa; padding: 4px; border: 1px solid #eee; }
  </style>
</head>
<body>
  <div class="header-bar">
    <h3>Draft Monthly Contractor Payment Emails</h3>
  </div>
  <div class="sub-note">
    Scans visible tabs only (hidden tabs excluded). Reads <strong>Email</strong> from <code>Q1</code>, <strong>Month</strong> from <code>K16</code>, and <strong>Breakdown rows</strong> from <code>H18:K</code>. Drafts are saved to your Gmail <strong>Drafts</strong> folder.
  </div>

  <div class="table-container">
    <table id="draftTable">
      <thead>
        <tr>
          <th style="width:28px;"><input type="checkbox" id="selectAll" checked onchange="toggleAll(this)" /></th>
          <th>Contractor Tab</th>
          <th>Email (Q1)</th>
          <th>Subject (K16)</th>
          <th>Rows (H18:K)</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody id="draftBody">
        <tr><td colspan="6" style="text-align:center; color:#888;">Scanning visible contractor tabs...</td></tr>
      </tbody>
    </table>
  </div>

  <div style="margin-top: 15px; display:flex; justify-content: space-between; align-items: center;">
    <span id="selectionSummary" style="font-size:12px; color:#555;"></span>
    <button id="createDraftsBtn" onclick="createSelectedDrafts()" disabled>✉️ Create Selected Gmail Drafts</button>
  </div>

  <script>
    let contractorPayloads = [];

    google.script.run
      .withSuccessHandler(renderPreview)
      .withFailureHandler(err => {
        document.getElementById('draftBody').innerHTML = '<tr><td colspan="6" style="color:red;">Error: ' + err.message + '</td></tr>';
      })
      .getContractorEmailPayloads();

    function renderPreview(items) {
      contractorPayloads = items || [];
      const tbody = document.getElementById('draftBody');
      tbody.innerHTML = '';

      if (contractorPayloads.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#888;">No visible tabs found in this workbook.</td></tr>';
        return;
      }

      contractorPayloads.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const checkedAttr = item.ready ? 'checked' : 'disabled';
        const statusBadge = item.ready
          ? '<span class="badge ok">Ready</span>'
          : '<span class="badge skip">' + item.skipReason + '</span>';

        let rowsPreviewHtml = item.rows.length + ' row(s)';
        if (item.rows.length > 0) {
          rowsPreviewHtml += '<div class="preview-mini">' +
            item.rows.map(r => r.join(' | ')).join('<br>') +
            '</div>';
        }

        tr.innerHTML =
          '<td><input type="checkbox" class="row-chk" data-idx="' + idx + '" ' + checkedAttr + ' onchange="updateSelectionCount()" /></td>' +
          '<td><strong>' + escapeHtml(item.tabName) + '</strong></td>' +
          '<td>' + escapeHtml(item.email || '—') + '</td>' +
          '<td>' + escapeHtml(item.subject || '—') + '</td>' +
          '<td>' + rowsPreviewHtml + '</td>' +
          '<td>' + statusBadge + '</td>';
        tbody.appendChild(tr);
      });

      updateSelectionCount();
    }

    function escapeHtml(str) {
      return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    function toggleAll(master) {
      document.querySelectorAll('.row-chk:not(:disabled)').forEach(cb => {
        cb.checked = master.checked;
      });
      updateSelectionCount();
    }

    function updateSelectionCount() {
      const selected = document.querySelectorAll('.row-chk:checked').length;
      document.getElementById('selectionSummary').innerText = selected + ' draft(s) selected';
      document.getElementById('createDraftsBtn').disabled = (selected === 0);
    }

    function createSelectedDrafts() {
      const selectedIndices = Array.from(document.querySelectorAll('.row-chk:checked'))
        .map(cb => parseInt(cb.getAttribute('data-idx'), 10));

      const toCreate = selectedIndices.map(i => contractorPayloads[i]);
      if (toCreate.length === 0) return;

      const btn = document.getElementById('createDraftsBtn');
      btn.disabled = true;
      btn.innerText = 'Creating Gmail Drafts...';

      google.script.run
        .withSuccessHandler(res => {
          alert('✅ Created ' + res.createdCount + ' Gmail draft(s)!\\n\\nCheck your Gmail Drafts folder to review and send.');
          google.script.host.close();
        })
        .withFailureHandler(err => {
          alert('Error creating drafts: ' + err.message);
          btn.disabled = false;
          btn.innerText = '✉️ Create Selected Gmail Drafts';
        })
        .createContractorGmailDrafts(toCreate);
    }
  </script>
</body>
</html>
  `;

  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(820)
    .setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(html, 'Draft Monthly Contractor Emails (v1.1.1)');
}

/**
 * Formats individual cell values from H18:K cleanly for display in emails.
 */
function formatBreakdownCell(val, displayVal) {
  if (val === null || val === undefined || String(val).trim() === '') {
    return '';
  }
  if (displayVal !== undefined && displayVal !== null && String(displayVal).trim() !== '') {
    return String(displayVal).trim();
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    const dd = String(val.getDate()).padStart(2, '0');
    const yyyy = val.getFullYear();
    return mm + '/' + dd + '/' + yyyy;
  }
  return String(val).trim();
}

/**
 * Scans every visible worksheet tab for Email (Q1), Month (K16), and Pay Breakdown rows (H18:K).
 * Hidden tabs are excluded.
 */
function getContractorEmailPayloads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const results = [];

  sheets.forEach(sheet => {
    if (sheet.isSheetHidden()) return;

    const tabName = sheet.getName();
    const email = String(sheet.getRange('Q1').getDisplayValue() || '').trim();
    const monthText = String(sheet.getRange('K16').getDisplayValue() || '').trim();

    const maxRows = sheet.getMaxRows();
    const rows = [];

    if (maxRows >= 18) {
      const numRows = maxRows - 17;
      const rawValues = sheet.getRange(18, 8, numRows, 4).getValues();
      const displayValues = sheet.getRange(18, 8, numRows, 4).getDisplayValues();

      for (let r = 0; r < rawValues.length; r++) {
        const formattedRow = [
          formatBreakdownCell(rawValues[r][0], displayValues[r][0]),
          formatBreakdownCell(rawValues[r][1], displayValues[r][1]),
          formatBreakdownCell(rawValues[r][2], displayValues[r][2]),
          formatBreakdownCell(rawValues[r][3], displayValues[r][3])
        ];
        const hasContent = formattedRow.some(cell => cell !== '');
        if (hasContent) {
          rows.push(formattedRow);
        }
      }
    }

    let ready = true;
    let skipReason = '';
    if (!email || email.indexOf('@') === -1) {
      ready = false;
      skipReason = 'Missing valid email in Q1';
    } else if (!monthText) {
      ready = false;
      skipReason = 'Missing month in K16';
    } else if (rows.length === 0) {
      ready = false;
      skipReason = 'No entries in H18:K';
    }

    const subject = monthText ? (monthText + ' Payment Invoice') : '';

    results.push({
      tabName: tabName,
      email: email,
      monthText: monthText,
      subject: subject,
      rows: rows,
      ready: ready,
      skipReason: skipReason
    });
  });

  return results;
}

/**
 * Creates Gmail drafts with HTML table formatting for each selected contractor.
 */
function createContractorGmailDrafts(payloads) {
  let createdCount = 0;

  payloads.forEach(item => {
    if (!item.email || !item.monthText || !item.rows || item.rows.length === 0) {
      return;
    }

    const subject = item.monthText + ' Payment Invoice';
    const plainIntro = 'Below are invoices for ' + item.monthText + ' payment.';

    const plainLines = [plainIntro, ''];
    item.rows.forEach(r => {
      plainLines.push(r.join('\t'));
    });
    const plainBody = plainLines.join('\n');

    let tableRowsHtml = '';
    item.rows.forEach((r, idx) => {
      const isHeader = (idx === 0);
      const cellTag = isHeader ? 'th' : 'td';
      const bgStyle = isHeader
        ? 'background-color:#f1f3f4; font-weight:bold;'
        : (idx % 2 === 0 ? 'background-color:#fafafa;' : 'background-color:#ffffff;');

      const cellsHtml = r.map(cell => {
        const escaped = String(cell || '')
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');
        return '<' + cellTag + ' style="border:1px solid #dcdcdc; padding:6px 10px; text-align:left; ' + bgStyle + '">' +
               escaped +
               '</' + cellTag + '>';
      }).join('');

      tableRowsHtml += '<tr>' + cellsHtml + '</tr>';
    });

    const htmlBody =
      '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #222222;">' +
        '<p>Below are invoices for ' +
          String(item.monthText).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') +
          ' payment.</p>' +
        '<table style="border-collapse: collapse; font-size: 13px; margin-top: 10px;">' +
          tableRowsHtml +
        '</table>' +
      '</div>';

    GmailApp.createDraft(item.email, subject, plainBody, {
      htmlBody: htmlBody
    });

    createdCount++;
  });

  return { createdCount: createdCount };
}