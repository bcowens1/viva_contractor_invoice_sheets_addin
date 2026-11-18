
  Adds custom menu to Google Sheets.
 
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Invoices')
    .addItem('Import Invoices from Excel...', 'showImportDialog')
    .addToUi();
}


  Opens the file upload modal dialog directly from embedded HTML.
 
function showImportDialog() {
  const htmlContent = `
!DOCTYPE html
html
head
  base target=_top
  script src=httpscdn.jsdelivr.netnpmxlsx@0.18.5distxlsx.full.min.jsscript
  style
    body { font-family -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif; margin 15px; color #333; font-size 13px; }
    .field-row { display flex; gap 15px; margin-bottom 15px; align-items center; }
    label { font-weight 600; margin-bottom 4px; display block; }
    input[type=date], input[type=file] { padding 6px; border 1px solid #ccc; border-radius 4px; }
    button { padding 8px 16px; background-color #1a73e8; color white; border none; border-radius 4px; cursor pointer; font-weight 500; }
    buttondisabled { background-color #ccc; cursor not-allowed; }
    .summary-box { display flex; gap 10px; margin 15px 0; }
    .card { flex 1; padding 10px; border-radius 4px; background #f8f9fa; border-left 4px solid #ccc; }
    .card.green { border-left-color #34a853; }
    .card.yellow { border-left-color #fbbc05; }
    .card.red { border-left-color #ea4335; }
    .card-val { font-size 18px; font-weight bold; }
    .table-container { max-height 240px; overflow-y auto; border 1px solid #e0e0e0; margin-top 10px; }
    table { width 100%; border-collapse collapse; font-size 12px; }
    th, td { border 1px solid #eee; padding 6px 8px; text-align left; }
    th { background-color #f1f3f4; position sticky; top 0; }
    .notice { padding 8px; margin-bottom 10px; border-radius 4px; font-size 12px; }
    .warning { background #fff3cd; color #856404; }
  style
head
body
  div class=field-row
    div
      label1. Select Excel File (.xlsx)label
      input type=file id=excelFile accept=.xlsx, .xls 
    div
    div
      labelStart Date (Optional)label
      input type=date id=startDate 
    div
    div
      labelEnd Date (Optional)label
      input type=date id=endDate 
    div
  div

  div id=statusNotice class=notice warning style=displaynone;div

  div class=summary-box
    div class=card green
      divReady to Insertdiv
      div id=readyCount class=card-val0div
    div
    div class=card yellow
      divDuplicates Skippeddiv
      div id=dupCount class=card-val0div
    div
    div class=card red
      divNo Tab Founddiv
      div id=unmatchedCount class=card-val0div
    div
  div

  div class=table-container
    table id=previewTable
      thead
        tr
          thTarget Tabth
          thInvoice#th
          thStaffth
          thDate (MMDDYYYY)th
          thPaidth
          thTotalth
          thStatusth
        tr
      thead
      tbody
        trtd colspan=7 style=text-aligncenter; color#888;Select an Excel file to preview data.tdtr
      tbody
    table
  div

  div style=margin-top 15px; displayflex; justify-content flex-end; gap 10px;
    button id=confirmBtn onclick=confirmAndPush() disabled🚀 Confirm & Push to Google Sheetsbutton
  div

  script
    let sheetMetadata = null;
    let stagedRowsByTab = {};
    let totalReadyRows = 0;

    google.script.run
      .withSuccessHandler(data = { sheetMetadata = data; })
      .getSheetMetadata();

    document.getElementById('excelFile').addEventListener('change', handleFile);
    document.getElementById('startDate').addEventListener('change', handleFile);
    document.getElementById('endDate').addEventListener('change', handleFile);

    function normalizeTokens(str) {
      return (String(str  '').toLowerCase().match(w+g)  []);
    }

    function matchContractorToTab(staffName, availableTabs) {
      if (!staffName) return null;
      const cleanStaff = String(staffName).trim().toLowerCase();
      const staffTokens = normalizeTokens(cleanStaff);

      for (const tab of availableTabs) {
        if (tab.trim().toLowerCase() === cleanStaff) return tab;
      }
      if (staffTokens.length  0) {
        const firstName = staffTokens[0];
        for (const tab of availableTabs) {
          if (tab.trim().toLowerCase() === firstName) return tab;
        }
      }
      for (const tab of availableTabs) {
        const tabLower = tab.trim().toLowerCase();
        if (tabLower.includes(cleanStaff)  cleanStaff.includes(tabLower)) return tab;
        const tabTokens = normalizeTokens(tab);
        const overlap = staffTokens.filter(t = tabTokens.includes(t));
        if (overlap.length  0) return tab;
      }
      return null;
    }

    function handleFile() {
      const fileInput = document.getElementById('excelFile');
      if (!fileInput.files.length  !sheetMetadata) return;

      const file = fileInput.files[0];
      const reader = new FileReader();

      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type 'array', cellDates true });
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
      return isNaN(d)  null  d;
    }

    function formatDateMMDDYYYY(dateObj) {
      if (!dateObj  !(dateObj instanceof Date)  isNaN(dateObj.getTime())) return '';
      const mm = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(dateObj.getUTCDate()).padStart(2, '0');
      const yyyy = dateObj.getUTCFullYear();
      return mm + '' + dd + '' + yyyy;
    }

    function getVal(row, possibleKeys) {
      for (const k of possibleKeys) {
        if (row[k] !== undefined) return row[k];
      }
      const rowKeys = Object.keys(row);
      for (const target of possibleKeys) {
        const targetClean = target.toLowerCase().replace(s+g, '');
        for (const rk of rowKeys) {
          if (rk.toLowerCase().replace(s+g, '') === targetClean) {
            return row[rk];
          }
        }
      }
      return '';
    }

    function processRecords(records) {
      const startStr = document.getElementById('startDate').value;
      const endStr = document.getElementById('endDate').value;
      const startFilter = startStr  new Date(startStr + 'T000000')  null;
      const endFilter = endStr  new Date(endStr + 'T235959')  null;

      stagedRowsByTab = {};
      totalReadyRows = 0;
      let dupCount = 0;
      let unmatchedCount = 0;
      let unmatchedNames = new Set();
      const previewRows = [];

      records.forEach(row = {
        const rawInv = getVal(row, ['Invoice#', 'Invoice #', 'Invoice Number', 'Inv #']);
        const invNo = String(rawInv  '').replace(.0$, '').trim();
        const staff = String(getVal(row, ['Staff', 'Contractor', 'Name'])).trim();
        const rawDate = getVal(row, ['Invoice Date', 'Date']);
        const dateObj = parseDateVal(rawDate);

        if (dateObj) {
          if (startFilter && dateObj  startFilter) return;
          if (endFilter && dateObj  endFilter) return;
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

        previewRows.push({ tab matchedTab, values rowValues });
      });

      document.getElementById('readyCount').innerText = totalReadyRows;
      document.getElementById('dupCount').innerText = dupCount;
      document.getElementById('unmatchedCount').innerText = unmatchedCount;

      const noticeBox = document.getElementById('statusNotice');
      if (unmatchedNames.size  0) {
        noticeBox.innerText = '⚠️ Skipped ' + unmatchedCount + ' rows No matching tab found for ' + Array.from(unmatchedNames).join(', ');
        noticeBox.style.display = 'block';
      } else {
        noticeBox.style.display = 'none';
      }

      const tbody = document.querySelector('#previewTable tbody');
      tbody.innerHTML = '';
      if (previewRows.length === 0) {
        tbody.innerHTML = 'trtd colspan=7 style=text-aligncenter; color#888;No new invoices to import matching criteria.tdtr';
        document.getElementById('confirmBtn').disabled = true;
      } else {
        previewRows.forEach(item = {
          const tr = document.createElement('tr');
          tr.innerHTML = 'tdstrong' + item.tab + 'strongtd' + item.values.map(v = 'td' + v + 'td').join('');
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
        .withSuccessHandler(summary = {
          let msg = 'Successfully inserted ' + totalReadyRows + ' invoices!nn';
          for (const tab in summary) {
            msg += '• ' + tab + ' +' + summary[tab] + ' rowsn';
          }
          alert(msg);
          google.script.host.close();
        })
        .withFailureHandler(err = {
          alert('Error inserting rows ' + err.message);
          btn.disabled = false;
          btn.innerText = '🚀 Confirm & Push to Google Sheets';
        })
        .appendBatchInvoices(stagedRowsByTab);
    }
  script
body
html
  `;

  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(750)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Contractor Invoices');
}


  Helper to find the next available row based exclusively on columns A to F.
 
function getNextRowColumnsAtoF(sheet) {
  const maxRows = sheet.getMaxRows();
   Read columns A through F for all rows currently existing in the sheet
  const values = sheet.getRange(1, 1, maxRows, 6).getValues();

   Scan backwards from bottom to find the last row containing any data in AF
  let lastOccupiedRow = 0;
  for (let i = values.length - 1; i = 0; i--) {
    const isRowOccupied = values[i].some(cell = String(cell).trim() !== '');
    if (isRowOccupied) {
      lastOccupiedRow = i + 1;
      break;
    }
  }

   Row 1 is header, so minimum insert target is Row 2
  return Math.max(lastOccupiedRow + 1, 2);
}


  Scans columns A to F across all tabs to build a set of existing invoice numbers.
 
function getSheetMetadata() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const tabNames = [];
  const existingInvoices = {};

  sheets.forEach(sheet = {
    const title = sheet.getName();
    tabNames.push(title);

    const maxRows = sheet.getMaxRows();
    if (maxRows  1) {
       Fetch up to row 100 or total rows in cols A to F
      const sampleRows = Math.min(maxRows, 500);
      const data = sheet.getRange(1, 1, sampleRows, 6).getValues();
      
      const headerRow = data[0].map(h = String(h).trim().toLowerCase().replace(s+g, ''));
      let invIdx = headerRow.indexOf('invoice#');
      if (invIdx === -1) {
        invIdx = headerRow.indexOf('invoicenumber');
      }

      if (invIdx !== -1) {
        for (let i = 1; i  data.length; i++) {
          const invVal = String(data[i][invIdx]).replace(.0$, '').trim();
          if (invVal) {
            existingInvoices[invVal] = true;
          }
        }
      }
    }
  });

  return {
    tabNames tabNames,
    existingInvoices existingInvoices
  };
}


  Batch-appends staged rows strictly to columns A through F at the bottom of each tab.
 
function appendBatchInvoices(rowsByTab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = {};

  for (const tabName in rowsByTab) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;

    const rows = rowsByTab[tabName];
    if (rows && rows.length  0) {
       Find the first blank row specifically within columns AF
      const startRow = getNextRowColumnsAtoF(sheet);
      const numRows = rows.length;
      const numCols = 6;  Exactly columns A to F

       Ensure the sheet has enough rows allocated
      const requiredRows = startRow + numRows - 1;
      if (sheet.getMaxRows()  requiredRows) {
        sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
      }

       Write strictly to AF
      sheet.getRange(startRow, 1, numRows, numCols).setValues(rows);
      summary[tabName] = numRows;
    }
  }

  return summary;
}