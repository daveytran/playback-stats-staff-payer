// Configuration
const CONFIG = {
  mainSheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
  workLogSheetId: '17BzCsrHTQQi4e1hg59_AmP9PtiKWdkKCjbAvu5_EjO0',
  workLogSheetName: 'MASTER'
};

// Helper function to get column index by name
function getColumnIndex(sheet, columnName, headerRow = 1) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(columnName);
  if (index === -1) {
    throw new Error(`Column "${columnName}" not found in sheet "${sheet.getName()}"`);
  }
  return index + 1; // Convert to 1-based index for Sheets API
}

// Helper function to get multiple column indices
function getColumnIndices(sheet, columnNames, headerRow = 1) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  
  columnNames.forEach(name => {
    const index = headers.indexOf(name);
    if (index === -1) {
      Logger.log(`Warning: Column "${name}" not found in sheet "${sheet.getName()}"`);
      indices[name] = -1;
    } else {
      indices[name] = index; // 0-based for array access
    }
  });
  
  return indices;
}

// Add menu on open
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Staff Pay Automation')
    .addItem('Calculate Staff Pay', 'calculateStaffPay')
    .addItem('Analyze Sheet Structure', 'analyzeSheets')
    .addItem('Get Sample Data', 'getSampleData')
    .addToUi();
}

// Main function to calculate staff pay
function calculateStaffPay() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Get work log data
    const workLogData = getUnpaidWorkFromMaster();
    
    if (workLogData.length === 0) {
      // Get debug log
      const debugLog = JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]');
      const debugInfo = debugLog.join('\n');
      
      const result = ui.alert(
        'No unpaid work found',
        'All tasks are either already paid or invoiced.\n\nClick "Yes" to see debug log, "No" to exit.',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        showDebugLog();
      }
      return;
    }
    
    // Get pay configuration
    const payConfig = getPayConfiguration();
    const staffMapping = getStaffMapping();
    
    // Calculate payments
    const { payments, errors } = calculatePayments(workLogData, payConfig, staffMapping);
    
    // Check for errors
    let errorMessage = '';
    if (errors.unmatchedTaskTypes.size > 0) {
      errorMessage += `\n⚠️ Unmatched Task Types in Pay Config:\n`;
      errors.unmatchedTaskTypes.forEach(type => {
        errorMessage += `  - ${type}\n`;
      });
    }
    
    if (errors.unmatchedStaffKeys.size > 0) {
      errorMessage += `\n⚠️ Staff Keys not found in mapping:\n`;
      errors.unmatchedStaffKeys.forEach(key => {
        errorMessage += `  - ${key}\n`;
      });
    }
    
    if (errors.tasksWithNoRate.length > 0) {
      errorMessage += `\n⚠️ Tasks with no pay rate:\n`;
      errors.tasksWithNoRate.forEach(task => {
        errorMessage += `  - ${task.staffName}: ${task.taskType} (${task.league} ${task.round})\n`;
      });
    }
    
    // Show summary for confirmation
    const summary = createPaymentSummary(payments);
    const fullMessage = summary + (errorMessage ? '\n\n' + errorMessage : '');
    
    // Show dialog with three options
    const htmlContent = `
      <div style="font-family: Arial, sans-serif;">
        <h3>Payment Summary</h3>
        <pre style="background-color: #f5f5f5; padding: 10px; border-radius: 5px;">${summary}</pre>
        ${errorMessage ? `<div style="color: red; margin-top: 10px;"><pre>${errorMessage}</pre></div>` : ''}
        <div style="margin-top: 20px;">
          <p>Choose an action:</p>
          <button onclick="google.script.run.createInvoicesAndMark()" style="padding: 10px 20px; margin-right: 10px;">Create Invoices</button>
          <button onclick="google.script.run.showDebugLog()" style="padding: 10px 20px; margin-right: 10px;">Show Debug Log</button>
          <button onclick="google.script.host.close()" style="padding: 10px 20px;">Cancel</button>
        </div>
      </div>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(600)
        .setHeight(500);
    
    ui.showModalDialog(htmlOutput, 'Payment Summary & Actions');
    
    // Store work data for later use
    PropertiesService.getScriptProperties().setProperty('pendingWorkData', JSON.stringify(workLogData));
    PropertiesService.getScriptProperties().setProperty('pendingPayments', JSON.stringify(payments));
    
  } catch (error) {
    ui.alert('Error', 'An error occurred: ' + error.toString(), ui.ButtonSet.OK);
    Logger.log(error);
  }
}

// Get unpaid work from MASTER sheet
function getUnpaidWorkFromMaster() {
  const workLogSheet = SpreadsheetApp.openById(CONFIG.workLogSheetId);
  const masterSheet = workLogSheet.getSheetByName(CONFIG.workLogSheetName);
  
  if (!masterSheet) {
    throw new Error('MASTER sheet not found in work log spreadsheet');
  }
  
  // Get column indices by name - MASTER sheet has headers in row 2
  // First, find the actual Status column dynamically
  const actualHeaders = masterSheet.getRange(2, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  const statusColumnName = actualHeaders.find(h => h && h.toString().toLowerCase().includes('status'));
  
  const columnNames = [
    'Assign',
    'Due Date',
    'LEAGUE',
    'Round',
    'Team 1',
    'Team 2',
    'STATS LEVEL',
    'Youtube link',
    'QA',
    'Playback Link',
    statusColumnName, // Use the actual status column name found
    'Team 1 Public Stats Link',
    'Team 2 Public Stats link',
    'Paid',
    'Paid Date',
    'Payment Method',
    'Done Date'
  ];
  
  const cols = getColumnIndices(masterSheet, columnNames, 2); // Headers in row 2
  const data = masterSheet.getDataRange().getValues();
  const unpaidWork = [];
  const debugLog = [];
  
  debugLog.push(`Status column found: "${statusColumnName}"`);
  debugLog.push(`Total rows in sheet: ${data.length}`);
  debugLog.push(`Column indices found: ${JSON.stringify(cols)}`);
  
  // Specific rows to check
  const targetRows = [257, 261, 300, 302, 307];
  
  // Skip header rows (rows 0 and 1 based on your data)
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const status = row[cols[statusColumnName]];
    const paid = row[cols['Paid']];
    const rowNumber = i + 1; // 1-based row number
    
    // Log first few rows for debugging
    if (i < 5) {
      debugLog.push(`Row ${rowNumber}: Status="${status}", Paid="${paid}"`);
    }
    
    // Log specific target rows in detail
    if (targetRows.includes(rowNumber)) {
      debugLog.push(`*** TARGET ROW ${rowNumber} ***`);
      debugLog.push(`  Status: "${status}" (type: ${typeof status}, length: ${status ? status.length : 'null'})`);
      debugLog.push(`  Paid: "${paid}" (type: ${typeof paid}, length: ${paid ? paid.length : 'null'})`);
      debugLog.push(`  Status === 'Done': ${status === 'Done'}`);
      debugLog.push(`  Paid !== 'Paid': ${paid !== 'Paid'}`);
      debugLog.push(`  Paid !== 'Invoiced': ${paid !== 'Invoiced'}`);
      debugLog.push(`  Overall condition: ${status === 'Done' && paid !== 'Paid' && paid !== 'Invoiced'}`);
      debugLog.push(`  Staff: "${row[cols['Assign']]}"`);
      debugLog.push(`  Task: "${row[cols['STATS LEVEL']]}"`);
      debugLog.push(`  League: "${row[cols['LEAGUE']]}"`);
      debugLog.push('');
    }
    
    // Check if work is done but not paid or invoiced
    if (status === 'Done' && paid !== 'Paid' && paid !== 'Invoiced') {
      unpaidWork.push({
        rowIndex: i + 1, // 1-based for Sheets API
        staffName: row[cols['Assign']],
        taskType: row[cols['STATS LEVEL']],
        league: row[cols['LEAGUE']],
        round: row[cols['Round']],
        team1: row[cols['Team 1']],
        team2: row[cols['Team 2']],
        playbackLink: row[cols['Playback Link']],
        doneDate: row[cols['Done Date']]
      });
    }
  }
  
  debugLog.push(`Found ${unpaidWork.length} unpaid tasks`);
  
  // Store debug log for access
  PropertiesService.getScriptProperties().setProperty('lastDebugLog', JSON.stringify(debugLog));
  
  return unpaidWork;
}

// Get pay configuration from Pay Config sheet
function getPayConfiguration() {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
  const payConfigSheet = mainSheet.getSheetByName('Pay Config');
  
  if (!payConfigSheet) {
    throw new Error('Pay Config sheet not found');
  }
  
  const cols = getColumnIndices(payConfigSheet, [
    'Task Type (Stats Level)',
    'Default Rate',
    'Staff Name',
    'Custom Rate'
  ]);
  
  const data = payConfigSheet.getDataRange().getValues();
  const payConfig = {};
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const taskType = data[i][cols['Task Type (Stats Level)']];
    const defaultRate = parseFloat(String(data[i][cols['Default Rate']]).replace(/[^\d.-]/g, ''));
    const staffName = data[i][cols['Staff Name']];
    const customRate = data[i][cols['Custom Rate']] ? 
      parseFloat(String(data[i][cols['Custom Rate']]).replace(/[^\d.-]/g, '')) : null;
    
    if (!payConfig[taskType]) {
      payConfig[taskType] = {
        defaultRate: defaultRate,
        customRates: {}
      };
    }
    
    if (staffName && customRate) {
      payConfig[taskType].customRates[staffName] = customRate;
    }
  }
  
  return payConfig;
}

// Get staff name mapping
function getStaffMapping() {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = mainSheet.getSheetByName('Staff Key to Staff name');
  
  if (!staffSheet) {
    throw new Error('Staff Key to Staff name sheet not found');
  }
  
  const cols = getColumnIndices(staffSheet, ['Key', 'Name']);
  const data = staffSheet.getDataRange().getValues();
  const mapping = {};
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const key = data[i][cols['Key']];
    const legalName = data[i][cols['Name']];
    if (key && legalName) {
      mapping[key] = legalName;
    }
  }
  
  return mapping;
}

// Calculate payments for unpaid work
function calculatePayments(workLogData, payConfig, staffMapping) {
  const payments = {};
  const errors = {
    unmatchedTaskTypes: new Set(),
    unmatchedStaffKeys: new Set(),
    tasksWithNoRate: []
  };
  
  workLogData.forEach(work => {
    const staffKey = work.staffName;
    const taskType = work.taskType;
    
    // Check if staff key exists in mapping
    if (!staffMapping[staffKey]) {
      errors.unmatchedStaffKeys.add(staffKey);
    }
    
    // Get rate for this task
    let rate = 0;
    let rateSource = 'none';
    
    if (payConfig[taskType]) {
      // Check for custom rate first
      if (payConfig[taskType].customRates[staffKey]) {
        rate = payConfig[taskType].customRates[staffKey];
        rateSource = 'custom';
      } else {
        rate = payConfig[taskType].defaultRate;
        rateSource = 'default';
      }
    } else {
      errors.unmatchedTaskTypes.add(taskType);
      errors.tasksWithNoRate.push({
        staffName: staffKey,
        taskType: taskType,
        league: work.league,
        round: work.round,
        teams: `${work.team1} vs ${work.team2}`
      });
    }
    
    // Get legal name
    const legalName = staffMapping[staffKey] || staffKey;
    
    // Group by staff
    if (!payments[legalName]) {
      payments[legalName] = {
        staffKey: staffKey,
        legalName: legalName,
        hasMapping: !!staffMapping[staffKey],
        tasks: [],
        totalAmount: 0
      };
    }
    
    payments[legalName].tasks.push({
      ...work,
      rate: rate,
      rateSource: rateSource,
      hasValidRate: rate > 0
    });
    payments[legalName].totalAmount += rate;
  });
  
  return { payments, errors };
}

// Create payment summary
function createPaymentSummary(payments) {
  let summary = 'Payment Summary:\n\n';
  let grandTotal = 0;
  
  Object.values(payments).forEach(payment => {
    summary += `${payment.legalName}`;
    if (!payment.hasMapping) {
      summary += ` (⚠️ No legal name mapping)`;
    }
    summary += `:\n`;
    summary += `  Tasks: ${payment.tasks.length}\n`;
    
    // Show breakdown of rate sources
    const customRateTasks = payment.tasks.filter(t => t.rateSource === 'custom').length;
    const defaultRateTasks = payment.tasks.filter(t => t.rateSource === 'default').length;
    const noRateTasks = payment.tasks.filter(t => !t.hasValidRate).length;
    
    if (customRateTasks > 0) summary += `    - ${customRateTasks} with custom rate\n`;
    if (defaultRateTasks > 0) summary += `    - ${defaultRateTasks} with default rate\n`;
    if (noRateTasks > 0) summary += `    - ${noRateTasks} with NO RATE ⚠️\n`;
    
    summary += `  Total: ${formatCurrency(payment.totalAmount)}\n\n`;
    grandTotal += payment.totalAmount;
  });
  
  summary += `Grand Total: ${formatCurrency(grandTotal)}`;
  return summary;
}

// Format currency
function formatCurrency(amount) {
  return new Intl.NumberFormat('vi-VN', {
    style: 'currency',
    currency: 'VND'
  }).format(amount);
}

// Create invoices in Invoicing sheet
function createInvoices(payments) {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
  let invoicingSheet = mainSheet.getSheetByName('Invoicing');
  
  if (!invoicingSheet) {
    throw new Error('Invoicing sheet not found');
  }
  
  // Initialize debug log for this function
  const debugLog = [];
  
  // Invoicing sheet has headers in row 2
  const headerRow = 2;
  const headers = invoicingSheet.getRange(headerRow, 1, 1, invoicingSheet.getLastColumn()).getValues()[0];
  
  // Find columns by name
  const invoiceColumns = {
    'Invoice Number': headers.indexOf('Invoice Number'),
    'Date': headers.indexOf('Date'),
    'Contractor': headers.indexOf('Contractor'),
    'Work done': headers.indexOf('Work done'),
    'Total': headers.indexOf('Total'),
    'Playback Links': headers.indexOf('Playback Links')
  };
  
  // Check if all required columns exist
  Object.entries(invoiceColumns).forEach(([name, index]) => {
    if (index === -1) {
      Logger.log(`Warning: Column "${name}" not found in Invoicing sheet`);
    }
  });
  
  // Find the last row with data
  const lastRow = invoicingSheet.getLastRow();
  const nextRow = lastRow > headerRow ? lastRow + 1 : headerRow + 1; // Start after headers
  
  const invoiceData = [];
  const timestamp = new Date();
  
  // Generate invoice number based on date and sequence
  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd');
  let invoiceSequence = 1;
  
  Object.values(payments).forEach(payment => {
    // Group tasks by type
    const tasksByType = {};
    payment.tasks.forEach(task => {
      if (!tasksByType[task.taskType]) {
        tasksByType[task.taskType] = [];
      }
      tasksByType[task.taskType].push(task);
    });
    
    // Create work done summary in format: "3 x 1-Side - Basic\n35 x LEAGUE - BASIC"
    const workDoneSummary = Object.entries(tasksByType).map(([type, tasks]) => {
      return `${tasks.length} x ${type}`;
    }).join('\n');
    
    // Create playback links grouped by task type with line breaks
    const playbackLinks = Object.entries(tasksByType).map(([type, tasks]) => {
      const header = `${type} (${tasks.length} tasks):`;
      const links = tasks.map((task, index) => {
        return `  ${index + 1}. ${task.playbackLink}`;
      }).join('\n');
      return header + '\n' + links;
    }).join('\n\n');
    
    // Create row with proper number of columns (13 based on your sheet structure)
    const row = new Array(13).fill('');
    
    // Only populate the columns that exist
    if (invoiceColumns['Invoice Number'] !== -1) {
      row[invoiceColumns['Invoice Number']] = `INV-${dateStr}-${String(invoiceSequence).padStart(3, '0')}`;
    }
    if (invoiceColumns['Date'] !== -1) {
      row[invoiceColumns['Date']] = timestamp;
    }
    if (invoiceColumns['Contractor'] !== -1) {
      row[invoiceColumns['Contractor']] = payment.legalName;
    }
    if (invoiceColumns['Work done'] !== -1) {
      row[invoiceColumns['Work done']] = workDoneSummary;
    }
    if (invoiceColumns['Total'] !== -1) {
      row[invoiceColumns['Total']] = payment.totalAmount;
    }
    if (invoiceColumns['Playback Links'] !== -1) {
      row[invoiceColumns['Playback Links']] = playbackLinks;
    }
    
    invoiceData.push(row);
    invoiceSequence++;
  });
  
  // Check if we need more rows and add them if necessary
  const rowsNeeded = nextRow + invoiceData.length - 1;
  const currentMaxRows = invoicingSheet.getMaxRows();
  
  if (rowsNeeded > currentMaxRows) {
    const additionalRows = rowsNeeded - currentMaxRows + 10; // Add 10 extra for buffer
    invoicingSheet.insertRowsAfter(currentMaxRows, additionalRows);
    debugLog.push(`Added ${additionalRows} rows to Invoicing sheet`);
  }
  
  // Write to sheet
  if (invoiceData.length > 0) {
    const maxCols = Math.min(13, invoicingSheet.getMaxColumns());
    const targetRowCount = Math.min(invoiceData.length, invoicingSheet.getMaxRows() - nextRow + 1);
    
    // First, copy formulas from the row above (if it exists and has data)
    if (nextRow > headerRow + 1 && nextRow <= invoicingSheet.getMaxRows()) {
      try {
        const sourceRow = nextRow - 1;
        
        if (targetRowCount > 0) {
          const sourceRange = invoicingSheet.getRange(sourceRow, 1, 1, maxCols);
          const targetRange = invoicingSheet.getRange(nextRow, 1, targetRowCount, maxCols);
          
          // Copy formulas only (not values)
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
          debugLog.push(`Copied formulas from row ${sourceRow} to rows ${nextRow}-${nextRow + targetRowCount - 1}`);
        }
      } catch (error) {
        Logger.log('Formula copy failed: ' + error.toString());
        debugLog.push(`Formula copy failed: ${error.toString()}`);
      }
    }
    
    // Then write our data only to specific columns (not all columns)
    if (targetRowCount > 0) {
      // Write each column individually to preserve formulas in other columns
      Object.entries(invoiceColumns).forEach(([columnName, columnIndex]) => {
        if (columnIndex !== -1 && columnIndex < maxCols) {
          const columnData = invoiceData.slice(0, targetRowCount).map(row => [row[columnIndex]]);
          if (columnData.length > 0) {
            const columnRange = invoicingSheet.getRange(nextRow, columnIndex + 1, targetRowCount, 1);
            columnRange.setValues(columnData);
          }
        }
      });
    }
    
    // Set text wrapping for the Work done column
    if (invoiceColumns['Work done'] !== -1 && targetRowCount > 0) {
      const workDoneRange = invoicingSheet.getRange(
        nextRow, 
        invoiceColumns['Work done'] + 1, 
        targetRowCount, 
        1
      );
      workDoneRange.setWrap(true);
    }
    
    // Set text wrapping for the Playback Links column if it exists
    if (invoiceColumns['Playback Links'] !== -1 && targetRowCount > 0) {
      const playbackLinksRange = invoicingSheet.getRange(
        nextRow, 
        invoiceColumns['Playback Links'] + 1, 
        targetRowCount, 
        1
      );
      playbackLinksRange.setWrap(true);
    }
    
    // Adjust row heights to accommodate wrapped text
    for (let i = 0; i < targetRowCount; i++) {
      if (nextRow + i <= invoicingSheet.getMaxRows()) {
        invoicingSheet.setRowHeight(nextRow + i, 120);
      }
    }
  }
  
  // Store debug log if there were any messages
  if (debugLog.length > 0) {
    const existingLog = JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]');
    existingLog.push(...debugLog);
    PropertiesService.getScriptProperties().setProperty('lastDebugLog', JSON.stringify(existingLog));
  }
}

// Mark work as invoiced in MASTER sheet
function markWorkAsInvoiced(workLogData) {
  const workLogSheet = SpreadsheetApp.openById(CONFIG.workLogSheetId);
  const masterSheet = workLogSheet.getSheetByName(CONFIG.workLogSheetName);
  
  // Get the Paid column index - MASTER sheet has headers in row 2
  const paidColumnIndex = getColumnIndex(masterSheet, 'Paid', 2);
  
  workLogData.forEach(work => {
    masterSheet.getRange(work.rowIndex, paidColumnIndex).setValue('Invoiced');
  });
}

// Function to analyze sheet structure
function analyzeSheets() {
  const results = {};
  
  try {
    // Analyze main sheet (where this script is bound)
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    results.mainSheet = {
      name: mainSheet.getName(),
      sheets: mainSheet.getSheets().map(sheet => ({
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        cols: sheet.getLastColumn(),
        headers: sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : []
      }))
    };
    
    // Try to analyze work log sheet
    if (CONFIG.workLogSheetId) {
      const workLogSpreadsheet = SpreadsheetApp.openById(CONFIG.workLogSheetId);
      const masterSheet = workLogSpreadsheet.getSheetByName(CONFIG.workLogSheetName);
      
      results.workLogSheet = {
        spreadsheetName: workLogSpreadsheet.getName(),
        masterSheet: masterSheet ? {
          name: masterSheet.getName(),
          rows: masterSheet.getLastRow(),
          cols: masterSheet.getLastColumn(),
          headers: masterSheet.getLastRow() > 0 ? masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0] : [],
          sampleData: masterSheet.getLastRow() > 1 ? masterSheet.getRange(2, 1, Math.min(5, masterSheet.getLastRow() - 1), masterSheet.getLastColumn()).getValues() : []
        } : null,
        allSheets: workLogSpreadsheet.getSheets().map(sheet => sheet.getName())
      };
    }
    
    // Log results for debugging
    console.log(JSON.stringify(results, null, 2));
    
    // Show results in a dialog
    const html = '<pre>' + JSON.stringify(results, null, 2) + '</pre>';
    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(600)
        .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sheet Structure Analysis');
    
    return results;
  } catch (error) {
    Logger.log('Error analyzing sheets: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
    return { error: error.toString() };
  }
}

// Function to get sample data from sheets
function getSampleData() {
  const samples = {};
  
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get Pay Config data
    const payConfigSheet = mainSheet.getSheetByName('Pay Config');
    if (payConfigSheet && payConfigSheet.getLastRow() > 0) {
      const data = payConfigSheet.getDataRange().getValues();
      samples.payConfig = {
        headers: data[0],
        totalRows: data.length,
        sampleRows: data.slice(1, Math.min(11, data.length))
      };
    }
    
    // Get Staff Key data
    const staffKeySheet = mainSheet.getSheetByName('Staff Key to Staff name');
    if (staffKeySheet && staffKeySheet.getLastRow() > 0) {
      const data = staffKeySheet.getDataRange().getValues();
      samples.staffKey = {
        headers: data[0],
        totalRows: data.length,
        sampleRows: data.slice(1, Math.min(6, data.length))
      };
    }
    
    // Get current unpaid work count
    try {
      const unpaidWork = getUnpaidWorkFromMaster();
      samples.unpaidWorkCount = unpaidWork.length;
      samples.unpaidWorkSample = unpaidWork.slice(0, 3);
    } catch (e) {
      samples.unpaidWorkError = e.toString();
    }
    
    // Show results in a formatted dialog
    let html = '<div style="font-family: monospace; font-size: 12px;">';
    html += '<h3>Sample Data Analysis</h3>';
    html += '<pre>' + JSON.stringify(samples, null, 2) + '</pre>';
    html += '</div>';
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(700)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sample Data from Sheets');
    
    return samples;
  } catch (error) {
    Logger.log('Error getting sample data: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
    return { error: error.toString() };
  }
}

// Function to show debug log
function showDebugLog() {
  const debugLog = JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]');
  const debugInfo = debugLog.join('\n');
  
  const htmlContent = `
    <div style="font-family: monospace; font-size: 12px;">
      <h3>Debug Log - Work Detection</h3>
      <div style="background-color: #f5f5f5; padding: 10px; border-radius: 5px; white-space: pre-wrap; max-height: 400px; overflow-y: auto;">
${debugInfo}
      </div>
      <div style="margin-top: 20px;">
        <button onclick="google.script.host.close()" style="padding: 10px 20px;">Close</button>
      </div>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(700)
      .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Debug Log');
}

// Function to create invoices and mark work (called from HTML dialog)
function createInvoicesAndMark() {
  try {
    const workLogData = JSON.parse(PropertiesService.getScriptProperties().getProperty('pendingWorkData') || '[]');
    const payments = JSON.parse(PropertiesService.getScriptProperties().getProperty('pendingPayments') || '{}');
    
    if (workLogData.length === 0 || Object.keys(payments).length === 0) {
      throw new Error('No pending work data found. Please run Calculate Staff Pay again.');
    }
    
    // Create invoices
    createInvoices(payments);
    
    // Mark work as invoiced
    markWorkAsInvoiced(workLogData);
    
    // Clean up
    PropertiesService.getScriptProperties().deleteProperty('pendingWorkData');
    PropertiesService.getScriptProperties().deleteProperty('pendingPayments');
    
    SpreadsheetApp.getUi().alert('Success', 'Invoices created and work marked as invoiced.', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(error);
  }
}