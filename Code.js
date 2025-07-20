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
    .addItem('Calculate Staff Pay', 'calculateStaffPayUI')
    .addSeparator()
    .addItem('Create Invoices', 'createInvoicesAndMarkUI')
    .addSeparator()
    .addItem('Analyze Sheet Structure', 'analyzeSheets')
    .addItem('Get Sample Data', 'getSampleData')
    .addToUi();
}

/**
 * Web App Interface - serves both HTML UI and API endpoints
 * 
 * Usage as Web App:
 * - No parameters: Returns HTML interface for manual use
 * - ?action=preview: Returns payment preview as JSON
 * - ?action=calculatePay: Executes payment calculation and creates invoices
 * - ?action=getStatus: Returns current status (unpaid tasks count)
 * - ?action=getDebugLog: Returns debug information
 * - ?action=getDeploymentUrl: Returns the current deployment URL for frontend use
 * 
 * Direct API Usage:
 * - calculateStaffPay() - Returns payment calculation results
 * - createInvoicesAndMark(workLogData, payments) - Creates invoices and marks work as paid
 * - getUnpaidWorkFromMaster() - Gets unpaid work data
 * - getPayConfiguration() - Gets pay rates configuration
 * - getStaffMapping() - Gets staff name mappings
 * 
 * Deployment Management:
 * - setCurrentDeploymentUrl(url) - Store deployment URL after deploying
 * - updateDeploymentAfterPush(deploymentId) - Helper to update URL with deployment ID
 */
function doGet(e) {
  try {
    const action = e.parameter ? e.parameter.action : null;
    
    // If no action parameter, serve the web interface
    if (!action) {
      return createWebAppInterface();
    }
    
    let result;
    switch(action) {
      case 'preview':
        result = handleCalculatePayPreviewRequest();
        break;
      case 'calculatePay':
        result = handleCalculatePayRequest();
        break;
      case 'getStatus':
        result = handleStatusRequest();
        break;
      case 'getDebugLog':
        result = handleDebugLogRequest();
        break;
      case 'test':
        result = ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            message: 'Test endpoint working',
            timestamp: new Date().toISOString()
          }))
          .setMimeType(ContentService.MimeType.JSON);
        break;
      case 'exportLatestInvoice':
        result = ContentService
          .createTextOutput(JSON.stringify(exportLatestInvoicePDF()))
          .setMimeType(ContentService.MimeType.JSON);
        break;
      case 'exportInvoicePDF':
        const invoiceNum = e.parameter.invoiceNumber;
        const days = e.parameter.daysBack ? parseInt(e.parameter.daysBack) : 30;
        result = ContentService
          .createTextOutput(JSON.stringify(exportInvoicesPDF(invoiceNum, days)))
          .setMimeType(ContentService.MimeType.JSON);
        break;
      default:
        result = ContentService
          .createTextOutput(JSON.stringify({
            error: 'Invalid action. Available actions: preview, calculatePay, getStatus, getDebugLog, test, exportLatestInvoice, exportInvoicePDF'
          }))
          .setMimeType(ContentService.MimeType.JSON);
    }
    
    return result;
  } catch (error) {
    // Log the error for debugging
    console.error('Error in doGet:', error);
    
    const errorResult = ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: 'Server error: ' + error.toString(),
        stack: error.stack || 'No stack trace available'
      }))
      .setMimeType(ContentService.MimeType.JSON);
    
    return errorResult;
  }
}

// Add CORS headers to allow cross-origin requests
function addCorsHeaders(response) {
  // ContentService responses don't support setHeaders, so we return the response as-is
  // CORS is handled at the web app deployment level in Google Apps Script
  return response;
}

/**
 * Creates the web app HTML interface
 * Provides a user-friendly interface for external consumers
 */
function createWebAppInterface() {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <title>Staff Pay Calculator</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      max-width: 1200px;
      margin: 0 auto;
      padding: 20px;
      background-color: #f5f7fa;
    }
    .container {
      background: white;
      border-radius: 8px;
      padding: 30px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    h1 {
      color: #2c3e50;
      text-align: center;
      margin-bottom: 30px;
    }
    .action-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
      gap: 20px;
      margin-bottom: 30px;
    }
    .action-card {
      border: 1px solid #e1e8ed;
      border-radius: 8px;
      padding: 20px;
      text-align: center;
      transition: transform 0.2s, box-shadow 0.2s;
    }
    .action-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .btn {
      background: #3498db;
      color: white;
      border: none;
      padding: 12px 24px;
      border-radius: 5px;
      cursor: pointer;
      font-size: 16px;
      transition: background-color 0.3s;
      margin: 5px;
    }
    .btn:hover {
      background: #2980b9;
    }
    .btn-success {
      background: #27ae60;
    }
    .btn-success:hover {
      background: #229954;
    }
    .btn-warning {
      background: #f39c12;
    }
    .btn-warning:hover {
      background: #e67e22;
    }
    .btn-info {
      background: #8e44ad;
    }
    .btn-info:hover {
      background: #7d3c98;
    }
    .results {
      margin-top: 20px;
      padding: 20px;
      border-radius: 5px;
      display: none;
    }
    .results.success {
      background: #d4edda;
      border: 1px solid #c3e6cb;
      color: #155724;
    }
    .results.error {
      background: #f8d7da;
      border: 1px solid #f5c6cb;
      color: #721c24;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 20px;
    }
    .spinner {
      border: 4px solid #f3f3f3;
      border-top: 4px solid #3498db;
      border-radius: 50%;
      width: 40px;
      height: 40px;
      animation: spin 1s linear infinite;
      margin: 0 auto 10px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    pre {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 5px;
      overflow-x: auto;
      white-space: pre-wrap;
    }
    .api-docs {
      margin-top: 40px;
      padding: 20px;
      background: #f8f9fa;
      border-radius: 8px;
    }
    .api-docs h3 {
      color: #2c3e50;
      margin-top: 0;
    }
    .endpoint {
      background: white;
      padding: 10px;
      margin: 10px 0;
      border-left: 4px solid #3498db;
      border-radius: 4px;
    }
    code {
      background: #e9ecef;
      padding: 2px 6px;
      border-radius: 3px;
      font-family: 'Courier New', monospace;
    }
    input[type="text"], input[type="number"] {
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
      box-sizing: border-box;
    }
    input[type="text"]:focus, input[type="number"]:focus {
      outline: none;
      border-color: #3498db;
      box-shadow: 0 0 5px rgba(52, 152, 219, 0.3);
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>Staff Pay Calculator</h1>
    
    <div class="action-grid">
      <div class="action-card">
        <h3>Preview Payments</h3>
        <p>Review payment calculations without creating invoices</p>
        <button class="btn" onclick="previewPayments()">Preview</button>
      </div>
      
      <div class="action-card">
        <h3>Calculate & Create Invoices</h3>
        <p>Calculate payments and create invoices in the spreadsheet</p>
        <button class="btn btn-success" onclick="calculatePayments()">Execute</button>
      </div>
      
      <div class="action-card">
        <h3>Check Status</h3>
        <p>View current status and unpaid task count</p>
        <button class="btn btn-warning" onclick="checkStatus()">Status</button>
      </div>
      
      <div class="action-card">
        <h3>Debug Information</h3>
        <p>View debug logs for troubleshooting</p>
        <button class="btn btn-info" onclick="getDebugLog()">Debug</button>
      </div>
      
      <div class="action-card">
        <h3>Test Connection</h3>
        <p>Test basic API connectivity</p>
        <button class="btn" onclick="testConnection()">Test</button>
      </div>
      
      
      <div class="action-card">
        <h3>Export Latest Invoice</h3>
        <p>Export the most recently created invoice as PDF</p>
        <button class="btn btn-warning" onclick="exportLatestInvoice()">Export PDF</button>
      </div>
      
      <div class="action-card">
        <h3>Export Invoice by Number</h3>
        <p>Export specific invoice by invoice number</p>
        <input type="text" id="invoiceNumber" placeholder="Invoice Number" style="width: 100%; margin-bottom: 10px; padding: 8px;">
        <button class="btn btn-warning" onclick="exportInvoiceByNumber()">Export PDF</button>
      </div>
      
      <div class="action-card">
        <h3>Export Recent Invoices</h3>
        <p>Export invoices from the last N days</p>
        <input type="number" id="daysBack" placeholder="Days back (default: 30)" value="30" style="width: 100%; margin-bottom: 10px; padding: 8px;">
        <button class="btn btn-warning" onclick="exportRecentInvoices()">Export PDF</button>
      </div>
    </div>
    
    <div class="loading" id="loading">
      <div class="spinner"></div>
      <div>Processing request...</div>
    </div>
    
    <div class="results" id="results"></div>
    
    <div class="api-docs">
      <h3>API Documentation</h3>
      <p>This web app provides both a user interface and programmatic API access:</p>
      
      <div class="endpoint">
        <strong>GET ?action=preview</strong><br>
        Returns payment preview without creating invoices
      </div>
      
      <div class="endpoint">
        <strong>GET ?action=calculatePay</strong><br>
        Calculates payments and creates invoices
      </div>
      
      <div class="endpoint">
        <strong>GET ?action=getStatus</strong><br>
        Returns current status (unpaid task count)
      </div>
      
      <div class="endpoint">
        <strong>GET ?action=getDebugLog</strong><br>
        Returns debug information
      </div>
      
      
      <div class="endpoint">
        <strong>GET ?action=test</strong><br>
        Test endpoint to verify API connectivity
      </div>
      
      <div class="endpoint">
        <strong>GET ?action=exportLatestInvoice</strong><br>
        Exports the most recently created invoice as PDF
      </div>
      
      <div class="endpoint">
        <strong>GET ?action=exportInvoicePDF&invoiceNumber=[number]</strong><br>
        Exports specific invoice by invoice number as PDF
      </div>
      
      <div class="endpoint">
        <strong>GET ?action=exportInvoicePDF&daysBack=[number]</strong><br>
        Exports invoices from the last N days as PDF (default: 30 days)
      </div>
      
      <h4>Direct Function Calls (Google Apps Script)</h4>
      <p>You can also call these functions directly:</p>
      <ul>
        <li><code>calculateStaffPay()</code> - Returns payment calculation results</li>
        <li><code>createInvoicesAndMark(workLogData, payments)</code> - Creates invoices</li>
        <li><code>getUnpaidWorkFromMaster()</code> - Gets unpaid work data</li>
        <li><code>getPayConfiguration()</code> - Gets pay rates</li>
        <li><code>getStaffMapping()</code> - Gets staff mappings</li>
      </ul>
    </div>
  </div>

  <script>
    function showLoading() {
      document.getElementById('loading').style.display = 'block';
      document.getElementById('results').style.display = 'none';
    }
    
    function hideLoading() {
      document.getElementById('loading').style.display = 'none';
    }
    
    function showResults(content, isError = false) {
      const resultsDiv = document.getElementById('results');
      resultsDiv.innerHTML = content;
      resultsDiv.className = 'results ' + (isError ? 'error' : 'success');
      resultsDiv.style.display = 'block';
      hideLoading();
    }
    
    function previewPayments() {
      showLoading();
      console.log('Calling handleCalculatePayPreviewRequest with directReturn=true');
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .handleCalculatePayPreviewRequest(true);
    }
    
    function calculatePayments() {
      if (confirm('This will create invoices and mark work as paid. Continue?')) {
        showLoading();
        google.script.run
          .withSuccessHandler(handleSuccess)
          .withFailureHandler(handleFailure)
          .handleCalculatePayRequest();
      }
    }
    
    function checkStatus() {
      showLoading();
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .handleStatusRequest();
    }
    
    function getDebugLog() {
      showLoading();
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .handleDebugLogRequest();
    }
    
    function testConnection() {
      showLoading();
      console.log('Calling simpleTest function');
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .simpleTest();
    }
    
    
    function handleSuccess(result) {
      console.log('handleSuccess called with result:', result);
      console.log('Result type:', typeof result);
      console.log('Result keys:', Object.keys(result || {}));
      
      // Check if result is null or undefined
      if (!result) {
        console.error('Result is null/undefined');
        showResults('<strong>Error:</strong> Server function returned null. Check server-side logs.', true);
        return;
      }
      
      // Result should now be a direct object when using directReturn=true
      if (result.error) {
        console.log('Error found:', result.error);
        showResults('<strong>Error:</strong> ' + result.error, true);
      } else {
        console.log('Attempting to JSON.stringify result...');
        try {
          const jsonString = JSON.stringify(result, null, 2);
          console.log('JSON stringify successful, length:', jsonString.length);
          showResults('<pre>' + jsonString + '</pre>');
        } catch (e) {
          console.error('JSON.stringify failed:', e);
          showResults('<strong>JSON Error:</strong> Could not serialize result - ' + e.message, true);
        }
      }
    }
    
    function handleFailure(error) {
      console.error('handleFailure called with error:', error);
      console.error('Error message:', error.message);
      console.error('Error stack:', error.stack);
      showResults('<strong>Request failed:</strong> ' + error.message, true);
    }
    
    function exportLatestInvoice() {
      showLoading();
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .exportLatestInvoicePDF();
    }
    
    function exportInvoiceByNumber() {
      const invoiceNumber = document.getElementById('invoiceNumber').value.trim();
      if (!invoiceNumber) {
        showResults('<strong>Error:</strong> Please enter an invoice number', true);
        return;
      }
      showLoading();
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .exportInvoicesPDF(invoiceNumber, null);
    }
    
    function exportRecentInvoices() {
      const daysBack = parseInt(document.getElementById('daysBack').value) || 30;
      showLoading();
      google.script.run
        .withSuccessHandler(handleSuccess)
        .withFailureHandler(handleFailure)
        .exportInvoicesPDF(null, daysBack);
    }
  </script>
</body>
</html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('Staff Pay Calculator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  // Handle POST requests for specific function calls
  const data = JSON.parse(e.postData.contents);
  const functionName = data.function;
  const preview = data.preview;
  
  try {
    let result;
    switch(functionName) {
      case 'calculateStaffPay':
        if (preview) {
          // For preview, return calculation data without executing
          result = handleCalculatePayPreviewRequest();
        } else {
          // For non-preview, just return calculation data (don't create invoices)
          const calcResult = calculateStaffPay();
          result = ContentService
            .createTextOutput(JSON.stringify(calcResult))
            .setMimeType(ContentService.MimeType.JSON);
        }
        break;
      case 'createInvoicesAndMark':
        // For this function, we need to get the calculation first, then create invoices
        const payResult = calculateStaffPay();
        if (!payResult.success) {
          result = ContentService
            .createTextOutput(JSON.stringify(payResult))
            .setMimeType(ContentService.MimeType.JSON);
        } else {
          const invoiceResult = createInvoicesAndMark(payResult.workLogData, payResult.payments);
          result = ContentService
            .createTextOutput(JSON.stringify({
              ...invoiceResult,
              summary: payResult.summary
            }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        break;
      default:
        result = ContentService
          .createTextOutput(JSON.stringify({
            error: `Invalid function name: ${functionName}. Use 'calculateStaffPay' or 'createInvoicesAndMark'`
          }))
          .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Add CORS headers
    return result.getHeaders ? addCorsHeaders(result) : result;
  } catch (error) {
    const errorResult = ContentService
      .createTextOutput(JSON.stringify({
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
    
    return addCorsHeaders(errorResult);
  }
}

// Handle calculate pay preview request from web (no execution)
function handleCalculatePayPreviewRequest(directReturn = false) {
  try {
    const workLogData = getUnpaidWorkFromMaster();
    
    if (workLogData.length === 0) {
      const result = {
        success: false,
        message: 'No unpaid work found',
        debugLog: JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]')
      };
      
      if (directReturn) {
        return result;
      } else {
        return ContentService
          .createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    const payConfig = getPayConfiguration();
    const staffMapping = getStaffMapping();
    const { payments, errors } = calculatePayments(workLogData, payConfig, staffMapping);
    
    // Create a clean, serializable result object
    const result = {
      success: true,
      message: 'Payment preview calculated successfully',
      summary: {
        totalTasks: workLogData.length,
        totalStaff: Object.keys(payments).length,
        grandTotal: Object.values(payments).reduce((sum, p) => sum + p.totalAmount, 0),
        errors: {
          unmatchedTaskTypes: Array.from(errors.unmatchedTaskTypes || []),
          unmatchedStaffKeys: Array.from(errors.unmatchedStaffKeys || []),
          tasksWithNoRate: Array.from(errors.tasksWithNoRate || [])
        }
      },
      // Create a clean payments object that's guaranteed to serialize
      payments: {}
    };
    
    // Manually construct payments object to avoid serialization issues
    Object.keys(payments).forEach(staffName => {
      const payment = payments[staffName];
      result.payments[staffName] = {
        staffKey: String(payment.staffKey || ''),
        legalName: String(payment.legalName || ''),
        hasMapping: Boolean(payment.hasMapping),
        totalAmount: Number(payment.totalAmount || 0),
        taskCount: Array.isArray(payment.tasks) ? payment.tasks.length : 0,
        tasks: Array.isArray(payment.tasks) ? payment.tasks.map(task => ({
          rowIndex: Number(task.rowIndex || 0),
          staffName: String(task.staffName || ''),
          taskType: String(task.taskType || ''),
          league: String(task.league || ''),
          round: String(task.round || ''),
          team1: String(task.team1 || ''),
          team2: String(task.team2 || ''),
          rate: Number(task.rate || 0),
          rateSource: String(task.rateSource || ''),
          hasValidRate: Boolean(task.hasValidRate)
        })) : []
      };
    });
    
    if (directReturn) {
      return result;
    } else {
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
      
  } catch (error) {
    const errorResult = {
      success: false,
      error: error.toString()
    };
    
    if (directReturn) {
      return errorResult;
    } else {
      return ContentService
        .createTextOutput(JSON.stringify(errorResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
}

// Handle calculate pay request from web
function handleCalculatePayRequest() {
  try {
    const workLogData = getUnpaidWorkFromMaster();
    
    if (workLogData.length === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'No unpaid work found',
          debugLog: JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]')
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const payConfig = getPayConfiguration();
    const staffMapping = getStaffMapping();
    const { payments, errors } = calculatePayments(workLogData, payConfig, staffMapping);
    
    // Auto-create invoices (since this is from web, assume user wants to proceed)
    createInvoice(payments);
    markWorkAsInvoiced(workLogData);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Invoices created successfully',
        summary: {
          totalTasks: workLogData.length,
          totalStaff: Object.keys(payments).length,
          grandTotal: Object.values(payments).reduce((sum, p) => sum + p.totalAmount, 0),
          errors: {
            unmatchedTaskTypes: Array.from(errors.unmatchedTaskTypes),
            unmatchedStaffKeys: Array.from(errors.unmatchedStaffKeys),
            tasksWithNoRate: errors.tasksWithNoRate
          }
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle status request
function handleStatusRequest() {
  try {
    const workLogData = getUnpaidWorkFromMaster();
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        status: {
          unpaidTasks: workLogData.length,
          lastCheck: new Date().toISOString()
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle debug log request
function handleDebugLogRequest() {
  const debugLog = JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]');
  
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      debugLog: debugLog
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle get deployment URL request

// Test endpoint for basic connectivity
function testEndpoint() {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: 'Test endpoint working - direct function call',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Simple test function for google.script.run
function simpleTest() {
  return {
    success: true,
    message: 'Simple test working',
    timestamp: new Date().toISOString()
  };
}

// Simplified preview function to test step by step
function testPreview() {
  try {
    // Test 1: Basic function call
    console.log('testPreview: Starting');
    
    // Test 2: Can we get work log data?
    const workLogData = getUnpaidWorkFromMaster();
    console.log('testPreview: Work log data length:', workLogData.length);
    
    if (workLogData.length === 0) {
      return {
        success: false,
        message: 'No unpaid work found',
        step: 'workLogData check'
      };
    }
    
    // Test 3: Can we get config data?
    const payConfig = getPayConfiguration();
    console.log('testPreview: Pay config keys:', Object.keys(payConfig));
    
    const staffMapping = getStaffMapping();
    console.log('testPreview: Staff mapping keys:', Object.keys(staffMapping));
    
    // Test 4: Can we calculate payments?
    const { payments, errors } = calculatePayments(workLogData, payConfig, staffMapping);
    console.log('testPreview: Payments calculated, staff count:', Object.keys(payments).length);
    
    // Test 5: Return minimal result
    return {
      success: true,
      message: 'Test preview completed',
      taskCount: workLogData.length,
      staffCount: Object.keys(payments).length
    };
    
  } catch (error) {
    console.error('testPreview error:', error);
    return {
      success: false,
      error: error.toString(),
      step: 'exception caught'
    };
  }
}

/**
 * Deployment Management Functions
 * Use these to manage deployment URLs dynamically
 */

// Set the current deployment URL (call this after deploying)
function setCurrentDeploymentUrl(url) {
  if (!url) {
    throw new Error('URL is required');
  }
  
  PropertiesService.getScriptProperties().setProperty('CURRENT_DEPLOYMENT_URL', url);
  
  console.log(`Deployment URL updated to: ${url}`);
  return {
    success: true,
    url: url,
    timestamp: new Date().toISOString()
  };
}

// Get the current deployment URL (for internal use)
function getCurrentDeploymentUrl() {
  return PropertiesService.getScriptProperties().getProperty('CURRENT_DEPLOYMENT_URL');
}

// Utility function to help with deployment
function updateDeploymentAfterPush(deploymentId) {
  const baseUrl = 'https://script.google.com/macros/s/';
  const fullUrl = `${baseUrl}${deploymentId}/exec`;
  
  return setCurrentDeploymentUrl(fullUrl);
}

/**
 * Core API function to calculate staff pay
 * Returns payment calculation results without UI interaction
 * 
 * @return {Object} {
 *   success: boolean,
 *   workLogData: Array,
 *   payments: Object,
 *   errors: Object,
 *   summary: Object
 * }
 */
function calculateStaffPay() {
  try {
    const workLogData = getUnpaidWorkFromMaster();
    
    if (workLogData.length === 0) {
      return {
        success: false,
        message: 'No unpaid work found',
        debugLog: JSON.parse(PropertiesService.getScriptProperties().getProperty('lastDebugLog') || '[]')
      };
    }
    
    const payConfig = getPayConfiguration();
    const staffMapping = getStaffMapping();
    const { payments, errors } = calculatePayments(workLogData, payConfig, staffMapping);
    
    return {
      success: true,
      workLogData: workLogData,
      payments: payments,
      errors: {
        unmatchedTaskTypes: Array.from(errors.unmatchedTaskTypes),
        unmatchedStaffKeys: Array.from(errors.unmatchedStaffKeys),
        tasksWithNoRate: errors.tasksWithNoRate
      },
      summary: {
        totalTasks: workLogData.length,
        totalStaff: Object.keys(payments).length,
        grandTotal: Object.values(payments).reduce((sum, p) => sum + p.totalAmount, 0)
      }
    };
    
  } catch (error) {
    Logger.log(error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Google Sheets UI version of calculateStaffPay
 * Shows interactive dialog for confirmation and actions
 */
function calculateStaffPayUI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const result = calculateStaffPay();
    
    if (!result.success) {
      if (result.message === 'No unpaid work found') {
        const response = ui.alert(
          'No unpaid work found',
          'All tasks are either already paid or invoiced.\n\nClick "Yes" to see debug log, "No" to exit.',
          ui.ButtonSet.YES_NO
        );
        
        if (response === ui.Button.YES) {
          showDebugLog();
        }
      } else {
        ui.alert('Error', result.error || 'An unknown error occurred', ui.ButtonSet.OK);
      }
      return;
    }
    
    const { payments, errors } = result;
    
    // Build error message
    let errorMessage = '';
    if (errors.unmatchedTaskTypes.length > 0) {
      errorMessage += `\n⚠️ Unmatched Task Types in Pay Config:\n`;
      errors.unmatchedTaskTypes.forEach(type => {
        errorMessage += `  - ${type}\n`;
      });
    }
    
    if (errors.unmatchedStaffKeys.length > 0) {
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
    
    const summary = createPaymentSummary(payments);
    
    const htmlContent = `
      <div style="font-family: Arial, sans-serif;">
        <h3>Payment Summary</h3>
        <pre style="background-color: #f5f5f5; padding: 10px; border-radius: 5px;">${summary}</pre>
        ${errorMessage ? `<div style="color: red; margin-top: 10px;"><pre>${errorMessage}</pre></div>` : ''}
        <div style="margin-top: 20px;">
          <p>Choose an action:</p>
          <button onclick="google.script.run.createInvoicesAndMarkUI()" style="padding: 10px 20px; margin-right: 10px;">Create Invoices</button>
          <button onclick="google.script.run.showDebugLog()" style="padding: 10px 20px; margin-right: 10px;">Show Debug Log</button>
          <button onclick="google.script.host.close()" style="padding: 10px 20px;">Cancel</button>
        </div>
      </div>
    `;
    
    // Use simple alert instead of modal dialog
    let alertMessage = 'Payment Summary:\n\n' + summary;
    if (errorMessage) {
      alertMessage += '\n\nErrors:\n' + errorMessage;
    }
    alertMessage += '\n\nUse "Create Invoices" from the menu to proceed.';
    
    ui.alert('Payment Summary', alertMessage, ui.ButtonSet.OK);
    
    // Store for UI callback
    PropertiesService.getScriptProperties().setProperty('pendingWorkData', JSON.stringify(result.workLogData));
    PropertiesService.getScriptProperties().setProperty('pendingPayments', JSON.stringify(result.payments));
    
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

// Create invoice in Invoicing sheet
function createInvoice(payments) {
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
  
  // Generate single invoice number for all rows in this invocation
  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd');
  const invoiceNumber = `INV-${dateStr}-${String(Math.floor(Math.random() * 1000)).padStart(3, '0')}`;
  
  debugLog.push(`Creating invoice ${invoiceNumber} for ${Object.keys(payments).length} staff members`);
  
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
    
    // Create row with proper number of columns
    const maxColumnIndex = Math.max(...Object.values(invoiceColumns).filter(i => i !== -1));
    const row = new Array(maxColumnIndex + 1).fill('');
    
    // Only populate the columns that exist
    if (invoiceColumns['Invoice Number'] !== -1) {
      row[invoiceColumns['Invoice Number']] = invoiceNumber;
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
  });
  
  // Check if we need more rows and add exactly the right amount
  const rowsNeeded = nextRow + invoiceData.length - 1;
  const currentMaxRows = invoicingSheet.getMaxRows();
  
  if (rowsNeeded > currentMaxRows) {
    const additionalRows = rowsNeeded - currentMaxRows; // Add exactly what we need
    invoicingSheet.insertRowsAfter(currentMaxRows, additionalRows);
    debugLog.push(`Added ${additionalRows} rows to Invoicing sheet (exactly ${invoiceData.length} invoice rows needed)`);
  }
  
  // Write invoice data to sheet
  if (invoiceData.length > 0) {
    const maxCols = invoicingSheet.getMaxColumns();
    
    // First, copy formulas from the row above (if it exists and has data)
    if (nextRow > headerRow + 1 && nextRow <= invoicingSheet.getMaxRows()) {
      try {
        const sourceRow = nextRow - 1;
        const sourceRange = invoicingSheet.getRange(sourceRow, 1, 1, maxCols);
        const targetRange = invoicingSheet.getRange(nextRow, 1, invoiceData.length, maxCols);
        
        // Copy formulas only (not values)
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        debugLog.push(`Copied formulas from row ${sourceRow} to rows ${nextRow}-${nextRow + invoiceData.length - 1}`);
      } catch (error) {
        Logger.log('Formula copy failed: ' + error.toString());
        debugLog.push(`Formula copy failed: ${error.toString()}`);
      }
    }
    
    // Then write our data only to specific columns (preserving formulas in other columns)
    Object.entries(invoiceColumns).forEach(([columnName, columnIndex]) => {
      if (columnIndex !== -1 && columnIndex < maxCols) {
        const columnData = invoiceData.map(row => [row[columnIndex]]);
        if (columnData.length > 0) {
          const columnRange = invoicingSheet.getRange(nextRow, columnIndex + 1, columnData.length, 1);
          columnRange.setValues(columnData);
        }
      }
    });
    
    debugLog.push(`Wrote ${invoiceData.length} invoice rows starting at row ${nextRow}`);
    
    // Set text wrapping for the Work done column
    if (invoiceColumns['Work done'] !== -1) {
      const workDoneRange = invoicingSheet.getRange(
        nextRow, 
        invoiceColumns['Work done'] + 1, 
        invoiceData.length, 
        1
      );
      workDoneRange.setWrap(true);
    }
    
    // Set text wrapping for the Playback Links column if it exists
    if (invoiceColumns['Playback Links'] !== -1) {
      const playbackLinksRange = invoicingSheet.getRange(
        nextRow, 
        invoiceColumns['Playback Links'] + 1, 
        invoiceData.length, 
        1
      );
      playbackLinksRange.setWrap(true);
    }
    
    // Adjust row heights to accommodate wrapped text
    for (let i = 0; i < invoiceData.length; i++) {
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
  
  return {
    invoiceNumber: invoiceNumber,     // The generated invoice number
    invoiceDate: timestamp,           // When invoice was created
    rowsCreated: invoiceData.length   // Number of staff rows
  };
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
    // Use simple alert instead of modal dialog
    SpreadsheetApp.getUi().alert('Sheet Structure Analysis', `Found ${results.sheets.length} sheets. Check server logs for detailed analysis.`, SpreadsheetApp.getUi().ButtonSet.OK);
    
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
    
    // Use simple alert instead of modal dialog
    SpreadsheetApp.getUi().alert('Sample Data', 'Sample data collected from sheets. Check the server logs for detailed output.', SpreadsheetApp.getUi().ButtonSet.OK);
    
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
  
  // Use simple alert instead of modal dialog
  const logText = debugLog.length > 0 ? debugLog.join('\n') : 'No debug log entries found';
  SpreadsheetApp.getUi().alert('Debug Log', logText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Core API function to create invoices and mark work as invoiced
 * 
 * @param {Array} workLogData - Array of work items to mark as invoiced
 * @param {Object} payments - Payment data organized by staff member
 * @return {Object} { success: boolean, message?: string, error?: string }
 */
function createInvoicesAndMark(workLogData, payments) {
  try {
    if (!workLogData || workLogData.length === 0) {
      return {
        success: false,
        error: 'No work data provided'
      };
    }
    
    if (!payments || Object.keys(payments).length === 0) {
      return {
        success: false,
        error: 'No payment data provided'
      };
    }
    
    const invoiceResult = createInvoice(payments);
    markWorkAsInvoiced(workLogData);
    
    return {
      success: true,
      message: 'Invoices created and work marked as invoiced',
      invoiceResult: invoiceResult
    };
    
  } catch (error) {
    Logger.log(error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Google Sheets UI version - uses stored properties from calculateStaffPayUI
 */
function createInvoicesAndMarkUI() {
  try {
    const workLogData = JSON.parse(PropertiesService.getScriptProperties().getProperty('pendingWorkData') || '[]');
    const payments = JSON.parse(PropertiesService.getScriptProperties().getProperty('pendingPayments') || '{}');
    
    const result = createInvoicesAndMark(workLogData, payments);
    
    // Clean up stored properties
    PropertiesService.getScriptProperties().deleteProperty('pendingWorkData');
    PropertiesService.getScriptProperties().deleteProperty('pendingPayments');
    
    if (result.success) {
      SpreadsheetApp.getUi().alert('Success', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Error', result.error, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(error);
  }
}


/**
 * Export specific invoice rows as PDF
 * @param {string} invoiceNumber - Optional specific invoice number to export
 * @param {number} daysBack - Number of days back to include invoices (default: 30)
 * @returns {Object} Success status and PDF file URL
 */
function exportInvoicesPDF(invoiceNumber = null, daysBack = 30, specificDate = null) {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const invoicingSheet = mainSheet.getSheetByName('Invoicing');
    
    if (!invoicingSheet) {
      throw new Error('Invoicing sheet not found');
    }
    
    const headerRow = 2;
    const lastRow = invoicingSheet.getLastRow();
    
    if (lastRow <= headerRow) {
      return {
        success: false,
        error: 'No invoice data found'
      };
    }
    
    // Get all data including headers
    const allData = invoicingSheet.getRange(headerRow, 1, lastRow - headerRow + 1, invoicingSheet.getLastColumn()).getValues();
    const allHeaders = allData[0];
    const allDataRows = allData.slice(1);
    
    // Define columns to export
    const exportColumns = ['Contractor', 'Work done', 'Total', 'Email', 'ACCOUNT NAME', 'ACC NUMBER', 'BANK', 'Playback Links'];
    
    // Find column indices for filtering and export columns
    const invoiceNumberCol = allHeaders.indexOf('Invoice Number');
    const dateCol = allHeaders.indexOf('Date');
    
    const exportColIndices = exportColumns.map(col => allHeaders.indexOf(col));
    const missingColumns = exportColumns.filter((col, idx) => exportColIndices[idx] === -1);
    
    if (missingColumns.length > 0) {
      throw new Error(`Missing columns: ${missingColumns.join(', ')}`);
    }
    
    if (invoiceNumberCol === -1) {
      throw new Error('Invoice Number column not found');
    }
    
    // Filter rows based on criteria
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysBack);
    let filteredAllDataRows = [];
    if (invoiceNumber) {
      if (specificDate) {
        // Export specific invoice by number AND date (unique combination - used for latest invoice)
        const targetDate = new Date(specificDate);
        filteredAllDataRows = allDataRows.filter(row => {
          const rowDate = new Date(row[dateCol]);
          return row[invoiceNumberCol] === invoiceNumber && 
                 rowDate.toDateString() === targetDate.toDateString();
        });
      } else {
        // Export all rows with this invoice number (regardless of date)
        filteredAllDataRows = allDataRows.filter(row => row[invoiceNumberCol] === invoiceNumber);
      }
    } else {
      // Export recent invoices (when no specific invoice number provided)
      filteredAllDataRows = allDataRows.filter(row => {
        const rowDate = new Date(row[dateCol]);
        return rowDate >= cutoffDate;
      });
    }
    
    // Extract only the specified columns from filtered data
    const headers = exportColumns;
    const accNumberIndex = exportColumns.indexOf('ACC NUMBER');
    
    const filteredRows = filteredAllDataRows.map(row => 
      exportColIndices.map((colIndex, exportIndex) => {
        const value = row[colIndex];
        // Force ACC NUMBER to be treated as text to preserve leading zeros
        if (exportIndex === accNumberIndex && value !== null && value !== undefined) {
          return "'" + String(value); // Prefix with single quote to force text format
        }
        return value;
      })
    );
    
    if (filteredRows.length === 0) {
      return {
        success: false,
        error: invoiceNumber ? 
          `No invoices found with number: ${invoiceNumber}` : 
          `No invoices found in the last ${daysBack} days`,
        debug: {
          totalRows: dataRows.length,
          invoiceNumberCol: invoiceNumberCol,
          dateCol: dateCol,
          searchCriteria: invoiceNumber || `last ${daysBack} days`
        }
      };
    }
    
    // Create a temporary spreadsheet with just the filtered data
    const tempSpreadsheet = SpreadsheetApp.create(`TempInvoices_${Date.now()}`);
    const tempSheet = tempSpreadsheet.getActiveSheet();
    
    // Copy headers
    tempSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Copy filtered data
    if (filteredRows.length > 0) {
      tempSheet.getRange(2, 1, filteredRows.length, headers.length).setValues(filteredRows);
    }
    
    // Format the temporary sheet
    formatInvoiceSheet(tempSheet, headers.length, filteredRows.length + 1);
    
    // Auto-resize all columns to fit content
    for (let i = 1; i <= headers.length; i++) {
      tempSheet.autoResizeColumn(i);
    }
    
    // Set minimum column widths for readability
    const minWidths = {
      'Contractor': 150,
      'Work done': 200,
      'Total': 80,
      'Email': 200,
      'ACCOUNT NAME': 150,
      'ACC NUMBER': 120,
      'BANK': 100,
      'Playback Links': 300
    };
    
    headers.forEach((header, index) => {
      const colIndex = index + 1;
      const currentWidth = tempSheet.getColumnWidth(colIndex);
      const minWidth = minWidths[header] || 100;
      
      if (currentWidth < minWidth) {
        tempSheet.setColumnWidth(colIndex, minWidth);
      }
    });
    
    // Wait a moment for formatting to apply
    Utilities.sleep(1000);
    
    // Debug: log temp spreadsheet info
    Logger.log(`Temp spreadsheet has ${tempSheet.getLastRow()} rows and ${tempSheet.getLastColumn()} columns`);
    Logger.log(`Data range A1:${String.fromCharCode(64 + tempSheet.getLastColumn())}${tempSheet.getLastRow()}`);
    
    // Set print area to ensure all data is included
    const lastCol = String.fromCharCode(64 + headers.length);
    const printRange = `A1:${lastCol}${filteredRows.length + 1}`;
    tempSheet.getRange(printRange).activate();
    
    // Convert to PDF using built-in method
    const pdfBlob = tempSpreadsheet.getAs('application/pdf');
    
    // Create filename
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    const filename = invoiceNumber ? 
      `Invoice_${invoiceNumber}_${dateStr}.pdf` : 
      `Invoices_Recent_${dateStr}.pdf`;
    
    // Convert PDF to base64 data URL
    const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    const dataUrl = `data:application/pdf;base64,${base64Data}`;
    
    // Clean up temporary spreadsheet after PDF is created
    try {
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    } catch (cleanupError) {
      Logger.log('Warning: Could not clean up temp spreadsheet: ' + cleanupError.toString());
    }
    
    return {
      success: true,
      message: `PDF created successfully: ${filename}`,
      filename: filename,
      downloadUrl: dataUrl,
      rowsExported: filteredRows.length,
      debug: {
        headersLength: headers.length,
        filteredRowsLength: filteredRows.length,
        tempSpreadsheetId: tempSpreadsheet.getId()
      }
    };
    
  } catch (error) {
    Logger.log('Error in exportInvoicesPDF: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Format the invoice sheet for PDF export
 */
function formatInvoiceSheet(sheet, numCols, numRows) {
  // Set header formatting
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A90E2');
  headerRange.setFontColor('white');
  
  // Set borders for all data
  const dataRange = sheet.getRange(1, 1, numRows, numCols);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Auto-resize columns
  for (let i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Set alternating row colors
  for (let i = 2; i <= numRows; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, numCols).setBackground('#f9f9f9');
    }
  }
}

/**
 * Export the most recently created invoice as PDF
 * (Use this after creating invoices)
 */
function exportLatestInvoicePDF() {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const invoicingSheet = mainSheet.getSheetByName('Invoicing');
    
    if (!invoicingSheet) {
      throw new Error('Invoicing sheet not found');
    }
    
    const headerRow = 2;
    const lastRow = invoicingSheet.getLastRow();
    
    if (lastRow <= headerRow) {
      return {
        success: false,
        error: 'No invoice data found'
      };
    }
    
    // Get the most recent invoice by InvoiceNumber + Date combination
    const headers = invoicingSheet.getRange(headerRow, 1, 1, invoicingSheet.getLastColumn()).getValues()[0];
    const invoiceNumberCol = headers.indexOf('Invoice Number');
    const dateCol = headers.indexOf('Date');
    
    if (invoiceNumberCol === -1) {
      throw new Error('Invoice Number column not found');
    }
    if (dateCol === -1) {
      throw new Error('Date column not found');
    }
    
    const latestInvoiceNumber = invoicingSheet.getRange(lastRow, invoiceNumberCol + 1).getValue();
    const latestInvoiceDate = invoicingSheet.getRange(lastRow, dateCol + 1).getValue();
    
    // Export that specific invoice by number and date
    return exportInvoicesPDF(latestInvoiceNumber, null, latestInvoiceDate);
    
  } catch (error) {
    Logger.log('Error in exportLatestInvoicePDF: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}