# Staff Pay Calculator - Google Apps Script

## Claude Assistant Instructions
When asked to "push" or "deploy", execute these commands in order:
```bash
# 1. Push to Google Apps Script
clasp push

# 2. Deploy the web app
clasp deploy --deploymentId AKfycbxbH2KekQ2ibY85WygitXa6B3IZmSoMq2wORi-Cf9OU1EXNBjWdOecRINOElkpIzEVw

# 3. Commit and push to Git
git add .
git commit -m "Update: [describe changes]"
git push
```

## Project Overview
Google Apps Script web application that automates staff payment calculations by reading unpaid work from an external spreadsheet, calculating payments based on configurable rates, creating invoices, and marking work as invoiced.

## Deployment Information

### Primary Web App Deployment
- **Deployment ID**: `AKfycbxbH2KekQ2ibY85WygitXa6B3IZmSoMq2wORi-Cf9OU1EXNBjWdOecRINOElkpIzEVw`
- **Web App URL**: `https://script.google.com/macros/s/AKfycbxbH2KekQ2ibY85WygitXa6B3IZmSoMq2wORi-Cf9OU1EXNBjWdOecRINOElkpIzEVw/exec`
- **Script ID**: `1Fya4c-_NCDInNN3t-6eqViizQRKaN-6capqnkHTk7CLvB6xV4ckcU70X`

### Deployment Commands

#### Complete Update Workflow (Simple Method)
```bash
# 1. Push code changes to Google Apps Script
clasp push

# 2. Update the existing web app deployment
clasp deploy --deploymentId AKfycbxbH2KekQ2ibY85WygitXa6B3IZmSoMq2wORi-Cf9OU1EXNBjWdOecRINOElkpIzEVw

# That's it! The URL stays the same.
```

#### ⚠️ Important: Web App Configuration
- **Required**: The `appsscript.json` must include webapp configuration:
  ```json
  {
    "webapp": {
      "access": "ANYONE",
      "executeAs": "USER_DEPLOYING"
    }
  }
  ```
- **With this configuration**: `clasp deploy --deploymentId` works perfectly
- **Same URL**: The deployment URL remains constant across updates

#### Individual Commands
```bash
# Push code only (updates script but not deployment)
clasp push

# Update deployment only (uses current code version)
clasp deploy --deploymentId AKfycbzvsnL2dgnQa4ZQW0oJjjE4BKW-dDd2-IGdDbCj9ana04tz6WEwX2HC6jp2KMVARNQV --description "Updated version"
```

#### List All Deployments
```bash
clasp deployments
```

## API Endpoints

### Web Interface
- **URL**: `https://script.google.com/macros/s/AKfycbzvsnL2dgnQa4ZQW0oJjjE4BKW-dDd2-IGdDbCj9ana04tz6WEwX2HC6jp2KMVARNQV/exec`
- **Description**: Interactive web interface with buttons for preview, execute, status, and debug

### API Endpoints
- **Preview**: `?action=preview` - Returns payment calculation without creating invoices
- **Execute**: `?action=calculatePay` - Calculates payments, creates invoices, and returns invoice metadata
- **Status**: `?action=getStatus` - Returns unpaid task count
- **Debug**: `?action=getDebugLog` - Returns debug information
- **Export Latest PDF**: `?action=exportLatestInvoice` - Exports most recent invoice as PDF
- **Export Specific PDF**: `?action=exportInvoicePDF&invoiceNumber=XXX` - Exports specific invoice by number

## Core Functions

### API Functions (work without UI)
- `calculateStaffPay()` - Returns payment calculation results as JSON
- `createInvoice(payments)` - Creates invoice rows with single invoice number
- `createInvoicesAndMark(workLogData, payments)` - Creates invoice and marks work as paid
- `getUnpaidWorkFromMaster()` - Gets unpaid work data from external sheet
- `getPayConfiguration()` - Gets pay rates from Pay Config sheet
- `getStaffMapping()` - Gets staff name mappings

### Google Sheets UI Functions
- `calculateStaffPayUI()` - Shows interactive dialog (used by menu)
- `createInvoicesAndMarkUI()` - UI callback for invoice creation

## Configuration

### External Spreadsheet
- **Work Log Sheet ID**: `17BzCsrHTQQi4e1hg59_AmP9PtiKWdkKCjbAvu5_EjO0`
- **Sheet Name**: `MASTER`

### Local Sheets Required
- **Pay Config**: Contains task types and pay rates
- **Staff Key to Staff name**: Maps staff keys to legal names  
- **Invoicing**: Where invoice rows are created

## Recent Updates

### Invoice Metadata Response Enhancement (Latest)
- Updated `handleCalculatePayRequest()` to return invoice metadata in response
- Added `invoiceInfo` field containing:
  - `invoiceNumber`: Generated invoice identifier (e.g., "INV-20250720-234")
  - `invoiceDate`: ISO timestamp when invoice was created
  - `rowsCreated`: Number of staff invoice rows created
- Enables PlayBack Game Review frontend to show success dialogs and targeted PDF export
- Maintains backward compatibility with existing integrations

### Invoice Creation Fix
- Function renamed from `createInvoices` to `createInvoice` (singular)
- All rows created in single invocation now share the same invoice number
- Fixed to create exactly the right number of rows (no extra empty rows)
- Formula copying preserved from previous row

### Web App Interface
- Added responsive HTML interface accessible without parameters
- Separated UI functions from core API functions
- Added comprehensive API documentation in web interface
- Fixed CORS header issues for Google Apps Script

## Testing

### Manual Testing via Web Interface
1. Visit the web app URL
2. Use "Preview" to see calculation without execution
3. Use "Execute" to create invoices and mark work as paid

### Direct Function Testing
```javascript
// Test payment calculation
const result = calculateStaffPay();
console.log(result);

// Test invoice creation (if result successful)
if (result.success) {
  createInvoice(result.payments);
}
```

## Troubleshooting

### Common Issues
1. **No unpaid work found**: Check debug log via `?action=getDebugLog`
2. **Permission errors**: Ensure web app is deployed with "Execute as: Me" and "Who has access: Anyone"
3. **Formula copying issues**: Verify there's a previous row with formulas to copy from

### Debug Information
- Access debug logs via web interface or `?action=getDebugLog`
- Debug logs stored in Script Properties and show work detection process
- Invoice creation includes row count and formula copying status