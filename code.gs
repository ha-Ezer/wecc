/*
================================================================================
WEST END CITY CHURCH - CONTACT FORM BACKEND
Google Apps Script for Form Submission to Google Sheets
================================================================================

DEPLOYMENT INSTRUCTIONS:
========================

1. Open Google Sheets: https://docs.google.com/spreadsheets/d/1ZCLVRlWRXz-Gp4kiFJOR1XT6wt64KP4liVHp2mSrma0/edit

2. Click Extensions → Apps Script

3. Delete any existing code in Code.gs

4. Copy and paste this ENTIRE code.gs file

5. Save the project (Ctrl+S or Cmd+S)
   - Name it: "West End City Church Contact Form"

6. Deploy as Web App:
   - Click "Deploy" → "New deployment"
   - Click gear icon → Select "Web app"
   - Configuration:
     * Description: "Contact Form Backend v1"
     * Execute as: "Me (your email)"
     * Who has access: "Anyone"
   - Click "Deploy"
   - Authorize the app (click "Authorize access")
   - Review permissions and click "Allow"

7. Copy the Web App URL:
   - It will look like: https://script.google.com/macros/s/AKfycby.../exec
   - You MUST use this URL in the frontend code (index.html)

8. Test the deployment:
   - Click "Test deployments" to verify it works

9. Set up the sheet:
   - Ensure first row has headers: Name | Phone Number | Location | Date | Time
   - Run formatSheet() once manually to set initial formatting

10. Enable email notifications:
    - Ensure your Google account can send emails via Apps Script
    - Test by running sendErrorEmail() manually

IMPORTANT NOTES:
- Anyone with the Web App URL can submit data
- The URL must be kept in the frontend code
- Redeploy if you make code changes (Deploy → Manage deployments → Edit → Version: New version)
- Test thoroughly before going live

WEB APP URL PLACEHOLDER:
[COPY YOUR WEB APP URL HERE AND PASTE IN index.html]

================================================================================
*/

// Configuration
const SHEET_ID = '1ZCLVRlWRXz-Gp4kiFJOR1XT6wt64KP4liVHp2mSrma0';
const SHEET_NAME = 'Contacts'; // Use first sheet if this doesn't exist
const ERROR_EMAIL = 'ntimeben@gmail.com';
const MAX_ROWS = 1000; // Warning threshold (90% = 900 rows)

// Color Scheme
const COLORS = {
  headerBackground: '#B73231',    // Primary red
  headerText: '#FFFFFF',          // White
  oddRow: '#FAF9F6',             // Light cream
  evenRow: '#FFFFFF',            // White
  border: '#E0E0E0'              // Light gray
};

/**
 * Main POST handler - receives form submissions from frontend
 */
function doPost(e) {
  try {
    // Log that we received a request
    Logger.log('POST request received');
    Logger.log('Request type: ' + (e.postData ? e.postData.type : 'form'));
    
    // Parse data from POST request - check for form data first, then JSON
    let postData;
    try {
      // Check if this is form-encoded data (from HTML form submission)
      if (e && e.parameter && (e.parameter.name || e.parameter.phone || e.parameter.location)) {
        postData = {
          name: e.parameter.name || '',
          phone: e.parameter.phone || '',
          location: e.parameter.location || ''
        };
        Logger.log('Parsed as form data:', postData);
      }
      // Otherwise try to parse as JSON (from fetch/AJAX)
      else if (e && e.postData && e.postData.contents) {
        try {
          postData = JSON.parse(e.postData.contents);
          Logger.log('Parsed as JSON:', postData);
        } catch (jsonError) {
          // If JSON parse fails, try to parse as URL-encoded form data
          const contents = e.postData.contents;
          Logger.log('JSON parse failed, trying URL-encoded: ' + contents);
          const params = {};
          contents.split('&').forEach(param => {
            const [key, value] = param.split('=');
            if (key) {
              params[decodeURIComponent(key)] = decodeURIComponent(value || '').replace(/\+/g, ' ');
            }
          });
          postData = {
            name: params.name || '',
            phone: params.phone || '',
            location: params.location || ''
          };
          Logger.log('Parsed as URL-encoded form data:', postData);
        }
      } else {
        throw new Error('No data received in POST request');
      }
    } catch (parseError) {
      Logger.log('Error parsing request data: ' + parseError.toString());
      sendErrorEmail('Error parsing request: ' + parseError.toString(), 'Data Parse Error');
      return createCORSResponse({
        status: 'error',
        message: 'Invalid request format. Please try again.'
      });
    }
    
    // Validate required fields
    if (!postData.name || !postData.phone || !postData.location) {
      Logger.log('Missing required fields. Name: ' + (postData.name || 'missing') + ', Phone: ' + (postData.phone || 'missing') + ', Location: ' + (postData.location || 'missing'));
      return createCORSResponse({
        status: 'error',
        message: 'Missing required fields. Please fill in all fields.'
      });
    }
    
    // Get Ghana timezone timestamp
    const timestamp = getGhanaDateTime();
    Logger.log('Timestamp generated: ' + timestamp.date + ' ' + timestamp.time);
    
    // Get the spreadsheet
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(SHEET_ID);
      Logger.log('Spreadsheet opened successfully');
    } catch (error) {
      Logger.log('Error opening spreadsheet: ' + error.toString());
      sendErrorEmail('Error opening spreadsheet with ID ' + SHEET_ID + ': ' + error.toString(), 'Sheet Unavailable');
      return createCORSResponse({
        status: 'error',
        message: 'Unable to save your information. Please try again later.'
      });
    }
    
    // Get the sheet (try named sheet first, then first sheet)
    let sheet;
    try {
      sheet = spreadsheet.getSheetByName(SHEET_NAME);
      if (!sheet) {
        Logger.log('Sheet "' + SHEET_NAME + '" not found, using first sheet');
        sheet = spreadsheet.getSheets()[0];
      } else {
        Logger.log('Sheet "' + SHEET_NAME + '" found');
      }
    } catch (error) {
      Logger.log('Error getting sheet: ' + error.toString());
      try {
        sheet = spreadsheet.getSheets()[0];
        Logger.log('Using first sheet as fallback');
      } catch (fallbackError) {
        Logger.log('Error getting first sheet: ' + fallbackError.toString());
        sendErrorEmail('Error accessing sheet: ' + fallbackError.toString(), 'Sheet Unavailable');
        return createCORSResponse({
          status: 'error',
          message: 'Unable to save your information. Please try again later.'
        });
      }
    }
    
    if (!sheet) {
      Logger.log('Sheet is null or undefined');
      sendErrorEmail('Sheet not found', 'Sheet Unavailable');
      return createCORSResponse({
        status: 'error',
        message: 'Unable to save your information. Please try again later.'
      });
    }
    
    // Check sheet health before writing
    const healthCheck = checkSheetHealth(sheet);
    if (!healthCheck.isHealthy) {
      Logger.log('Sheet health check failed: ' + healthCheck.message);
      sendErrorEmail(healthCheck.message, healthCheck.errorType);
      return createCORSResponse({
        status: 'error',
        message: 'Unable to save your information. Please try again later.'
      });
    }
    
    // Prepare data row: [Name, Phone Number, Location, Date, Time]
    const rowData = [
      postData.name.trim(),
      postData.phone.trim(),
      postData.location.trim(),
      timestamp.date,
      timestamp.time
    ];
    Logger.log('Prepared row data: ' + JSON.stringify(rowData));
    
    // Append data to sheet
    try {
      sheet.appendRow(rowData);
      Logger.log('Data appended to sheet successfully');
    } catch (error) {
      Logger.log('Error appending row: ' + error.toString());
      sendErrorEmail('Error appending row: ' + error.toString(), 'Permission Denied');
      return createCORSResponse({
        status: 'error',
        message: 'Unable to save your information. Please try again later.'
      });
    }
    
    // Format the sheet
    try {
      formatSheet(sheet);
      Logger.log('Sheet formatting applied');
    } catch (error) {
      // Log formatting error but don't fail the submission
      Logger.log('Formatting error (non-critical): ' + error.toString());
    }
    
    // Return success response
    Logger.log('Submission successful!');
    return createCORSResponse({
      status: 'success',
      message: 'Thank you! Your information has been recorded.',
      timestamp: timestamp.date + ' ' + timestamp.time
    });
    
  } catch (error) {
    // Handle any unexpected errors
    Logger.log('Unexpected error in doPost: ' + error.toString());
    Logger.log('Error stack: ' + (error.stack || 'No stack trace'));
    sendErrorEmail('Unexpected error: ' + error.toString() + '\nStack: ' + (error.stack || 'No stack trace'), 'Unknown Error');
    return createCORSResponse({
      status: 'error',
      message: 'Unable to save your information. Please try again later.'
    });
  }
}

/**
 * Initialize the sheet with headers if they don't exist
 * Run this once manually from Apps Script editor to set up the sheet
 */
function initializeSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];
    
    if (!sheet) {
      // Create new sheet if it doesn't exist
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }
    
    // Check if headers already exist
    const firstRow = sheet.getRange(1, 1, 1, 5).getValues()[0];
    const hasHeaders = firstRow[0] && firstRow[0].toString().trim().length > 0;
    
    if (!hasHeaders) {
      // Set up column headers
      const headers = ['Name', 'Phone Number', 'Location', 'Date', 'Time'];
      sheet.getRange(1, 1, 1, 5).setValues([headers]);
      Logger.log('Sheet headers initialized successfully');
    } else {
      Logger.log('Sheet headers already exist');
    }
    
    // Apply formatting
    formatSheet(sheet);
    Logger.log('Sheet formatting applied successfully');
    
    return 'Sheet initialized successfully!';
    
  } catch (error) {
    Logger.log('Error initializing sheet: ' + error.toString());
    throw error;
  }
}

/**
 * Format the sheet with alternating row colors, borders, and sorting
 */
function formatSheet(sheet) {
  if (!sheet) {
    // Get sheet if not provided
    try {
      const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
      sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];
    } catch (error) {
      console.error('Could not get sheet for formatting:', error);
      return;
    }
  }
  
  try {
    // Get data range
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) return; // Only header row, no data to format
    
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Format header row (row 1)
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground(COLORS.headerBackground);
    headerRange.setFontColor(COLORS.headerText);
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    
    // Apply alternating row colors to data rows (row 2 onwards)
    for (let row = 2; row <= lastRow; row++) {
      const rowRange = sheet.getRange(row, 1, 1, lastCol);
      
      // Even row numbers (2, 4, 6...) get white, odd rows (3, 5, 7...) get cream
      // Note: row 2 is first data row (even), so it gets white
      if (row % 2 === 0) {
        rowRange.setBackground(COLORS.evenRow); // White
      } else {
        rowRange.setBackground(COLORS.oddRow); // Cream
      }
      
      // Set borders for all cells
      rowRange.setBorder(true, true, true, true, true, true, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    }
    
    // Auto-resize columns
    for (let col = 1; col <= lastCol; col++) {
      sheet.autoResizeColumn(col);
    }
    
    // Align columns
    // Name (A), Phone (B), Location (C) - left align
    sheet.getRange(2, 1, lastRow - 1, 3).setHorizontalAlignment('left');
    // Date (D), Time (E) - center align
    if (lastCol >= 4) {
      sheet.getRange(2, 4, lastRow - 1, Math.min(2, lastCol - 3)).setHorizontalAlignment('center');
    }
    
    // Sort by Date (column D) and Time (column E) - oldest first
    if (lastRow > 2 && lastCol >= 5) {
      const sortRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      sortRange.sort([
        {column: 4, ascending: true}, // Date column (ascending = oldest first)
        {column: 5, ascending: true}  // Time column (ascending = oldest first)
      ]);
    }
    
  } catch (error) {
    console.error('Error formatting sheet:', error);
    // Don't throw - formatting errors shouldn't break submissions
  }
}

/**
 * Send error notification email
 */
function sendErrorEmail(errorMessage, errorType) {
  try {
    // Use default values if parameters are not provided (for testing)
    const message = errorMessage || 'Test error message';
    const type = errorType || 'Test Error';
    
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`;
    
    // Get Ghana timezone timestamp
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Africa/Accra', 'yyyy-MM-dd HH:mm:ss');
    
    const subject = `West End City Church - ${type}`;
    const body = `Alert: Contact Form Error

Error Type: ${type}
Error Message: ${message}
Timestamp: ${timestamp}
Sheet URL: ${sheetUrl}

Please check the Google Sheet and resolve the issue.

This is an automated message from West End City Church Contact Form.`;
    
    // Try to send email
    try {
      MailApp.sendEmail({
        to: ERROR_EMAIL,
        subject: subject,
        body: body
      });
      Logger.log('Error email sent successfully');
    } catch (mailError) {
      // If MailApp fails, try GmailApp as alternative
      try {
        GmailApp.sendEmail(
          ERROR_EMAIL,
          subject,
          body
        );
        Logger.log('Error email sent via GmailApp successfully');
      } catch (gmailError) {
        Logger.log('Failed to send email via MailApp: ' + mailError.toString());
        Logger.log('Failed to send email via GmailApp: ' + gmailError.toString());
        throw gmailError;
      }
    }
    
  } catch (emailError) {
    Logger.log('Failed to send error email: ' + emailError.toString());
    // Log to console for debugging
    console.error('Email error details:', emailError);
    // Silently fail - don't break the main flow if email fails
  }
}

/**
 * Test function for sendErrorEmail - can be run from Apps Script editor
 */
function testSendErrorEmail() {
  sendErrorEmail('This is a test error message', 'Test Error');
  Logger.log('Test email function completed. Check logs for details.');
}

/**
 * Check sheet health - verify accessibility and row limits
 */
function checkSheetHealth(sheet) {
  try {
    if (!sheet) {
      return {
        isHealthy: false,
        message: 'Sheet is null or undefined',
        errorType: 'Sheet Unavailable'
      };
    }
    
    // Check if sheet is accessible
    try {
      const lastRow = sheet.getLastRow();
      const rowCount = Math.max(0, lastRow - 1); // Subtract header row
      
      // Check if approaching row limit (90% of MAX_ROWS)
      const warningThreshold = MAX_ROWS * 0.9;
      if (rowCount >= warningThreshold) {
        sendErrorEmail(
          `Sheet is approaching row limit. Current rows: ${rowCount}/${MAX_ROWS}`,
          'Row Limit Reached'
        );
        // Still allow submissions, just warn
      }
      
      // Check if at or over limit
      if (rowCount >= MAX_ROWS) {
        return {
          isHealthy: false,
          message: `Sheet has reached maximum capacity (${MAX_ROWS} rows)`,
          errorType: 'Row Limit Reached'
        };
      }
      
      return {
        isHealthy: true,
        message: 'Sheet is healthy',
        rowCount: rowCount
      };
      
    } catch (error) {
      return {
        isHealthy: false,
        message: `Cannot access sheet: ${error.toString()}`,
        errorType: 'Sheet Unavailable'
      };
    }
    
  } catch (error) {
    return {
      isHealthy: false,
      message: `Health check failed: ${error.toString()}`,
      errorType: 'Unknown Error'
    };
  }
}

/**
 * Get current date and time in Ghana timezone (Africa/Accra, GMT+0)
 * Returns object with separate date and time strings
 */
function getGhanaDateTime() {
  const now = new Date();
  
  // Format date and time in Ghana timezone
  const dateStr = Utilities.formatDate(now, 'Africa/Accra', 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, 'Africa/Accra', 'HH:mm:ss');
  
  return {
    date: dateStr,      // Format: YYYY-MM-DD
    time: timeStr       // Format: HH:MM:SS (24-hour)
  };
}

/**
 * Create CORS-enabled JSON response
 */
function createCORSResponse(content) {
  const output = ContentService.createTextOutput(JSON.stringify(content));
  output.setMimeType(ContentService.MimeType.JSON);
  
  // Note: Google Apps Script Web Apps automatically handle CORS for GET/POST
  // Headers are set automatically, so we don't need to set them manually
  // The Web App deployment settings handle CORS
  
  return output;
}

/**
 * Handle OPTIONS requests for CORS preflight
 */
function doGet(e) {
  return createCORSResponse({
    status: 'ok',
    message: 'West End City Church Contact Form API'
  });
}
