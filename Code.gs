
// Main function to generate weekly leaderboard
function createWeeklyLeaderboard() {
  const startTime = new Date();
  const debugInfo = [];
  let leaderboard = [];
  
  try {
    // 1. SETUP AND DATA COLLECTION
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Form Responses 1");
    
    if (!sheet) {
      throw new Error("Sheet 'Form Responses 1' not found. Please check your spreadsheet.");
    }
    
    const timeZone = Session.getScriptTimeZone();
    const userEmail = Session.getActiveUser().getEmail();
    
    debugInfo.push(`Script started at: ${startTime}`);
    debugInfo.push(`Processing sheet: ${sheet.getName()}`);
    debugInfo.push(`User email: ${userEmail}`);
    
    // Calculate current week range (Monday to Sunday)
    const {weekStart, weekEnd, weekRangeStr} = getCurrentWeekRange(timeZone);
    debugInfo.push(`Week range: ${weekRangeStr}`);
    debugInfo.push(`Start: ${weekStart} | End: ${weekEnd}`);

    // 2. DATA PROCESSING WITH VALIDATION
    const data = sheet.getDataRange().getValues();
    debugInfo.push(`Total rows found: ${data.length}`);
    
    // Process data with strict validation
    const processingResult = processReferralData(data, weekStart, weekEnd);
    leaderboard = processingResult.leaderboard;
    
    debugInfo.push(`Valid rows processed: ${processingResult.validRows}`);
    debugInfo.push(`Staff with deliveries: ${leaderboard.length}`);
    debugInfo.push(`Top performer: ${leaderboard[0]?.name || 'None'} with ${leaderboard[0]?.count || 0} deliveries`);

    // 3. OUTPUT GENERATION
    if (leaderboard.length === 0) {
      // Special handling for empty leaderboard
      const errorContent = generateErrorEmail(weekRangeStr, processingResult, data);
      MailApp.sendEmail({
        to: userEmail,
        subject: "EPC Leaderboard - No Data Found",
        htmlBody: errorContent
      });
      debugInfo.push("Sent empty data notification");
    } else {
      // Generate normal leaderboard with exact design
      const emailContent = generateLeaderboardEmail(leaderboard, weekRangeStr);
      MailApp.sendEmail({
        to: userEmail,
        subject: `EPC Referral Leaderboard - Week of ${Utilities.formatDate(weekStart, timeZone, "MMM d")}`,
        htmlBody: emailContent
      });
      debugInfo.push("Sent leaderboard email");
      
      // Update spreadsheet
      updateLeaderboardSheet(ss, leaderboard, weekRangeStr);
      debugInfo.push("Updated leaderboard sheet");
    }

  } catch (error) {
    // Comprehensive error handling
    debugInfo.push(`ERROR: ${error.message}`);
    debugInfo.push(`Stack: ${error.stack}`);
    
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "EPC Leaderboard Processing Error",
      body: debugInfo.join('\n\n')
    });
  } finally {
    // Log execution details
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    debugInfo.push(`Script completed in ${duration} seconds`);
    Logger.log(debugInfo.join('\n'));
  }
}

// Helper Functions

function getCurrentWeekRange(timeZone) {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1));
  weekStart.setHours(0, 0, 0, 0);
  
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  weekEnd.setHours(23, 59, 59, 999);
  
  const dateFormat = "MMM d, yyyy";
  const weekRangeStr = `${Utilities.formatDate(weekStart, timeZone, dateFormat)} to ${Utilities.formatDate(weekEnd, timeZone, dateFormat)}`;
  
  return {weekStart, weekEnd, weekRangeStr};
}

function processReferralData(data, weekStart, weekEnd) {
  const staffCounts = {};
  let validRows = 0;
  let skippedRows = 0;
  const sampleData = [];
  const statusVariations = new Set();

  for (let i = 1; i < data.length; i++) {
    try {
      const row = data[i];
      
      // Parse timestamp in format "6/25/2025 20:32:27"
      const timestampStr = row[0]; // Column A - timestamp
      const timestamp = new Date(timestampStr);
      
      if (isNaN(timestamp.getTime())) {
        Logger.log(`Invalid timestamp in row ${i+1}: ${timestampStr}`);
        skippedRows++;
        continue;
      }
      
      // Column B - staff name
      const staffName = row[1]?.toString().trim();
      if (!staffName) {
        Logger.log(`Missing staff name in row ${i+1}`);
        skippedRows++;
        continue;
      }
      
      // Column J - status
      const status = row[9]?.toString().trim().toLowerCase();
      statusVariations.add(status);
      
      // Check if record should be counted
      const isInWeek = timestamp >= weekStart && timestamp <= weekEnd;
      const isDelivered = status === "delivered";
      
      if (isInWeek && isDelivered) {
        staffCounts[staffName] = (staffCounts[staffName] || 0) + 1;
        validRows++;
        
        // Collect sample data for debugging
        if (validRows <= 5) {
          sampleData.push({
            name: staffName,
            date: timestamp,
            status: row[9]?.toString().trim()
          });
        }
      }
    } catch (e) {
      skippedRows++;
      Logger.log(`Error processing row ${i+1}: ${e.message}`);
    }
  }

  return {
    leaderboard: Object.keys(staffCounts).map(name => ({
      name: name,
      count: staffCounts[name],
      initial: name.charAt(0).toUpperCase()
    })).sort((a, b) => b.count - a.count),
    validRows,
    skippedRows,
    statusVariations: Array.from(statusVariations),
    sampleData
  };
}

function generateLeaderboardEmail(leaderboard, weekRangeStr) {
  const topThree = leaderboard.slice(0, 3);
  
  return `
  <html>
  <head>
    <style>
      body {
        font-family: 'Roboto', sans-serif;
        margin: 0;
        padding: 20px;
        background-color: #f5f9fc;
      }
      .container {
        max-width: 900px;
        margin: 0 auto;
        background: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0, 163, 224, 0.1);
      }
      .header {
        text-align: center;
        margin-bottom: 30px;
        padding: 25px 20px;
        border-radius: 8px;
        background: linear-gradient(135deg, #00a3e0, #0078a5);
        box-shadow: 0 4px 12px rgba(0, 163, 224, 0.2);
      }
      .header h1 {
        color: white;
        font-weight: 500;
        margin-bottom: 5px;
      }
      .header h3 {
        color: white;
        font-weight: 400;
        margin-top: 0;
      }
      .row {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        align-items: flex-end;
        margin: 0 -10px;
      }
      .col {
        flex: 1;
        min-width: 200px;
        padding: 10px;
        box-sizing: border-box;
      }
      .d-flex {
        display: flex;
      }
      .align-center {
        align-items: center;
      }
      .justify-center {
        justify-content: center;
      }
      .v-avatar {
          border-radius: 50%;
          display: flex;
          justify-content: center;
          align-items: center;
          border: 3px solid #fff;
          box-shadow: 0 3px 10px rgba(0, 163, 224, 0.3);
          margin: 0 auto;
          aspect-ratio: 1/1;
          text-align: center;
          line-height: 1;
          display: flex;
          justify-content: center;
          align-items: center;
          }

      .v-avatar span {
        display: block;
        width: 100%;
        text-align: center;
        line-height: 150px;
        margin: 0;
        padding: 0;
        position: relative;
        top: 0.1em;
        }    

      .blue-gradient {
        background: linear-gradient(135deg, #00a3e0, #0078a5) !important;
      }
      .white--text {
        color: white !important;
      }
      .display-3 {
        font-size: 48px !important;
        font-weight: 300;
        line-height: 1;
      }
      .text-center {
        text-align: center !important;
      }
      .text--secondary {
        color: #6b7c8a !important;
      }
      .mt-3 {
        margin-top: 12px !important;
      }
      .mt-5 {
        margin-top: 20px !important;
      }
      .ml-5 {
        margin-left: 20px !important;
      }
      .ml-n5 {
        margin-left: -20px !important;
      }
      .mt-n15 {
        margin-top: -60px !important;
      }
      .bringfront {
        z-index: 10;
        position: relative;
      }
      .gold--text {
        color: #ffc107 !important;
      }
      .silver--text {
        color: #c0c0c0 !important;
      }
      .bronze--text {
        color: #cd7f32 !important;
      }
      .border {
        border: 3px solid #fff !important;
      }
      .leaderboard-table {
        width: 100%;
        border-collapse: collapse;
        margin: 40px 0;
        border-radius: 8px;
        overflow: hidden;
      }
      .leaderboard-table th {
        background: linear-gradient(135deg, #00a3e0, #0078a5);
        color: white;
        padding: 15px;
        text-align: left;
        font-weight: 500;
      }
      .leaderboard-table td {
        padding: 12px 15px;
        border-bottom: 1px solid #e0f2fb;
      }
      .leaderboard-table tr:nth-child(even) {
        background-color: #f5f9fc;
      }
      .leaderboard-table tr:hover {
        background-color: #e0f2fb;
      }
      .footer {
        margin-top: 30px;
        text-align: center;
        color: #6b7c8a;
        font-size: 12px;
        padding-top: 20px;
        border-top: 1px solid #e0f2fb;
      }
      h1 {
        font-size: 24px;
        font-weight: 500;
        margin: 0;
        color: #0078a5;
      }
      h4 {
        font-size: 16px;
        font-weight: 400;
        margin: 4px 0;
      }
      i.mdi {
        display: inline-block;
        font-style: normal;
      }
      .rank-1 {
        color: #ffc107;
        font-weight: 500;
      }
      .rank-2 {
        color: #c0c0c0;
        font-weight: 500;
      }
      .rank-3 {
        color: #cd7f32;
        font-weight: 500;
      }
      .medal {
        width: 24px;
        height: 24px;
        display: inline-block;
        margin-right: 8px;
        vertical-align: middle;
      }
    </style>
    <link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500&display=swap" rel="stylesheet">
    <link href="https://cdn.materialdesignicons.com/5.3.45/css/materialdesignicons.min.css" rel="stylesheet">
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>EPC Staff Referral Leaderboard</h1>
        <h3>Week of ${weekRangeStr}</h3>
      </div>
      
      <div class="row fill-height minheight">
        <!-- 2nd Place -->
        <div class="col col-4">
          <div class="d-flex align-center justify-center">
            <div>
              <div class="v-avatar border mt-5 ml-5 blue-gradient" style="height: 130px; min-width: 130px; width: 130px;">
                <span class="white--text display-3">${topThree[1]?.initial || ''}</span>
              </div>
              <h1 class="mt-3 ml-5 text-center silver--text">2nd</h1>
              <h4 class="ml-5 text-center text--secondary">${topThree[1]?.count || '0'}</h4>
              <h4 class="ml-5 text-center text--secondary">Referrals</h4>
            </div>
          </div>
        </div>
        
        <!-- 1st Place -->
        <div class="col col-4">
          <div class="d-flex align-center justify-center">
            <div class="mt-n15">
              <div class="d-flex justify-center">
                <i aria-hidden="true" class="mdi mdi-crown gold--text" style="font-size: 40px;"></i>
              </div>
              <div class="v-avatar border bringfront blue-gradient" style="height: 150px; min-width: 150px; width: 150px;">
                <span class="white--text display-3">${topThree[0]?.initial || ''}</span>
              </div>
              <h1 class="mt-3 text-center gold--text">1st</h1>
              <h4 class="text-center text--secondary">${topThree[0]?.count || '0'}</h4>
              <h4 class="ml-n1 text-center text--secondary">Referrals</h4>
            </div>
          </div>
        </div>
        
        <!-- 3rd Place -->
        <div class="col col-4">
          <div class="d-flex align-center justify-center">
            <div>
              <div class="v-avatar border ml-n5 mt-5 blue-gradient" style="height: 125px; min-width: 125px; width: 125px;">
                <span class="white--text display-3">${topThree[2]?.initial || ''}</span>
              </div>
              <h1 class="mt-3 ml-n5 text-center bronze--text">3rd</h1>
              <h4 class="ml-n5 text-center text--secondary">${topThree[2]?.count || '0'}</h4>
              <h4 class="ml-n5 text-center text--secondary">Referrals</h4>
            </div>
          </div>
        </div>
      </div>
      
      <table class="leaderboard-table">
        <tr>
          <th>Rank</th>
          <th>Staff Name</th>
          <th>Delivered EPCs</th>
        </tr>
        ${leaderboard.map((staff, index) => `
          <tr>
            <td>
              ${index === 0 ? '<span class="medal">🥇</span>' : ''}
              ${index === 1 ? '<span class="medal">🥈</span>' : ''}
              ${index === 2 ? '<span class="medal">🥉</span>' : ''}
              ${index > 2 ? index + 1 : ''}
            </td>
            <td class="${index === 0 ? 'rank-1' : ''} ${index === 1 ? 'rank-2' : ''} ${index === 2 ? 'rank-3' : ''}">
              ${staff.name}
            </td>
            <td>${staff.count}</td>
          </tr>
        `).join('')}
      </table>
      
      <div class="footer">
        <p>This leaderboard resets automatically every Monday morning</p>
        <p>Generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy 'at' h:mm a")}</p>
      </div>
    </div>
  </body>
  </html>`;
}

function generateErrorEmail(weekRangeStr, processingResult, rawData) {
  const sampleTimestamps = rawData.slice(1, 6).map(row => {
    try {
      return new Date(row[0]).toString() + ` (Original: ${row[0]})`;
    } catch (e) {
      return `Invalid date: ${row[0]}`;
    }
  }).join('<br>');

  return `
  <html>
  <head>
    <style>
      body { 
        font-family: 'Roboto', sans-serif; 
        line-height: 1.6; 
        max-width: 800px; 
        margin: 0 auto; 
        padding: 20px;
        background-color: #f5f9fc;
      }
      .error-container {
        background: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0, 163, 224, 0.1);
      }
      h1 { 
        color: #d32f2f;
        font-weight: 500;
        margin-top: 0;
      }
      .debug-info { 
        background: #f5f9fc; 
        padding: 20px; 
        border-radius: 8px; 
        margin: 20px 0; 
        border-left: 4px solid #00a3e0;
      }
      strong {
        color: #0078a5;
      }
      ul {
        padding-left: 20px;
      }
      li {
        margin-bottom: 8px;
      }
    </style>
    <link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500&display=swap" rel="stylesheet">
  </head>
  <body>
    <div class="error-container">
      <h1>No Referral Data Found</h1>
      <p>No delivered referrals were found for the week of <strong>${weekRangeStr}</strong>.</p>
      
      <div class="debug-info">
        <h3 style="color: #0078a5; margin-top: 0;">Debug Information</h3>
        <p><strong>Total rows processed:</strong> ${rawData.length - 1}</p>
        <p><strong>Valid delivered referrals:</strong> ${processingResult.validRows}</p>
        <p><strong>Status variations found:</strong> ${processingResult.statusVariations.join(', ')}</p>
        
        <h4 style="color: #0078a5;">Sample Matching Records:</h4>
        ${processingResult.sampleData.length > 0 ? 
          processingResult.sampleData.map(item => `
            <p>${item.name} - ${item.date} - ${item.status}</p>
          `).join('') : 
          '<p>No matching records found</p>'
        }
        
        <h4 style="color: #0078a5;">First 5 Timestamps (Column A):</h4>
        <p>${sampleTimestamps}</p>
      </div>
      
      <p>Please verify:</p>
      <ul>
        <li>Column A contains dates in format "M/D/YYYY H:mm:ss"</li>
        <li>Column B contains staff names</li>
        <li>Column J contains "Delivered" status (case insensitive)</li>
        <li>Dates fall within ${weekRangeStr}</li>
      </ul>
    </div>
  </body>
  </html>`;
}

function updateLeaderboardSheet(ss, leaderboard, weekRangeStr) {
  let leaderboardSheet = ss.getSheetByName("Leaderboard");
  
  if (!leaderboardSheet) {
    leaderboardSheet = ss.insertSheet("Leaderboard");
    leaderboardSheet.setFrozenRows(1);
    leaderboardSheet.getRange("A1:C1")
      .setBackground("#00a3e0")
      .setFontColor("white")
      .setFontWeight("bold");
  } else {
    leaderboardSheet.clear();
  }
  
  // Set headers
  leaderboardSheet.getRange("A1:C1").setValues([["Rank", "Staff Name", "Delivered Referrals"]]);
  
  // Add week info
  leaderboardSheet.getRange("E1:F1").setValues([["Week:", weekRangeStr]]);
  leaderboardSheet.getRange("E1").setFontWeight("bold");
  
  // Populate data
  if (leaderboard.length > 0) {
    const data = leaderboard.map((staff, index) => [index + 1, staff.name, staff.count]);
    leaderboardSheet.getRange(2, 1, data.length, 3).setValues(data);
    
    // Format top 3 with colors
    const top3Colors = ["#FFD700", "#C0C0C0", "#CD7F32"]; // Gold, Silver, Bronze
    top3Colors.forEach((color, index) => {
      if (leaderboard.length > index) {
        leaderboardSheet.getRange(index + 2, 1, 1, 3).setBackground(color);
      }
    });
  }
  
  // Auto-resize columns
  leaderboardSheet.autoResizeColumns(1, 3);
}

function setupWeeklyTrigger() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers()
    .forEach(trigger => {
      if (trigger.getHandlerFunction() === "createWeeklyLeaderboard") {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  
  // Create new trigger for Monday 8 AM
  ScriptApp.newTrigger("createWeeklyLeaderboard")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  
  Logger.log("Weekly trigger set for Monday at 8:00 AM");
}