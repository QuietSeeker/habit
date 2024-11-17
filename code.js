// Global constants
const SHEET_NAMES = {
  USERS: 'User Habits',
  MONTHLY_TRACKING: 'Monthly Tracking',
  SUMMARY: 'Summary View'
};

const STATUS = {
  COMPLETE: '✓',
  INCOMPLETE: '✗',
  EXEMPT: 'E',
  EMPTY: '-'  // Changed from '' to '-' for better visibility
};

const CHARGE_AMOUNT = 3;
const COLORS = {
  COMPLETE: '#93C47D',  // Green
  INCOMPLETE: '#E06666', // Red
  EXEMPT: '#6FA8DC',    // Blue
  HEADER: '#344261',    // Dark Blue
  EMPTY: '#FFFFFF'      // White
};

// Create menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Habit Tracker')
    .addItem('Initial Setup', 'setupSpreadsheet')
    .addItem('Setup New User', 'setupNewUser')
    .addItem('Create Monthly Sheet', 'createMonthlySheet')
    .addItem('Update Summary', 'updateSummary')
    .addToUi();
}

// Enhanced setupMonthlySheet function
function setupMonthlySheet(sheet, users, currentDate) {
  const daysInMonth = new Date(currentDate.getYear(), currentDate.getMonth() + 1, 0).getDate();

  // Create header row with dates
  const headerRow = [`Price: £${CHARGE_AMOUNT}`, 'Name/Date'];
  for (let day = 1; day <= daysInMonth; day++) {
    const dayDate = new Date(currentDate.getYear(), currentDate.getMonth(), day);
    const dayStr = Utilities.formatDate(dayDate, Session.getScriptTimeZone(), 'MMM/d\nEEE');
    headerRow.push(dayStr);
  }
  headerRow.push('Total');

  // Set headers
  const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
  headerRange.setValues([headerRow]);
  headerRange.setBackground(COLORS.HEADER);
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setWrap(true);  // Enable text wrapping for dates

  // Calculate total rows needed
  const totalRows = users.length * 5;

  // Create all rows at once
  const allRows = [];
  users.forEach(user => {
    const userName = user[0];
    for (let i = 1; i <= 5; i++) {
      const habit = user[i];
      const row = [userName, habit];
      // Add empty cells with default status
      for (let day = 1; day <= daysInMonth; day++) {
        row.push(STATUS.EMPTY);
      }
      // Add total formula
      const rowIndex = allRows.length + 2;
      row.push(`=COUNTIF(C${rowIndex}:${String.fromCharCode(66 + daysInMonth)}${rowIndex},"${STATUS.COMPLETE}")`);
      allRows.push(row);
    }
  });

  // Set all rows at once
  if (allRows.length > 0) {
    const dataRange = sheet.getRange(2, 1, allRows.length, headerRow.length);
    dataRange.setValues(allRows);

    // Set formatting for the entire data range
    dataRange.setFontFamily('Arial');
    dataRange.setFontSize(10);
    dataRange.setHorizontalAlignment('center');

    // Set up data validation for status cells
    const validationRange = sheet.getRange(2, 3, allRows.length, daysInMonth);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([STATUS.COMPLETE, STATUS.INCOMPLETE, STATUS.EXEMPT, STATUS.EMPTY], true)
      .build();
    validationRange.setDataValidation(rule);

    // Set conditional formatting
    setConditionalFormatting(sheet, 2, 3, allRows.length, daysInMonth);
  }

  // Freeze headers and names
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Auto-size columns
  sheet.autoResizeColumns(1, headerRow.length);

  // Set column widths for better visibility
  for (let col = 3; col <= daysInMonth + 2; col++) {
    sheet.setColumnWidth(col, 35);
  }
}

// Enhanced setConditionalFormatting function
function setConditionalFormatting(sheet, startRow, startCol, numRows, numCols) {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);

  // Clear existing rules
  sheet.clearConditionalFormatRules();

  // Add new rules with improved formatting
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.COMPLETE)
      .setBackground(COLORS.COMPLETE)
      .setFontColor('#000000')
      .setRanges([range])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.INCOMPLETE)
      .setBackground(COLORS.INCOMPLETE)
      .setFontColor('#FFFFFF')
      .setRanges([range])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.EXEMPT)
      .setBackground(COLORS.EXEMPT)
      .setFontColor('#000000')
      .setRanges([range])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS.EMPTY)
      .setBackground(COLORS.EMPTY)
      .setRanges([range])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}

// Setup initial spreadsheet structure
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Create Users sheet
  let usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_NAMES.USERS);
    usersSheet.getRange('A1:G1').setValues([['User', 'Habit 1', 'Habit 2', 'Habit 3', 'Habit 4', 'Habit 5', 'Total Charges (£)']]);

    // Format header
    const headerRange = usersSheet.getRange('A1:G1');
    headerRange.setBackground(COLORS.HEADER);
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    usersSheet.setFrozenRows(1);
  }

  // Create Summary sheet
  let summarySheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SHEET_NAMES.SUMMARY);
    summarySheet.getRange('A1:D1').setValues([['User', 'Month', 'Completion Rate', 'Charges']]);

    // Format header
    const headerRange = summarySheet.getRange('A1:D1');
    headerRange.setBackground(COLORS.HEADER);
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    summarySheet.setFrozenRows(1);
  }

  ui.alert('Setup Complete', 'Basic sheets have been initialized. Use "Create Monthly Sheet" to create tracking sheet for current month.', ui.ButtonSet.OK);
}

// Create new user setup
function setupNewUser() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

  if (!usersSheet) {
    ui.alert('Error', 'Users sheet not found. Please run Initial Setup first.', ui.ButtonSet.OK);
    return;
  }

  // Prompt for user name
  const userResponse = ui.prompt(
    'New User Setup',
    'Enter user name:',
    ui.ButtonSet.OK_CANCEL
  );

  if (userResponse.getSelectedButton() !== ui.Button.OK) return;
  const userName = userResponse.getResponseText().trim();

  if (!userName) {
    ui.alert('Error', 'Please enter a valid user name.', ui.ButtonSet.OK);
    return;
  }

  // Check if user already exists
  const existingUsers = usersSheet.getRange(2, 1, usersSheet.getLastRow() || 1, 1).getValues();
  if (existingUsers.flat().includes(userName)) {
    ui.alert('Error', 'This user already exists.', ui.ButtonSet.OK);
    return;
  }

  // Array to store habits
  const habits = [];

  // Prompt for each habit
  for (let i = 1; i <= 5; i++) {
    const habitResponse = ui.prompt(
      'Habit Setup',
      `Enter Habit ${i} for ${userName}:`,
      ui.ButtonSet.OK_CANCEL
    );

    if (habitResponse.getSelectedButton() !== ui.Button.OK) return;
    const habit = habitResponse.getResponseText().trim();

    if (!habit) {
      ui.alert('Error', `Please enter a valid habit ${i}.`, ui.ButtonSet.OK);
      return;
    }

    habits.push(habit);
  }

  // Add new user data
  const newRow = [userName, ...habits, 0]; // Initialize charges as 0
  const newRowRange = usersSheet.getRange(usersSheet.getLastRow() + 1, 1, 1, 7);
  newRowRange.setValues([newRow]);

  // Format the new row
  newRowRange.setFontFamily('Arial');
  newRowRange.setFontSize(10);
  newRowRange.setHorizontalAlignment('center');

  // Format the charges cell
  const chargesCell = newRowRange.getCell(1, 7);
  chargesCell.setNumberFormat('£#,##0.00');

  // Auto-resize columns
  usersSheet.autoResizeColumns(1, 7);

  ui.alert(
    'Success',
    `User "${userName}" has been added with ${habits.length} habits.\nUse "Create Monthly Sheet" to update tracking sheets.`,
    ui.ButtonSet.OK
  );

  // Optionally update existing monthly sheets
  const updateResponse = ui.alert(
    'Update Existing Sheets',
    'Would you like to update existing monthly tracking sheets with this new user?',
    ui.ButtonSet.YES_NO
  );

  if (updateResponse === ui.Button.YES) {
    updateExistingSheets(userName, habits);
  }
}

// Create monthly tracking sheet
function createMonthlySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get current month and year
  const today = new Date();
  const monthYear = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMyy");
  const sheetName = `Tracking ${monthYear}`;

  // Check if sheet already exists
  let trackingSheet = ss.getSheetByName(sheetName);
  if (trackingSheet) {
    const response = ui.alert(
      'Sheet Exists',
      'Sheet for this month already exists. Would you like to reset it?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      ss.deleteSheet(trackingSheet);
    } else {
      return;
    }
  }

  // Create new monthly tracking sheet
  trackingSheet = ss.insertSheet(sheetName);

  // Get users and their habits
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!usersSheet || usersSheet.getLastRow() < 2) {
    ui.alert('Error', 'No users found. Please add users first.', ui.ButtonSet.OK);
    return;
  }

  const users = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 6).getValues();

  // Setup monthly sheet
  setupMonthlySheet(trackingSheet, users, today);

  // Initialize charges calculation
  setupChargesCalculation(trackingSheet, usersSheet);

  ui.alert('Success', `Monthly tracking sheet for ${monthYear} has been created!`, ui.ButtonSet.OK);
}

// Setup charges calculation
function setupChargesCalculation(trackingSheet, usersSheet) {
  const lastRow = trackingSheet.getLastRow();
  const lastCol = trackingSheet.getLastColumn();

  // Add charges column if it doesn't exist
  if (trackingSheet.getRange(1, lastCol).getValue() !== 'Charges') {
    trackingSheet.getRange(1, lastCol + 1).setValue('Charges');

    // Calculate charges for each user
    const users = new Set(trackingSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat());
    let currentRow = 2;

    users.forEach(user => {
      const userRows = trackingSheet.getRange(2, 1, lastRow - 1, 1)
        .getValues()
        .map((row, index) => row[0] === user ? index + 2 : null)
        .filter(row => row !== null);

      const lastUserRow = userRows[userRows.length - 1];
      const formula = `=IF(SUM(${String.fromCharCode(65 + lastCol - 1)}${userRows[0]}:${String.fromCharCode(65 + lastCol - 1)}${lastUserRow})/5/${lastCol - 3}<0.8,${CHARGE_AMOUNT},0)`;

      trackingSheet.getRange(currentRow, lastCol + 1).setFormula(formula);
      currentRow += 5;
    });
  }
}

// Update summary view with enhanced calculations
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

  // Clear existing summary data
  if (summarySheet.getLastRow() > 1) {
    summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 4).clear();
  }

  // Get all tracking sheets
  const trackingSheets = ss.getSheets()
    .filter(sheet => sheet.getName().startsWith('Tracking '));

  const summaryData = [];
  const users = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 7).getValues();
  let totalCharges = {};

  trackingSheets.forEach(sheet => {
    const monthYear = sheet.getName().replace('Tracking ', '');
    const data = sheet.getDataRange().getValues();

    // Skip empty sheets
    if (data.length <= 1) return;

    // Calculate completion rates and charges for each user
    users.forEach(user => {
      const userName = user[0];
      const userRows = data.filter(row => row[0] === userName);

      if (userRows.length > 0) {
        let complete = 0;
        let total = 0;
        let charges = 0;

        userRows.forEach(row => {
          // Count completions and calculate charges
          for (let i = 2; i < row.length - 2; i++) {
            if (row[i] === STATUS.COMPLETE) complete++;
            if (row[i] !== STATUS.EXEMPT && row[i] !== STATUS.EMPTY) total++;
          }
        });

        const completionRate = total > 0 ? (complete / total * 100).toFixed(1) + '%' : '0%';
        charges = parseFloat(completionRate) < 80 ? CHARGE_AMOUNT : 0;

        totalCharges[userName] = (totalCharges[userName] || 0) + charges;
        summaryData.push([userName, monthYear, completionRate, charges]);
      }
    });
  });

  // Update summary sheet
  if (summaryData.length > 0) {
    const summaryRange = summarySheet.getRange(2, 1, summaryData.length, 4);
    summaryRange.setValues(summaryData);

    // Format summary range
    summaryRange.setFontFamily('Arial');
    summaryRange.setFontSize(10);
    summaryRange.setHorizontalAlignment('center');

    // Format charges column
    summarySheet.getRange(2, 4, summaryData.length, 1).setNumberFormat('£#,##0.00');
  }

  // Update total charges in Users sheet
  Object.entries(totalCharges).forEach(([userName, charges]) => {
    const userRow = users.findIndex(row => row[0] === userName) + 2;
    if (userRow > 1) {
      usersSheet.getRange(userRow, 7).setValue(charges);
    }
  });

  SpreadsheetApp.getUi().alert('Success', 'Summary has been updated!', SpreadsheetApp.getUi().ButtonSet.OK);
}
