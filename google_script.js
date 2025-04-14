var currentIndex = 0;  // Track the position globally to continue where the last run left off

function startPDFGeneration() {
  showProgressDialog('Initializing PDF Generation...');
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Today_data');
  var parentFolderId = '1JZx1m-Tvt5uW5VMrFHFzr51_2qUqIgaC';  // Replace with your actual parent folder ID
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var date = new Date();
  var timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  var folderName = 'Generated PDFs_' + timestamp;
  var folder = parentFolder.createFolder(folderName);  // Create a new folder for each batch
  
  var data = sheet.getDataRange().getValues();
  var uniqueNames = {};

  if (data.length <= 1) {
    Logger.log("No data found.");
    closeProgressDialog();
    return;
  }

  // Organize data by unique name in Column F (index 5)
  for (var i = 1; i < data.length; i++) {
    var name = data[i][5];  // Column F (index 5)
    if (name) {
      if (!uniqueNames[name]) {
        uniqueNames[name] = [];
      }
      uniqueNames[name].push(data[i]);
    }
  }

  var pdfLinks = [];
  var batchCount = 0;
  var nameKeys = Object.keys(uniqueNames).sort();

  // Generate PDFs for all names in the list
  for (var i = currentIndex; i < nameKeys.length; i++) {
    var name = nameKeys[i];
    var rows = uniqueNames[name];

    // Check if the PDF has already been generated
    if (isPDFGeneratedByNumber(sheet, name)) {
      Logger.log("PDF already generated for: " + name + ". Skipping.");
      continue;  // Skip processing this name if PDF has already been generated
    }

    var tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    tempSheet.appendRow(header[0]);

    // Add Serial Numbers and Rows
    var serialNumber = 1;
    for (var j = 0; j < rows.length; j++) {
      tempSheet.appendRow([serialNumber].concat(rows[j].slice(1)));  // Add serial number to the start of each row
      serialNumber++;
    }

    // Add the title to the first row in the temporary sheet
    var title = 'Today Data (' + timestamp + ') for ' + name;
    tempSheet.insertRows(1, 1);  // Insert a new row at the top
    tempSheet.getRange(1, 1).setValue(title);  // Set the title in the first cell

    // Dynamically calculate the range for PDF (from B1 to last row with data in Column J)
    var lastRow = tempSheet.getLastRow();
    var pdfRange = tempSheet.getRange(1, 2, lastRow, 9);  // Range B1:J(lastRow)

    try {
      var pdfUrl = generatePDF(pdfRange, folder);  // Use the function for generating PDF

      if (pdfUrl) {
        pdfLinks.push([name, pdfUrl]);

        // Mark the unique number in Column K (assuming K is the 11th column, index 10)
        markPDFGenerated(sheet, name);

        Logger.log("Generated PDF for: " + name + " with URL: " + pdfUrl);

        // Clean up by deleting the temporary sheet
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);

        batchCount++;
        if (batchCount >= 50) {  // Adjust this if you want to stop after a certain number of PDFs
          updateProgressDialog('Processed 50 PDFs so far, please wait...');
          break;
        }
      }
    } catch (e) {
      Logger.log("Error generating PDF for " + name + ": " + e.message);
    }
  }

  closeProgressDialog(); // Close progress dialog when done

  // Now update Sheet2 with the PDF links by matching the names in Column C
  if (pdfLinks.length > 0) {
    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
    if (!sheet2) {
      sheet2 = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Sheet2');
      Logger.log("Created Sheet2");
    }

    var sheet2Data = sheet2.getRange(1, 3, sheet2.getLastRow(), 1).getValues();
    var namesInSheet2 = sheet2Data.map(function(row) { return row[0]; });

    // Loop through pdfLinks and place links in Sheet2
    for (var i = 0; i < pdfLinks.length; i++) {
      var pdfName = pdfLinks[i][0];  
      var pdfUrl = pdfLinks[i][1];   

      Logger.log("Searching for PDF for: " + pdfName);

      var index = namesInSheet2.indexOf(pdfName);

      if (index !== -1) {
        sheet2.getRange(index + 1, 1).setValue(pdfUrl);
      } else {
        var lastRow = sheet2.getLastRow() + 1;
        sheet2.getRange(lastRow, 1).setValue(pdfUrl);
      }
    }
  } else {
    Logger.log("No PDFs generated.");
  }
}

// Function to generate PDF
function generatePDF(range, folder) {
  var url = 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActiveSpreadsheet().getId() + '/export?format=pdf&gid=' + range.getSheet().getSheetId() + '&range=' + range.getA1Notation();

  var options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };

  var attempts = 0;
  var maxAttempts = 5;
  var delay = 1000;

  while (attempts < maxAttempts) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var pdf = response.getBlob();
      var fileName = range.getSheet().getName() + ".pdf";
      var pdfFile = folder.createFile(pdf);
      pdfFile.setName(fileName);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return pdfFile.getUrl();
    } catch (e) {
      Logger.log("Error during PDF generation: " + e.message);
      if (e.message.includes("Request failed for https://docs.google.com returned code 429")) {
        attempts++;
        Logger.log("Rate limit hit, attempt " + attempts + " of " + maxAttempts + ". Retrying in " + delay + "ms.");
        Utilities.sleep(delay);  
        delay *= 2;
      } else {
        throw e;
      }
    }
  }
  
  throw new Error("Max retry attempts reached for creating PDF.");
}

// Function to check if the PDF for the name already exists in Column K (unique number)
function isPDFGeneratedByNumber(sheet, name) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] == name && data[i][11]) {
      return true; 
    }
  }
  return false;
}

// Function to mark that the PDF has been generated by adding a unique number in Column K
function markPDFGenerated(sheet, name) {
  var data = sheet.getDataRange().getValues();
  var uniqueNumber = Utilities.getUuid();  // Generate a unique UUID as the number
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] == name) {  
      sheet.getRange(i + 1, 11).setValue(uniqueNumber);  
    }
  }
}

// Function to show the progress dialog
function showProgressDialog(message) {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutput('<h3>' + message + '</h3>')
                              .setWidth(250)
                              .setHeight(100);
  ui.showModalDialog(htmlOutput, 'Generating PDF...');
}

// Function to update the progress dialog with dynamic text (or progress bar)
function updateProgressDialog(message) {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutput('<h3>' + message + '</h3>')
                              .setWidth(250)
                              .setHeight(100);
  ui.showModalDialog(htmlOutput, 'Generating PDF...');
}

// Function to close the progress dialog
function closeProgressDialog() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('PDF Generation Completed');
  ui.showModalDialog(HtmlService.createHtmlOutput('<h3>Processing Complete!</h3>'), 'Process Finished');
}
