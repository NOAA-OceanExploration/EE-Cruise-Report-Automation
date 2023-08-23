function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Generation Menu')
      .addItem('Add Shore Scientists to Table', 'createTableInDoc')
      .addItem('Add Operations Results', 'operationsResults')
      .addItem('Copy SITREP from Drive', 'copySitrep')
      .addToUi();
}

function createTableInDoc() {
  // Get the active Google Document
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  // Extract the expedition dates from the document
  var content = body.getText();
  var datesMatch = content.match(/Expedition Dates: (.+? \d+)-(.+? \d+), (\d{4})/);
  var startDate = new Date(datesMatch[1] + ", " + datesMatch[3]);
  var endDate = new Date(datesMatch[2] + ", " + datesMatch[3]);
  
  // Extract names from the text file
  var fileId = "1xkxdoiW5LW4PZSFZLRBmHODTdIdQvMbL";
  var file = DriveApp.getFileById(fileId);
  var fileContent = file.getBlob().getDataAsString();
  var lines = fileContent.split("\n");
  var names = [];
  var currentDate = null;
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].endsWith(":")) {
      currentDate = new Date(lines[i].slice(0, -1));
    } else if (currentDate >= startDate && currentDate <= endDate) {
      names.push(lines[i]);
    }
  }
  
  // Access the spreadsheet and extract affiliations
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1AEO1oraEuJZLvypgBMcsXI1h6_8tc6Y-ZmuPicydmpg/edit#gid=958012620");
  var sheet = ss.getSheetByName("Viewer 1");
  var firstNames = sheet.getRange("C5:C").getValues();
  var lastNames = sheet.getRange("D5:D").getValues();
  var affiliations = sheet.getRange("I5:I").getValues();
  
  function capitalizeName(name) {
      return name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
  }

  function isNameInTable(name, table) {
      for (var i = 1; i < table.length; i++) { // Start from 1 to skip the header row
          if (table[i][0] === name) {
              return true;
          }
      }
      return false;
  }

  var tableData = [["Name", "Role", "Affiliation"]];
  for (var i = 0; i < names.length; i++) {
      for (var j = 0; j < firstNames.length; j++) {
          var capitalizedFirstName = capitalizeName(firstNames[j][0]);
          var capitalizedLastName = capitalizeName(lastNames[j][0]);
          var fullName = capitalizedFirstName + " " + capitalizedLastName; // Names are now capitalized
          
          if (names[i] === fullName.toLowerCase().replace(" ", "")) { // Remove space for comparison
              if (fullName.trim() !== "" && !isNameInTable(fullName, tableData)) { // Ensure the name is not empty and not already in the table
                  tableData.push([fullName, "Shore-Based Scientist", affiliations[j][0]]);
              }
              break;
          }
      }
  }

  // Get all tables in the document
  var tables = body.getTables();

  if (tables.length >= 10) {
      // Get the position of the 9th table
      var position = body.getChildIndex(tables[9]);
      
      // Remove the 9th table
      tables[9].removeFromParent();
      
      // Insert the new table at the position of the removed 9th table
      body.insertTable(position, tableData);
  } else {
      // If there are fewer than 9 tables, append the table at the end
      body.appendTable(tableData);
  }
}


function getFileByLink(link) {
  // Extract fileId from the link
  var fileId = link.match(/[-\w]{25,}/);
  
  return fileId
}

function getAllShareableLinksFromFolder(folder) {
  var fileIterator = folder.getFiles();
  var fileLinks = [];
  
  while (fileIterator.hasNext()) {
    var file = fileIterator.next();
    
    // Assuming each file is already shared
    var link = file.getUrl();
    fileLinks.push(link);
  }

  return fileLinks;
}

function copySitrep() {
  // Step 1: Get the folder ID from the provided link
  var folderId = "1n3kU0XsSX62sgaH0asbdU46KRZU6K1dF";  // Replace this with your folder link

  var folder = DriveApp.getFolderById(folderId);
  var fileLinks = getAllShareableLinksFromFolder(folder);

  Logger.log(fileLinks[0])
  for (var i = 0; i < fileLinks.length; i++) {
    var docID = getFileByLink(fileLinks[i]);
    
    // Extract the content of the file
    var content = DocumentApp.openById(docID).getBody().getText();

    // Extract the date after "NOAA OCEAN EXPLORATION AND RESEARCH SITUATION REPORT FOR"
    var dateMatch = content.match(/NOAA OCEAN EXPLORATION AND RESEARCH SITUATION REPORT FOR\s*([\s\S]*?)(?=\n|\s*$)/);
    
    var reportDate;
    if (dateMatch && dateMatch[1]) {
        reportDate = dateMatch[1].trim();
    } else {
        reportDate = "Date Not Found";
    }
    
    // Extract the text between "SURVEY:" and "ROV:"
    var contentMatch = content.match(/SURVEY:\s*([\s\S]*?)\s*ROV:/);
    if (!contentMatch || contentMatch.length < 2) {
      Logger.log('Section between SURVEY and ROV not found for link ' + (i + 1));
      continue;  // Skip to the next link
    }
    var surveyToRovText = contentMatch[1].trim();
    
    // Open the active Google Document
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    
    // Find the "$SITREPs$" keyword and paste the text after it
    var searchElement = body.findText("\\$SITREPs\\$");
    if (!searchElement) {
      Logger.log('$SITREPs$ not found.');
      return;  // Stop the script if "$SITREPs$" isn't found
    }
    var searchElement = body.findText("\\$SITREPs\\$");
    if (searchElement) {
        var elementText = searchElement.getElement().asText();
        var startOffset = searchElement.getStartOffset();
        var endOffset = searchElement.getEndOffsetInclusive();
        elementText.insertText(endOffset + 1, "\n" + reportDate + "\n" + surveyToRovText);
    } else {
        Logger.log('$SITREPs$ not found.');
    }
    Logger.log(startOffset)
    Logger.log('$SITREPs$ found.');
  }
}

function operationsResults() {
  var folderId = '1ca1seh41rTqTtHGRdCwIQvXNG1kSwTbW';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.PLAIN_TEXT);

  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Capture the position of the table (if it exists)
  var tablePosition;
  if (body.getTables().length > 5) {
    tablePosition = body.getChildIndex(body.getTables()[5]);
    body.getTables()[5].removeFromParent();
  }

  // Headers for the new table
  var headers = [
    'Dive #',
    'Site Name',
    'Date (yyyymmdd)',
    'On Bottom Latitude (dd)',
    'On Bottom Longitude (dd)',
    'Max Depth (m)',
    'Min Depth (m)',
    'Dive Duration (hh:mm:ss)',
    'Bottom Time (hh:mm:ss)',
    'Water Column Exploration Time (hh:mm:ss)'
  ];

  // Insert the new table at the captured position
  var table = body.insertTable(tablePosition, [headers]);

  // Prepare a list of unique site names
  var siteNames = [
    "Site Alpha", "Site Beta", "Site Gamma",
    // ... add more site names as needed
  ];

  var siteIndex = 0;

  while (files.hasNext()) {
    // Ensure we don't run out of unique site names
    if (siteIndex >= siteNames.length) {
      console.error('Ran out of unique site names.');
      break;
    }
    var siteName = siteNames[siteIndex]; // Get site name for this row

    var file = files.next();
    var content = file.getBlob().getDataAsString();
    
    var diveNumberMatch = content.match(/Dive Summary:\s+(EX\d+_DIVE\d+)/);
    var diveNumber = diveNumberMatch ? diveNumberMatch[1] : 'N/A';
    
    var dateMatch = content.match(/In Water:\s+(\d{4}-\d{2}-\d{2})/);
    var date = dateMatch ? dateMatch[1].replace(/-/g, '') : 'N/A';
    
    var onBottomLatMatch = content.match(/On Bottom:\s+N\/A\s+(\d+\.\d+)/);
    var onBottomLat = onBottomLatMatch ? onBottomLatMatch[1] : 'N/A';
    
    var onBottomLongMatch = content.match(/On Bottom:\s+N\/A\s+;\s+(\d+\.\d+)/);
    var onBottomLong = onBottomLongMatch ? onBottomLongMatch[1] : 'N/A';
    
    var maxDepthMatch = content.match(/Max Vehicle Depth:\s+(\d+\.\d+)/);
    var maxDepth = maxDepthMatch ? maxDepthMatch[1] : 'N/A';
    
    var minDepthMatch = content.match(/Min Seafloor Depth:\s+(\d+\.\d+)/);
    var minDepth = minDepthMatch ? minDepthMatch[1] : 'N/A';
    
    var diveDurationMatch = content.match(/Dive Duration:\s+(\d+:\d+:\d+)/);
    var diveDuration = diveDurationMatch ? diveDurationMatch[1] : 'N/A';
    
    var bottomTimeMatch = content.match(/Bottom Time:\s+(\d+:\d+:\d+)/);
    var bottomTime = bottomTimeMatch ? bottomTimeMatch[1] : 'N/A';
    
    // Calculate Water Column Exploration Time
    var transectDurations = content.match(/Duration:\s+(\d+:\d+:\d+)/g) || [];
    var totalSeconds = 0;
    for (var i = 0; i < transectDurations.length; i++) {
      var timeParts = transectDurations[i].match(/(\d+):(\d+):(\d+)/);
      totalSeconds += parseInt(timeParts[1]) * 3600 + parseInt(timeParts[2]) * 60 + parseInt(timeParts[3]);
    }
    var hours = Math.floor(totalSeconds / 3600);
    var minutes = Math.floor((totalSeconds - (hours * 3600)) / 60);
    var seconds = totalSeconds - (hours * 3600) - (minutes * 60);
    var waterColumnExplorationTime = hours + ':' + minutes + ':' + seconds;

    // Append a new row to the table and populate with extracted data
    var newRow = table.appendTableRow();
    newRow.appendTableCell(diveNumber);
    newRow.appendTableCell(siteName);
    newRow.appendTableCell(date);
    newRow.appendTableCell(onBottomLat);
    newRow.appendTableCell(onBottomLong);
    newRow.appendTableCell(maxDepth);
    newRow.appendTableCell(minDepth);
    newRow.appendTableCell(diveDuration);
    newRow.appendTableCell(bottomTime);
    newRow.appendTableCell(waterColumnExplorationTime);

    siteIndex++; // Move to the next site name for the next row
  }
}

