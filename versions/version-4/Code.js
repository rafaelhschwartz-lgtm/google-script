/**
 * Recursively scans Google Drive for files, folders, and shortcuts.
 * Logs progress to Google Sheets and sends an automated email alert if errors occur.
 */
function exportDriveInventoryWithEmailAlerts() {
  const startTime = new Date();
  const logs = [];
  const errorDetails = [];
  
  // Helper function for synchronized logging and error tracking
  function log(message, level = "INFO") {
    const timestamp = new Date().toISOString();
    const formattedEntry = `[${timestamp}] [${level}] ${message}`;
    Logger.log(formattedEntry);
    logs.push([timestamp, level, message]);
    
    if (level === "WARN" || level === "ERROR") {
      errorDetails.push(formattedEntry);
    }
  }


  log("Starting Google Drive inventory scanning process...");


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Inventory Sheet
  let inventorySheet = ss.getSheetByName("Drive Inventory");
  if (!inventorySheet) {
    inventorySheet = ss.insertSheet("Drive Inventory");
  } else {
    inventorySheet.clear();
  }


  // Setup Execution Log Sheet
  let logSheet = ss.getSheetByName("Execution Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Execution Log");
  } else {
    logSheet.clear();
  }


  const inventoryHeaders = [
    "Item Type", "Item Name", "Item ID", "MIME Type",
    "Path / Parent Folder", "Size (KB)", "Created Date",
    "Last Updated", "Owner", "URL", "Shortcut Target ID"
  ];
  const inventoryData = [inventoryHeaders];


  let stats = { filesProcessed: 0, foldersProcessed: 0, shortcutsProcessed: 0, errorsEncountered: 0 };


  try {
    // Start recursion from Root Folder
    const rootFolder = DriveApp.getRootFolder();
    log(`Accessed Root Folder: "${rootFolder.getName()}" (${rootFolder.getId()})`);
    
    processFolderWithLogging(rootFolder, "/", inventoryData, log, stats);
    
    // Output Inventory Data
    if (inventoryData.length > 1) {
      log(`Writing ${inventoryData.length - 1} records to "Drive Inventory" sheet...`);
      inventorySheet.getRange(1, 1, inventoryData.length, inventoryHeaders.length).setValues(inventoryData);
      inventorySheet.getRange(1, 1, 1, inventoryHeaders.length)
        .setFontWeight("bold")
        .setBackground("#1A73E8")
        .setFontColor("#FFFFFF");
      inventorySheet.autoResizeColumns(1, inventoryHeaders.length);
    }
  } catch (err) {
    log(`CRITICAL ERROR during drive scan: ${err.message}`, "ERROR");
    stats.errorsEncountered++;
  }


  const endTime = new Date();
  const elapsedSeconds = ((endTime.getTime() - startTime.getTime()) / 1000).toFixed(2);


  log(`Execution completed in ${elapsedSeconds} seconds.`);
  log(`Summary — Folders: ${stats.foldersProcessed}, Files: ${stats.filesProcessed}, Shortcuts: ${stats.shortcutsProcessed}, Errors: ${stats.errorsEncountered}`);


  // Output Execution Log to Sheet
  const logHeaders = ["Timestamp", "Level", "Message"];
  const logData = [logHeaders, ...logs];
  
  logSheet.getRange(1, 1, logData.length, logHeaders.length).setValues(logData);
  logSheet.getRange(1, 1, 1, logHeaders.length)
    .setFontWeight("bold")
    .setBackground("#34A853")
    .setFontColor("#FFFFFF");
  logSheet.autoResizeColumns(1, logHeaders.length);


  // Send automated email alert if any errors or warnings were logged
  if (stats.errorsEncountered > 0) {
    sendErrorAlertEmail(stats, errorDetails, elapsedSeconds);
  }
}


/**
 * Sends an email notification to the active user detailing any scan issues.
 */
function sendErrorAlertEmail(stats, errorDetails, elapsedSeconds) {
  const recipient = Session.getActiveUser().getEmail();
  const subject = `⚠️ Google Drive Inventory Alert: ${stats.errorsEncountered} Issue(s) Detected`;
  
  const body = `Hello,


The Google Drive inventory scan completed in ${elapsedSeconds} seconds with ${stats.errorsEncountered} issue(s) encountered.


Scan Statistics:
- Folders Processed: ${stats.foldersProcessed}
- Files Processed: ${stats.filesProcessed}
- Shortcuts Processed: ${stats.shortcutsProcessed}
- Errors / Warnings: ${stats.errorsEncountered}


Log Excerpts:
--------------------------------------------------
${errorDetails.join("\n")}
--------------------------------------------------


You can review the full report in the "Execution Log" tab of your Google Sheet.


Automated alert sent by Google Apps Script.`;


  MailApp.sendEmail(recipient, subject, body);
  Logger.log(`Alert email successfully sent to ${recipient}`);
}


/**
 * Recursive folder processing function with error catching.
 */
function processFolderWithLogging(folder, currentPath, inventoryData, log, stats) {
  const folderName = folder.getName();
  const folderPath = currentPath + folderName + "/";
  stats.foldersProcessed++;


  log(`Processing folder: ${folderPath}`);


  // Record Folder Entry
  try {
    inventoryData.push([
      "Folder",
      folderName,
      folder.getId(),
      MimeType.FOLDER,
      currentPath,
      "-",
      folder.getDateCreated(),
      folder.getLastUpdated(),
      folder.getOwner() ? folder.getOwner().getEmail() : "N/A",
      folder.getUrl(),
      "-"
    ]);
  } catch (e) {
    log(`Failed to read folder metadata for ${folderPath}: ${e.message}`, "WARN");
    stats.errorsEncountered++;
  }


  // Read Files
  try {
    const files = folder.getFiles();
    while (files.hasNext()) {
      try {
        const file = files.next();
        const mimeType = file.getMimeType();
        const isShortcut = mimeType === MimeType.SHORTCUT;
        const itemType = isShortcut ? "Shortcut" : "File";
        let targetId = "-";


        if (isShortcut) {
          stats.shortcutsProcessed++;
          try {
            targetId = file.getTargetId() || "N/A";
          } catch (shortcutErr) {
            log(`Could not resolve target ID for shortcut "${file.getName()}": ${shortcutErr.message}`, "WARN");
            targetId = "ERROR_RESOLVING";
            stats.errorsEncountered++;
          }
        } else {
          stats.filesProcessed++;
        }


        inventoryData.push([
          itemType,
          file.getName(),
          file.getId(),
          mimeType,
          folderPath,
          (file.getSize() / 1024).toFixed(2),
          file.getDateCreated(),
          file.getLastUpdated(),
          file.getOwner() ? file.getOwner().getEmail() : "N/A",
          file.getUrl(),
          targetId
        ]);
      } catch (fileErr) {
        log(`Error reading file in ${folderPath}: ${fileErr.message}`, "WARN");
        stats.errorsEncountered++;
      }
    }
  } catch (filesErr) {
    log(`Permission error listing files in ${folderPath}: ${filesErr.message}`, "WARN");
    stats.errorsEncountered++;
  }


  // Recurse Subfolders
  try {
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      processFolderWithLogging(subFolder, folderPath, inventoryData, log, stats);
    }
  } catch (subFolderErr) {
    log(`Permission error listing subfolders in ${folderPath}: ${subFolderErr.message}`, "WARN");
    stats.errorsEncountered++;
  }
}
