/**
 * Batch-processed Google Drive inventory scanner using Continuation Tokens.
 * Automatically saves state and self-triggers if execution time approaches quota limits.
 */
function processDriveInventoryBatch() {
  const MAX_RUNTIME_MS = 4.5 * 60 * 1000; // 4.5 minutes safety threshold
  const startTime = new Date().getTime();
  const userProps = PropertiesService.getUserProperties();
  const savedToken = userProps.getProperty('DRIVE_CONTINUATION_TOKEN');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Drive Inventory");
  
  // Initialize sheet headers if starting fresh
  if (!sheet) {
    sheet = ss.insertSheet("Drive Inventory");
    sheet.appendRow([
      "Item Type", "Item Name", "Item ID", "MIME Type",
      "Size (KB)", "Created Date", "Last Updated", "Owner", "URL", "Shortcut Target ID"
    ]);
    sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#1A73E8").setFontColor("#FFFFFF");
  } else if (!savedToken) {
    sheet.clear();
    sheet.appendRow([
      "Item Type", "Item Name", "Item ID", "MIME Type",
      "Size (KB)", "Created Date", "Last Updated", "Owner", "URL", "Shortcut Target ID"
    ]);
    sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#1A73E8").setFontColor("#FFFFFF");
  }


  let fileIterator;
  if (savedToken) {
    Logger.log("Resuming scan from saved continuation token...");
    fileIterator = DriveApp.continueFileIterator(savedToken);
  } else {
    Logger.log("Starting a new full Google Drive scan...");
    fileIterator = DriveApp.getFiles();
  }


  const batchData = [];
  let itemsProcessed = 0;


  while (fileIterator.hasNext()) {
    // Check if execution time is approaching the limit
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      const continuationToken = fileIterator.getContinuationToken();
      userProps.setProperty('DRIVE_CONTINUATION_TOKEN', continuationToken);
      Logger.log(`Time limit approaching. Continuation token saved after ${itemsProcessed} items.`);
      
      // Save current batch to Google Sheet
      if (batchData.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, batchData.length, batchData.length).setValues(batchData);
      }
      
      // Schedule next execution in 1 minute
      clearBatchTriggers();
      ScriptApp.newTrigger('processDriveInventoryBatch')
        .timeBased()
        .after(60 * 1000)
        .create();
        
      return; // Pause execution cleanly
    }


    try {
      const file = fileIterator.next();
      const mimeType = file.getMimeType();
      const isShortcut = (mimeType === MimeType.SHORTCUT);
      const itemType = isShortcut ? "Shortcut" : "File";
      let targetId = "-";


      if (isShortcut) {
        try {
          targetId = file.getTargetId() || "N/A";
        } catch (e) {
          targetId = "ERROR_RESOLVING";
        }
      }


      batchData.push([
        itemType,
        file.getName(),
        file.getId(),
        mimeType,
        (file.getSize() / 1024).toFixed(2),
        file.getDateCreated(),
        file.getLastUpdated(),
        file.getOwner() ? file.getOwner().getEmail() : "N/A",
        file.getUrl(),
        targetId
      ]);
      itemsProcessed++;
    } catch (err) {
      Logger.log(`Error reading item: ${err.message}`);
    }
  }


  // Write remaining batch data upon final loop
  if (batchData.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, batchData.length, batchData.length).setValues(batchData);
    sheet.autoResizeColumns(1, 10);
  }


  // Clean up token and triggers when fully finished
  userProps.deleteProperty('DRIVE_CONTINUATION_TOKEN');
  clearBatchTriggers();
  Logger.log(`Scan complete! Logged ${itemsProcessed} items in final batch.`);
}


/**
 * Removes temporary time triggers created for batch continuation.
 */
function clearBatchTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processDriveInventoryBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


/**
 * Utility to manually clear token and start fresh if needed.
 */
function resetInventoryScan() {
  PropertiesService.getUserProperties().deleteProperty('DRIVE_CONTINUATION_TOKEN');
  clearBatchTriggers();
  Logger.log("Continuation token and triggers reset.");
}
