/**
 * Recursively lists all files, folders, and shortcuts in Google Drive with relevant metadata.
 * Outputs results into the active Google Sheet.
 */
function exportDriveInventoryToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Drive Inventory");
  
  if (!sheet) {
    sheet = ss.insertSheet("Drive Inventory");
  } else {
    sheet.clear();
  }
  
  // Define header row
  const headers = [
    "Item Type",
    "Item Name",
    "Item ID",
    "MIME Type",
    "Path / Parent Folder",
    "Size (KB)",
    "Created Date",
    "Last Updated",
    "Owner",
    "URL",
    "Shortcut Target ID"
  ];
  
  const data = [headers];
  
  // Start recursion from Root Folder (or set specific folder ID via DriveApp.getFolderById("YOUR_FOLDER_ID"))
  const rootFolder = DriveApp.getRootFolder();
  processFolder(rootFolder, "/", data);
  
  // Write data to sheet in bulk for performance
  if (data.length > 1) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4A90E2").setFontColor("#FFFFFF");
    sheet.autoResizeColumns(1, headers.length);
  }
  
  Logger.log(`Finished! Processed ${data.length - 1} Drive items.`);
}

/**
 * Helper recursive function to process a folder and its contents.
 */
function processFolder(folder, currentPath, data) {
  const folderPath = currentPath + folder.getName() + "/";
  
  // Record folder metadata
  data.push([
    "Folder",
    folder.getName(),
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
  
  // Process files inside folder
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const isShortcut = file.getMimeType() === MimeType.SHORTCUT;
    const itemType = isShortcut ? "Shortcut" : "File";
    const targetId = isShortcut ? (file.getTargetId() || "N/A") : "-";
    
    data.push([
      itemType,
      file.getName(),
      file.getId(),
      file.getMimeType(),
      folderPath,
      (file.getSize() / 1024).toFixed(2),
      file.getDateCreated(),
      file.getLastUpdated(),
      file.getOwner() ? file.getOwner().getEmail() : "N/A",
      file.getUrl(),
      targetId
    ]);
  }
  
  // Recursively process subfolders
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    processFolder(subFolder, folderPath, data);
  }
}
