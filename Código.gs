// Main function
function getFiles() {

  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssid = ss.getId();
  var sheet = ss.getActiveSheet();
  
  // Set up the spreadsheet to display the results
  var headers = [["Parent Folder", "File Name", "Status", "Last Updated", "Date created", "File URL"]];
  sheet.getRange("A1:F").clear();
  sheet.getRange("A1:F1").setValues(headers);
  
  
  
  // Look in the same folder the sheet exists in. For example, if this template is in
  // My Drive, it will return all of the files in My Drive.
  var ssparents = DriveApp.getFileById(ssid).getParents();
  
  
  // Loop through all the files and add the values to the spreadsheet.
  while(ssparents.hasNext()) {
    var folder = ssparents.next();
    var folders = folder.getFolders();
    var files = folder.getFiles();
    
    // Start row counter and write files in parent folder
    var i=1;
    i = write(folder,files,ss,sheet,i)
    
    // Iterate through subfolders and write their files
    while(folders.hasNext()) {
      var folder = folders.next();
      var files = folder.getFiles();
      i = write(folder,files,ss,sheet,i)
      
    }
  }
}

// function that writes the folders details to the spreadsheet's sheet
function write(folder,files,ss,sheet,i){
  while(files.hasNext()) {
    var file = files.next();
    if(ss.getId() == file.getId()){ 
      continue; 
    }
    try{
    var doc = DocumentApp.openById(file.getId());
    sheet.getRange(i+1, 1, 1, 6).setValues([[folder.getName(),file.getName(),status(doc),file.getLastUpdated(),file.getDateCreated(), file.getUrl()]]);
    i++;  
    }catch(e){console.log(e)}
  }
  return i;
}

function status(doc){
  
  var completion = doc.getHeader()
  if (completion==null){completion=""}
  else {
    completion=doc.getHeader().getText()}
  return completion
}