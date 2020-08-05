// Main function
function getFiles() {

  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssid = ss.getId();
  var sheet = ss.getActiveSheet();
  
  // Set up the spreadsheet to display the results
  var headers = [["Parent Folder", "File Name", "Status", "Success", "Last Updated", "Date created", "File URL"]];
  sheet.getRange("A1:G").clear();
  sheet.getRange("A1:G1").setValues(headers);
  
  
  
  // Look in the same folder the sheet exists in. For example, if this template is in
  // My Drive, it will return all of the files in My Drive.
  var ssparent = DriveApp.getFileById(ssid).getParents().next();
  
  // Start row counter and adding files from folders
  var i = 1;
  i = subfolderation(i,ssparent,ss,sheet);   
    
  }

function subfolderation(i,folder,ss,sheet){  
  // Write folder files
  var files = folder.getFiles();
  i = write(folder,files,ss,sheet,i);
  
  // get subfolders
  var subfolders = folder.getFolders();
  
  // Repeat for subfolders
  while(subfolders.hasNext()){
    var subfolder = subfolders.next();
    i = subfolderation(i,subfolder,ss,sheet);
  }
  return i;
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
    var filestatus = status(doc,file,folder);
    var fileresult = result(doc);
    var lastUpdated = file.getLastUpdated();
    
    sheet.getRange(i+1, 1, 1, 7).setValues([[getFolderName(folder),file.getName(),filestatus,fileresult,lastUpdated,file.getDateCreated(), file.getUrl()]]);
    i++;  
    }catch(e){console.log(e)}
  }
  return i;
}

function getFolderName(folder){
  if (folder.getName().toLowerCase() == "completed"){
  var parentFolder = folder.getParents()
  var folder = parentFolder.next();}
  return folder.getName();
}

function status(doc,file,folder){
  var completion = doc.getHeader()
  if (completion == null){completion = "";}
  else {completion = completion.getText().toLowerCase();}
  checkMoveCompleted(completion,file,folder)
  return completion;
}

function checkMoveCompleted(completion,file,folder){
  
  if (completion.toLowerCase() == "completed")
  {if(DriveApp.getFileById(file.getId()).getParents().next().getName().toLowerCase() != "completed"){
    try{moveCompleted(file,folder);
        console.log( file.getName().concat(" moved to completed folder"))}
    catch(e){console.log( file.getName().toUpperCase().concat(" coudn't be moved to completed folder".concat(e)));}}
  }
}

function moveCompleted(file,folder){
  var subfolders = folder.getFolders()
  var moved = false;
  
  while(subfolders.hasNext()){
    var subfolder = subfolders.next();
    if(subfolder.getName().toLowerCase() == "completed"){      
      DriveApp.getFolderById(subfolder.getId()).addFile(file);
      folder.removeFile(file);
      moved = true;}}
  
  if( moved == false){
    var newFolder = folder.createFolder("Completed");
    DriveApp.getFolderById(newFolder.getId()).addFile(file);
    folder.removeFile(file);
     }
}

function result(doc){
  var success = doc.getFooter();
  if (success == null){success = "";}
  else {success = success.getText().toLowerCase();}
  return success;
}