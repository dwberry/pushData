// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org
function pushData_logPush()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Pushed%20Data", scriptName, scriptTrackingId, systemName)
}


function pushData_logError()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Error%20Pushing%20Data", scriptName, scriptTrackingId, systemName)
}


function pushData_overTime()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Over%20Time", scriptName, scriptTrackingId, systemName)
}


function logInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  var pushdata_uid = UserProperties.getProperty("pushdata_uid");
  if(pushdata_uid == null || pushdata_uid == ""){
    NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
  }else{
    NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
  }
}


function getRootFolderId() {
  var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var rootFolder = DocsList.getFileById(ssID).getParents()[0];
  var folderId = rootFolder.getId();
  return folderId;
}

function addFileToPushDataRoot(ssID) {
  var file = DocsList.getFileById(ssID);
  var folder = DocsList.createFolder("pushData Root");
  file.addToFolder(folder);
}

function getFileIdFromName(folderId, fileName) {
  var folder = DocsList.getFolderById(folderId);
  var files = folder.getFiles();
  var count = 0;
  for (var i = 0; i<files.length; i++) {
    if (files[i].getName()==fileName) {
      count++;
    }
  }
  if (count == 0) {
    fileId = "File not found";
    return fileId;
  }
  if (count > 1) {
    fileId = "Duplicate filenames found";
    return fileId;
  }
  if (count == 1) {
    for (var i = 0; i<files.length; i++) {
      if (files[i].getName()==fileName) {
        var fileId = files[i].getId();
        break;
      }
    }
  }
  return fileId;
}
