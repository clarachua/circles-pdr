
// Get macro variables
var macrosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Macro Page');
var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('For Merge');
// Get rows for data entry
var rowFolderrow = macrosheet.createTextFinder('Enter FolderId').findNext().getRowIndex();
var rowSheetname = macrosheet.createTextFinder('Sheet Name to Protect').findNext().getRowIndex();
var rowDupSheet = macrosheet.createTextFinder('Sheet Name to Duplicate').findNext().getRowIndex();
var rowRenamedSheet = macrosheet.createTextFinder('Rename Duplicated Sheet to').findNext().getRowIndex();
var rowTrackerFolderId = macrosheet.createTextFinder('File FolderId').findNext().getRowIndex();
var rowTrackerFile = macrosheet.createTextFinder('Tracking Sheet Id').findNext().getRowIndex();
var rowTrackerSheet = macrosheet.createTextFinder('Tracking Sheet Name').findNext().getRowIndex();
var rowMissingSheet = macrosheet.createTextFinder('Missing Sheet Name').findNext().getRowIndex();
var rowTrackTab = macrosheet.createTextFinder('Tab to Track').findNext().getRowIndex();
var rowSelEmail = macrosheet.createTextFinder('Add permission').findNext().getRowIndex();
var rowEmailGroup = macrosheet.createTextFinder('Enter email 1').findNext().getRowIndex();
var rowtextChange = macrosheet.createTextFinder('Original').findNext().getRowIndex() +1;
var rowSupsheet = macrosheet.createTextFinder('Supervisor Sheet Name').findNext().getRowIndex();
var rowArchiveFolder = macrosheet.createTextFinder('Archive FolderId').findNext().getRowIndex();
var rowSupRefSheet = macrosheet.createTextFinder('Reference Supervisors Sheet').findNext().getRowIndex();
var ddcopy = macrosheet.getRange("A33:A34").getValues();
var ddname = macrosheet.getRange('A32').getValue();
var colEefileid = infosheet.createTextFinder('New_FileId').findNext().getColumnIndex();
var colFirstname = infosheet.createTextFinder('First Name').findNext().getColumnIndex();
var colLastname = infosheet.createTextFinder('Last Name').findNext().getColumnIndex();
var colJobtitle = infosheet.createTextFinder('Job Title').findNext().getColumnIndex();
var colMgrname = infosheet.createTextFinder('Reporting to').findNext().getColumnIndex();
var colEeid = infosheet.createTextFinder('Employee #').findNext().getColumnIndex();
var col2019link = infosheet.createTextFinder('PDR 2020 Link').findNext().getColumnIndex();

// Get variable values in B column
var folderId = macrosheet.getRange('B'+rowFolderrow).getValue();
var sheetname = macrosheet.getRange('B'+rowSheetname).getValue();
var DupSheet = macrosheet.getRange('B'+rowDupSheet).getValue();
var RenamedSheet = macrosheet.getRange('B'+rowRenamedSheet).getValue();
var trackerFolderId = macrosheet.getRange('B'+rowTrackerFolderId).getValue();
var trackerFile = macrosheet.getRange('B'+rowTrackerFile).getValue();
var trackerSheet = macrosheet.getRange('B'+rowTrackerSheet).getValue();
var missingSheet = macrosheet.getRange('B'+rowMissingSheet).getValue();
var trackTab = macrosheet.getRange('B'+rowTrackTab).getValue();
var Supsheet = macrosheet.getRange('B'+rowSupsheet).getValue();
var refSupsheet = macrosheet.getRange('B'+rowSupRefSheet).getValue();
var archiveFolder = macrosheet.getRange('B'+rowArchiveFolder).getValue();
var selEmail = macrosheet.getRange('B'+rowSelEmail).getValue();
var numEmails = macrosheet.getRange('B'+(rowEmailGroup-1)).getValue();
var emailGroup3 = [macrosheet.getRange('B'+rowEmailGroup).getValue(),
                   macrosheet.getRange('B'+(rowEmailGroup+1)).getValue(),
                   macrosheet.getRange('B'+(rowEmailGroup+2)).getValue() ]
var emailGroup2 = [macrosheet.getRange('B'+rowEmailGroup).getValue(),
                   macrosheet.getRange('B'+(rowEmailGroup+1)).getValue()]
var emailGroup1 = [macrosheet.getRange('B'+rowEmailGroup).getValue()]
var find1 = macrosheet.getRange('A'+rowtextChange).getValue();
var find2 = macrosheet.getRange('A'+(rowtextChange+1)).getValue();
var find3 = macrosheet.getRange('A'+(rowtextChange+2)).getValue();
var find4 = macrosheet.getRange('A'+(rowtextChange+3)).getValue();
var find5 = macrosheet.getRange('A'+(rowtextChange+4)).getValue();
var find6 = macrosheet.getRange('A'+(rowtextChange+5)).getValue();
var find7 = macrosheet.getRange('A'+(rowtextChange+6)).getValue();
var repl1 = macrosheet.getRange('B'+rowtextChange).getValue();
var repl2 = macrosheet.getRange('B'+(rowtextChange+1)).getValue();
var repl3 = macrosheet.getRange('B'+(rowtextChange+2)).getValue();
var repl4 = macrosheet.getRange('B'+(rowtextChange+3)).getValue();
var repl5 = macrosheet.getRange('B'+(rowtextChange+4)).getValue();
var repl6 = macrosheet.getRange('B'+(rowtextChange+5)).getValue();
var repl7 = macrosheet.getRange('B'+(rowtextChange+6)).getValue();
var parentFolder = DriveApp.getFolderById(folderId);
var childFolders = parentFolder.getFolders();
  var i = 0;
  var j = 0;

// Functions start here

// CopyAllFolders, protectTab and changeFormulas
function ready(){
  CopyAllFolders();
  protectTab();
  changeFormulas();
}

// Set Data Validation for new SQ4 template for specific fileId only
function dataval() {
// Log information about the data validation rule for cell A1.
  var destination = SpreadsheetApp.openById("1sXzn_CINLdqyNcKDWbgMMfjcVpe7M4uruB8hFR0BtmI");
  var itt = destination.getSheetByName("PDR Form - Q4");
  var cell1 = itt.getRange("K7:K24");
  var cell2 = itt.getRange("N7:N24");
  var cell3 = itt.getRange("K31:K46");
  var cell4 = itt.getRange("N31:N46");
  // var range = SpreadsheetApp.getActive().getRangeByName('Drop Down!Cult_Rating');
  // var values = range.getValues();
  var rule1 = SpreadsheetApp.newDataValidation().requireValueInRange(destination.getRangeByName('Cult_Rating'), true).build();
  var rule2 = SpreadsheetApp.newDataValidation().requireValueInRange(destination.getRangeByName('Rating'), true).build();
  // var rule = SpreadsheetApp.newDataValidation().requireValueInList(SpreadsheetApp.destination.getRangeByName('Cult_Rating', true).build());
  // cell1.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(destination.getRangeByName('Cult_Rating'), true).build());
  cell1.setDataValidation(rule1)
  cell2.setDataValidation(rule1)
  cell3.setDataValidation(rule2)
  cell4.setDataValidation(rule2)
}

// CopySheetFolder will copy selected sheet ('dupsheet') to all files in 1 folder (i.e. no subfolders) and set data validation
// CopySheetAll will copy selected sheet across 1 level of subfolders (i.e. Main Folder > Folders > File) if there are no subfolders, it will copy files across.

function CopySheetFolder() {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var stcopy = source.getSheetByName(DupSheet);
  var i = 0
  var j = 0
  //while(childFolders.hasNext()) {
    // Only uncomment the next portion if you need the subFolders from the childFolders
    //getSubFolders(child);
    var fileIter = parentFolder.getFiles();

    while(fileIter.hasNext()){
      var file = fileIter.next();
      var filename = file.getName();
      var fileId = file.getId();
      var ss = SpreadsheetApp.openById(fileId);
      // SpreadsheetApp.setActiveSpreadsheet(ss)
      var itt = ss.getSheetByName(RenamedSheet);
      if (!itt) {
        var newsheet = stcopy.copyTo(ss);
        newsheet.setName(RenamedSheet);
        Logger.log(j + "." + newsheet.getName() + "has been copied");
        // Set data validation
          var cell1 = newsheet.getRange("K7:K24");
          var cell2 = newsheet.getRange("N7:N24");
          var cell3 = newsheet.getRange("K31:K46");
          var cell4 = newsheet.getRange("N31:N46");
          var rule1 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('Cult_Rating'), true).build();
          var rule2 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('Rating'), true).build();
          cell1.setDataValidation(rule1);
          cell2.setDataValidation(rule1);
          cell3.setDataValidation(rule2);
          cell4.setDataValidation(rule2);
          Logger.log(j + ". " + "Data validation set for " + fileId)
          j++;
        }
      else {
        Logger.log('Tab in ' + filename + 'already exists');
        j++;
      }
    }

}

// CopySheetFolder will copy selected sheet ('dupsheet') to all files in 1 folder (i.e. no subfolders)
// CopySheetAll will copy selected sheet across 1 level of subfolders (i.e. Main Folder > Folders > File) if there are no subfolders, it will copy files across.

function CopySheetAll() {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var stcopy = source.getSheetByName(DupSheet);
  var i = 0
  var j = 0

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      // you will need to do a SubFolder.Iter and then a fileIter on the SubFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        // SpreadsheetApp.setActiveSpreadsheet(ss)
        var itt = ss.getSheetByName(RenamedSheet);
        if (!itt) {
          var newsheet = stcopy.copyTo(ss);
          newsheet.setName(RenamedSheet);
          Logger.log(j + "." + newsheet.getName() + "has been copied");
          // Set data validation
          var cell1 = newsheet.getRange("K7:K24");
          var cell2 = newsheet.getRange("N7:N24");
          var cell3 = newsheet.getRange("K31:K46");
          var cell4 = newsheet.getRange("N31:N46");
          var rule1 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('Cult_Rating'), true).build();
          var rule2 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRangeByName('Rating'), true).build();
          cell1.setDataValidation(rule1);
          cell2.setDataValidation(rule1);
          cell3.setDataValidation(rule2);
          cell4.setDataValidation(rule2);
          Logger.log(j + ". " + "Data validation set for " + fileId)
          j++;
        }
        else {
          Logger.log('Tab in ' + filename + 'already exists');
          j++;
        }
      }
    }
  }
  else {
    CopySheetFolder();
  }
}

// Add named range to Drop Down menu
function createNamedRange() {

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        var ss1 = ss.getSheetByName("Drop Down");
        var range1 = ss1.getRange('D1');
        var range2 = ss1.getRange('D2:D3');
        // var range3 = ss1.getRange('B2:B5')
        range1.setValue(ddname)
        range2.setValues(ddcopy)
        ss.setNamedRange(ddname, range2);
        ss.setNamedRange('Rating', ss1.getRange('B2:B5'))
        Logger.log(j+1, ".", filename, "has new range1", ddname);
        j++;
      }
    }
    Logger.log("There are " + i + " folders.");
    Logger.log("There are " + j + " files.");
  }
  else {
      var fileIter = parentFolder.getFiles();
      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        var ss1 = ss.getSheetByName("Drop Down");
        var range1 = ss1.getRange('D1');
        var range2 = ss1.getRange('D2:D3');
        range1.setValue(ddname);
        range2.setValues(ddcopy);
        // ss1.activate();
        ss.setNamedRange(range1, range2);
        // ss1.setNamedRange(ddname, range2);
        Logger.log(j+1, ".", filename, "has new range", ddname);
        j++;
      }
    Logger.log("There are " + i + " folders.");
    Logger.log("There are " + j + " files.");
  }
}


// Function to go down the spreadsheet, open the file by id, copy name from
// table to sheet2 If this works, do another function to get copy various
// ranges in the fileID and copy to tracker
function tabletosheet() {
  // test on filtered open spreadsheet, print first name and last name
  var range = infosheet.getRange("A:AC");
  var nRows = range.getLastRow();

  for (var j = 1; j < nRows; j++) {
    var row = j+1
    if (infosheet.isRowHiddenByFilter(row) != true) {
      // do stuff
      var firstname = infosheet.getRange(row, colFirstname).getValue();
      var lastname = infosheet.getRange(row, colLastname).getValue();
      var jobtitle = infosheet.getRange(row, colJobtitle).getValue();
      var mgrname = infosheet.getRange(row, colMgrname).getValue();
      var eefileid = infosheet.getRange(row, colEefileid).getValue();
      var eeid = infosheet.getRange(row, colEeid).getValue();
      var oldfile = infosheet.getRange(row, col2019link).getValue();

      // var
      Logger.log(j + " Name:" + firstname + " " + lastname + " Mgr name: " + mgrname + " Job title: " + jobtitle + eefileid);

      // Open fileId
      if (eefileid != "") {
        var ss = SpreadsheetApp.openById(eefileid);
        var pdrsheet = ss.getSheetByName(RenamedSheet);
        var pdrname = pdrsheet.getRange('F2');
        var pdrmgr = pdrsheet.getRange('F3');
        var pdrtitle = pdrsheet.getRange('H2');
        var pdrempid = pdrsheet.getRange('D2');
        var pdroldfile = pdrsheet.getRange('K54');

        // Set text values
        pdrname.setValue(firstname + " " + lastname);
        pdrmgr.setValue(mgrname);
        pdrtitle.setValue(jobtitle);
        pdrempid.setValue(eeid);
        pdroldfile.setValue(oldfile)
      }
      continue
    }
  }
}

// Rename Sheet
function renamesinglesheet(){
  // test on filtered open spreadsheet, print first name and last name
  msheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MissingSheets')
  var range = msheet.getRange("A:B");
  var nRows = range.getLastRow();
  var fileidcol = msheet.createTextFinder('FileId').findNext().getColumnIndex();

  for (var j = 1; j < nRows; j++) {
    var row = j+1
    if (msheet.isRowHiddenByFilter(row) != true) {
      // do stuff
      var filename1 = msheet.getRange(row, 1).getValue();
      var eefileid = msheet.getRange(row, fileidcol).getValue();
      // var
      Logger.log(j + " File Name: " + filename1 + " " + eefileid);
      // do more stuff
      // Open fileId
      if (eefileid != "") {
        var ss = SpreadsheetApp.openById(eefileid);
        var sht = ss.getSheetByName('SQ4 PDR Template')
        sht.setName(RenamedSheet);
        Logger.log(j + " FileId " + eefileid + "has been renamed")
      }
    }
  }
}

// Function to go down the spreadsheet, open the old file by id, copy items from file to new sheet by id
function copyCultFromOldId() {
  // test on filtered open spreadsheet, print first name and last name
  var range = infosheet.getRange("A:Z");
  var nRows = range.getLastRow();
  var fileidcolold = infosheet.createTextFinder('Old_FileId').findNext().getColumnIndex();
  var fileidcolnew = infosheet.createTextFinder('New_FileId').findNext().getColumnIndex();

  for (var j = 1; j < nRows; j++) {
    var row = j+1
    if (infosheet.isRowHiddenByFilter(row) != true) {
      // get info
      // var firstname = infosheet.getRange(row, 4).getValue();
      // var lastname = infosheet.getRange(row, 5).getValue();
      // var jobtitle = infosheet.getRange(row, 7).getValue();
      // var mgrname = infosheet.getRange(row, 11).getValue();
      var eefileidold = infosheet.getRange(row, fileidcolold).getValue();
      var eefileidnew = infosheet.getRange(row, fileidcolnew).getValue();

      // var
      Logger.log(j + " Old FileId:" + eefileidold + " New FileId: " + eefileidnew);
      // do more stuff
      // Open fileId
      if (eefileidold != "") {
        var ss = SpreadsheetApp.openById(eefileidold);
        var pdrsheetold = ss.getSheetByName(DupSheet);
        var pdrcvold = pdrsheetold.getRange('K7:P24').getValues();

      // // Test values
      // Logger.log(j + "values are: " + pdrcvold)
        // Set values
        var ssnew = SpreadsheetApp.openById(eefileidnew);
        var pdrsheetnew = ssnew.getSheetByName(RenamedSheet);
        var pdrnewrange = pdrsheetnew.getRange('K7:P24');
        pdrnewrange.setValues(pdrcvold)
        Logger.log(j + " Sheet range copied")
      }
      continue
    }
  }
}



// New tracker function to go down the spreadsheet, open the file by id, copy items from file to Tracker2 sheet
function newTracker() {
  // test on filtered open spreadsheet, print first name and last name
  var range = infosheet.getRange("A:Z");
  var nRows = range.getLastRow();
  var fileidcol = infosheet.createTextFinder('FileId').findNext().getColumnIndex();

  for (var j = 1; j < nRows; j++) {
    var row = j+1
    if (infosheet.isRowHiddenByFilter(row) != true) {
      // do stuff
      var firstname = infosheet.getRange(row, 4).getValue();
      var lastname = infosheet.getRange(row, 5).getValue();
      var jobtitle = infosheet.getRange(row, 7).getValue();
      var mgrname = infosheet.getRange(row, 11).getValue();
      var eefileid = infosheet.getRange(row, fileidcol).getValue();
      // var
      Logger.log(j + " Name:" + firstname + " " + lastname + " Mgr name: " + mgrname + " Job title: " + jobtitle + eefileid);
      // do more stuff
      // Open fileId
      if (eefileid != "") {
        var ss = SpreadsheetApp.openById(eefileid);
        var pdrsheet = ss.getSheetByName(RenamedSheet);
        var pdrname = pdrsheet.getRange('F2');
        var pdrmgr = pdrsheet.getRange('F3');
        var pdrtitle = pdrsheet.getRange('H2');

        // Set text values
        pdrname.setValue(firstname + " " + lastname);
        pdrmgr.setValue(mgrname);
        pdrtitle.setValue(jobtitle);
      }
      continue
    }
  }
}

// Function to go down the spreadsheet, add permission to folders for supervisors
function folderPermission() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Test");
  // test on filtered open spreadsheet, print first name and last name
  var range = sheet.getDataRange();
  var nRows = range.getLastRow();
  var fileidcol = sheet.createTextFinder('FolderId').findNext().getColumnIndex();
  var supemailcol = sheet.createTextFinder('Supervisor email').findNext().getColumnIndex();

  for (var j = 1; j < nRows; j++) {
    var row = j+1
      var folderid = sheet.getRange(row, fileidcol).getValue();
      var supemail = sheet.getRange(row, supemailcol).getValue();
      Logger.log(j + " Folderid: " + folderid + ", sup email: " + supemail + ", nrows = " + nRows);
      var folder = DriveApp.getFolderById(folderid);
      folder.addEditor(supemail);
      Logger.log(j + ", " + supemail + " added to " + folderid);
  }
}

//addPermission only adds permission to the protected sheet for 1 user and does not remove protection for any existing user
function addPermission() {
  Logger.log("Tab", sheetname);

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        var spreadsheet = ss.getSheetByName(sheetname);
        var editPerm = spreadsheet.protect().addEditor(selEmail);
        Logger.log(j+1, ".", filename, spreadsheet, "can be edited by", selEmail);
        j++;
      }
    }
    Logger.log("There are " + i + " folders.");
    Logger.log("There are " + j + " files.");
  }
  else {
      var fileIter = parentFolder.getFiles();
      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        var spreadsheet = ss.getSheetByName(sheetname);
        var editPerm = spreadsheet.protect().addEditor(selEmail);
        Logger.log(j+1, ".", filename, spreadsheet, "can be edited by", selEmail);
        j++;
      }
    Logger.log("There are " + i + " folders.");
    Logger.log("There are " + j + " files.");
  }
}

// protectTab removes all existing users who have permission to the sheet, and adds the users entered in the email Group (up to 3 users max)
function protectTab() {
  var i = 0;
  var j = 0;
  Logger.log(sheetname);
  if (numEmails ==1) {
    var emailGroup = emailGroup1;
  }
  else if (numEmails ==2) {
    var emailGroup = emailGroup2;
  }
  else {
    var emailGroup = emailGroup3;
  }


  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        var spreadsheet = ss.getSheetByName(sheetname);
        var protection = spreadsheet.protect();
        // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
        // permission comes from a group, the script throws an exception upon removing the group.
        var me = Session.getEffectiveUser();
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
        protection.addEditors(emailGroup);
        Logger.log(j+1, ".", filename, spreadsheet, "is protected for me and Email Group");
        j++;
      }
    }
    Logger.log("There are " + i + " folders.");
    Logger.log("There are " + j + " files.");
  }
  else {
      var fileIter = parentFolder.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        var spreadsheet = ss.getSheetByName(sheetname);
        var protection = spreadsheet.protect();
        // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
        // permission comes from a group, the script throws an exception upon removing the group.
        var me = Session.getEffectiveUser();
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
        protection.addEditors(emailGroup);
        Logger.log(j+1, ".", filename, spreadsheet, "is protected for me and Email Group");
        j++;
      }
      Logger.log("There are " + i + " folders.");
      Logger.log("There are " + j + " files.");
    }
}


// removeProtect removes all editors from the protected tab (basically unprotects the sheet so all editors of the Spreadsheet can edit the tab)
function removeProtect() {
  var i = 0
  var j = 0

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        var spreadsheet = ss.getSheetByName(DupSheet).protect().remove();
        // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
        // permission comes from a group, the script throws an exception upon removing the group.
        Logger.log(j, ".", filename, "has been unprotected");
        j++;
      }
    }
    Logger.log("There are " + i + " subfolders.");
    Logger.log("There are " + j + " files.");
  }
  else {
      var fileIter = parentFolder.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        var spreadsheet = ss.getSheetByName(DupSheet).protect().remove();
        // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
        // permission comes from a group, the script throws an exception upon removing the group.
        Logger.log(j, ".", filename, "has been unprotected");
        j++;
      }
      Logger.log("There are " + i + " subfolders.");
      Logger.log("There are " + j + " files.");
    }
}

/*
function CopyTab() {
  var ss = SpreadsheetApp.getActive();
  var spreadsheet = ss.getSheetByName(Dupsheet).activate();
  var newsheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  Logger.log(newsheet.getName());
  newsheet.setName(RenamedSheet)
}
*/

// CopyAllFiles will copy all files in 1 folder (i.e. no subfolders)
// CopyAllFolders will check copy all files in across 1 level of subfolders (i.e. Main Folder > Folders > File) if there are no subfolders, it will copy files across.

function CopyAllFiles() {
  var i = 0
  var j = 0
  //while(childFolders.hasNext()) {
    // Only uncomment the next portion if you need the subFolders from the childFolders
    //getSubFolders(child);
    var fileIter = parentFolder.getFiles();

    while(fileIter.hasNext()){
      var file = fileIter.next();
      var filename = file.getName();
      var fileId = file.getId();
      var ss = SpreadsheetApp.openById(fileId);
      SpreadsheetApp.setActiveSpreadsheet(ss)
      var itt = ss.getSheetByName(RenamedSheet);
      if (!itt) {
        var spreadsheet = ss.getSheetByName(DupSheet).activate();
        var newsheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
        newsheet.setName(RenamedSheet);
        Logger.log(j, ".", filename, " ", newsheet.getName(), "has been copied");
        var range1 = newsheet.getRange("H3");
        range1.clearContent();
        var rep =   SpreadsheetApp.getActiveSheet().createTextFinder('I confirm that the PDR').findNext().getValue();
        replaceInSheet(newsheet, find1, repl1);
        replaceInSheet(newsheet, find2, repl2);
        replaceInSheet(newsheet, find3, repl3);
        replaceInSheet(newsheet, rep, repl4);
        replaceInSheet(newsheet, find5, repl5);
        replaceInSheet(newsheet, find6, repl6);
        replaceInSheet(newsheet, find7, repl7);
        var row = findRow3(repl4);
        var cell1 = 'K'+row;
        var cell2 = 'N'+row;
        //Logger.log(cell1, cell2);
        if (row != null) {
          var rangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRangeList([cell1, cell2]);
          rangeList.uncheck();
        Logger.log(newsheet.getName(), "has the checkboxes unticked");
        }
        j++;
      }
      else {
        Logger.log('Tab in ', filename, 'already exists');
        j++;
      }
    }

}

// CopyAllFiles will copy all files in 1 folder (i.e. no subfolders)
// CopyAllFolders will copy all files in across 1 level of subfolders (i.e. Main Folder > Folders > File) if there are no subfolders, it will copy files across.

function CopyAllFolders() {
  var i = 0
  var j = 0

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      // you will need to do a SubFolder.Iter and then a fileIter on the SubFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(ss)
        var itt = ss.getSheetByName(RenamedSheet);
        if (!itt) {
          var spreadsheet = ss.getSheetByName(DupSheet).activate();
          var newsheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
          newsheet.setName(RenamedSheet);
          Logger.log(j, ".", filename, " ", newsheet.getName(), "has been copied");
          var range1 = newsheet.getRange("H3");
          range1.clearContent();
          var rep =   SpreadsheetApp.getActiveSheet().createTextFinder('I confirm that the PDR').findNext().getValue();
          replaceInSheet(newsheet, find1,repl1);
          replaceInSheet(newsheet, find2, repl2);
          replaceInSheet(newsheet, find3, repl3);
          replaceInSheet(newsheet, rep, repl4);
          replaceInSheet(newsheet, find5, repl5);
          replaceInSheet(newsheet, find6, repl6);
          replaceInSheet(newsheet, find7, repl7);
          var row = findRow3(repl4);
          //Logger.log(row);
          var cell1 = 'K'+row;
          var cell2 = 'N'+row;
          //Logger.log(cell1, cell2);
          if (row != null) {
            var rangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRangeList([cell1, cell2]);
            rangeList.uncheck();
          }
          j++;
          Logger.log(newsheet.getName(), "has the checkboxes unticked");
        }
      }
    }
  }
  else {
    CopyAllFiles();
  }
}

function findRow3(to_find){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var info = to_find;
  for(var i = 0; i<data.length;i++){
    //for(var j = 0; j<data.length; j++) {
    if(data[i][7] == info){ //[1] because column B
      Logger.log((i+1))
      return i+1;
      }
    //}
  }
}

// TrackFiles grabs the scores from all files in folder
function TrackFiles() {
  var fileIter = parentFolder.getFiles();
  while(fileIter.hasNext()){
    var file = fileIter.next();
    var filename = file.getName();
    var fileId = file.getId();
    //Logger.log(fileId);
    var ss = SpreadsheetApp.openById(fileId);
    SpreadsheetApp.setActiveSpreadsheet(ss);
    if (ss.getSheetByName(trackTab) != null) {
      var source = ss.getSheetByName(trackTab).activate();
      var empid = source.getRange("D2").getValue();
      var empname = source.getRange("F2").getValue();
      var supname = source.getRange("F3").getValue();
      var Culturerow = source.createTextFinder('CULTURAL ALIGNMENT RATING').findNext().getRow();
      var OKRrow = source.createTextFinder('OKR SCORE').findNext().getRowIndex();
      var Completerow = source.createTextFinder('I confirm that the PDR has been discussed and').findNext().getRowIndex();
      Logger.log(Culturerow, OKRrow, Completerow);
      var empOKR = source.getRange('K'+OKRrow).getValue();
      var supOKR = source.getRange('N'+OKRrow).getValue();
      var empCult = source.getRange('K'+Culturerow).getValue();
      var supCult = source.getRange('N'+Culturerow).getValue();
      var empComp = source.getRange('K'+Completerow).getValue();
      var supComp = source.getRange('N'+Completerow).getValue();
      // Put all values into a 1 row array
      var values = [[fileId,empid, empname, supname, empCult, supCult, empOKR, supOKR, empComp, supComp]];
      var length = values[0].length;
      // Define destination sheet and last row
      var dest = SpreadsheetApp.openById(trackerFile).getSheetByName(trackerSheet);
      var lastrow = dest.getLastRow();
      //Logger.log(empid, empname, supname, empCult, supCult, empOKR, supOKR, empComp, supComp, lastrow, length)
      // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
      // Set value from all the variables previously
      var destrange = dest.getRange(lastrow+1, 1, 1,length);
      destrange.setValues(values);
      Logger.log(destrange + values + "has been tracked");
      j++;
    }
    else {
      var missing = [[filename, fileId]]
      var length = missing[0].length;
      // Define destination sheet and last row
      var dest = SpreadsheetApp.openById(trackerFile).getSheetByName(missingSheet);
      var lastrow = dest.getLastRow();
      //Logger.log(filename, fileId, "has no" PDR Form - Q2 sheet")
      // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
      // Set value from all the variables previously
      var destrange = dest.getRange(lastrow+1, 1, 1,length);
      destrange.setValues(missing);
      j++;
      Logger.log(j + filename + fileId + "does not have a" + trackTab + "sheet");
    }
  }
}


// TrackAll grabs files from parent folder and all subfolders
function TrackAll() {
  // Clear previous content first
  clearTracker(missingSheet);
  clearTracker(trackerSheet);
  // Main function
  var i = 0
  var j = 0
  if (childFolders.hasNext()){
    while(childFolders.hasNext()) {
    var child = childFolders.next();
    i++;
    Logger.log(child.getName());
    // Only uncomment the next portion if you need the subFolders from the childFolders
    //getSubFolders(child);
    var fileIter = child.getFiles();
    while(fileIter.hasNext()){
      var file = fileIter.next();
      var filename = file.getName();
      var fileId = file.getId();
      //Logger.log(fileId);
      var ss = SpreadsheetApp.openById(fileId);
      SpreadsheetApp.setActiveSpreadsheet(ss);
      if (ss.getSheetByName(trackTab) != null) {
        var source = ss.getSheetByName(trackTab).activate();
        var empid = source.getRange("D2").getValue();
        var empname = source.getRange("F2").getValue();
        var supname = source.getRange("F3").getValue();
        var Culturerow = source.createTextFinder('CULTURAL ALIGNMENT RATING').findNext().getRow();
        var OKRrow = source.createTextFinder('OKR SCORE').findNext().getRowIndex();
        var Completerow = source.createTextFinder('I confirm that the PDR has been discussed and').findNext().getRowIndex();
        Logger.log(Culturerow, OKRrow, Completerow);
        var empOKR = source.getRange('K'+OKRrow).getValue();
        var supOKR = source.getRange('N'+OKRrow).getValue();
        var empCult = source.getRange('K'+Culturerow).getValue();
        var supCult = source.getRange('N'+Culturerow).getValue();
        var empComp = source.getRange('K'+Completerow).getValue();
        var supComp = source.getRange('N'+Completerow).getValue();
        // Put all values into a 1 row array
        var values = [[fileId,empid, empname, supname, empCult, supCult, empOKR, supOKR, empComp, supComp]];
        var length = values[0].length;
        // Define destination sheet and last row
        var dest = SpreadsheetApp.openById(trackerFile).getSheetByName(trackerSheet);
        var lastrow = dest.getLastRow();
        //Logger.log(empid, empname, supname, empCult, supCult, empOKR, supOKR, empComp, supComp, lastrow, length)
        // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
        // Set value from all the variables previously
        var destrange = dest.getRange(lastrow+1, 1, 1,length);
        destrange.setValues(values);
        Logger.log(destrange + values + "has been tracked");
        j++;
      }
      else {
        var missing = [[filename, fileId]]
        var length = missing[0].length;
        // Define destination sheet and last row
        var dest = SpreadsheetApp.openById(trackerFile).getSheetByName(missingSheet);
        var lastrow = dest.getLastRow();
        //Logger.log(filename, fileId, "has no" PDR Form - Q2 sheet")
        // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
        // Set value from all the variables previously
        var destrange = dest.getRange(lastrow+1, 1, 1,length);
        destrange.setValues(missing);
        j++;
        Logger.log(j + filename + fileId + "does not have a" + trackTab + "sheet");
      }
    }
  }
}
  else {
    TrackFiles();
  }
}


function GetFolderId() {
  var i = 0
  var j = 0
  while(childFolders.hasNext()) {
    var child = childFolders.next();
    var childId = child.getId();
    var fileIter = child.getFiles();

    while(fileIter.hasNext()) {
      var file = fileIter.next()
      j += 1
    }
    Logger.log(child.getName() +  "," + childId + ", has," +  j + ", files" )
    i++;
//    Logger.log(child.getName(), childId);
  }
//  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setActiveSelection('G1:G');
//  range.setValue(childId);
}

// Track2019 grabs the scores from all files in in 2019 folder
function Track2019() {
  var parentFolder = DriveApp.getFolderById('1vZP_-A0BVXGsnOhMeQZr5wSNK4IKZUce');
  var fileIter = parentFolder.getFiles();
  while(fileIter.hasNext()){
    var file = fileIter.next();
    var filename = file.getName();
    var fileId = file.getId();
    //Logger.log(fileId);
    var ss = SpreadsheetApp.openById(fileId);
    SpreadsheetApp.setActiveSpreadsheet(ss);
    if (ss.getSheetByName("MANAGER EVALUATION") != null) {
      var source = ss.getSheetByName('MANAGER EVALUATION').activate();
      var empname = source.getRange("C2").getValue();
      var supname = source.getRange("C3").getValue();
      var Culturerow = source.createTextFinder('CULTURAL ALIGNMENT RATING').findNext().getRow();
      var OKRrow = source.createTextFinder('OKR PERFORMANCE RATING').findNext().getRowIndex();
      var Completerow = source.createTextFinder('OVERALL RATING').findNext().getRowIndex();
      Logger.log(Culturerow, OKRrow, Completerow);
      var empid = source.getRange('E'+Completerow).getValue();
//      var empOKR = source.getRange('K'+OKRrow).getValue();
      var supOKR = source.getRange('G'+OKRrow).getValue();
//      var empCult = source.getRange('K'+Culturerow).getValue();
      var supCult = source.getRange('G'+Culturerow).getValue();
//      var empComp = source.getRange('K'+Completerow).getValue();
      var supComp = source.getRange('G'+Completerow).getValue();
      // Put all values into a 1 row array
      var values = [[fileId,empid, empname, supname, supCult, supOKR, supComp]];
      var length = values[0].length;
      // Define destination sheet and last row
      var dest = SpreadsheetApp.openById('1pzu3R4s7AFhOWGfpd5cXvT7d6betArbNxmt1S9SNh0E').getSheetByName('Results2019');
      var lastrow = dest.getLastRow();
      //Logger.log(empid, empname, supname, empCult, supCult, empOKR, supOKR, empComp, supComp, lastrow, length)
      // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
      // Set value from all the variables previously
      var destrange = dest.getRange(lastrow+1, 1, 1,length);
      destrange.setValues(values);
      Logger.log(destrange, values, "has been tracked");
      j++;
    }
    else {
      var missing = [[filename, fileId]]
      var length = missing[0].length;
      // Define destination sheet and last row
      var dest = SpreadsheetApp.openById(trackerFile).getSheetByName(missingSheet);
      var lastrow = dest.getLastRow();
      //Logger.log(filename, fileId, "has no" PDR Form - Q2 sheet")
      // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
      // Set value from all the variables previously
      var destrange = dest.getRange(lastrow+1, 1, 1,length);
      destrange.setValues(missing);
      j++;
      Logger.log(j, filename, fileId, "does not have a manager eval", "sheet");
    }
  }
}

function clearTracker(sheet_name) {
  var s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var lastrow = s1.getLastRow();
  s1.insertRowsAfter(lastrow, lastrow+10);
  s1.activate().deleteRows(2, lastrow);
}


function checkSheetExists() {
  var i = 0
  var j = 0

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      // you will need to do a SubFolder.Iter and then a fileIter on the SubFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(ss)
        var itt = ss.getSheetByName(DupSheet);
        if (!itt) {
        var missing = [[filename, fileId]]
        var length = missing[0].length;
        // Define destination sheet and last row

        var dest = SpreadsheetApp.openById(trackerFile).getSheetByName('No Q2 Form');
        var lastrow = dest.getLastRow();
        //Logger.log(filename, fileId, "has no" PDR Form - Q2 sheet")
        // Define destination range, getting the last row + 1, starting with Col A, 1 row, and length of the values
        // Set value from all the variables previously
        var destrange = dest.getRange(lastrow+1, 1, 1,length);
        destrange.setValues(missing);
          }
          j++;
        }
      }
    }
  }


/*function test() {
  var sheet_name = 'TrackTest';
  var s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var lastrow = s1.getLastRow();
  var s2 = s1.getRange('A2').activate()
  s1.insertRowsAfter(lastrow, 30);
*/


function test4() {
      //Logger.log(fileId);
      var ss = SpreadsheetApp.openById('1sXzn_CINLdqyNcKDWbgMMfjcVpe7M4uruB8hFR0BtmI');
      SpreadsheetApp.setActiveSpreadsheet(ss);

      if (ss.getSheetByName(RenamedSheet) != null) {
        // Get the correct sheet
        var sheet = ss.getSheetByName(RenamedSheet);

        // Define the index rows with the scores
        var Culturerow = sheet.createTextFinder('CULTURAL ALIGNMENT RATING').findNext().getRowIndex();
        var OKRrow = sheet.createTextFinder('OKR SCORE').findNext().getRowIndex();
        var OKRtoprow = sheet.createTextFinder('COMMITTED TARGET').findNext().getRowIndex();
        var OKRtitlerow = OKRtoprow - 1;
        var OKRnumrows = OKRrow - (OKRtoprow + 1);
        Logger.log(Culturerow, OKRrow, OKRtoprow, OKRnumrows)

        // Define the cells that need formula changes
        var empCultscore = sheet.getRange('K'+Culturerow);
        var supCultscore = sheet.getRange('N'+Culturerow);
        var empCultperc = sheet.getRange('L'+Culturerow)
        var supCultperc = sheet.getRange('O'+Culturerow)
        var empOKRscore = sheet.getRange('K'+OKRrow);
        var supOKRscore = sheet.getRange('N'+OKRrow);
        var empOKRperc = sheet.getRange('L'+OKRrow);
        var supOKRperc = sheet.getRange('O'+OKRrow);
        var empOKRweight = sheet.getRange(OKRtoprow+1, 12, OKRnumrows);
        Logger.log(empOKRweight.getA1Notation());

        var v1 = 'this';
        var v2 = [];
        for (var i = 0; i < 5 ; i++){
          //for (var j = 1; j < 3; j++) {
            v2[i] += v1 ;
          //}
        }
        Logger.log(v2);
      }
}


function changeFormulas() {
  // Iterating through Folders
  var i = 0
  var j = 0
  while(childFolders.hasNext()) {
    var child = childFolders.next();
    i++;
    Logger.log(child.getName());
    // Only uncomment the next portion if you need the subFolders from the childFolders
    //getSubFolders(child);
    var fileIter = child.getFiles();

    while(fileIter.hasNext()){
      var file = fileIter.next();
      var filename = file.getName();
      var fileId = file.getId();
      //Logger.log(fileId);
      var ss = SpreadsheetApp.openById(fileId);
      SpreadsheetApp.setActiveSpreadsheet(ss);

      if (ss.getSheetByName(RenamedSheet) != null) {
        // Get the correct sheet
        var sheet = ss.getSheetByName(RenamedSheet);

        // Define the index rows with the scores
        var Culturerow = sheet.createTextFinder('CULTURAL ALIGNMENT RATING').findNext().getRowIndex();
        var OKRrow = sheet.createTextFinder('OKR SCORE').findNext().getRowIndex();
        var OKRtoprow = sheet.createTextFinder('COMMITTED TARGET').findNext().getRowIndex();
        var OKRtitlerow = OKRtoprow - 1;
        var OKRnumrows = OKRrow - (OKRtoprow + 1);
        Logger.log(Culturerow, OKRrow, OKRtoprow, OKRnumrows)

        // Define the cells that need formula changes
        var empCultscore = sheet.getRange('K'+Culturerow);
        var supCultscore = sheet.getRange('N'+Culturerow);
        var empOKRscore = sheet.getRange('K'+OKRrow);
        var supOKRscore = sheet.getRange('N'+OKRrow);
        var empOKRperc = sheet.getRange('L'+OKRrow);
        var supOKRperc = sheet.getRange('O'+OKRrow);
        //var empOKRweight = sheet.getRange(OKRtoprow, 11, OKRnumrows);

        // This sets the formula to be the sum of the 18 rows above K/N of employee/sup cult score
        empCultscore.setFormulaR1C1("=COUNTIF(R[-18]C[0]:R[-1]C[0],\"Yes\")/18*4");
        supCultscore.setFormulaR1C1("=COUNTIF(R[-18]C[0]:R[-1]C[0],\"Yes\")/18*4");
        empOKRscore.setFormulaR1C1("=SUMPRODUCT(R[-"+OKRnumrows+"]C[0]:R[-1]C[0], R[-"+OKRnumrows+"]C[1]:R[-1]C[1])");
        supOKRscore.setFormulaR1C1("=SUMPRODUCT(R[-"+OKRnumrows+"]C[0]:R[-1]C[0], R[-"+OKRnumrows+"]C[1]:R[-1]C[1])");
        empOKRperc.setFormulaR1C1("=SUM(R[-"+OKRnumrows+"]C[0]:R[-1]C[0])");
        supOKRperc.setFormulaR1C1("=SUM(R[-"+OKRnumrows+"]C[0]:R[-1]C[0])");
        empOKRperc.setNumberFormat('##%');
        supOKRperc.setNumberFormat('##%');


        // Change score weights for Cultural Alignment
        // var empCultweight = sheet.getRange("L7:L12");
        // var supCultweight = sheet.getRange("O7:O12");
        // var weights = [[1/6],[1/6], [1/6],[1/6],[1/6], [1/6]];
        // var formats = [['##.#%'],['##.#%'], ['##.#%'],['##.#%'],['##.#%'], ['##.0%']];
        // empCultweight.setValues(weights);
        // supCultweight.setValues(weights);
        // empCultweight.setNumberFormats(formats);
        // supCultweight.setNumberFormats(formats);

        Logger.log(filename + fileId + "has been changed");
      }
      else {
        Logger.log(filename + fileId + "does not have" + RenamedSheet);
      }
    }
   }
}




function addFolderPerms() {
  // Open supervisor folder, get data range
  var s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Supsheet).activate();
  var data = s1.getDataRange();
  var values = data.getValues();
  var lastrow = data.getLastRow();
  var folderIdcol = data.createTextFinder('FolderId').findNext().getColumn()-1;
  var supemailcol = (data.createTextFinder('Supervisor email').findNext().getColumn())-1;

  //Logger.log(data, lastrow, folderIdcol, supemailcol, values[1][0]);
  // for each row in the supervisor folder
  for (var i = 1; i < lastrow; i++) {
    // Open folderid by the value in row i and column folderId
    // Give editing permissions to supervisor email in row i and column supervisor email
    // Log row, sup email, and permissions granted
    var fid = values[i][folderIdcol];
    var supemail = values[i][supemailcol];
    var ff = DriveApp.getFolderById(fid);
    if (supemail != "NA") {
      ff.addEditor(supemail)
      Logger.log(supemail + "has been added to" + fid);
      }
    }
}

/* Portion of code to change text by defining range and set the text values
function changeText() {
        // Change some text as well
        var title = sheet.getRange('D1');
        var titleCultemp = sheet.getRange('K5');
        var titleCultsup = sheet.getRange('N5');
        var titleOKR = sheet.getRange('K'+(OKRtoprow-1));

        // Set text values
        title.setValue('QQ2 PERFORMANCE DEVELOPMENT & REVIEW - 2021');
        titleCultemp.setValue('Q2 SELF EVALUATION');
        titleCultsup.setValue('Q2 MANAGER EVALUATION');
        titleOKR.setValue('Q2 UPDATE');
}
*/

function test() {
  var sheet_name = 'TrackTest';
  var s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var lastrow = s1.getLastRow();
  var s2 = s1.getRange('A2').activate()
  s1.insertRowsAfter(lastrow, 30);
  //s1.insertRowsAfter(s1.getMaxRows(), 15);
  //s2 = s1.getRange('A'+lastrow);
  //s2.insertRows(2, 500)
  //s1.deleteRows(2, s1.getLastRow());
    //var range = s1.getRange(2, 1,s1.getLastRow(), s1.getLastColumn());
  //range.clearContent();
}

function test2(){
var folderId = macrosheet.getRange('B'+rowFolderrow).getValue();
var sheetname = macrosheet.getRange('B'+rowSheetname).getValue();
var trackerFile = macrosheet.getRange('B'+rowTrackerFile).getValue();
var trackerSheet = macrosheet.getRange('B'+rowTrackerSheet).getValue();
var missingSheet = macrosheet.getRange('B'+rowMissingSheet).getValue();
var selEmail = macrosheet.getRange('B'+rowSelEmail).getValue();
  Logger.log(folderId, sheetname, trackerFile, trackerSheet, missingSheet, selEmail);
}

// Function to move file (id) to supervisor name (foldername) in parentfolder (targetFolderId)
function fileMover(id, foldername, targetFolderId) {
  var file = DriveApp.getFileById(id);
  var newFolderId = FolderExists(foldername, targetFolderId);
  var newFolder = DriveApp.getFolderById(newFolderId);
  file.moveTo(newFolder);
  Logger.log("id is " + id + ", foldername is " + foldername + ", newfolderId is " + newFolderId + " & " + "file has been moved");
}

// Function to see if foldername exists in parent folder, if not create folder
function FolderExists(foldername, targetFolderId){
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var folderfind = targetFolder.getFoldersByName(foldername);
    if (folderfind.hasNext()){
       var folderfind = folderfind.next().getId();
     } else {
       var folderfind = targetFolder.createFolder(foldername).getId();
     }
  Logger.log("Return Folder: "+ folderfind);
  return folderfind
}

// Iterate down sheet and move files from Old Supervisor to New Supervisor
function movefiles() {
  // Go down spreadsheet Test
  var sheet = SpreadsheetApp.getActive().getSheetByName(Supsheet);
  var range = sheet.getDataRange().getValues();
  // var range = sheet.getRange("A:G");
  // var nRows = range.getLastRow();
  var fileidcol = sheet.createTextFinder('FileId').findNext().getColumnIndex();
  var archiveFolderId = archiveFolder; // Test Archive
  var currentFolderId = folderId; // Test Folder

// If Supervisor Name = Reporting to, do nothing
// If Reporting to = NA, move to Archive
// Else do fileMover2

  for (var j = 1; j < range.length; j++) {
    var row = j+1
    if (sheet.isRowHiddenByFilter(row) != true) {
      var curr_sup = sheet.getRange(row, 2).getValue();
      var new_sup = sheet.getRange(row, 7).getValue();
      var new_sup_email = sheet.getRange(row, 8).getValue();
      // do stuff
      if (new_sup == "NA") {
        var id = sheet.getRange(row, fileidcol).getValue();
        fileMover(id, curr_sup, archiveFolderId);
        // fileMover2Archive(id, archiveFolderId);
        Logger.log(j + " Fileid:" + id + ", currsup: NA, archiveFolderId:" + archiveFolderId);
      }
      else if (new_sup != curr_sup) {
        var id = sheet.getRange(row, fileidcol).getValue();
        fileMover(id, new_sup, currentFolderId);
        var file = SpreadsheetApp.openById(id);
        file.addEditor(new_sup_email);
        Logger.log(j + " Fileid:" + id + " , currsup:" + curr_sup + ", newsup:" + new_sup + ", email:" + new_sup_email
        + ", currentfolderid:" + currentFolderId);
      }
      // Logger.log(j + " Fileid:" + id + " , currsup:" + curr_sup + ", newsup:" + new_sup + ", email:" + new_sup_email);
    }
  }
}

// Var File Iteration code, to put in right place
  // var fileIter = child.getFiles();
  // while(fileIter.hasNext()){
  //   var file = fileIter.next();
  //   var filename = file.getName();
  //   var fileId = file.getId();

// Create array with folder IDs and folder names under defined parentFolder
function FolderIdName() {
  // Enter Cell with FolderIds
  // Get SubFolder Names and SubFolder FolderIds
  // Copy into Array - Subfolder Name, Subfolder ID

  // Create empty array for folders
  var parentFolder = DriveApp.getFolderById(folderId);
  var childFolders = parentFolder.getFolders();
  var sheet = SpreadsheetApp.getActive().getSheetByName(Supsheet);
  sheet.clear();
  var rows = [];
  rows.push(["S/N", "Supervisor Name", "FolderId", "No. of Files", "File Name", "FileId"]);
  var i = 1;
  // if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var j = 0;
      var child = childFolders.next();
      var fileIter = child.getFiles();
        while(fileIter.hasNext()) {
          var file = fileIter.next();
          j++;
          rows.push([i, child.getName(), child.getId(), j, file.getName(), file.getId()]);
        }
      i++;
      Logger.log("Row copied for " + child.getName());
    }
  sheet.getRange(1,1,rows.length, rows[0].length).setValues(rows);
}

// Get current supervisor folder name of fileid by index match
function indexMatch2() {
  // var basesheet = infosheet;
  var basesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(refSupsheet);
  var sheet = SpreadsheetApp.getActive().getSheetByName(Supsheet);
  var found = [];
  var found2 = [];
  var searchData = basesheet.getDataRange().getValues();
  var findData = sheet.getDataRange().getValues();

  for (i=0; i < findData.length; i++) {
    var find = findData[i][5];
    var searchref = basesheet.createTextFinder(find).findNext();
    if (searchref !=null) {
      var searchrefrow = searchref.getRowIndex()
      var searchvalue = basesheet.getRange(searchrefrow, 9).getValue();
      var searchvalue2 = basesheet.getRange(searchrefrow, 10).getValue();
    }
    else {
      var searchvalue = "NA";
      var searchvalue2 = "NA";
    }
  found.push([searchvalue]);
  found2.push([searchvalue2]);
  }

  Logger.log(" FL=" + found.length + " FL[0]= " + found[0].length);
  Logger.log(found);
  sheet.getRange(1, 7, found.length, 1).setValues(found);
  sheet.getRange(1, 8, found2.length, 1).setValues(found2);
}


// Get current supervisor folder name of fileid by index match
function indexMatch3() {
  var basesheet = SpreadsheetApp.getActive().getSheetByName("For Merge");
  var sheet = SpreadsheetApp.getActive().getSheetByName("Test");
  var found = [];
  var searchData = basesheet.getDataRange().getValues();
  var findData = sheet.getDataRange().getValues();

  for (i = 0; i < findData.length; i++) {
    for (j = 0; j < searchData.length; j++) {

      var find = findData[i][5];
      var searchref = searchData[j][19];
      var search1 = [searchData[2][8]];
      var search2 = searchData[3][8];
      var search3 = [searchData[22][8]];
    }
  }
  search4 = "Clara Chua"
  sindex1 = searchData.indexOf(search1);
  sindex2 = searchData.indexOf(search2);
  sindex3 = searchData.indexOf(search3);
  sindex4 = searchData.indexOf(search4);
  Logger.log("S1: " + sindex1 + ", S2: " + sindex2 + ", S3: " + sindex3 + ", S4: " + sindex4);
  Logger.log("S1: " + search1 + ", S2: " + search2 + ", S3: " + search3 + ", S4: " + search4);
}


function indexMatch4() {
  var basesheet = SpreadsheetApp.getActive().getSheetByName("For Merge");
  var sheet = SpreadsheetApp.getActive().getSheetByName("Test");
  var found = [];
  var search = basesheet.getDataRange().getValues();
  var find = sheet.getRange(1, 6, sheet.getLastRow()).getValues().flat().filter(e => e);
  for (j = 0; j < search.length; j++) {
    if(~find.indexOf(search[j][18]))found.push([search[j][8]]);
  }
  // Logger.log(i + " Found length is " + found.length);
  // Logger.log(found);
  sheet.getRange(1, 7, found.length, 1).setValues(found)
}


// Get current supervisor folder name of fileid by index match
function indexMatch() {
  var basesheet = SpreadsheetApp.getActive().getSheetByName("For Merge");
  var sheet = SpreadsheetApp.getActive().getSheetByName("Test");
  var found = [];
  // found.push(["Current Sup"])
  var searchData = basesheet.getDataRange().getValues();
  var findData = sheet.getDataRange().getValues();
  var temp;

  for (i = 0; i < findData.length; i++) {
    for (j = 0; j < searchData.length; j++) {
      // var temp = null
      var find = findData[i][5];
      var searchref = searchData[j][19];
      if (find == searchref && find != "") {
        found[i] = [searchData[j][8]];
        // found.push([searchData[j][8]]);
        // break;
      }
    }
    // found.push([temp]);
  }
  found_array = [found];
  // Logger.log(i + " Found length is " + found.length);
  // Logger.log(i + " Found length is " + found.length + "; Found[0].length is " + found[0].length);
  Logger.log(found);
  Logger.log(i + " Foundarr length is " + found_array.length + "; Foundarr[0].length is " + found_array[0].length);
  // Logger.log(found_array);
  sheet.getRange(1, 7, found.length, 1).setValues(found)
}


/*
      if (find == searchref && find != "" ) {
      // if (find == searchref) {
        found[[i]] = [searchData[j][8]]
        // found.push([searchData[j][8]])
        // found.push([searchData[j][8]])
        // found.push([0][i] = searchData[j][8]
      }
*/
      // else {
      //   found[[i]] = ["NA"];
      //   // found.push(["NA"])
      // }

// Old version of FolderIdName()
function FolderIdNameOld() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Test");
  sheet.clear();
  var rows = [];
  rows.push(["S/N", "Supervisor Name", "FolderId", "No. of Files"]);
  var parFolder = DriveApp.getFolderById("1NWqsfoVvDaJZ5Vv46pR036BnB_6GxV7S");
  var childFolders = parFolder.getFolders();
  var i = 0;
    while(childFolders.hasNext()) {
    var j = 0;
    var child = childFolders.next();
    var childFiles = child.getFiles();
      while(childFiles.hasNext()) {
        var file = childFiles.next();
        j++
      }
    // if(child != null) {
    rows.push([i, child.getName(), child.getId(), j]);
    // }
    i++;
    Logger.log("Row copied for " + child.getName());
  }
  // var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setActiveSelection('G1:G');
  // range.setValue(childId);
  sheet.getRange(1,1,rows.length, rows[0].length).setValues(rows);
}


/*
function GetFolderId() {
  var parentFolder = DriveApp.getFolderById(folderId);
  var childFolders = parentFolder.getFolders();
  var i = 0
  while(childFolders.hasNext()) {
    var child = childFolders.next();
    var childId = child.getId();
    i++;
    Logger.log(child.getName(), childId);
  }
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setActiveSelection('G1:G');
  range.setValue(childId);
}
*/

function findCell(phrase) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] == phrase) {
        row = values[i][j+1];
        Logger.log(row);
        Logger.log(i); // This is your row number
      }
    }
  }
}


function findRow2(to_find){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  //var employeeName = sheet.getRange("C2").getValue();
  for(var i = 0; i<data.length;i++){
    for(var j = 0; j<data.length; j++) {
      //if(data[i][j].match(to_find)!=null) {
      if(data[i][j] == to_find){ //[1] because column B
      Logger.log((i+1))
      return i+1;
      }
    }
  }
}

function replaceInSheet(sheet, to_replace, replace_with) {
  //get the current data range values as an array
  var values = sheet.getDataRange().getValues();

  //loop over the rows in the array
  for(var row in values){

    //use Array.map to execute a replace call on each of the cells in the row.
    var replaced_values = values[row].map(function(original_value){
      return original_value.toString().replace(to_replace,replace_with);
    });

    //replace the original row values with the replaced values
    values[row] = replaced_values;
  }

  //write the updated values to the sheet
  sheet.getDataRange().setValues(values);
}

function test3() {
  //var deleteSheet = trackTab;
  //var deleteSheet = Browser.inputBox("Name of sheet to delete (for all files)");
  var i = 0
  var j = 0
  if (childFolders.hasNext()) {
    Logger.log("There are child Folders");
  }
  else {
    Logger.log("There are no child folders");
  }
  //Logger.log("parent folder", parentFolder, "childfolders", childFolders);
}

function DeleteTab(deleteSheet) {
  //var deleteSheet = trackTab;
  var deleteSheet = Browser.inputBox("Name of sheet to delete (for all files)");
  var i = 0
  var j = 0
  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        //Logger.log(filename);
        var fileId = file.getId();
        //Logger.log(fileId);
        var ss = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(ss);
        if (ss.getSheetByName(deleteSheet) != null) {
          var sheet = ss.getSheetByName(deleteSheet);
          ss.deleteSheet(sheet);
          Logger.log(j, ".", filename, " ", ss.getName(), "has been deleted");
          j++;
        }
        else {
          Logger.log(j, ".", filename, " ", ss.getName(), "does not have the sheet ", deleteSheet);
        }
      }
    }
  }
  else {
    var fileIter = parentFolder.getFiles();

    while(fileIter.hasNext()){
      var file = fileIter.next();
      var filename = file.getName();
      //Logger.log(filename);
      var fileId = file.getId();
      //Logger.log(fileId);
      var ss = SpreadsheetApp.openById(fileId);
      SpreadsheetApp.setActiveSpreadsheet(ss);
      if (ss.getSheetByName(deleteSheet) != null) {
        var sheet = ss.getSheetByName(deleteSheet);
        ss.deleteSheet(sheet);
        Logger.log(j, ".", filename, " ", deleteSheet, "has been deleted");
        j++;
      }
      else {
        Logger.log(j, ".", filename, " ", "does not have the sheet ", deleteSheet);
      }
    }
  }
}

// Function to delete range of cells
function deleteSelectedRange() {
  var i = 0
  var j = 0

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(i + ". " + child.getName());
      // Only uncomment the next portion if you need the subFolders from the childFolders
      // you will need to do a SubFolder.Iter and then a fileIter on the SubFolders
      //getSubFolders(child);
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        j++;
        // var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(ss);
        var itt = ss.getSheetByName(RenamedSheet);
        var findtext = itt.createTextFinder("What can Circles.Life do to make you").findNext();
        if (findtext != null) {
          var rowForDeletion = findtext.getRowIndex();
          var range = itt.getRange(rowForDeletion, 5, 4, 3);
          range.clear()
          Logger.log(j + " - " + fileId + " has text deleted.");
        }
        else {
          Logger.log(j + " - " + fileId + " does not have the text.")
        }
      }
    }
  }

}


function TextChanger() {
  var i = 0
  var j = 0

  if (childFolders.hasNext()) {
    while(childFolders.hasNext()) {
      var child = childFolders.next();
      i++;
      Logger.log(child.getName());
      var fileIter = child.getFiles();

      while(fileIter.hasNext()){
        var file = fileIter.next();
        var filename = file.getName();
        var fileId = file.getId();
        var ss = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(ss)
        var itt = ss.getSheetByName(RenamedSheet);
        if (itt != null) {
          itt.activate();
          var rep =   SpreadsheetApp.getActiveSheet().createTextFinder('I confirm that the PDR').findNext().getValue();
          // replaceInSheet(newsheet, find1,repl1);
          // replaceInSheet(newsheet, find2, repl2);
          // replaceInSheet(newsheet, find3, repl3);
          replaceInSheet(itt, rep, repl4);
          // replaceInSheet(newsheet, find5, repl5);
          // replaceInSheet(newsheet, find6, repl6);
          // replaceInSheet(newsheet, find7, repl7);
          j++;
          Logger.log(j + " " + fileId + " has text replaced")
        }
        else {
          j++;
          Logger.log(j + " " + fileId + " does not have the sheet")
        }
      }
    }
  }
}




/*
function ProtectSheet() {
  // Protects the sheet PDR Form Q1 only
  //var mainfolder = SpreadsheetApp.getActiveSheet().getRange('B1').getValue();
  //var sheetname = SpreadsheetApp.getActiveSheet().getRange('B2').getValue();
  var spreadsheet = ss.getActive().getSheetByName("PDR Form - Q1");
  var protection = spreadsheet.protect();

  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script throws an exception upon removing the group.
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  return;
};
*/

/*
function getSubFolders(parent) {
  parent = parent.getId();
  var childFolder = DriveApp.getFolderById(parent).getFolders();
  while(childFolder.hasNext()) {
    var child = childFolder.next();
    Logger.log(child.getName());
    getSubFolders(child);
  }
  return;
}
*/

/*
function LinkIter() {
  var urlList = ["https://drive.google.com/open?id=1YNj68ZA33V1MIEhWJY6wjTcwO2twJB58Pnk2uGMzQlg",
                 "https://drive.google.com/open?id=1_LiZWSCEHpB_UsnH3cm_WVjdCU4mMpVm-mnb6m7mYpE",
                 "https://drive.google.com/open?id=1MmToG_d_Sss9VI6ZS371JPGvQAnzvJQm9gcM69P2gIg"];
  for (var i = 0; i < urlList.length; i++) {
    var test = urlList[i];
    //var ss = SpreadsheetApp.openByUrl(urlList[i]).getSheetByName("PDR Form - Q1");
    Logger.log(test);
  }
}
*/
