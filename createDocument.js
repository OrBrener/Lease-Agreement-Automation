
 function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Autofill Docs');
  menu.addItem('Create New Docs', 'createnewGoogleDocs')
  menu.addToUi();
}

function getOrCreateSubFolder(childFolderName, parentFolderName) {
  var parentFolder, parentFolders;
  var childFolder, childFolders;
  // Gets FolderIterator for parentFolder
  parentFolders = DriveApp.getFoldersByName(parentFolderName);
  /* Checks if FolderIterator has Folders with given name
  Assuming there's only a parentFolder with given name... */ 
  while (parentFolders.hasNext()) {
    parentFolder = parentFolders.next();
  }
  // If parentFolder is not defined it sets it to root folder
  if (!parentFolder) { parentFolder = DriveApp.getRootFolder(); }
  // Gets FolderIterator for childFolder
  childFolders = parentFolder.getFoldersByName(childFolderName);
  /* Checks if FolderIterator has Folders with given name
  Assuming there's only a childFolder with given name... */ 
  while (childFolders.hasNext()) {
    childFolder = childFolders.next();
  }
  // If childFolder is not defined it creates it inside the parentFolder
  if (!childFolder) { parentFolder.createFolder(childFolderName); }
  return childFolder;
}

var UnitAddressFromSheet = SpreadsheetApp.getActiveSheet().getRange(2, 6).getValue();
var destFolderName = UnitAddressFromSheet;

var destFolder = getOrCreateSubFolder( destFolderName, DriveApp.getFolderById('1h05K-Iaut-iUSbgJQUcPnmorF1wc5P9W').getName());

function createnewGoogleDocs(){
  const googleDocTemplate = DriveApp.getFileById('1kkZixmy3M2ilWZWDb9GCKCxV0hiJ6zUZFQYZuGPXCuY');
  var id = DriveApp.getFoldersByName(destFolder).next().getId();
  const destinationfolder = DriveApp.getFolderById(id);
  const sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const rows = sheet.getDataRange().getValues();

  rows.forEach(function(row, index) {
  if (index === 0) return;
  if (row[32]) return; 
  


  const friendlyDate = new Date(row[7]).toLocaleDateString();
  const friendlyDate1 = new Date(row[8]).toLocaleDateString();
  const friendlyDate2 = new Date(row[9]).toLocaleDateString();
  const friendlyDate3 = new Date(row[10]).toLocaleDateString();
  const friendlyDate4 = new Date(row[11]).toLocaleDateString();
  const friendlyDate5 = new Date(row[12]).toLocaleDateString();

  var name = row[5] +  friendlyDate + ' Lease Agreement';

  const copy = googleDocTemplate.makeCopy(name, destinationfolder)
  const doc = DocumentApp.openById(copy.getId())
  const body = doc.getBody();
    body.replaceText('{{TENANTNAME}}', row[0]);
    body.replaceText('{{LANDLORDNAME}}', row[1]);
    body.replaceText('{{SQFT1}}', row[2]);
    body.replaceText('{{FLOOR1}}', row[3]);
    body.replaceText('{{UNITADDRESS}}', row[5]);
    body.replaceText('{{OFFICEUSE}}', row[4]);
    body.replaceText('{{XYEARXMONTH}}', row[6]);
    body.replaceText('{{LEASESTART1}}', friendlyDate);
    body.replaceText('{{LEASEEND1}}', friendlyDate1);
    body.replaceText('{{LEASESTART2}}',friendlyDate2);
    body.replaceText('{{LEASEEND2}}', friendlyDate3);
    body.replaceText('{{LEASESTART3}}', friendlyDate4);
    body.replaceText('{{LEASEEND3}}', friendlyDate5);
    body.replaceText('{{AMTYEAR1}}', Math.round(row[13] * 100) / 100);
    body.replaceText('{{AMTMONTH1}}', Math.round(row[14] * 100) / 100);
    body.replaceText('{{AMTYEAR2}}', Math.round(row[15] * 100) / 100);
    body.replaceText('{{AMTMONTH2}}', Math.round(row[16] * 100) / 100);
    body.replaceText('{{AMTYEAR3}}', Math.round(row[17] * 100) / 100);
    body.replaceText('{{AMTMONTH3}}', Math.round(row[18] * 100) / 100);
    body.replaceText('{{FIRSTLAST}}', Math.round(row[19] * 100) / 100);
    body.replaceText('{{TM11}}', Math.round(row[20] * 100) / 100);
    body.replaceText('{{UTILIT11}}', Math.round(row[21] * 100) / 100);
    body.replaceText('{{TM12}}', Math.round(row[22] * 100) / 100);
    body.replaceText('{{UTILIT12}}', Math.round(row[23] * 100) / 100);
    body.replaceText('{{TM13}}', Math.round(row[24] * 100) / 100);
    body.replaceText('{{UTILIT13}}', Math.round(row[25] * 100) / 100);
    body.replaceText('{{TOTALAMOUNT1}}', Math.round(row[26] * 100) / 100);
    body.replaceText('{{TOTALAMOUNT2}}', Math.round(row[27] * 100) / 100);
    body.replaceText('{{TOTALAMOUNT3}}', Math.round(row[28] * 100) / 100);
    body.replaceText('{{TENANTPHONE}}', row[29]);
    body.replaceText('{{TENANTEMAIL}}', row[30]);
    body.replaceText('{{LEASECONDITION}}', row[31]);

    doc.saveAndClose();
    const url = doc.getUrl();
    sheet.getRange(index + 1, 33).setValue(url) 
  })

}
