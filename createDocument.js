function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Automations');
  menu.addItem('Create Lease Agreement', 'createNewLeaseAgreement')
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

var unitAddress = SpreadsheetApp.getActiveSheet().getRange(2, 6).getValue();
var destFolderName = unitAddress;
var parentFolderName = DriveApp.getFolderById('1h05K-Iaut-iUSbgJQUcPnmorF1wc5P9W').getName();
var destFolder = getOrCreateSubFolder( destFolderName, parentFolderName);

function createNewLeaseAgreement(){
  const leaseTemplate = DriveApp.getFileById('1ey1ypl7X_j7Iv0Gg1oHsCs6Z1028pIcVJYNRRn1Tpmk');
  var folderId = DriveApp.getFoldersByName(destFolder).next().getId();
  const leaseFolder = DriveApp.getFolderById(folderId);
  const inputSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const rows = inputSheet.getDataRange().getValues();  

  const leaseTermLength = rows[1][6]
  var numYears = {
    "1 year (12)": 1,
    "2 years (24)": 2,
    "3 years (36)": 3
  } 
  const leaseStartDate = new Date(rows[1][7]).toLocaleDateString();
  const year1EndDate = new Date(rows[1][8]).toLocaleDateString();
  const year2StartDate = new Date(rows[1][9]).toLocaleDateString();
  const year2EndDate = new Date(rows[1][10]).toLocaleDateString();
  const year3StartDate = new Date(rows[1][11]).toLocaleDateString();
  const year3EndDate = new Date(rows[1][12]).toLocaleDateString();
  const amountYear1 = Math.round(rows[1][13] * 100) / 100
  const amountMonth1 = Math.round(rows[1][14] * 100) / 100
  const amountYear2 = Math.round(rows[1][15] * 100) / 100
  const amountMonth2 = Math.round(rows[1][16] * 100) / 100
  const amountYear3 = Math.round(rows[1][17] * 100) / 100
  const amountMonth3 = Math.round(rows[1][18] * 100) / 100
  const firstLast = Math.round(rows[1][19] * 100) / 100
  const tmi1 = Math.round(rows[1][20] * 100) / 100
  const utilities1 = Math.round(rows[1][21] * 100) / 100
  const tmi2 = Math.round(rows[1][22] * 100) / 100
  const utilities2 = Math.round(rows[1][23] * 100) / 100
  const tmi3 = Math.round(rows[1][24] * 100) / 100
  const utilities3 = Math.round(rows[1][25] * 100) / 100
  const totalAmount1 = Math.round(rows[1][26] * 100) / 100
  const totalAmount2 = Math.round(rows[1][27] * 100) / 100
  const totalAmount3 = Math.round(rows[1][28] * 100) / 100
  
  var leaseName = rows[1][5] +  ' ' + leaseStartDate + ' Lease Agreement';
  const leaseTemplateCpy = leaseTemplate.makeCopy(leaseName, leaseFolder)
  const lease = DocumentApp.openById(leaseTemplateCpy.getId())
  const body = lease.getBody();

  var endDate = year1EndDate;

  var rentAmount = `From ${leaseStartDate} to ${year1EndDate} inclusive ${amountYear1} per annum being ${amountMonth1} per month.\n` 

  var tmi = `Year 1 -TMI (Property Tax, Maintenance and Insurance) - ${tmi1} CAD and  ${utilities1}  CAD for Utilities.\n` 
  
  var rentAmount2 = `From ${leaseStartDate}  to  ${year1EndDate} - gross rent ${amountMonth1}  CAD plus ${utilities1}  CAD for Utilities plus ${tmi1} CAD for TMI  plus HST  totaling to ${totalAmount1} per month.\n\n`


  if (numYears[leaseTermLength] >= 2) {

    endDate = year2EndDate

    rentAmount += `From ${year2StartDate} to ${year2EndDate} inclusive ${amountYear2} per annum being ${amountMonth2} per month.\n`
    
    tmi += `Year 2 -TMI (Property Tax, Maintenance and Insurance) - ${tmi2} CAD and  ${utilities2}  CAD for Utilities.\n` 
    
    rentAmount2 += `From ${year2StartDate}  to  ${year2EndDate} - gross rent ${amountMonth2}  CAD plus ${utilities2}  CAD for Utilities plus ${tmi2} CAD for TMI  plus HST  totaling to ${totalAmount2} per month.\n\n`
  

  }
  if (numYears[leaseTermLength] >= 3) {
    endDate = year3EndDate

    rentAmount += `From ${year3StartDate} to ${year3EndDate} inclusive ${amountYear3} per annum being ${amountMonth3} per month.\n`

    tmi += `Year 3 -TMI (Property Tax, Maintenance and Insurance) - ${tmi3} CAD and  ${utilities3}  CAD for Utilities.\n` 
    
    rentAmount2 += `From ${year3StartDate}  to  ${year3EndDate} - gross rent ${amountMonth3}  CAD plus ${utilities3}  CAD for Utilities plus ${tmi3} CAD for TMI  plus HST  totaling to ${totalAmount3} per month.\n\n`
  }

  body.replaceText('{{LEASEEND}}', endDate);
  body.replaceText('{{RENTAL}}', rentAmount);
  body.replaceText('{{TMI}}', tmi);
  body.replaceText('{{RENTAL2}}', rentAmount2);
  body.replaceText('{{TENANTNAME}}', rows[1][0]);
  body.replaceText('{{LANDLORDNAME}}', rows[1][1]);
  body.replaceText('{{SQFT1}}', rows[1][2]);
  body.replaceText('{{FLOOR1}}', rows[1][3]);
  body.replaceText('{{UNITADDRESS}}', rows[1][5]);
  body.replaceText('{{OFFICEUSE}}', rows[1][4]);
  body.replaceText('{{XYEARXMONTH}}', rows[1][6]);
  body.replaceText('{{LEASESTART}}', leaseStartDate);
  body.replaceText('{{FIRSTLAST}}', firstLast);
  body.replaceText('{{TENANTPHONE}}', rows[1][30]);
  body.replaceText('{{TENANTEMAIL}}', rows[1][29]);
  body.replaceText('{{LEASECONDITION}}', rows[1][31]);


  lease.saveAndClose();
  const leaseUrl = lease.getUrl();
  inputSheet.getRange(2, 33).setValue(leaseUrl) 
}
