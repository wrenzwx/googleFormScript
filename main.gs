var ss = SpreadsheetApp.getActiveSpreadsheet();
var template = ss.getSheetByName('Individual Cutter Template'); 
const EMAILADD = 'zhang.wenxuan@herrenknecht.com'; //singaram.alagu@herrenknecht.com
const TBMNO = 'C';
const BATCHNO = 'D';
const TRACKNO = 'E';
const SERIALNO = 'F';
const MDATE = 'H';
const MCUTTER = 'I';
const MWEAR = 'J';
const MWEARTYPE = 'K';
const MTORQUE = 'L';
const MROTATION = 'M';
const MTIGHTNESS = 'N';
const MRFTYPE = 'O';
const MCOMMENT = 'P';

const SDATE = 'Q';
const SCUTTER = 'R';
const SWEAR = 'S';
const SWEARTYPE = 'T';
const STORQUE = 'U';
const SROTATION = 'V';
const STIGHTNESS = 'W';
const SRFTYPE = 'X';
const SCOMMENT = 'Y';

const M16S = 'Z';
const M16E = 'AO';
const M14S = 'AP';
const M14E = 'BC';
const M17S = 'BD';
const M17E = 'BT';
const S13S = 'BU';
const S13E = 'CG';

const MECHNAME = 'CH';
const RFBDATE = 'CI';
const TORQUEF = 'CJ';
const PRESSUREF = 'CK';
const ADDREMARKS = 'CM';
const SENDEMAIL = 'CN';
const BATCHFLAG = 'CO';

/**
 * main function here
 * execute trigger: Google form submition
**/

function onFormSubmit(e) {

  const values = e.namedValues;

  var info = fetchInfo(); 

  clearTemplate();
  
  preInfo(info);
  
  entryCheck(info);

  partAmount(info);
  
  exitCheck(info);

  const photoIdsBf = getPhotoIds(values["UPLOAD PHOTOS - Before"][0]);
  if (photoIdsBf.length > 0) {
    const bfFolderId = getFolderId(photoIdsBf[0]); 
    processPhotos(photoIdsBf, bfFolderId, "Before");
  }

  const photoIdsDm = getPhotoIds(values["UPLOAD PHOTOS - Dismantled"][0]);
  if (photoIdsDm.length > 0) {
    const dmFolderId = getFolderId(photoIdsDm[0]); 
    processPhotos(photoIdsDm, dmFolderId, "Dismantled");
  }
  
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf

  const pdfId = createPDF(ss.getId(), template, info);   // individual disc cutter report

  sendEmail(info,pdfId);
  
  // summayCheck(info);
  // var batchM = ss.getSheetByName('M-2857 Batch Summary'); 
  // const pdfSum = createSummary(ss.getId(), batchM);
  // sendEmail(pdfSum);

  // var batchS = ss.getSheetByName('S-1345A Batch Summary'); 
  // const pdfSum = createSummary(ss.getId(), batchS);
  // sendEmail(pdfSum);

}

function fetchInfo(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var response = ss.getSheetByName('Form responses 5');
  var lastRow = response.getLastRow();
  // Logger.log(lastRow +" lastrow fetch here");

  var tbmNo = response.getRange(TBMNO + lastRow).getValue();
  var batchNo = response.getRange(BATCHNO + lastRow).getValue();
  var trackNo = response.getRange(TRACKNO + lastRow).getValue();
  var serialNo = response.getRange(SERIALNO + lastRow).getValue();
  var emailFlag;
  var discCutter;
  var batchFlag = response.getRange(BATCHFLAG + lastRow).getValue();

  if (batchFlag=='Yes'){
    batchFlag = 1;
  }
  else{ batchFlag = 0; }

  if (response.getRange(SENDEMAIL + lastRow).getValue()=='Yes'){
    emailFlag = 1;
  }
  else{ emailFlag = 0; }

  if (tbmNo=='M-2857')
    {discCutter = response.getRange(MCUTTER + lastRow).getValue();}
  if (tbmNo=='S-1345A')
    {discCutter = response.getRange(SCUTTER + lastRow).getValue();}

  var artNr = discCutter.substring(8, 16); 

  // Logger.log(artNr);

  var mechName = response.getRange(MECHNAME + lastRow).getValue();

  if (tbmNo == 'M-2857'){
    var date = response.getRange(MDATE + lastRow).getValue();
    var cutterWear = response.getRange(MWEAR + lastRow).getValue();
    var torqueEntry = response.getRange(MTORQUE + lastRow).getValue();
    var wearType = response.getRange(MWEARTYPE + lastRow).getValue();
    var rotationEntry = response.getRange(MROTATION + lastRow).getValue();
    var tightEntry = response.getRange(MTIGHTNESS + lastRow).getValue();
    var rfbType = response.getRange(MRFTYPE + lastRow).getValue();
    var furComment = response.getRange(MCOMMENT + lastRow).getValue();
  }

  else{
    var date = response.getRange(SDATE + lastRow).getValue();
    var cutterWear = response.getRange(SWEAR + lastRow).getValue();
    var torqueEntry = response.getRange(STORQUE + lastRow).getValue();
    var wearType = response.getRange(SWEARTYPE + lastRow).getValue();
    var rotationEntry = response.getRange(SROTATION + lastRow).getValue();
    var tightEntry = response.getRange(STIGHTNESS + lastRow).getValue();
    var rfbType = response.getRange(SRFTYPE + lastRow).getValue();
    var furComment = response.getRange(SCOMMENT + lastRow).getValue();
  }

  var torqueF =  response.getRange(TORQUEF+ lastRow).getValue();
  var rfbDate =  response.getRange(RFBDATE+ lastRow).getValue();
  var pressureF =  response.getRange(PRESSUREF+ lastRow).getValue();
  var addRemarks =  response.getRange(ADDREMARKS+ lastRow).getValue();

  return {
    response: response,
    tbmNo: tbmNo,
    batchNo: batchNo,
    trackNo: trackNo,
    serialNo: serialNo,
    artNr: artNr,
    mechName: mechName,
    lastRow: lastRow,
    emailFlag: emailFlag,
    batchFlag: batchFlag,
    date: date,
    cutterWear: cutterWear,
    torqueEntry: torqueEntry,
    wearType: wearType,
    rotationEntry: rotationEntry,
    tightEntry: tightEntry,
    rfbType: rfbType,
    furComment: furComment,
    torqueF: torqueF,
    rfbDate: rfbDate,
    pressureF: pressureF,
    addRemarks: addRemarks
  };
}


function getPhotoIds(photosString) {
  return photosString.split(',').map(url => {
    const match = /id=(.*?)$/g.exec(url.trim());
    return match ? match[1] : null;
  }).filter(id => id !== null); 
}

function getFolderId(photoId) {
  const file = DriveApp.getFileById(photoId);
  const parents = file.getParents();
  if (parents.hasNext()) {
    const folder = parents.next();
    return folder.getId();
  }
  return null;
}

function processPhotos(photoIds, folderId, type) {

  let mergedCellRanges = [];
  if (type == "Before"){
    mergedCellRanges = ['A49:H55', 'I49:O55', 'A56:H62', 'I56:O62','A63:H69','I63:O69','A70:H76','I70:O76'];
  } else{
    mergedCellRanges = ['A82:H88', 'I82:O88', 'A89:H95', 'I89:O95','A96:H102','I96:O102','A103:H109','I103:O109'];
  }

  
  photoIds.forEach((id, index) => {
    Logger.log("index = " +index+" processing...")
    if (index >= mergedCellRanges.length) return;
    
    // var folder = DriveApp.getFolderById(folderId);
    // const newFileName = `${tbmNo}_Batch${batchNo}_Track${trackNo}_${type}-${index + 1}`;
    
    const file = DriveApp.getFileById(id);
    const range = template.getRange(mergedCellRanges[index]);
    const topRow = range.getRow();
    const leftCol = range.getColumn();  

    var blob = DriveApp.getFileById(id).getBlob();
    var res0 = ImgApp.getSize(blob);
    var w0 = res0.width;
    var h0 = res0.height;
    w1 = parseInt(200*w0/h0);
    // Logger.log(w0+" "+h0+" "+w1);
    
    var res = ImgApp.doResize(id, 1024); 
    template.insertImage(res.blob, leftCol, topRow,10,5).setWidth(w1).setHeight(200);

    Logger.log("INSERTED.")

    
    file.setTrashed(true);

    const root = DriveApp.getRootFolder();
    const allFilesInRoot = root.getFiles();   
    const originalName = file.getName();
    const partialName = originalName.split(' ')[0];

    while (allFilesInRoot.hasNext()) {
        const possibleMatch = allFilesInRoot.next();
        if (possibleMatch.getName().startsWith(partialName)) {
            Logger.log('Found a file with matching prefix: ' + possibleMatch.getName()+ '. Deleting...');
            possibleMatch.setTrashed(true);
        }
      }

  });
}


function preInfo(info){

  template.getRange("C2:D2").setValue(info.batchNo);
  template.getRange("C45:D45").setValue(info.batchNo);
  template.getRange("C78:D78").setValue(info.batchNo);
 
  template.getRange("L2:M2").setValue(info.trackNo);
  template.getRange("L45:M45").setValue(info.trackNo);
  template.getRange("L78:M78").setValue(info.trackNo);

  template.getRange("L3:M3").setValue(info.serialNo);  
  template.getRange("L46:M46").setValue(info.serialNo);  
  template.getRange("L79:M79").setValue(info.serialNo);  

  // TBMNo + ArtNr + part list fill in

  if(info.tbmNo == "S-1345A"){ macroS(); }
  else if(info.artNr == "29604491"){ macroM16(); }
  else if(info.artNr == "29604529"){ macroM14(); }
  else{ macroM17(); }
  

}

function entryCheck(info){

  template.getRange("G16:H16").setValue(info.date);
  template.getRange("I16:J16").setValue(info.mechName);
  template.getRange("M16:M18").setValue(info.rfbType);
  template.getRange("D16").setValue(info.cutterWear);
  template.getRange("I12").setValue(info.torqueEntry);  
  template.getRange("E18:J18").setValue(info.furComment);  

  var wearTypeS = info.wearType.split(", ");
  var rotationS = info.rotationEntry.split(", ");
  var tightS = info.tightEntry.split(", ");

  wearTypeS.forEach(function(item) {
    Logger.log(item)
    if (item.includes("Normal Wear")) {
      template.getRange("D9").setValue("✓");
    }
    if (item.includes("Cracked")) {
      template.getRange("D10").setValue("✓");  
    }
    if (item.includes("Mushrooming")) {
      template.getRange("D11").setValue("✓");  
    }
    if (item.includes("Chipping")) {
      template.getRange("D12").setValue("✓");
    }
    if (item.includes("Circlip Damage")) {
      template.getRange("D13").setValue("✓");
    }
    if (item.includes("Total Damage")) {
      template.getRange("D14").setValue("✓");
    }
    if (item.includes("Prevention")) {
      template.getRange("D15").setValue("✓");
    }
  });

  rotationS.forEach(function(item) {
    var hasA = item.includes("Smooth");
    var hasB = item.includes("Jerkily");
    var hasC = item.includes("Blocked");
    if (hasA) {
      template.getRange("I9").setValue("✓");
    }
    if (hasB) {
      template.getRange("I10").setValue("✓");  
    }
    if (hasC) {
      template.getRange("I11").setValue("✓");  
    }
    if (!hasA && !hasB && !hasC) {
      template.getRange("H13:J13").setValue(rotationS);
    }
  });

  tightS.forEach(function(item) {
    var hasA = item.includes("TIGHT");
    var hasB = item.includes("LEAKAGE");
    var hasC = item.includes("Not Carried Out");
    if (hasA) {
      template.getRange("N9").setValue("✓");
    }
    if (hasB) {
      template.getRange("N10").setValue("✓");  
    }
    if (hasC) {
      template.getRange("N11").setValue("✓");  
    }
  });

}

function partAmount(info){
  
  if(info.artNr == "29604491"){  //16
    var rowData = info.response.getRange(M16S+info.lastRow+":"+M16E+info.lastRow).getValues();
    // Logger.log("reponse: "+ rowData);
    var columnData = rowData[0].map(function(item) {
      return [item];
    });
    columnData = columnData.map(function(row) {
      return row[0] === 0 ? [''] : row;
    });
    template.getRange("J23:J38").setValues(columnData);
  }

  else if(info.artNr == "29604529"){  //14
    var rowData = info.response.getRange(M14S+info.lastRow+":"+M14E+info.lastRow).getValues();
    var columnData = rowData[0].map(function(item) {
      return [item];
    });
    columnData = columnData.map(function(row) {
      return row[0] === 0 ? [''] : row;
    });
    template.getRange("J23:J36").setValues(columnData);
  }

  else if(info.artNr == "29604480"){  //17
    var rowData = info.response.getRange(M17S+info.lastRow+":"+M17E+info.lastRow).getValues();
    var columnData = rowData[0].map(function(item) {
      return [item];
    });
    columnData = columnData.map(function(row) {
      return row[0] === 0 ? [''] : row;
    });
    template.getRange("J23:J39").setValues(columnData);
  }

  else{                  //13
    var rowData = info.response.getRange(S13S+info.lastRow+":"+S13E+info.lastRow).getValues();
    var columnData = rowData[0].map(function(item) {
      return [item];
    });
    columnData = columnData.map(function(row) {
      return row[0] === 0 ? [''] : row;
    });
    // Logger.log(columnData);
    template.getRange("J23:J35").setValues(columnData);
  }
}

function exitCheck(info){
  
  template.getRange("N26:N27").setValue(info.torqueF);
  template.getRange("M36").setValue(info.rfbDate);
  template.getRange("B42:N42").setValue(info.addRemarks);

  if (info.pressureF == 2){
    template.getRange("N29:N30").setValue("✓");  
  }
  else{template.getRange("N29:N30").setValue(info.pressureF);}

}

/**
 * Sends emails with PDF as an attachment.
 * Checks/Sets 'Email Sent' column to 'Yes' to avoid resending.
 * 
 * Called by user via custom menu item.
 */
function sendEmail(info, pdfId) {

  if(info.emailFlag == 0){return;}

  const EMAIL_SUBJECT = 'Disc Cutter Insepection Report Notification';
  const EMAIL_BODY = 'Hello!\rPlease see the attached PDF document.'; 

  if(info.batchFlag == 1){EMAIL_SUBJECT = 'Disc Cutter Batch Summary Notification';}
  
  const attachment = DriveApp.getFileById(pdfId);

  var recipient = EMAILADD;
  var mechName = info.mechName;

  GmailApp.sendEmail(recipient, EMAIL_SUBJECT, EMAIL_BODY, {
    attachments: [attachment.getAs(MimeType.PDF)],
    name: mechName
  });

  Logger.log("email sent.");

}

function clearTemplate(){

  const rngClear = template.getRangeList(['A1:O1','C2:D2', 'G2:H3', 'L2:M3', 'D9:D16','I9:I12', 'N9:N11','G16:J16','M16:M18','E18:J18','A23:K39','N26:N30','M36','B42:N42','C45:D45','G45:H46','L45:M46','C78:D78','G78:H79','L78:M79','A49:O76','A82:O109']).getRanges()

  rngClear.forEach(function (cell) {
    cell.clearContent();
  });

  const ranges = [
    template.getRange('A49:O76'),
    template.getRange('A82:O109')
  ];

  const images = template.getImages();
  Logger.log("clearing pics in template...");

  images.forEach(image => {
    const anchorCell = image.getAnchorCell();
    const imageRow = anchorCell.getRow();
    const imageCol = anchorCell.getColumn();

    for (let i = 0; i < ranges.length; i++) {
      const range = ranges[i];
      if (imageRow >= range.getRow() && imageRow <= range.getLastRow() &&
          imageCol >= range.getColumn() && imageCol <= range.getLastColumn()) {
        image.remove();
        break;
      }
    }
  });

}


function createPDF(ssId, sheet,info) {
  const fr = 0, fc = 0, lc = 15, lr = 111;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=A4&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=false&" +
    "horizontal_alignment=CENTER&"+
    "gridlines=false&" +
    "printtitle=false&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  var pdfName = `${info.tbmNo}_Batch#${info.batchNo}_Track#${info.trackNo}_DCIR`;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = DriveApp.getFolderById('1JkxYNvGQdpNpspUPIz0TejHSs3NXNzY8');

  const pdfFile = folder.createFile(blob);
  pdfId = pdfFile.getId();
  return pdfId;
}

function createSummary(){

}

function summayCheck(info){

  if (info.batchFlag == 0){return;}
  if (info.tbmNo == 'M-2857'){
    var batchSum = ss.getSheetByName('M-2857 Batch Summary'); 


  }

  if(info.tbmNo == 'S-1345A'){
    var batchSum = ss.getSheetByName('S-1345A Batch Summary'); 
  }


}

function copyResponses() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = spreadsheet.getSheetByName("Form responses 5"); 
  var batchSum = spreadsheet.getSheetByName("M-2857 Batch Summary"); 
  
  var responses = responseSheet.getDataRange().getValues(); 
  
  for (var i = 1; i < responses.length; i++) {
    var response = responses[i];
    if (response[1] == 2) { // 检查第B列是否等于2
      var dataToCopy = response.slice(3, 12); // 获取D:L列的数据
      
      // 找到模板工作表中最后一列
      var lastColumn = batchSum.getLastColumn();
      
      // 将数据复制到模板工作表的最后一列后面新的一列的5:13行
      for (var j = 0; j < dataToCopy.length; j++) {
        batchSum.getRange(5 + j, lastColumn + 1).setValue(dataToCopy[j]);
      }
    }
  }
}



