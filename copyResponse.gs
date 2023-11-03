function copyResponses(info) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = spreadsheet.getSheetByName("Form responses 5"); 
  var responses = responseSheet.getDataRange().getValues(); 
  
  for (var i = 1; i < responses.length; i++) {
    var response = responses[i];
 
         
    if (info.tbmNo == 'S-1345A'){
      var batchSum = spreadsheet.getSheetByName("S-1345A Batch Summary"); 
      if (response[3] == info.batchNo && response[2] == info.tbmNo) {        // D = batchno?
            Logger.log("batch matched, col."+i);
      
      var serialN = response[5];
      //13 
      var rowData = response.slice(72, 85);
      var lastColumn = batchSum.getLastColumn();      
      // Logger.log("lastCol"+lastColumn);
      
      batchSum.getRange(3,lastColumn+1).setValue(serialN);

      for (var j = 0; j < rowData.length; j++) {
        batchSum.getRange(4 + j, lastColumn + 1).setValue(rowData[j]);
        // Logger.log("filling cell:"+j);
      }


    }

    if (info.tbmNo == 'M-2857'){

      var cutterType = response[8]; //I
      Logger.log('response8(I): '+ cutterType);
      var artR = cutterType.substring(8, 16); //response art nr
      Logger.log('art response: '+ artR);

      if(artR == "29604491"){  //16
        var batchSum = spreadsheet.getSheetByName("29604491_BatchSummary"); 
        var rowData = response.slice(25, 41);
        var lastColumn = batchSum.getLastColumn();      
        // Logger.log("lastCol"+lastColumn);
        
        batchSum.getRange(3,lastColumn+1).setValue(serialN);

        for (var j = 0; j < rowData.length; j++) {
          batchSum.getRange(4 + j, lastColumn + 1).setValue(rowData[j]);
          // Logger.log("filling cell:"+j);
        }
      }

      else if(artR == "29604529"){  //14
        var batchSum = spreadsheet.getSheetByName("29604529_BatchSummary"); 
        var rowData = response.slice(41, 55);
        var lastColumn = batchSum.getLastColumn();      
        // Logger.log("lastCol"+lastColumn);
        
        batchSum.getRange(3,lastColumn+1).setValue(serialN);

        for (var j = 0; j < rowData.length; j++) {
          batchSum.getRange(4 + j, lastColumn + 1).setValue(rowData[j]);
          // Logger.log("filling cell:"+j);
        }
      }

      else if(artR == "29604480"){  //17
        var batchSum = spreadsheet.getSheetByName("29604480_BatchSummary"); 
        var rowData = response.slice(55,72);
        var lastColumn = batchSum.getLastColumn();      
        // Logger.log("lastCol"+lastColumn);
        
        batchSum.getRange(3,lastColumn+1).setValue(serialN);

        for (var j = 0; j < rowData.length; j++) {
          batchSum.getRange(4 + j, lastColumn + 1).setValue(rowData[j]);
          // Logger.log("filling cell:"+j);
        }
      }

    }

  }
  

  // calculate row sum

}

