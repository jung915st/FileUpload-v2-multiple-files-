
//function doGet(e) {
//  return HtmlService.createHtmlOutputFromFile('form.html');
//}

function doGet(e) {
　return HtmlService.createTemplateFromFile('form').evaluate().setTitle("檔案上傳程式");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



function uploadFiles(form) {
  
  try {
    //var sheetid = "1uoePKTZs0d7-qLVjapb_I4f3cKk159ZqlooErAH71FI";
    var sheetid="1gq3vH289_WuT1EB5br9-nuDtcWdDfr4S13Vbee7xLww";
    //var sheetName = "upload-list";
    var sheetName = "上傳清單test";
    var dropbox = "uploadbox";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    var foldername = form.myName +'_'+ form.dept +'_'+ form.sub +'_'+ form.semester +'_'+ Utilities.formatDate(new Date(),'GMT+8','yyyyMMdd');
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    var subfolder = folder.createFolder(foldername);
    
    var blob1 = form.myFile1;
    var blob2 = form.myFile2;
    
    var blobs = [];
    var fileNames = [];
    var fileUrls= [];
    
//    if(blob1.length >0) {
//      blobs.push(blob1)
//      }
//    if(blob2.length >0) {
//      blobs.push(blob2)
//      }
    
//    for (var i=0;i<=blobs.length;i++){
//      var file = blobs[i];
//      subfolder.createFile(file);
//      file.setDescription("上傳者： " + form.myName);
//      fileNames.push(file.getName());
//      fileUrls.push(file.getUrl());
//      }
      
    
    if (blob1.length >0){
        var file1 = subfolder.createFile(blob1);
        file1.setDescription("上傳者： " + form.myName);
        var fileUrl1 = file1.getUrl();
        var fileName1 = file1.getName();
        fileNames.push(file1.getName());
      } 

    
    if (blob2.length > 0){
        var file2 = subfolder.createFile(blob2);
        file2.setDescription("上傳者： " + form.myName);
        var fileUrl2 = file2.getUrl();
        var fileName2 = file2.getName();
        fileNames.push(file2.getName());
      } 


    var FileIterator = DriveApp.getFilesByName(sheetName);
    //var sheetApp = "";
    //while (FileIterator.hasNext())
    //{
    // var sheetFile = FileIterator.next();
    //  if (sheetFile.getName() == sheetName)
    //  {
    //   sheetApp = SpreadsheetApp.open(sheetFile);
    //  }
    //}
    //if(sheetApp == "")
    //{
    var sheetApp = SpreadsheetApp.openById(sheetid);
    //var sheetApp = SpreadsheetApp.create(sheetName);
    sheetApp.getSheets()[0].getRange(1, 1, 1, 10).setValues([['上傳時間','姓名','科目','年級','類別','電子信箱','檔案名稱1','檔案網址1','檔案名稱2','檔案網址2']]);
    //  }
    //sheetApp.getSheets()[0].getRange(1, 1, 1, 8).setValues([['上傳時間','姓名','科目','年級','類別','電子信箱','檔案名稱','檔案網址']]);
    var sheet = sheetApp.getSheets()[0];
    var lastRow = sheet.getLastRow();
    var targetRange = sheet.getRange(lastRow+1, 1, 1, 10)
          .setValues([[new Date().toLocaleString(),form.myName,form.sub,form.dept,form.semester,form.email,fileName1,fileUrl1,fileName2,fileUrl2]]);
    //var targetRange = sheet.getRange(lastRow+1, 1, 1, 8).setValues([[new Date().toLocaleString(),form.myName,form.sub,form.dept,form.semester,form.email,fileNames,fileUrls]]);
    

        
    return "檔案上傳成功！檔案名稱：<br>**"+fileNames+"<br>**";
    
    
  } catch (error) {
    
    return "檔案上傳失敗！ 原因："+error.toString();
  }
  
}
