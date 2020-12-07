function ClearContent() { 
  //Paste the importhtml and the array formulas in the import html sheet
  var SupplyCoD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Folders').getRange('A:I');
  SupplyCoD.clearContent();
};
                      

function listFilesInFolder() {
//writes the headers for the spreadsheet
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Folders");
   sheet.appendRow(["Region", "Country", "Project","Folder","URL","Date Created", "Date Updated", "Project ID"]);
   var folder1ID = "THE INITIAL OF THE LEVEL 1 FOLDER GOES HERE";
   var folder1 = DriveApp.getFolderById(folder1ID);
   var content1 = folder1.getFolders()
   var cnt = 0;
   var folders2;

    while (content1.hasNext()) {
        var folders2 = content1.next();
        cnt++;
      
        Logger.log(folders2);
        Logger.log(cnt);
      
      var folder3ID = folders2.getId();
      var folder3 = DriveApp.getFolderById(folder3ID);
      var content2 = folders2.getFolders()
      var cnt2 = 0;
      var folders3;

      
      while (content2.hasNext()) {
        var folders3 = content2.next();
        cnt++;
        
        var folder4ID = folders3.getId();
        var folder4 = DriveApp.getFolderById(folder4ID);
        var content3 = folders3.getFolders()
        var cnt3 = 0;
        var folders4;
      
        Logger.log(folders4);
        Logger.log(cnt3);
          
        while (content3.hasNext()) {
          var folders4 = content3.next();
          cnt3++;
          
          var folder5ID = folders4.getId();
          var folder5 = DriveApp.getFolderById(folder5ID);
          var content4 = folders4.getFolders()
          var cnt4 = 0;
          var folders5;
          
          while (content4.hasNext()) {
            var folders5 = content4.next();
            cnt4++;
            
            if (folders4.getName().includes('Phases')) {
              var folder6ID = folders5.getId();
              var folder6 = DriveApp.getFolderById(folder6ID);
              var content5 = folders5.getFolders()
              var cnt5 = 0;
              var folders6;
          
            while (content5.hasNext()) {
              var folders6 = content5.next();
              cnt5++;
        
            // writes the various chunks to the spreadsheet- just delete anything you don't want
            data = [
              folders2.getName(),
              folders3.getName(),
              folders5.getName(),
              folders6.getName(),
              folders6.getUrl(),
              folders6.getDateCreated(),
              folders6.getLastUpdated(),
            ];
              
              sheet.appendRow(data);
              
              }
              } else {
              
              data = [
              folders2.getName(),
              folders3.getName(),
              folders4.getName(),
              folders5.getName(),
              folders5.getUrl(),
              folders5.getDateCreated(),
              folders5.getLastUpdated(),
              ];

              sheet.appendRow(data);
              
              };
            };
          };
        };
      };
    };
            
function idarray() { 
  //Paste the importhtml and the array formulas in the import html sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Folders');
  var idarray = '=iferror(arrayformula(right(C2:C,14)),"")';
  s.getRange('H2').setValue(idarray);

};
            
