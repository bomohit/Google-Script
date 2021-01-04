function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Create Individual Sheet and Spreadsheet [1]", functionName: "addSheets"}
  ];
  ss.addMenu("EXTRA", menuEntries);
}

//FiFo

function addSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var rData = sh.getDataRange().getValues();
  var arr = []; // saving name that already use
  var emailArr = []; // saving email
  var urlArr = []; // save spreadsheet url
  var sheetID = [];
  var mainUrl = ss.getUrl();
  let cat1 = "NAME";
  
  var message = [];    
  for(var i=1, len=rData.length; i<len; i++) { // ENABLE THIS FOR FULL SCAN
//  for(var i=1, len=2; i<len; i++) { // for testing only: scan 2 only
    if(rData[i][0] != "" || rData[i][1] != "" || rData[i][2] != "") { 
//      ss.toast(i);
      if(arr.length== 0 || (arr.indexOf(rData[i][2])+1) == 0){
        try {
//          let sName = rData[i][3];
          let sName = rData[i][2];
//          ss.toast(arr[i]);
          ss.insertSheet(sName);
          arr.push(sName); // stored the created sheet name
          emailArr.push(rData[i][4]); // stored the email
          ss.setActiveSheet(ss.getSheets()[0]);
          let sBlock = rData[i][0];
          // create individual sheet
          //copy the first row (header)
          var source = sh.getRange("A1:BH1");
          ss.setActiveSheet(ss.getSheetByName(sName));
//          source.copyTo(ss.getRange("A1:BH1"));
          ss.getRange("A"+ 1).setFormula("=filter(DATA!A1:BY1,DATA!D1=\"Name\")");
          if (sBlock == 1) {
           ss.getRange("A"+ 2).setFormula("=filter(DATA!A:BH,DATA!D:D=\""+sName+"\")"); 
          } else if ( sBlock == 2) {
            ss.getRange("A"+ 2).setFormula("=filter(DATA!A:BH,DATA!E:E=\""+sName+"\")"); 
          } else if ( sBlock == 3) {
            ss.getRange("A"+ 2).setFormula("=filter(DATA!A:BH,DATA!F:F=\""+sName+"\")"); 
          } else if ( sBlock == 4) {
            ss.getRange("A"+ 2).setFormula("=filter(DATA!A:BH,DATA!G:G=\""+sName+"\")"); 
          }
          
          ss.getRange("A1:BH1").setBackgroundRGB(11, 83, 148).setFontColor("white");// customize color
        
        } catch(e) {
          message.push("row " + (i+1));
        }
      }    
    }else{break;}
  }
}
