function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Add Student's at once (FIFO)", functionName: "addSheets"},
    {name: "Delete Sheets", functionName: "delSheets"},
    {name: "Create SpreadSheet", functionName: "createSpreadsheet"},
    {name: "Email SpreadSheet", functionName: "emailSpreadSheet"},
    {name: "Insert Overall", functionName: "insertOverall"}
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
  
  var message = [];    
//  for(var i=1, len=rData.length; i<len; i++) { // var i represent rows
  for(var i=1, len=rData.length; i<len; i++) {
    if(rData[i][0] != "" || rData[i][1] != "" || rData[i][2] != "") { 
//      ss.toast(i);
      if(arr.length== 0 || (arr.indexOf(rData[i][2])+1) == 0){
        try {
//          ss.toast(arr[i]);
          ss.insertSheet(rData[i][2]);
          arr.push(rData[i][2]); // stored the created sheet name
          emailArr.push(rData[i][1]); // stored the email
          ss.setActiveSheet(ss.getSheets()[0]);
          
          // create individual sheet
          //copy the first row (header)
          var source = sh.getRange("A1:BH1");
          ss.setActiveSheet(ss.getSheetByName(rData[i][2]));
          source.copyTo(ss.getRange("A1:BH1"));
          ss.getRange("A"+ 2).setFormula("=filter(DATA!A:BH,DATA!C:C=\""+rData[i][2]+"\")");
          ss.getRange("A1:BH1").setBackgroundRGB(11, 83, 148).setFontColor("white");// customize color
          
          //create new spreadsheet for each individual 
          var cs = SpreadsheetApp.create(rData[i][2]); // create sheet with student name as title
          var os = SpreadsheetApp.openByUrl(cs.getUrl()); // open created spreadsheet using url
          os.setActiveSheet(os.getSheets()[0]).setName("Data"); // change the first sheet name to "Data"
          os.getRange("A"+1).setFormula("=IMPORTRANGE(\""+ss.getUrl()+"\",\""+rData[i][2]+"!A:BH\")");
          os.getRange("A1:BH1").setBackgroundRGB(11, 83, 148).setFontColor("white"); // customize color
          
          urlArr.push(cs.getUrl());
          sheetID.push(cs.getId());
          
          //insert overall sheet
          os.insertSheet('OVERALL');
          os.setActiveSheet(os.getSheetByName('OVERALL'));
          os.getRange("A"+1).setFormula("=IMPORTRANGE(\""+mainUrl+"\",\"OVERALL!A:B\")");
          
//          //update url in Namelist 
//          var us = SpreadsheetApp.getActiveSpreadsheet();
//          us.setActiveSheet(us.getSheetByName('NAMELIST'));
//          var uh = us.getActiveSheet(); 
//          var uData = uh.getDataRange().getValues();
//          
//          for(var u = 0, ulen=uData.length; u<ulen; u++){
//            if(uData[u][0] == rData[i][2]){
//              us.getRange("E"+(u+1)).setValue(cs.getUrl());
//            }
//              
//          }        
                    
//          Logger.log(uData.length);

        } catch(e) {
          message.push("row " + (i+1));
        }
      }    
    }else{break;}
  }
    ss.setActiveSheet(ss.getSheets()[0]);
  //enter email address for each student
  var us = SpreadsheetApp.getActiveSpreadsheet();
  us.setActiveSheet(us.getSheetByName('NAMELIST'));
  var uh = us.getActiveSheet(); 
  var uData = uh.getDataRange().getValues();
          
  for(var u = 0, ulen=uData.length; u<ulen; u++){
    if(uData[u][0] != ""){
    var loc = arr.indexOf(uData[u][0]); //get the location
    us.getRange("D"+(u+1)).setValue(emailArr[loc]);
    us.getRange("E"+(u+1)).setValue(urlArr[loc]);
    us.getRange("F"+(u+1)).setValue(sheetID[loc]);
    Logger.log(emailArr[loc]);  
    }
  }
  
  
//  ss.toast("These sheets already exist: " + message);
//  ss.setActiveSheet(ss.getSheets()[0]);
  

}

//Del
function delSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shs = ss.getNumSheets();

  for(var i=shs-1;i>0;i--){
    ss.setActiveSheet(ss.getSheets()[i]);
    var sName = ss.getActiveSheet().getName();
    if(sName != 'NAMELIST' && sName != 'OVERALL' && sName != 'DATA'){ 
      ss.deleteActiveSheet();
      Logger.log(sName);
    }
  }
  ss.setActiveSheet(ss.getSheets()[0]);
//  ss.getRange("D2:D").clear();
}


function createSpreadsheet(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('NAMELIST'));
  var sh = ss.getActiveSheet();
  var rData = sh.getDataRange().getValues();
  var arr = []; // saving name that already use

  
  for(var i=0, len=rData.length; i<len; i++){
    var name = rData[i][0]; // get name 
    var cs = SpreadsheetApp.openByUrl(rData[i][4]); // open spreadsheet using get url
    var ch = cs.getActiveSheet();
    var cData = ch.getDataRange().getValues();
    var tArr = [];
    
    for(var b=1,blen=cData.length; b<blen; b++){
      if(cData[b][3]!=""){
        if(tArr.length == 0 || (tArr.indexOf(cData[b][3])+1) == 0){
            //create sheet for procedure
          var chk = cs.getSheetByName(cData[b][3]);
          if(!chk){
            cs.insertSheet(cData[b][3]);
            tArr.push(cData[b][3]);
            var oSource = ch.getRange("D1:BH1");
            cs.setActiveSheet(cs.getSheetByName(cData[b][3]));
            oSource.copyTo(cs.getRange("A1:BE1"));
            cs.getRange("A"+2).setFormula("=filter(Data!D:BH,Data!D:D=\""+cData[b][3]+"\")");
            cs.getRange("A1:BE1").setBackgroundRGB(11, 83, 148).setFontColor("white");
            Logger.log(b);
          }
        }
      }else{break;}
    }
    
  }
  
}

function emailSpreadSheet(){
  
  var us = SpreadsheetApp.getActiveSpreadsheet();
  us.setActiveSheet(us.getSheetByName('NAMELIST'));
  var uh = us.getActiveSheet(); 
  var uData = uh.getDataRange().getValues();
  var email = []; // student email
  var sUrl = []; // student spreadsheet url
  var sId = []; // spreadsheet id
  
  // stored sequence
  for(var u = 0, ulen=uData.length; u<ulen; u++){
    if(uData[u][0] != ""){
      email.push(uData[u][3]);
      sUrl.push(uData[u][4]);  
      sId.push(uData[u][5]);
    }
  }
  
  //send email
//  for(var e = 0, elen=email.length; e<elen; e++){
  for(var e = 0, elen=1; e<elen; e++){
//    var file = DriveApp.getFileById(sId[e]);
//    var addView =  file.addViewer(email[e]); // add viewer for the file(specific to the target-user email)
//    var addCom = file.addCommenter(email[e])
//    var url = sUrl[e];

    var emailAddress = email[e]; // student email
    var message = "Use or Press the Link below to view your logbook\n\n\n\n"+sUrl[e]; // logbook url / spreadsheet
    var subject = 'LOGBOOK';
    MailApp.sendEmail(emailAddress, subject, message);
    Logger.log(emailAddress);
    Logger.log(message);

  }

  
}

function insertOverall(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('NAMELIST'));
  var sh = ss.getActiveSheet();
  var Data = sh.getDataRange().getValues();
  var url = []; // stored spreadsheet url
  
//  get and store spreadsheet url from namelist
  for(var i=0, len=Data.length; i<len; i++){
    url.push(Data[i][4]);
  }

  
//  open each spreadsheet from url
  for(var i=0, len=1; i<len; i++){
    ss = SpreadsheetApp.openByUrl(url[i]);
    //open overall sheet
    ss.setActiveSheet(ss.getSheetByName('OVERALL'));
    var sData = ss.getActiveSheet().getDataRange().getValues();
    var procedure = [];
    var minRequirement = [];
    var calReq = [];
    var marks = [];
    var tTime = 0;
//    stored the procedure into an array
//     for(var j = 0, jlen=sData.length; j<jlen;j++){
    for(var j = 0, jlen=sData.length; j<jlen;j++){
      if(sData[j][0] != "" && sData[j][0] != undefined){
//        var tmp = sData[j][0].toUpperCase();
//        if(sData[j][0] != tmp){ // check if the data is uppercase
          procedure[j] = sData[j][0]; // stored into procedure array
          minRequirement[j] = sData[j][1]; // stored into minRequirement array
//        Logger.log(procedure[j]);
//        }
      }
    }
//    get the number of occurance from each sheet
    for(var k=0, klen=procedure.length; k<klen; k++){
      
      if(procedure[k] != undefined){
        var proc = procedure[k];
        proc = proc.split(' (');
        if(proc[0] == 'CASE PRESENTATION'){
          proc[0] = 'Case Presentation';
        }
        else if(proc[0] == 'Evacuation of Retained Products of Conception'){
          proc[0] = 'Evacuation of Retained Product of Conception (ERPOC)';
        }
        else if((proc[0].indexOf('ON CALL')+1) > 0 ){
          proc[0] = 'On Call';
          proc[1] ='';
        }
        
        var chk = ss.getSheetByName(proc[0]);
        var con = proc[1];
        var cn1 = 'performed', cn2 = 'observed', cn3 = 'performed - with partogram submitted';
        
        if(con != undefined){
          if((con.indexOf('performed - with partogram submitted')+1) > 0){
            con = 'performed - with partogram submitted';
          }else if((con.indexOf('observed')+1) > 0){
            con = 'observed';
          }else if((con.indexOf('performed')+1) > 0){
            con = 'performed';
          }
        }
//        Logger.log(proc[0]);

        if(chk){
          ss.setActiveSheet(ss.getSheetByName(proc[0]));
          var oData = ss.getActiveSheet().getDataRange().getValues(); // sheet data
          
          if(proc[0] == 'Case Presentation'){
            for( r = 1, rlen = oData.length; r< rlen; r++){
              if(oData[r][45] != undefined && oData[r][45] != ""){
                if(r%2 != 0 || marks[0] == undefined){
                  if(marks[0] == undefined){
                    marks[0] = oData[r][45];
                  }else{
                    if(oData[r][45] > marks[0]){
                      marks[0] = oData[r][45];
                    }
                  }
                }
                else if(r%2 == 0 || marks[1] == undefined){
                  if(marks[1] == undefined){
                    marks[1] = oData[r][45];
                  }else{
                    if(oData[r][45] > marks[1]){
                      marks[1] = oData[r][45];
                    }
                  }
                }
              }
              Logger.log('marks[0] = '+marks[0]);
              Logger.log('marks[1] = '+marks[1]);
            }
            calReq[k+1] = marks[0]+marks[1];
          }
          else if(proc[0] == 'On Call'){
            calReq[k+1] = '=SUM(\'On Call\'!BE2:BE)';
          }
          
          for(var e = 1, elen = oData.length; e<elen; e++){
//            Logger.log(oData.length);
              for(var x = 0; x<4; x++){
                switch(x){
                  case 0 : 
                    var z = 2;
                    break;
                  case 1: 
                    var z = 11;
                    break;
                  case 2:
                    var z = 21;
                    break;
                  case 3:
                    var z = 30;
                    break;
                }

                if(oData[e][z] != undefined){
                  if(con == 'performed'){
                    if((oData[e][z].indexOf('Performed Under Supervision')+1) > 0){
                      if(calReq[k] == undefined){
                        calReq[k] = 0;
                      }
                      calReq[k] += 1;
                      break;
                    }
                 }
                  if(con == 'observed' || con == undefined || con == ""){
                    if((oData[e][z].indexOf('Observed')+1) > 0){
                      if(calReq[k] == undefined){
                        calReq[k] = 0;
                      }
                      calReq[k] += 1;
                      break;
                    }
                    Logger.log(proc[0]);
                    Logger.log(calReq[k]);
                    Logger.log(con);
                    Logger.log(oData[e][z]);
                  }
                  if(con == 'performed - with partogram submitted'){
                    if((oData[e][z].indexOf('Performed Under Supervision')+1) > 0){
                      if(calReq[k] == undefined){
                        calReq[k] = 0;
                      }
                      calReq[k] += 1;
                      break;
                    }
                  }
                  
                }else{
                  break;
                }
              }
//            Logger.log(proc[0]);
//            Logger.log(calReq[k]);
//            Logger.log(con);
          }
          
        }
      }
    }
    
          
//      write into overall
    ss.setActiveSheet(ss.getSheetByName('OVERALL'));
    
    for(var v = 0, vlen = calReq.length; v<vlen; v++){
      if(calReq[v] != undefined){
        ss.getRange("D"+(v+1)).setValue(calReq[v]);
        
      }
//      Logger.log(calReq[v]);      
    }
    
    
    }
    
    }