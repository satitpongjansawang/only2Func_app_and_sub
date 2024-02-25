
function submitForm(e) {
    
  var ss1,ss2,ss3;
  var i,rg,rows,values;
  var rt,ml,mlsubj,sender;
  var pk;
  
  ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName(INPUT_DATA);
  rg = ss1.getRange("A:A");



  if (e.docFile.length != 0) {

      //Logger.log("Check File OK");
    
      var fileBlob = e.docFile;

      var folder = DriveApp.getFolderById(ATTACH_FOLDER);
      var tmp_folder = DriveApp.getFolderById(TEMP_FOLDER);
      var f_name = fileBlob.getName();

      var t_name = "TMP_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss") + "-" + e.txtIssueBy;
      var App_name = e.txtMaster;
      var Comp_ = e.txtCompany;

      var Master_Convert = [];
      Master_Convert = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getRange("A:H").getValues().filter(
              function (e) {
                
              if (e[0].toString() != "") {
                  return (App_name.indexOf(e[0].toString()) > -1) && (Comp_.indexOf(e[1].toString())>-1);
              } else {
                  return false;
              }
          });
      Master_Convert.push(["", ""]);
    
      //Logger.log("Get Master Data - Master_Convert[0][6].toString().toLowerCase() = " + Master_Convert[0][6].toString().toLowerCase());

      var Convert_flg = (Master_Convert[0][0] == "" ? 'true' : Master_Convert[0][6].toString().toLowerCase());
      var file_des,doc;
     

      if (Convert_flg == "true") {

          var doc_t = GASConvertExcelToSpreadsheetActive.convertExcel2Sheets(fileBlob, t_name, TEMP_FOLDER);
          var rows_ = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getDataRange().getLastRow();
          var Master_ver = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getRange(1, 1, rows_, 12).getValues();
          //return "Joe-ok"
        
          var template_sheet,
          form_Sheet;
          //Column : [Application Name]	[COMPANY]	[Sheet Name]	[Pos_Ver]	[VER.DOC]	[Check_Sign]	[Conv_Google]	[Price_Autorize]	[Template_ID]   [Sheet Form]
            //Logger.log("Check version & signature  Start : ")
          for (var k = 1; k < rows_; k++) {
              if (Master_ver[k][0] == "") {}
              else {
                  if ((Master_ver[k][0] == App_name) && (Master_ver[k][1] == Comp_) && (Master_ver[k][2] != "")) {

                      var chk_sheet = doc_t.getSheetByName(Master_ver[k][2]);
                      template_sheet = Master_ver[k][8];
                      form_Sheet = Master_ver[k][2];
                      if (chk_sheet != null) {
                             // Check company
                         if (Master_ver[k][9] !== null && Master_ver[k][9] !== '') {
                           if (chk_sheet.getRange(Master_ver[k][9]).getValue().trim()!= Comp_.trim()) {
                             //return chk_sheet.getRange(Master_ver[k][9]).getValue().trim()+ " :: " + Comp_.trim()
                             return 31;
                           }
                         } 
                         
                         if (Master_ver[k][11] == 'PR amount greater than or equal to 50,000 THB or other currency'){
                            if (Master_ver[k][10] !== null && Master_ver[k][10] !== '') {
                              if (chk_sheet.getRange(Master_ver[k][10]).getValue() !== Master_ver[k][11]) {
                                //return chk_sheet.getRange(Master_ver[k][9]).getValue().trim()+ " :: " + Comp_.trim()
                                return 311;
                              }
                            }
                         }
                         if (Master_ver[k][11] == 'PR amount less than 50,000 THB'){
                            if (Master_ver[k][10] !== null && Master_ver[k][10] !== '') {
                              if (chk_sheet.getRange(Master_ver[k][10]).getValue() !== Master_ver[k][11]) {
                                //return chk_sheet.getRange(Master_ver[k][9]).getValue().trim()+ " :: " + Comp_.trim()
                                return 312;
                              }
                            }
                         }


                        //return "CHK";
                          var ver_ = "" + chk_sheet.getRange(Master_ver[k][3]).getValue();
                          //return ver_+ ":"+Master_ver[k][4];
                          if (ver_ != "") {
                              // Compare version with master
                              if (ver_ != Master_ver[k][4]) {
                                //  ss1.deleteRow(rows);
                                  return 3;
                              } 

                              //var t_2h_s = new Date();
                              // Check sign if Master set range
                              if (((Master_ver[k][5] == null) || (Master_ver[k][5].toString().trim() == "")) == false) {
                                  // return "Error";
                                  //return Master_ver[k][5]=="" ;
                                //Logger.log(chk_sheet.getRange(Master_ver[k][2]+"!"+Master_ver[k][5]).Name)
                                  if ((chk_sheet.getRange(Master_ver[k][2]+"!"+Master_ver[k][5]).getValue() == null) || (chk_sheet.getRange(Master_ver[k][2]+"!"+Master_ver[k][5]).getValue() == "")) {
                                    //  ss1.deleteRow(rows);
                                      return 4;
                                  }
                              }
                              // var t_2h_e = new Date();
                              //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(),"2H",t_2h_s,t_2h_e,t_2h_e-t_2h_s,version_wfl,"Check File Signature"]);

                          } else {
                           // ss1.deleteRow(rows);//appendRow([rows+"_"]);
                            return 3;
                          }
                      } else {
                          //ss1.deleteRow(rows);
                          return 3;
                      }
                  }
              }
          }

          //Logger.log("Check version & signature Stop : ")
          //var t_2g_e = new Date();
         // SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2G", t_2g_s, t_2g_e, t_2g_e - t_2g_s, version_wfl, "Check File Version"]);
          //var t10 = new Date();
           //Logger.log("Check Move File temp to keep Form Folder  Start : ")
          //var t_2h_s = new Date();
          /*doc_t.getName()
          var file_src = tmp_folder.getFilesByName(doc_t.getName()).next();*/ // SOM20240122 Change from move by file name to Google ID
          var file_src = DriveApp.getFileById(doc_t.getId()); // SOM20240122

          file_des = file_src.makeCopy(folder);
          
          file_des.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
          file_src.setTrashed(true);
                  
         var t_copy=0;
           //Logger.log("Check Move File temp to keep Form Folder  Stop : ")
        doc = SpreadsheetApp.open(file_des);
        doc.setSpreadsheetTimeZone("Asia/Bangkok");
          //--------------------------------------------------------

          var sht_input_id = 0;

          var sht_itm = doc.getSheetByName("INPUT sheet");
          if (sht_itm != null) {
              sht_input_id = "" + doc.getSheetByName("INPUT sheet").getSheetId();
          }

          var doc_id = doc.getId();
          var doc_url = doc.getUrl() + "#gid=" + sht_input_id;

      } else {

          //Uncovert

          // ATTACH_FOLDER_UNCONVERT
          var folder_uncon = DriveApp.getFolderById(ATTACH_FOLDER_UNCONVERT);
          var file_form = folder_uncon.createFile(fileBlob);
          var ff = folder_uncon.getFilesByName(file_form.getName()).next()
          var folder_con = DriveApp.getFolderById(ATTACH_FOLDER);

          var file_sprtsht = SpreadsheetApp.create('Creatsheet-' + Session.getActiveUser().getEmail(), 100, 10)
              file_sprtsht.getSheetByName("Sheet1").getRange("A1").setValue("This Sheet will use Download to Stamp Route.")
          var file_src = DriveApp.getFilesByName(file_sprtsht.getName()).next();
            
          file_des = file_src.makeCopy(folder_con);
          file_des.setSpreadsheetTimeZone("Asia/Bangkok");
          file_des.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
          file_src.setTrashed(true);
          doc = SpreadsheetApp.open(file_des);
          var doc_id = file_des.getId();
          var doc_url = ff.getUrl(); //+"#gid="+sht_input_id;

      }
      
      //var t_2h_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2H", t_2h_s, t_2h_e, (t_2h_e - t_2h_s)-t_copy, version_wfl, "Move File Form After Check temp to Folder"]);

      //var t9 = new Date();
      //var t_2m_s = new Date();

      //var doc = convertExcel2Sheets(fileBlob, f_name, folder.getId());
      //Logger.log("Check Write Data History  Start : ")
      var sht = doc.insertSheet("History", doc.getSheets().length).setTabColor('GREEN'); // Change from >> "History" to "History"
      
      var txtRemark = ""
      if ((typeof(e.txtRemark)!= "undefined") && e.txtRemark != "") {
        txtRemark +="[Comment]\r\n "+e.txtRemark;
      }
      if ((typeof(e.txtRef3_)!= "undefined") && e.txtRef3_ != "") {
        txtRemark +="\r\n[Priority Reason :: Start priority from URGENT]\r\n "+ e.txtRef3_;
      }

      var col_h = ["Step", "E-Mail", "Action", "Date/Time", "Comment"];

      sht.getRange("A1").setValue("Step");
      sht.getRange("B1").setValue("E-Mail");
      sht.getRange("C1").setValue("Action");
      sht.getRange("D1").setValue("Date/Time");
      sht.getRange("E1").setValue("Comment");

      sht.getRange("A2").setValue("1");
      sht.getRange("B2").setValue(e.txtIssueBy);
      sht.getRange("C2").setValue("Issued");
      sht.getRange("D2").setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      sht.getRange("E2").setValue(txtRemark);
 
      sht.autoResizeColumn(1);
      sht.autoResizeColumn(2);
      sht.autoResizeColumn(3);
      sht.autoResizeColumn(4);
      sht.autoResizeColumn(5);


      var file_set = DriveApp.getFileById(doc.getId());
      var file_id_doc =file_set.getId();
      file_set.setOwner("admin@ngkntk-asia.com")
      file_set.setShareableByEditors(true);
      //file_set.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT) // Original
      file_set.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT)   // New From 14-07-2020 Change from domain to private
      
      //file_set.addEditor("s-tatcha@ngkntk-asia.com")                           // New From 14-07-2020 Add permission for person in route only
    //RV1 
    var output = [];
    //RV1
    var optionalArgs = {    sendNotificationEmails: false};  
    e.reviewLevel1.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.reviewLevel2.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      //return "OK-"+output
    e.reviewLevel3.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.reviewLevel4.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.reviewLevel5.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.reviewLevel6.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.reviewLevel7.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.finalApprove.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
    e.applicationAdmin.split(',').forEach(function(q){if (q.toString()!= "")   {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
   // return "OK-"+JSON.stringify(e.applicationAdminCC.split(','));
    e.applicationAdminCC.split(',').forEach(function(q){if (q.toString()!= "") {try{
      Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false}) }catch (e) {output.push(q + "-"+ e);};}
                                                                              ;})

      
      
      //file_set.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT)
    // return "Joe-ok1"
      //-------------------------------------
      //var t7 = new Date();
      var folder2 = folder.createFolder(doc_id);

     //folder2.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDI)
      var user1 = [];
      user1.push("admin@ngkntk-asia.com");

  for(var a in e){
    if (a.indexOf("attFile")>-1 && e[a].length !=0) {
          var zipdoc = folder2.createFile(e[a]);

     }
  }

  } else {
      var doc_id = "";
      var doc_url = "";
  }
     //Logger.log("Upload attach File  Stop : ")
     
     //Logger.log("Check Write Data in Database Sheet Start: ")
     
  var txtId = e.txtId;
  var txtDept = e.txtDepartment;
  //return txtDept;
  var txtDeptCode = txtDept.split("|")[2];
  var txtDepartment = txtDept.split("|");
  var txtRef1 = e.txtRef1;
  var txtRef2 = e.txtRef2;
  var txtRef3 = e.txtRef3;
  var txtIssueDate = e.txtIssueDate;
  var txtIssueBy = e.txtIssueBy;
  if (txtIssueBy != "") {
      user1.push(txtIssueBy);
  }
  var chkCompany = e.chkCompany;
  var chkMaster = e.txtMaster; //e.chkMaster;
 
  var txtPriority = e.priority;
  // change folder name ----------------
  var f_code_name =  Comp_ + "_" + txtDeptCode + "_" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmssSSS") + "_" + chkMaster
  //return f_code_name
  //t1 = new Date();
   
  var UniqueStr = getUniqueStr();
  
  //ss1.appendRow(["=Row()&\"_\"&\""+f_code_name+"\""]);
   ss1.appendRow([UniqueStr +"_"+f_code_name]); //test 220818
var namepaths = UniqueStr +"_"+f_code_name;
  //t2 = new Date();
  var qvizQuery = "SELECT * WHERE A like '%" + f_code_name + "%'";
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ss1.getParent().getId()
  + '/gviz/tq?tqx=out:json&headers=1&sheet=' + ss1.getName() + '&range=A:A&tq=' + encodeURIComponent(qvizQuery);
  var qvizret = UrlFetchApp.fetch(qvizURL, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getContentText();
  var qvizjson = JSON.parse(qvizret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2))
   //return qvizjson
  if ((qvizjson.table.rows[0].c[0].v != "undefined") || (qvizjson.table.rows[0].c[0].v !="")) {
    //f_code_name = qvizjson.table.rows[0].c[0].v
      f_code_name =  UniqueStr; //test 220818
  } else {
    return 9;
  }
  //f_code_name =  UniqueStr; //test 220818
 // t3 = new Date();
  //Logger.log("Write Data : " + (t2 - t1) )
 // Logger.log("Read ID Data : " + (t3 - t2) )
  //  Logger.log(f_code_name)
  //ss1.appendRow([f_code_name]);
  //rows = rg.getLastRow() + 1;
   rows =  f_code_name.split("_")[0]
  //var f_code_name = rows + "_" +f_code_name
   
   
  doc.setName(namepaths);
  file_des.setName(namepaths)
  folder2.setName(namepaths);
  if (Convert_flg != "true") {
      file_form.setName(namepaths);
  }

  //var t_2j_e = new Date();
  //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2J", t_2j_s, t_2j_e, t_2j_e - t_2j_s, version_wfl, "Upload File Support  & Unzip"]);

  //var t_2k_s = new Date();

  //ADDED FOR V2
  var chkReview1 = e.reviewLevel1;
  var STReview1 = e.STreviewLevel1;
  var RMReview1 = (STReview1 == "1" ? "" : "Remove-" + e.CommentreviewLevel1);
  var PTHReview1 = (STReview1 == "1" ? "" : doc_url);
  var IDReview1 = (STReview1 == "1" ? "" : doc_id);
  var PCReview1 = (STReview1 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkReview1    != "") {user1.push(chkReview1);}

  var chkReview2 = e.reviewLevel2;
  var STReview2 = e.STreviewLevel2;
  var RMReview2 = (STReview2 == "1" ? "" : "Remove-" + e.CommentreviewLevel2);
  var PTHReview2 = (STReview2 == "1" ? "" : doc_url);
  var IDReview2 = (STReview2 == "1" ? "" : doc_id);
  var PCReview2 = (STReview2 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkReview2    != "") {user1.push(chkReview2);}

  var chkReview3 = e.reviewLevel3;
  var STReview3 = e.STreviewLevel3;
  var RMReview3 = (STReview3 == "1" ? "" : "Remove-" + e.CommentreviewLevel3);
  var PTHReview3 = (STReview3 == "1" ? "" : doc_url);
  var IDReview3 = (STReview3 == "1" ? "" : doc_id);
  var PCReview3 = (STReview3 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkReview3    != "") {user1.push(chkReview3);}

  var chkReview4 = e.reviewLevel4;
  var STReview4 = e.STreviewLevel4;
  var RMReview4 = (STReview4 == "1" ? "" : "Remove-" + e.CommentreviewLevel4);
  var PTHReview4 = (STReview4 == "1" ? "" : doc_url);
  var IDReview4 = (STReview4 == "1" ? "" : doc_id);
  var PCReview4 = (STReview4 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkReview4    != "") {user1.push(chkReview4);}

  var chkReview5 = e.reviewLevel5;
  var STReview5 = e.STreviewLevel5;
  var RMReview5 = (STReview5 == "1" ? "" : "Remove-" + e.CommentreviewLevel5);
  var PTHReview5 = (STReview5 == "1" ? "" : doc_url);
  var IDReview5 = (STReview5 == "1" ? "" : doc_id);
  var PCReview5 = (STReview5 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkApprove1   != "") {user1.push(chkApprove1);}

  var chkReview6 = e.reviewLevel6;
  var STReview6 = e.STreviewLevel6;
  var RMReview6 = (STReview6 == "1" ? "" : "Remove-" + e.CommentreviewLevel6);
  var PTHReview6 = (STReview6 == "1" ? "" : doc_url);
  var IDReview6 = (STReview6 == "1" ? "" : doc_id);
  var PCReview6 = (STReview6 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkReview6    != "") {user1.push(chkReview6);}

  var chkReview7 = e.reviewLevel7;
  var STReview7 = e.STreviewLevel7;
  var RMReview7 = (STReview7 == "1" ? "" : "Remove-" + e.CommentreviewLevel7);
  var PTHReview7 = (STReview7 == "1" ? "" : doc_url);
  var IDReview7 = (STReview7 == "1" ? "" : doc_id);
  var PCReview7 = (STReview7 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
  //if (chkReview7    != "") {user1.push(chkReview7);}

  var chkApprove2 = e.finalApprove;
  //if (chkApprove2   != "") {user1.push(chkApprove2);}
  var txtRegister = e.applicationAdmin;
  //if (txtRegister   != "") {user1.push(txtRegister);}
  var txtRegisterCC = e.applicationAdminCC;
  //if (txtRegisterCC != "") {user1.push(txtRegisterCC);}

  //return   STReview1+":"+STReview2+":"+STReview3+":"+STReview4+":"+STReview5+":"+STReview6+":"+STReview7+":"


  var bi;
  //END ADDED FOR V2

  //////SUBMIT REQUEST STEP
  //var t6 = new Date();
  /*
  pk = txtRef1 + ":" + chkCompany + '_' + chkMaster + '_' + Utilities.formatDate(new Date(), "GMT+7","yyMMddHHmmss");

  if (txtRef1 == "-") {
  pk = chkCompany + '_' + txtDepartment[1] + '_' + chkMaster + '_' + Utilities.formatDate(new Date(), "GMT+7","yyMMddHHmmss");
  } else {
  pk = txtRef1 + ":" + chkCompany + '_' + txtDepartment[1] + '_' + chkMaster + '_' + Utilities.formatDate(new Date(), "GMT+7","yyMMddHHmmss");
  }
   */
  //ss2 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);

  //pk = f_code_name; 
pk = namepaths;
/*
ss1.appendRow([
          　　　 pk,
          chkCompany,
          'Issue',
          chkMaster,
          txtRemark,
          txtIssueBy,
          txtIssueDate,
          doc_url,
          doc_id,
          Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")
          //ADDED FOR V2
          //[comment by reviewer],]reviewer email],[Attach Path],[Attcch ID],[process date & time]
      , RMReview5, chkReview5, PTHReview5, IDReview5, PCReview5, '', chkApprove2, '', '', '', '',
    txtRegister, '', '', '', '', RMReview1, chkReview1, PTHReview1, IDReview1, PCReview1, RMReview2,
    chkReview2, PTHReview2, IDReview2, PCReview2, RMReview3, chkReview3, PTHReview3, IDReview3, PCReview3,
    RMReview4, chkReview4, PTHReview4, IDReview4, PCReview4, txtRegisterCC, RMReview6, chkReview6, PTHReview6,
    IDReview6, PCReview6, RMReview7, chkReview7, PTHReview7, IDReview7, PCReview7, txtRef1, txtRef2, txtRef3
          //ADDED FOR RELEASE 02.052018

      ]);
      */

//return  "A"+pk+":BH"+pk;
 var data_input=[[
          　　　 pk,
          chkCompany,
          'Issue',
          chkMaster,
          txtRemark,
          txtIssueBy,
          txtIssueDate,
          doc_url,
          doc_id,
          Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")
          //ADDED FOR V2
          //[comment by reviewer],]reviewer email],[Attach Path],[Attcch ID],[process date & time]
      , RMReview5, chkReview5, PTHReview5, IDReview5, PCReview5, '', chkApprove2, '', '', '', '',
    txtRegister, '', '', '', '', RMReview1, chkReview1, PTHReview1, IDReview1, PCReview1, RMReview2,
    chkReview2, PTHReview2, IDReview2, PCReview2, RMReview3, chkReview3, PTHReview3, IDReview3, PCReview3,
    RMReview4, chkReview4, PTHReview4, IDReview4, PCReview4, txtRegisterCC, RMReview6, chkReview6, PTHReview6,
    IDReview6, PCReview6, RMReview7, chkReview7, PTHReview7, IDReview7, PCReview7, txtRef1, txtRef2, txtRef3,
    '', '', '', '','', '', '', '','', '', '', '','', '', '', '','', '', '', '','', '', '', '','', '', '', '',txtPriority
          //ADDED FOR RELEASE 02.052018

      ]];
var temprow = findRow(rows); // test 202218

rows = temprow; // test 202218
ss1.getRange("A" +  rows +":CK"+rows).setValues(data_input)
//ss1.getRange("A"+pk+":BH"+pk).setValues()

/// INPUT CULUMN CH
var min = 1;
var max = 100;
const now =  Utilities.formatDate(new Date(), "GMT+7:00", "yyyyMMddHHmmss") + Math.random() * (max - min) + min;
var result = parseInt(now).toString();
////
  //var t5 = new Date();
var formulas = [
  ['=IF(IFERROR(FIND(\"Rejected\",C' + rows + '), -1) >= 0, \"\", CONCATENATE(CH' + rows + ',"|",BU' + rows + ',"|",CF' + rows + ',"|",CG' + rows + '))',
  '',
  '=IF(AE' + rows + '="",0,1)',
  '=IF(AJ' + rows + '="",0,1)',
  '=IF(AO' + rows + '="",0,1)',
  '=IF(AT' + rows + '="",0,1)',
  '=IF(AZ' + rows + '="",0,1)',
  '=IF(BE' + rows + '="",0,1)',
  '=IF(O' + rows + '="",0,1)',
  '=IF(T' + rows + '="",0,1)',
  '=IF(Y' + rows + '="",0,1)',
  '=SUMPRODUCT($BK$1:$BS$1,BK' + rows + ':BS' + rows + ')',
  '=IF(BT' + rows + '<1,"0",IF(BT' + rows + '<2,"1",IF(BT' + rows + '<4,"1A",IF(BT' + rows + '<8,"1B",IF(BT' + rows + '<16,"1C",IF(BT' + rows + '<32,"1D",IF(BT' + rows + '<64,"2",IF(BT' + rows + '<128,"2A",IF(BT' + rows + '<256,"2B",IF(BT' + rows + '<512,"3","9"))))))))))',
  '=IF(AND(AB' + rows + '<>"",AE' + rows + '=""),1,0)',
  '=IF(AND(AG' + rows + '<>"",AJ' + rows + '=""),1,0)',
  '=IF(AND(AL' + rows + '<>"",AO' + rows + '=""),1,0)',
  '=IF(AND(AQ' + rows + '<>"",AT' + rows + '=""),1,0)',
  '=IF(AND(AW' + rows + '<>"",AZ' + rows + '=""),1,0)',
  '=IF(AND(BB' + rows + '<>"",BE' + rows + '=""),1,0)',
  '=IF(AND(L' + rows + '<>"",O' + rows + '=""),1,0)',
  '=IF(AND(Q' + rows + '<>"",T' + rows + '=""),1,0)',
  '=IF(AND(V' + rows + '<>"",Y' + rows + '=""),1,0)',
  '=SUMPRODUCT($BV$1:$CD$1,BV' + rows + ':CD' + rows + ')',
  '=IF(CE' + rows + '>=512,"1A",IF(CE' + rows + '>=256,"1B",IF(CE' + rows + '>=128,"1C",IF(CE' + rows + '>=64,"1D",IF(CE' + rows + '>=32,"2",IF(CE' + rows + '>=16,"2A",IF(CE' + rows + '>=8,"2B",IF(CE' + rows + '>=4,"3",IF(CE' + rows + '>=2,"4","9")))))))))',
  '=IF(CF' + rows + '="1A",AB' + rows + ',IF(CF' + rows + '="1B",AG' + rows + ',IF(CF' + rows + '="1C",AL' + rows + ',IF(CF' + rows + '="1D",AQ' + rows + ',IF(CF' + rows + '="2",L' + rows + ',IF(CF' + rows + '="2A",AW' + rows
       + ',IF(CF' + rows + '="2B",BB' + rows + ',IF(CF' + rows + '="3",Q' + rows + ',IF(CF' + rows + '="4",V' + rows + ',"")))))))))',
  result]
];
  ss1.getRange("BI" + rows+":CH"+rows).setFormulas(formulas)

  //var t_2k_e = new Date();
  //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2K", t_2k_s, t_2k_e, t_2k_e - t_2k_s, version_wfl, "Write to Master Sheet After Upload"]);
  //   //Logger.log("Check Write Data in Database Sheet Stop: ")
  //          //Logger.log("Check Send Mail Start: ")
  //ADDED FOR V2
  //var t_2l_s = new Date();
  
  //Check Urgent
  urgt = ""
  if (txtPriority==0) {urgt="[!!!URGENT!!!]"}
  
  mlsubj = "WFL <" + urgt + "You're Reviewer>"
      if (chkReview1 != "" && STReview1 == "1") {
          ml = chkReview1; //NEXT APPROVAL is Review #1
          rt = "1A";
      } else if (chkReview2 != "" && STReview2 == "1") {
          ml = chkReview2; //NEXT APPROVAL is Review #2
          rt = "1B";
      } else if (chkReview3 != "" && STReview3 == "1") {
          ml = chkReview3; //NEXT APPROVAL is Review #3
          rt = "1C";
      } else if (chkReview4 != "" && STReview4 == "1") {
          ml = chkReview4; //NEXT APPROVAL is Review #4
          rt = "1D";
      } else if (chkReview5 != "" && STReview5 == "1") {
          ml = chkReview5; //NEXT APPROVAL is Review #5
          rt = "2";
      } else if (chkReview6 != "" && STReview6 == "1") {
          ml = chkReview6; //NEXT APPROVAL is Review #5
          rt = "2A";
      } else if (chkReview7 != "" && STReview7 == "1") {
          ml = chkReview7; //NEXT APPROVAL is Review #5
          rt = "2B";
      }

      else if (chkApprove2 != "") {
          ml = chkApprove2; //NEXT APPROVAL is Approver.
          rt = "3";
          mlsubj = "WFL <" + urgt + "You're Approver>"
      }


     // var t4 = new Date();
  //START WRITE LOG

  ss3 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  //ss3.getSheetByName("Logs").activate();
  ss3.getSheetByName("Logs").appendRow([
          pk,
          chkCompany,
          'Issue',
          chkMaster,
          txtRemark,
          txtIssueBy,
          txtIssueDate,
          doc_url,
          doc_id,
          Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"), WEB_PATH + "?id=" + UniqueStr + "&rt=" + rt
          //ADDED FOR RELEASE 02.052018
      , txtRef1, txtRef2, txtRef3
          //ADDED FOR RELEASE 02.052018
      ]);

  //END FOR WRITE LOG

//var t3 = new Date();

  //SEND EMAIL TO NEXT STEP (REVIEW)
  ml = "admin@ngkntk-asia.com,"+ml;
  MailApp.sendEmail(ml.replace(/\s/g, ""), "" + mlsubj + " - " + txtRef1 + ':' + pk, "You have received Application from WorkFlow Launcher.\r\n"
       + "Please confirm the following link.\r\n\r\n＜Link＞\r\n"
       + WEB_PATH + "?id=" + (UniqueStr) + "&rt=" + rt + "&fr=" + 0);

  //SEND EMAIL TO ISSUE PERSON (ORIGINATOR)

  MailApp.sendEmail(txtIssueBy, "" + "WFL <"+ urgt + txtRef1 + ":" + pk + ">", "You have received this email because you're submitted WorkFlow Launcher (" + chkMaster + ").\r\n"
       + "You can click on the link to view latest document status.\r\n\r\n＜Link＞\r\n"
       //+ WEB_PATH + "?id=" + (rows) + "&rt=" + 0 + "&fr=" + 0);
       + WEB_PATH + "?id=" + (UniqueStr) + "&rt=" + "x" + "&fr=" + 0);
  //var t_2l_e = new Date();
  //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2L", t_2l_s, t_2l_e, t_2l_e - t_2l_s, version_wfl, "Send Mail to Next Process"]);
  //          //Logger.log("Check Send Mail Stop: ")
  //var t2 = new Date();
/*  
var diff = Math.floor((t2 - t1) / (1000.0)); //seconds
  var diff_mail = Math.floor((t2 - t3) / (1000.0)); //seconds
  var diff_Log = Math.floor((t3 - t4) / (1000.0)); //seconds
  var diff_formular = Math.floor((t4 - t5) / (1000.0)); //seconds
  var diff_wrt = Math.floor((t5 - t6) / (1000.0)); //seconds
  var diff_upform = Math.floor((t9 - t1) / (1000.0)); //seconds
*/
//Logger.log("Stop : ")
  return 1;

}