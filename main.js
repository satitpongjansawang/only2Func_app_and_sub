function approveForm(e) {
  "use strict";

  var returnVal;

  //var t1 = new Date();
  var ss1,ss2,e_sign;
  var i,rg,rows,values,rw;
  var rt,ml,msg;
  var act,mlcc,mlsubj,mlregistcc;
  var pk,next,nextProcApprove;
  var docPID,docPIDPk;

  var txtId = e.txtId;
  var txtRt = e.txtRt; //ORIGINAL
  var txtCompany = e.txtCompany;
  var txtMaster = e.txtMaster;
  var txtComment = e.txtComment
  if (typeof(e.txtComment) == "undefined") { txtComment = ""; }
  var txtJudge = e.txtJudge;
  var txtTempFileId = e.txtTempFileId;
  var txtPriority = e.priority;
  var txtUrl = "";
  var txtIssueBy;
  var urgent_reason = e.txtRef3_ ;
  if (typeof(e.txtRef3_) == "undefined") { urgent_reason = ""; }
  //return e.txtJudge;
  //-----------------------New
  //    return DATA_SPREADSHEET;
    ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName(INPUT_DATA);
    var doc_id= ss1.getRange("I" + txtId).getValue();
    var doc_url = ss1.getRange("H" + txtId).getValue();
    var priority = ss1.getRange("CK" + txtId).getValue();

  //------------------------
 
//Logger.log('e.txtId : ' +  e.txtId);
//Logger.log('docPID : ' + docPID);
//Logger.log('e.txtRt : ' + e.txtRt);
//Logger.log('UniqueStr : ' + UniqueStr);
var UniqueStr = GetRow2UniID(e.txtId);
  // for check stamp sheet
  var chk_sht_his = SpreadsheetApp.openById(doc_id).getSheetByName("History");
  if (chk_sht_his == null) { return 5; }

  // for check template sheet
  var search_row,check_template_sheet,file_for_check_esig;
//return txtTempFileId
 /*
  if (txtTempFileId!="") {
      file_for_check_esig = SpreadsheetApp.openById(txtTempFileId);
  } else {
      file_for_check_esig = SpreadsheetApp.openById(doc_id);
  }*/
  file_for_check_esig = SpreadsheetApp.openById(doc_id);
  //check_template_sheet = file_for_check_esig.getSheetByName("Template");
  // if (check_template_sheet != null) {
  var sheet =file_for_check_esig.getSheets()[0];
  var master_form = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getDataRange().getValues().filter(function (dataRow) {return dataRow[0] === e.txtMaster && dataRow[1] ===e.txtCompany;});
  //Logger.log(master_form)
  var master_sheet_name = master_form[0][2];
  var master_sheet_id = master_form[0][8];

  //return master_sheet_name;
//            if (sheet.getName()=="History") {
//                sheet = file_for_check_esig.getSheets()[1];
//            }


///////////////////////////////////////////////////
///////////////////////MODIFIED BY WP 20210916, IN CASE SIGNATURE NOT FOUND
   sheet = file_for_check_esig.getSheetByName(master_sheet_name);
   var search_string = Session.getActiveUser().getEmail();

   if (search_string != "admin@ngkntk-asia.com") {
     var textFinder = sheet.createTextFinder(search_string);
     search_row = textFinder.findNext();
     
     if (search_row != null) {
       // Clear signature if reject.
       // if (txtJudge=='r'){ 
       if ((txtJudge=='r')||(txtJudge=='c')) {  //ADDED BY WP 20210916 TO REMOVE THE SIGNATURE IF CANCEL ACTION
         do {
           search_row.clearContent();
           ////Logger.log(""+search_row.getA1Notation()+":"+search_row.getValue());
           search_row = textFinder.findNext(); 
         } while (search_row!= null); 
         ////Logger.log("'" + sheet.getName() + "'!" + search_row.getA1Notation());
       }
     } else {
       if(txtJudge=='a'){ return 6;}
     }
   }
     // }
      //return doc_id

/////////////////////////////////////////////////////////
////////////////////////////////////////////////////////


  // Start to Upload...
  var folder = DriveApp.getFolderById(e.txtFolderattId);
  for(var a in e){
    if (a.indexOf("attFile")>-1 && e[a].length !=0) {
      var zipdoc = folder.createFile(e[a]);
      
    }
  }
  var pRt = e.txtRt;
        
     /*     
      if (e.IattFile.length != 0) {
         // var t_3i_s = new Date();
          var fileBlob2 = e.IattFile;
          var fileBlobType = fileBlob2.getContentType();
          var folder = DriveApp.getFolderById(e.txtFolderattId);
          var zipdoc = folder.createFile(fileBlob2);
        
        
        /* UnzIp
          if (fileBlobType.indexOf("zip") == -1) {
              var zipdoc = folder.createFile(fileBlob2);
          } else {
              fileBlob2.setContentType("application/zip");
              var unZippedfile = Utilities.unzip(fileBlob2);
             // for each(var file_ in unZippedfile) {
                  var zipdoc = folder.createFile(file_);
                  //zipdoc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
             // }
          }
          
         // var t_3i_e = new Date();
          var pRt = e.txtRt;
        
          if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
              SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
              SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
          }
          if ((pRt == "3")) {
              SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
              SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
          }
          if ((pRt == "4")) {
              SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
              SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
          }

      }
*/

      //return "OK1"
      if (typeof(e.IesigFile) != "undefined") {
          if (e.IesigFile.length != 0) {
            //  var t_3i_s = new Date();
              var fileBlob2 = e.IesigFile;
              var fileBlobType = fileBlob2.getContentType();
              var folder = DriveApp.getFolderById(ATTACH_FOLDER_UNCONVERT);
              var zipdoc = folder.createFile(fileBlob2);
              //return "OK"
              var re_name = zipdoc.getName().replace(".xlsm", "[" + txtRt + "]_.xlsm");

              zipdoc.setName(re_name);
              doc_url = zipdoc.getUrl();
            //  var t_3i_e = new Date();
              var pRt = e.txtRt;
            /*
              if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
                  SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                  SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
              }
              if ((pRt == "3")) {
                  SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                  SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
              }
              if ((pRt == "4")) {
                  SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                  SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
              }
*/
          }
      }

      //ss1.getSheetByName(INPUT_DATA).activate();

  docPID = txtId; //PID CONTROLLED BY SYSTEM
  docPIDPk = docPID + 1;
  //txtId = Number(txtId) - 1; //PIC PLUS ONE FOR HEADER FOR DATABASE ACCESSS

  ////Logger.log(txtJudge + " txtCompany:" + txtCompany + " txtMaster:" + txtMaster + " txtRt:" + txtRt);
  var text_ref1;
  pk = ss1.getRange("A" + docPID).getValue(); // get Primary Key
  text_ref1 = ss1.getRange("BF" + docPID).getValue(); // get Primary Key
  if (text_ref1.length > 120) {text_ref1=text_ref1.substr(0, 120)+"...";}
  ////Logger.log(pk)
  //return ;
  //var t6 = new Date();
  //var t_3j_s = new Date();
  //var t_3m_s = new Date();
  mlcc = getCCMail(txtId); // send id to get CC Mail
  //return mlcc
  ////Logger.log(mlcc);
  //var t_3m_e = new Date();
  //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3M", t_3m_s, t_3m_e, t_3m_e - t_3m_s, version_wfl, "getCCMail"]);
  //var t_3n_s = new Date();
  //var t5 = new Date();
  mlregistcc = getRegistCCMail(txtId); // send id to get CC Mail
  //return mlregistcc;
  ////Logger.log(mlregistcc);
  //var t_3n_e = new Date();
 // SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3N", t_3n_s, t_3n_e, t_3n_e - t_3n_s, version_wfl, "getRegistCCMail"]);

  //var t4 = new Date();
  var end_rt=0;
  // Set Urgent Flag
  ss1.getRange("CK" + txtId).setValue(txtPriority);
  urgent_mail = ""
  if  (txtPriority==0) {
    urgent_mail = "[!!!URGENT!!!]"
  } else {
    urgent_mail = ""
  }
//   if (typeof(e.txtComment) == "undefined") { txtComment = ""; }
//   if (typeof(e.txtComment) == "undefined") { txtComment = ""; }
 if (txtComment!="") {
                       txtComment = "[Comment]\r\n "+txtComment; 
 }
 if (txtPriority!=priority) { 
   if (txtPriority == 0) { 

              if (urgent_reason!="") {
                       txtComment += "\r\n[Priority Reason :: Change priority from NORMAL to URGENT ]\r\n "+urgent_reason;
              }
   }
   if (txtPriority == 1) { 
             //txtComment = "[Comment]\r\n "+txtComment+"\r\n[Priority Reason :: Change priority from URGENT to NORMAL ]\r\n " 
              if (urgent_reason!="") {
                       txtComment += "\r\n[Priority Reason :: Change priority from URGENT to NORMAL ]\r\n "+urgent_reason;
              }
   }
   } else {
   if (txtPriority == 0) { 
             //txtComment = "[Comment]\r\n "+txtComment+"\r\n[Priority Reason :: ]\r\n " +urgent_reason 
              if (urgent_reason!="") {
                       txtComment += "\r\n[Priority Reason :: ]\r\n "+urgent_reason;
              }
   }
   if (txtPriority == 1) { 
             //txtComment = "[Comment]\r\n "+txtComment+"\r\n[Priority Reason :: ]\r\n " 
              if (urgent_reason!="") {
                       txtComment += "\r\n[Priority Reason :: ]\r\n "+urgent_reason;
              }
   }
   }
  
  if (txtRt == "1A") {
      //var t_3j_s = new Date();
      // Approve Review
      ss1.getRange("AA" + txtId).setValue(txtComment);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("AE" + txtId).setValue(doc_url);
      ss1.getRange("AD" + txtId).setValue(doc_id);
      ss1.getRange("AE" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason)
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      /*-------------ADDED BY P-WINYOU FOR ENTERTAIN EXPENSE AND OVERSEA BUSINESS TRIP
      var input = SpreadsheetApp.openById(doc_id).getSheetByName("Ent.pre-approval");
      //if (input != null) {

      e_sign.getRange("U31").setValue("Signature Here");
      e_sign.getRange("U32").setValue("Date/Time Here");

      //}
      //-------------END OF ADDED BY P-WINYOU 20190723*/

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date-Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "1A");
      //return docPID+":"+next;
      nextProcApprove = next.split("|");
    
       txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1A&rt=" + nextProcApprove[0];
      //txtUrl = WEB_PATH + "?id=" + docPID + "&fr=1A&rt=" + nextProcApprove[0];
     
    
      //var t_3f_e = new Date();
     // SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer1"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          msg = "Your master registration request has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer1');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer1"); //Action
      }

      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }

  if (txtRt == "1B") {
      //var t_3j_s = new Date();
      ss1.getRange("AF" + txtId).setValue(txtComment);
      //ss1.getRange("AL" + txtId).setValue(values[txtId][3]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("AH" + txtId).setValue(doc_url);
      ss1.getRange("AI" + txtId).setValue(doc_id);
      ss1.getRange("AJ" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason)
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      //e_sign.getRange("C"+rw).setValue("Reviewer2");//Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();

      next = getNextApproveStep(docPID, "1B");

      nextProcApprove = next.split("|");
      //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=1B&rt=" + nextProcApprove[0];
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1B&rt=" + nextProcApprove[0];
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer2"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
            // txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
              txtUrl = WEB_PATH + "?id=" +  UniqueStr + "&rt=0&view=1";
            
          msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer2');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer2"); //Action
      }
      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);
  }

  if (txtRt == "1C") {
      //var t_3j_s = new Date();
      ss1.getRange("AK" + txtId).setValue(txtComment);
      //ss1.getRange("AL" + txtId).setValue(values[txtId][3]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("AM" + txtId).setValue(doc_url);
      ss1.getRange("AN" + txtId).setValue(doc_id);
      ss1.getRange("AO" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason)
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      //e_sign.getRange("C"+rw).setValue("Reviewer3");//Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "1C");
      nextProcApprove = next.split("|");
     // txtUrl = WEB_PATH + "?id=" + txtId + "&fr=1C&rt=" + nextProcApprove[0];
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1C&rt=" + nextProcApprove[0];
      
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer3"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
          msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer3');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer3"); //Action
      }

      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }

  if (txtRt == "1D") {
      //var t_3j_s = new Date();
      ss1.getRange("AP" + txtId).setValue(txtComment);
      //ss1.getRange("AQ" + txtId).setValue(values[txtId][3]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("AR" + txtId).setValue(doc_url);
      ss1.getRange("AS" + txtId).setValue(doc_id);
      ss1.getRange("AT" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason);
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      //e_sign.getRange("C"+rw).setValue("Reviewer4");//Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "1D");
      nextProcApprove = next.split("|");
      //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=1D&rt=" + nextProcApprove[0];
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1D&rt=" + nextProcApprove[0];
      
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer4"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
        //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
        
          msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer4');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer4"); //Action
      }

      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }

  if (txtRt == "2") {
      //var t_3j_s = new Date();
      ss1.getRange("K" + txtId).setValue(txtComment);
      //ss1.getRange("L" + txtId).setValue(values[txtId][3]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("M" + txtId).setValue(doc_url);
      ss1.getRange("N" + txtId).setValue(doc_id);
      ss1.getRange("O" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason);
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      e_sign.getRange("C" + rw).setValue("Reviewer5"); //Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "2");
      nextProcApprove = next.split("|");
     // txtUrl = WEB_PATH + "?id=" + txtId + "&fr=2&rt=" + nextProcApprove[0];
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=2&rt=" + nextProcApprove[0];
      
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer5"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
        //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
        
          msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer5');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer5"); //Action
      }
      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're Final  " + text_ref1 + ":" + mlsubj + ">  " + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>" + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }

  if (txtRt == "2A") {
      //var t_3j_s = new Date();
      ss1.getRange("AV" + txtId).setValue(txtComment);
      //ss1.getRange("AW" + txtId).setValue(values[txtId][3]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("AX" + txtId).setValue(doc_url);
      ss1.getRange("AY" + txtId).setValue(doc_id);
      ss1.getRange("AZ" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason);
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      //e_sign.getRange("C"+rw).setValue("Reviewer6");//Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "2A");
      nextProcApprove = next.split("|");
      //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=2A&rt=" + nextProcApprove[0];
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=2A&rt=" + nextProcApprove[0];
      
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer6"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
       // txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
        
          msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer6');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer6"); //Action
      }

      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }

  if (txtRt == "2B") {
      //var t_3j_s = new Date();
      ss1.getRange("BA" + txtId).setValue(txtComment);
      //ss1.getRange("BB" + txtId).setValue(values[txtId][3]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("BC" + txtId).setValue(doc_url);
      ss1.getRange("BD" + txtId).setValue(doc_id);
      ss1.getRange("BE" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason);
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      e_sign.getRange("C" + rw).setValue("Reviewer7"); //Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "2B");
      nextProcApprove = next.split("|");
      //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=2B&rt=" + nextProcApprove[0];
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=2B&rt=" + nextProcApprove[0];
      
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }

      if (txtJudge == "a") {
          ml = nextProcApprove[1];
          msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Signed By Reviewer";
          e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer7"); //Action
      } else {
          ml = ss1.getRange("F" + txtId).getValue();
        //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
        
          msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
          ss1.getRange("C" + txtId).setValue('Rejected by Reviewer7');
          act = "Reviewer Rejected";
          e_sign.getRange("C" + rw).setValue("Rejected By Reviewer7"); //Action
      }

      //var t_3l_s = new Date();
      if (txtJudge == "a") {
          MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
      } else {
          MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
              cc: mlcc
          });
      }
      //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }

/////////////////////////////////////////////////////////////
////////////////////////START OF APPROVAL ///////////////////
/////////////////////////////////////////////////////////////

  if (txtRt == "3") {
    
      //////////////////////////////////////////SET INITIAL VALUES
      //var t_3j_s = new Date();
      ss1.getRange("P" + txtId).setValue(txtComment);
      //ss1.getRange("Q" + txtId).setValue(values[txtId][5]);
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("R" + txtId).setValue(doc_url);
      ss1.getRange("S" + txtId).setValue(doc_id);
      ss1.getRange("T" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      ss1.getRange("BH" + txtId).setValue(urgent_reason);
      //------------------------------------------------
    
      ///////////////////////////////////UPDATE DATA TO THE HISTORY SHEET IN DATASHEET ITSELF
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      //e_sign.getRange("C"+rw).setValue("Approver");//Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      next = getNextApproveStep(docPID, "3");
    
      ///////////////////////////////////////GENTERATE URL FOR EMAIL SENDING      
      // return next.toString();
      if (next.split("|")[0]!="undefined") {
        //has register
        nextProcApprove = next.split("|");
        //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=3&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=3&rt=" + nextProcApprove[0];
        
      } else {
        nextProcApprove=["",""];
        //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=3&rt=9";
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=3&rt=9";
        
        end_rt = 1;
      }
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

    
      ///////////////////////////////////////DETERMINE EMAIL SUBJECT      
      if (nextProcApprove[0] == "3") {
          mlsubj = "Approver";
      } else {
          mlsubj = "Reviewer";
      }
    
    
      //////////////////////////////////////CHECK IF ACTION IS APPROVE OR REJECT OR CANCEL APPROVE
      /////////////////////////////////////////////////////////////////////////////////////////////
    
    
      if (txtJudge == "a") { /////////////APPROVE

          ml   = nextProcApprove[1];

          mlcc = "admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue(); //ORIGINAL
          mlcc = mlcc.toString().replace(" ",""); //ADDED BY WNP 20210707 TO FIX CASE SPACE WITH EMAIL ADDRESS - ERROR INVALID ADDRESS
        
        
          //return ml;
          msg = "WorkFlow Launcher have been approved. Please perform next process.\r\n"
              //+ "This Application is not completed until Application Admin processed the signed document.\r\n\r\n＜Link＞\r\n" + txtUrl
               + "This Application is not completed until Application Admin processed the signed document.\r\n"
               + "\r\n＜Application sheet＞\r\n" + txtUrl

              if (nextProcApprove[1] != "") {

                  msg = "WorkFlow Launcher have been approved. Please perform next process.\r\n"
                       + "This Application is not completed until Application Admin processed the signed document.\r\n\r\n＜Link＞\r\n" + txtUrl
                       + "\r\n";//＜Application sheet＞\r\n" + doc_url

              } else {

                  msg = "WorkFlow Launcher have been approved. Please perform next process.\r\n"
                       + "\r\n\r\n＜Application sheet＞\r\n" + txtUrl;

              }
        
          //////////////UPDATE STATUS TO DATABASE SHEET AND DATA SHEET///////////////////

          ss1.getRange("C" + txtId).setValue('Final Approved');
          act = "Final Approve - Signed";
          e_sign.getRange("C" + rw).setValue("Final Approved"); //ActionS
        
        
              ///////////////////////////////////SEND EMAIL

              //var t_3l_s = new Date();
                m_s = ("|,"+ml+","+mlcc+",|").replace(",,,",",").replace(",,",",").replace("|,","").replace(",|","")
                //return "WFL : <" + urgent_mail + "" + text_ref1 + ":" + pk + "> has been Approved.";
                MailApp.sendEmail(m_s, "" + "WFL : <" + urgent_mail + "" + text_ref1 + ":" + pk + "> has been Approved.", msg);

        
    
        
      } else if (txtJudge == "r") { ////////////////////////////REJECT
        
          ml = ss1.getRange("F" + txtId).getValue();
        
          mlcc = "admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue(); //ORIGINAL
          mlcc = mlcc.toString().replace(" ",""); //ADDED BY WNP 20210707 TO FIX CASE SPACE WITH EMAIL ADDRESS - ERROR INVALID ADDRESS
        
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
          msg = "Your application has been rejected by approver.\r\n\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n\r\n"+ txtUrl;//＜Application Sheet＞\r\n" + doc_url;
          ss1.getRange("C" + txtId).setValue('Rejected by Final Approver');
          act = "Final Approve - Rejceted";
          e_sign.getRange("C" + rw).setValue("Final Approve Rejceted"); //Action
        
              ///////////////////////////////////SEND EMAIL

              //var t_3l_s = new Date();

                //MailApp.sendEmail(mlcc + mlregistcc,"" + ss1.getSheetByName(INPUT_DATA).getRange("D" + txtId).getValue() + " Application <Rejected by Final Approver>", msg);
                m_s = ("|,"+ml+","+mlcc+",|").replace(",,,",",").replace(",,",",").replace("|,","").replace(",|","") 
                MailApp.sendEmail(m_s, "" + "WFL <Rejected by Final Approver>" + text_ref1 + ":" + pk, msg, { cc: mlcc});
        
        
          
          /////////////////////////////////ADDED BY WP 20210916
          /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEET
          /*
          /////////////////////////////////////////REMOVE SIGNATURE FROM TEMPLATE FILE
          //sheet = file_for_check_esig.getSheetByName(master_sheet_name);
          var search_string = Session.getActiveUser().getEmail();

          if (search_string != "admin@ngkntk-asia.com") {
            var textFinder = sheet.createTextFinder(search_string);
            search_row = textFinder.findNext();
            
            if (search_row != null) {

                do {
                  search_row.clearContent();
                  search_row = textFinder.findNext(); 
                  SpreadsheetApp.flush(); //ADDED BY WP 20210917                    
                } while (search_row!= null); 
              
            }
          }
        
        
          SpreadsheetApp.flush(); //ADDED BY WP 20210917  
          /////////////////////////////////END OF ADDED BY WP 20210916
          /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEET
          */
      }
    

        //////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////
        //////ADDED BY WP 20210914 - 'CANCEL APPROVE' BUTTON
        //////TO RESTORE STAGE BACK BEFORE FINAL APPROVE
        //////NO NEED TO SEND EMAIL
                 
        
      else if (txtJudge == "c") {   //ACTION CANCELL
        
        
          /////////////////////////////////ADDED BY WP 20210916
          /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEETS
        

        

          ////////////////////////////////////////REMOVE SIGNATRUE FROM TEMPLATE FILE
          //////////////////////////////////////////FIND TEMPLATE FILE
          /*

          var FolderTempUserEmail;
          var FolderTempID = "1T0JMNMr4Bw94RMcu8TFkZl6K_B4HH37S";
          var FolderTempUser = DriveApp.getFolderById(FolderTempID); //ACCESS TO FOLDER 'TEMPLATE_USERS' 
          var user_email = Session.getActiveUser().getEmail();
          FolderTempUserEmail =  FolderTempUser.getFoldersByName(user_email); //ACCESS TO FOLDER OF THE CURRENT USER
        
          var file_temp,sht_temp;
          var active_record = ss1.getRange("A" + txtId+":CK"+txtId).getValues();
          var doc_type = active_record[0][3]; //GET TEMPLATE NAME OF THIS DOCUMENT
        
          file_temp = FolderTempUserEmail.getFilesByName(doc_type); //GET TEMPLATE FILE FROM FOLDER OF THE CURRENT USER
          sht_temp = file_temp.getSheetByName(master_sheet_name);

          var search_string = user_email;

          if (search_string != "admin@ngkntk-asia.com") {
            var textFinder = sht_temp.createTextFinder(search_string);
            search_row = textFinder.findNext();
            
            if (search_row != null) {

                do {
                  search_row.clearContent();
                  search_row = textFinder.findNext(); 
                  SpreadsheetApp.flush(); //ADDED BY WP 20210917                    
                } while (search_row!= null); 
              
            }
          }
          */
          ////////////////////////////////////////END OF REMOVE SIGNATRUE FROM TEMPLATE FILE

          ////////////////////////////////////////REMOVE SIGNATRUE FROM DATA FILE          
          //sheet = file_for_check_esig.getSheetByName(master_sheet_name);
          var search_string = Session.getActiveUser().getEmail();

          if (search_string != "admin@ngkntk-asia.com") {
            var textFinder = sheet.createTextFinder(search_string);
            search_row = textFinder.findNext();
            
            if (search_row != null) {

                do {
                  search_row.clearContent();  //MODIFIED BY WP 20210917
                  SpreadsheetApp.flush(); //ADDED BY WP 20210917
                  search_row = textFinder.findNext(); 
                } while (search_row!= null);                
            }
          }
          /////////////////////////////////END OF ADDED BY WP 20210916
          /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEET

        ss1.getRange("C" + txtId).setValue('Reviewed'); //BACK STAGE
        ss1.getRange("H" + txtId).setValue(doc_url);
        act = "Cancel Action";
        e_sign.getRange("C" + rw).setValue("Cancel Action"); //KEEP ACTION LOG    
        
        ///////////////////CLEAR DATA - RESTORE TO BLANK (Approve Cells)
        ss1.getRange("P" + txtId).setValue("");
        ss1.getRange("R" + txtId).setValue("");
        ss1.getRange("S" + txtId).setValue("");
        ss1.getRange("T" + txtId).setValue("");
        ss1.getRange("BH" + txtId).setValue("");
        //returnVal = 99;
        
        ///////////////////CLEAR DATA - RESTORE TO BLANK (Register Cells)

       // if ( end_rt==1) { //20211001
          ss1.getRange("V" + txtId).setValue("");  
          ss1.getRange("U" + txtId).setValue("");
          ss1.getRange("W" + txtId).setValue("");
          ss1.getRange("X" + txtId).setValue("");
          ss1.getRange("Y" + txtId).setValue("");
        //------------------------------------------------
        
          ml = ss1.getRange("F" + txtId).getValue();
        
          mlcc = "admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue(); //ORIGINAL
          mlcc = mlcc.toString().replace(" ",""); //ADDED BY WNP 20210707 TO FIX CASE SPACE WITH EMAIL ADDRESS - ERROR INVALID ADDRESS
        
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
          msg = "Approver has cancelled the previous action, the document is now pending for Approver.\r\n\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n\r\n"+ txtUrl;//＜Application Sheet＞\r\n" + doc_url;
          ss1.getRange("C" + txtId).setValue('Reviewed');
          act = "Final Approve - Rejceted";
          e_sign.getRange("C" + rw).setValue("Action Cancelled"); //Action
        
              ///////////////////////////////////SEND EMAIL
              //MailApp.sendEmail(mlcc + mlregistcc,"" + ss1.getSheetByName(INPUT_DATA).getRange("D" + txtId).getValue() + " Application <Action Cancelled by Final Approver>", msg);
              m_s = ("|,"+ml+","+mlcc+",|").replace(",,,",",").replace(",,",",").replace("|,","").replace(",|","") 
              MailApp.sendEmail(m_s, "" + "WFL <Action Cancelled by Final Approver>" + text_ref1 + ":" + pk, msg, { cc: mlcc});
  
      }
    
   
        //////END OF ADDED BY WP 20210914 - 'CANCEL APPROVE' BUTTON
        ////////////////////////////////////////////////////////////////////////////          
        ////////////////////////////////////////////////////////////////////////////  

  }

/////////////////////////////////////////////////////////////
////////////////////////END OF APPROVAL /////////////////////
/////////////////////////////////////////////////////////////

  //if (txtRt == "4" || end_rt == 1) {
  if ( (txtRt == "4" || end_rt == 1) && (txtJudge != "c")) { //20211001
    
      //var t_3j_s = new Date();
      act = "Register Signed";

      //ss2 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName(name);
      //ss2.getSheetByName(INPUT_DATA).activate();
      //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=9&view=1";
      txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=9&view=1";
      
      msg = "Process completed.\r\n\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n\r\n"+ txtUrl;//＜Application Sheet＞\r\n" + doc_url;
      ss1.getRange("C" + txtId).setValue('Registerd');
      ss1.getRange("U" + txtId).setValue(txtComment);
      if ( end_rt==1) {
      ss1.getRange("V" + txtId).setValue("<Auto Register After Approved>");  
      } else {
      ss1.getRange("V" + txtId).setValue(Session.getActiveUser().getEmail());
      }
      ss1.getRange("H" + txtId).setValue(doc_url);
      ss1.getRange("W" + txtId).setValue(doc_url);
      ss1.getRange("X" + txtId).setValue(doc_id);
      ss1.getRange("Y" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //return "track : " + ss1.getName()
      //------------------------------------------------
      e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
      rw = e_sign.getLastRow() + 1;

      e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
      e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
      e_sign.getRange("C" + rw).setValue("Registerd"); //Action
      e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
      e_sign.getRange("E" + rw).setValue(txtComment); //Comment
    
      //var t_3j_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
      //var t_3f_s = new Date();
      //next = getNextApproveStep(docPID,"2A");
      //var t_3f_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);
      //var t_3l_s = new Date();
      mlcc ="admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue();
      //MailApp.sendEmail(mlcc , "" + "WFL : <"+ urgent_mail  + text_ref1 + ":" + pk + "> has been Registered.", msg);
      
      if (pk.indexOf('ANGK') > -1) {
        if (pk.indexOf('Items') > -1) {
          //Logger.log("Found DBSheet Items of ANGK");
        } else {
          MailApp.sendEmail(mlcc , "" + "WFL : <"+ urgent_mail  + text_ref1 + ":" + pk + "> has been Registered.", msg);
        }
      } else {
        MailApp.sendEmail(mlcc , "" + "WFL : <"+ urgent_mail  + text_ref1 + ":" + pk + "> has been Registered.", msg);
      }
      
    //var t_3l_e = new Date();
      //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

  }
/*
   if (txtTempFileId!="") { 
        // Copy Data To WorkSheet  
               //.openByUrl(ss1.getSheetByName(INPUT_DATA).getRange("H" + docPID).getValue())
               var sheets_c = SpreadsheetApp.openById(txtTempFileId).getSheets() ; //SpreadsheetApp.open(doc_id).getSheets();
               var file_des = SpreadsheetApp.openByUrl(ss1.getRange("H" + docPID).getValue());
     //return "OK"          
     //return sheets_c
              for (var i = 0; i < sheets_c.length ; i++ ) {
                 var sheet_c = sheets_c[i];
                 if (sheet_c.getName()!= "Template") {
                   if (file_des.getSheetByName(sheet_c.getName())!=null) {
                     var rng_all = sheet_c.getDataRange().getValues();
                     //var doc_i = SpreadsheetApp.open(file_temp);
                     file_des.getSheetByName(sheet_c.getName()).getRange(1, 1, rng_all.length, rng_all[0].length).setValues(rng_all);
                   }  
                 }    
               }
  }
  */
  //START WRITE LOG
  //var t3 = new Date();
  ss3 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  //ss3.getSheetByName("Logs").activate();
  ss3.getSheetByName("Logs").appendRow([
          　　　 pk,
          txtCompany,
          act,
          txtMaster,
          txtComment, // + " - by - " + Session.getActiveUser().getEmail(),
          Session.getActiveUser().getEmail(),
          Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
          doc_url,
          doc_id,
          Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"), txtUrl
      ]);

  //END FOR WRITE LOG
  //var t2 = new Date();
/* 
var diff = Math.floor((t2 - t1) / (1000.0)); //seconds
  var diff_log = Math.floor((t2 - t3) / (1000.0)); //seconds
  var diff_wrt = Math.floor((t3 - t4) / (1000.0)); //seconds
  var diff_CCregmail = Math.floor((t4 - t5) / (1000.0)); //seconds
  var diff_regmail = Math.floor((t5 - t6) / (1000.0)); //seconds
*/
  ////Logger.log("Over All :"+diff+" Second");
  ////Logger.log("+ WRT LOG :"+diff_log+" Second");
  ////Logger.log("+ WRT SHEET :"+diff_wrt+" Second");
  ////Logger.log("+ GET CC REGIST MAIL :"+diff_CCregmail+" Second");
  ////Logger.log("+ GET  REGIST MAIL :"+diff_regmail+" Second");

  if (txtJudge == "c") {
  
    return 99; //IN CASE CANCEL ACTION
  
  } else if (txtJudge == "r") {
  
    return 98; //IN CASE REJECT
  
  } else if (txtJudge == "a") {
  
    return 100; //IN CASE APPROVE
  
  } else {
    
    return 10; //IN CASE REROUTE
  }

  ss1 = null;
  ss3 = null;
  sheet = null;
  e_sign = null;

}