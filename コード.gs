/*
* transfer pdf & send mail
* Author  :
* Version :0.1
* Create  :0.1 2019/11/26 kawagoe 新規作成
* Update  :1.0 2019/11/26 メインリリース
* Etc     :
*/

/*
#####################################################
# 定義ファイル読み込み
#####################################################
*/
var Cfg = {
  debug       : true,
  
  configSheetName : "設定",
  fromName    : "B1",
  fromAddress : "B2",
  subject     : "B3",
  mailBody1   : "B4",
  mailBody2   : "B5",
  targetSheetName : "送信リスト",
  colNumberRow    : 2,
  stRow           : 3,
  checkName : [""], //誤送信防止用の送信者アドレスを指定（カンマ）
  pdfURL : "https://docs.google.com/TYPE/d/SSID/export?"
};


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu")
  .addItem("メール一斉送信", "listCheckSendEmail")
  .addSeparator()
  .addItem("初回に実行してください（実行権限の承認）", "dummyFunction")
  .addToUi()
};

function dummyFunction(){
}



function listCheckSendEmail() {
  var ret, folder = "";
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var userName = Session.getEffectiveUser().getEmail();
  var nowActiveSheet = ss.getActiveSheet();
  var sheet = ss.getSheetByName(Cfg.targetSheetName).activate();
  SpreadsheetApp.flush();
  
  //
  var CDATA = getColNumber_(sheet);
  
  var check = 0;
//  for ( var i in Cfg.checkName ) {
//    Logger.log(Cfg.checkName[i] + " = " + userName)
//    if ( userName == Cfg.checkName[i] ) {
//      check++;
//    }
//  }
//  if ( 0 == check ) {
//    ss.toast(userName + "では利用できません: " + check, "WARNING", 5);
//    LogSheet("WARNING",userName + "では利用できません: " + check);
//    ui.alert("警告",userName + "では利用できません: " + check ,ui.ButtonSet.OK);
//    return;
//  }
  ss.toast(userName + "でメール送信します","INFO", 5);
  
  var quota = MailApp.getRemainingDailyQuota();
  var startCheck = ui.alert("【警告】",
                            "本機能によりメールが送信されます。\n"
                            + "添付ファイル、宛先を十分に確認して下さい。\n\n"
                            + "中断する場合は：[キャンセル]ボタンをクリックします。\n"
                            + "現在のメール送信可能件数：" + quota + "/1日"
                            , ui.ButtonSet.OK_CANCEL);
  nowActiveSheet.activate();
  SpreadsheetApp.flush();
  if ( startCheck == "OK" ) {
  }else{
    ss.toast("処理がキャンセルされました", "INFO", 5);
    LogSheet("INFO","処理がキャンセルされました");
    return;
  }
                                                

  try{
    var token = ScriptApp.getOAuthToken();
    var tgtSs = ss.getSheetByName(Cfg.targetSheetName);
    var lastRow = tgtSs.getLastRow() - Cfg.stRow + 1;
    if(lastRow <= 0){
      ss.toast("データが存在しません", "INFO", 5);
      LogSheet("INFO","データが存在しません");
      return;
    }
    var tgtValues = tgtSs.getRange(Cfg.stRow, 1, lastRow, tgtSs.getLastColumn()).getValues();
    
    //mail情報取得
    var mailSs = ss.getSheetByName(Cfg.configSheetName);
    var fromName = mailSs.getRange(Cfg.fromName).getValue();
    var fromAddress = mailSs.getRange(Cfg.fromAddress).getValue();
    var subject = mailSs.getRange(Cfg.subject).getValue();
    var mailBody1 = mailSs.getRange(Cfg.mailBody1).getValue();
    var mailBody2 = mailSs.getRange(Cfg.mailBody2).getValue();
    
    //メール送信処理
    var skipCnt = 0;
    var noTgtCnt = 0;
    var toErrCnt = 0;
    var exeCnt = 0;
    
    //行ごとに処理
    for (var i = 0; i < tgtValues.length; i++) {
      var body = mailBody1 + "\n" + mailBody2;
      var fileObjArray = [];
      var fileNameArray = [];
      var attachmentArray = [];
      
      // 配信状態を設定
      var check = tgtValues[i][CDATA["CHECK"]];
      if(check == false){
        skipCnt++;
        continue;
      }
      
      // 対象行の宛先や本文、添付情報を取得
      var to  = tgtValues[i][CDATA["TO"]];
      var cc  = tgtValues[i][CDATA["CC"]];
      var bcc = tgtValues[i][CDATA["BCC"]];
      LogSheet("INFO","["+i+"]" + to);
      var insertData1 = tgtValues[i][CDATA["BODY1"]];
      var insertData2 = tgtValues[i][CDATA["BODY2"]];
      var insertData3 = tgtValues[i][CDATA["BODY3"]];
      var insertData4 = tgtValues[i][CDATA["BODY4"]];
//      LogSheet("debug", insertData1)
            
      //添付フラグ読み取り
      var attachmentFlag = tgtValues[i][CDATA["ATTACHMENTFLAG"]];
      var pdfFlag = tgtValues[i][CDATA["PDFFLAG"]];
      var deleteFlag = tgtValues[i][CDATA["PDFDELTEFLAG"]];
      
      //添付フラグありなら添付処理開始
      var folderid
      if(attachmentFlag == true && tgtValues[i][CDATA["ATTACHMENTFLAG"]] != ""){
        var fileUrls = tgtValues[i][CDATA["ATTACHMENTURL"]].split(',');
        loop: for(var j in fileUrls){
          var id = fileUrls[j].match(/[-\w]{25,}/);
          try{
            DriveApp.getFileById(id);
          }catch(e){
            LogSheet("debug", "ファイルが存在しません。: " + id);
            continue loop;
          }
          Logger.log("id : " + id);
          if(pdfFlag == false){
            var fileObj = DriveApp.getFileById(id);
          }else{
            //格納先
            if(folderid){
              //指定フォルダにPDFを作成
              var folder = DriveApp.getFolderById(folderid);
            }else{
              //指定なければ、元シートのカレントディレクトリに作成
              folderid = DriveApp.getFileById(id).getParents().next().getId();
              var folder = DriveApp.getFolderById(folderid);
            }
            //PDF化
//            var pdfBlob = DriveApp.getFileById(id).getAs(MimeType.PDF);
            //ドライブにファイルを生成する
//            var fileObj = folder.createFile(pdfBlob);
            var pdfRet = createPDF_(id, "", folderid);
            var fileObj = pdfRet.obj;
          }
          
          fileObjArray.push(fileObj);
          fileNameArray.push(fileObj.getName());
          attachmentArray.push(fileObj.getBlob());
        } //for j
      }//if
      
      subject = subject.replace("{{データ1}}", insertData1, "g");
      subject = subject.replace("{{データ2}}", insertData2, "g");
      subject = subject.replace("{{データ3}}", insertData3, "g");
      subject = subject.replace("{{データ4}}", insertData4, "g");
      body = body.replace("{{データ1}}", insertData1, "g");
      body = body.replace("{{データ2}}", insertData2, "g");
      body = body.replace("{{データ3}}", insertData3, "g");
      body = body.replace("{{データ4}}", insertData4, "g");
      body = body.replace("{{添付ファイル}}", fileNameArray.join("\n"), "g");
      
      //メール配信
      if(Cfg.debug){
        ui.alert(subject, body, ui.ButtonSet.OK);
      }else{
        var ret = sendEmail_(to, subject, body, attachmentArray, cc, bcc, fromName, fromAddress);
        
        // 配信状態を設定
        var date = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
        if(ret){
          tgtSs.getRange(Cfg.stRow + i, CDATA["LOG1"] + 1, 1, 2).setValues([["配信済: " + date, "PDF化処理: " + pdfFlag + ", Folder: " + folder]]);
          tgtSs.getRange(Cfg.stRow + i, CDATA["LOG2"] + 1).setValue(false);
          SpreadsheetApp.flush();
          exeCnt++
        }else{
          tgtSs.getRange(Cfg.stRow + i, CDATA["LOG1"] + 1, 1, 2).setValues([["配信失敗: " + date, ""]]);
          toErrCnt++
        }
      }
      
      //変換ファイルの後処理
      if(attachmentFlag == true && pdfFlag == true && deleteFlag == true){
        for(var k in fileObjArray){
          var delFileName = fileObjArray[k].getName();
          ss.toast(delFileName + "を削除します","INFO", 5)
          LogSheet("INFO",delFileName + "を削除します");
          fileObjArray[k].setTrashed(true);
        }
      }
      if(attachmentFlag == true && pdfFlag == true && deleteFlag == false){
        for(var l in fileObjArray){
          var delFileName = fileObjArray[l].getName();
          ss.toast(delFileName + "を退避します","INFO", 5)
          LogSheet("INFO",delFileName + "を退避します");
          fileObjArray[l].setName("[PDF]" + delFileName);
        }
      }
      
    }//for i
    // 終了確認ダイアログを表示
    
//    ui.alert("確認", "メール一斉送信が完了しました。", ui.ButtonSet.OK);
    ss.toast("メール一斉送信が完了しました。","INFO", 5)
    LogSheet("INFO","MailSend メール一斉送信が完了しました。 RESULT: AllNum: " + (Cfg.stRow + i + 1) + ", exeCnt: " + exeCnt + ", skipCnt: " + skipCnt + ", toErrCnt: " + toErrCnt + ", noTgtCnt: " + noTgtCnt);
    
    // 処理を開始する行番号を初期
  }catch(e){
    throw new Error("エラー " + e.name + ", 行 " + e.lineNumber + ", 内容 " + e.message);
    LogSheet("ERROR","エラー " + e.name + ", 行 " + e.lineNumber + ", 内容 " + e.message);
  }
  
};



/*
#####################################################
# メール送信
#　to, subject, body, attachments[blob]
#####################################################
*/
function sendEmail_(to, subject, body, attachments, cc, bcc, fromName, fromAddress){
  try{
    var options = {
      cc          : cc,
      bcc         : bcc,
      name        : fromName,
      attachments : attachments,
      from        : fromAddress
      //    noReply     : true
    };
    GmailApp.sendEmail(to, subject, body, options);
    return true;
  }catch(e){
    return false;
  }
}



/*
* #####################################################
* # createPDF
* #####################################################
* @param {string} - ssid
* @param {string} - filename
* @param {string} - folderid
* @param {string} - sheetid
* @param {string} - range
* @return {string} - none
*/
//https://developers.google.com/drive/api/v3/reference/files/export
//https://developers.google.com/drive/api/v3/manage-downloads#
function createPDF_(ssid, filename, folderid, gid, range){
  var delFlag = 0;
  var opts = "", setType = "";
  
//  debugger;
  if(range != "" && range != undefined){
    opts["range"] =  range;   // シート範囲　例）A1:D5
  }
  if(gid != "" && gid != undefined){
    opts["gid"] =  gid;       // シートIDを指定
  }
  
  //格納先
  if(folderid){
    //指定フォルダ
    var folder = DriveApp.getFolderById(folderid);
  }else{
    //指定なければ、元シートのカレントディレクトリに作成
    folderid = fileObj.getParents().next().getId();
    var folder = DriveApp.getFolderById(folderid);
  }
  
  var fileObj = DriveApp.getFileById(ssid);
  var filename = baseName(fileObj.getName()) + ".pdf";
  var type = fileObj.getMimeType();
//  debugger;
  Logger.log(type);
  LogSheet("debug",type + ", " + ssid);
  
  var docRet = getDocumentType(ssid, folder);
  ssid    = docRet.ssid;
  setType = docRet.setType;
  opts    = docRet.opts;
  delFlag = docRet.delFlag;
  
  //URL
  var url = Cfg.pdfURL.replace("SSID", ssid);
  url = url.replace("TYPE", setType);
  var url_ext = [];
  for( var optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }
  var options = url_ext.join("&");
  var token = ScriptApp.getOAuthToken();
  url+=options;
  Logger.log(url);
//  debugger;
  var response = UrlFetchApp.fetch(url, {
                                   'headers': {
                                   'Authorization': 'Bearer ' +  token
                                   }
                                   });
  
  //BLOB作成
  var blob = response.getBlob();
  
  //PDF作成
  var fileObj = folder.createFile(blob).setName(filename);
  
  //Officeの変換後のGoogleドキュメントは削除
  if(delFlag == 1){
    DriveApp.getFileById(ssid).setTrashed(true);
    LogSheet("debug", "Office変換後のGoogleドキュメントを削除しました。" + ssid);
  }
  
  return {
    url : url + options, //ログ用
    blob : blob,         //メール添付用
    obj  : fileObj
  };
}


function getDocumentType(ssid, folder){
  var setType, delFlag;
    //(documentごとのOption整理が必要)
  var opts = {
    format:               "pdf",         //export format [pdf,,xlsx,ods,csv,tsv] [zip : zipの場合このオプションのみ]
    size:                 "A4",          //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
    fzr:                  "false",       // 固定行の表示有無
    fzc:                  "false",       //true/false
    portrait:             "true",        //trueで縦出力(Potrait)、falseで横出力(Landscape)
    fitw:                 "true",        // 幅を用紙に合わせるか trueでフィット、falseで原寸大
    gridlines:            "false",       // グリッドラインの表示有無//true/false
    printtitle:           "false",       // スプレッドシート名をPDF上部に表示するか
    sheetnames:           "false",       // シート名をPDF上部に表示するか
    pagenum:              "false",       // ページ番号の有無
    attachment:           "false",       //true/false ?
    scale:                "2",           //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
    printnotes:           "false",       //コメントを印刷するかどうか true/false
    pageorder:            "2",           //1で前から後ろ(Down, then over)、2で後ろから前(Over, then down)
//    top_margin:           "0.75",        //All four margins must be set (0.75)
//    bottom_margin:        "0.75",        //All four margins must be set (0.75)
//    left_margin:          "0.75",        //All four margins must be set (0.75)
//    right_margin:         "0.75",        //All four margins must be set (0.75)
//    horizontal_alignment: "LEFT",      //LEFT/CENTER/RIGHT
//    vertical_alignment:   "TOP",         //TOP/MIDDLE/BOTTOM
  };
  
  //check document type
  //  try{
  switch(type){
    case "application/vnd.google-apps.spreadsheet" :
      setType = "spreadsheets";
      break;
    case "application/vnd.google-apps.presentation" :
      setType = "presentation";
      Cfg.pdfURL = "https://docs.google.com/presentation/d/SSID/export/pdf"
      opts = ""; //オプションなし
      break;
    case "application/vnd.google-apps.document" :
      setType = "document";
      break;
    case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" :
      ssid = convertOffice2Google_(ssid, folder, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx");
      setType = "spreadsheets";
      delFlag = 1;
      break;
    case "application/vnd.openxmlformats-officedocument.presentationml.presentation" :
      ssid = convertOffice2Google_(ssid, folder, "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx");
      setType = "presentation";
      Cfg.pdfURL = "https://docs.google.com/presentation/d/SSID/export/pdf"
      opts = ""; //オプションなし
      delFlag = 1;
      break;
    case "application/vnd.openxmlformats-officedocument.wordprocessingml.document" :
      ssid = convertOffice2Google_(ssid, folder, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx");
      setType = "document";
      delFlag = 1;
      break;
    default:
  }
  
  return {
    ssid    :ssid, 
    setType :setType,
    opts    :opts,
    delFlag :delFlag
  }
  
//  }catch(e){
//    return false;
//  }}
}

/*
#####################################################
# 関数：Excelをシートへ変換
#####################################################
//code: https://gist.github.com/azadisaryev/ab57e95096203edc2741
https://developers.google.com/apps-script/reference/base/mime-type
*/
function convertOffice2Google_(id, folder, contentType, extension){
  //アクセストークンを取得 //add manifest:
  var token = ScriptApp.getOAuthToken();
  var blob = DriveApp.getFileById(id).getBlob();
  var fileName = "[OFFICE]" + baseName(blob.getName());
  var folderId = folder.getId();
  
  //ファイル変換パラメータ
  var uploadParams = {
    method:'post',
    contentType: contentType,
    contentLength: blob.getBytes().length,
    headers: {'Authorization': 'Bearer ' + token},
    payload: blob.getBytes()
  };
  // マイドライブへの変換リクエスト
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());
//  debugger;
  
  // 指定したフォルダへファイルを変換する際のパラメータ
  var payloadData = {
    title: fileName,
    parents: [{id: folderId}]
  };
  // ファイル名と格納先パラメータ
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + token},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  // 変換したファイルを受け取る
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
  //変換後のスプレッドシートのファイルIDを取得して返す。
  return fileDataResponse.id;
}

function baseName(str){
   var base = new String(str).substring(str.lastIndexOf('/') + 1); 
    if(base.lastIndexOf(".") != -1)       
        base = base.substring(0, base.lastIndexOf("."));
   return base;
}


//Get column number for each items
function getColNumber_(sheet){
  var initData = sheet.getSheetValues(Cfg.colNumberRow, 1, 1, sheet.getLastColumn());
  return {
    NO    : initData[0].indexOf("NO"),
    CHECK : initData[0].indexOf("CHECK"),
    TO    : initData[0].indexOf("TO"),
    CC    : initData[0].indexOf("CC"),
    BCC   : initData[0].indexOf("BCC"),
    BODY1 : initData[0].indexOf("BODY1"),
    BODY2 : initData[0].indexOf("BODY2"),
    BODY3 : initData[0].indexOf("BODY3"),
    BODY4 : initData[0].indexOf("BODY4"),
    ATTACHMENTFLAG : initData[0].indexOf("ATTACHMENTFLAG"),
    PDFFLAG        : initData[0].indexOf("PDFFLAG"),
    PDFDELTEFLAG   : initData[0].indexOf("PDFDELTEFLAG"),
    ATTACHMENTURL  : initData[0].indexOf("ATTACHMENTURL"),
    LOG1  : initData[0].indexOf("LOG1"),
    LOG2  : initData[0].indexOf("LOG2")
  }
}
