// BillUtil 用 Script Properties
const BILL_UTIL_PROPS = PropertiesService.getScriptProperties();

// 請求書の雛形ID
var normalBillSheetId = BILL_UTIL_PROPS.getProperty("NORMAL_BILL_SHEET_ID");
var invoiceBillSheetId = BILL_UTIL_PROPS.getProperty("INVOICE_BILL_SHEET_ID");

// 請求書初期化用シートID
var resetNormalBillSheetId = BILL_UTIL_PROPS.getProperty("RESET_NORMAL_BILL_SHEET_ID");
var resetInvoiceBillSheetId = BILL_UTIL_PROPS.getProperty("RESET_INVOICE_BILL_SHEET_ID");

// データベースシートID
var databaseSheetId = BILL_UTIL_PROPS.getProperty("DATABASE_SHEET_ID");

// 消費税率
const taxPercentage = Number(BILL_UTIL_PROPS.getProperty("TAX_PERCENTAGE") || "10");

/**
 * 勤怠シートからクローリングして勤務時間を計算
 * @param {string} attendanceSheetId - 勤怠シートのID
 * @param {number} aggregationMonth - 集計月
 * @param {number} aggregationYear - 集計年
 * @returns {Object.<string, number>} - 連想配列{PJ名:勤務時間(msec)}
*/
function calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear) {
  //シートの準備
  const attendanceSheet = SpreadsheetApp.openById(attendanceSheetId);
  const sheet = attendanceSheet.getSheets()[0];
  const values = sheet.getRange('B2:E'+ sheet.getLastRow()).getValues();
  //計算、msec単位であることに注意
  let projectWorkTimes = {};
  for (let i = 0; i < values.length; i++){
    const projectName = values[i][0];
    const startTimeDate = new Date(values[i][2]);
    const startTimeMonth = startTimeDate.getMonth();
    const startTimeYear = startTimeDate.getFullYear();
    const endTimeDate = new Date(values[i][3]);
    const endTimeMonth = endTimeDate.getMonth();
    const endTimeYear = endTimeDate.getFullYear();
    //プロジェクトが初めて登場した場合に初期値を設定
    //年判定
    if(startTimeYear === aggregationYear || endTimeYear === aggregationYear) {
      if (!projectWorkTimes[projectName]) {
        projectWorkTimes[projectName] = 0;
      }
      //月跨ぎ判定、1ヶ月丸ごと勤務による2ヶ月ずれは考慮していない
      //またいでない時
      if (startTimeMonth === endTimeMonth) {
        if(startTimeMonth === aggregationMonth - 1) {
          const workTimes = endTimeDate - startTimeDate;
          projectWorkTimes[projectName] += workTimes;
        }
      }//またいでる時
      else {
        if (startTimeMonth === aggregationMonth - 1 && startTimeYear === aggregationYear) {
          const lastTimeOfMonth = new Date(startTimeDate.getFullYear(),  startTimeDate.getMonth() + 1, 1);
          const workTimes = lastTimeOfMonth - startTimeDate;
          projectWorkTimes[projectName] += workTimes;
        }
        else if (endTimeMonth === aggregationMonth - 1 && endTimeYear === aggregationYear) {
          const firstTimeOfMonth = new Date(endTimeDate.getFullYear(), endTimeDate.getMonth(),1);
          const workTimes = endTimeDate - firstTimeOfMonth;
          projectWorkTimes[projectName] += workTimes;
        }
      }
    }
  }
  return projectWorkTimes;
}

/**
 * 勤怠シートと個人情報シートのからクローリングして請求書に書き込んでドライブに出力
 * @param {string} attendanceSheetId - 勤怠シートのID
 * @param {string} driveFolderId - 出力ドライブのID
 * @param {number} aggregationMonth - 集計月
 * @param {number} aggregationYear - 集計年
 * @param {number} thatYear - 実行年（デバッグ用）
 * @param {number} thatMonth - 実行月（デバッグ用）
 * @param {number} thatDate - 実行日（デバッグ用）
/**
 * @returns {Object.<string, string|number>} - 連想配列 {status: エラーステータス, slackUserId: スラックID, aggregationYear: 集計年, aggregationMonth: 集計月, billUrl: 出力請求書URL}
*/
function generateBill(attendanceSheetId,driveFolderId,aggregationMonth,aggregationYear,thatYear = new Date().getFullYear(),thatMonth = new Date().getMonth() + 1,thatDate = new Date().getDate()) {
  //シートと変数の準備
  const databaseSheet = SpreadsheetApp.openById(databaseSheetId);
  const addressSheetId = databaseSheet.getSheets()[0].getRange(`E${search(databaseSheet.getSheets()[0],attendanceSheetId)[0]}`).getValue();
  const addressSheet = SpreadsheetApp.openById(addressSheetId);
  let slackUserId = databaseSheet.getSheets()[0].getRange(`C${search(databaseSheet.getSheets()[0],attendanceSheetId)[0]}`).getValue();
  let billUrl;
  //インボイス請求書かを判断
  //インボイス番号はT（十三桁の半角数字）の形
  let bill;
  let invoinceNumber;
  const addressArray = addressSheet.getSheets()[0].getRange('A1:E7').getValues();
  const invoiceOrNot = addressArray[4][1];
  let status;
  try { 
    //インボイスでない場合
    if (invoiceOrNot === 'いいえ') {
      bill = SpreadsheetApp.openById(normalBillSheetId).getActiveSheet();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
      //請求書シートを初期化
      resetBill(normalBillSheetId);
      //請求書シートに計算結果を記述
      //通し番号更新
      const lastInvoiceNo = Number(addressArray[6][1] || 0);
      const nextInvoiceNo = lastInvoiceNo + 1;

      addressSheet.getRange('B7').setValue(nextInvoiceNo);
      const workTimeDict = msecDictToHourAndMinutesDict(hourlyWageDict);
      let rowPointer = 15;//請求項目の書き始めは15行目
      for (let key in hourlyWageDict) {
        if (hourlyWageDict[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDict[key]/3600000);
          const setArray = [[`インターン${aggregationYear}年${aggregationMonth}月稼働分(${key}円×${workTimeDict[key]})`,'','',1,'式',price]];
          bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      //請求書シートに振込先を記述
      let deadline;
      if(thatDate > 10){
        deadline = new Date(thatYear,thatMonth + 1,0).toLocaleDateString('ja-JP')
      }else{
        deadline = new Date(thatYear,thatMonth ,0).toLocaleDateString('ja-JP')
      }
      const leftBlock = [['晃商インターン'],[deadline],[addressArray[0][1]],[`${addressArray[0][2]} ${addressArray[0][3]} ${addressArray[0][4]}`],[`${addressArray[2][3]} ${addressArray[2][4]}`]];
      bill.getRange('B6:B10').setValues(leftBlock);
      //請求書シートに通し番号等の右側ブロックを記述
      const rightBlock = [['No',addressArray[6][1] + 1],['請求日',new Date(thatYear,thatMonth - 1,thatDate).toLocaleDateString('ja-JP')],['',''],[`〒${addressArray[1][1]}`,''],[addressArray[1][2],''],[addressArray[1][3],''],[addressArray[2][1] + addressArray[2][2],''],[`TEL:${addressArray[3][1]}`,'']];
      bill.getRange('F2:G9').setValues(rightBlock);
      //請求書シートをpdf出力
      billUrl = spreadsheetToPdf(normalBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      //請求書シートを初期化
      resetBill(normalBillSheetId);
      status = 'success';
    }
    //インボイスの場合（プルダウンではい・いいえなのでいいえ以外）
    else {
      bill = SpreadsheetApp.openById(invoiceBillSheetId).getActiveSheet();
      invoinceNumber = addressSheet.getSheets()[0].getRange('B6').getValue();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
      //請求書シートを初期化
      resetBill(invoiceBillSheetId);
      //請求書シートに計算結果を記述
      //通し番号更新
      addressSheet.getRange('B7').setValue(addressArray[6][1] + 1);
      const workTimeDict = msecDictToHourAndMinutesDict(hourlyWageDict);
      let rowPointer = 15;
      for (let key in hourlyWageDict) {
        if (hourlyWageDict[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDict[key]/3600000);
          const setArray = [[`インターン${aggregationYear}年${aggregationMonth}月稼働分(${key}円×${workTimeDict[key]})`,'','',1,'式',price,`${taxPercentage}%`]];
          bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      //請求書シートに振込先を記述
      let deadline;
      if(thatDate > 10){
        deadline = new Date(thatYear,thatMonth + 1,0).toLocaleDateString('ja-JP')
      }else{
        deadline = new Date(thatYear,thatMonth ,0).toLocaleDateString('ja-JP')
      }
      const leftBlock = [['晃商インターン'],[deadline],[addressArray[0][1]],[`${addressArray[0][2]} ${addressArray[0][3]} ${addressArray[0][4]}`],[`${addressArray[2][3]} ${addressArray[2][4]}`]];
      bill.getRange('B6:B10').setValues(leftBlock);
      //請求書シートに通し番号等の右側ブロックを記述
      const rightBlock = [['','No',addressArray[6][1] + 1],['','請求日',new Date(thatYear,thatMonth - 1,thatDate).toLocaleDateString('ja-JP')],['','',''],[`〒${addressArray[1][1]}`,'',''],[addressArray[1][2],'',''],[addressArray[1][3],'',''],[addressArray[2][1]+addressArray[2][2],'',''],[`TEL:${addressArray[3][1]}`,'',''],[`登録番号:T${addressArray[5][1]}`,'','']];
      bill.getRange('F2:H10').setValues(rightBlock);
      //請求書シートをpdf出力
      billUrl = spreadsheetToPdf(invoiceBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      //請求書シートを初期化
      resetBill(invoiceBillSheetId);
      status = 'success'
    }
  }
  catch(e){
    status = e.message;
  }
  //ManegementUtil接続用配列
  return {status:status,slackUserId:slackUserId,aggregationYear:aggregationYear,aggregationMonth:aggregationMonth,billUrl:billUrl}
}

/**
 * シート内検索関数(指定シート内のキーワードが最初に見つかった位置を1スタートの[行,列]で返す)
 * @param {Object} targetSheet - 検索対象となるシート
 * @param {(string|number)} keyword - 検索キーワード
 * @return {[number,number]} - 配列[行,列]
*/
function search(targetSheet,keyword) {
  const results = targetSheet.createTextFinder(keyword).matchEntireCell(true).findAll()
  try{
    let row = results[0].getRow()
    let column = results[0].getColumn()
    return [row,column]}
  catch{
    return[-1,-1]
  }
}

/**
 * 与えられた連想配列{何か:時間（mesc）}を{何か:m時間n分}の連想配列に
 * @param {Object.<string, number>} workTimeDict - 連想配列{何か:時間（mesc）}
 * @returns {Object.<string, string>} - 連想配列{何か:m時間n分}
*/
function msecDictToHourAndMinutesDict(workTimeDict) {
  let workHourAndMinutesDict = { ...workTimeDict };
  //msecを3600000で割った整数部がm、3600000で割った余りを60000で割ったものがn
  for (let key in workTimeDict) {
    workHourAndMinutesDict[key] = `${Math.trunc(workTimeDict[key]/3600000)}時間${(workTimeDict[key]%3600000)/60000}分`;
  }
  return workHourAndMinutesDict;
}

/**
 * 与えられた連想配列{プロジェクト名:時間}をデータベースを参照して{時給:時間}の連想配列に
 * @param {Object.<string, (string|number)>} workTimeDict - 連想配列{プロジェクト名:時間}
 * @param {string} attendanceSheetId - 出席シートのID
 * @returns {Object.<string, (string|number)>} - 連想配列{時給:時間}
*/
function projectNameToHourlyWage(workTimeDict, attendanceSheetId) {
    // Databaseブック各シートを読み込む
    const database = SpreadsheetApp.openById(databaseSheetId);
    const userSheet = database.getSheets()[0];     // Usersシート
    const projectSheet = database.getSheets()[1];  // Projectsシート
    const upgradeSheet = database.getSheets()[2];  // 時給アップシート
    
    // UsersシートからattendanceSheetIdを検索してユーザーIDを取得
    const [userRow, userCol] = search(userSheet, attendanceSheetId);
    
    let userId = null;
    if (userRow !== -1) {
        // attendanceSheetIdが見つかった行のA列（ID列）を取得
        userId = userSheet.getRange(userRow, 1).getValue();
    } else {
        // ユーザーが見つからない場合はエラーを投げる
        throw new Error(`出席シートID「${attendanceSheetId}」に対応するユーザーが見つかりません`);
    }
    
    // プロジェクト情報を取得（A列:id, B列:プロジェクト名, C列:時給）
    const projectList = projectSheet.getRange('A2:C' + projectSheet.getLastRow()).getValues();
    
    // 時給アップ情報を取得
    const upgradeList = upgradeSheet.getRange('A2:D' + upgradeSheet.getLastRow()).getValues();
    
    // 時給アップマップを作成（user_id + project_id の組み合わせをキーとする）
    const upgradeMap = {};
    for (let i = 0; i < upgradeList.length; i++) {
        const [id, user_id, project_id, increment] = upgradeList[i];
        const key = `${user_id}_${project_id}`;
        upgradeMap[key] = increment;
    }
    
    // 時給をキーとする連想配列作成
    const hourlyWageDict = {};
    
    // 連想配列に時給ごとに時間を格納していく
    for (let i = 0; i < projectList.length; i++) {
        const projectId = projectList[i][0];      // A列のid
        const projectName = projectList[i][1];    // B列のプロジェクト名
        let projectHourlyWage = projectList[i][2]; // C列の時給
        
        // 時給アップ処理：ユーザーIDとプロジェクトIDの組み合わせで増分を確認
        const upgradeKey = `${userId}_${projectId}`;
        if (upgradeMap[upgradeKey]) {
            projectHourlyWage += upgradeMap[upgradeKey];
        }
        
        // 時給辞書の初期化
        if (!hourlyWageDict[projectHourlyWage]) {
            hourlyWageDict[projectHourlyWage] = 0;
        }
        
        // 作業時間が存在する場合に加算
        if (workTimeDict[projectName]) {
            hourlyWageDict[projectHourlyWage] += workTimeDict[projectName];
        }
    }
    
    return hourlyWageDict;
}

/**
 * 請求書初期化関数
 * @param {string} billSheetId - 請求書シートID
*/
function resetBill(billSheetId){
  //通常
  const bill = SpreadsheetApp.openById(billSheetId);
  if(billSheetId == normalBillSheetId) {
    const copyResource = SpreadsheetApp.openById(resetNormalBillSheetId).getSheetByName('初期化用');
    const copied = copyResource.copyTo(bill);
    copied.getRange('A1:H40').copyTo(bill.getSheetByName('請求書').getRange('A1:H40'));
    bill.deleteSheet(copied);    
  }//インボイス
  else if(billSheetId == invoiceBillSheetId) {
    const copyResource = SpreadsheetApp.openById(resetInvoiceBillSheetId).getSheetByName('初期化用');
    const copied = copyResource.copyTo(bill);
    copied.getRange('A1:H40').copyTo(bill.getSheetByName('請求書').getRange('A1:H40'));
    bill.deleteSheet(copied);
  }
}

/**
 * 請求書pdf化関数（urlを返す）
 * @param {string} billSheetId - 請求書シートID
 * @param {string} driveFolderId - 出力ドライブID
 * @param {number} year - 年（ファイル名用）
 * @param {number} month - 月（ファイル名用）
 * @param {string} name - 氏名（ファイル名用）
 * @returns {string} - 出力請求書url
 */
function spreadsheetToPdf(billSheetId, driveFolderId, year, month, name) {
  var sheet = SpreadsheetApp.openById(billSheetId).getActiveSheet();
  var spreadsheetId = billSheetId;
  var sheetId = sheet.getSheetId();
  SpreadsheetApp.flush();
  // PDFのエクスポートURL
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?';

  // PDFエクスポートオプション
  var exportOptions = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'A4',           // 用紙サイズ
    portrait: true,        // 縦向き
    fitw: true,            // ページ幅に合わせる
    gridlines: false,      // グリッド線を非表示
    printtitle: false,     // タイトルを非表示
    sheetnames: false,     // シート名を非表示
    fzr: false,            // 固定行を無効化
    top_margin: 0.5,       // 上部マージン
    bottom_margin: 0.5,    // 下部マージン
    left_margin: 0.5,      // 左マージン
    right_margin: 0.5,     // 右マージン
    horizontal_alignment: 'CENTER',  // 水平方向の中央揃え
    vertical_alignment: 'MIDDLE'     // 垂直方向の中央揃え
  }
  var params = [];
  for (var param in exportOptions) {
    params.push(param + '=' + exportOptions[param]);
  }
  var finalUrl = url + params.join('&') + '&gid=' + sheetId;
  // OAuthトークンの取得
  var token = ScriptApp.getOAuthToken();
  // PDFファイルを取得
  var response = UrlFetchApp.fetch(finalUrl, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  let nameOfPdf = `${year}年${month}月_晃商長期インターンシップ_${name}.pdf`;
  var blob = response.getBlob().setName(nameOfPdf);
  //同名ファイルが存在すれば削除
  try{
    DriveApp.getFolderById(driveFolderId).getFilesByName(nameOfPdf).next().setTrashed(true);
  }
  catch{}//同名ファイルなければエラー出して握りつぶす
  finally{
    // PDFファイルを指定されたドライブフォルダに保存
    var file = DriveApp.getFolderById(driveFolderId).createFile(blob);
    // 保存されたファイルのURLを返す
    return file.getUrl();
  }
}


//Aさんの勤務時間テスト
function testATime() {
  console.log(msecDictToHourAndMinutesDict(calculateMonthlyWorkTime("1BzGprz6Y8uJUgjuMpiry8qhD5ObbIlc-IWeGYp5pIWs",9,2024)));
}

//AさんgenerateBillテスト
function testABill(){
  generateBill("1BzGprz6Y8uJUgjuMpiry8qhD5ObbIlc-IWeGYp5pIWs",'128Te5rjJkyz0h5TPvyCXQ1wHJntO3EVo',9,2024);
}

function generate2MonthBill(attendanceSheetId,aggregationMonth,aggregationYear,thatYear = new Date().getFullYear(),thatMonth = new Date().getMonth() + 1,thatDate = new Date().getDate()) {
  //シートと変数の準備
  const databaseSheet = SpreadsheetApp.openById(databaseSheetId);
  const addressSheetId = databaseSheet.getSheets()[0].getRange(`E${search(databaseSheet.getSheets()[0],attendanceSheetId)[0]}`).getValue();
  const addressSheet = SpreadsheetApp.openById(addressSheetId);
  let slackUserId = databaseSheet.getSheets()[0].getRange(`C${search(databaseSheet.getSheets()[0],attendanceSheetId)[0]}`).getValue();
  let driveFolderId = databaseSheet.getSheets()[0].getRange(`D${search(databaseSheet.getSheets()[0],attendanceSheetId)[0]}`).getValue();
  let billUrl;
  let lastmonthlastday = new Date(aggregationYear,aggregationMonth-1,0);
  //インボイス請求書かを判断
  //インボイス番号はT（十三桁の半角数字）の形
  let bill;
  let invoinceNumber;
  const addressArray = addressSheet.getSheets()[0].getRange('A1:E7').getValues();
  const invoiceOrNot = addressArray[4][1];
  let status;
  try { 
    //インボイスでない場合
    if (invoiceOrNot === 'いいえ') {
      bill = SpreadsheetApp.openById(normalBillSheetId).getActiveSheet();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      const projectWorkTimesPrevious = calculateMonthlyWorkTime(attendanceSheetId,lastmonthlastday.getMonth()+1,lastmonthlastday.getFullYear())
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
      let hourlyWageDictPrevious = projectNameToHourlyWage(projectWorkTimesPrevious, attendanceSheetId);
      //請求書シートを初期化
      resetBill(normalBillSheetId);
      //請求書シートに計算結果を記述
      //通し番号更新
      addressSheet.getRange('B7').setValue(addressArray[6][1] + 1);
      const workTimeDict = msecDictToHourAndMinutesDict(hourlyWageDict);
      const workTimeDictPrevious = msecDictToHourAndMinutesDict(hourlyWageDictPrevious);
      let rowPointer = 15;//請求項目の書き始めは15行目

      for (let key in hourlyWageDictPrevious) {
        if (hourlyWageDictPrevious[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDictPrevious[key]/3600000);
          const setArray = [[`インターン${lastmonthlastday.getFullYear()}年${lastmonthlastday.getMonth()+1}月稼働分(${key}円×${workTimeDictPrevious[key]})`,'','',1,'式',price]];
          bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      for (let key in hourlyWageDict) {
        if (hourlyWageDict[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDict[key]/3600000);
          const setArray = [[`インターン${aggregationYear}年${aggregationMonth}月稼働分(${key}円×${workTimeDict[key]})`,'','',1,'式',price]];
          bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      //請求書シートに振込先を記述
      let deadline;
      if(thatDate > 10){
        deadline = new Date(thatYear,thatMonth + 1,0).toLocaleDateString('ja-JP')
      }else{
        deadline = new Date(thatYear,thatMonth ,0).toLocaleDateString('ja-JP')
      }
      const leftBlock = [['晃商インターン'],[deadline],[addressArray[0][1]],[`${addressArray[0][2]} ${addressArray[0][3]} ${addressArray[0][4]}`],[`${addressArray[2][3]} ${addressArray[2][4]}`]];
      bill.getRange('B6:B10').setValues(leftBlock);
      //請求書シートに通し番号等の右側ブロックを記述
      const rightBlock = [['No',addressArray[6][1] + 1],['請求日',new Date(thatYear,thatMonth - 1,thatDate).toLocaleDateString('ja-JP')],['',''],[`〒${addressArray[1][1]}`,''],[addressArray[1][2],''],[addressArray[1][3],''],[addressArray[2][1] + addressArray[2][2],''],[`TEL:${addressArray[3][1]}`,'']];
      bill.getRange('F2:G9').setValues(rightBlock);
      //請求書シートをpdf出力
      billUrl = spreadsheetToPdf(normalBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      //請求書シートを初期化
      resetBill(normalBillSheetId);
      status = 'success';
    }
    //インボイスの場合（プルダウンではい・いいえなのでいいえ以外）
    else {
      bill = SpreadsheetApp.openById(invoiceBillSheetId).getActiveSheet();
      invoinceNumber = addressSheet.getSheets()[0].getRange('B6').getValue();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      const projectWorkTimesPrevious = calculateMonthlyWorkTime(attendanceSheetId,lastmonthlastday.getMonth()+1,lastmonthlastday.getFullYear())
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
      let hourlyWageDictPrevious = projectNameToHourlyWage(projectWorkTimesPrevious, attendanceSheetId);
      //請求書シートを初期化
      resetBill(invoiceBillSheetId);
      //請求書シートに計算結果を記述
      //通し番号更新
      addressSheet.getRange('B7').setValue(addressArray[6][1] + 1);
      const workTimeDict = msecDictToHourAndMinutesDict(hourlyWageDict);
      const workTimeDictPrevious = msecDictToHourAndMinutesDict(hourlyWageDictPrevious);
      let rowPointer = 15;
      for (let key in hourlyWageDictPrevious) {
        if (hourlyWageDictPrevious[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDictPrevious[key]/3600000);
          const setArray = [[`インターン${lastmonthlastday.getFullYear()}年${lastmonthlastday.getMonth()+1}月稼働分(${key}円×${workTimeDictPrevious[key]})`,'','',1,'式',price,`${taxPercentage}%`]];
          bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      for (let key in hourlyWageDict) {
        if (hourlyWageDict[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDict[key]/3600000);
          const setArray = [[`インターン${aggregationYear}年${aggregationMonth}月稼働分(${key}円×${workTimeDict[key]})`,'','',1,'式',price,`${taxPercentage}%`]];
          bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      //請求書シートに振込先を記述
      let deadline;
      if(thatDate > 10){
        deadline = new Date(thatYear,thatMonth + 1,0).toLocaleDateString('ja-JP')
      }else{
        deadline = new Date(thatYear,thatMonth ,0).toLocaleDateString('ja-JP')
      }
      const leftBlock = [['晃商インターン'],[deadline],[addressArray[0][1]],[`${addressArray[0][2]} ${addressArray[0][3]} ${addressArray[0][4]}`],[`${addressArray[2][3]} ${addressArray[2][4]}`]];
      bill.getRange('B6:B10').setValues(leftBlock);
      //請求書シートに通し番号等の右側ブロックを記述
      const rightBlock = [['','No',addressArray[6][1] + 1],['','請求日',new Date(thatYear,thatMonth - 1,thatDate).toLocaleDateString('ja-JP')],['','',''],[`〒${addressArray[1][1]}`,'',''],[addressArray[1][2],'',''],[addressArray[1][3],'',''],[addressArray[2][1]+addressArray[2][2],'',''],[`TEL:${addressArray[3][1]}`,'',''],[`登録番号:T${addressArray[5][1]}`,'','']];
      bill.getRange('F2:H10').setValues(rightBlock);
      //請求書シートをpdf出力
      billUrl = spreadsheetToPdf(invoiceBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      //請求書シートを初期化
      resetBill(invoiceBillSheetId);
      status = 'success'
    }
  }
  catch(e){
    console.log(e)
    status = e.message;
  }
  //ManegementUtil接続用配列
  return {status:status,slackUserId:slackUserId,aggregationYear:aggregationYear,aggregationMonth:aggregationMonth,billUrl:billUrl}
}

const BillUtil = {
    get databaseSheetId() {
        return databaseSheetId;
    },

    calculateMonthlyWorkTime,
    generateBill,
    generate2MonthBill,
    search,
    msecDictToHourAndMinutesDict,
    projectNameToHourlyWage,
    resetBill,
    spreadsheetToPdf
};