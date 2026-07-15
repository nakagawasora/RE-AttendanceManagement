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

const USERS_COLUMNS = {
  id: 1,
  attendanceSheetId: 2,
  slackUserId: 3,
  driveFolderId: 4,
  addressSheetId: 5,
  email: 6,
  passwordHash: 7,
  savedData: 8,
  name: 9,
  deletedAt: 10
};

const PROJECT_COLUMNS = {
  id: 1,
  name: 2,
  hourlyWage: 3,
  active: 4,
  assignedUserIds: 5
};

const ROLE_ALLOWANCE_COLUMNS = {
  id: 1,
  name: 2,
  userId: 3,
  price: 4,
  deletedAt: 5
};
const ADMIN_EXECUTION_LOG_SHEET_NAME_FOR_BILL = 'AdminExecutionLogs';
const ADMIN_WHITELIST_SHEET_NAME_FOR_BILL = 'AdminWhitelist';
const DATABASE_META_SHEET_NAMES_FOR_BILL = [
  ADMIN_EXECUTION_LOG_SHEET_NAME_FOR_BILL,
  ADMIN_WHITELIST_SHEET_NAME_FOR_BILL
];
const BILL_DATABASE_CORE_SHEET_DEFINITIONS = [
  {
    name: 'Users',
    aliases: ['Users', 'User', 'ID', 'ユーザー', 'ユーザ', 'ユーザー一覧', 'Usersシート']
  },
  {
    name: 'Projects',
    aliases: ['Projects', 'Project', 'PJ', 'PJs', 'プロジェクト', 'プロジェクト一覧', 'Projectsシート']
  },
  {
    name: 'WageUpgrades',
    aliases: ['WageUpgrades', 'WageUpgrade', 'HourlyWageIncrements', 'upgrade', '時給増分', '時給アップ', '時給アップシート']
  },
  {
    name: 'RoleAllowances',
    aliases: ['RoleAllowances', 'RoleAllowance', '役職手当', '役職手当シート']
  }
];

function getDatabaseSheetsForBill() {
  const database = SpreadsheetApp.openById(databaseSheetId);
  const fallbackSheets = getOrderedNonMetaDatabaseSheetsForBill(database);
  return BILL_DATABASE_CORE_SHEET_DEFINITIONS.map((definition, index) => {
    const sheet = findDatabaseSheetByDefinitionForBill(database, definition) || fallbackSheets[index];
    if (!sheet) {
      throw new Error(`${definition.name}シートが見つかりません`);
    }
    return sheet;
  });
}

function getDatabaseSheetForBill(index) {
  const sheet = getDatabaseSheetsForBill()[index];
  if (!sheet) {
    throw new Error(`データベースシート${index + 1}が見つかりません`);
  }
  return sheet;
}

function getOrderedNonMetaDatabaseSheetsForBill(database) {
  return database.getSheets()
    .filter(sheet => !DATABASE_META_SHEET_NAMES_FOR_BILL.includes(sheet.getName()));
}

function normalizeSheetNameForLookupForBill(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, '');
}

function findDatabaseSheetByDefinitionForBill(database, definition) {
  const aliases = definition.aliases.map(normalizeSheetNameForLookupForBill);
  return database.getSheets().find(sheet => aliases.includes(normalizeSheetNameForLookupForBill(sheet.getName()))) || null;
}

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
  const user = getUserRowByAttendanceSheetId(attendanceSheetId);
  if (user.deletedAt) {
    return {status:'論理削除済みユーザーの請求書は発行できません',slackUserId:user.slackUserId,aggregationYear:aggregationYear,aggregationMonth:aggregationMonth,billUrl:''};
  }
  const addressSheetId = user.addressSheetId;
  const addressSheet = SpreadsheetApp.openById(addressSheetId);
  let slackUserId = user.slackUserId;
  let billUrl;
  //インボイス請求書かを判断
  //インボイス番号はT（十三桁の半角数字）の形
  let bill;
  let invoinceNumber;
  const addressArray = addressSheet.getSheets()[0].getRange('A1:E7').getValues();
  const invoiceOrNot = addressArray[4][1];
  let status;
  let workingBillSheetId;
  try { 
    //インボイスでない場合
    if (invoiceOrNot === 'いいえ') {
      workingBillSheetId = createTemporaryBillSpreadsheet(normalBillSheetId, aggregationYear, aggregationMonth, addressArray[2][1] + addressArray[2][2]);
      bill = SpreadsheetApp.openById(workingBillSheetId).getActiveSheet();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
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
      const roleAllowanceItems = calculateRoleAllowanceItems(attendanceSheetId, aggregationMonth, aggregationYear, projectWorkTimes);
      roleAllowanceItems.forEach(item => {
        const setArray = [[`${item.name}(${item.price}円×${item.workTimeText})`,'','',1,'式',item.amount]];
        bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
        rowPointer++;
      });
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
      billUrl = spreadsheetToPdf(workingBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      status = 'success';
    }
    //インボイスの場合（プルダウンではい・いいえなのでいいえ以外）
    else {
      workingBillSheetId = createTemporaryBillSpreadsheet(invoiceBillSheetId, aggregationYear, aggregationMonth, addressArray[2][1] + addressArray[2][2]);
      bill = SpreadsheetApp.openById(workingBillSheetId).getActiveSheet();
      invoinceNumber = addressSheet.getSheets()[0].getRange('B6').getValue();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
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
      const roleAllowanceItems = calculateRoleAllowanceItems(attendanceSheetId, aggregationMonth, aggregationYear, projectWorkTimes);
      roleAllowanceItems.forEach(item => {
        const setArray = [[`${item.name}(${item.price}円×${item.workTimeText})`,'','',1,'式',item.amount,`${taxPercentage}%`]];
        bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
        rowPointer++;
      });
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
      billUrl = spreadsheetToPdf(workingBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      status = 'success'
    }
  }
  catch(e){
    status = e.message;
  }
  finally {
    trashSpreadsheetFile(workingBillSheetId);
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

function msecToHourAndMinutes(msec) {
  return `${Math.trunc(msec/3600000)}時間${Math.trunc((msec%3600000)/60000)}分`;
}

function calculateTotalWorkTime(projectWorkTimes) {
  let total = 0;
  for (let key in projectWorkTimes) {
    total += Number(projectWorkTimes[key] || 0);
  }
  return total;
}

function getUserIdByAttendanceSheetId(attendanceSheetId) {
  const user = getUserRowByAttendanceSheetId(attendanceSheetId);
  if (user.deletedAt) {
    throw new Error('論理削除済みユーザーは対象外です');
  }
  return user.id;
}

function getUserRowByAttendanceSheetId(attendanceSheetId) {
  const userSheet = getDatabaseSheetForBill(0);
  const [userRow] = search(userSheet, attendanceSheetId);
  if (userRow === -1) {
    throw new Error(`出席シートID「${attendanceSheetId}」に対応するユーザーが見つかりません`);
  }
  const row = userSheet.getRange(userRow, 1, 1, 10).getValues()[0];
  return {
    id: row[USERS_COLUMNS.id - 1],
    attendanceSheetId: row[USERS_COLUMNS.attendanceSheetId - 1],
    slackUserId: row[USERS_COLUMNS.slackUserId - 1],
    driveFolderId: row[USERS_COLUMNS.driveFolderId - 1],
    addressSheetId: row[USERS_COLUMNS.addressSheetId - 1],
    email: row[USERS_COLUMNS.email - 1],
    name: row[USERS_COLUMNS.name - 1],
    deletedAt: row[USERS_COLUMNS.deletedAt - 1]
  };
}

function getRoleAllowancesForUser(userId) {
  const sheets = getDatabaseSheetsForBill();
  const roleAllowanceSheet = sheets[3];
  if (!roleAllowanceSheet || roleAllowanceSheet.getLastRow() < 2) {
    return [];
  }
  const values = roleAllowanceSheet.getRange(2, 1, roleAllowanceSheet.getLastRow() - 1, 5).getValues();
  return values
    .filter(row => String(row[ROLE_ALLOWANCE_COLUMNS.userId - 1]) === String(userId))
    .filter(row => !row[ROLE_ALLOWANCE_COLUMNS.deletedAt - 1])
    .map(row => ({
      id: row[ROLE_ALLOWANCE_COLUMNS.id - 1],
      name: row[ROLE_ALLOWANCE_COLUMNS.name - 1],
      userId: row[ROLE_ALLOWANCE_COLUMNS.userId - 1],
      price: Number(row[ROLE_ALLOWANCE_COLUMNS.price - 1] || 0)
    }))
    .filter(allowance => allowance.name && allowance.price !== 0);
}

function calculateRoleAllowanceItems(attendanceSheetId, aggregationMonth, aggregationYear, precomputedProjectWorkTimes) {
  const userId = getUserIdByAttendanceSheetId(attendanceSheetId);
  const projectWorkTimes = precomputedProjectWorkTimes || calculateMonthlyWorkTime(attendanceSheetId, aggregationMonth, aggregationYear);
  const totalWorkTime = calculateTotalWorkTime(projectWorkTimes);
  const totalWorkTimeText = msecToHourAndMinutes(totalWorkTime);
  return getRoleAllowancesForUser(userId).map(allowance => ({
    id: allowance.id,
    name: allowance.name,
    userId: allowance.userId,
    price: allowance.price,
    workTimeMsec: totalWorkTime,
    workTimeText: totalWorkTimeText,
    amount: Math.round(allowance.price * totalWorkTime / 3600000)
  })).filter(item => item.workTimeMsec > 0);
}

function calculateMonthlyWorkItems(attendanceSheetId, aggregationMonth, aggregationYear, precomputedProjectWorkTimes) {
  const database = SpreadsheetApp.openById(databaseSheetId);
  const projectSheet = getDatabaseSheetForBill(1);
  const upgradeSheet = getDatabaseSheetForBill(2);
  const userId = getUserIdByAttendanceSheetId(attendanceSheetId);
  const projectWorkTimes = precomputedProjectWorkTimes || calculateMonthlyWorkTime(attendanceSheetId, aggregationMonth, aggregationYear);
  const projectList = projectSheet.getLastRow() < 2 ? [] : projectSheet.getRange(2, 1, projectSheet.getLastRow() - 1, 5).getValues();
  const upgradeList = upgradeSheet.getLastRow() < 2 ? [] : upgradeSheet.getRange('A2:D' + upgradeSheet.getLastRow()).getValues();
  const upgradeMap = {};
  for (let i = 0; i < upgradeList.length; i++) {
    const [id, user_id, project_id, increment] = upgradeList[i];
    upgradeMap[`${user_id}_${project_id}`] = Number(increment || 0);
  }

  return projectList.map(project => {
    const projectId = project[PROJECT_COLUMNS.id - 1];
    const projectName = project[PROJECT_COLUMNS.name - 1];
    const workTimeMsec = Number(projectWorkTimes[projectName] || 0);
    const unitPrice = Number(project[PROJECT_COLUMNS.hourlyWage - 1] || 0) + Number(upgradeMap[`${userId}_${projectId}`] || 0);
    return {
      type: 'project',
      id: projectId,
      name: projectName,
      unitPrice: unitPrice,
      workTimeMsec: workTimeMsec,
      workTimeText: msecToHourAndMinutes(workTimeMsec),
      amount: Math.round(unitPrice * workTimeMsec / 3600000)
    };
  }).filter(item => item.workTimeMsec > 0);
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
    const userSheet = getDatabaseSheetForBill(0);     // Usersシート
    const projectSheet = getDatabaseSheetForBill(1);  // Projectsシート
    const upgradeSheet = getDatabaseSheetForBill(2);  // 時給アップシート
    
    const userId = getUserIdByAttendanceSheetId(attendanceSheetId);
    
    // プロジェクト情報を取得（A列:id, B列:プロジェクト名, C列:時給）
    const projectList = projectSheet.getLastRow() < 2 ? [] : projectSheet.getRange(2, 1, projectSheet.getLastRow() - 1, 5).getValues();
    
    // 時給アップ情報を取得
    const upgradeList = upgradeSheet.getLastRow() < 2 ? [] : upgradeSheet.getRange('A2:D' + upgradeSheet.getLastRow()).getValues();
    
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
        const projectId = projectList[i][PROJECT_COLUMNS.id - 1];
        const projectName = projectList[i][PROJECT_COLUMNS.name - 1];
        let projectHourlyWage = Number(projectList[i][PROJECT_COLUMNS.hourlyWage - 1] || 0);
        
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

function createTemporaryBillSpreadsheet(templateSheetId, year, month, name) {
  const fileName = `tmp_${year}年${month}月_請求書_${name}_${Utilities.getUuid()}`;
  return DriveApp.getFileById(templateSheetId).makeCopy(fileName).getId();
}

function trashSpreadsheetFile(spreadsheetId) {
  if (!spreadsheetId) {
    return;
  }
  try {
    DriveApp.getFileById(spreadsheetId).setTrashed(true);
  } catch (e) {
    console.log(`一時請求書ファイルの削除に失敗しました: ${e.message}`);
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
  const user = getUserRowByAttendanceSheetId(attendanceSheetId);
  if (user.deletedAt) {
    return {status:'論理削除済みユーザーの請求書は発行できません',slackUserId:user.slackUserId,aggregationYear:aggregationYear,aggregationMonth:aggregationMonth,billUrl:''};
  }
  const addressSheetId = user.addressSheetId;
  const addressSheet = SpreadsheetApp.openById(addressSheetId);
  let slackUserId = user.slackUserId;
  let driveFolderId = user.driveFolderId;
  let billUrl;
  let lastmonthlastday = new Date(aggregationYear,aggregationMonth-1,0);
  //インボイス請求書かを判断
  //インボイス番号はT（十三桁の半角数字）の形
  let bill;
  let invoinceNumber;
  const addressArray = addressSheet.getSheets()[0].getRange('A1:E7').getValues();
  const invoiceOrNot = addressArray[4][1];
  let status;
  let workingBillSheetId;
  try { 
    //インボイスでない場合
    if (invoiceOrNot === 'いいえ') {
      workingBillSheetId = createTemporaryBillSpreadsheet(normalBillSheetId, aggregationYear, aggregationMonth, addressArray[2][1] + addressArray[2][2]);
      bill = SpreadsheetApp.openById(workingBillSheetId).getActiveSheet();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      const projectWorkTimesPrevious = calculateMonthlyWorkTime(attendanceSheetId,lastmonthlastday.getMonth()+1,lastmonthlastday.getFullYear())
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
      let hourlyWageDictPrevious = projectNameToHourlyWage(projectWorkTimesPrevious, attendanceSheetId);
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
      const roleAllowanceItemsPrevious = calculateRoleAllowanceItems(attendanceSheetId, lastmonthlastday.getMonth()+1, lastmonthlastday.getFullYear(), projectWorkTimesPrevious);
      roleAllowanceItemsPrevious.forEach(item => {
        const setArray = [[`${lastmonthlastday.getFullYear()}年${lastmonthlastday.getMonth()+1}月 ${item.name}(${item.price}円×${item.workTimeText})`,'','',1,'式',item.amount]];
        bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
        rowPointer++;
      });
      for (let key in hourlyWageDict) {
        if (hourlyWageDict[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDict[key]/3600000);
          const setArray = [[`インターン${aggregationYear}年${aggregationMonth}月稼働分(${key}円×${workTimeDict[key]})`,'','',1,'式',price]];
          bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      const roleAllowanceItems = calculateRoleAllowanceItems(attendanceSheetId, aggregationMonth, aggregationYear, projectWorkTimes);
      roleAllowanceItems.forEach(item => {
        const setArray = [[`${aggregationYear}年${aggregationMonth}月 ${item.name}(${item.price}円×${item.workTimeText})`,'','',1,'式',item.amount]];
        bill.getRange(`A${rowPointer}:F${rowPointer}`).setValues(setArray);
        rowPointer++;
      });
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
      billUrl = spreadsheetToPdf(workingBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      status = 'success';
    }
    //インボイスの場合（プルダウンではい・いいえなのでいいえ以外）
    else {
      workingBillSheetId = createTemporaryBillSpreadsheet(invoiceBillSheetId, aggregationYear, aggregationMonth, addressArray[2][1] + addressArray[2][2]);
      bill = SpreadsheetApp.openById(workingBillSheetId).getActiveSheet();
      invoinceNumber = addressSheet.getSheets()[0].getRange('B6').getValue();
      //勤務時間計算
      const projectWorkTimes = calculateMonthlyWorkTime(attendanceSheetId,aggregationMonth,aggregationYear);
      const projectWorkTimesPrevious = calculateMonthlyWorkTime(attendanceSheetId,lastmonthlastday.getMonth()+1,lastmonthlastday.getFullYear())
      //プロジェクト名から時給ごとにわける
      let hourlyWageDict = projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
      let hourlyWageDictPrevious = projectNameToHourlyWage(projectWorkTimesPrevious, attendanceSheetId);
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
      const roleAllowanceItemsPrevious = calculateRoleAllowanceItems(attendanceSheetId, lastmonthlastday.getMonth()+1, lastmonthlastday.getFullYear(), projectWorkTimesPrevious);
      roleAllowanceItemsPrevious.forEach(item => {
        const setArray = [[`${lastmonthlastday.getFullYear()}年${lastmonthlastday.getMonth()+1}月 ${item.name}(${item.price}円×${item.workTimeText})`,'','',1,'式',item.amount,`${taxPercentage}%`]];
        bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
        rowPointer++;
      });
      for (let key in hourlyWageDict) {
        if (hourlyWageDict[key] !== 0) {
          const price = Math.round(Number(key)*hourlyWageDict[key]/3600000);
          const setArray = [[`インターン${aggregationYear}年${aggregationMonth}月稼働分(${key}円×${workTimeDict[key]})`,'','',1,'式',price,`${taxPercentage}%`]];
          bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
          rowPointer++;
        }
      }
      const roleAllowanceItems = calculateRoleAllowanceItems(attendanceSheetId, aggregationMonth, aggregationYear, projectWorkTimes);
      roleAllowanceItems.forEach(item => {
        const setArray = [[`${aggregationYear}年${aggregationMonth}月 ${item.name}(${item.price}円×${item.workTimeText})`,'','',1,'式',item.amount,`${taxPercentage}%`]];
        bill.getRange(`A${rowPointer}:G${rowPointer}`).setValues(setArray);
        rowPointer++;
      });
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
      billUrl = spreadsheetToPdf(workingBillSheetId,driveFolderId,aggregationYear,aggregationMonth,addressArray[2][1] + addressArray[2][2]);
      status = 'success'
    }
  }
  catch(e){
    console.log(e)
    status = e.message;
  }
  finally {
    trashSpreadsheetFile(workingBillSheetId);
  }
  //ManegementUtil接続用配列
  return {status:status,slackUserId:slackUserId,aggregationYear:aggregationYear,aggregationMonth:aggregationMonth,billUrl:billUrl}
}

const BillUtil = {
    get databaseSheetId() {
        return databaseSheetId;
    },
    set databaseSheetId(nextDatabaseSheetId) {
        databaseSheetId = nextDatabaseSheetId;
    },

    calculateMonthlyWorkTime,
    generateBill,
    generate2MonthBill,
    search,
    msecDictToHourAndMinutesDict,
    msecToHourAndMinutes,
    calculateTotalWorkTime,
    calculateMonthlyWorkItems,
    calculateRoleAllowanceItems,
    getRoleAllowancesForUser,
    getUserIdByAttendanceSheetId,
    projectNameToHourlyWage,
    resetBill,
    spreadsheetToPdf
};
