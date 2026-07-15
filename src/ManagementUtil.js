const scriptProperties = PropertiesService.getScriptProperties();

const token = scriptProperties.getProperty("SLACK_BOT_TOKEN");
const billChannelId = scriptProperties.getProperty("BILL_CHANNEL_ID");
const slackIdOfAdmin = scriptProperties.getProperty("SLACK_ID_OF_ADMIN");
const mailAddress = scriptProperties.getProperty("MAIL_ADDRESS");
const ccRecipient = scriptProperties.getProperty("CC_RECIPIENT");
const billFolderId = scriptProperties.getProperty("BILL_FOLDER_ID");
const apiToken = scriptProperties.getProperty("API_TOKEN");

const USER_COL = {
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

const PROJECT_COL = {
  id: 1,
  name: 2,
  hourlyWage: 3,
  active: 4,
  assignedUserIds: 5
};
const NO_ASSIGNED_USERS = '__NONE__';
const ADMIN_EXECUTION_LOG_SHEET_NAME = 'AdminExecutionLogs';
const ADMIN_WHITELIST_SHEET_NAME = 'AdminWhitelist';
const DATABASE_META_SHEET_NAMES = [ADMIN_EXECUTION_LOG_SHEET_NAME, ADMIN_WHITELIST_SHEET_NAME];

function openDatabaseSpreadsheet() {
  return SpreadsheetApp.openById(BillUtil.databaseSheetId);
}

function getDatabaseSheets() {
  return openDatabaseSpreadsheet().getSheets()
    .filter(sheet => !DATABASE_META_SHEET_NAMES.includes(sheet.getName()));
}

function getDatabaseSheet(index) {
  const sheet = getDatabaseSheets()[index];
  if (!sheet) {
    throw new Error(`データベースシート${index + 1}が見つかりません`);
  }
  return sheet;
}

function repairDatabaseSheetOrder() {
  const database = openDatabaseSpreadsheet();
  const adminWhitelistSheet = database.getSheetByName(ADMIN_WHITELIST_SHEET_NAME);
  if (adminWhitelistSheet) {
    database.setActiveSheet(adminWhitelistSheet);
    database.moveActiveSheet(database.getNumSheets());
  }
  const logSheet = database.getSheetByName(ADMIN_EXECUTION_LOG_SHEET_NAME);
  if (logSheet) {
    database.setActiveSheet(logSheet);
    database.moveActiveSheet(database.getNumSheets());
  }
}

/**
 * 先月の請求書を設置してスラックで通知
 * @param {number} year - 実行年（デバッグ用）
 * @param {number} month - 実行月（デバッグ用）
 * @param {number} date - 実行日（デバッグ用）
*/
function setBill(year = new Date().getFullYear(),month = new Date().getMonth() + 1,date = new Date().getDate()) {
  const lastMonthLastDay = new Date(year,month - 1,0);//-1により見かけ上の集計月からDate型の集計月へ
  const databaseSheet = getDatabaseSheet(0);
  const idArray = databaseSheet.getRange('B2:J'+databaseSheet.getLastRow()).getValues();
  let genarateCounter = 0;
  let targetCounter = 0;
  for (let i = 0; i < idArray.length; i++){
    if (idArray[i][8]) {
      continue;
    }
    targetCounter++;
    let generatedDict;
    //請求書出力
    try{
      generatedDict = BillUtil.generateBill(idArray[i][0],idArray[i][2],lastMonthLastDay.getMonth()+1,lastMonthLastDay.getFullYear(),year,month,date);
    }
    catch(e){
      generatedDict = {status:e.message,slackUserId:idArray[i][1],aggregationYear:lastMonthLastDay.getFullYear(),aggregationMonth:lastMonthLastDay.getMonth()+1,billUrl:''};
    }
    //成功数のカウント
    if(generatedDict.status === 'success'){
      genarateCounter++;
    }
    //結果をスラックで連絡、送信エラーが出ればメールで連絡
    try{
      slackReport(generatedDict);
    }
    catch(e){
      try{
        let name = SpreadsheetApp.openById(idArray[i][3]).getSheets()[0].getRange('B3:C3').getValues();
        let fullname = name[0][0]+name[0][1];
        GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${fullname}さんの請求書発行に際し、Slackでのメッセージ送信に失敗しました。確認してください。\nエラーメッセージ:${e.message}\nなお、請求書発行ステータスは以下の通りです（エラーが発生している場合は再実行、成功している場合は対象者と連絡を取って訂正の必要がないか確認してください）。\n\n${generatedDict.status}`,{ name: 'ManagementUtil',cc:ccRecipient });
      }
      catch{
        GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${i + 1}番目の請求書発行に際し、Slackでのメッセージ送信に失敗しました。また、本名の取得にも失敗しているため対象者のAddressSheetIDが正しいかどうか確認してください。\nエラーメッセージ:${e.message}\nなお、請求書発行ステータスは以下の通りです（エラーが発生している場合は再実行、成功している場合は対象者と連絡を取って訂正の必要がないか確認してください）。\n\n${generatedDict.status}`,{ name: 'ManagementUtil',cc:ccRecipient }); 
      }
    }
  }
  //おまとめメッセージの送信、送信エラーが出ればメールに送るべきメッセージを連絡
  try{
    let message = `<!channel>\n${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月分の請求書が出力されました。\n成功:${genarateCounter}件\n失敗:${targetCounter - genarateCounter}件\n訂正の必要がなければ当メッセージにリアクションを、何か問題があれば返信してください。`;
    sendMessage(message,billChannelId);
  }
  catch(e){
    GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月分の請求書発行おまとめメッセージ送信に際し、Slackでのメッセージ送信に失敗しました。\nエラーメッセージ:${e.message}\n以下の文章かぎ括弧内をコピーして請求書チャンネルに貼り付けてください。\n「@channel\n${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月分の請求書が出力されました。\n成功:${genarateCounter}件\n失敗:${targetCounter - genarateCounter}件\n訂正の必要がなければ当メッセージにリアクションを、何か問題があれば返信してください。」`,{cc:ccRecipient});   
  }
}

/**
 * 特定の勤怠シートについて先月の請求書を再発行して通知（各シートからインポートしてメニューから実行する用）
 * @param {string} attendanceSheetId - 勤怠シートのID
 * @param {number} year - 実行年（デバッグ用）
 * @param {number} month - 実行月（デバッグ用）
 * @param {number} date - 実行日（デバッグ用）
*/
function reSetBill(attendanceSheetId, year = new Date().getFullYear(),month = new Date().getMonth() + 1,date = new Date().getDate()){
  const lastMonthLastDay = new Date(year,month - 1,0);//-1により見かけ上の集計月からDate型の集計月へ
  const databaseSheet = getDatabaseSheet(0);
  const row = BillUtil.search(databaseSheet,attendanceSheetId)[0];
  if (row === -1) {
    throw new Error('対象ユーザーが見つかりません');
  }
  const user = databaseSheet.getRange(row, 1, 1, 10).getValues()[0];
  if (user[USER_COL.deletedAt - 1]) {
    throw new Error('論理削除済みユーザーの請求書は再発行できません');
  }
  let generatedDict;
  //請求書出力
  try{
    generatedDict = BillUtil.generateBill(user[USER_COL.attendanceSheetId - 1],user[USER_COL.driveFolderId - 1],lastMonthLastDay.getMonth()+1,lastMonthLastDay.getFullYear(),year,month,date);
  }
  catch(e){
    generatedDict = {status:e.message,slackUserId:user[USER_COL.slackUserId - 1],aggregationYear:lastMonthLastDay.getFullYear(),aggregationMonth:lastMonthLastDay.getMonth()+1,billUrl:''};
  }
  //結果をスラックで連絡、送信エラーが出ればメールで連絡
  try{
    slackReport(generatedDict,true);
  }
  catch(e){
    try{
      let name = SpreadsheetApp.openById(user[USER_COL.addressSheetId - 1]).getSheets()[0].getRange('B3:C3').getValues();
      let fullname = name[0][0]+name[0][1];
      GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${fullname}さんの請求書再発行に際し、Slackでのメッセージ送信に失敗しました。確認してください\nエラーメッセージ:${e.message}\nなお、請求書発行ステータスは以下の通りです（エラーが発生している場合は再実行、成功している場合は対象者と連絡を取って再訂正の必要がないか確認してください）。\n\n${generatedDict.status}`,{ name: 'ManagementUtil' ,cc:ccRecipient});
    }
    catch{
      GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`請求書再発行に際し、Slackでのメッセージ送信に失敗しました。また、本名の取得にも失敗しているため対象者のAddressSheetIDが正しいかどうか確認してください。\nエラーメッセージ:${e.message}\nなお、請求書発行ステータスは以下の通りです（エラーが発生している場合は再実行、成功している場合は対象者と連絡を取って再訂正の必要がないか確認してください）。\n\n${generatedDict.status}`,{ name: 'ManagementUtil',cc:ccRecipient }); 
    }
  }
}

/**
 * 請求書出力を受けてスラックにメッセージを送信
 * @param {Object.<string,string|number>} dict - BillUtil.generateBillによる連想配列
 * @param {Boolean} isReset - 再発行フラグ
*/
function slackReport(dict,isReset = false) {
  // setBill用
  if(isReset === false){
    //成功メッセージ
    if (dict.status === 'success'){
      let message = `${dict.aggregationYear}年${dict.aggregationMonth}月分請求書が発行されました。\n下記URLから内容を確認し、訂正の必要がなければ請求書チャンネルのおまとめメッセージにリアクションを、何か問題があれば返信してください。\n（当bot及びDMは送信専用です）\n${dict.billUrl} `;
      sendDirectmessage(message,dict.slackUserId);
    }//失敗メッセージ
    else {
      //メンションミスでエラーを発生させる
      if(!isUserInChannel(billChannelId,dict.slackUserId)){
        throw new Error(' slackReport error;user is not found');
      }
      let message = `${mention(slackIdOfAdmin)}様\n${mention(dict.slackUserId)}さんの${dict.aggregationYear}年${dict.aggregationMonth}月分請求書の発行に失敗しました。\nしばらく時間をおいて再試行してください。\n再試行しても問題が解決されない場合は勤怠管理システム担当者に連絡してください。\nエラーメッセージ:${dict.status}`;
      sendMessage(message,billChannelId);
    }
  }//reSetBill用
  else{
    //成功メッセージ
    if (dict.status === 'success'){
      let message = `${dict.aggregationYear}年${dict.aggregationMonth}月分請求書が発行されました。\n下記URLから内容を確認し、再訂正の必要があるかどうかを管理者に連絡してください。\n（当bot及びDMは送信専用です）\n${dict.billUrl} `;
      sendDirectmessage(message,dict.slackUserId);
      message = `${dict.aggregationYear}年${dict.aggregationMonth}月分請求書が再発行されました。\n対象者と連絡を取り、再訂正の必要があるかを確認してください。\n（当bot及びDMは送信専用です）`;
      sendDirectmessage(message,slackIdOfAdmin);
    }//失敗メッセージ
    else {
      //メンションミスでエラーを発生させる
      if(!isUserInChannel(billChannelId,dict.slackUserId)){
        throw new Error(' slackReport error;user is not found');
      }
      let message = `${mention(slackIdOfAdmin)}様\n${mention(dict.slackUserId)}さんの${dict.aggregationYear}年${dict.aggregationMonth}月分請求書の再発行に失敗しました。\nしばらく時間をおいて再試行してください。\n再試行しても問題が解決されない場合は勤怠管理システム担当者に連絡してください。\nエラーメッセージ:${dict.status}`;
      sendMessage(message,billChannelId);
    }
  }
}

/**
 * 指定のチャンネルに指定のメッセージを送信
 * @param {string} message - メッセージ内容
 * @param {string} channelId - 送信先チャンネルID
*/
function sendMessage(message = '<!channel> \nHello world!',channelId = billChannelId) {
  const response = SlackApp.create(token).postMessage(channelId, message);
  //送信失敗でエラーを発生させる
  if(response.ok === false){
    throw new Error('sendMessage error;response.ok = false');
  }
}

/**
 * useridをメンションの形に
 * @param {string} userId - スラックID
 * @returns {string} - メンションの形
*/
function mention(userId) {
  return `<@${userId}>`;
}

/**
 * 実行前月の請求書を確認用から回収フォルダにコピー（なければ作成）
 * @param {number} year - 実行年（デバッグ用）
 * @param {number} month - 実行月（デバッグ用）
*/
function gatherBills(year = new Date().getFullYear(),month = new Date().getMonth() + 1) {
  const billFolder = DriveApp.getFolderById(billFolderId);
  //当月フォルダがあればID取得、なければ作成
  const lastMonthLastDay = new Date(year,month - 1,0);

  const folderName = `${lastMonthLastDay.getFullYear()}${(lastMonthLastDay.getMonth()+1).toString().padStart(2,'0')}請求書`;
  let monthlyBillFolderId;
  try{
    monthlyBillFolderId = billFolder.getFoldersByName(folderName).next().getId();
  } 
  catch{
    monthlyBillFolderId = billFolder.createFolder(folderName).getId();
  } 
  //繰り返し収集
  const databaseSheet = getDatabaseSheet(0);
  const idArray = databaseSheet.getRange('B2:J'+databaseSheet.getLastRow()).getValues();
  let failureCounter = 0;
  let targetCounter = 0;
  for(let i = 0; i < idArray.length; i++){
    if (idArray[i][8]) {
      continue;
    }
    targetCounter++;
    let fileName = `${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月_晃商長期インターンシップ_${SpreadsheetApp.openById(idArray[i][3]).getRange('B3').getValue()}${SpreadsheetApp.openById(idArray[i][3]).getRange('C3').getValue()}.pdf`;
    let status = sendBill(idArray[i][2],monthlyBillFolderId,fileName);

    //回収失敗のときは個別的にメッセージ
    if(status !== 'success'){
      failureCounter++;
      try{
        sendMessage(`${mention(slackIdOfAdmin)}様 \n${SpreadsheetApp.openById(idArray[i][3]).getRange('B3').getValue()}${SpreadsheetApp.openById(idArray[i][3]).getRange('C3').getValue()}さんの請求書回収に際し、エラーが発生しました。\nしばらく時間をおいて再試行してください。\n再試行しても問題が解決されない場合は勤怠管理システム担当者に連絡してください。\nエラーメッセージ:${status}`,billChannelId);
      }
      catch(e){
        try{
          GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${SpreadsheetApp.openById(idArray[i][3]).getRange('B3').getValue()}${SpreadsheetApp.openById(idArray[i][3]).getRange('C3').getValue()}さんの請求書回収に際して発生したエラーを伝えるSlackでのメッセージ送信に失敗しました。\nしばらく時間をおいて再試行してください。\n再試行しても問題が解決されない場合は勤怠管理システム担当者に連絡してください。\nエラーメッセージ:${e.message} \nなお、請求書回収に際して発生したエラーは以下の通りです。\n\n${status}`,{ name: 'ManagementUtil',cc:ccRecipient }); 
        }
        catch{
          GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${i + 1}番目の請求書回収に際し、Slackでのメッセージ送信に失敗しました。また、本名の取得にも失敗しているため対象者のAddressSheetIDが正しいかどうか確認してください。\nエラーメッセージ:${e.message}\nなお、請求書回収に際して発生したエラーは以下の通りです。\n\n${status}`,{ name: 'ManagementUtil',cc:ccRecipient }); 
        }
      }
    }
  }
  //すべての回収に成功したときのメッセージ
  if(failureCounter === 0){
    try{
      sendMessage(`${mention(slackIdOfAdmin)}様\nすべての請求書回収に成功しました。\n請求書フォルダを確認してください。`,billChannelId);
    }
    catch(e){
      GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`すべての請求書回収に成功したことを伝えるSlackでのメッセージ送信に失敗しました。回収自体は完了しているため、請求書フォルダを確認してください。\nエラーメッセージ:${e.message}`,{ name: 'ManagementUtil',cc:ccRecipient });
    }
  }//一部失敗時のメッセージ
  else{
    try{
      sendMessage(`${targetCounter - failureCounter}件の請求書回収に成功し、${failureCounter}件の請求書回収に失敗しました。\n請求書フォルダを確認してください。`,billChannelId);
    }
    catch(e){
      GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${targetCounter - failureCounter}件の請求書回収に成功し、${failureCounter}件の請求書回収に失敗したことを伝えるSlackでのメッセージ送信に失敗しました。請求書フォルダを確認してください。\nエラーメッセージ:${e.message}`,{ name: 'ManagementUtil',cc:ccRecipient });
    }
  }
}

/**
 * 上書き保存関数
 * @param {string} sourceFolderId - コピー元フォルダID
 * @param {string} destinationFolderId - コピー先フォルダID
 * @param {string} targetFileName - コピーファイル名
 * @returns {string} - エラーステータス
*/
function sendBill(sourceFolderId,destinationFolderId,targetFileName) {
  //対象のフォルダからファイル名で選択
  let status;
  try{
    let files = DriveApp.getFolderById(sourceFolderId).getFiles();
    let file;
    while(files.hasNext()){
      let checking = files.next();
      if(checking.getName() === targetFileName){
        file = checking;
        break;
      }
    }
    //コピーを上書き保存
    try{
      DriveApp.getFolderById(destinationFolderId).getFilesByName(file.getName()).next().setTrashed(true);
    }
    catch{}//同名ファイルなければエラー出して握りつぶす
    file.makeCopy(file.getName(),DriveApp.getFolderById(destinationFolderId));
    status = 'success';
  }
  catch(e){
    status = e.message;
  }
  return status;
}

/**
 * 指定のユーザーに指定のDMを送信
 * @param {string} message - メッセージ内容
 * @param {string} userId - 送信先ユーザーID
*/
function sendDirectmessage(message = 'hello world!',userId = 'U0764N7CDC2'){
  const slack = SlackApp.create(token);
  const endpoint = 'https://slack.com/api/conversations.open';
  const options = {
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      users: userId, // メッセージを送りたい相手のユーザーID
    }),
  };
  const openResponse = UrlFetchApp.fetch(endpoint, options);
  const openResult = JSON.parse(openResponse.getContentText());
  const dmChannelId = openResult.channel.id;
  const response = slack.postMessage(dmChannelId,message);
  //送信失敗でエラーを発生させる
  if(response.ok === false){
    throw new Error('sendMessage error;response.ok = false');
  }
}

function sendDirectmessages(message, userIds) {
  const uniqueUserIds = Array.from(new Set((userIds || []).map(userId => String(userId || '').trim()).filter(Boolean)));
  if (!uniqueUserIds.length) {
    return;
  }
  const openRequests = uniqueUserIds.map(userId => ({
    url: 'https://slack.com/api/conversations.open',
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ users: userId }),
    muteHttpExceptions: true
  }));
  const openResponses = UrlFetchApp.fetchAll(openRequests);
  const postRequests = openResponses.map((response, index) => {
    const result = JSON.parse(response.getContentText());
    if (!result.ok || !result.channel || !result.channel.id) {
      throw new Error(`Slack DM開始に失敗しました: ${uniqueUserIds[index]} ${result.error || ''}`);
    }
    return {
      url: 'https://slack.com/api/chat.postMessage',
      method: 'post',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        channel: result.channel.id,
        text: message
      }),
      muteHttpExceptions: true
    };
  });
  const postResponses = UrlFetchApp.fetchAll(postRequests);
  postResponses.forEach((response, index) => {
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      throw new Error(`Slack DM送信に失敗しました: ${uniqueUserIds[index]} ${result.error || ''}`);
    }
  });
}

function authorizeRequiredScopes() {
  PropertiesService.getScriptProperties().getProperties();
  Session.getActiveUser().getEmail();
  ScriptApp.getOAuthToken();
  getDatabaseSheet(0).getLastRow();
  DriveApp.getRootFolder().getName();
  UrlFetchApp.fetch('https://slack.com/api/auth.test', {
    method: 'post',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });
}

/**
 * 指定のユーザーが指定のチャンネルに存在するか判定
 * @param {string} channelId - チャンネルID
 * @param {string} userId - ユーザーID
 * @returns {Boolian} - 存在可否
*/
function isUserInChannel(channelId = billChannelId, userId) {
  const url = 'https://slack.com/api/conversations.members';
  
  // リクエストパラメータ
  const params = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + token
    },
    payload: {
      channel: channelId
    }
  };
  // APIリクエスト
  const response = UrlFetchApp.fetch(url, params);
  const jsonResponse = JSON.parse(response.getContentText());
  // メンバーリストを取得
  const members = jsonResponse.members; // メンバーのユーザーIDリスト
  // ユーザーIDがメンバーリストに含まれているかをチェック
  if (members.includes(userId)) {
    return true;
  } 
  else {
    return false;
  }
}

/**
 * メール認証を行い、今月のプロジェクトごとの勤務時間及び合計金額(税込み)を返すAPI
 * @param {string} email - メールアドレス
 * @param {string} password - パスワード（平文）
 * @param {number} targetMonth - 対象月 (省略時は現在の月)
 * @param {number} targetYear - 対象年 (省略時は現在の年)
 * @returns {Object} - {status: string, data: {workTimeByProject: Object, totalAmount: number, month: number, year: number}}
 */
function getMonthlyWorkTimeAndSalaryWithAuth(email, password, targetMonth = null, targetYear = null, tokenForApi = null) {
  try {
    // 認証処理
    const authResult = tokenForApi ? authenticateUserByApiToken(email, tokenForApi) : authenticateUser(email, password);
    if (authResult.status !== 'success') {
      return authResult;
    }
    
    const attendanceSheetId = authResult.attendanceSheetId;
    const userId = authResult.userId;
    
    // 対象月・年の設定（省略時は現在の月・年）
    const currentDate = new Date();
    const month = targetMonth !== null ? targetMonth : currentDate.getMonth() + 1;
    const year = targetYear !== null ? targetYear : currentDate.getFullYear();
    if (isYearMonthAfter(year, month, currentDate.getFullYear(), currentDate.getMonth() + 1)) {
      return {
        status: 'error',
        message: '実績閲覧は今月まで指定できます'
      };
    }
    
    // 勤務時間計算（プロジェクト名 : 勤務時間(msec)）
    const projectWorkTimes = BillUtil.calculateMonthlyWorkTime(attendanceSheetId, month, year);
    
    // プロジェクト名から時給ごとに分ける（時給 : 勤務時間(msec)）
    const hourlyWageDict = BillUtil.projectNameToHourlyWage(projectWorkTimes, attendanceSheetId);
    
    // 0時間のプロジェクトを除外して、見やすい形式に変換
    const workTimeByProject = {};
    const workTimeFormatted = BillUtil.msecDictToHourAndMinutesDict(projectWorkTimes);
    
    for (let projectName in projectWorkTimes) {
      if (projectWorkTimes[projectName] > 0) {
        workTimeByProject[projectName] = workTimeFormatted[projectName];
      }
    }
    
    // 合計金額計算（税込み）
    let totalAmount = 0;
    const taxRate = 1 + (Number(PropertiesService.getScriptProperties().getProperty("TAX_PERCENTAGE") || "10") / 100);
    const projectItems = BillUtil.calculateMonthlyWorkItems(attendanceSheetId, month, year, projectWorkTimes);
    const roleAllowances = BillUtil.calculateRoleAllowanceItems(attendanceSheetId, month, year, projectWorkTimes);
    const lineItems = projectItems.concat(roleAllowances.map(item => ({
      type: 'allowance',
      id: item.id,
      name: item.name,
      unitPrice: item.price,
      workTimeMsec: item.workTimeMsec,
      workTimeText: item.workTimeText,
      amount: item.amount
    }))).map(item => {
      const amountWithTax = Math.round(item.amount * taxRate);
      totalAmount += amountWithTax;
      return {
        type: item.type,
        id: item.id,
        name: item.name,
        unitPrice: item.unitPrice,
        workTime: item.workTimeText,
        amount: item.amount,
        amountWithTax: amountWithTax
      };
    });
    
    return {
      status: 'success',
      data: {
        userId: userId,
        workTimeByProject: workTimeByProject,
        lineItems: lineItems,
        roleAllowances: roleAllowances.map(item => ({
          name: item.name,
          price: item.price,
          workTime: item.workTimeText,
          amount: item.amount,
          amountWithTax: Math.round(item.amount * taxRate)
        })),
        totalAmount: totalAmount,
        month: month,
        year: year,
        summary: `${year}年${month}月の勤務実績`
      }
    };
    
  } catch (error) {
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * メールアドレスとパスワードでユーザー認証を行う
 * @param {string} email - メールアドレス
 * @param {string} password - パスワード（平文）
 * @returns {Object} - {status: string, attendanceSheetId?: string, message?: string}
 */
function authenticateUser(email, password) {
  try {
    // データベースシートを取得
    const databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId);
    const userSheet = getDatabaseSheet(0);
    
    // 全データを取得（F列：メール、G列：パスワードハッシュ、B列：attendanceSheetId）
    const lastRow = userSheet.getLastRow();
    if (lastRow < 2) {
      return {
        status: 'error',
        message: 'ユーザーデータが存在しません'
      };
    }
    
    const userData = userSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    
    // パスワードのSHA256ハッシュ値を計算
    const passwordHash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password,
      Utilities.Charset.UTF_8
    );
    const passwordHashHex = passwordHash.map(byte => {
      return (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0');
    }).join('');
    
    // メールアドレスとパスワードハッシュで照合
    for (let i = 0; i < userData.length; i++) {
      const userEmail = userData[i][USER_COL.email - 1];
      const storedPasswordHash = userData[i][USER_COL.passwordHash - 1];
      const deletedAt = userData[i][USER_COL.deletedAt - 1];
      
      if (userEmail === email && storedPasswordHash === passwordHashHex) {
        if (deletedAt) {
          return {
            status: 'error',
            message: 'このユーザーは無効化されています'
          };
        }
        const attendanceSheetId = userData[i][USER_COL.attendanceSheetId - 1];
        
        if (!attendanceSheetId) {
          return {
            status: 'error',
            message: '勤怠シートIDが設定されていません'
          };
        }
        
        return {
          status: 'success',
          userId: userData[i][USER_COL.id - 1],
          attendanceSheetId: attendanceSheetId
        };
      }
    }
    
    return {
      status: 'error',
      message: 'メールアドレスまたはパスワードが正しくありません'
    };
    
  } catch (error) {
    return {
      status: 'error',
      message: `認証エラー: ${error.message}`
    };
  }
}

function authenticateUserByApiToken(email, tokenForApi) {
  if (!apiToken || tokenForApi !== apiToken) {
    return {
      status: 'error',
      message: 'APIトークンが正しくありません'
    };
  }
  const userSheet = getDatabaseSheet(0);
  const lastRow = userSheet.getLastRow();
  if (lastRow < 2) {
    return {
      status: 'error',
      message: 'ユーザーデータが存在しません'
    };
  }
  const values = userSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][USER_COL.email - 1] === email) {
      if (values[i][USER_COL.deletedAt - 1]) {
        return {
          status: 'error',
          message: 'このユーザーは無効化されています'
        };
      }
      return {
        status: 'success',
        userId: values[i][USER_COL.id - 1],
        attendanceSheetId: values[i][USER_COL.attendanceSheetId - 1]
      };
    }
  }
  return {
    status: 'error',
    message: 'メールアドレスが見つかりません'
  };
}

/**
 * 月次勤務実績API用GETハンドラー
 * URL: /exec?email=xxx@example.com&password=password&month=xx&year=xxxx
 */
function handleMonthlyWorkTimeApi(e) {
  try {
    const params = e.parameter;
    const email = params.email;
    const password = params.password;
    const tokenForApi = params.apiToken || params.token;
    
    if (!email || (!password && !tokenForApi)) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'error',
          message: 'emailとpassword、またはemailとapiTokenパラメータが必要です'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const month = params.month ? parseInt(params.month) : null;
    const year = params.year ? parseInt(params.year) : null;
    
    const result = getMonthlyWorkTimeAndSalaryWithAuth(email, password, month, year, tokenForApi);
    
    return ContentService
      .createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: `サーバーエラー: ${error.message}`
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GETリクエスト用ハンドラー
 * email/passwordが指定された場合は既存API、それ以外は勤怠入力フォームを表示する
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.mode === 'admin') {
    return renderAdminPage(e);
  }

  if (params.view === 'worktime') {
    return renderMonthlyWorkTimePage(e);
  }

  if (params.email && (params.password || params.apiToken || params.token)) {
    return handleMonthlyWorkTimeApi(e);
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();
  template.formHTML = createHtml(e);
  template.projectAccessJson = JSON.stringify(getProjectAccessForForm());
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('勤怠入力');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

// POSTメソッド
function doPost(e) {
  let databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId);

  if (e.parameter.mode === 'admin') {
    return handleAdminPost(e);
  }

  // 勤怠提出
  if (e.parameter.submit) {
    // バリデーション
    let alertContent = '';

    // 入力メールアドレスがデータベース内に存在するか
    let passOfEmail = BillUtil.search(getDatabaseSheet(0), e.parameter.email);

    // 存在する場合
    if (passOfEmail[0] > 0) {
      if (isUserDeletedRow(getDatabaseSheet(0), passOfEmail[0])) {
        alertContent += alert('・このユーザーは無効化されています');
      }
      // 入力されたパスワードのハッシュ化した値がメアドの隣のハッシュパスに一致するか
      else if (!(getDatabaseSheet(0).getRange('F2:G' + getDatabaseSheet(0).getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password))) {
        alertContent += alert('・アカウント情報が不正です');
      } else {
        // パスワード認証成功時のバリデーション
        alertContent += validateFormData(e);
        alertContent += validateProjectAccess(e.parameter.email, e.parameter.password, e.parameter.project);

        // 二重勤怠チェック（認証成功時のみ）
        if (alertContent === '' && e.parameter.startTime && e.parameter.endTime) {
          alertContent += checkDuplicateAttendance(e, passOfEmail[0]);
        }
      }
    } else {
      alertContent += alert('・アカウント情報が不正です');
    }

    // 基本的なフォームバリデーション（認証失敗時でも実行）
    if (!passOfEmail[0] || passOfEmail[0] <= 0 || isUserDeletedRow(getDatabaseSheet(0), passOfEmail[0]) || !(getDatabaseSheet(0).getRange('F2:G' + getDatabaseSheet(0).getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password))) {
      alertContent += validateFormData(e);
    }

    // 警告内容がなければ書き込んで完了ページへ、あれば警告とともに再表示
    if (alertContent === '') {
      const attendance = [[formatTimeToGoogleForm(new Date()), e.parameter.project, e.parameter.content, formatTimeToGoogleForm(e.parameter.startTime), formatTimeToGoogleForm(e.parameter.endTime)]];
      const sheet = SpreadsheetApp.openById(getDatabaseSheet(0).getRange('B' + passOfEmail[0]).getValue()).getSheets()[0];
      sheet.getRange(`A${sheet.getLastRow() + 1}:E${sheet.getLastRow() + 1}`).setValues(attendance);

      const template = HtmlService.createTemplateFromFile('submitted');
      template.deployURL = ScriptApp.getService().getUrl();
      const htmlOutput = template.evaluate();
      htmlOutput.setTitle('提出完了');
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    } else {
      const template = HtmlService.createTemplateFromFile('index');
      template.deployURL = ScriptApp.getService().getUrl();
      template.formHTML = createHtml(e, alertContent);
      template.projectAccessJson = JSON.stringify(getProjectAccessForForm());
      const htmlOutput = template.evaluate();
      htmlOutput.setTitle('勤怠入力');
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    }
  }

  // 勤怠一時保存
  if (e.parameter.save) {
    let databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId);
    let alertContent = '';
    let passOfEmail = BillUtil.search(getDatabaseSheet(0), e.parameter.email);

    // アカウント認証チェック
    let authenticationPassed = false;
    if (passOfEmail[0] > 0) {
      if (isUserDeletedRow(getDatabaseSheet(0), passOfEmail[0])) {
        alertContent += alert('・このユーザーは無効化されています');
      } else if (getDatabaseSheet(0).getRange('F2:G' + getDatabaseSheet(0).getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password)) {
        authenticationPassed = true;
      } else {
        alertContent += alert('・アカウント情報が不正です');
      }
    } else {
      alertContent += alert('・アカウント情報が不正です');
    }

    // 認証が成功した場合のみ保存処理を実行
    if (authenticationPassed) {
      alertContent += alert('Saved!!');

      // セキュアな新形式データを作成して保存
      const saveData = createSecureSaveData(e);
      getDatabaseSheet(0).getRange('H' + passOfEmail[0]).setValue(JSON.stringify(saveData));
    }

    // 画面の再表示（認証成功・失敗に関わらず）
    const template = HtmlService.createTemplateFromFile('index');
    template.deployURL = ScriptApp.getService().getUrl();
    template.formHTML = createHtml(e, alertContent);
    template.projectAccessJson = JSON.stringify(getProjectAccessForForm());
    const htmlOutput = template.evaluate();
    htmlOutput.setTitle('勤怠入力');
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return htmlOutput;
  }

  // 勤怠呼び出し
  if (e.parameter.load) {
    let databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId);
    let alertContent = '';
    let passOfEmail = BillUtil.search(getDatabaseSheet(0), e.parameter.email);

    if (passOfEmail[0] > 0) {
      if (isUserDeletedRow(getDatabaseSheet(0), passOfEmail[0])) {
        alertContent += alert('・このユーザーは無効化されています');
      } else if (!(getDatabaseSheet(0).getRange('F2:G' + getDatabaseSheet(0).getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password))) {
        alertContent += alert('・アカウント情報が不正です');
      }
      // 一度も保存がなされていないとき
      if (getDatabaseSheet(0).getRange('H' + passOfEmail[0]).getValue() === '') {
        alertContent += alert('・データが保存されていません');
      }
    } else {
      alertContent += alert('・アカウント情報が不正です');
    }

    // アカウント情報が正しければ保存データを呼び出し
    if (alertContent === '') {
      alertContent += alert('Loaded!!');
      const savedDataJson = getDatabaseSheet(0).getRange('H' + passOfEmail[0]).getValue();

      // 保存データを復元して現在のパスワードと結合
      const loaded = createLoadedEvent(savedDataJson, e.parameter.email, e.parameter.password);

      const template = HtmlService.createTemplateFromFile('index');
      template.deployURL = ScriptApp.getService().getUrl();
      template.formHTML = createHtml(loaded, alertContent);
      template.projectAccessJson = JSON.stringify(getProjectAccessForForm());
      const htmlOutput = template.evaluate();
      htmlOutput.setTitle('勤怠入力');
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    } else {
      const template = HtmlService.createTemplateFromFile('index');
      template.deployURL = ScriptApp.getService().getUrl();
      template.formHTML = createHtml(e, alertContent);
      template.projectAccessJson = JSON.stringify(getProjectAccessForForm());
      const htmlOutput = template.evaluate();
      htmlOutput.setTitle('勤怠入力');
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    }
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();
  template.formHTML = createHtml(e, alert('・不正な操作です'));
  template.projectAccessJson = JSON.stringify(getProjectAccessForForm());
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('勤怠入力');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

// フォームデータのバリデーション関数
function validateFormData(e) {
  let alertContent = '';

  // PJ名が選択されているか
  if (!e.parameter.project) {
    alertContent += alert('・PJを選択してください');
  }

  // 作業内容が入力されているか
  if (!e.parameter.content) {
    alertContent += alert('・作業内容を入力してください');
  }

  // 開始時刻と終了時刻が入力されているか、入力されている場合は0<{勤務時間}<24hであるか
  if (!e.parameter.startTime || !e.parameter.endTime) {
    alertContent += alert('・開始時刻、終了時刻を正しく入力してください');
  } else {
    if (new Date(e.parameter.endTime) - new Date(e.parameter.startTime) <= 0 || new Date(e.parameter.endTime) - new Date(e.parameter.startTime) >= 86400000) {
      alertContent += alert('・勤務時間が不正です');
    }
  }

  return alertContent;
}

function validateProjectAccess(email, password, projectName) {
  if (!projectName) {
    return '';
  }
  const selectableProjects = getSelectableProjects(email, password).map(project => String(project.name));
  if (!selectableProjects.includes(String(projectName))) {
    return alert('・選択されたPJは現在利用できません');
  }
  return '';
}

// 二重勤怠チェック関数
function checkDuplicateAttendance(e, passOfEmailIndex) {
  let alertContent = '';

  const attendanceSheet = SpreadsheetApp.openById(getDatabaseSheet(0).getRange('B' + passOfEmailIndex).getValue()).getSheets()[0];

  // 初回勤怠のときは二重勤怠バリデーションをスキップ
  if (!(attendanceSheet.getLastRow() === 1)) {
    const previousAttendanceArray = attendanceSheet.getRange('A2:E' + attendanceSheet.getLastRow()).getValues();
    for (let i = 0; i < previousAttendanceArray.length; i++) {
      let startTime = new Date(e.parameter.startTime);
      let endTime = new Date(e.parameter.endTime);
      let previousStartTime = new Date(previousAttendanceArray[i][3]);
      let previousEndTime = new Date(previousAttendanceArray[i][4]);
      if (!(startTime >= previousEndTime || endTime <= previousStartTime)) {
        alertContent += alert(`・入力された勤怠はタイムスタンプ:${previousAttendanceArray[i][0]}に提出された勤怠と重複する時間があります。開始時刻と終了時刻を再度確認してください。`);
      }
    }
  }

  return alertContent;
}

// セキュアな保存データ作成関数
function createSecureSaveData(e) {
  // 必要なフィールドのみを明示的に抽出してパスワード情報を完全に除去
  const saveData = {
    parameter: {
      email: sanitizeInput(e.parameter.email || ''),
      project: sanitizeInput(e.parameter.project || ''),
      content: sanitizeInput(e.parameter.content || ''),
      startTime: e.parameter.startTime || '',
      endTime: e.parameter.endTime || ''
    }
  };

  // 新形式であることを明示するためのフラグを追加
  saveData.formatVersion = 'secure_v1';

  return saveData;
}

// 保存データからイベントオブジェクトを復元する関数（後方互換性あり）
function createLoadedEvent(savedDataJson, email, password) {
  const savedData = JSON.parse(savedDataJson);

  // 旧形式と新形式の両方に対応
  let project = '';
  let content = '';
  let startTime = '';
  let endTime = '';

  if (isOldFormat(savedData)) {
    // 旧形式の場合：既存のデータ構造から抽出
    console.log('旧形式のデータを検出しました');
    project = savedData.parameter.project || '';
    content = savedData.parameter.content || '';
    startTime = savedData.parameter.startTime || '';
    endTime = savedData.parameter.endTime || '';
  } else {
    // 新形式の場合：シンプルな構造から抽出
    project = savedData.parameter.project || '';
    content = savedData.parameter.content || '';
    startTime = savedData.parameter.startTime || '';
    endTime = savedData.parameter.endTime || '';
  }

  // 新しいイベントオブジェクトを作成
  const loaded = {
    parameter: {
      email: email,
      password: password,
      project: project,
      content: content,
      startTime: startTime,
      endTime: endTime
    }
  };

  return loaded;
}

// 旧形式のデータかどうかを判定する関数
function isOldFormat(savedData) {
  // 新形式にはformatVersionフラグが存在する
  if (savedData.formatVersion === 'secure_v1') {
    return false;
  }

  // 旧形式には parameters, postData, contextPath などのプロパティが存在する
  return savedData.hasOwnProperty('parameters') ||
    savedData.hasOwnProperty('postData') ||
    savedData.hasOwnProperty('contextPath') ||
    savedData.hasOwnProperty('contentLength');
}

// 質問部HTMLタグを動的に生成
function createHtml(e, alert = '') {
  // eが存在しない場合のデフォルト処理
  if (!e || !e.parameter) {
    e = { parameter: {} };
  }

  // eに入力値があればそれ、なければ''を入れる
  // テキスト入力系に対してはサニタイズを行う
  const email = sanitizeInput(e.parameter.email || '') || '';
  const password = sanitizeInput(e.parameter.password || '') || '';
  const project = sanitizeInput(e.parameter.project || '') || '';
  const content = sanitizeInput(e.parameter.content || '') || '';
  const startTime = e.parameter.startTime || '';
  const endTime = e.parameter.endTime || '';

  // データベースに基づくPJ選択肢の動的生成
  let pulldown = '';
  if (project === '') {
    pulldown += '<option value="" selected disabled>選択してください</option>';
  } else {
    pulldown += '<option value="" disabled>選択してください</option>';
  }

  // タグ全体の生成
  let html = `
        <div class="mb-3">
            <label for="email" class="form-label">Email</label>
            <input type="email" class="form-control" id="email" name="email" value="${email}" required>
        </div>
        <div class="mb-3">
            <label for="password" class="form-label">パスワード</label>
            <input type="password" class="form-control" id="password" name="password" value="${password}" autocomplete="off" required>
        </div>
        ${alert}
        <table class="table">
            <tr>
                <td>PJ名</td>
                <td>
                    <select class="form-select" id="projectSelect" name="project" data-selected-project="${project}" autocomplete="off">
                        ${pulldown}
                    </select>
                </td>
            </tr>
            <tr>
                <td>作業内容</td>
                <td>
                    <input class="form-control" name="content" type="text" value="${content}" autocomplete="off">
                </td>
            </tr>
            <tr>
                <td>開始時刻</td>
                <td>
                    <input type="datetime-local" name="startTime" value="${startTime}">
                </td>
            </tr>
            <tr>
                <td>終了時刻</td>
                <td>
                    <input type="datetime-local" name="endTime" value="${endTime}">
                </td>
            </tr>
        </table>
    `;
  return html;
}

function getProjectAccessForForm() {
  const sheets = getDatabaseSheets();
  const usersSheet = sheets[0];
  const projectsSheet = sheets[1];
  if (!projectsSheet || projectsSheet.getLastRow() < 2) {
    return [];
  }

  const userRows = usersSheet && usersSheet.getLastRow() >= 2
    ? usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 10).getValues()
    : [];
  const emailHashByUserId = {};
  userRows.forEach(row => {
    const userId = String(row[USER_COL.id - 1] || '');
    const email = row[USER_COL.email - 1];
    const deletedAt = row[USER_COL.deletedAt - 1];
    if (userId && email && !deletedAt) {
      emailHashByUserId[userId] = hashEmailForProjectAccess(email);
    }
  });

  const projectRows = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 5).getValues();
  return projectRows
    .filter(row => row[PROJECT_COL.name - 1])
    .filter(row => isTruthyCell(row[PROJECT_COL.active - 1]))
    .filter(row => String(row[PROJECT_COL.assignedUserIds - 1] || '').trim() !== NO_ASSIGNED_USERS)
    .map(row => {
      const rawAssignedUserIds = String(row[PROJECT_COL.assignedUserIds - 1] || '').trim();
      const assignedUserIds = rawAssignedUserIds === NO_ASSIGNED_USERS ? [] : rawAssignedUserIds.split(',').map(value => value.trim()).filter(Boolean);
      const assignedEmailHashes = assignedUserIds.map(userId => emailHashByUserId[userId]).filter(Boolean);
      return {
        name: String(row[PROJECT_COL.name - 1]),
        allUsers: !rawAssignedUserIds,
        assignedEmailHashes: assignedEmailHashes
      };
    });
}

function normalizeEmailForProjectAccess(email) {
  return String(email || '').trim().toLowerCase();
}

function hashEmailForProjectAccess(email) {
  const normalized = normalizeEmailForProjectAccess(email);
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalized,
    Utilities.Charset.UTF_8
  );
  return digest.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}

function getSelectableProjects(email, password) {
  const sheets = getDatabaseSheets();
  const projectsSheet = sheets[1];
  if (!projectsSheet || projectsSheet.getLastRow() < 2) {
    return [];
  }

  let userId = '';
  if (email && password) {
    const authResult = authenticateUser(email, password);
    if (authResult.status === 'success') {
      userId = String(authResult.userId);
    }
  }

  const projectValues = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 5).getValues();
  return projectValues
    .filter(row => row[PROJECT_COL.name - 1])
    .filter(row => isTruthyCell(row[PROJECT_COL.active - 1]))
    .filter(row => {
      const rawAssigned = String(row[PROJECT_COL.assignedUserIds - 1] || '').trim();
      if (!rawAssigned) {
        return true;
      }
      if (rawAssigned === NO_ASSIGNED_USERS) {
        return false;
      }
      const assigned = rawAssigned.split(',').map(v => v.trim()).filter(Boolean);
      return userId && assigned.includes(userId);
    })
    .map(row => ({
      id: row[PROJECT_COL.id - 1],
      name: row[PROJECT_COL.name - 1],
      hourlyWage: row[PROJECT_COL.hourlyWage - 1]
    }));
}

function isTruthyCell(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return true;
  }
  const normalized = String(value).toLowerCase();
  return value === true || normalized === 'true' || normalized === 'active' || normalized === '有効' || normalized === '1';
}

function isUserDeletedRow(userSheet, row) {
  return Boolean(userSheet.getRange(row, USER_COL.deletedAt).getValue());
}

function renderMonthlyWorkTimePage(e) {
  const params = e && e.parameter ? e.parameter : {};
  let month = null;
  let year = null;
  if (params.monthForWorkTime) {
    const parts = params.monthForWorkTime.split('-');
    year = Number(parts[0]);
    month = Number(parts[1]);
  } else {
    month = params.month ? Number(params.month) : null;
    year = params.year ? Number(params.year) : null;
  }
  if (!month || month < 1 || month > 12) {
    month = null;
  }
  if (!year || year < 1900) {
    year = null;
  }

  let result;
  if (month && year && isYearMonthAfter(year, month, new Date().getFullYear(), new Date().getMonth() + 1)) {
    result = {
      status: 'error',
      message: '実績閲覧は今月まで指定できます'
    };
  } else {
    result = getMonthlyWorkTimeAndSalaryWithAuth(params.email, params.password, month, year, params.apiToken || params.token);
  }
  const template = HtmlService.createTemplateFromFile('worktime');
  template.deployURL = ScriptApp.getService().getUrl();
  template.resultJson = JSON.stringify(result);
  template.email = params.email || '';
  template.password = params.password || '';
  template.apiToken = params.apiToken || params.token || '';
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('実績');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}

function normalizeEmailForAdmin(email) {
  return String(email || '').trim().toLowerCase();
}

function getAdminWhitelistSheet() {
  return openDatabaseSpreadsheet().getSheetByName(ADMIN_WHITELIST_SHEET_NAME);
}

function setupAdminWhitelistSheet() {
  const database = openDatabaseSpreadsheet();
  let sheet = database.getSheetByName(ADMIN_WHITELIST_SHEET_NAME);
  if (!sheet) {
    sheet = database.insertSheet(ADMIN_WHITELIST_SHEET_NAME);
    sheet.appendRow(['email', 'active', 'memo']);
  }
  database.setActiveSheet(sheet);
  database.moveActiveSheet(database.getNumSheets());
  return sheet;
}

function getAdminAllowedEmailsFromDatabase() {
  const sheet = getAdminWhitelistSheet();
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .filter(row => row[0])
    .filter(row => row[1] === '' || row[1] === null || typeof row[1] === 'undefined' || isTruthyCell(row[1]))
    .map(row => normalizeEmailForAdmin(row[0]))
    .filter(Boolean);
}

function getAdminWhitelistEntries() {
  const sheet = getAdminWhitelistSheet();
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().map(row => ({
    email: normalizeEmailForAdmin(row[0]),
    active: row[1] === '' || row[1] === null || typeof row[1] === 'undefined' || isTruthyCell(row[1]),
    memo: row[2] || ''
  })).filter(row => row.email);
}

function isAdminUser() {
  const activeEmail = normalizeEmailForAdmin(getActiveUserEmail());
  const allowedEmails = getAdminAllowedEmailsFromDatabase();
  return Boolean(activeEmail && allowedEmails.includes(activeEmail));
}

function renderAdminPage(e, message = '') {
  const params = e && e.parameter ? e.parameter : {};
  const activeEmail = getActiveUserEmail();
  const authed = isAdminUser();
  const template = HtmlService.createTemplateFromFile('admin');
  template.deployURL = ScriptApp.getService().getUrl();
  template.activeEmail = activeEmail || '';
  template.message = message || (params.message ? alert(params.message) : '');
  template.authed = authed;
  template.snapshotJson = JSON.stringify(authed ? getAdminSnapshot() : {});
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('勤怠管理');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

function handleAdminPost(e) {
  const result = runAdminAction(e.parameter);
  return renderAdminPage(e, alert(result.message));
}

function runAdminAction(params) {
  if (!isAdminUser()) {
    return { status: 'error', message: 'このGoogleアカウントは管理画面の利用を許可されていません' };
  }

  try {
    executeAdminAction(params || {});
    return { status: 'success', message: '完了しました' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function loadLatestBillReactionSnapshot() {
  if (!isAdminUser()) {
    return { status: 'error', message: 'このGoogleアカウントは管理画面の利用を許可されていません' };
  }
  return { status: 'success', latestBillReaction: getLatestBillReactionSnapshot() };
}

function loadAdminSnapshot(options) {
  if (!isAdminUser()) {
    return { status: 'error', message: 'このGoogleアカウントは管理画面の利用を許可されていません' };
  }
  const includeLatestBillReaction = !options || options.includeLatestBillReaction !== false;
  return { status: 'success', snapshot: getAdminSnapshot(includeLatestBillReaction) };
}

function executeAdminAction(params) {
  const action = params.action;
  if (action === 'addProject') {
    addProject(params);
  } else if (action === 'updateProject') {
    updateProject(params);
  } else if (action === 'bulkUpdateProjectStatus') {
    bulkUpdateProjectStatus(params);
  } else if (action === 'updateProjectAssignments') {
    updateProjectAssignments(params);
  } else if (action === 'addWageUpgrade') {
    addWageUpgrade(params);
  } else if (action === 'updateWageUpgrades') {
    updateWageUpgrades(params);
  } else if (action === 'upsertUser') {
    upsertUser(params);
  } else if (action === 'deleteUser') {
    logicalDeleteUser(params.userId);
  } else if (action === 'addRoleAllowance') {
    addRoleAllowance(params);
  } else if (action === 'updateRoleAllowance') {
    updateRoleAllowance(params);
  } else if (action === 'deleteRoleAllowance') {
    deleteRoleAllowance(params.allowanceId);
  } else if (action === 'addAdminWhitelistEntry') {
    addAdminWhitelistEntry(params);
  } else if (action === 'updateAdminWhitelistEntry') {
    updateAdminWhitelistEntry(params);
  } else if (action === 'deleteAdminWhitelistEntry') {
    deleteAdminWhitelistEntry(params);
  } else if (action === 'setBillForPeriod') {
    runSetBillForPeriod(params);
  } else if (action === 'gatherBillsForPeriod') {
    runGatherBillsForPeriod(params);
  } else if (action === 'resetBillForPeriod') {
    runResetBillForPeriod(params);
  } else if (action === 'remindUnreactedUser') {
    remindUnreactedUser(params);
  } else if (action === 'bulkRemindUnreactedUsers') {
    bulkRemindUnreactedUsers(params);
  } else if (action === 'setBill') {
    const year = Number(params.year) || new Date().getFullYear();
    const month = Number(params.month) || new Date().getMonth() + 1;
    validateBillExecutionDate(year, month);
    setBill(year, month);
  } else if (action === 'gatherBills') {
    const year = Number(params.year) || new Date().getFullYear();
    const month = Number(params.month) || new Date().getMonth() + 1;
    validateBillExecutionDate(year, month);
    gatherBills(year, month);
  } else if (action === 'resetBill') {
    const year = Number(params.year) || new Date().getFullYear();
    const month = Number(params.month) || new Date().getMonth() + 1;
    validateBillExecutionDate(year, month);
    reSetBill(params.attendanceSheetId, year, month);
  }
}

function getAdminSnapshot(includeLatestBillReaction = false) {
  const sheets = getDatabaseSheets();
  const users = sheetRowsToObjects(sheets[0], ['id','attendanceSheetId','slackUserId','driveFolderId','addressSheetId','email','passwordHash','savedData','name','deletedAt']);
  const activeUsers = users.filter(user => !user.deletedAt);
  const userNameById = {};
  const activeUserIds = {};
  activeUsers.forEach(user => userNameById[String(user.id)] = user.name || '');
  activeUsers.forEach(user => activeUserIds[String(user.id)] = true);
  const projects = sheetRowsToObjects(sheets[1], ['id','name','hourlyWage','active','assignedUserIds']).map(project => {
    const rawAssignedUserIds = String(project.assignedUserIds || '').trim();
    const noAssignedUsers = rawAssignedUserIds === NO_ASSIGNED_USERS;
    const assignedUsers = rawAssignedUserIds && !noAssignedUsers ? rawAssignedUserIds.split(',').map(id => id.trim()).filter(id => activeUserIds[id]).map(id => ({
      id: id,
      label: userNameById[id] ? `${id}_${userNameById[id]}` : id
    })) : [];
    const assignedLabels = assignedUsers.map(user => user.label);
    return {
      ...project,
      activeInterpreted: isTruthyCell(project.active) ? '有効' : '無効',
      assignedInterpreted: noAssignedUsers ? '割当なし' : (rawAssignedUserIds ? assignedLabels.join(', ') || '割当なし' : '全員'),
      assignedSummary: noAssignedUsers ? '0人' : (rawAssignedUserIds ? `${assignedUsers.length}人` : '全員'),
      assignedDetail: noAssignedUsers ? '割当なし' : (rawAssignedUserIds ? assignedLabels.join(', ') || '割当なし' : '全員')
    };
  });
  return {
    activeUsers: activeUsers,
    projects: projects,
    wageUpgrades: sheetRowsToObjects(sheets[2], ['id','userId','projectId','increment']).filter(row => activeUserIds[String(row.userId)]),
    roleAllowances: sheets[3] ? sheetRowsToObjects(sheets[3], ['id','name','userId','price','deletedAt'])
      .filter(row => !row.deletedAt)
      .filter(row => activeUserIds[String(row.userId)]) : [],
    adminWhitelist: getAdminWhitelistEntries(),
    latestBillReaction: includeLatestBillReaction ? getLatestBillReactionSnapshot() : null,
    adminExecutionLogs: getAdminExecutionLogs(10)
  };
}

function getAdminExecutionLogSheet() {
  const database = openDatabaseSpreadsheet();
  let sheet = database.getSheetByName(ADMIN_EXECUTION_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = database.insertSheet(ADMIN_EXECUTION_LOG_SHEET_NAME);
    sheet.appendRow(['timestamp', 'action', 'period', 'status', 'message', 'operatorEmail']);
  }
  database.setActiveSheet(sheet);
  database.moveActiveSheet(database.getNumSheets());
  return sheet;
}

function appendAdminExecutionLog(action, period, status, message) {
  const sheet = getAdminExecutionLogSheet();
  sheet.appendRow([new Date(), action, period, status, message || '', getActiveUserEmail() || '']);
}

function getAdminExecutionLogs(limit) {
  const database = openDatabaseSpreadsheet();
  const sheet = database.getSheetByName(ADMIN_EXECUTION_LOG_SHEET_NAME);
  if (!sheet) {
    return [];
  }
  database.setActiveSheet(sheet);
  database.moveActiveSheet(database.getNumSheets());
  if (sheet.getLastRow() < 2) {
    return [];
  }
  const rowCount = Math.min(Number(limit || 10), sheet.getLastRow() - 1);
  const startRow = sheet.getLastRow() - rowCount + 1;
  return sheet.getRange(startRow, 1, rowCount, 6).getValues().reverse().map(row => ({
    timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss') : '',
    action: row[1],
    period: row[2],
    status: row[3],
    message: row[4],
    operatorEmail: row[5]
  }));
}

function runExclusiveAdminBatch(action, period, callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('別の一括処理が実行中です。完了後に再実行してください');
  }
  appendAdminExecutionLog(action, period, 'running', '開始');
  try {
    callback();
    appendAdminExecutionLog(action, period, 'success', '完了');
  } catch (error) {
    appendAdminExecutionLog(action, period, 'error', error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function getLatestBillReactionSnapshot() {
  try {
    const activeUsers = sheetRowsToObjects(getDatabaseSheet(0), ['id','attendanceSheetId','slackUserId','driveFolderId','addressSheetId','email','passwordHash','savedData','name','deletedAt'])
      .filter(user => !user.deletedAt && user.slackUserId);
    const historyResponse = UrlFetchApp.fetch(`https://slack.com/api/conversations.history?channel=${encodeURIComponent(billChannelId)}&limit=20`, {
      method: 'get',
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    const history = JSON.parse(historyResponse.getContentText());
    if (!history.ok) {
      return { status: 'error', message: history.error || 'Slack履歴の取得に失敗しました', unreactedUsers: [] };
    }
    const targetMessage = (history.messages || []).find(message => String(message.text || '').indexOf('請求書が出力されました') !== -1);
    if (!targetMessage) {
      return { status: 'not_found', message: '請求書おまとめメッセージが見つかりません', unreactedUsers: [] };
    }
    const reactedUserIds = {};
    (targetMessage.reactions || []).forEach(reaction => {
      (reaction.users || []).forEach(userId => reactedUserIds[userId] = true);
    });
    return {
      status: 'success',
      messageTs: targetMessage.ts,
      unreactedUsers: activeUsers.filter(user => !reactedUserIds[user.slackUserId]).map(user => ({
        id: user.id,
        name: user.name,
        email: user.email,
        slackUserId: user.slackUserId
      }))
    };
  } catch (error) {
    return { status: 'error', message: error.message, unreactedUsers: [] };
  }
}

function sheetRowsToObjects(sheet, keys) {
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, keys.length).getValues().map(row => {
    const obj = {};
    keys.forEach((key, i) => obj[key] = row[i]);
    return obj;
  });
}

function nextNumericId(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return 1;
  }
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(Number).filter(v => !isNaN(v));
  return values.length ? Math.max.apply(null, values) + 1 : 1;
}

function addProject(params) {
  const sheet = getDatabaseSheet(1);
  sheet.appendRow([nextNumericId(sheet), params.projectName, Number(params.hourlyWage || 0), true, params.assignedUserIds || '']);
}

function updateProject(params) {
  const sheet = getDatabaseSheet(1);
  const row = findRowById(sheet, params.projectId);
  const nextActive = params.active === 'true';
  let nextAssignedUserIds = normalizeAssignedUserIdsForAdmin(params.assignedUserIds || '');
  if (nextActive && nextAssignedUserIds === NO_ASSIGNED_USERS) {
    nextAssignedUserIds = '';
  }
  if (isTruthyCell(sheet.getRange(row, PROJECT_COL.active).getValue()) !== nextActive) {
    sheet.getRange(row, PROJECT_COL.active).setValue(nextActive);
  }
  if (normalizeAssignedUserIdsForAdmin(sheet.getRange(row, PROJECT_COL.assignedUserIds).getValue()) !== nextAssignedUserIds) {
    sheet.getRange(row, PROJECT_COL.assignedUserIds).setValue(nextAssignedUserIds);
  }
}

function bulkUpdateProjectStatus(params) {
  const projectIds = String(params.projectIds || '').split(',').map(value => value.trim()).filter(Boolean);
  if (!projectIds.length) {
    throw new Error('対象PJが選択されていません');
  }
  const sheet = getDatabaseSheet(1);
  const nextActive = params.active === 'true';
  if (sheet.getLastRow() < 2) {
    return;
  }
  const selectedIds = {};
  projectIds.forEach(projectId => selectedIds[String(projectId)] = true);
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, PROJECT_COL.assignedUserIds);
  const values = range.getValues();
  let changed = false;
  values.forEach(row => {
    if (selectedIds[String(row[PROJECT_COL.id - 1])] && isTruthyCell(row[PROJECT_COL.active - 1]) !== nextActive) {
      row[PROJECT_COL.active - 1] = nextActive;
      if (nextActive && String(row[PROJECT_COL.assignedUserIds - 1] || '').trim() === NO_ASSIGNED_USERS) {
        row[PROJECT_COL.assignedUserIds - 1] = '';
      }
      changed = true;
    }
  });
  if (changed) {
    range.setValues(values);
  }
}

function updateProjectAssignments(params) {
  const sheet = getDatabaseSheet(1);
  const row = findRowById(sheet, params.projectId);
  const nextAssignedUserIds = normalizeAssignedUserIdsForAdmin(params.assignedUserIds || '');
  if (nextAssignedUserIds === NO_ASSIGNED_USERS && isTruthyCell(sheet.getRange(row, PROJECT_COL.active).getValue())) {
    sheet.getRange(row, PROJECT_COL.active).setValue(false);
  }
  if (normalizeAssignedUserIdsForAdmin(sheet.getRange(row, PROJECT_COL.assignedUserIds).getValue()) !== nextAssignedUserIds) {
    sheet.getRange(row, PROJECT_COL.assignedUserIds).setValue(nextAssignedUserIds);
  }
}

function addWageUpgrade(params) {
  const sheet = getDatabaseSheet(2);
  sheet.appendRow([nextNumericId(sheet), params.userId, params.projectId, Number(params.increment || 0)]);
}

function updateWageUpgrades(params) {
  const sheet = getDatabaseSheet(2);
  const userId = String(params.userId || '');
  const projectIds = String(params.projectIds || '').split(',').map(value => value.trim()).filter(Boolean);
  const increments = String(params.increments || '').split(',').map(value => Number(value || 0));
  const existing = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const rowByProjectId = {};
  for (let i = 0; i < existing.length; i++) {
    if (String(existing[i][1]) === userId) {
      rowByProjectId[String(existing[i][2])] = i + 2;
    }
  }
  for (let i = 0; i < projectIds.length; i++) {
    const projectId = projectIds[i];
    const increment = Number(increments[i] || 0);
    if (rowByProjectId[projectId]) {
      sheet.getRange(rowByProjectId[projectId], 4).setValue(increment);
    } else if (increment !== 0) {
      sheet.appendRow([nextNumericId(sheet), userId, projectId, increment]);
    }
  }
}

function upsertUser(params) {
  const sheet = getDatabaseSheet(0);
  let row = params.userId ? findRowById(sheet, params.userId, false) : -1;
  const passwordHash = params.password ? hashWithSHA256(params.password) : params.passwordHash;
  if (row > 0) {
    const currentPasswordHash = sheet.getRange(row, USER_COL.passwordHash).getValue();
    const currentName = sheet.getRange(row, USER_COL.name).getValue();
    sheet.getRange(row, 1, 1, 10).setValues([[
      params.userId,
      params.attendanceSheetId,
      params.slackUserId,
      params.driveFolderId,
      params.addressSheetId,
      params.email,
      passwordHash || currentPasswordHash,
      sheet.getRange(row, USER_COL.savedData).getValue(),
      params.name || currentName,
      ''
    ]]);
  } else {
    sheet.appendRow([
      nextNumericId(sheet),
      params.attendanceSheetId,
      params.slackUserId,
      params.driveFolderId,
      params.addressSheetId,
      params.email,
      passwordHash,
      '',
      params.name,
      ''
    ]);
  }
}

function logicalDeleteUser(userId) {
  const sheet = getDatabaseSheet(0);
  const row = findRowById(sheet, userId);
  sheet.getRange(row, USER_COL.deletedAt).setValue(new Date());
}

function addRoleAllowance(params) {
  const sheets = getDatabaseSheets();
  const database = openDatabaseSpreadsheet();
  const sheet = sheets[3] || database.getSheetByName('RoleAllowances') || database.insertSheet('RoleAllowances');
  if (sheets[3] !== sheet) {
    database.setActiveSheet(sheet);
    database.moveActiveSheet(Math.min(4, database.getNumSheets()));
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', '手当名', '対象user_id', 'price', '論理削除日時']);
  }
  sheet.appendRow([nextNumericId(sheet), params.allowanceName, params.userId, Number(params.price || 0), '']);
}

function updateRoleAllowance(params) {
  const sheet = getDatabaseSheet(3);
  const row = findRowById(sheet, params.allowanceId);
  const nextValues = [params.allowanceName, params.userId, Number(params.price || 0)];
  const currentValues = sheet.getRange(row, 2, 1, 3).getValues()[0];
  if (String(currentValues[0] || '') !== String(nextValues[0] || '')
    || String(currentValues[1] || '') !== String(nextValues[1] || '')
    || Number(currentValues[2] || 0) !== Number(nextValues[2] || 0)) {
    sheet.getRange(row, 2, 1, 3).setValues([nextValues]);
  }
}

function deleteRoleAllowance(allowanceId) {
  const sheet = getDatabaseSheet(3);
  const row = findRowById(sheet, allowanceId);
  sheet.getRange(row, 5).setValue(new Date());
}

function findAdminWhitelistRowByEmail(sheet, email, shouldThrow = true) {
  const normalizedEmail = normalizeEmailForAdmin(email);
  if (sheet.getLastRow() < 2) {
    if (shouldThrow) {
      throw new Error(`管理者 ${normalizedEmail} が見つかりません`);
    }
    return -1;
  }
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeEmailForAdmin(values[i][0]) === normalizedEmail) {
      return i + 2;
    }
  }
  if (shouldThrow) {
    throw new Error(`管理者 ${normalizedEmail} が見つかりません`);
  }
  return -1;
}

function assertNotCurrentAdminEmail(email, operation) {
  if (normalizeEmailForAdmin(email) === normalizeEmailForAdmin(getActiveUserEmail())) {
    throw new Error(`ログイン中の管理者自身は${operation}できません`);
  }
}

function addAdminWhitelistEntry(params) {
  const email = normalizeEmailForAdmin(params.email);
  if (!email) {
    throw new Error('メールアドレスを入力してください');
  }
  const sheet = setupAdminWhitelistSheet();
  if (findAdminWhitelistRowByEmail(sheet, email, false) > 0) {
    throw new Error('このメールアドレスは既に登録されています');
  }
  sheet.appendRow([email, true, params.memo || '']);
}

function updateAdminWhitelistEntry(params) {
  const email = normalizeEmailForAdmin(params.email);
  const nextActive = params.active === 'true';
  if (!nextActive) {
    assertNotCurrentAdminEmail(email, '無効化');
  }
  const sheet = setupAdminWhitelistSheet();
  const row = findAdminWhitelistRowByEmail(sheet, email);
  sheet.getRange(row, 2, 1, 2).setValues([[nextActive, params.memo || '']]);
}

function deleteAdminWhitelistEntry(params) {
  const email = normalizeEmailForAdmin(params.email);
  assertNotCurrentAdminEmail(email, '削除');
  const sheet = setupAdminWhitelistSheet();
  const row = findAdminWhitelistRowByEmail(sheet, email);
  sheet.deleteRow(row);
}

function aggregationPeriodToExecutionDate(params) {
  const today = new Date();
  const period = params.aggregationPeriod || '';
  if (!period) {
    const aggregationDate = new Date(today.getFullYear(), today.getMonth(), 0);
    return {
      year: today.getFullYear(),
      month: today.getMonth() + 1,
      aggregationYear: aggregationDate.getFullYear(),
      aggregationMonth: aggregationDate.getMonth() + 1
    };
  }
  const parts = period.split('-');
  const aggregationYear = Number(parts[0]);
  const aggregationMonth = Number(parts[1]);
  if (!aggregationYear || !aggregationMonth || aggregationMonth < 1 || aggregationMonth > 12) {
    throw new Error('年月の指定が不正です');
  }
  const executionDate = new Date(aggregationYear, aggregationMonth, 1);
  return {
    year: executionDate.getFullYear(),
    month: executionDate.getMonth() + 1,
    aggregationYear: aggregationYear,
    aggregationMonth: aggregationMonth
  };
}

function isYearMonthAfter(year, month, limitYear, limitMonth) {
  return Number(year) > Number(limitYear) || (Number(year) === Number(limitYear) && Number(month) > Number(limitMonth));
}

function validateBillAggregationPeriod(execution) {
  const today = new Date();
  if (!isYearMonthAfter(today.getFullYear(), today.getMonth() + 1, execution.aggregationYear, execution.aggregationMonth)) {
    throw new Error('請求書の発行・回収・再発行は先月以前のみ指定できます');
  }
}

function validateBillExecutionDate(year, month) {
  const aggregationDate = new Date(Number(year), Number(month) - 1, 0);
  validateBillAggregationPeriod({
    aggregationYear: aggregationDate.getFullYear(),
    aggregationMonth: aggregationDate.getMonth() + 1
  });
}

function runSetBillForPeriod(params) {
  const execution = aggregationPeriodToExecutionDate(params);
  validateBillAggregationPeriod(execution);
  const period = `${execution.aggregationYear}-${String(execution.aggregationMonth).padStart(2, '0')}`;
  runExclusiveAdminBatch('一括発行', period, () => setBill(execution.year, execution.month));
}

function runGatherBillsForPeriod(params) {
  const execution = aggregationPeriodToExecutionDate(params);
  validateBillAggregationPeriod(execution);
  const period = `${execution.aggregationYear}-${String(execution.aggregationMonth).padStart(2, '0')}`;
  runExclusiveAdminBatch('一括回収', period, () => gatherBills(execution.year, execution.month));
}

function runResetBillForPeriod(params) {
  const execution = aggregationPeriodToExecutionDate(params);
  validateBillAggregationPeriod(execution);
  reSetBill(params.attendanceSheetId, execution.year, execution.month);
}

function remindUnreactedUser(params) {
  if (!params.slackUserId) {
    throw new Error('Slack IDが指定されていません');
  }
  const message = '請求書チャンネルのおまとめ投稿へのリアクションが未確認です。内容を確認し、問題なければリアクションをお願いします。';
  sendDirectmessage(message, params.slackUserId);
}

function bulkRemindUnreactedUsers(params) {
  const slackUserIds = String(params.slackUserIds || '').split(',').map(value => value.trim()).filter(Boolean);
  if (!slackUserIds.length) {
    throw new Error('対象ユーザーが選択されていません');
  }
  const message = '請求書チャンネルのおまとめ投稿へのリアクションが未確認です。内容を確認し、問題なければリアクションをお願いします。';
  sendDirectmessages(message, slackUserIds);
}

function findRowById(sheet, id, shouldThrow = true) {
  if (sheet.getLastRow() < 2) {
    if (shouldThrow) {
      throw new Error(`ID ${id} が見つかりません`);
    }
    return -1;
  }
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return i + 2;
    }
  }
  if (shouldThrow) {
    throw new Error(`ID ${id} が見つかりません`);
  }
  return -1;
}

function normalizeAssignedUserIdsForAdmin(value) {
  const ids = String(value || '').split(',').map(item => item.trim()).filter(Boolean);
  if (!ids.length) {
    return String(value || '').trim() === NO_ASSIGNED_USERS ? NO_ASSIGNED_USERS : '';
  }
  if (ids.length === 1 && ids[0] === NO_ASSIGNED_USERS) {
    return NO_ASSIGNED_USERS;
  }
  const users = sheetRowsToObjects(getDatabaseSheet(0), ['id','attendanceSheetId','slackUserId','driveFolderId','addressSheetId','email','passwordHash','savedData','name','deletedAt'])
    .filter(user => !user.deletedAt);
  const selectedIdSet = {};
  ids.forEach(id => selectedIdSet[String(id)] = true);
  const normalized = users.map(user => String(user.id)).filter(id => selectedIdSet[id]);
  return normalized.length === users.length ? '' : normalized.join(',');
}

// 警告タグの生成
function alert(text) {
  return `<p class="text-danger">${text}</p>`;
}

// 日時を Googleフォーム形式に変換 (yyyy/MM/dd HH:mm:ss)
function formatTimeToGoogleForm(datetime) {
  const time = new Date(datetime);
  const formattedDateTime = `${time.getFullYear()}/${(time.getMonth() + 1).toString().padStart(2, '0')}/${time.getDate().toString().padStart(2, '0')} ${time.getHours().toString().padStart(2, '0')}:${time.getMinutes().toString().padStart(2, '0')}:${time.getSeconds().toString().padStart(2, '0')}`;
  return formattedDateTime;
}

// HTML特殊文字をエスケープする
function sanitizeInput(input) {
  if (!input) return '';
  return input.replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ハッシュ化
function hashWithSHA256(input) {
  let rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  // ハッシュ値を16進数形式に変換
  let hash = rawHash.map(function(byte) {
    let hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return ('00' + hex).slice(-2);
  }).join('');
  return hash;
}

/**
 * パスワードハッシュ生成用のヘルパー関数（データベース設定時用）
 * @param {string} password - 平文パスワード
 * @returns {string} - SHA256ハッシュ値（16進数文字列）
 */
function generatePasswordHash(password) {
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  const hashHex = hash.map(byte => {
    return (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0');
  }).join('');
  
  console.log(`パスワード "${password}" のハッシュ値: ${hashHex}`);
  return hashHex;
}
