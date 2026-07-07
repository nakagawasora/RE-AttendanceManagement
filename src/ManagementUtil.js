const scriptProperties = PropertiesService.getScriptProperties();

const token = scriptProperties.getProperty("SLACK_BOT_TOKEN");
const billChannelId = scriptProperties.getProperty("BILL_CHANNEL_ID");
const slackIdOfAdmin = scriptProperties.getProperty("SLACK_ID_OF_ADMIN");
const mailAddress = scriptProperties.getProperty("MAIL_ADDRESS");
const ccRecipient = scriptProperties.getProperty("CC_RECIPIENT");
const billFolderId = scriptProperties.getProperty("BILL_FOLDER_ID");

/**
 * 先月の請求書を設置してスラックで通知
 * @param {number} year - 実行年（デバッグ用）
 * @param {number} month - 実行月（デバッグ用）
 * @param {number} date - 実行日（デバッグ用）
*/
function setBill(year = new Date().getFullYear(),month = new Date().getMonth() + 1,date = new Date().getDate()) {
  const lastMonthLastDay = new Date(year,month - 1,0);//-1により見かけ上の集計月からDate型の集計月へ
  const databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0];
  const idArray = databaseSheet.getRange('B2:E'+databaseSheet.getLastRow()).getValues();
  let genarateCounter = 0;
  for (let i = 0; i < idArray.length; i++){
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
    let message = `<!channel>\n${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月分の請求書が出力されました。\n成功:${genarateCounter}件\n失敗:${idArray.length - genarateCounter}件\n訂正の必要がなければ当メッセージにリアクションを、何か問題があれば返信してください。`;
    sendMessage(message,billChannelId);
  }
  catch(e){
    GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月分の請求書発行おまとめメッセージ送信に際し、Slackでのメッセージ送信に失敗しました。\nエラーメッセージ:${e.message}\n以下の文章かぎ括弧内をコピーして請求書チャンネルに貼り付けてください。\n「@channel\n${lastMonthLastDay.getFullYear()}年${lastMonthLastDay.getMonth()+1}月分の請求書が出力されました。\n成功:${genarateCounter}件\n失敗:${idArray.length - genarateCounter}件\n訂正の必要がなければ当メッセージにリアクションを、何か問題があれば返信してください。」`,{cc:ccRecipient});   
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
  const databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0];
  let i = BillUtil.search(databaseSheet,attendanceSheetId)[0] - 2;
  const idArray = databaseSheet.getRange('B2:E'+databaseSheet.getLastRow()).getValues();
  let generatedDict;
  //請求書出力
  try{
    generatedDict = BillUtil.generateBill(idArray[i][0],idArray[i][2],lastMonthLastDay.getMonth()+1,lastMonthLastDay.getFullYear(),year,month,date);
  }
  catch(e){
    generatedDict = {status:e.message,slackUserId:idArray[i][1],aggregationYear:lastMonthLastDay.getFullYear(),aggregationMonth:lastMonthLastDay.getMonth()+1,billUrl:''};
  }
  //結果をスラックで連絡、送信エラーが出ればメールで連絡
  try{
    slackReport(generatedDict,true);
  }
  catch(e){
    try{
      let name = SpreadsheetApp.openById(idArray[i][3]).getSheets()[0].getRange('B3:C3').getValues();
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
  const databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0];
  const idArray = databaseSheet.getRange('B2:E'+databaseSheet.getLastRow()).getValues();
  let failureCounter = 0;
  for(let i = 0; i < idArray.length; i++){
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
      sendMessage(`${idArray.length - failureCounter}件の請求書回収に成功し、${failureCounter}件の請求書回収に失敗しました。\n請求書フォルダを確認してください。`,billChannelId);
    }
    catch(e){
      GmailApp.sendEmail(mailAddress,'【勤怠管理】Slackの送信に失敗しました',`${idArray.length - failureCounter}件の請求書回収に成功し、${failureCounter}件の請求書回収に失敗したことを伝えるSlackでのメッセージ送信に失敗しました。請求書フォルダを確認してください。\nエラーメッセージ:${e.message}`,{ name: 'ManagementUtil',cc:ccRecipient });
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
function getMonthlyWorkTimeAndSalaryWithAuth(email, password, targetMonth = null, targetYear = null) {
  try {
    // 認証処理
    const authResult = authenticateUser(email, password);
    if (authResult.status !== 'success') {
      return authResult;
    }
    
    const attendanceSheetId = authResult.attendanceSheetId;
    
    // 対象月・年の設定（省略時は現在の月・年）
    const currentDate = new Date();
    const month = targetMonth !== null ? targetMonth : currentDate.getMonth() + 1;
    const year = targetYear !== null ? targetYear : currentDate.getFullYear();
    
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
    for (let hourlyWage in hourlyWageDict) {
      if (hourlyWageDict[hourlyWage] !== 0) {
        const baseAmount = Math.round(Number(hourlyWage) * hourlyWageDict[hourlyWage] / 3600000);
        totalAmount += Math.round(baseAmount * 1.1); // 10%税込み
      }
    }
    
    return {
      status: 'success',
      data: {
        workTimeByProject: workTimeByProject,
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
    const userSheet = databaseSheet.getSheets()[0];
    
    // 全データを取得（F列：メール、G列：パスワードハッシュ、B列：attendanceSheetId）
    const lastRow = userSheet.getLastRow();
    if (lastRow < 2) {
      return {
        status: 'error',
        message: 'ユーザーデータが存在しません'
      };
    }
    
    const userData = userSheet.getRange(`F2:G${lastRow}`).getValues();
    const attendanceSheetIds = userSheet.getRange(`B2:B${lastRow}`).getValues();
    
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
      const userEmail = userData[i][0];
      const storedPasswordHash = userData[i][1];
      
      if (userEmail === email && storedPasswordHash === passwordHashHex) {
        const attendanceSheetId = attendanceSheetIds[i][0];
        
        if (!attendanceSheetId) {
          return {
            status: 'error',
            message: '勤怠シートIDが設定されていません'
          };
        }
        
        return {
          status: 'success',
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

/**
 * 月次勤務実績API用GETハンドラー
 * URL: /exec?email=xxx@example.com&password=password&month=xx&year=xxxx
 */
function handleMonthlyWorkTimeApi(e) {
  try {
    const params = e.parameter;
    const email = params.email;
    const password = params.password;
    
    if (!email || !password) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'error',
          message: 'emailとpasswordパラメータが必要です'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const month = params.month ? parseInt(params.month) : null;
    const year = params.year ? parseInt(params.year) : null;
    
    const result = getMonthlyWorkTimeAndSalaryWithAuth(email, password, month, year);
    
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

  if (params.email && params.password) {
    return handleMonthlyWorkTimeApi(e);
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();
  template.formHTML = createHtml(e);
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('勤怠入力');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

// POSTメソッド
function doPost(e) {
  let databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId);

  // 勤怠提出
  if (e.parameter.submit) {
    // バリデーション
    let alertContent = '';

    // 入力メールアドレスがデータベース内に存在するか
    let passOfEmail = BillUtil.search(databaseSheet.getSheets()[0], e.parameter.email);

    // 存在する場合
    if (passOfEmail[0] > 0) {
      // 入力されたパスワードのハッシュ化した値がメアドの隣のハッシュパスに一致するか
      if (!(databaseSheet.getSheets()[0].getRange('F2:G' + databaseSheet.getSheets()[0].getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password))) {
        alertContent += alert('・アカウント情報が不正です');
      } else {
        // パスワード認証成功時のバリデーション
        alertContent += validateFormData(e);

        // 二重勤怠チェック（認証成功時のみ）
        if (alertContent === '' && e.parameter.startTime && e.parameter.endTime) {
          alertContent += checkDuplicateAttendance(e, passOfEmail[0]);
        }
      }
    } else {
      alertContent += alert('・アカウント情報が不正です');
    }

    // 基本的なフォームバリデーション（認証失敗時でも実行）
    if (!passOfEmail[0] || passOfEmail[0] <= 0 || !(databaseSheet.getSheets()[0].getRange('F2:G' + databaseSheet.getSheets()[0].getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password))) {
      alertContent += validateFormData(e);
    }

    // 警告内容がなければ書き込んで完了ページへ、あれば警告とともに再表示
    if (alertContent === '') {
      const attendance = [[formatTimeToGoogleForm(new Date()), e.parameter.project, e.parameter.content, formatTimeToGoogleForm(e.parameter.startTime), formatTimeToGoogleForm(e.parameter.endTime)]];
      const sheet = SpreadsheetApp.openById(SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0].getRange('B' + passOfEmail[0]).getValue()).getSheets()[0];
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
    let passOfEmail = BillUtil.search(databaseSheet.getSheets()[0], e.parameter.email);

    // アカウント認証チェック
    let authenticationPassed = false;
    if (passOfEmail[0] > 0) {
      if (databaseSheet.getSheets()[0].getRange('F2:G' + databaseSheet.getSheets()[0].getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password)) {
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
      databaseSheet.getSheets()[0].getRange('H' + passOfEmail[0]).setValue(JSON.stringify(saveData));
    }

    // 画面の再表示（認証成功・失敗に関わらず）
    const template = HtmlService.createTemplateFromFile('index');
    template.deployURL = ScriptApp.getService().getUrl();
    template.formHTML = createHtml(e, alertContent);
    const htmlOutput = template.evaluate();
    htmlOutput.setTitle('勤怠入力');
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return htmlOutput;
  }

  // 勤怠呼び出し
  if (e.parameter.load) {
    let databaseSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId);
    let alertContent = '';
    let passOfEmail = BillUtil.search(SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0], e.parameter.email);

    if (passOfEmail[0] > 0) {
      if (!(databaseSheet.getSheets()[0].getRange('F2:G' + databaseSheet.getSheets()[0].getLastRow()).getValues()[passOfEmail[0] - 2][1] === hashWithSHA256(e.parameter.password))) {
        alertContent += alert('・アカウント情報が不正です');
      }
      // 一度も保存がなされていないとき
      if (SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0].getRange('H' + passOfEmail[0]).getValue() === '') {
        alertContent += alert('・データが保存されていません');
      }
    } else {
      alertContent += alert('・アカウント情報が不正です');
    }

    // アカウント情報が正しければ保存データを呼び出し
    if (alertContent === '') {
      alertContent += alert('Loaded!!');
      const savedDataJson = SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0].getRange('H' + passOfEmail[0]).getValue();

      // 保存データを復元して現在のパスワードと結合
      const loaded = createLoadedEvent(savedDataJson, e.parameter.email, e.parameter.password);

      const template = HtmlService.createTemplateFromFile('index');
      template.deployURL = ScriptApp.getService().getUrl();
      template.formHTML = createHtml(loaded, alertContent);
      const htmlOutput = template.evaluate();
      htmlOutput.setTitle('勤怠入力');
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    } else {
      const template = HtmlService.createTemplateFromFile('index');
      template.deployURL = ScriptApp.getService().getUrl();
      template.formHTML = createHtml(e, alertContent);
      const htmlOutput = template.evaluate();
      htmlOutput.setTitle('勤怠入力');
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return htmlOutput;
    }
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();
  template.formHTML = createHtml(e, alert('・不正な操作です'));
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

// 二重勤怠チェック関数
function checkDuplicateAttendance(e, passOfEmailIndex) {
  let alertContent = '';

  const attendanceSheet = SpreadsheetApp.openById(SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[0].getRange('B' + passOfEmailIndex).getValue()).getSheets()[0];

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

  const projectsSheet = SpreadsheetApp.openById(BillUtil.databaseSheetId).getSheets()[1];
  const projectArray = projectsSheet.getRange('B2:B' + projectsSheet.getLastRow()).getValues();
  for (let i = 0; i < projectArray.length; i++) {
    const projectName = sanitizeInput(projectArray[i][0] || '');
    let selected = '';
    if (project === projectName) {
      selected = 'selected';
      pulldown += `<option value="${projectName}" ${selected}>${projectName}</option>`;
    } else {
      pulldown += `<option value="${projectName}">${projectName}</option>`;
    }
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
                    <select class="form-select" name="project" value="${project}" autocomplete="off">
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
