// /* /**
//  * 
//  */
// function getGDriveFiles(fileName){
//   const searchFiles = DriveApp.searchFiles(`title contains '${fileName}'`);
//   let files = [];
//   while(searchFiles.hasNext()){
//     let file = searchFiles.next()
//     files.push(file);
//     console.log(file.getName());
//   }
//   return files;
// }

// /**
//  * URLパスから「/」以降の文字列を取得する
//  * @param {string} path - 文字列パス（例: "test/sample.jpg_dasd"）
//  * @return {string} 「/」以降の文字列（例: "sample.jpg_dasd"）
//  */
// function getPathAfterTarget(path, target) {
//    return path.indexOf(target) !== -1 ? path.substring(path.indexOf(target) + 1) : path;
// }

// /**
//  * 指定のスプレッドシートのシートのx行目まで固定する
//  */
// function setFrozenRowAndColumns(ss, fixRow = 1, fixColumn = 1) {
//    // すべてのシートを取得する
//   const sheets = ss.getSheets();
//   // シートを1つずつ処理していく
//   for (let i = 0; i < sheets.length; i++) {
//     // 行方向の固定(0なら解除)
//     sheets[i].setFrozenRows(fixRow);
//     // 列方向の固定(0なら解除)
//     sheets[i].setFrozenColumns(fixColumn);
//   }
// }

// /**
//  * クエリ文字列をオブジェクトに変換
//  * @param {string} queryString - クエリ文字列
//  * @returns {Object} パース済みのクエリパラメータ
//  */
// function parseQueryString(queryString) {
//   const params = {};
//   if (!queryString) {
//     return params;
//   }
  
//   queryString.split('&').forEach(function(pair) {
//     const [key, value] = pair.split('=');
//     if (key && value) {
//       params[decodeURIComponent(key)] = decodeURIComponent(value);
//     }
//   });
  
//   return params;
// }


// /**
//  * エラーレスポンスの作成
//  * @param {string} message - エラーメッセージ
//  * @param {number} code - HTTPステータスコード
//  * @param {Object} headers - レスポンスヘッダー
//  * @returns {GoogleAppsScript.Content.TextOutput} エラーレスポンス
//  */
// function createErrorResponse(message, code, headers) {
//   const timestamp = new Date().toISOString();
  
//   return JSON.stringify({
//     status: "error",
//     error: {
//       code: code,
//       message: message,
//       timestamp: timestamp,
//       path: ScriptApp.getService().getUrl()
//     }
//   })
// }