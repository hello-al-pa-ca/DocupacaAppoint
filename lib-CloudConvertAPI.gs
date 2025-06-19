/**
 * =================================================================
 * CloudConvert API Client (Revised based on working code)
 * =================================================================
 * * CloudConvert API v2 との連携をカプセル化し、再利用しやすくしたクライアントクラスです。
 * * 動作確認が取れた、より信頼性の高い個別タスク実行方式を採用しています。
 */

/**
 * @class CloudConvertClient
 * @description CloudConvert APIと対話するためのクライアントクラス。
 */
class CloudConvertClient {
  /**
   * CloudConvertClientのコンストラクタ
   * @param {string} apiKey - CloudConvertのAPIキー。
   * @throws {Error} - APIキーが指定されていない場合にエラーをスローします。
   */
  constructor(apiKey) {
    if (!apiKey) {
      throw new Error('CloudConvert API key is required.');
    }
    /** @private */
    this.apiKey = apiKey;
    /** @private */
    this.baseUrl = 'https://api.cloudconvert.com/v2';
  }

  /**
   * APIリクエストを送信するための内部ヘルパーメソッド。
   * @private
   */
  _request(url, options) {
    try {
      options.muteHttpExceptions = true;
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode >= 400) {
        throw new Error(`CloudConvert API Error (HTTP ${responseCode}): ${responseText}`);
      }
      return responseText ? JSON.parse(responseText) : null;
    } catch (e) {
      console.error(`CloudConvert API request failed. URL: ${url}, Error: ${e.message}`);
      throw e;
    }
  }
  
  /**
   * 個別のタスクを作成するための共通メソッド。
   * @private
   */
  _createTask(operation, payload = {}) {
    const url = `${this.baseUrl}/${operation}`;
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    console.log(`Creating task: ${operation}`);
    return this._request(url, options);
  }

  /**
   * 指定されたタスクIDのステータスを取得します。
   * @private
   */
  _getTaskStatus(taskId) {
    const url = `${this.baseUrl}/tasks/${taskId}`;
    const options = {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${this.apiKey}` }
    };
    return this._request(url, options);
  }

  /**
   * 指定されたタスクが完了するまで同期的に待機します。
   * @param {string} taskId - 待機対象のタスクID。
   * @param {number} timeoutSeconds - タイムアウトまでの秒数（デフォルト: 300秒）。
   * @returns {Object} - 完了したタスクのデータ。
   */
  waitForTask(taskId, timeoutSeconds = 300) {
    const startTime = Date.now();
    while (Date.now() - startTime < timeoutSeconds * 1000) {
      const task = this._getTaskStatus(taskId);
      
      if (!task || !task.data) {
        throw new Error(`Failed to get status for task ${taskId}. Response was empty.`);
      }

      const status = task.data.status;
      console.log(`Task ${taskId} status: ${status}, percent: ${task.data.percent || 0}%`);
      
      if (status === 'finished') {
        console.log(`Task ${taskId} finished successfully.`);
        return task.data;
      }
      if (status === 'error') {
        throw new Error(`Task ${taskId} failed with code ${task.data.code}: ${task.data.message}`);
      }
      
      Utilities.sleep(3000); // 3秒待機
    }
    throw new Error(`Task ${taskId} timed out after ${timeoutSeconds} seconds.`);
  }

  /**
   * Google Drive上の単一のPDFファイルをPNGに変換し、指定されたGoogle Driveフォルダに保存します。
   * @param {string} sourceFileId - 変換元のPDFファイルのGoogle Drive ID。
   * @param {string} destinationFolderId - 保存先のGoogle DriveフォルダのID。
   * @returns {Array<string>} - 保存されたPNGファイルのGoogle Drive IDの配列。
   */
  async convertPdfFromDriveToPng(sourceFileId, destinationFolderId) {
    // 1. PDFファイルを取得
    const pdfFile = DriveApp.getFileById(sourceFileId);
    const pdfBlob = pdfFile.getBlob();
    const pdfName = pdfFile.getName();
    console.log(`--- Starting conversion for: ${pdfName} (ID: ${sourceFileId}) ---`);

    // 2. CloudConvertにファイルをアップロードするためのタスクを作成
    const uploadTaskResponse = this._createTask('import/upload');
    const uploadTask = uploadTaskResponse.data;

    // 3. 取得したURLにファイル本体をアップロード
    UrlFetchApp.fetch(uploadTask.result.form.url, {
      method: 'POST',
      payload: { ...uploadTask.result.form.parameters, file: pdfBlob }
    });
    console.log(`File uploaded successfully. Import task ID: ${uploadTask.id}`);
    
    // 4. 変換タスクを作成
    const convertTaskPayload = {
      input: uploadTask.id,
      input_format: 'pdf',
      output_format: 'png',
      engine: 'poppler',
      pages: '1-' // 全ページを変換
    };
    const convertTaskResponse = this._createTask('convert', convertTaskPayload);
    const convertTaskId = convertTaskResponse.data.id;
    
    // 5. 変換タスクの完了を待機
    await this.waitForTask(convertTaskId);
    
    // 6. エクスポート(URL取得)タスクを作成
    const exportTaskPayload = { input: convertTaskId };
    const exportTaskResponse = this._createTask('export/url', exportTaskPayload);
    const exportTaskId = exportTaskResponse.data.id;

    // 7. エクスポートタスクの完了を待機し、結果を取得
    const exportTaskResult = await this.waitForTask(exportTaskId);
    const convertedFiles = exportTaskResult.result.files;

    // 8. 変換された全ファイルをGoogle Driveに保存
    const outputFolder = DriveApp.getFolderById(destinationFolderId);
    const savedFileIds = [];
    const baseFilename = pdfName.replace(/\.pdf$/i, '');
    
    for (let i = 0; i < convertedFiles.length; i++) {
      const fileInfo = convertedFiles[i];
      
      const newFilename = (convertedFiles.length > 1)
        ? `${baseFilename}-${i + 1}.png`
        : `${baseFilename}.png`;

      console.log(`Downloading converted file from ${fileInfo.url}`);
      const fileBlob = UrlFetchApp.fetch(fileInfo.url).getBlob();
      fileBlob.setName(newFilename);
      
      const savedFile = outputFolder.createFile(fileBlob);
      savedFileIds.push(savedFile.getId());
      console.log(`Saved file to Drive: ${savedFile.getName()} (ID: ${savedFile.getId()})`);
    }

    console.log(`--- Conversion for ${pdfName} completed successfully! ---`);
    return savedFileIds;
  }
}

/**
 * =================================================================
 * 実行関数 (Functions for Execution)
 * =================================================================
 */

/**
 * 【メイン関数】複数のPDFファイルを一括でPNGに変換します。
 * @param {Array<string>} pdfFileIds - 変換したいPDFファイルのGoogle Drive IDの配列。
 * @param {string} destinationFolderId - PNG画像の保存先フォルダのID。
 * @returns {Array<Object>} 各ファイルの変換結果を格納した配列。
 */
async function convertPdfsInDriveToPngs(pdfFileIds, destinationFolderId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLOUDE_CONVERT_APIKEY");
  if (!apiKey) throw new Error("APIキーがスクリプトプロパティに設定されていません。");

  const client = new CloudConvertClient(apiKey);
  const allResults = [];

  for (const fileId of pdfFileIds) {
    try {
      const savedFileIds = await client.convertPdfFromDriveToPng(fileId, destinationFolderId);
      allResults.push({
        sourceFileId: fileId,
        status: 'Success',
        outputFileIds: savedFileIds
      });
    } catch (e) {
      const fileName = DriveApp.getFileById(fileId).getName();
      console.error(`Failed to convert ${fileName} (ID: ${fileId}): ${e.toString()}`);
      allResults.push({
        sourceFileId: fileId,
        status: 'Failed',
        error: e.toString()
      });
    }
  }
  return allResults;
}


/**
 * 【ヘルパー関数】指定したフォルダ内をファイル名で検索し、ファイルIDの配列を返します。
 * @param {string} fileName - 検索するファイル名。
 * @param {string} searchFolderId - 検索対象のフォルダID。
 * @returns {Array<string>} 見つかったファイルのIDの配列。
 */
function findFileIdsByName(fileName, searchFolderId) {
  const folder = DriveApp.getFolderById(searchFolderId);
  const files = folder.getFilesByName(fileName);
  const fileIds = [];
  while (files.hasNext()) {
    fileIds.push(files.next().getId());
  }
  console.log(`Found ${fileIds.length} file(s) named "${fileName}" in folder ID ${searchFolderId}.`);
  return fileIds;
}

/**
 * =================================================================
 * 使用例 (Usage Examples)
 * =================================================================
 */

/**
 * 【使用例】ファイル名（ユニークIDを含む）でPDFを検索し、変換する
 */
async function example_findByNameAndConvert() {
  // ▼▼▼ あなたの環境に合わせて、以下の値を書き換えてください ▼▼▼
  const SEARCH_FOLDER_ID = '1CzMYL4pWxbzzK4YDfPGBDFzeHy9Cxx2p'; // AppSheetがPDFをアップロードするフォルダID
  const FILENAME_TO_FIND = 'AIで名刺読み取り商談に、製造業出身の二人開発　アルパカ【TECH Biz LIVE!】：ニュース：中日BIZナビ.pdf';       // AppSheetが作成するユニークID付きのファイル名
  const DESTINATION_FOLDER_ID = '1CzMYL4pWxbzzK4YDfPGBDFzeHy9Cxx2p';   // 変換後のPNGを保存するフォルダID
  // ▲▲▲ ▲▲▲

  if (SEARCH_FOLDER_ID === 'YOUR_APPSHEET_UPLOAD_FOLDER_ID') {
    console.log("エラー: スクリプト内のSEARCH_FOLDER_IDなどの定数を、実際のGoogle DriveのIDに書き換えてください。");
    return;
  }

  // 1. ファイル名でPDFのIDを検索
  const fileIds = findFileIdsByName(FILENAME_TO_FIND, SEARCH_FOLDER_ID);

  if (fileIds.length > 0) {
    // 2. 見つかったファイルを一括変換関数に渡す
    const results = await convertPdfsInDriveToPngs(fileIds, DESTINATION_FOLDER_ID);
    
    console.log("\n--- CONVERSION SUMMARY ---");
    console.log(JSON.stringify(results, null, 2));
  } else {
    console.log(`File named "${FILENAME_TO_FIND}" not found in folder ${SEARCH_FOLDER_ID}.`);
  }
}
