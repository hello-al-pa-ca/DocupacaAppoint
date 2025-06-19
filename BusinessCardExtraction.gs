/**
 * @fileoverview AppSheetからトリガーされる名刺情報抽出・更新バックエンド (緯度経度取得対応)
 * @version 1.3.0
 * * @description
 * AppSheetから直接このスクリプトの'runBusinessCardExtraction'関数を呼び出して使用します。
 * 抽出した住所を元に、Google Mapsサービスを利用して正確な緯度・経度を取得し、AppSheetのレコードを更新します。
 * * @see
 * - AIへの指示(プロンプト)は、管理しやすいようにGoogleスプレッドシートから動的に読み込みます。
 * - 処理中にエラーが発生した場合、3回まで自動でリトライします。
 * - 最終的に失敗した場合は、AppSheetのステータスを更新して処理を終了します。
 */

// =================================================================
// ▼▼▼ 設定項目 (ご利用前に必ず設定してください) ▼▼▼
// =================================================================

/**
 * @const {string} APPSHEET_APP_ID - あなたのAppSheetアプリのID
 */
const APPSHEET_APP_ID = PropertiesService.getScriptProperties().getProperty('APPSHEET_APP_ID') || 'APPSHEET_APP_ID';

/**
 * @const {string} APPSHEET_KEY_COLUMN_NAME - AppSheetテーブルの主キーとして使用する列の名前
 */
const APPSHEET_KEY_COLUMN_NAME = 'id';

/**
 * @const {string} APPSHEET_IMAGE_FOLDER_NAME - AppSheetがファイルを参照する際に使用するフォルダ名 (例: "テーブル名.Images")
 */
const APPSHEET_IMAGE_FOLDER_NAME = 'FileContents';

/**
 * @const {string} APPSHEET_UPLOAD_FOLDER_ID - AppSheetがファイルをアップロードするGoogle DriveフォルダのID
 */
const APPSHEET_UPLOAD_FOLDER_ID = '1ihmiqnPVl7DSdNw5tQ9Z-XiKizmdUXkL';

/**
 * @const {string} IMAGE_DESTINATION_FOLDER_ID - PDFから変換した画像を保存するGoogle DriveフォルダのID
 */
const IMAGE_DESTINATION_FOLDER_ID = APPSHEET_UPLOAD_FOLDER_ID;

/**
 * @const {string} SCHEMA_SPREADSHEET_ID - プロンプトのSchemaを定義しているGoogleスプレッドシートのID
 */
const SCHEMA_SPREADSHEET_ID = '129J3rU9h1QRU6dtvmRnElRPzhKHVmZy0JQaiHPAMLhs';

/**
 * @const {string} SCHEMA_SHEET_NAME - Schemaが定義されているシートの名前
 */
const SCHEMA_SHEET_NAME = 'Schema';

/**
 * @const {string} GEMINI_MODEL_NAME - データ抽出に使用するGeminiのモデル名
 */
const GEMINI_MODEL_NAME = 'gemini-2.0-flash';


// --- APIキー関連 (スクリプトプロパティでの設定を強く推奨します) ---
const APPSHEET_API_KEY = PropertiesService.getScriptProperties().getProperty('APPSHEET_API_KEY') || 'YOUR_APPSHEET_APIKEY';
const CLOUDCONVERT_API_KEY = PropertiesService.getScriptProperties().getProperty('CLOUDCONVERT_API_KEY') || 'YOUR_CLOUDCONVERT_APIKEY';
// Gemini APIキーは 'GOOGLE_API_KEY' という名前でスクリプトプロパティに設定してください。


// --- Gemini プロンプト定義 (英語) ---
const PROMPT_SYSTEM_INSTRUCTION = `You are an AI specialized in generating structured data by extracting information from text, images, and PDFs.`;
const PROMPT_USER_INSTRUCTION = `Follow the rules and the provided Schema below to extract information from the 1-2 attached business card images and output a single JSON object.

**Rules:**
* **Schema:** Strictly adhere to the attached Schema definition for each field.
* **Prefixes in 'description':**
  - **(Infer):** If the description starts with \`(Infer)\`, output a value based on context, even if not explicitly written on the image.
  - **(Select):** If the description starts with \`(Select)\`, you must choose the most appropriate value from the list provided in square brackets \`[A, B, C]\`.
  - **No Prefix:** Extract the information as it is written on the image.
* **format:** If a 'format' is specified in the Schema, strictly follow it (e.g.,<x_bin_880>/MM/dd, regex).
* **Multiple Images:** If two images are provided, they represent the front and back of the business card. Consolidate information from both images to provide the most complete and accurate data.
* **Non-existent Fields:** If you cannot find the information for a field, output an empty string \`""\`.
* **Output Format:** You must output **only a single, raw JSON object** and nothing else. Do not include any explanatory text or markdown formatting like \`\`\`json.

**Output Example:**
{
  "companyName": "Google Japan G.K.",
  "name": "Taro Yamada",
  "nameKana": "ヤマダ タロウ",
  "position": "Software Engineer",
  "department": "Cloud Platform",
  "gender": "男性",
  "postalCode": "150-0002",
  "address": "Shibuya Stream, 3-21-3 Shibuya, Shibuya-ku, Tokyo",
  "buildingName": "Shibuya Stream",
  "email": "taro.yamada@example.com",
  "phoneNumber": "03-1234-5678",
  "mobileNumber": "",
  "faxNumber": "",
  "link1": "https://cloud.google.com/",
  "link2": "",
  "corporateNumber": "1011001089234",
  "ocrFront": "Google Japan G.K. Software Engineer Taro Yamada ...",
  "ocrBack": "",
  "otherNotes": "Met at the tech conference."
}`;


// =================================================================
// メイン実行関数 (Entry Point)
// =================================================================

/**
 * AppSheetから直接呼び出すためのメイン関数。3回までのリトライ処理を含む。
 * @param {string} tableName - AppSheetで更新対象のテーブル名
 * @param {string} recordId - AppSheetのレコードのユニークID (主キーの値)
 * @param {string} fileNameFront - 表面のファイル名
 * @param {string} [fileNameBack] - 裏面のファイル名 (任意)
 * @param {string} [userEmail] - アクションを実行するユーザーのメールアドレス (任意)
 * @returns {{status: string, data: (Object|undefined), message: (string|undefined)}} 処理結果を示すオブジェクト
 */
function runBusinessCardExtraction(tableName, recordId, fileNameFront, fileNameBack, userEmail) {
  const MAX_RETRIES = 3;
  let lastError;
  const execUser = userEmail || 'user@example.com'; 

  // 処理全体を最大3回まで試行
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      console.log(`Attempt ${attempt} for recordId: ${recordId}`);
      
      const fileIdFront = findFileIdByName_(APPSHEET_UPLOAD_FOLDER_ID, fileNameFront);
      const fileIdBack = fileNameBack ? findFileIdByName_(APPSHEET_UPLOAD_FOLDER_ID, fileNameBack) : null;
      const fileIds = [fileIdFront, fileIdBack].filter(id => id); 

      if (!tableName || !recordId || fileIds.length === 0) {
        throw new Error('tableName, recordIdが指定されていないか、指定されたファイル名が見つかりません。');
      }

      const result = processBusinessCard(tableName, fileIds, recordId, execUser);
      console.log(`Attempt ${attempt} succeeded.`);
      return { status: 'success', data: result };

    } catch (error) {
      console.error(`Attempt ${attempt} failed: ${error.stack}`);
      lastError = error;
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(2000);
      }
    }
  }

  // 全てのリトライが失敗した場合、AppSheetのステータスを「失敗」で更新
  console.error(`All ${MAX_RETRIES} attempts failed for recordId: ${recordId}. Updating status to 'failed'.`);
  try {
    const appSheetClient = new AppSheetClient(APPSHEET_APP_ID, APPSHEET_API_KEY);
    const recordToUpdate = {
      [APPSHEET_KEY_COLUMN_NAME]: recordId,
      importStatus: '解析に失敗しました。再抽出してください。'
    };
    appSheetClient.updateRecords(tableName, [recordToUpdate], execUser);
    console.log(`Successfully updated recordId ${recordId} with failure status.`);
    return { status: 'error', message: `All attempts failed. Final error: ${lastError.message}` };
  } catch (updateError) {
    console.error(`Fatal: Failed to update AppSheet with failure status for recordId ${recordId}: ${updateError.stack}`);
    return { status: 'error', message: `All attempts failed, and could not update failure status. Final error: ${lastError.message}` };
  }
}

// =================================================================
// コアロジック (Core Logic)
// =================================================================

/**
 * 名刺ファイルの処理とデータ更新の主要な流れを管理する
 * @param {string} tableName - AppSheetのテーブル名
 * @param {string[]} fileIds - 処理対象のファイルIDの配列
 * @param {string} recordId - AppSheetのレコードID
 * @param {string} execUser - 実行ユーザーのメールアドレス
 * @returns {{updatedData: Object}} 更新されたデータを含むオブジェクト
 */
function processBusinessCard(tableName, fileIds, recordId, execUser) {
  const destinationFolder = DriveApp.getFolderById(IMAGE_DESTINATION_FOLDER_ID);
  const processedImageFiles = [];

  // --- 1. ファイル処理 ---
  fileIds.forEach((fileId, index) => {
    const originalFile = DriveApp.getFileById(fileId);
    let processedImageFile;
    if (originalFile.getMimeType() === MimeType.PDF) {
      const ccClient = new CloudConvertClient(CLOUDCONVERT_API_KEY);
      const convertedFileIds = ccClient.convertPdfFromDriveToPng(fileId, IMAGE_DESTINATION_FOLDER_ID);
      if (!convertedFileIds || convertedFileIds.length === 0) throw new Error('CloudConvertでのファイル変換に失敗しました。');
      processedImageFile = DriveApp.getFileById(convertedFileIds[0]);
    } else {
      processedImageFile = originalFile;
    }
    processedImageFiles.push(processedImageFile);
    console.log(`ファイル(${index + 1})処理完了: ${processedImageFile.getName()}`);
  });

  // --- 2. データ抽出 (Gemini API) ---
  const geminiClient = new GeminiClient(GEMINI_MODEL_NAME);
  const schemaForPrompt = getSchemaFromSheet();
  const fullUserPrompt = `${PROMPT_USER_INSTRUCTION}\n\nSchema:\n${JSON.stringify(schemaForPrompt, null, 2)}`;
  geminiClient.setSystemInstructionText(PROMPT_SYSTEM_INSTRUCTION);
  processedImageFiles.forEach(file => geminiClient.attachFiles(file.getBlob()));
  geminiClient.setPromptText(fullUserPrompt);
  const geminiResponse = geminiClient.generateCandidates();
  if (!geminiResponse.candidates || geminiResponse.candidates.length === 0) {
    throw new Error(`Geminiからの応答がありませんでした。レスポンス: ${JSON.stringify(geminiResponse)}`);
  }
  const textResponse = geminiResponse.candidates[0].content.parts[0].text;
  let extractedData;
  try {
    extractedData = JSON.parse(textResponse.replace(/^```json\s*|```\s*$/g, ''));
  } catch (e) {
    throw new Error(`Geminiからの応答がJSON形式ではありません: ${textResponse}`);
  }
  console.log('データ抽出完了:', extractedData);
  
  // --- 3. 緯度・経度取得 (Google Maps Service) ---
  let location = { latitude: '', longitude: '' };
  if (extractedData.address) {
      location = getLatLngFromAddress_(extractedData.address);
      console.log(`緯度経度取得完了: lat=${location.latitude}, lng=${location.longitude}`);
  }

  // --- 4. レコード更新 (AppSheet API) ---
  console.log('AppSheetのレコード更新を開始します...');
  const appSheetClient = new AppSheetClient(APPSHEET_APP_ID, APPSHEET_API_KEY);
  const image_path_front = processedImageFiles[0] ? `${APPSHEET_IMAGE_FOLDER_NAME}/${processedImageFiles[0].getName()}` : "";
  const image_path_back = processedImageFiles[1] ? `${APPSHEET_IMAGE_FOLDER_NAME}/${processedImageFiles[1].getName()}` : "";
  const now = new Date();

  // 更新用データオブジェクトを作成
  const recordToUpdate = {
    ...extractedData,
    [APPSHEET_KEY_COLUMN_NAME]: recordId,
    latitude: location.latitude,
    longitude: location.longitude,
    acquisitionDatetime: Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss'),
    image_path_front: image_path_front,
    image_path_back: image_path_back,
    importStatus: '',
  };
  if (recordToUpdate.gender === '不明') {
    recordToUpdate.gender = '';
  }

  appSheetClient.updateRecords(tableName, [recordToUpdate], execUser);
  console.log('レコード更新が完了しました。');
  return { updatedData: recordToUpdate };
}


// =================================================================
// ヘルパー関数 (Helper Functions)
// =================================================================

/**
 * 住所文字列から緯度・経度を取得する
 * @param {string} address - 変換したい住所
 * @returns {{latitude: number|string, longitude: number|string}} 緯度経度オブジェクト
 * @private
 */
function getLatLngFromAddress_(address) {
  try {
    // Mapsのジオコーダーを呼び出し
    const geocoder = Maps.newGeocoder().setLanguage('ja');
    const response = geocoder.geocode(address);
    
    // 結果が存在し、ステータスがOKの場合のみ値を返す
    if (response && response.status === 'OK' && response.results && response.results.length > 0) {
      const location = response.results[0].geometry.location;
      return {
        latitude: location.lat,
        longitude: location.lng
      };
    } else {
      console.warn(`住所から緯度経度を取得できませんでした: ${address}, Status: ${response.status}`);
      return { latitude: '', longitude: '' };
    }
  } catch(e) {
    console.error(`getLatLngFromAddress_でエラーが発生しました: ${e}`);
    return { latitude: '', longitude: '' };
  }
}

/**
 * 指定されたフォルダ内でファイル名からファイルIDを検索します。
 * @private
 */
function findFileIdByName_(folderId, fileName) {
  if (!folderId || !fileName) return null;
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      if (files.hasNext()) console.warn(`フォルダID'${folderId}'内に'${fileName}'という名前のファイルが複数見つかりました。最初のファイルを使用します。`);
      return file.getId();
    }
    console.error(`フォルダID'${folderId}'内でファイル'${fileName}'が見つかりませんでした。`);
    return null;
  } catch (e) {
    console.error(`フォルダID'${folderId}'の検索中にエラーが発生しました: ${e.message}`);
    return null;
  }
}

/**
 * スプレッドシートからSchema定義を読み込む。パフォーマンス向上のためキャッシュを利用。
 */
function getSchemaFromSheet() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'prompt_schema_cache';
  const cachedSchema = cache.get(cacheKey);
  if (cachedSchema) {
    console.log('Schemaをキャッシュから読み込みました。');
    return JSON.parse(cachedSchema);
  }
  console.log('キャッシュが存在しないため、スプレッドシートからSchemaを読み込みます。');
  const ss = SpreadsheetApp.openById(SCHEMA_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SCHEMA_SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const schema = {};
  data.forEach(row => {
    const key = row[0], label = row[1], description = row[2], format = row[3];
    if (key) schema[key] = { label, description, format };
  });
  cache.put(cacheKey, JSON.stringify(schema), 21600);
  return schema;
}


/** テスト用の関数 */
function testProcess() {
  const tableName = 'BusinessCard';
  const recordId = '8E67005D-AED0-4A62-9D58-90167A26C1C1';
  const fileNameFront = '8E67005D-AED0-4A62-9D58-90167A26C1C1.image_path_front.205556.jpg';
  const fileNameBack = '';
  const userEmail = 'hello@al-pa-ca.com';
  
  const result = runBusinessCardExtraction(tableName, recordId, fileNameFront, fileNameBack, userEmail);
  console.log(JSON.stringify(result, null, 2));
}


