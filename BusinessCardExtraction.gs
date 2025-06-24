/**
 * =================================================================
 * Business Card Extraction (v19.0.0)
 * =================================================================
 * v18をベースに、BusinessCardテーブルから`corporateNumber`カラムが
 * 削除された仕様変更に対応しました。
 *
 * 【v19.0.0での主な変更点】
 * - AIへの抽出指示プロンプトから`corporateNumber`を削除。
 * - AppSheetのAccountテーブル、BusinessCardテーブルを更新する際の
 * ペイロードから`corporateNumber`を削除し、エラーを解消。
 * =================================================================
 */

// =================================================================
// ▼▼▼ 設定項目 (ご利用前に必ず設定してください) ▼▼▼
// =================================================================

/** @const {string} APPSHEET_APP_ID - あなたのAppSheetアプリのID */
const APPSHEET_APP_ID = PropertiesService.getScriptProperties().getProperty('APPSHEET_APP_ID') || 'YOUR_APPSHEET_APP_ID';

/** @const {string} APPSHEET_KEY_COLUMN_NAME - BusinessCardテーブルの主キー列の名前 */
const APPSHEET_KEY_COLUMN_NAME = 'id';

/** @const {string} APPSHEET_IMAGE_FOLDER_NAME - AppSheetがファイルを参照する際に使用するフォルダ名 (例: "テーブル名.Images") */
const APPSHEET_IMAGE_FOLDER_NAME = 'FileContents';

/** @const {string} APPSHEET_UPLOAD_FOLDER_ID - AppSheetがファイルをアップロードするGoogle DriveフォルダのID */
const APPSHEET_UPLOAD_FOLDER_ID = '1ihmiqnPVl7DSdNw5tQ9Z-XiKizmdUXkL';

/** @const {string} IMAGE_DESTINATION_FOLDER_ID - PDFから変換した画像を保存するGoogle DriveフォルダのID */
const IMAGE_DESTINATION_FOLDER_ID = APPSHEET_UPLOAD_FOLDER_ID;

/** @const {string} SCHEMA_SPREADSHEET_ID - プロンプトのSchemaを定義しているGoogleスプレッドシートのID */
const SCHEMA_SPREADSHEET_ID = '129J3rU9h1QRU6dtvmRnElRPzhKHVmZy0JQaiHPAMLhs';

/** @const {string} SCHEMA_SHEET_NAME - Schemaが定義されているシートの名前 */
const SCHEMA_SHEET_NAME = 'Schema';

/** @const {string} GEMINI_MODEL_NAME - データ抽出に使用するGeminiのモデル名 */
const GEMINI_MODEL_NAME = 'gemini-1.5-flash-latest';


// --- APIキー関連 (スクリプトプロパティでの設定を強く推奨します) ---
const APPSHEET_API_KEY = PropertiesService.getScriptProperties().getProperty('APPSHEET_API_KEY') || 'YOUR_APPSHEET_APIKEY';
const CLOUDCONVERT_API_KEY = PropertiesService.getScriptProperties().getProperty('CLOUDCONVERT_API_KEY') || 'YOUR_CLOUDCONVERT_APIKEY';


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
* **URL Format:** For the 'link1' and 'link2' fields, you must output the full URL including 'https://' or 'http://'. Do not output just the domain name.
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
 * @returns {Promise<{status: string, data: (Object|undefined), message: (string|undefined)}>} 処理結果を示すオブジェクト
 */
async function runBusinessCardExtraction(tableName, recordId, fileNameFront, fileNameBack, userEmail) {
  const MAX_RETRIES = 3;
  let lastError;
  const execUser = userEmail || Session.getActiveUser().getEmail() || 'appsheet-owner@example.com'; 

  // 処理全体を最大3回まで試行
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      console.log(`Attempt ${attempt} for recordId: ${recordId}`);
      
      const fileIdFront = findFileIdByName(APPSHEET_UPLOAD_FOLDER_ID, fileNameFront);
      const fileIdBack = fileNameBack ? findFileIdByName(APPSHEET_UPLOAD_FOLDER_ID, fileNameBack) : null;
      const fileIds = [fileIdFront, fileIdBack].filter(id => id); 

      if (!tableName || !recordId || fileIds.length === 0) {
        throw new Error('tableName, recordIdが指定されていないか、指定されたファイル名が見つかりません。');
      }

      const result = await processBusinessCard(tableName, fileIds, recordId, execUser);
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
    await appSheetClient.updateRecords(tableName, [recordToUpdate], execUser);
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
 * @returns {Promise<{updatedData: Object}>} 更新されたデータを含むオブジェクト
 */
async function processBusinessCard(tableName, fileIds, recordId, execUser) {
  const destinationFolder = DriveApp.getFolderById(IMAGE_DESTINATION_FOLDER_ID);
  const processedImageFiles = [];

  // --- 1. ファイル処理 (PDF -> PNG変換) ---
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
  const geminiResponse = await geminiClient.generateCandidates();
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

  extractedData.link1 = _formatUrl(extractedData.link1);
  extractedData.link2 = _formatUrl(extractedData.link2);


  // --- 3. アカウントの検索または作成 ---
  const appSheetClient = new AppSheetClient(APPSHEET_APP_ID, APPSHEET_API_KEY);
  const accountId = await _findOrCreateAccount(appSheetClient, extractedData, execUser);
  
  // --- 4. 緯度・経度取得 (Google Maps Service) ---
  let location = { latitude: '', longitude: '' };
  if (extractedData.address) {
      location = getLatLngFromAddress_(extractedData.address);
      console.log(`緯度経度取得完了: lat=${location.latitude}, lng=${location.longitude}`);
  }

  // --- 5. レコード更新 (AppSheet API) ---
  console.log('AppSheetのレコード更新を開始します...');
  const image_path_front = processedImageFiles[0] ? `${APPSHEET_IMAGE_FOLDER_NAME}/${processedImageFiles[0].getName()}` : "";
  const image_path_back = processedImageFiles[1] ? `${APPSHEET_IMAGE_FOLDER_NAME}/${processedImageFiles[1].getName()}` : "";
  const now = new Date();

  // ★★★ 修正点: 更新ペイロードから corporateNumber を削除 ★★★
  const { corporateNumber, ...restOfExtractedData } = extractedData;

  // 更新用データオブジェクトを作成
  const recordToUpdate = {
    ...restOfExtractedData,
    [APPSHEET_KEY_COLUMN_NAME]: recordId,
    account_id: accountId, 
    latitude: location.latitude,
    longitude: location.longitude,
    acquisitionDatetime: Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss'),
    image_path_front: image_path_front,
    image_path_back: image_path_back,
    importStatus: '完了', // ステータスを完了に
  };
  if (recordToUpdate.gender === '不明') {
    recordToUpdate.gender = '';
  }

  await appSheetClient.updateRecords(tableName, [recordToUpdate], execUser);
  console.log('レコード更新が完了しました。');
  return { updatedData: recordToUpdate };
}


// =================================================================
// ヘルパー関数 (Helper Functions)
// =================================================================

/**
 * 会社名またはメールドメインを基にAccountテーブルを検索し、存在しない場合は基本的な情報で新規作成する。
 * @param {AppSheetClient} client - AppSheetClientのインスタンス。
 * @param {Object} cardData - 名刺から抽出したデータ。
 * @param {string} userEmail - 実行ユーザーのメールアドレス。
 * @returns {Promise<string|null>} - 見つかった、または作成されたAccountのID。
 * @private
 */
async function _findOrCreateAccount(client, cardData, userEmail) {
  const companyName = cardData.companyName;
  const domain = _extractDomainFromEmail(cardData.email);

  if (!companyName && !domain) {
    console.warn("会社名と有効なメールドメインの両方がないため、アカウント連携をスキップします。");
    return null;
  }
  
  try {
    // 1. ドメインで既存アカウントを検索 (最優先)
    if (domain) {
      console.log(`ドメインでアカウントを検索中: ${domain}`);
      const domainSelector = `FILTER("Account", [domain] = "${domain}")`;
      const accountsByDomain = await client.findData("Account", userEmail, { "Selector": domainSelector });
      if (accountsByDomain && accountsByDomain.length > 0) {
        const existingAccountId = accountsByDomain[0].id;
        console.log(`ドメインで既存のアカウントが見つかりました: ${existingAccountId}`);
        return existingAccountId;
      }
    }

    // 2. 会社名で既存アカウントを検索 (フォールバック)
    if (companyName) {
      console.log(`会社名でアカウントを検索中: ${companyName}`);
      const nameSelector = `FILTER("Account", [name] = "${companyName}")`;
      const accountsByName = await client.findData("Account", userEmail, { "Selector": nameSelector });
      if (accountsByName && accountsByName.length > 0) {
        const existingAccountId = accountsByName[0].id;
        console.log(`会社名で既存のアカウントが見つかりました: ${existingAccountId}`);
        return existingAccountId;
      }
    }

    // 3. 新規アカウントを作成 (企業情報収集は行わない)
    console.log(`新規アカウントを作成します: ${companyName || domain}`);
    
    const newAccountPayload = {
      name: companyName,
      domain: domain, 
      postal_code: cardData.postalCode,
      address: cardData.address,
      building_name: cardData.buildingName,
      email: cardData.email,
      phone_number: cardData.phoneNumber,
      fax_number: cardData.faxNumber,
      note: cardData.otherNotes,
      // ★★★ 修正点: corporateNumberをペイロードから削除 ★★★
      // corporate_number: cardData.corporateNumber, 
      website_url: cardData.link1,
      enrichment_status: 'Pending'
    };
    
    const addResponse = await client.addRecords("Account", [newAccountPayload], userEmail);
    
    if (addResponse && addResponse.Rows && addResponse.Rows.length > 0 && addResponse.Rows[0].id) {
        const newAccountId = addResponse.Rows[0].id;
        console.log(`新規アカウントを作成しました: ${newAccountId}`);
        return newAccountId;
    } else {
        throw new Error("新規アカウントの作成に成功しましたが、応答からIDを取得できませんでした。");
    }
  } catch (e) {
    console.error(`アカウントの検索または作成中にエラーが発生しました: ${e.stack}`);
    // エラーを再スローして、上位のcatchで処理させる
    throw e;
  }
}

/**
 * URLを適切な形式に整形するヘルパー関数。
 */
function _formatUrl(urlString) {
  if (!urlString || typeof urlString !== 'string' || urlString.trim() === '') {
    return null;
  }
  let trimmedUrl = urlString.trim();
  if (!/^https?:\/\//i.test(trimmedUrl)) {
    trimmedUrl = `https://${trimmedUrl}`;
  }
  try {
    new URL(trimmedUrl);
    return trimmedUrl;
  } catch (_) {
    Logger.log(`無効なURL形式のためスキップします: ${urlString}`);
    return null;
  }
}


/**
 * メールアドレスからドメインを抽出する。
 */
function _extractDomainFromEmail(email) {
  if (!email || !email.includes('@')) return null;
  
  const domain = email.split('@')[1];
  const freeEmailDomains = [
    'gmail.com', 'yahoo.co.jp', 'yahoo.com', 'hotmail.com', 'outlook.jp', 'outlook.com', 
    'icloud.com', 'me.com', 'mac.com', 'aol.com', 'excite.co.jp'
  ];

  if (freeEmailDomains.includes(domain.toLowerCase())) {
    return null;
  }
  
  return domain;
}


/**
 * 住所文字列から緯度・経度を取得する
 */
function getLatLngFromAddress_(address) {
  try {
    const geocoder = Maps.newGeocoder().setLanguage('ja');
    const response = geocoder.geocode(address);
    
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
 */
function findFileIdByName(folderId, fileName) {
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
 * スプレッドシートからSchema定義を読み込む。
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
    // ★★★ 修正点: corporateNumberをスキーマから除外 ★★★
    const key = row[0];
    if (key && key !== 'corporateNumber') {
      schema[key] = { label: row[1], description: row[2], format: row[3] };
    }
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
  
  runBusinessCardExtraction(tableName, recordId, fileNameFront, fileNameBack, userEmail);
}
