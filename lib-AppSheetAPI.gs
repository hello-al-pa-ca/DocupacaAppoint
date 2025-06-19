/**
 * =================================================================
 * Refactored AppSheet Client Code
 * =================================================================
 * * 機能性を維持しつつ、エラーハンドリングとAPIの正確性を向上させました。
 * - `APIRequest.request`: レスポンスをHTTPステータスコードで検証し、
 * JSONレスポンスを自動的にパースするように修正。
 * - `AppSheetClient.getData`: AppSheet APIの仕様に基づき、
 * 'GET'メソッドを'POST'メソッドに修正。
 * - 全体的にJSDocコメントを追加し、可読性を向上。
 */


/**
 * =================================================================
 * Usage Examples (個別実行サンプル)
 * =================================================================
 * * 以下の各関数は、操作ごとに独立して実行できます。
 * * 実際に使用する際は、プレースホルダ（YOUR_...）を実際の値に置き換えてください。
 * * Google Apps Scriptエディタで、実行したい関数（例: `example_addRecords`）を選択して実行します。
 */

// --- ▼▼▼ レコードを追加する (Add) ▼▼▼ ---
async function example_addRecords() {
  const APP_ID = 'YOUR_APP_ID';         // あなたのAppSheetアプリID
  const API_KEY = 'YOUR_API_KEY';       // あなたのAppSheet APIキー
  const EXEC_USER = 'user@example.com'; // 実行ユーザーのメールアドレス
  const TABLE_NAME = 'Users';           // 操作対象のテーブル名

  const client = new AppSheetClient(APP_ID, API_KEY);

  // 追加したいレコードのデータを作成します。複数同時に追加も可能です。
  const newRecords = [
    { "Name": "Taro Yamada", "Age": 35, "Email": "taro.yamada@example.com" },
    { "Name": "Hanako Suzuki", "Age": 28, "Email": "hanako.suzuki@example.com" }
  ];

  try {
    console.log(`Adding ${newRecords.length} record(s) to "${TABLE_NAME}" table...`);
    const result = await client.addRecords(TABLE_NAME, newRecords, EXEC_USER);
    console.log('Add Result:', JSON.stringify(result, null, 2));
  } catch (error) {
    console.error('Failed to add records:', error);
  }
}


// --- ▼▼▼ レコードを更新する (Update) ▼▼▼ ---
async function example_updateRecords() {
  const APP_ID = 'YOUR_APP_ID';
  const API_KEY = 'YOUR_API_KEY';
  const EXEC_USER = 'user@example.com';
  const TABLE_NAME = 'Users';

  const client = new AppSheetClient(APP_ID, API_KEY);

  // 更新したいレコードのデータを作成します。
  // AppSheetテーブルのキー列（例: "Email"）と、更新したい列の値を指定します。
  const recordsToUpdate = [
    { "Email": "taro.yamada@example.com", "Age": 36 } // taro.yamadaの年齢を36に更新
  ];

  try {
    console.log(`Updating ${recordsToUpdate.length} record(s) in "${TABLE_NAME}" table...`);
    const result = await client.updateRecords(TABLE_NAME, recordsToUpdate, EXEC_USER);
    console.log('Update Result:', JSON.stringify(result, null, 2));
  } catch (error) {
    console.error('Failed to update records:', error);
  }
}


// --- ▼▼▼ レコードを削除する (Delete) ▼▼▼ ---
async function example_deleteRecords() {
  const APP_ID = 'YOUR_APP_ID';
  const API_KEY = 'YOUR_API_KEY';
  const EXEC_USER = 'user@example.com';
  const TABLE_NAME = 'Users';

  const client = new AppSheetClient(APP_ID, API_KEY);

  // 削除したいレコードのキーを指定します。
  const recordsToDelete = [
    { "Email": "hanako.suzuki@example.com" } // hanako.suzukiのレコードを削除
  ];

  try {
    console.log(`Deleting ${recordsToDelete.length} record(s) from "${TABLE_NAME}" table...`);
    const result = await client.deleteRecords(TABLE_NAME, recordsToDelete, EXEC_USER);
    console.log('Delete Result:', JSON.stringify(result, null, 2));
  } catch (error) {
    console.error('Failed to delete records:', error);
  }
}



/**
 * @class APIRequest
 * @description APIリクエストを処理するための汎用クラス。
 */
class APIRequest {
  /**
   * @param {string} baseUrl - APIのベースURL。
   * @param {Object} [headers={}] - 全てのリクエストに適用される共通ヘッダー。
   */
  constructor(baseUrl, headers = {}) {
    this.baseUrl = baseUrl;
    this.headers = headers;
  }

  /**
   * APIにリクエストを送信します。
   * @param {string} endpoint - リクエストのエンドポイント。
   * @param {string} method - HTTPメソッド ('get', 'post', 'put', 'delete'など)。
   * @param {Object|null} [data=null] - 送信するデータ（ペイロード）。
   * @returns {Promise<Object|string>} - APIからのレスポンス。JSONの場合はパースされたオブジェクト、それ以外はテキストを返します。
   * @throws {Error} - APIリクエストが失敗した場合（HTTPステータスコードが200番台でない場合）。
   */
  async request(endpoint, method, data = null) {
    const url = `${this.baseUrl}${endpoint}`;
    
    const params = {
      contentType: 'application/json',
      headers: this.headers,
      method: method.toLowerCase(),
      muteHttpExceptions: true // HTTPエラーを例外としてスローせず、レスポンスを直接処理する
    };

    if (data) {
      params.payload = JSON.stringify(data);
    }

    try {
      const response = UrlFetchApp.fetch(url, params);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      // HTTPステータスコードが2xxでない場合はエラーをスロー
      if (responseCode < 200 || responseCode >= 300) {
        throw new Error(`API Error: Received status code ${responseCode}. Response: ${responseText}`);
      }
      
      // レスポンスが空の場合でも成功として扱うことがあるため、空文字列チェックは削除
      if (responseText === "") {
        return "Success (No Content)";
      }

      // JSONとしてパースを試みる
      try {
        return JSON.parse(responseText);
      } catch (e) {
        // パースに失敗した場合はテキストをそのまま返す
        return responseText;
      }

    } catch (error) {
      console.error(`Request failed: ${error.message}`);
      throw error; // エラーを再スローして呼び出し元で処理できるようにする
    }
  }
}

/**
 * @class AppSheetClient
 * @description AppSheet APIと対話するためのクライアントクラス。
 * @extends APIRequest
 */
class AppSheetClient extends APIRequest {
  /**
   * @param {string} appId - AppSheetのアプリケーションID。
   * @param {string} apiKey - AppSheetのアプリケーションアクセスキー。
   * @throws {Error} - appIdまたはapiKeyが指定されていない場合。
   */
  constructor(appId, apiKey) {
    if (!appId || !apiKey) {
      throw new Error('appId and apiKey are required.');
    }
    
    const baseUrl = `https://api.appsheet.com/api/v2/apps/${appId}/tables`;
    const headers = {
      'ApplicationAccessKey': apiKey
    };
    
    super(baseUrl, headers);
    
    this._appId = appId;
    this._apiKey = apiKey;
  }

  get appId() {
    return this._appId;
  }

  get apiKey() {
    return this._apiKey;
  }

  /**
   * AppSheetのアクション（Add, Edit, Delete, etc.）を実行するための共通メソッド。
   * @private
   * @param {string} tableName - 対象のテーブル名。
   * @param {string} action - 実行するアクション名 ('Add', 'Edit', 'Delete'など)。
   * @param {Array<Object>} rows - 操作対象の行データ。
   * @param {string} executionUser - アクションを実行するユーザーのメールアドレス。
   * @returns {Promise<Object|string>} - APIからのレスポンス。
   */
  async executeAction(tableName, action, rows, executionUser) {
    const endpoint = `/${encodeURIComponent(tableName)}/Action`;
    const payload = {
      Action: action,
      Properties: {
        "Locale": "ja-JP",
        "Timezone": "Tokyo Standard Time",
        "RunAsUserEmail": executionUser,
      },
      Rows: rows
    };

    return this.request(endpoint, 'post', payload);
  }
  
  /**
   * テーブルに新しいレコードを追加します。
   * @param {string} tableName - 対象のテーブル名。
   * @param {Array<Object>} recordData - 追加するレコードのデータ。
   * @param {string} executionUser - 実行ユーザーのメールアドレス。
   * @returns {Promise<Object|string>} - APIからのレスポンス。
   */
  async addRecords(tableName, recordData, executionUser) {
    return this.executeAction(tableName, 'Add', recordData, executionUser);
  }

  /**
   * 既存のレコードを更新します。
   * @param {string} tableName - 対象のテーブル名。
   * @param {Array<Object>} recordData - 更新するレコードのデータ（キーを含む）。
   * @param {string} executionUser - 実行ユーザーのメールアドレス。
   * @returns {Promise<Object|string>} - APIからのレスポンス。
   */
  async updateRecords(tableName, recordData, executionUser) {
    return this.executeAction(tableName, 'Edit', recordData, executionUser);
  }

  /**
   * 既存のレコードを削除します。
   * @param {string} tableName - 対象のテーブル名。
   * @param {Array<Object>} deleteData - 削除するレコードのキー。
   * @param {string} executionUser - 実行ユーザーのメールアドレス。
   * @returns {Promise<Object|string>} - APIからのレスポンス。
   */
  async deleteRecords(tableName, deleteData, executionUser) {
    return this.executeAction(tableName, 'Delete', deleteData, executionUser);
  }

  /**
   * 特定のレコードに対してカスタムアクションを呼び出します。
   * @param {string} tableName - 対象のテーブル名。
   * @param {Array<Object>} targets - アクション対象のレコードキー。
   * @param {string} actionName - 呼び出すカスタムアクションの名前。
   * @param {string} executionUser - 実行ユーザーのメールアドレス。
   * @returns {Promise<Object|string>} - APIからのレスポンス。
   */
  async callAction(tableName, targets, actionName, executionUser) {
    return this.executeAction(tableName, actionName, targets, executionUser);
  }

  /**
   * テーブルからデータを取得します (Findアクションを使用)。
   * @param {string} tableName - 対象のテーブル名。
   * @param {string} executionUser - 実行ユーザーのメールアドレス。
   * @param {Object} [properties={}] - Selectorなどの追加プロパティ。
   * @returns {Promise<Object|string>} - 取得したデータ。
   */
  async findData(tableName, executionUser, properties = {}) {
    const endpoint = `/${encodeURIComponent(tableName)}/Action`;
    const payload = {
      Action: 'Find',
      Properties: {
        "Locale": "ja-JP",
        "Timezone": "Tokyo Standard Time",
        "RunAsUserEmail": executionUser,
        ...properties
      },
      Rows: []
    };
    // AppSheetのFindアクションはPOSTリクエストを使用します。
    return this.request(endpoint, 'post', payload);
  }
}


