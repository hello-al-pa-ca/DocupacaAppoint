/**
 * =================================================================
 * 営業進捗ステータス自動更新スクリプト (GAS版) v5
 * =================================================================
 * AppSheetのVirtual Columnの代わりに、GASを使って顧客の営業ステータスを
 * 実カラムに書き込みます。これにより、AppSheetアプリのパフォーマンスを改善します。
 *
 * 【主な変更点】
 * - 【v5での修正】AppSheet APIライブラリの非同期処理に対応するため、
 * async/awaitを正しく再導入しました。これにより、APIからのデータ取得を
 * 確実に行うようになり、「レコードが見つかりません」エラーを解消します。
 *
 * 【処理の流れ】
 * 1. AppSheetの `SalesAction` テーブルでレコードが追加・更新される。
 * 2. AppSheetのAutomationがこのスクリプトの `updateCustomerStatus` 関数を呼び出す。
 * 3. GASが `SalesAction` レコードから顧客IDを取得し、顧客の `enabled_contact` 列を確認。
 * 4. GASが `SalesActionFlow` のルールを参照して、次のステータスを計算する。
 * 5. GASが顧客テーブル（`BusinessCard`）の `progress_status` 列を更新する。
 *
 * 【AppSheetでの設定方法】
 * (設定方法は以前のバージョンから変更ありません)
 * =================================================================
 */

/**
 * 【テスト用関数】固定のIDを使ってステータス更新をテストします。
 */
function test_updateCustomerStatus() {
  // ▼▼▼ テスト用に書き換えてください ▼▼▼
  const TEST_ACTION_ID = '7FBCF696-7397-49A3-BC8C-7E5E3AB3AAB4'; // テストしたいSalesActionレコードのID
  const TEST_EXEC_USER = 'hello@al-pa-ca.com';       // 実行ユーザーのメールアドレス
  // ▲▲▲

  if (TEST_ACTION_ID === 'YOUR_TEST_SALESACTION_ID') {
    Logger.log("テストを実行するには、TEST_ACTION_IDを実際の値に設定してください。");
    return;
  }
  
  Logger.log(`テスト開始: Action ID [${TEST_ACTION_ID}]`);
  updateCustomerStatus(TEST_ACTION_ID, TEST_EXEC_USER);
  Logger.log("テスト終了");
}


/**
 * 【AppSheetから呼び出す関数】顧客のステータスを更新します。
 * @param {string} actionId - トリガーとなったSalesActionのレコードID。
 * @param {string} execUserEmail - アクションを実行するユーザーのメールアドレス。
 */
function updateCustomerStatus(actionId, execUserEmail) {
  // async/awaitをトップレベルで安全に扱うため、無名関数でラップして実行します。
  (async () => {
    if (!actionId || !execUserEmail) {
      Logger.log("必要な引数（actionId, execUserEmail）が不足しています。");
      return;
    }
    
    try {
      const statusUpdater = new StatusUpdater(execUserEmail);
      
      // トリガーとなったアクション情報を取得（awaitで応答を待つ）
      const latestAction = await statusUpdater.findRecordById('SalesAction', actionId, 'id');
      if (!latestAction) {
        throw new Error(`ID [${actionId}] のSalesActionレコードが見つかりません。`);
      }

      const businessCardId = latestAction.business_card_id; // アクションレコードから親の顧客IDを取得
      if(!businessCardId) {
        throw new Error(`アクション [${actionId}] に顧客ID(business_card_id)が紐付いていません。`);
      }

      // 顧客情報を取得してアプローチ対象か確認（awaitで応答を待つ）
      const customerRecord = await statusUpdater.findRecordById('BusinessCard', businessCardId, 'id');
      if (!customerRecord) {
        throw new Error(`ID [${businessCardId}] の顧客レコードが見つかりません。`);
      }

      // `enabled_contact`列がFALSEかどうかをチェック
      const isContactEnabled = (String(customerRecord.enabled_contact).toUpperCase() !== 'FALSE');

      if (!isContactEnabled) {
        Logger.log(`顧客 [${businessCardId}] はアプローチ対象外(enabled_contact=FALSE)のため、ステータスを「対象外」に更新します。`);
        await statusUpdater.updateCustomerRecord(businessCardId, { "progress_status": "対象外" });
        return; // アプローチ対象外ならここで処理を終了
      }
      
      const { progress, action_name, result } = latestAction;
      
      if (progress === undefined || action_name === undefined) {
        Logger.log(`エラー：取得したアクションレコードに 'progress' または 'action_name' が含まれていません。AppSheetのテーブル定義を確認してください。`);
        Logger.log(`取得データ: ${JSON.stringify(latestAction)}`);
        return;
      }
      
      if (!result) {
        Logger.log(`アクション [${actionId}] に結果が設定されていないため、ステータス更新をスキップします。`);
        return;
      }

      const nextStatus = statusUpdater.getNextStatus(progress, action_name, result);

      if (nextStatus) {
        Logger.log(`顧客 [${businessCardId}] のステータスを [${nextStatus}] に更新します。`);
        await statusUpdater.updateCustomerRecord(businessCardId, { "progress_status": nextStatus });
      } else {
        Logger.log(`次のステータスが見つかりませんでした。Flowを確認してください。 (progress: ${progress}, action: ${action_name}, result: ${result})`);
      }

    } catch (e) {
      Logger.log(`❌ ステータス更新中にエラーが発生しました: ${e.message}\n${e.stack}`);
    }
  })().catch(e => Logger.log(`❌ 非同期処理の実行中にエラー: ${e.message}`));
}


/**
 * @class StatusUpdater
 * @description 営業ステータスの計算と更新ロジックを管理するクラス。
 */
class StatusUpdater {
  constructor(execUserEmail) {
    if (!execUserEmail) {
      throw new Error("実行ユーザーのメールアドレスは必須です。");
    }
    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.client = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);

    const masterSheetId = this.props.MASTER_SHEET_ID;
    if (!masterSheetId) throw new Error("マスターシートのIDがスクリプトプロパティに設定されていません。");
    
    this.salesFlows = this._loadSheetData(masterSheetId, 'ActionFlow');
  }

  _loadSheetData(sheetId, sheetName) {
    try {
      const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
      const [headers, ...rows] = sheet.getDataRange().getValues();
      return rows.map(row => headers.reduce((obj, header, i) => (obj[header] = row[i], obj), {}));
    } catch (e) {
      throw new Error(`マスターシート(ID: ${sheetId}, Name: ${sheetName})の読み込みに失敗しました。: ${e.message}`);
    }
  }

  /**
   * 指定されたテーブルから特定のIDを持つレコードを検索します。
   * @param {string} tableName - 検索するテーブル名。
   * @param {string} recordId - 検索するレコードのID。
   * @param {string} [keyColumn='id'] - 検索に使用するキー列の名前。
   * @returns {Promise<Object|null>} - 見つかったレコードオブジェクト、またはnull。
   */
  async findRecordById(tableName, recordId, keyColumn = 'id') {
    const selector = `FILTER("${tableName}", [${keyColumn}] = "${recordId}")`;
    const result = await this.client.findData(tableName, this.execUserEmail, { "Selector": selector });
    return (result && result.length > 0) ? result[0] : null;
  }

  getNextStatus(currentProgress, currentActionName, currentResult) {
    // SalesActionFlowのaction_id列はアクション名が入っていると仮定
    const flow = this.salesFlows.find(row =>
      row.progress === currentProgress &&
      row.action_id === currentActionName &&
      row.result === currentResult
    );
    return flow ? flow.next_status : null;
  }

  /**
   * 顧客レコードを更新します。
   * @param {string} customerId - 更新する顧客のID。
   * @param {Object} fieldsToUpdate - 更新するフィールドのキーと値。
   * @returns {Promise<Object>} - APIからのレスポンス。
   */
  async updateCustomerRecord(customerId, fieldsToUpdate) {
    const CUSTOMER_TABLE_NAME = 'BusinessCard'; 
    const CUSTOMER_KEY_COLUMN = 'id';

    const recordData = { [CUSTOMER_KEY_COLUMN]: customerId, ...fieldsToUpdate };
    return await this.client.updateRecords(CUSTOMER_TABLE_NAME, [recordData], this.execUserEmail);
  }
}
