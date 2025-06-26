/**
 * =================================================================
 * 営業進捗ステータス自動更新スクリプト (GAS版) v6.1
 * =================================================================
 * ActionFlowシートのカラム名(action_name)とコード内の参照名(action_id)の
 * 不一致によってステータス更新が失敗していた不具合を修正。
 *
 * 【v6.1での主な変更点】
 * - `getNextStatus`関数を修正し、ActionFlowシートの`action_name`カラムを
 * 正しく参照するように変更。
 * - データ比較時に前後の空白を除去する処理を追加し、より堅牢に。
 * =================================================================
 */

// =================================================================
// 定数宣言
// =================================================================
const PM_CONSTANTS = {
  TABLE: {
    SALES_ACTION: 'SalesAction',
    BUSINESS_CARD: 'BusinessCard',
    ACTION_FLOW: 'ActionFlow',
  },
  COLUMN: {
    // SalesAction
    ACTION_ID: 'id',
    BUSINESS_CARD_ID: 'business_card_id',
    EXECUTED_DT: 'executed_dt',
    PROGRESS: 'progress',
    ACTION_NAME: 'action_name',
    RESULT: 'result',
    // BusinessCard
    CUSTOMER_ID: 'id',
    ENABLED_CONTACT: 'enabled_contact',
    PROGRESS_STATUS: 'progress_status',
    // ActionFlow
    FLOW_PROGRESS: 'progress',
    FLOW_ACTION_NAME: 'action_name', // ★ v6.1 修正：カラム名を action_id から action_name に
    FLOW_RESULT: 'result',
    FLOW_NEXT_STATUS: 'next_status',
  },
  PROPS_KEY: {
    APPSHEET_APP_ID: 'APPSHEET_APP_ID',
    APPSHEET_API_KEY: 'APPSHEET_API_KEY',
    MASTER_SHEET_ID: 'MASTER_SHEET_ID',
  }
};


/**
 * 【テスト用関数】固定のIDを使ってステータス更新をテストします。
 */
function test_updateCustomerStatus() {
  // ▼▼▼ テスト用に書き換えてください ▼▼▼
  const TEST_ACTION_ID = 'YOUR_TEST_SALESACTION_ID'; // テストしたいSalesActionレコードのID
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
 * @param {string} triggerActionId - トリガーとなったSalesActionのレコードID。
 * @param {string} execUserEmail - アクションを実行するユーザーのメールアドレス。
 */
function updateCustomerStatus(triggerActionId, execUserEmail) {
  (async () => {
    if (!triggerActionId || !execUserEmail) {
      Logger.log("必要な引数（triggerActionId, execUserEmail）が不足しています。");
      return;
    }
    
    try {
      const statusUpdater = new StatusUpdater(execUserEmail);
      
      const triggerAction = await statusUpdater.findRecordById(PM_CONSTANTS.TABLE.SALES_ACTION, triggerActionId);
      if (!triggerAction) {
        throw new Error(`ID [${triggerActionId}] のSalesActionレコードが見つかりません。`);
      }

      const businessCardId = triggerAction[PM_CONSTANTS.COLUMN.BUSINESS_CARD_ID];
      if(!businessCardId) {
        throw new Error(`アクション [${triggerActionId}] に顧客ID(${PM_CONSTANTS.COLUMN.BUSINESS_CARD_ID})が紐付いていません。`);
      }

      const customerRecord = await statusUpdater.findRecordById(PM_CONSTANTS.TABLE.BUSINESS_CARD, businessCardId);
      if (!customerRecord) {
        throw new Error(`ID [${businessCardId}] の顧客レコードが見つかりません。`);
      }

      if (String(customerRecord[PM_CONSTANTS.COLUMN.ENABLED_CONTACT]).toUpperCase() === 'FALSE') {
        Logger.log(`顧客 [${businessCardId}] はアプローチ対象外のため、ステータスを「対象外」に更新します。`);
        await statusUpdater.updateCustomerRecord(businessCardId, { [PM_CONSTANTS.COLUMN.PROGRESS_STATUS]: "対象外" });
        return;
      }
      
      const latestAction = await statusUpdater.findLatestActionForCustomer(businessCardId);
      if (!latestAction) {
        Logger.log(`顧客 [${businessCardId}] に紐づくアクションが見つからないため、処理を終了します。`);
        return;
      }
      Logger.log(`顧客[${businessCardId}]の最新アクションを取得しました (ID: ${latestAction[PM_CONSTANTS.COLUMN.ACTION_ID]})`);

      const { [PM_CONSTANTS.COLUMN.PROGRESS]: progress, [PM_CONSTANTS.COLUMN.ACTION_NAME]: action_name, [PM_CONSTANTS.COLUMN.RESULT]: result } = latestAction;
      
      if (progress === undefined || action_name === undefined) {
        Logger.log(`エラー：最新のアクションレコードに 'progress' または 'action_name' が含まれていません。`);
        Logger.log(`取得データ: ${JSON.stringify(latestAction)}`);
        return;
      }
      
      if (!result) {
        Logger.log(`最新アクション [${latestAction[PM_CONSTANTS.COLUMN.ACTION_ID]}] に結果が設定されていないため、ステータス更新をスキップします。`);
        return;
      }

      const nextStatus = statusUpdater.getNextStatus(progress, action_name, result);

      if (nextStatus) {
        Logger.log(`顧客 [${businessCardId}] のステータスを [${nextStatus}] に更新します。`);
        await statusUpdater.updateCustomerRecord(businessCardId, { [PM_CONSTANTS.COLUMN.PROGRESS_STATUS]: nextStatus });
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
    
    const appId = this.props[PM_CONSTANTS.PROPS_KEY.APPSHEET_APP_ID];
    const apiKey = this.props[PM_CONSTANTS.PROPS_KEY.APPSHEET_API_KEY];
    if (!appId || !apiKey) throw new Error("AppSheetのIDまたはAPIキーが設定されていません。");
    this.client = new AppSheetClient(appId, apiKey);

    const masterSheetId = this.props[PM_CONSTANTS.PROPS_KEY.MASTER_SHEET_ID];
    if (!masterSheetId) throw new Error("マスターシートのIDが設定されていません。");
    
    this.salesFlows = this._loadSheetData(masterSheetId, PM_CONSTANTS.TABLE.ACTION_FLOW);
  }

  _loadSheetData(sheetId, sheetName) {
    try {
      const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
      const [headers, ...rows] = sheet.getDataRange().getValues();
      return rows.map(row => headers.reduce((obj, header, i) => (obj[String(header).trim()] = String(row[i]).trim(), obj), {}));
    } catch (e) {
      throw new Error(`マスターシート(ID: ${sheetId}, Name: ${sheetName})の読み込みに失敗しました。: ${e.message}`);
    }
  }

  /**
   * 指定されたテーブルから特定のIDを持つレコードを検索します。
   */
  async findRecordById(tableName, recordId) {
    const keyColumn = (tableName === PM_CONSTANTS.TABLE.BUSINESS_CARD) ? PM_CONSTANTS.COLUMN.CUSTOMER_ID : PM_CONSTANTS.COLUMN.ACTION_ID;
    const selector = `FILTER("${tableName}", [${keyColumn}] = "${recordId}")`;
    const result = await this.client.findData(tableName, this.execUserEmail, { "Selector": selector });
    return (result && result.length > 0) ? result[0] : null;
  }
  
  /**
   * 特定の顧客に紐づく、実行日時が最も新しいSalesActionレコードを取得します。
   */
  async findLatestActionForCustomer(businessCardId) {
    const selector = `FILTER("${PM_CONSTANTS.TABLE.SALES_ACTION}", [${PM_CONSTANTS.COLUMN.BUSINESS_CARD_ID}] = "${businessCardId}")`;
    const allActions = await this.client.findData(PM_CONSTANTS.TABLE.SALES_ACTION, this.execUserEmail, { "Selector": selector });

    if (!allActions || allActions.length === 0) {
      return null;
    }

    const sortedActions = allActions.sort((a, b) => {
        const dateA = new Date(a[PM_CONSTANTS.COLUMN.EXECUTED_DT] || 0);
        const dateB = new Date(b[PM_CONSTANTS.COLUMN.EXECUTED_DT] || 0);
        return dateB - dateA;
    });

    return sortedActions[0];
  }

  /**
   * 次に取るべきステータスを計算します。
   */
  getNextStatus(currentProgress, currentActionName, currentResult) {
    // ★ v6.1 修正点: `action_id` を `action_name` に修正し、trim()で空白を除去
    const flow = this.salesFlows.find(row =>
      row[PM_CONSTANTS.COLUMN.FLOW_PROGRESS] === currentProgress &&
      row[PM_CONSTANTS.COLUMN.FLOW_ACTION_NAME] === currentActionName &&
      row[PM_CONSTANTS.COLUMN.FLOW_RESULT] === currentResult
    );
    return flow ? flow[PM_CONSTANTS.COLUMN.FLOW_NEXT_STATUS] : null;
  }

  /**
   * 顧客レコードのステータスを更新します。
   */
  async updateCustomerRecord(customerId, fieldsToUpdate) {
    const recordData = { [PM_CONSTANTS.COLUMN.CUSTOMER_ID]: customerId, ...fieldsToUpdate };
    return await this.client.updateRecords(PM_CONSTANTS.TABLE.BUSINESS_CARD, [recordData], this.execUserEmail);
  }
}
