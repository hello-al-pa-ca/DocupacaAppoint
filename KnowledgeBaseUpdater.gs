/**
 * =================================================================
 * Knowledge Base Updater for RAG (v5)
 * =================================================================
 * v4に加え、ご指定の仕様に合わせてスクリプトプロパティの参照名を変更しました。
 *
 * 【v5での主な変更点】
 * - 参照するプロジェクトIDのプロパティ名を `BIGQUERY_PROJECT_ID` から `GCP_PROJECT_ID` に変更。
 * - テスト関数の実行ユーザーもサービスアカウントを参照するように修正。
 * =================================================================
 */

// =================================================================
// 定数宣言
// =================================================================
const SUCCESS_RESULTS = ['受注', 'アポイント取得']; // 「成功」とみなすresult列の値
const BATCH_LIMIT = 50; // 一度に処理する最大件数

const BIGQUERY_DATASET_ID = 'rag_knowledge_base';
const BIGQUERY_TABLE_ID = 'knowledge_base';

// =================================================================
// グローバル関数 (時間主導型トリガーまたは手動で実行)
// =================================================================

/**
 * 【夜間バッチ実行用】ナレッジベースの更新プロセスを開始します。
 */
function runDailyKnowledgeBaseUpdate() {
  const execUserEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
  
  if (!execUserEmail) {
    Logger.log('❌ エラー: スクリプトプロパティに "SERVICE_ACCOUNT" が設定されていません。バッチ処理を実行できません。');
    return;
  }

  Logger.log(`ナレッジベース更新バッチをサービスアカウント (${execUserEmail}) で開始します。`);
  
  try {
    const updater = new KnowledgeBaseUpdater(execUserEmail);
    updater.updateKnowledgeBase()
      .catch(e => {
        Logger.log(`❌ ナレッジベース更新中に致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
      });
  } catch (e) {
    Logger.log(`❌ 初期化中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * BigQueryへのデータ取り込みをテストするための関数
 * AppSheetからのデータ取得をスキップし、ダミーデータを使ってBigQueryへの挿入をテストします。
 */
async function test_insertDummyDataToBigQuery() {
  // ★★★ 修正点: 実行ユーザーをサービスアカウントから取得 ★★★
  const execUserEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
  if (!execUserEmail) {
    Logger.log('❌ エラー: スクリプトプロパティに "SERVICE_ACCOUNT" が設定されていません。テストを実行できません。');
    return;
  }
  Logger.log(`BigQueryへのダミーデータ挿入テストをサービスアカウント (${execUserEmail}) で開始します。`);

  try {
    const updater = new KnowledgeBaseUpdater(execUserEmail);

    // 1. テスト用のダミーデータを作成
    const dummyAction = {
      ID: `test-action-${Utilities.getUuid()}`,
      accountId: `test-account-${Utilities.getUuid()}`,
      action_name: '初回提案',
      probability: 'A',
      addPrompt: '新製品のリリースについて拝見し、非常に興味を持ちました。',
      body: '貴社の新製品と弊社のサービスを組み合わせることで、大きなシナジーが生まれると考えております。'
    };

    const dummyAccount = {
      id: dummyAction.accountId,
      organization_id: `test-org-${Utilities.getUuid()}`,
      industry: 'IT・ソフトウェア',
      company_size: '51-200人'
    };

    // 2. ダミーデータからベクトル化用のテキストを生成
    const chunkText = updater._generateChunkText(dummyAction, dummyAccount);
    Logger.log(`生成されたChunk Text:\n${chunkText}`);

    // 3. テキストをベクトル化
    const embedding = updater.embeddingClient.generate(chunkText, 'RETRIEVAL_DOCUMENT');
    Logger.log(`ベクトル化成功。次元数: ${embedding.length}`);
    
    // 4. BigQueryに挿入する行データを作成
    const rowToInsert = {
      chunk_id: Utilities.getUuid(),
      organization_id: dummyAccount.organization_id,
      account_id: dummyAccount.id,
      source_document: `SalesAction:${dummyAction.ID}`,
      chunk_text: chunkText,
      embedding: embedding,
      company_size: dummyAccount.company_size,
      industry: dummyAccount.industry,
      deal_type: dummyAction.action_name
    };

    // 5. BigQueryに保存
    await updater._saveToBigQuery([rowToInsert]);

    Logger.log('✅ ダミーデータのBigQuery挿入テストが正常に完了しました。');

  } catch (e) {
    Logger.log(`❌ テスト実行中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}


// =================================================================
// KnowledgeBaseUpdater クラス
// =================================================================

class KnowledgeBaseUpdater {
  constructor(execUserEmail) {
    this.execUserEmail = execUserEmail;
    this.props = PropertiesService.getScriptProperties().getProperties();
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);
    this.embeddingClient = new EmbeddingClient('text-embedding-004'); 
  }

  /**
   * ナレッジベースを更新するメインの実行メソッド。
   */
  async updateKnowledgeBase() {
    const newSuccessActions = await this._findNewSuccessActions();

    if (!newSuccessActions || newSuccessActions.length === 0) {
      Logger.log('更新対象の成功事例はありませんでした。');
      return;
    }

    Logger.log(`${newSuccessActions.length}件の新しい成功事例を処理します。`);
    
    const rowsToInsert = [];
    const processedActionIds = [];

    for (const action of newSuccessActions) {
      try {
        const accountId = action.accountId; // AppSheetの参照列名
        if (!accountId) {
          Logger.log(`ID: ${action.ID} にはaccountIdが紐付いていないためスキップします。`);
          continue;
        }

        // 関連するAccount情報を取得
        const accountRecord = await this._findRecordById('Account', accountId);

        const chunkText = this._generateChunkText(action, accountRecord);
        if (!chunkText) {
          Logger.log(`ID: ${action.ID} はテキストが空のためスキップします。`);
          continue;
        }

        const embedding = this.embeddingClient.generate(chunkText, 'RETRIEVAL_DOCUMENT');
        
        rowsToInsert.push({
          chunk_id: Utilities.getUuid(),
          organization_id: accountRecord ? accountRecord.organization_id : null,
          account_id: accountId,
          source_document: `SalesAction:${action.ID}`,
          chunk_text: chunkText,
          embedding: embedding,
          company_size: accountRecord ? accountRecord.company_size : null,
          industry: accountRecord ? accountRecord.industry : null,
          deal_type: action.action_name
        });

        processedActionIds.push(action.ID);

      } catch (e) {
        Logger.log(`❌ アクションID ${action.ID} の処理中にエラーが発生しました: ${e.message}`);
        await this._updateVectorizedStatus([action.ID], 'Failed');
      }
    }

    if (rowsToInsert.length > 0) {
      await this._saveToBigQuery(rowsToInsert);
      await this._updateVectorizedStatus(processedActionIds, 'Completed');
    }
    
    Logger.log('ナレッジベースの更新が完了しました。');
  }

  /**
   * AppSheetからベクトル化されていない成功事例を取得します。
   */
  async _findNewSuccessActions() {
    const resultConditions = SUCCESS_RESULTS.map(r => `[result] = "${r}"`).join(', ');
    const selector = `FILTER("SalesAction", AND(OR(${resultConditions}), ISBLANK([vectorized_status])))`;
    
    Logger.log(`検索条件 (Selector): ${selector}`);

    const properties = { 
      "Selector": selector,
      "Properties": { "PageSize": BATCH_LIMIT, "Locale": "ja-JP" }
    };
    
    try {
      const results = await this.appSheetClient.findData('SalesAction', this.execUserEmail, properties);
      return (results && Array.isArray(results)) ? results : [];
    } catch (e) {
      Logger.log(`AppSheetからのデータ取得に失敗しました: ${e.message}`);
      return [];
    }
  }

  /**
   * レコードをIDで検索します。
   */
  async _findRecordById(tableName, recordId) {
    const keyColumn = (tableName === 'Account' || tableName === 'Organization') ? 'id' : 'ID';
    const selector = `FILTER("${tableName}", [${keyColumn}] = "${recordId}")`;
    const properties = { "Selector": selector };
    const result = await this.appSheetClient.findData(tableName, this.execUserEmail, properties);
    if (result && Array.isArray(result) && result.length > 0) {
      return result[0];
    }
    Logger.log(`テーブル[${tableName}]からID[${recordId}]のレコードが見つかりませんでした。応答: ${JSON.stringify(result)}`);
    return null;
  }

  /**
   * ベクトル化するための元となるテキストを生成します。
   */
  _generateChunkText(action, accountRecord) {
    let content = '';
    content += `顧客の業種: ${accountRecord ? accountRecord.industry : '不明'}\n`;
    content += `顧客の規模: ${accountRecord ? accountRecord.company_size : '不明'}\n`;
    content += `アクション名: ${action.action_name || ''}\n`;
    content += `確度: ${action.probability || ''}\n`;
    content += `商談メモ: ${action.addPrompt || ''}\n`;
    content += `提案メール本文: ${action.body || ''}`;
    return content.trim();
  }

  /**
   * 処理済みのレコードにステータスを書き込みます。
   */
  async _updateVectorizedStatus(recordIds, status) {
    if (!recordIds || recordIds.length === 0) return;
    
    const rowsToUpdate = recordIds.map(id => ({
      ID: id,
      vectorized_status: status
    }));

    try {
      await this.appSheetClient.updateRecords('SalesAction', rowsToUpdate, this.execUserEmail);
      Logger.log(`${recordIds.length}件のレコードのvectorized_statusを「${status}」に更新しました。`);
    } catch (e) {
      Logger.log(`❌ AppSheetのステータス更新に失敗しました: ${e.message}`);
    }
  }

  /**
   * データをBigQueryに保存します。
   */
  async _saveToBigQuery(rows) {
    // ★★★ 修正点: 参照するプロパティ名を変更 ★★★
    const projectId = this.props.GCP_PROJECT_ID;

    if (!projectId) {
      // ★★★ 修正点: ログメッセージも合わせる ★★★
      Logger.log('⚠️ BigQueryへの保存はスキップされました。GCP_PROJECT_IDが設定されていません。');
      return;
    }
    
    try {
      BigQuery.Tabledata.insertAll({ rows: rows.map(row => ({ json: row })) }, projectId, BIGQUERY_DATASET_ID, BIGQUERY_TABLE_ID);
      Logger.log(`${rows.length}件のデータをBigQueryに正常に挿入しました。`);
    } catch (e) {
      Logger.log(`❌ BigQueryへのデータ挿入中にエラーが発生しました: ${e.message}`);
      if (e.details && e.details.errors) {
        e.details.errors.forEach((err, index) => {
          Logger.log(`  Error[${index}]: ${err.message}`);
        });
      }
    }
  }
}
