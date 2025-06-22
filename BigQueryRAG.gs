/**
 * =================================================================
 * 成功事例ナレッジベース自動更新スクリプト (RAG - Part 1) v2
 * =================================================================
 * このスクリプトは、1日1回、夜間などに定時実行トリガーで呼び出されることを想定しています。
 * 過去の成功したSalesActionを抽出し、AIによる要約とベクトル化を行い、
 * BigQuery上のナレッジベースに蓄積します。
 *
 * 【v2での主な変更点】
 * - ユーザーから提供されたBigQueryのスキーマに合わせて、保存するデータの構造を変更しました。
 * - SalesActionに加え、関連するBusinessCard（顧客情報）も参照し、
 * `company_size`や`industry`といった属性情報もナレッジに含めるようにしました。
 * =================================================================
 */

// =================================================================
// グローバル設定
// =================================================================
const GcpProjectId = PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');
const BqDatasetId = 'rag_knowledge_base'; // BigQueryのデータセットID
const BqTableId = 'knowledge_base'; // ナレッジベースのテーブルID

/**
 * 【トリガー実行用】ナレッジベースの更新を日次で実行します。
 * この関数を毎日深夜などに実行するよう、Apps Scriptのトリガーを設定してください。
 */
function runDailyKnowledgeBaseUpdate() {
  try {
    // 実行ユーザーはスクリプトの所有者や固定の管理者を指定するのが一般的です
    const execUserEmail = 'admin@your-company.com';
    Logger.log('ナレッジベースの更新処理を開始します...');
    
    const manager = new KnowledgeBaseManager(execUserEmail);
    manager.updateKnowledgeBase();
    
    Logger.log('ナレッジベースの更新処理が正常に完了しました。');
  } catch (e) {
    Logger.log(`ナレッジベースの更新中に致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * @class KnowledgeBaseManager
 * @description 成功事例の抽出、要約、ベクトル化、DB保存を管理するクラス。
 */
class KnowledgeBaseManager {
  /**
   * @param {string} execUserEmail 実行ユーザーのメールアドレス
   */
  constructor(execUserEmail) {
    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);
    this.embeddingClient = new EmbeddingClient('text-embedding-004'); // ベクトル化に使用するモデル
    this.customerDataCache = {}; // 顧客データのキャッシュ
  }

  /**
   * ナレッジベースの更新プロセス全体を実行します。
   */
  updateKnowledgeBase() {
    // 1. AppSheetから成功事例を抽出
    const successfulActions = this._extractSuccessfulActions();
    if (!successfulActions || successfulActions.length === 0) {
      Logger.log("更新対象の新しい成功事例はありませんでした。");
      return;
    }
    Logger.log(`${successfulActions.length}件の成功事例を抽出しました。`);

    // 2. 各事例を要約し、ベクトル化
    const knowledgeEntries = successfulActions.map(action => {
      const customerInfo = this._getCustomerInfo(action.business_card_id);
      if (!customerInfo) {
        Logger.log(`アクション[${action.id}]の顧客情報が見つからないためスキップします。`);
        return null;
      }
      
      const summary = this._summarizeAction(action, customerInfo);
      if (!summary) return null;

      const embedding = this.embeddingClient.generate(summary, 'RETRIEVAL_DOCUMENT');
      
      return {
        chunk_id: action.id, 
        organization_id: customerInfo.organization_id || '', // 組織ID
        account_id: action.business_card_id,
        source_document: action.id,
        chunk_text: summary,
        embedding: embedding,
        company_size: customerInfo.company_size || '',
        industry: customerInfo.industry || '',
        deal_type: action.action_name || ''
      };
    }).filter(entry => entry); // nullになったものを除去

    // 3. BigQueryに保存
    if (knowledgeEntries.length > 0) {
      this._upsertToBigQuery(knowledgeEntries);
    }
  }

  /**
   * AppSheetから成功したSalesActionレコードを抽出します。
   * @private
   * @returns {Object[]} 成功したアクションのレコード配列
   */
  _extractSuccessfulActions() {
    // 「受注」または「アポイント取得」したアクションを成功と定義
    const successConditions = ['受注', 'アポイント取得'];
    // 過去7日間の成功事例のみを対象にする例（期間は調整可能）
    const selector = `FILTER("SalesAction", AND(IN([result], ${JSON.stringify(successConditions)}), ([実施日時] > (NOW() - "168:00:00"))))`;

    const results = this.appSheetClient.findData('SalesAction', this.execUserEmail, { "Selector": selector });
    return results || [];
  }
  
  /**
   * AppSheetから顧客情報を取得します（キャッシュ付き）
   * @private
   * @param {string} customerId - 顧客ID
   * @returns {Object|null} - 顧客情報のレコード
   */
  _getCustomerInfo(customerId) {
    if (this.customerDataCache[customerId]) {
        return this.customerDataCache[customerId];
    }
    const selector = `FILTER("BusinessCard", [id] = "${customerId}")`;
    const result = this.appSheetClient.findData('BusinessCard', this.execUserEmail, { "Selector": selector });
    if (result && result.length > 0) {
        this.customerDataCache[customerId] = result[0];
        return result[0];
    }
    return null;
  }

  /**
   * 個々のアクションの内容をAIに要約させます。
   * @private
   * @param {Object} action - 要約対象のSalesActionレコード
   * @param {Object} customerInfo - 顧客情報のレコード
   * @returns {string|null} - AIが生成した要約テキスト
   */
  _summarizeAction(action, customerInfo) {
    try {
      const context = `
        - 顧客名: ${customerInfo.companyName || '不明'}
        - 業種: ${customerInfo.industry || '不明'}
        - 企業規模: ${customerInfo.company_size || '不明'}
        - 実施したアクション: ${action.action_name} (${action.contactMethod})
        - 提案内容の要点: ${action.addPrompt || '記載なし'}
        - 最終的な結果: ${action.result}
      `;
      const prompt = `以下の営業活動の記録について、「どのような特徴の顧客に、どんな提案をして、どう成功したのか」が分かるように、150文字程度で簡潔な要約を作成してください。\n\n${context}`;
      
      const summarizerClient = new GeminiClient('gemini-1.5-flash-latest');
      summarizerClient.setPromptText(prompt);
      const response = summarizerClient.generateCandidates();
      return (response.candidates[0].content.parts || []).map(p => p.text).join('');
    } catch (e) {
      Logger.log(`アクション[${action.id}]の要約中にエラー: ${e.message}`);
      return null;
    }
  }

  /**
   * 処理したデータをBigQueryに挿入します（Upsert相当の処理）。
   * @private
   * @param {Object[]} entries - BigQueryに挿入するエントリの配列
   */
  _upsertToBigQuery(entries) {
    const chunkIds = entries.map(entry => entry.chunk_id);
    
    // 1. まず、今回処理するIDの古いデータを削除（再処理の場合に対応）
    const deleteSql = `DELETE FROM \`${GcpProjectId}.${BqDatasetId}.${BqTableId}\` WHERE chunk_id IN UNNEST(@chunk_ids)`;
    const deleteRequest = {
      query: deleteSql,
      useLegacySql: false,
      queryParameters: [{
        name: 'chunk_ids',
        parameterType: { type: 'ARRAY', arrayType: { type: 'STRING' } },
        parameterValue: { arrayValues: chunkIds.map(id => ({ value: id })) }
      }]
    };
    BigQuery.Jobs.query(deleteRequest, GcpProjectId);
    Logger.log(`${chunkIds.length}件の既存ナレッジデータを削除しました（重複防止）。`);

    // 2. 新しいデータを挿入
    const rows = entries.map(entry => ({
      json: entry
    }));
    
    const insertRequest = { rows: rows };
    const response = BigQuery.Tabledata.insertAll(insertRequest, GcpProjectId, BqDatasetId, BqTableId);
    
    if (response.insertErrors) {
      Logger.log(`BigQueryへのデータ挿入中にエラーが発生しました: ${JSON.stringify(response.insertErrors)}`);
    } else {
      Logger.log(`${rows.length}件の新しいナレッジをBigQueryに正常に保存しました。`);
    }
  }
}
