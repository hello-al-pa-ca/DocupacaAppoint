/**
 * =================================================================
 * 成功事例ナレッジベース自動更新スクリプト (RAG - Part 1) v9
 * =================================================================
 * このスクリプトは、1日1回、夜間などに定時実行トリガーで呼び出されることを想定しています。
 * 過去の成功したSalesActionを抽出し、AIによる要約とベクトル化を行い、
 * BigQuery上のナレッジベースに蓄積します。
 *
 * 【v9での主な変更点】
 * - 企業規模や業種などの顧客情報を、BusinessCardテーブルではなく、
 * その親であるAccountテーブルから取得するようにロジックを修正しました。
 * これにより、より正確で一元化された情報がナレッジに反映されます。
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
  // async/awaitをトップレベルで安全に扱うため、無名関数でラップして実行します。
  (async () => {
    try {
      const execUserEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
      if (!execUserEmail) {
        throw new Error("スクリプトプロパティ 'SERVICE_ACCOUNT' が設定されていません。");
      }
      
      Logger.log(`ナレッジベースの更新処理を開始します... (実行アカウント: ${execUserEmail})`);
      
      const manager = new KnowledgeBaseManager(execUserEmail);
      await manager.updateKnowledgeBase();
      
      Logger.log('ナレッジベースの更新処理が正常に完了しました。');
    } catch (e) {
      Logger.log(`ナレッジベースの更新中に致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
    }
  })().catch(e => Logger.log(`❌ 非同期処理の実行中にエラー: ${e.message}`));
}

/**
 * 【開発・テスト用】KnowledgeBaseManagerの処理を直接実行します。
 */
function test_KnowledgeBaseManager() {
    (async () => {
    try {
      const execUserEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
      if (!execUserEmail) {
        throw new Error("スクリプトプロパティ 'SERVICE_ACCOUNT' が設定されていません。");
      }
      
      Logger.log(`テスト実行を開始します... (実行アカウント: ${execUserEmail})`);
      
      const manager = new KnowledgeBaseManager(execUserEmail);
      await manager.updateKnowledgeBase();
      
      Logger.log('テスト実行が完了しました。');
    } catch (e) {
      Logger.log(`テスト実行中にエラーが発生しました: ${e.message}\n${e.stack}`);
    }
  })().catch(e => Logger.log(`❌ 非同期処理の実行中にエラー: ${e.message}`));
}

/**
 * 【開発・テスト用】動作確認のためのダミーデータをAppSheetに登録します。
 */
function addDummyDataForTesting() {
  (async () => {
    try {
      const execUserEmail = Session.getActiveUser().getEmail() || PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
      Logger.log(`ダミーデータの登録を開始します... (実行アカウント: ${execUserEmail})`);

      const client = new AppSheetClient(
        PropertiesService.getScriptProperties().getProperty('APPSHEET_APP_ID'),
        PropertiesService.getScriptProperties().getProperty('APPSHEET_API_KEY')
      );
      
      // 1. ダミーのAccountデータを作成
      const newAccountId = Utilities.getUuid();
      const dummyAccount = {
        id: newAccountId,
        name: "株式会社ダミーインダストリー",
        industry: "製造業",
        company_size: "501-1000名",
        domain: "dummy-industry.example.com"
      };
      await client.addRecords("Account", [dummyAccount], execUserEmail);
      Logger.log(`ダミーのAccountを作成しました: ${dummyAccount.name} (ID: ${newAccountId})`);

      // 2. ダミーの顧客データを作成
      const newBusinessCardId = Utilities.getUuid();
      const dummyCustomer = {
        id: newBusinessCardId,
        account_id: newAccountId, // 作成したAccountに紐付ける
        companyName: dummyAccount.name,
        name: "ダミー 次郎",
        department: "生産管理部",
        position: "課長",
        email: `jiro.dummy@${dummyAccount.domain}`,
        enabled_contact: true
      };
      await client.addRecords("BusinessCard", [dummyCustomer], execUserEmail);
      Logger.log(`ダミーの顧客を作成しました: ${dummyCustomer.name} (ID: ${newBusinessCardId})`);
      
      // 3. 上記顧客に対する「成功事例」となるアクションを作成
      const dummySalesAction = {
        id: Utilities.getUuid(),
        business_card_id: newBusinessCardId,
        progress: "提案済/検討中",
        action_name: "提案後フォローアップ",
        contactMethod: "メール",
        result: "受注",
        excuted_dt: new Date().toISOString(),
        addPrompt: "最終提案後、AIが生成したフォローアップメールを送信したところ、翌日に受注の連絡があった。特に、顧客の過去の発言を踏まえたパーソナライズが効果的だった模様。"
      };
      await client.addRecords("SalesAction", [dummySalesAction], execUserEmail);
      Logger.log(`ダミーの成功事例を作成しました: ${dummySalesAction.action_name} -> ${dummySalesAction.result}`);
      
      Logger.log("ダミーデータの登録が完了しました。");

    } catch (e) {
      Logger.log(`ダミーデータの登録中にエラーが発生しました: ${e.message}\n${e.stack}`);
    }
  })().catch(e => Logger.log(`❌ 非同期処理の実行中にエラー: ${e.message}`));
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
    this.accountDataCache = {}; // Accountデータのキャッシュ
  }

  /**
   * ナレッジベースの更新プロセス全体を実行します。
   */
  async updateKnowledgeBase() {
    // 1. AppSheetから成功事例を抽出
    const successfulActions = await this._extractSuccessfulActions();
    if (!successfulActions || successfulActions.length === 0) {
      Logger.log("更新対象の新しい成功事例はありませんでした。");
      return;
    }
    Logger.log(`${successfulActions.length}件の成功事例を抽出しました。`);

    // 2. 各事例を要約し、ベクトル化
    const knowledgeEntries = await Promise.all(successfulActions.map(async (action) => {
      const businessCard = await this._getBusinessCardInfo(action.business_card_id);
      if (!businessCard || !businessCard.account_id) {
        Logger.log(`アクション[${action.id}]の顧客情報またはアカウント情報が見つからないためスキップします。`);
        return null;
      }
      
      let accountInfo = await this._getAccountInfo(businessCard.account_id);
      if (!accountInfo) {
        Logger.log(`アカウント[${businessCard.account_id}]の情報が見つからないためスキップします。`);
        return null;
      }
      
      // 業種や企業規模が不明な場合に情報を補完
      if (accountInfo.name && (!accountInfo.industry || !accountInfo.company_size)) {
        const enrichedInfo = this._enrichCustomerInfo(accountInfo.name);
        accountInfo.industry = accountInfo.industry || enrichedInfo.industry;
        accountInfo.company_size = accountInfo.company_size || enrichedInfo.company_size;
      }
      
      const summary = this._summarizeAction(action, businessCard, accountInfo);
      if (!summary) return null;

      const embedding = this.embeddingClient.generate(summary, 'RETRIEVAL_DOCUMENT');
      
      return {
        chunk_id: action.id, 
        organization_id: accountInfo.organization_id || '', 
        account_id: businessCard.account_id,
        source_document: action.id,
        chunk_text: summary,
        embedding: embedding,
        company_size: accountInfo.company_size || '',
        industry: accountInfo.industry || '',
        deal_type: action.action_name || ''
      };
    }));
    
    const validEntries = knowledgeEntries.filter(entry => entry); // nullになったものを除去

    // 3. BigQueryに保存
    if (validEntries.length > 0) {
      this._upsertToBigQuery(validEntries);
    }
  }

  /**
   * AppSheetから成功したSalesActionレコードを抽出します。
   * @private
   * @returns {Promise<Object[]>} 成功したアクションのレコード配列
   */
  async _extractSuccessfulActions() {
    // 「受注」または「アポイント取得」したアクションを成功と定義
    const successConditions = ['受注', 'アポイント取得'];
    // 過去7日間の成功事例のみを対象にする例（期間は調整可能）
    const selector = `FILTER("SalesAction", AND(IN([result], ${JSON.stringify(successConditions)}), ([excuted_dt] > (NOW() - "168:00:00"))))`;

    const results = await this.appSheetClient.findData('SalesAction', this.execUserEmail, { "Selector": selector });
    return results || [];
  }
  
  /**
   * AppSheetから名刺（個人）情報を取得します。
   * @private
   * @param {string} businessCardId - 名刺ID
   * @returns {Promise<Object|null>} - 名刺情報のレコード
   */
  async _getBusinessCardInfo(businessCardId) {
    if (!businessCardId) return null;
    const selector = `FILTER("BusinessCard", [id] = "${businessCardId}")`;
    const result = await this.appSheetClient.findData('BusinessCard', this.execUserEmail, { "Selector": selector });
    return (result && result.length > 0) ? result[0] : null;
  }

  /**
   * AppSheetからアカウント（企業）情報を取得します（キャッシュ付き）
   * @private
   * @param {string} accountId - アカウントID
   * @returns {Promise<Object|null>} - アカウント情報のレコード
   */
  async _getAccountInfo(accountId) {
    if (!accountId) return null;
    if (this.accountDataCache[accountId]) {
        return this.accountDataCache[accountId];
    }
    const selector = `FILTER("Account", [id] = "${accountId}")`;
    const result = await this.appSheetClient.findData('Account', this.execUserEmail, { "Selector": selector });
    if (result && result.length > 0) {
        this.accountDataCache[accountId] = result[0];
        return result[0];
    }
    return null;
  }

  /**
   * Google検索を使い、企業の業種と規模情報を補完します。
   * @private
   * @param {string} companyName - 調査対象の企業名
   * @returns {{industry: string, company_size: string}} - 調査結果
   */
  _enrichCustomerInfo(companyName) {
    if (!companyName) return { industry: '', company_size: '' };
    
    try {
      Logger.log(`企業情報の補完を開始: ${companyName}`);
      const prompt = `日本の企業「${companyName}」について、公開情報から「業種」と「従業員数に基づいた企業規模」を調べてください。企業規模は以下の選択肢から最も適切なものを選んでください: [1-10名, 11-50名, 51-100名, 101-500名, 501-1000名, 1001名以上]。以下のJSON形式だけで回答してください:\n{"industry": "（業種）", "company_size": "（企業規模の選択肢）"}`;
      
      const client = new GeminiClient('gemini-1.5-flash-latest');
      client.enableGoogleSearchTool();
      client.setPromptText(prompt);
      
      const response = client.generateCandidates();
      const textResponse = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      const cleanedJsonString = textResponse.replace(/^```json\s*|```\s*$/g, '').trim();
      const enrichedData = JSON.parse(cleanedJsonString);

      Logger.log(`補完結果: ${JSON.stringify(enrichedData)}`);
      return {
        industry: enrichedData.industry || '',
        company_size: enrichedData.company_size || ''
      };

    } catch (e) {
      Logger.log(`企業情報の補完中にエラーが発生しました (${companyName}): ${e.message}`);
      return { industry: '', company_size: '' };
    }
  }

  /**
   * 個々のアクションの内容をAIに要約させます。
   * @private
   * @param {Object} action - 要約対象のSalesActionレコード
   * @param {Object} businessCard - 名刺情報のレコード
   * @param {Object} accountInfo - アカウント情報のレコード
   * @returns {string|null} - AIが生成した要約テキスト
   */
  _summarizeAction(action, businessCard, accountInfo) {
    try {
      const context = `
        - 顧客名: ${accountInfo.name || '不明'} (${businessCard.name || '担当者不明'})
        - 業種: ${accountInfo.industry || '不明'}
        - 企業規模: ${accountInfo.company_size || '不明'}
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
