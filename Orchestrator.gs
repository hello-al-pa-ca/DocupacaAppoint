/**
 * =================================================================
 * Orchestrator (v1.0)
 * =================================================================
 * 複数のモジュールを連携させ、特定のビジネスフローを実行するための
 * オーケストレーション・スクリプトです。
 *
 * 【機能】
 * - ワンクリック・ファーストアプローチ提案機能
 * =================================================================
 */

/**
 * 【AppSheetから呼び出す関数】
 * 指定された名刺情報(BusinessCard)に基づき、企業情報をリッチ化し、
 * 初回アプローチ用のメール文面をAIに生成させる一連のプロセスを実行します。
 *
 * @param {string} contactId - アプローチ対象となるBusinessCardテーブルのレコードID。
 * @param {string} execUserEmail - 実行ユーザーのメールアドレス。
 */
async function generateFirstApproachProposal(contactId, execUserEmail) {
  // 1. 引数チェック
  if (!contactId || !execUserEmail) {
    Logger.log('❌ [ERROR] 引数が不足しています。contactId, execUserEmailは必須です。');
    return;
  }
  Logger.log(`[START] ワンクリック・ファーストアプローチ提案を開始します。(Contact ID: ${contactId})`);

  try {
    // 2. 関連情報の取得
    const props = PropertiesService.getScriptProperties().getProperties();
    const appSheetClient = new AppSheetClient(props.APPSHEET_APP_ID, props.APPSHEET_API_KEY);

    // 名刺(Contact)情報を取得
    const contactRecord = await findRecordById_(appSheetClient, 'BusinessCard', contactId, execUserEmail);
    if (!contactRecord) throw new Error(`ID [${contactId}] の名刺レコードが見つかりません。`);
    
    const accountId = contactRecord.account_id;
    if (!accountId) throw new Error(`名刺 [${contactId}] に企業(Account)が紐付いていません。`);
    
    // 企業(Account)情報を取得
    let accountRecord = await findRecordById_(appSheetClient, 'Account', accountId, execUserEmail);
    if (!accountRecord) throw new Error(`ID [${accountId}] の企業レコードが見つかりません。`);

    Logger.log(`[INFO] 対象企業: ${accountRecord.name} (ID: ${accountId})`);

    // 3. 企業情報の拡充 (Enrichment)
    if (accountRecord.enrichment_status !== 'Completed') {
      Logger.log(`[INFO] 企業情報が未拡充のため、Enrichmentプロセスを開始します...`);
      const enricher = new AccountEnricher(execUserEmail); // AccountEnricherは内部で専用AppIDを参照
      await enricher.processSingleAccount(accountId);
      Logger.log(`[SUCCESS] Enrichmentプロセスが完了しました。`);
      
      // 更新された企業情報を再取得
      accountRecord = await findRecordById_(appSheetClient, 'Account', accountId, execUserEmail);
      if (!accountRecord) throw new Error(`Enrichment後、ID [${accountId}] の企業レコードの再取得に失敗しました。`);
    } else {
      Logger.log(`[INFO] 企業情報は既に拡充済みです。`);
    }

    // 4. SalesActionレコードを先行して作成
    Logger.log(`[INFO] AI提案を格納するためのSalesActionレコードを作成します...`);
    const actionPayload = {
      accountId: accountId,
      business_card_id: contactId,
      action_name: '初回アプローチ', //マスター定義を想定
      contact_method: 'メール', //マスター定義を想定
      execute_ai_status: 'AI提案生成中...'
    };
    const addActionResponse = await appSheetClient.addRecords('SalesAction', [actionPayload], execUserEmail);
    if (!addActionResponse.Rows || addActionResponse.Rows.length === 0 || !addActionResponse.Rows[0].ID) {
        throw new Error("SalesActionレコードの作成に失敗したか、応答からIDを取得できませんでした。");
    }
    const salesActionId = addActionResponse.Rows[0].ID;
    Logger.log(`[SUCCESS] SalesActionレコードを作成しました。(ID: ${salesActionId})`);

    // 5. AISalesActionを実行して、メール文面を生成
    Logger.log(`[INFO] SalesCopilotを起動し、メール文面の生成を指示します...`);
    const copilot = new SalesCopilot(execUserEmail);

    // executeAISalesActionに必要な引数を準備
    const params = {
      recordId: salesActionId,
      organizationId: accountRecord.organization_id,
      accountId: accountId,
      AIRoleName: 'トップセールス', // デフォルト値 or マスターから取得
      actionName: '初回アプローチ',
      contactMethod: 'メール',
      mainPrompt: null, // 初回はマスターのプロンプトを使う想定
      addPrompt: '名刺交換のお礼と、貴社の事業について拝見した内容を踏まえたご提案です。',
      companyName: accountRecord.name,
      companyAddress: accountRecord.address,
      customerContactName: contactRecord.name,
      ourContactName: execUserEmail.split('@')[0], // 仮。後でApplicationUserテーブルから取得するよう改修を推奨
      probability: 'C', // 初期確度
      eventName: contactRecord.event_name,
      referenceUrls: '',
      execUserEmail: execUserEmail
    };

    // 非同期で実行（AppSheet側は待たない）
    copilot.executeAISalesAction(...Object.values(params)).catch(e => {
        Logger.log(`❌ AI提案生成の非同期実行中にエラー: ${e.message}\n${e.stack}`);
        // エラー発生時にステータスを更新
        appSheetClient.updateRecords('SalesAction', [{ ID: salesActionId, "execute_ai_status": "エラー" }], execUserEmail);
    });
    
    Logger.log(`[END] ワンクリック・ファーストアプローチ提案の処理をバックグラウンドで開始しました。`);

  } catch (error) {
    Logger.log(`❌ [FATAL] 致命的なエラーが発生しました: ${error.message}\n${error.stack}`);
    // ここでAppSheetにエラーを返す処理を追加することも可能
  }
}


/**
 * 汎用のレコードID検索ヘルパー関数
 * @private
 */
async function findRecordById_(appSheetClient, tableName, recordId, execUserEmail) {
    const keyColumn = (tableName === 'Account' || tableName === 'Organization' || tableName === 'BusinessCard') ? 'id' : 'ID';
    const selector = `FILTER("${tableName}", [${keyColumn}] = "${recordId}")`;
    const result = await appSheetClient.findData(tableName, execUserEmail, { "Selector": selector });
    if (result && Array.isArray(result) && result.length > 0) {
      return result[0];
    }
    return null;
}
