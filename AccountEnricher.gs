/**
 * =================================================================
 * AccountEnricher (調査項目拡張版 v2)
 * =================================================================
 * Accountテーブルの新しいスキーマ定義に合わせて、AIによる企業情報の
 * 調査項目を大幅に拡張しました。
 *
 * 【v2での主な変更点】
 * - 資金調達情報、事業戦略、採用情報、技術スタックなど、アプローチの質を
 * 高めるための戦略的な調査項目を追加。
 * - AIへの指示を更新し、新しい項目を含むすべての情報を構造化データ(JSON)
 * として取得するように修正。
 * =================================================================
 */

// =================================================================
// 定数宣言
// =================================================================
const BATCH_PROCESSING_LIMIT = 10; // 一度に処理する最大件数

// =================================================================
// グローバル関数 (時間主導型トリガーで実行)
// =================================================================

/**
 * 【時間主導型トリガー用】企業情報収集バッチ処理を実行します。
 */
function runAccountEnrichmentBatch() {
  const execUserEmail = "hello@al-pa-ca.com"//PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
  
  if (!execUserEmail) {
    Logger.log('❌ エラー: スクリプトプロパティに "SERVICE_ACCOUNT" が設定されていません。バッチ処理を実行できません。');
    return;
  }

  Logger.log(`企業情報収集バッチをサービスアカウント (${execUserEmail}) で開始します。`);
  
  try {
    const enricher = new AccountEnricher(execUserEmail);
    enricher.enrichAllPendingAccounts()
      .catch(e => {
        Logger.log(`❌ バッチ処理の実行中に致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
      });
  } catch (e) {
    Logger.log(`❌ 初期化中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 【AppSheetから実行】指定された1つのアカウント情報を強制的に更新します。
 * @param {string} accountId - 更新対象のアカウントID。
 * @param {string} execUserEmail - 実行ユーザーのメールアドレス。
 */
function enrichSingleAccount(accountId, execUserEmail) {
  accountId = "9250CC98-C95A-43D9-B261-E7EFD163B5E3-b3ac984e";
  execUserEmail = "hello@al-pa-ca.com"
  if (!accountId) {
    Logger.log('❌ エラー: accountIdが指定されていません。');
    return;
  }
  if (!execUserEmail) {
    Logger.log('❌ エラー: execUserEmailが指定されていません。');
    return;
  }
  
  Logger.log(`アカウント個別更新を開始します。Account ID: ${accountId}`);

  try {
    const enricher = new AccountEnricher(execUserEmail);
    enricher.processSingleAccount(accountId)
      .catch(e => {
        Logger.log(`❌ アカウント個別更新中にエラーが発生しました (ID: ${accountId}): ${e.message}\n${e.stack}`);
      });
  } catch (e) {
    Logger.log(`❌ 初期化中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}


// =================================================================
// AccountEnricher クラス
// =================================================================

class AccountEnricher {
  constructor(execUserEmail) {
    this.execUserEmail = execUserEmail;
    this.props = PropertiesService.getScriptProperties().getProperties();
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);
    this.geminiClient = new GeminiClient('gemini-2.0-flash');
  }

  /**
   * 保留中のすべてのアカウントに対して情報収集を実行するメインメソッド。
   */
  async enrichAllPendingAccounts() {
    const pendingAccounts = await this._findPendingAccounts();

    if (!pendingAccounts || pendingAccounts.length === 0) {
      Logger.log('情報収集対象のアカウントはありませんでした。処理を終了します。');
      return;
    }

    Logger.log(`${pendingAccounts.length}件のアカウントの情報収集を開始します。`);

    for (const account of pendingAccounts) {
      await this.processSingleAccount(account.id);
    }
    Logger.log('すべてのアカウントの情報収集処理が完了しました。');
  }

  /**
   * 指定された単一のアカウント情報を収集・更新します。
   * @param {string} accountId - 更新対象のアカウントID。
   */
  async processSingleAccount(accountId) {
    try {
      const account = await this._findRecordById('Account', accountId);
      if (!account) {
        throw new Error(`ID: ${accountId} のアカウントが見つかりませんでした。`);
      }

      const companyName = account.name;
      if (!companyName) {
        Logger.log(`ID: ${account.id} には会社名(name)がないためスキップします。`);
        await this._updateAccountStatus(account.id, 'Skipped');
        return;
      }
      
      Logger.log(`処理中: Account ID [${account.id}], 会社名 [${companyName}]`);

      // 既存のメソッドを再利用して情報収集
      const enrichedData = await this._enrichAccountData(companyName, account.website_url);
      
      if (enrichedData) {
        enrichedData.enrichment_status = 'Completed';
        await this._updateAccountInAppSheet(account.id, enrichedData);
        Logger.log(`-> 成功: Account ID [${account.id}] の情報を更新しました。`);
      } else {
        await this._updateAccountStatus(account.id, 'Failed');
        Logger.log(`-> 失敗: Account ID [${account.id}] の情報収集に失敗しました。ステータスを'Failed'に更新します。`);
      }
    } catch (error) {
      Logger.log(`-> エラー: Account ID [${accountId}] の個別更新処理中に予期せぬエラーが発生しました: ${error.stack}`);
      // 失敗した場合でもステータス更新を試みる
      await this._updateAccountStatus(accountId, 'Failed').catch(e => Logger.log(`ステータス更新にも失敗しました: ${e.message}`));
    }
  }


  /**
   * AppSheetから情報収集が保留中（Pending）のアカウントを取得します。
   */
  async _findPendingAccounts() {
    const selector = `FILTER("Account", [enrichment_status] = "Pending")`;
    const properties = { 
      "Selector": selector,
      "Properties": { "PageSize": BATCH_PROCESSING_LIMIT, "Locale": "ja-JP" }
    };
    
    try {
      const results = await this.appSheetClient.findData('Account', this.execUserEmail, properties);
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
    const keyColumn = 'id'; // Accountテーブルのキーは 'id'
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
   * AI（Gemini）を用いて企業の詳細情報を収集します。
   */
  async _enrichAccountData(companyName, websiteUrl) {
    // ★★★ 修正点: company_sizeの指示をText型を許容する形に戻しました ★★★
    const prompt = `
      あなたはプロの企業調査アナリストです。
      以下の企業について、公開情報から徹底的に調査し、指定されたJSON形式で回答してください。

      # 調査対象企業
      - 会社名: ${companyName}
      - ウェブサイト: ${websiteUrl || '不明'}

      # 収集項目
      ## 基本情報
      - company_description: 事業内容の包括的な説明
      - corporate_number: 法人番号
      - main_service: 主要な製品やサービスの概要
      - industry: 業種
      - company_size: 企業規模・従業員数
      - target_audience: ターゲット顧客
      - linkedin_url: LinkedInの企業ページURL

      ## 戦略・財務情報
      - funding_ir_info: 直近の資金調達の状況や、投資家向け(IR)情報の要約
      - business_strategy: 中期経営計画や今後の事業戦略の要約

      ## 人材・技術情報
      - hiring_info: 現在の採用情報、特に強化している職種の要約
      - tech_stack: Webサイトや採用情報から推測される利用技術 (例: AWS, Salesforce, React)

      ## マーケティング・広報情報
      - customer_case_studies: 顧客向けの導入事例や成功事例の要約
      - event_info: 直近のセミナー登壇やイベント出展情報の要約
      - last_signal_summary: 上記以外の最新ニュースやプレスリリース

      # 出力形式 (JSON)
      見つからない情報は "不明" または null としてください。
      {
        "company_description": "...",
        "corporate_number": "...",
        "main_service": "...",
        "industry": "...",
        "company_size": "...",
        "target_audience": "...",
        "linkedin_url": "...",
        "funding_ir_info": "...",
        "business_strategy": "...",
        "hiring_info": "...",
        "tech_stack": "...",
        "customer_case_studies": "...",
        "event_info": "...",
        "last_signal_summary": "..."
      }
    `;

    try {
      this.geminiClient.setPromptText(prompt);
      const response = await this.geminiClient.generateCandidates();
      const responseText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      
      const jsonMatch = responseText.match(/{[\s\S]*}/);
      if (!jsonMatch) {
        throw new Error("AIの応答から有効なJSONを抽出できませんでした。");
      }
      return JSON.parse(jsonMatch[0]);

    } catch (error) {
      Logger.log(`Geminiでの情報収集またはJSONパース中にエラーが発生しました (会社名: ${companyName}): ${error.stack}`);
      return null;
    }
  }

  /**
   * 収集したデータでAppSheetのレコードを更新します。
   */
  async _updateAccountInAppSheet(accountId, data) {
    const rowToUpdate = {
      id: accountId, // Accountテーブルのキーは 'id'
      ...data
    };
    
    await this.appSheetClient.updateRecords('Account', [rowToUpdate], this.execUserEmail);
  }
  
  /**
   * アカウントのステータスのみを更新します（主にエラー発生時に使用）。
   */
  async _updateAccountStatus(accountId, status) {
    try {
      await this._updateAccountInAppSheet(accountId, { enrichment_status: status });
    } catch (error) {
      Logger.log(`Account ID [${accountId}] のステータス更新中にエラーが発生しました: ${error.stack}`);
    }
  }
}
