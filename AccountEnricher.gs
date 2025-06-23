/**
 * @fileoverview AccountEnrichmentBatch.gs
 * AppSheet上のアカウントテーブルから情報が未収集のレコードを取得し、
 * AIとGoogle検索を用いて企業情報を自動収集・更新するバッチ処理スクリプト。
 * 30分ごとの時間主導型トリガーでの実行を想定しています。
 * 依存するApiClientクラスは、外部ライブラリとして定義されていることを前提とします。
 */

// =================================================================
// グローバル設定
// =================================================================

/**
 * @const {number} BATCH_PROCESSING_LIMIT
 * 一度のバッチ処理で処理するアカウントの最大件数。
 * GASの実行時間制限（6分）を考慮して調整してください。
 */
const BATCH_PROCESSING_LIMIT = 10;


// =================================================================
// トリガー関数 (GASから直接呼び出される)
// =================================================================

/**
 * 企業情報収集バッチ処理を実行するトリガー関数。
 * Google Apps Scriptの時間主導型トリガーにこの関数を設定します。
 */
function runAccountEnrichmentBatch() {
  try {
    const userEmail = 'hello@al-pa-ca.com';
    
    Logger.log(`バッチ処理を開始します。実行者: ${userEmail}`);
    
    // AccountEnricherクラスのインスタンスを作成し、処理を実行
    const enricher = new AccountEnricher(userEmail);
    
    // 非同期処理を実行し、完了を待つ
    enricher.enrichAllPendingAccounts().catch(e => {
        Logger.log(`バッチ処理の実行中に致命的なエラーが発生しました: ${e.stack || e}`);
    });
    
  } catch (error) {
    Logger.log(`トリガー関数の実行中にエラーが発生しました: ${error.stack || error}`);
  }
}


// =================================================================
// AccountEnricher クラス
// =================================================================

/**
 * @class AccountEnricher
 * @description 企業情報を自動収集し、AppSheetのAccountテーブルを更新する責務を持つクラス
 */
class AccountEnricher {
  /**
   * @constructor
   * @param {string} execUserEmail - スクリプトを実行するユーザーのメールアドレス
   */
  constructor(execUserEmail) {
    /** @private */
    this.execUserEmail = execUserEmail;
    
    // --- AppSheet APIクライアントの初期化 ---
    const appId = PropertiesService.getScriptProperties().getProperty('APPSHEET_APP_ID');
    const apiKey = PropertiesService.getScriptProperties().getProperty('APPSHEET_API_KEY');
    if (!appId || !apiKey) {
      throw new Error('スクリプトプロパティに APPSHEET_APP_ID と APPSHEET_API_KEY を設定してください。');
    }
    /** @private */
    this.appSheetClient = new AppSheetClient(appId, apiKey); 
    
    // --- Gemini APIクライアントの初期化 ---
    const geminiModel = 'gemini-1.5-flash-latest';
    /** @private */
    this.geminiClient = new GeminiClient(geminiModel);
  }

  /**
   * 保留中のすべてのアカウントに対して情報収集を実行するメインメソッド。
   */
  async enrichAllPendingAccounts() {
    const pendingAccounts = await this._findPendingAccounts();

    if (pendingAccounts.length === 0) {
      Logger.log('情報収集対象のアカウントはありませんでした。処理を終了します。');
      return;
    }

    Logger.log(`${pendingAccounts.length}件のアカウントの情報収集を開始します。`);

    // ★★★ デバッグのため、取得したレコードのうち、enrichment_status が "Pending" のものだけを処理対象とします ★★★
    const accountsToProcess = pendingAccounts.filter(account => account.enrichment_status === 'Pending');

    if (accountsToProcess.length === 0) {
      Logger.log('レコードは取得できましたが、処理対象（Pendingステータス）のアカウントはありませんでした。');
      Logger.log('取得した全レコードのステータスを確認してください。');
      return;
    }

    Logger.log(`取得した ${pendingAccounts.length} 件のうち、${accountsToProcess.length} 件が処理対象です。`);

    for (const account of accountsToProcess) {
      try {
        const companyName = account.会社名 || account.Name;
        Logger.log(`処理中: Account ID [${account.id}], 会社名 [${companyName}]`);

        const enrichedData = await this._enrichAccountData(account);
        
        if (enrichedData) {
          enrichedData.enrichment_status = 'Completed';
          await this._updateAccountInAppSheet(account.id, enrichedData);
          Logger.log(`-> 成功: Account ID [${account.id}] の情報を更新しました。`);
        } else {
          await this._updateAccountStatus(account.id, 'Failed');
          Logger.log(`-> 失敗: Account ID [${account.id}] の情報収集に失敗しました。ステータスを'Failed'に更新します。`);
        }
      } catch (error) {
        Logger.log(`-> エラー: Account ID [${account.id}] の処理中に予期せぬエラーが発生しました: ${error.stack}`);
        await this._updateAccountStatus(account.id, 'Failed');
      }
    }
    Logger.log('すべてのアカウントの情報収集処理が完了しました。');
  }

  /**
   * AppSheetからアカウントを取得します。
   */
  async _findPendingAccounts() {
    try {
      // ★★★ 修正点: 問題切り分けのため、一度すべてのレコードを取得します ★★★
      // "Selector" を指定しないことで、全件取得を試みます。
      const findProperties = { 
        "PageSize": BATCH_PROCESSING_LIMIT,
        "Locale": "ja-JP"
      };
      
      Logger.log(`[デバッグ] findDataをSelectorなしで呼び出します。渡すプロパティ: ${JSON.stringify(findProperties)}`);
      const response = await this.appSheetClient.findData('Account', this.execUserEmail, findProperties);
      
      if (!response) {
          Logger.log(`処理対象のアカウントが見つかりませんでした（AppSheetからの応答が空です）。`);
          return [];
      }
      if (typeof response === 'string') {
          if (response.trim() === '' || response.includes("Success (No Content)")) {
              Logger.log(`処理対象のアカウントが見つかりませんでした（AppSheetからの応答が空またはNo Contentです）。`);
              return [];
          }
          try {
              const parsedResponse = JSON.parse(response);
              if (Array.isArray(parsedResponse)) {
                  Logger.log(`AppSheetから ${parsedResponse.length} 件のアカウントを取得しました。`);
                  return parsedResponse;
              } else {
                  Logger.log(`AppSheetからの応答をパースしましたが、配列ではありませんでした。応答: ${response}`);
                  return [];
              }
          } catch (e) {
              Logger.log(`AppSheetからの応答のJSONパースに失敗しました。エラー: ${e.message}, 応答: ${response}`);
              return [];
          }
      }
      if (Array.isArray(response)) {
        Logger.log(`AppSheetから ${response.length} 件のアカウントを取得しました。`);
        return response;
      }

      Logger.log(`処理対象のアカウントが見つかりませんでした（AppSheetからの応答が予期せぬ形式です）。応答: ${JSON.stringify(response)}`);
      return [];
    } catch (error) {
      Logger.log(`_findPendingAccountsの実行中にエラーが発生しました: ${error.stack}`);
      return [];
    }
  }

  /**
   * AI（Gemini）を用いて企業の詳細情報を収集します。
   */
  async _enrichAccountData(account) {
    const companyName = account.会社名 || account.Name;
    const domain = account.ドメイン || account.Domain;

    if (!companyName) {
        Logger.log(`Account ID [${account.id}] に会社名が存在しないため、情報収集をスキップします。`);
        return null;
    }
    
    const prompt = `あなたはプロの企業調査アナリストです。以下の企業について、公開情報から調査し、指定されたJSON形式で回答してください。\n# 調査対象企業\n- 会社名: ${companyName}\n- ウェブサイトドメイン: ${domain || '不明'}\n# 収集項目\n- 事業内容 (business_description)\n- 法人番号 (corporate_number)\n- 最新の採用情報や採用ページのURL (recruitment_info)\n- 直近1年以内の主要なプレスリリースやニュースの要約 (latest_news)\n- 本社の住所 (address)\n- 公式ウェブサイトのURL (website_url)\n# 出力形式 (JSON)\n見つからない情報は "不明" としてください。\n{\n  "business_description": "...",\n  "corporate_number": "...",\n  "recruitment_info": "...",\n  "latest_news": "...",\n  "address": "...",\n  "website_url": "..."\n}`;

    try {
        this.geminiClient.setPromptText(prompt);
        const response = await this.geminiClient.generateCandidates();

        if (!response.candidates || response.candidates.length === 0 || !response.candidates[0].content.parts[0].text) {
          throw new Error('Gemini APIから有効なテキスト応答がありませんでした。');
        }
        
        const responseText = response.candidates[0].content.parts[0].text;
        const jsonString = responseText.replace(/```json\n?/, '').replace(/```$/, '');
        const result = JSON.parse(jsonString);
        
        return {
          '事業内容': result.business_description,
          '法人番号': result.corporate_number,
          '採用情報': result.recruitment_info,
          '最新ニュース': result.latest_news,
          '住所': result.address,
          'ウェブサイト': result.website_url,
        };
    } catch (error) {
        Logger.log(`Geminiでの情報収集またはJSONパース中にエラーが発生しました (会社名: ${companyName}): ${error.stack}`);
        return null;
    }
  }

  /**
   * 収集したデータでAppSheetのレコードを更新します。
   */
  async _updateAccountInAppSheet(accountId, data) {
    // lib_AppSheetAPIの仕様に基づき、updateRecordsではなくeditRecordsを使用するべきかもしれません。
    // ここでは、updateRecordsが存在すると仮定します。もしなければeditRecordsに変更してください。
    const rowToUpdate = { id: accountId, ...data };
    await this.appSheetClient.updateRecords('Account', [rowToUpdate]);
  }
  
  /**
   * アカウントのステータスのみを更新します。
   */
  async _updateAccountStatus(accountId, status) {
    try {
      await this._updateAccountInAppSheet(accountId, { enrichment_status: status });
    } catch (error) {
      Logger.log(`Account ID [${accountId}] のステータス更新中にエラーが発生しました: ${error.stack}`);
    }
  }
}
