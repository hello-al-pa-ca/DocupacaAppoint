/**
 * @fileoverview Accountテーブルの企業情報を定期的に自動収集・更新するバッチ処理スクリプト
 * @version 1.2.0
 * @description
 * 30分おきにトリガーで実行されることを想定。
 * 'Account'テーブルで'enrichment_status'が'Pending'のレコードを取得し、
 * AIとGoogle検索を使って企業情報をリッチ化して書き戻す。
 *
 * 【v1.2.0での主な変更点】
 * - AppSheet APIがレコード0件の場合に文字列を返す問題に対応。
 * APIからの応答が配列でない場合は空の配列を返すように修正し、エラーを解消しました。
 */

// =================================================================
// グローバル設定
// =================================================================
const BATCH_PROCESSING_LIMIT = 5; // 1回の実行で処理する最大アカウント数（APIの負荷を考慮）

// =================================================================
// メイン実行関数 (トリガーから実行)
// =================================================================

/**
 * 【トリガー実行用】企業情報のバッチ処理を開始します。
 */
async function runBatchAccountEnrichment() {
  try {
    const execUserEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT');
    if (!execUserEmail) {
      throw new Error("スクリプトプロパティ 'SERVICE_ACCOUNT' が設定されていません。");
    }
    
    Logger.log(`企業情報のバッチ更新処理を開始します... (実行アカウント: ${execUserEmail})`);
    
    const enricher = new AccountEnricher(execUserEmail);
    await enricher.enrichPendingAccounts();
    
    Logger.log('企業情報のバッチ更新処理が正常に完了しました。');
  } catch (e) {
    Logger.log(`❌ バッチ処理中に致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 【開発・テスト用】KnowledgeBaseManagerの処理を直接実行します。
 */
async function test_AccountEnrichment() {
    try {
      const execUserEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT') || Session.getActiveUser().getEmail();
      if (!execUserEmail) {
        throw new Error("実行アカウントのメールアドレスが取得できませんでした。");
      }
      
      Logger.log(`テスト実行を開始します... (実行アカウント: ${execUserEmail})`);
      
      const enricher = new AccountEnricher(execUserEmail);
      await enricher.enrichPendingAccounts();
      
      Logger.log('テスト実行が完了しました。');
    } catch (e) {
      Logger.log(`テスト実行中にエラーが発生しました: ${e.message}\n${e.stack}`);
    }
}


/**
 * @class AccountEnricher
 * @description アカウント情報の収集と更新ロジックを管理するクラス。
 */
class AccountEnricher {
  /**
   * @param {string} execUserEmail 実行ユーザーのメールアドレス
   */
  constructor(execUserEmail) {
    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);
  }

  /**
   * 未処理のアカウント情報をリッチ化するプロセス全体を実行します。
   */
  async enrichPendingAccounts() {
    // 1. AppSheetから処理対象のアカウントを抽出
    const pendingAccounts = await this._findPendingAccounts();
    if (!pendingAccounts || pendingAccounts.length === 0) {
      Logger.log("情報収集対象のアカウントはありませんでした。");
      return;
    }
    Logger.log(`${pendingAccounts.length}件のアカウントを処理します。`);

    // 2. 各アカウントの情報を収集・更新
    for (const account of pendingAccounts) {
      // ループの開始時にIDの存在を必ず確認
      if (!account || !account.id) {
        Logger.log(`処理対象のレコードが無効か、IDが含まれていません。スキップします: ${JSON.stringify(account)}`);
        continue;
      }

      try {
        const enrichedData = await this._enrichAccountData(account.name);
        
        let payload = {};
        if (enrichedData && Object.keys(enrichedData).length > 0) {
          payload = { 
            ...enrichedData,
            enrichment_status: 'Completed' // ステータスを完了に更新
          };
          await this.updateAccountRecord(account.id, payload);
          Logger.log(`アカウント [${account.name}] の情報を更新しました。`);
        } else {
          // 情報が取得できなかった場合もステータスは更新する
          payload = { enrichment_status: 'Failed' };
          await this.updateAccountRecord(account.id, payload);
          Logger.log(`アカウント [${account.name}] の情報が見つからなかったため、ステータスを'Failed'に更新しました。`);
        }
        
        Utilities.sleep(1000); // APIへの連続リクエストを避けるための待機
      } catch (e) {
        Logger.log(`アカウント [${account.name}] (ID: ${account.id}) の処理中にエラーが発生しました: ${e.message}`);
        // エラーが発生したレコードのステータスを更新
        try {
          await this.updateAccountRecord(account.id, { enrichment_status: 'Error' });
        } catch (updateError) {
          Logger.log(`エラーステータスの更新にも失敗しました (Account ID: ${account.id}): ${updateError.message}`);
        }
      }
    }
  }

  /**
   * 'enrichment_status'が'Pending'のアカウントレコードを取得します。
   * @private
   * @returns {Promise<Object[]>} 未処理のアカウントのレコード配列
   */
  async _findPendingAccounts() {
    const selector = `FILTER("Account", [enrichment_status] = "Pending")`;
    const results = await this.appSheetClient.findData('Account', this.execUserEmail, { 
      "Selector": selector,
      "Properties": {
        "PageSize": BATCH_PROCESSING_LIMIT, // 一度に処理する件数を制限
      }
     });
    
    // ★★★ エラー修正: AppSheet APIがレコード0件の場合に文字列を返すことがあるため、応答が配列でない場合は空の配列を返す ★★★
    if (Array.isArray(results)) {
        return results;
    }
    Logger.log(`処理対象のアカウントが見つかりませんでした（AppSheetからの応答が配列ではありません）。応答: ${results}`);
    return [];
  }

  /**
   * AIとGoogle検索を使い、企業情報を収集します。
   * @private
   * @param {string} companyName - 調査対象の企業名。
   * @returns {Promise<Object>} - 収集した企業情報を含むオブジェクト。
   */
  async _enrichAccountData(companyName) {
    if (!companyName) {
      console.warn("会社名がないため、企業情報の自動収集をスキップします。");
      return {};
    }
  
    try {
      console.log(`企業情報の自動収集を開始: ${companyName}`);
      const prompt = `日本の企業「${companyName}」について、公開情報を基に以下の情報を調査し、JSON形式で回答してください。不明な項目は空文字("")にしてください。
- company_description: 事業内容の3行程度の要約
- corporate_number: 国税庁の法人番号
- website_url: 公式サイトのURL
- main_service: 主要な製品またはサービス名
- target_audience: ターゲット顧客層（例: B2B, B2C, 中小企業向け）
- is_hiring: 現在、採用活動を積極的に行っているか (Yes/No)
- hiring_roles: 特に募集している職種（3つまでカンマ区切り）
- recent_news: 直近3ヶ月の重要なニュースやプレスリリースの要約（1件）と情報源URL
- last_signal_type: 上記ニュースが[資金調達, 新サービス発表, 役員交代, 業務提携]のどれに該当するか
`;
      
      const GEMINI_API_MODEL = 'gemini-1.5-flash-latest';
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_API_MODEL}:generateContent`;
    
      const payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "tools": [{"google_search_retrieval": {}}]
      };
    
      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };
    
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();
    
      if (responseCode !== 200) {
        throw new Error(`Gemini API request failed with status ${responseCode}: ${responseBody}`);
      }

      const result = JSON.parse(responseBody);
      const textResponse = (result.candidates[0].content.parts || []).map(p => p.text).join('');
      const cleanedJsonString = textResponse.replace(/^```json\s*|```\s*$/g, '').trim();
      const enrichedData = JSON.parse(cleanedJsonString);

      // last_signal_datetimeを追加
      if (enrichedData.recent_news) {
        enrichedData.last_signal_datetime = new Date().toISOString();
      }
      // approach_recommendedを追加
      if (enrichedData.is_hiring && enrichedData.is_hiring.toLowerCase() === 'yes') {
        enrichedData.approach_recommended = true;
      } else {
        enrichedData.approach_recommended = false;
      }

      console.log(`情報収集の結果: ${JSON.stringify(enrichedData)}`);
      return enrichedData;

    } catch (e) {
      console.error(`企業情報の自動収集中にエラーが発生しました (${companyName}): ${e.stack}`);
      return {};
    }
  }

  /**
   * 指定されたIDのAccountレコードを更新します。
   * @param {string} accountId - 更新するアカウントのID。
   * @param {Object} fieldsToUpdate - 更新するフィールドのキーと値。
   * @returns {Promise<Object>} - AppSheet APIからの応答。
   */
  async updateAccountRecord(accountId, fieldsToUpdate) {
    if(!accountId) {
      throw new Error("updateAccountRecordがIDなしで呼び出されました。");
    }
    const recordData = { 
      id: accountId, 
      ...fieldsToUpdate 
    };
    return await this.appSheetClient.updateRecords('Account', [recordData], this.execUserEmail);
  }
}
