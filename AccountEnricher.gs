/**
 * =================================================================
 * AccountEnricher (v6.7 - ハルシネーション対策版)
 * =================================================================
 * v6.6をベースに、AIが事実に基づかない情報を生成する「ハルシネーション」を
 * 抑制するための修正を加えました。
 *
 * 【v6.7での主な変更点】
 * - _enrichWithAI()内のプロンプトを修正。
 * - 「検索ツールが提供する情報のみを基に回答すること」
 * 「推測に基づく情報を含めないこと」という厳格なルールを追加。
 * - これにより、生成される情報の事実性が向上します。
 * =================================================================
 */

// =================================================================
// 定数宣言
// =================================================================
const ENRICHER_CONSTANTS = {
  PROPS_KEY: {
    ENRICHER_APPSHEET_APP_ID: 'ENRICHER_APPSHEET_APP_ID',
    ENRICHER_APPSHEET_API_KEY: 'ENRICHER_APPSHEET_API_KEY',
    SERVICE_ACCOUNT: 'SERVICE_ACCOUNT',
    GEMINI_MODEL: 'GEMINI_MODEL',
  },
  DEFAULT_MODEL: 'gemini-2.0-flash',
  TABLE: {
    ACCOUNT: 'Account',
  },
  COLUMN: {
    ID: 'id',
    NAME: 'name',
    WEBSITE_URL: 'website_url',
    ADDRESS: 'address',
    STATUS: 'enrichment_status',
    APPROACH_RECOMMENDED: 'approach_recommended',
  },
  DATE_COLUMNS: [
    'last_signal_datetime',
    'establishment_date',
    'foundation_date',
    'representative_birth_date',
  ],
  STATUS: {
    PENDING: 'Pending',
    COMPLETED: 'Completed',
    FAILED: 'Failed',
    SKIPPED: 'Skipped',
  },
  BATCH_PROCESSING_LIMIT: 10,
  TRIGGER_FUNCTION_NAME: 'runAccountEnrichmentBatch' // トリガーで呼び出す関数名
};


// =================================================================
// グローバル関数 (トリガーまたは手動で実行)
// =================================================================

/**
 * 企業情報収集バッチのメイン関数。
 * 手動または時間主導型トリガーで実行します。
 * 処理が残っている場合、1分後に自分自身を呼び出すトリガーをセットします。
 */
function runAccountEnrichmentBatch() {
  _deleteTriggersByName(ENRICHER_CONSTANTS.TRIGGER_FUNCTION_NAME);

  const execUserEmail = PropertiesService.getScriptProperties().getProperty(ENRICHER_CONSTANTS.PROPS_KEY.SERVICE_ACCOUNT);
  
  if (!execUserEmail) {
    Logger.log(`❌ エラー: 実行ユーザーのメールアドレスが設定されていません。スクリプトプロパティ '${ENRICHER_CONSTANTS.PROPS_KEY.SERVICE_ACCOUNT}' を確認してください。`);
    return;
  }
  Logger.log(`[START] 企業情報収集バッチをサービスアカウント (${execUserEmail}) で開始します。`);
  
  try {
    const enricher = new AccountEnricher(execUserEmail);
    enricher.enrichAllPendingAccounts().then(() => {
      Logger.log("1バッチ分の処理が完了しました。残りの処理があるか確認します...");
      const remainingAccounts = enricher._findPendingAccounts();
      remainingAccounts.then(accounts => {
        if (accounts && accounts.length > 0) {
          Logger.log(`[CONTINUE] 未処理のレコードが${accounts.length}件残っています。1分後に次のバッチ処理を予約します。`);
          ScriptApp.newTrigger(ENRICHER_CONSTANTS.TRIGGER_FUNCTION_NAME)
            .timeBased()
            .after(1 * 60 * 1000) // 1分後
            .create();
        } else {
          Logger.log(`[COMPLETE] ✅ すべてのレコードの処理が完了しました。バッチ処理を終了します。`);
        }
      });
    }).catch(e => {
      Logger.log(`❌ バッチ処理の実行中にエラーが発生しました: ${e.message}\n${e.stack}`);
    });
  } catch (e) {
    Logger.log(`❌ 初期化エラー: ${e.message}\n${e.stack}`);
  }
}

function enrichSingleAccount(accountId, execUserEmail) {
  if (!accountId || !execUserEmail) {
    Logger.log(`❌ [ERROR] 引数が不足しています。accountId: ${accountId}, execUserEmail: ${execUserEmail}`);
    return;
  }
  Logger.log(`[START] アカウント個別更新を開始します。Account ID: ${accountId}, 実行者: ${execUserEmail}`);
  try {
    const enricher = new AccountEnricher(execUserEmail);
    enricher.processSingleAccount(accountId).catch(e => Logger.log(`❌ 個別更新エラー: ${e.message}\n${e.stack}`));
  } catch (e) {
    Logger.log(`❌ 初期化エラー: ${e.message}\n${e.stack}`);
  }
}


// =================================================================
// AccountEnricher クラス
// =================================================================

class AccountEnricher {
  constructor(execUserEmail) {
    this.execUserEmail = execUserEmail;
    this.props = PropertiesService.getScriptProperties().getProperties();
    
    const appId = this.props[ENRICHER_CONSTANTS.PROPS_KEY.ENRICHER_APPSHEET_APP_ID];
    const apiKey = this.props[ENRICHER_CONSTANTS.PROPS_KEY.ENRICHER_APPSHEET_API_KEY];
    
    if (!appId || !apiKey) {
      throw new Error("Enricher専用のApp IDまたはAPIキーがスクリプトプロパティに設定されていません。('ENRICHER_APPSHEET_APP_ID', 'ENRICHER_APPSHEET_API_KEY')");
    }

    this.appSheetClient = new AppSheetClient(appId, apiKey);
    Logger.log(`✅ AccountEnricherの初期化完了 (接続先AppID: ${appId})`);
  }

  async processSingleAccount(accountId) {
    try {
      Logger.log(`[1/4] Account ID [${accountId}] のレコード情報を取得中...`);
      const account = await this._findRecordById(ENRICHER_CONSTANTS.TABLE.ACCOUNT, accountId);
      if (!account) throw new Error(`レコードが見つかりませんでした。`);
      Logger.log(`  -> ✅ 取得成功。`);

      const companyName = account[ENRICHER_CONSTANTS.COLUMN.NAME];
      if (!companyName) {
        Logger.log(`[SKIP] ID: ${accountId} には会社名がないためスキップします。`);
        await this._updateAccountStatus(accountId, ENRICHER_CONSTANTS.STATUS.SKIPPED);
        return;
      }
      
      Logger.log(`[2/4] 会社名 [${companyName}] の情報をAIで調査中...`);
      const websiteUrl = account[ENRICHER_CONSTANTS.COLUMN.WEBSITE_URL];
      const address = account[ENRICHER_CONSTANTS.COLUMN.ADDRESS];
      const enrichedData = await this._enrichWithAI(companyName, address, websiteUrl);
      
      if (enrichedData) {
        Logger.log(`  -> ✅ AIからの情報取得成功。`);
        Logger.log(`[3/4] 取得データをAppSheet用に整形(サニタイズ)中...`);
        const sanitizedData = this._sanitizeDataForAppSheet(enrichedData);
        Logger.log(`  -> ✅ 整形完了。`);

        Logger.log(`[4/4] AppSheetのレコードを更新中...`);
        
        sanitizedData[ENRICHER_CONSTANTS.COLUMN.STATUS] = ENRICHER_CONSTANTS.STATUS.COMPLETED;
        await this._updateAccountInAppSheet(accountId, sanitizedData);
        Logger.log(`[SUCCESS] ✅ Account ID [${accountId}] の情報更新が正常に完了しました。`);

      } else {
        await this._updateAccountStatus(accountId, ENRICHER_CONSTANTS.STATUS.FAILED);
        Logger.log(`[FAIL] AIからの情報収集に失敗しました。ステータスを'Failed'に更新します。`);
      }
    } catch (error) {
      Logger.log(`❌ [ERROR] Account ID [${accountId}] の処理中にエラー: ${error.stack}`);
      await this._updateAccountStatus(accountId, ENRICHER_CONSTANTS.STATUS.FAILED).catch(e => Logger.log(`  -> ⚠️ ステータス更新にも失敗: ${e.message}`));
    }
  }

  async _findPendingAccounts() {
    const selector = `TOP(FILTER("${ENRICHER_CONSTANTS.TABLE.ACCOUNT}", [${ENRICHER_CONSTANTS.COLUMN.STATUS}] = "${ENRICHER_CONSTANTS.STATUS.PENDING}"), ${ENRICHER_CONSTANTS.BATCH_PROCESSING_LIMIT})`;
    
    const properties = { "Selector": selector };
    
    try {
      const results = await this.appSheetClient.findData(ENRICHER_CONSTANTS.TABLE.ACCOUNT, this.execUserEmail, properties);
      return (results && Array.isArray(results)) ? results : [];
    } catch (e) {
      Logger.log(`❌ [ERROR] AppSheetからのデータ取得に失敗しました: ${e.message}`);
      throw e;
    }
  }

  async _enrichWithAI(companyName, address, websiteUrl) {
    // ★★★ 修正点: ハルシネーションを抑制するルールを追加 ★★★
    const prompt = `
      あなたはプロの企業調査アナリストです。
      以下の企業について、公開情報から徹底的に調査し、指定されたJSON形式で回答してください。
      このデータは日本のビジネスユーザー向けのアプリケーションで利用されるため、回答の品質が非常に重要です。

      # 調査対象企業
      - 会社名: ${companyName}
      - 所在地ヒント: ${address || '不明'}
      - URLヒント: ${websiteUrl || '不明'}

      # 収集項目とルール
      - 【最重要ルール】: 必ずGoogle検索ツールが提供する情報のみを基に回答してください。検索結果に存在しない情報や、推測に基づく情報を回答に含めることは固く禁じます。情報が見つからない場合は、その項目には必ず null を設定してください。
      - 【言語ルール】: すべての回答は、必ず自然で流暢な日本語で記述してください。英語、ロシア語(例: основ)、韓国語(例: 다양한)など、日本語以外の言語や不自然な記号を絶対に混ぜないでください。
      - 【欠損データ】: 見つからない情報は null を返してください。
      - 【日付形式】: 日付に関する項目は「YYYY-MM-DD」形式で回答してください。
      - 【フラグ形式】: "approach_recommended" には、「はい」か「いいえ」のいずれか一つだけを回答してください。

      # 出力形式 (JSONのみを回答)
      {
        "industry": "...", "company_size": "...", "company_description": "...", "corporate_number": "...",
        "website_url": "...", "linkedin_url": "...", "main_service": "...", "target_audience": "...",
        "intent_keyword": "...", "last_signal_type": "...", "last_signal_datetime": "...", "last_signal_summary": "...",
        "approach_recommended": "はい", "funding_ir_info": "...", "business_strategy": "...", "hiring_info": "...",
        "tech_stack": "...", "customer_case_studies": "...", "event_info": "...", "listing_status": "...",
        "capital_stock": "...", "establishment_date": "...", "foundation_date": "...", "legal_entity_type": "...",
        "representative_name": "...", "representative_title": "...", "representative_birth_date": "...",
        "representative_background": "...", "representative_career": "...", "shareholder_composition": "...",
        "main_suppliers": "...", "main_customers": "...", "facilities_overview": "...", "company_overview": "...",
        "business_strengths": "...", "business_weaknesses": "...", "future_outlook": "..."
      }
      
      # 最終確認
      生成したJSONの各値が、上記のルール（特に最重要ルール）に従っていることを必ず確認してください。`;

    let responseText = '';
    try {
      const geminiModel = this.props[ENRICHER_CONSTANTS.PROPS_KEY.GEMINI_MODEL] || ENRICHER_CONSTANTS.DEFAULT_MODEL;
      const localGeminiClient = new GeminiClient(geminiModel);
      
      localGeminiClient.enableGoogleSearchTool();
      localGeminiClient.setPromptText(prompt);
      const response = await localGeminiClient.generateCandidates();
      responseText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      
      let jsonString = '';
      const jsonRegex = /```(json)?\s*([\s\S]*?)\s*```/;
      const match = responseText.match(jsonRegex);

      if (match && match[2]) {
        jsonString = match[2];
      } else {
        const firstBrace = responseText.indexOf('{');
        const lastBrace = responseText.lastIndexOf('}');
        if (firstBrace !== -1 && lastBrace > firstBrace) {
          jsonString = responseText.substring(firstBrace, lastBrace + 1);
        } else {
            throw new Error("AIの応答から有効なJSONオブジェクトを抽出できませんでした。");
        }
      }
      
      return JSON.parse(jsonString);

    } catch (error) {
      Logger.log(`❌ [ERROR] Geminiでの情報収集またはJSONパース中にエラー: ${error.stack}`);
      Logger.log(`  -> AIからの生の応答テキスト: \n${responseText}`);
      return null;
    }
  }

  _sanitizeDataForAppSheet(data) {
    const sanitized = {};
    for (const key in data) {
      if (data[key] === '不明' || data[key] === null) {
        sanitized[key] = null;
      } else {
        sanitized[key] = data[key];
      }
    }

    const yesNoKey = ENRICHER_CONSTANTS.COLUMN.APPROACH_RECOMMENDED;
    if (sanitized.hasOwnProperty(yesNoKey)) {
        const originalValue = sanitized[yesNoKey];
        sanitized[yesNoKey] = (originalValue === 'はい');
    }

    ENRICHER_CONSTANTS.DATE_COLUMNS.forEach(key => {
        if (sanitized[key]) {
            sanitized[key] = this._formatDateString(sanitized[key]);
        }
    });

    sanitized.website_url = this._formatUrl(sanitized.website_url);
    sanitized.linkedin_url = this._formatUrl(sanitized.linkedin_url);
    
    return sanitized;
  }
  
  _formatUrl(urlString) {
    if (!urlString || typeof urlString !== 'string' || urlString.trim().toLowerCase() === 'null' || urlString.trim() === '') return null;
    let trimmedUrl = urlString.trim();
    if (!/^https?:\/\//i.test(trimmedUrl)) {
      trimmedUrl = `https://${trimmedUrl}`;
    }
    try {
      new URL(trimmedUrl);
      return trimmedUrl;
    } catch (_) {
      return null;
    }
  }
  
  _formatDateString(dateString) {
      if (!dateString || typeof dateString !== 'string') return null;
      
      const ymdMatch = dateString.match(/(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?/);
      if (ymdMatch) {
          const year = ymdMatch[1];
          const month = ymdMatch[2].padStart(2, '0');
          const day = ymdMatch[3].padStart(2, '0');
          const d = new Date(`${year}-${month}-${day}`);
          if (!isNaN(d.getTime())) {
            return `${year}/${month}/${day}`;
          }
      }
      
      Logger.log(`[WARN] 無効な日付形式の値のため、nullに変換します: "${dateString}"`);
      return null;
  }

  async _updateAccountInAppSheet(accountId, data) {
    const rowToUpdate = {
      [ENRICHER_CONSTANTS.COLUMN.ID]: accountId,
      ...data
    };
    Logger.log(`  -> 🔄 AppSheetへの更新ペイロード:\n${JSON.stringify(rowToUpdate, null, 2)}`);
    await this.appSheetClient.updateRecords(ENRICHER_CONSTANTS.TABLE.ACCOUNT, [rowToUpdate], this.execUserEmail);
  }
  
  async _updateAccountStatus(accountId, status) {
    try {
      Logger.log(`[INFO] Account ID [${accountId}] のステータスを "${status}" に更新します。`);
      await this._updateAccountInAppSheet(accountId, { [ENRICHER_CONSTANTS.COLUMN.STATUS]: status });
    } catch (error) {
      Logger.log(`❌ [ERROR] Account ID [${accountId}] のステータス更新中にエラー: ${error.stack}`);
    }
  }

  async enrichAllPendingAccounts() {
    Logger.log("⏳ 保留中のアカウントを検索中...");
    try {
        const pendingAccounts = await this._findPendingAccounts();

        if (!pendingAccounts || pendingAccounts.length === 0) {
            Logger.log("✅ 今回のバッチで処理するアカウントはありませんでした。");
            return;
        }

        Logger.log(`[INFO] ${pendingAccounts.length}件のアカウントの情報収集を開始します。`);

        for (const [index, account] of pendingAccounts.entries()) {
            Logger.log(`[BATCH] ${index + 1} / ${pendingAccounts.length} 件目の処理を開始します...`);
            await this.processSingleAccount(account[ENRICHER_CONSTANTS.COLUMN.ID]);
            
            if (index < pendingAccounts.length - 1) {
                const delay = 3000;
                Logger.log(`[PAUSE] 次の処理まで ${delay / 1000} 秒待機します...`);
                Utilities.sleep(delay);
            }
        }
        
        Logger.log(`[END BATCH] 今回のバッチ処理(${pendingAccounts.length}件)が完了しました。`);
    } catch (e) {
        Logger.log(`❌ [ERROR] 保留中アカウントの処理中にエラーが発生しました: ${e.message}`);
        throw e;
    }
  }

  async _findRecordById(tableName, recordId) {
    const keyColumn = ENRICHER_CONSTANTS.COLUMN.ID;
    const selector = `FILTER("${tableName}", [${keyColumn}] = "${recordId}")`;
    const properties = { "Selector": selector };
    const result = await this.appSheetClient.findData(tableName, this.execUserEmail, properties);
    if (result && Array.isArray(result) && result.length > 0) {
      return result[0];
    }
    Logger.log(`[WARN] テーブル[${tableName}]からID[${recordId}]のレコードが見つかりませんでした。`);
    return null;
  }
}

/**
 * 指定された名前のトリガーをすべて削除するヘルパー関数
 */
function _deleteTriggersByName(triggerFunctionName) {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === triggerFunctionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`古いトリガー (ID: ${trigger.getUniqueId()}) を削除しました。`);
    }
  }
}
