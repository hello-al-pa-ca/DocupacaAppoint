/**
 * =================================================================
 * AI Sales Action (v20.1 - 仕様変更版)
 * =================================================================
 * AISalesAction実行時にAccountテーブルを更新する際、上書きする情報を
 * 「最新動向」に関する項目のみに限定しました。
 *
 * 【v20.1での主な変更点】
 * - `executeAISalesAction`内のロジックを修正。
 * - AIが企業全体を調査した後、`last_signal_summary`, `hiring_info`など
 * 最新動向に関するカラムだけを抜き出して更新用のペイロードを作成。
 * - これにより、手動で入力した基本情報が保護されます。
 * =================================================================
 */

// =================================================================
// 定数宣言
// =================================================================
const AISALESACTION_CONSTANTS = {
  MASTER_SHEET_NAMES: {
    actionCategories: 'ActionCategory',
    aiRoles: 'AIRole',
    salesFlows: 'ActionFlow'
  },
  RETRY_CONFIG: {
    count: 3,
    delay: 2000
  },
  PROPS_KEY: {
    GEMINI_MODEL: 'AISALESACTION_MODEL',
    APPSHEET_APP_ID: 'APPSHEET_APP_ID',
    APPSHEET_API_KEY: 'APPSHEET_API_KEY',
    MASTER_SHEET_ID: 'MASTER_SHEET_ID',
    GOOGLE_API_KEY: 'GOOGLE_API_KEY',
  },
  DEFAULT_MODEL: 'gemini-2.5-flash-preview-05-20',
  // ★ v20.1 修正点: 更新対象とする「最新動向」カラムのリスト
  TREND_COLUMNS_TO_UPDATE: [
    'last_signal_summary',
    'last_signal_type',
    'last_signal_datetime',
    'approach_recommended',
    'intent_keyword',
    'hiring_info',
    'event_info'
  ]
};


// =================================================================
// グローバル関数 (AppSheetまたは手動で実行)
// =================================================================

function executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName = '', companyAddress = '', customerContactName = '', ourContactName = '', probability = '', eventName = '', referenceUrls = '', execUserEmail) {
  if (!execUserEmail) {
    // ... (エラーハンドリングは変更なし)
    return;
  }
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.executeAISalesAction(recordId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, organizationId, referenceUrls)
      .catch(e => {
        Logger.log(`❌ executeAISalesActionの非同期実行中にエラー: ${e.message}\n${e.stack}`);
        copilot._updateAppSheetRecord('SalesAction', recordId, { "execute_ai_status": "エラー", "suggest_ai_text": `処理エラー: ${e.message}` });
      });
  } catch (e) {
    Logger.log(`❌ executeAISalesActionで致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

function suggestNextAction(completedActionId, execUserEmail) {
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.suggestNextAction(completedActionId).catch(e => Logger.log(`❌ suggestNextActionエラー: ${e.message}\n${e.stack}`));
  } catch (e) {
    Logger.log(`❌ suggestNextActionで致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}


// =================================================================
// SalesCopilot クラス (メインのアプリケーションロジック)
// =================================================================

class SalesCopilot {
  constructor(execUserEmail) {
    if (!execUserEmail) throw new Error("実行ユーザーのメールアドレスは必須です。");

    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.appSheetClient = new AppSheetClient(this.props[AISALESACTION_CONSTANTS.PROPS_KEY.APPSHEET_APP_ID], this.props[AISALESACTION_CONSTANTS.PROPS_KEY.APPSHEET_API_KEY]);
    
    this.geminiModel = this.props[AISALESACTION_CONSTANTS.PROPS_KEY.GEMINI_MODEL] || AISALESACTION_CONSTANTS.DEFAULT_MODEL;
    Logger.log(`[INFO] SalesCopilot initialized with model: ${this.geminiModel}`);

    const masterSheetId = this.props[AISALESACTION_CONSTANTS.PROPS_KEY.MASTER_SHEET_ID];
    if (!masterSheetId) throw new Error("マスターシートのIDがスクリプトプロパティに設定されていません。");

    this.actionCategories = this._loadSheetData(masterSheetId, AISALESACTION_CONSTANTS.MASTER_SHEET_NAMES.actionCategories);
    this.aiRoles = this._loadSheetData(masterSheetId, AISALESACTION_CONSTANTS.MASTER_SHEET_NAMES.aiRoles);
    this.salesFlows = this._loadSheetData(masterSheetId, AISALESACTION_CONSTANTS.MASTER_SHEET_NAMES.salesFlows);
  }

  async executeAISalesAction(recordId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, organizationId, referenceUrls) {
    try {
      // ... (事前準備のコードは変更なし) ...
      const currentAction = await this._findRecordById('SalesAction', recordId);
      if (!currentAction) throw new Error(`SalesActionレコードが見つかりません (ID: ${recordId})`);
      const actionDetails = this._getActionDetails(actionName, contactMethod);
      if (!actionDetails) throw new Error(`アクション定義が見つかりません: ${actionName}/${contactMethod}`);
      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      const customerId = accountId;
      const organizationRecord = organizationId ? await this._findRecordById('Organization', organizationId) : null;
      const accountRecord = customerId ? await this._findRecordById('Account', customerId) : null;
      const effectiveCompanyName = accountRecord ? accountRecord.name : companyName;
      const effectiveAddress = accountRecord?.address || companyAddress;
      const effectiveWebsiteUrl = accountRecord?.website_url;
      const historySummary = customerId ? await this._summarizePastActions(customerId, recordId) : '';
      const { processedAddPrompt, referenceContent, markdownLinkList } = this._processUrlInputs(addPrompt, referenceUrls);
      
      // AIによる企業調査を実行 (ここは変更なし)
      const companyInfoResult = effectiveCompanyName ? await this._getCompanyInfo(effectiveCompanyName, effectiveAddress, effectiveWebsiteUrl) : null;
      
      // =================================================================
      // ▼▼▼【v20.1 修正点】更新用ペイロードを「最新動向」のみに限定 ▼▼▼
      // =================================================================
      let trendUpdatePayload = null;
      if (companyInfoResult && companyInfoResult.structuredData) {
        trendUpdatePayload = {};
        for (const key of AISALESACTION_CONSTANTS.TREND_COLUMNS_TO_UPDATE) {
            if (companyInfoResult.structuredData.hasOwnProperty(key)) {
                trendUpdatePayload[key] = companyInfoResult.structuredData[key];
            }
        }
      }
      
      const companyInfoForPrompt = companyInfoResult ? this._formatCompanyInfoForPrompt(companyInfoResult.structuredData) : '';
      const searchSourcesMarkdown = companyInfoResult ? companyInfoResult.sourcesMarkdown : '';
      
      // ... (プレースホルダー設定、最終プロンプト構築、AIへのリクエストは変更なし) ...
      const placeholders = {
        '[顧客の会社名]': effectiveCompanyName, '[取引先会社名]': effectiveCompanyName, '[企業名]': effectiveCompanyName,
        '[会社の住所]': effectiveAddress, '[取引先担当者名]': customerContactName, '[取引先氏名]': customerContactName,
        '[自社担当者名]': ourContactName, '[自社名]': organizationRecord ? organizationRecord.name : '株式会社ハロー！アルパカ',
        '[契約の確度]': probability, '[イベント名]': eventName, '[商談メモの内容を加味した、1言メッセージ]': processedAddPrompt,
        '[参考資料リンク]': markdownLinkList
      };
      const finalPrompt = this._buildFinalPrompt(mainPrompt || actionDetails.prompt, placeholders, contactMethod, probability, accountRecord, organizationRecord, companyInfoForPrompt, referenceContent, historySummary);
      const geminiClient = new GeminiClient(this.geminiModel);
      geminiClient.setSystemInstructionText(aiRoleDescription);
      geminiClient.setPromptText(finalPrompt);
      const response = await this._apiCallWithRetry(async () => await geminiClient.generateCandidates());
      const generatedText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      if (!generatedText) throw new Error('Geminiからの応答が空でした。');

      // SalesActionテーブルの更新 (ここは変更なし)
      const formattedData = this._formatResponse(generatedText, contactMethod);
      const updatePayloadForSalesAction = {
        "suggest_ai_text": formattedData.suggest_ai_text + searchSourcesMarkdown, "subject": formattedData.subject,
        "body": formattedData.body, "execute_ai_status": "提案済み", "link_markdown": markdownLinkList
      };
      await this._updateAppSheetRecord('SalesAction', recordId, updatePayloadForSalesAction);
      Logger.log(`処理完了 (AI提案生成): SalesAction ID ${recordId} を更新しました。`);
      
      // =================================================================
      // ▼▼▼【v20.1 修正点】限定されたペイロードでAccountテーブルを更新 ▼▼▼
      // =================================================================
      if (trendUpdatePayload && accountId) {
        trendUpdatePayload.enrichment_status = 'Completed'; //ステータスは更新
        await this._updateAppSheetRecord('Account', accountId, trendUpdatePayload);
        Logger.log(`処理完了 (最新動向更新): Account ID ${accountId} の最新動向を更新しました。`);
      }

    } catch (e) {
      Logger.log(`❌ AI提案生成エラー: ${e.message}\n${e.stack}`);
      throw e;
    }
  }

  /**
   * 次のアクションを提案します。
   */
  async suggestNextAction(completedActionId) {
    try {
      const completedAction = await this._findRecordById('SalesAction', completedActionId);
      if (!completedAction) throw new Error(`ID ${completedActionId} のアクションが見つかりません。`);
      
      const accountId = completedAction.accountId;
      if(!accountId) {
        Logger.log(`警告: 完了アクション[${completedActionId}]にアカウントIDが紐付いていません。`);
        return;
      }

      const nextActionFlow = this._getActionFlowDetails(completedAction['progress'], completedAction['action_name'], completedAction['result']);
      if (!nextActionFlow) {
        await this._updateAppSheetRecord('SalesAction', completedActionId, {"next_action_description": "営業フロー完了"});
        return;
      }

      const nextActionDetails = this._findNextActionInfo(nextActionFlow.next_action);
      await this._updateAppSheetRecord('SalesAction', completedActionId, {
        "next_action_category_id": nextActionDetails.id,
        "next_action_description": nextActionDetails.description
      });
    } catch (e) {
      Logger.log(`❌ 次アクション提案エラー: ${e.message}\n${e.stack}`);
      throw e;
    }
  }

  /**
   * 過去のアクション履歴を要約します。
   */
  async _summarizePastActions(customerId, currentActionId) {
    const task = async () => {
      Logger.log(`顧客ID [${customerId}] の過去の商談履歴の要約を開始します。`);
      const selector = `FILTER("SalesAction", AND([accountId] = "${customerId}", [ID] <> "${currentActionId}"))`;
      const pastActions = await this.appSheetClient.findData('SalesAction', this.execUserEmail, { "Selector": selector });

      if (!pastActions || pastActions.length === 0) {
        return "";
      }

      const historyText = pastActions
        .sort((a, b) => new Date(a.executed_dt) - new Date(b.executed_dt))
        .map(action => `日時: ${action.executed_dt}\nアクション: ${action.action_name}\nメモ: ${action.addPrompt || ''}\n結果: ${action.result || ''}\nAI提案: ${action.body || ''}`)
        .join('\n\n---\n\n');

      const summarizationPrompt = `以下の商談履歴の要点を、重要なポイントを3行程度でまとめてください。\n\n--- 履歴 ---\n${historyText}`;
      
      const summarizerClient = new GeminiClient(this.geminiModel);
      summarizerClient.setPromptText(summarizationPrompt);
      const response = await summarizerClient.generateCandidates();
      return (response.candidates[0].content.parts || []).map(p => p.text).join('');
    };

    try {
      return await this._apiCallWithRetry(task, "商談履歴の要約");
    } catch (e) {
      Logger.log(`商談履歴の要約中にエラーが発生しました: ${e.message}`);
      return "";
    }
  }
  
  /**
   * Google検索を使い、企業情報と参照元URLを取得します。
   */
  async _getCompanyInfo(companyName, address, websiteUrl) {
    const task = async () => {
      const apiKey = this.props[AISALESACTION_CONSTANTS.PROPS_KEY.GOOGLE_API_KEY];
      if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
        Logger.log('⚠️ 企業情報のリアルタイム検索はスキップされました。スクリプトプロパティに「GOOGLE_API_KEY」が設定されていません。');
        return null;
      }

      const researchPrompt = `
        あなたはプロの企業調査アナリストです。
        以下の企業について、公開情報から徹底的に調査し、指定されたJSON形式で回答してください。

        # 調査対象企業
        - 会社名: ${companyName}
        - 所在地ヒント: ${address || '不明'}
        - URLヒント: ${websiteUrl || '不明'}

        # 収集項目と出力形式 (JSON)
        - company_description: 事業内容の包括的な説明
        - main_service: 主要な製品やサービスの概要
        - hiring_info: 現在の採用情報、特に強化している職種の要約
        - last_signal_summary: 上記以外の最新ニュースやプレスリリース
        
        もし、企業の特定が困難な場合は、その旨をJSONの各値に含めてください。
        回答はJSONオブジェクトのみとし、前後に説明文などを加えないでください。
        {
          "company_description": "...", "main_service": "...", "hiring_info": "...", "last_signal_summary": "..."
        }
      `;

      const researchClient = new GeminiClient(this.geminiModel);
      researchClient.enableGoogleSearchTool();
      researchClient.setPromptText(researchPrompt);
      const response = await researchClient.generateCandidates();
      
      const responseText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      let structuredData = null;
      try {
        const jsonMatch = responseText.match(/{[\s\S]*}/);
        if (jsonMatch) {
          structuredData = JSON.parse(jsonMatch[0]);
        } else {
           throw new Error("AIの応答から有効なJSONを抽出できませんでした。");
        }
      } catch(e) {
         Logger.log(`企業情報のJSONパース中にエラー: ${e.message}`);
         return null;
      }
      
      let sourcesMarkdown = '';
      const attributions = response.candidates[0].groundingAttributions;
      if (attributions && attributions.length > 0) {
        const sources = attributions.map(attr => attr.web).filter(web => web && web.uri).slice(0, 5);
        if (sources.length > 0) {
            sourcesMarkdown = "\n\n---\n\n**▼ 調査情報のソース**\n";
            sources.forEach((source, index) => {
                sourcesMarkdown += `${index + 1}. [${source.title || source.uri}](${source.uri})\n`;
            });
        }
      }
      
      Logger.log(`企業情報の調査結果(JSON):\n${JSON.stringify(structuredData, null, 2)}`);
      return { structuredData, sourcesMarkdown };
    };
    
    try {
      return await this._apiCallWithRetry(task, "企業情報検索");
    } catch (e) {
      Logger.log(`企業情報の調査中にエラーが発生しました: ${e.message}`);
      return null;
    }
  }

  _formatCompanyInfoForPrompt(structuredData) {
    if (!structuredData) return '';
    let text = '';
    if (structuredData.company_description) text += `- 事業内容: ${structuredData.company_description}\n`;
    if (structuredData.main_service) text += `- 主要サービス: ${structuredData.main_service}\n`;
    if (structuredData.hiring_info) text += `- 採用情報: ${structuredData.hiring_info}\n`;
    if (structuredData.last_signal_summary) text += `- 最新動向: ${structuredData.last_signal_summary}\n`;
    return text;
  }
  
  _buildFinalPrompt(template, placeholders, contactMethod, probability, accountRecord, organizationRecord, companyInfoFromSearch, referenceContent, historySummary) {
    
    let toneInstruction = '';
    let currentProbability = probability || 'C';
    switch (currentProbability) {
      case 'A':
        toneInstruction = '自信を持って、次のアポイント獲得を強く意識した文面を作成してください。貴社のお役に立てると確信している、という熱意を伝えてください。';
        break;
      case 'B':
        toneInstruction = '相手の関心を引きつつ、丁寧に関係を構築するような、少し強めの文面を作成してください。お役に立てる「かもしれない」という、丁寧ながらも積極的な姿勢を示してください。';
        break;
      case 'C':
      case 'D':
      default:
        toneInstruction = 'まずはご挨拶と情報提供を主目的とした、丁寧で控えめな文面を作成してください。売り込みの色合いは極力なくし、今後の関係構築のきっかけ作りを意識してください。';
        break;
    }

    let filledTemplate = template.replace(/\[[^\]]+\]/g, (match) => {
        return (placeholders[match] !== undefined && placeholders[match] !== null) ? placeholders[match] : match;
    });
    
    if (!placeholders['[イベント名]']) {
      filledTemplate = filledTemplate.replace(/\[イベント名\]では（イベント名が空白の場合はここは削除）、/g, '');
    }
    filledTemplate = filledTemplate.replace(/\[[^\]]+\]/g, ''); 

    let additionalInfo = '\n\n【補足情報】\nこの情報を最大限に活用し、下記の指示に従って、具体的でパーソナライズされた文章を作成してください。\n';
    let hasInfo = false;
    
    if (organizationRecord) {
      additionalInfo += `--- 自社情報 ---\n`;
      if (organizationRecord.name) additionalInfo += `- 組織名: ${organizationRecord.name}\n`;
      if (organizationRecord.hp_link) additionalInfo += `- ホームページ: ${organizationRecord.hp_link}\n`;
      if (organizationRecord.category) additionalInfo += `- 業種: ${organizationRecord.category}\n`;
      if (organizationRecord.characteristics) additionalInfo += `- 特徴: ${organizationRecord.characteristics}\n`;
      if (organizationRecord.products) additionalInfo += `- 主要製品: ${organizationRecord.products}\n`;
      hasInfo = true;
    }

    if (accountRecord) {
      additionalInfo += `--- 企業情報（DBより取得） ---\n`;
      additionalInfo += `- 事業内容: ${accountRecord.company_description || '未登録'}\n`;
      additionalInfo += `- 最新の動向: ${accountRecord.last_signal_summary || '未登録'}\n`;
      hasInfo = true;
    }
    
    if (companyInfoFromSearch) {
      additionalInfo += `\n--- 企業調査情報（リアルタイム検索） ---\n${companyInfoFromSearch}\n`;
      hasInfo = true;
    }
    
    if (placeholders['[取引先担当者名]']) {
      additionalInfo += `- 宛先担当者名: ${placeholders['[取引先担当者名]']}\n`; hasInfo = true;
    }
    if (placeholders['[自社担当者名]']) {
      additionalInfo += `- 差出人担当者名: ${placeholders['[自社担当者名]']}\n`; hasInfo = true;
    }
    if (currentProbability) {
      additionalInfo += `- 現在の契約確度: ${currentProbability}\n`;
      hasInfo = true;
    }

    if (referenceContent) {
      additionalInfo += `\n--- 参考資料・引継ぎ資料の内容 ---\n${referenceContent}\n`;
      hasInfo = true;
    }
    if (placeholders['[参考資料リンク]']) {
      additionalInfo += `\n--- 利用可能な参考資料リンク ---\n${placeholders['[参考資料リンク]']}\n`;
      hasInfo = true;
    }
    if (historySummary) {
      additionalInfo += `\n--- これまでの商談履歴の要約 ---\n${historySummary}\n`;
      hasInfo = true;
    }
    
    let finalInstruction = '';
    if (contactMethod === 'メール') {
      finalInstruction = `\n\n【重要】\n- 以下の【メール本文の骨子】と【補足情報】を基に、完成されたメール文章を、【件名】と【本文】の形式で生成してください。`;
      finalInstruction += `\n- 全体のトーンは、補足情報にある「現在の契約確度: ${currentProbability}」を考慮し、「${toneInstruction}」という指示に従ってください。`;
      finalInstruction += `\n- **【最優先事項】「企業調査情報（リアルタイム検索）」の結果を最優先で参考にして、具体的でタイムリーな内容を盛り込んでください。DBの情報と異なる場合は、必ずリアルタイム検索の結果を使用してください。**`;
      finalInstruction += `\n- 本文は、読みやすさを向上させるため、必要に応じて太字（**テキスト**）や箇条書き（- テキスト）などのMarkdown形式で記述してください。`;
      finalInstruction += `\n- **テンプレート内のプレースホルダーは、補足情報を使って必ず具体的な内容に置き換えてください。** 最終的な文章に[]が残らないようにしてください。`;
      finalInstruction += `\n- 「利用可能な参考資料リンク」セクションに記載されているMarkdownリンクは、すべて本文中に自然な形で含めてください。`;
      finalInstruction += `\n- ★提供された情報以外のURL（例: https://example.com）は、絶対に生成しないでください。★`;
      finalInstruction += `\n- 件名は簡潔で分かりやすくしてください。`;
      finalInstruction += `\n- 生成する文章以外の解説や、確度に応じた文章の調整案などは一切含めないでください。`;
    }

    const finalPrompt = `【メール本文の骨子】\n${filledTemplate}` + (hasInfo ? additionalInfo : "") + finalInstruction;
    
    return finalPrompt;
  }

  async _updateAppSheetRecord(tableName, recordId, fieldsToUpdate) {
    const recordData = (tableName === 'Account' || tableName === 'Organization') 
      ? { id: recordId, ...fieldsToUpdate }
      : { ID: recordId, ...fieldsToUpdate };
    return await this.appSheetClient.updateRecords(tableName, [recordData], this.execUserEmail);
  }

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

  async _apiCallWithRetry(apiCallFunction, taskName = 'API呼び出し') {
    let lastError;
    for (let i = 0; i < AISALESACTION_CONSTANTS.RETRY_CONFIG.count; i++) {
      try {
        return await apiCallFunction();
      } catch (e) {
        lastError = e;
        if (e.message && (e.message.includes('status 50') || e.message.includes('Service invoked too many times'))) {
          const delay = AISALESACTION_CONSTANTS.RETRY_CONFIG.delay * Math.pow(2, i);
          Logger.log(`🔁 ${taskName}で一時的なエラーが発生しました (試行 ${i + 1}/${AISALESACTION_CONSTANTS.RETRY_CONFIG.count})。${delay}ms後に再試行します。エラー: ${e.message}`);
          Utilities.sleep(delay);
        } else {
          throw lastError;
        }
      }
    }
    Logger.log(`❌ ${taskName}のリトライがすべて失敗しました。`);
    throw lastError;
  }

  _processUrlInputs(addPrompt, referenceUrls) {
    const combinedUrlsString = [addPrompt, referenceUrls].filter(Boolean).join(',');
    const urlRegex = /https?:\/\/(?:drive|docs)\.google\.com\/(?:file|document|spreadsheets|presentation)\/d\/([a-zA-Z0-9_-]{28,})/g;
    const uniqueUrls = [...new Set(combinedUrlsString.match(urlRegex) || [])];
    if (uniqueUrls.length === 0) {
      return { processedAddPrompt: addPrompt, referenceContent: '', markdownLinkList: '' };
    }
    let processedAddPromptText = addPrompt || '';
    let referenceContentText = '';
    const markdownLinkArray = [];
    uniqueUrls.forEach(url => {
      try {
        const fileId = this._extractFileIdFromUrl(url);
        if (!fileId) return;
        const file = DriveApp.getFileById(fileId);
        const fileName = file.getName();
        const markdownLink = `[${fileName}](${url})`;
        markdownLinkArray.push(markdownLink);
        const textContent = this._extractTextFromFile(file);
        if (textContent) {
          referenceContentText += `--- 参考資料: ${fileName} ---\n${textContent}\n\n`;
        }
        const urlPattern = new RegExp(url.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
        if(processedAddPromptText.match(urlPattern)){
            processedAddPromptText = processedAddPromptText.replace(urlPattern, markdownLink);
        }
      } catch (e) {
        Logger.log(`URL処理中のエラー [${url}]: ${e.message}`);
      }
    });
    return {
      processedAddPrompt: processedAddPromptText,
      referenceContent: referenceContentText,
      markdownLinkList: markdownLinkArray.join('\n')
    };
  }
  _extractFileIdFromUrl(url) {
    if (!url) return null;
    const match = url.match(/\/d\/([a-zA-Z0-9_-]{28,})/);
    return match ? match[1] : null;
  }
  _extractTextFromFile(file) {
    const mimeType = file.getMimeType();
    const fileName = file.getName();
    Logger.log(`ファイルからテキストを抽出中: ${fileName} (MIME Type: ${mimeType})`);
    try {
      if (mimeType.startsWith('video/')) {
        return `（ファイル名: 「${fileName}」の動画ファイル）`;
      }
      switch (mimeType) {
        case MimeType.GOOGLE_DOCS:
          return DocumentApp.openById(file.getId()).getBody().getText();
        case MimeType.GOOGLE_SHEETS:
          const sheet = SpreadsheetApp.openById(file.getId());
          return sheet.getSheets().map(s => {
            const sheetName = s.getName();
            const data = s.getDataRange().getValues().map(row => row.join(', ')).join('\n');
            return `シート名: ${sheetName}\n${data}`;
          }).join('\n\n');
        case MimeType.GOOGLE_SLIDES:
          const presentation = SlidesApp.openById(file.getId());
          return presentation.getSlides().map((slide, index) => {
            const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
            const slideText = slide.getShapes().map(shape => shape.getText().asString()).join(' ');
            return `スライド ${index + 1}:\n${slideText}\nノート: ${notes}`;
          }).join('\n\n');
        case MimeType.PLAIN_TEXT:
        case 'text/csv':
          return file.getBlob().getDataAsString('UTF-8');
        case 'application/pdf':
          if (Drive.Files) { 
            Logger.log(`PDFのOCR処理を開始します: ${fileName}`);
            const tempDoc = Drive.Files.insert({ title: `temp_ocr_${Utilities.getUuid()}` }, file.getBlob(), { ocr: true, ocrLanguage: 'ja' });
            const text = DocumentApp.openById(tempDoc.id).getBody().getText();
            Drive.Files.remove(tempDoc.id); 
            return text;
          } else {
            Logger.log("PDFの読み込みにはDrive APIの有効化が必要です。");
            return '';
          }
        default:
          Logger.log(`サポートされていないMIMEタイプのためスキップ: ${mimeType}`);
          return `（ファイル名: 「${fileName}」、種類: ${mimeType}）`;
      }
    } catch (e) {
      Logger.log(`ファイルからのテキスト抽出中にエラーが発生しました: ${fileName}, Error: ${e.message}`);
      return '';
    }
  }
  _formatResponse(rawText, contactMethod) {
    return this._splitSubjectAndBody(rawText, contactMethod);
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
  _splitSubjectAndBody(text, contactMethod) {
    const response = { "suggest_ai_text": text, "subject": "", "body": text };
    if (contactMethod !== 'メール') return response;
    const subjectMarker = '【件名】';
    const bodyMarker = '【本文】';
    const subjectIndex = text.indexOf(subjectMarker);
    if (subjectIndex !== -1) {
      const bodyIndex = text.indexOf(bodyMarker, subjectIndex);
      let subjectText = '';
      let bodyText = '';
      if (bodyIndex !== -1) {
        subjectText = text.substring(subjectIndex + subjectMarker.length, bodyIndex).trim();
        bodyText = text.substring(bodyIndex + bodyMarker.length).trim();
      } else {
        const lines = text.substring(subjectIndex + subjectMarker.length).trim().split('\n');
        subjectText = lines[0] || '';
        bodyText = lines.slice(1).join('\n').trim();
      }
      response.subject = subjectText.replace(/[\r\n]/g, ' ').trim();
      response.body = bodyText;
    } else {
      const lines = text.trim().split('\n');
      if (lines.length > 1 && lines[0].length < 50 && !lines[0].includes('様')) {
        response.subject = lines[0].trim();
        response.body = lines.slice(1).join('\n').trim();
      }
    }
    return response;
  }
  _getActionDetails(actionName, contactMethod) {
    return this.actionCategories.find(row => row.action_name === actionName && row.contact_method === contactMethod) || null;
  }
  _getAIRoleDescription(roleName) {
    const role = this.aiRoles.find(row => row.name === roleName);
    return role ? role.description : `あなたは優秀な「${roleName}」です。`;
  }
   _getActionFlowDetails(currentProgress, currentActionName, currentResult) {
    return this.salesFlows.find(row => row.progress === currentProgress && row.action_id === currentActionName && row.result === currentResult) || null;
  }
  _findNextActionInfo(nextActionName) {
    const defaultContactMethod = 'メール';
    const nextAction = this.actionCategories.find(row => row.action_name === nextActionName && row.contact_method === defaultContactMethod) || this.actionCategories.find(row => row.action_name === nextActionName);
    if (nextAction && nextAction.id) {
      return { id: nextAction.id, description: `${nextAction.action_name} (${nextAction.contact_method}) を実施してください。` };
    }
    return { id: null, description: `推奨アクション: ${nextActionName}` };
  }
}
