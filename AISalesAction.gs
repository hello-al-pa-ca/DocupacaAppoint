/**
 * =================================================================
 * AI Sales Action (v21.4 - フォールバック処理追加)
 * =================================================================
 * 【v21.4での主な変更点】
 * - _getAIRoleDescription: AIの役割名で検索して見つからない場合、
 * 渡された文字列が長ければ、それを役割定義そのものとして扱う
 * フォールバック処理を追加しました。これにより、AppSheet側から
 * 役割名ではなく説明が渡された場合でもエラーなく動作します。
 *
 * 【v21.3での主な変更点】
 * - SalesCopilotクラスに、欠落していた `_getAIRoleDescription` 関数を追加。
 *
 * 【v21.2での主な変更点】
 * - `_loadSheetData` 関数を追加し、初期化エラーを修正。
 * - デフォルトモデルを 'gemini-2.5-flash' に更新。
 * =================================================================
 */

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
  DEFAULT_MODEL: 'gemini-2.5-flash',
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

  /**
   * スプレッドシートからデータを読み込み、オブジェクトの配列に変換します。
   * @param {string} spreadsheetId - スプレッドシートのID。
   * @param {string} sheetName - シート名。
   * @returns {Object[]} - データの配列。
   * @private
   */
  _loadSheetData(spreadsheetId, sheetName) {
    try {
      const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`シートが見つかりません: ${sheetName}`);
      }
      const data = sheet.getDataRange().getValues();
      const headers = data.shift(); // 最初の行をヘッダーとして取得
      return data.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      });
    } catch (e) {
      Logger.log(`シートデータの読み込み中にエラーが発生しました (${sheetName}): ${e.message}`);
      return []; // エラーが発生した場合は空の配列を返す
    }
  }

  // ▼▼▼【修正点】役割名が見つからない場合のフォールバック処理を追加 ▼▼▼
  /**
   * AIの役割名から説明を取得します。
   * @param {string} roleName - AIの役割名または説明文。
   * @returns {string | null} - AIの役割の説明。見つからない場合はnull。
   * @private
   */
  _getAIRoleDescription(roleName) {
    if (!roleName) return null;
    
    // 1. まず、役割名で完全に一致するものを探す (本来の動作)
    const role = this.aiRoles.find(r => r.name === roleName); 
    if (role) {
      return role.description;
    }

    // 2. 見つからず、渡された文字列が長い場合、それを説明文自体とみなす (フォールバック)
    if (roleName.length > 50) { // 50文字を「長い」と判断する閾値
      Logger.log("AI役割名での検索に失敗。渡されたテキスト自体を役割定義として使用します。");
      return roleName;
    }
    
    // 3. 短い文字列で見つからない場合は、定義がないものとしてnullを返す
    return null;
  }

  async executeAISalesAction(recordId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, organizationId, referenceUrls) {
    try {
      // --- 事前準備 (ここは変更なし) ---
      const currentAction = await this._findRecordById('SalesAction', recordId);
      if (!currentAction) throw new Error(`SalesActionレコードが見つかりません (ID: ${recordId})`);
      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      const organizationRecord = organizationId ? await this._findRecordById('Organization', organizationId) : null;
      const accountRecord = accountId ? await this._findRecordById('Account', accountId) : null;
      const effectiveCompanyName = accountRecord ? accountRecord.name : companyName;
      const effectiveAddress = accountRecord?.address || companyAddress;
      const effectiveWebsiteUrl = accountRecord?.website_url;
      const historySummary = accountId ? await this._summarizePastActions(accountId, recordId) : '';
      const { processedAddPrompt, referenceContent, markdownLinkList } = this._processUrlInputs(addPrompt, referenceUrls);
      const companyInfoResult = effectiveCompanyName ? await this._getCompanyInfo(effectiveCompanyName, effectiveAddress, effectiveWebsiteUrl) : null;
      const companyInfoForPrompt = companyInfoResult ? this._formatCompanyInfoForPrompt(companyInfoResult.structuredData) : '';
      const searchSourcesMarkdown = companyInfoResult ? companyInfoResult.sourcesMarkdown : '';
      
      const finalPrompt = this._buildFinalPrompt(
        { ...accountRecord, name: effectiveCompanyName, address: effectiveAddress }, 
        { name: customerContactName }, 
        organizationRecord,
        companyInfoForPrompt,
        historySummary,
        referenceContent,
        processedAddPrompt
      );

      const geminiClient = new GeminiClient(this.geminiModel);
      geminiClient.setSystemInstructionText(aiRoleDescription);
      geminiClient.setPromptText(finalPrompt);
      geminiClient.promptContents.generationConfig = {
        ...geminiClient.promptContents.generationConfig,
        "responseMimeType": "application/json"
      };

      const response = await this._apiCallWithRetry(async () => await geminiClient.generateCandidates());
      const responseText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      if (!responseText) throw new Error('Geminiからの応答が空でした。');

      let proposals = [];
      try {
        proposals = JSON.parse(responseText);
        if (!Array.isArray(proposals) || proposals.length === 0) throw new Error();
      } catch (e) {
        throw new Error(`AIの応答が期待したJSON配列形式ではありません: ${responseText}`);
      }
      
      await this._saveProposals(recordId, proposals);
      Logger.log(`[SUCCESS] ${proposals.length}件のAI提案を ActionProposal テーブルに保存しました。`);

      const updatePayloadForSalesAction = {
        "execute_ai_status": "提案済み",
        "suggest_ai_text": `AIが${proposals.length}パターンの提案を生成しました。\n\n` + searchSourcesMarkdown,
        "link_markdown": markdownLinkList
      };
      await this._updateAppSheetRecord('SalesAction', recordId, updatePayloadForSalesAction);
      Logger.log(`処理完了: SalesAction ID ${recordId} を更新しました。`);
      
    } catch (e) {
      Logger.log(`❌ AI提案生成エラー: ${e.message}\n${e.stack}`);
      throw e;
    }
  }

  async _saveProposals(salesActionId, proposals) {
    const recordsToCreate = proposals.map(p => ({
      sales_action_id: salesActionId,
      proposal_type: p.proposal_type || '不明',
      subject: p.subject || '',
      body: p.body || '',
      is_selected: false
    }));
    
    await this.appSheetClient.addRecords('ActionProposal', recordsToCreate, this.execUserEmail);
  }

  _buildFinalPrompt(account, contact, organization, latestInfo, history, reference, note) {
    const prompt = `
# 指示
あなたは、中小企業の社長に営業メールの文面を提案する、非常に優秀な「AI営業秘書」です。
以下の情報を基に、3つの異なる戦略的アプローチに基づいたメール文案を生成してください。

# 提案すべき3つの「型」
1.  **A. 王道で攻める型（信頼性重視）**: 相手企業の公式な発表（新サービス、プレスリリース等）を祝福し、信頼関係の構築を目指す。
2.  **B. 共感で心をつかむ型（課題直結）**: 相手の発信（SNS等）から個人的な悩みを見つけ出し、共感から入ることで心理的距離を縮める。
3.  **C. 合理性で刺す型（時間節約）**: 多忙な相手のため、結論から単刀直入にメリットを提示する。

# 提供情報
- ## 宛先企業情報
  - 会社名: ${account.name}
  - 住所: ${account.address || '不明'}
  - 事業内容: ${account.company_description || '不明'}
  - 最新の動向(リアルタイム検索結果): ${latestInfo || '特記事項なし'}

- ## 宛先担当者情報
  - 氏名: ${contact.name}

- ## 差出人(自社)情報
  - 会社名: ${organization.name || '株式会社ハロー！アルパカ'}
  - 自社サービス: ${organization.products || 'AIによる営業支援ツール'}

- ## その他補足情報
  - 過去のやり取りの要約: ${history || '特になし'}
  - 添付・参考資料の概要: ${reference || '特になし'}
  - 担当者からの指示・メモ: ${note || '特になし'}

# 出力形式
- 必ず、以下のJSON配列形式で回答してください。
- 各オブジェクトは、提案の「型」、メールの「件名(subject)」、「本文(body)」を含めてください。
- 本文は、読みやすさを考慮し、Markdown（**太字**や箇条書き）を使用してください。
- **重要: 本文は、多忙な社長がスマートフォンで読みやすいよう、それぞれ300文字程度に収まるように、簡潔に記述してください。**
- JSON以外の説明文や前置きは一切不要です。

[
  {
    "proposal_type": "A. 王道で攻める型",
    "subject": "件名をここに記述",
    "body": "本文をここに記述"
  },
  {
    "proposal_type": "B. 共感で心をつかむ型",
    "subject": "件名をここに記述",
    "body": "本文をここに記述"
  },
  {
    "proposal_type": "C. 合理性で刺す型",
    "subject": "件名をここに記述",
    "body": "本文をここに記述"
  }
]
`;
    return prompt.trim();
  }

  // ... (その他のヘルパー関数は変更なし) ...
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
}
