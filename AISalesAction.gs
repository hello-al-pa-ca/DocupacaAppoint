/**
 * =================================================================
 * AI Sales Action (リファクタリング版 v15)
 * =================================================================
 * v14の改善に加え、AppSheetのレコードが見つからない404エラーを防止する
 * チェック処理を追加し、「引継ぎ資料」の情報をAIがより活用できるように
 * プロンプトを改善しました。
 *
 * 【v15での主な変更点】
 * - 処理開始時にレコードIDの存在チェックを追加し、404エラーを未然に防ぎます。
 * - 「引継ぎ資料」の内容をAIがより重視するようにプロンプトを修正。
 * =================================================================
 */

// =================================================================
// 定数宣言
// =================================================================
const MASTER_SHEET_NAMES = {
  actionCategories: 'ActionCategory',
  aiRoles: 'AIRole',
  salesFlows: 'ActionFlow'
};

// =================================================================
// グローバル関数 (AppSheetまたは手動で実行)
// =================================================================

/**
 * 【AppSheetから実行】AIによる文章生成のメインプロセスを開始します。
 */
function executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName = '', companyAddress = '', customerContactName = '', ourContactName = '', probability = '', eventName = '', referenceUrls = '', execUserEmail) {
  
  const functionArgs = {
    recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, 
    addPrompt, companyName, companyAddress, customerContactName, ourContactName, 
    probability, eventName, referenceUrls, execUserEmail
  };
  Logger.log(`executeAISalesAction が以下の引数で呼び出されました: \n${JSON.stringify(functionArgs, null, 2)}`);

  if (!execUserEmail) {
    const errorMessage = "実行ユーザーのメールアドレス(execUserEmail)が渡されませんでした。AppSheetのBot設定で引数にUSEREMAIL()が正しく設定されているか確認してください。";
    Logger.log(`❌ ${errorMessage}`);
    try {
      const props = PropertiesService.getScriptProperties().getProperties();
      const client = new AppSheetClient(props.APPSHEET_APP_ID, props.APPSHEET_API_KEY);
      const errorPayload = {
        "ID": recordId,
        "execute_ai_status": "エラー",
        "suggest_ai_text": errorMessage
      };
      client.updateRecords('SalesAction', [errorPayload], null); // App Ownerとして実行
    } catch (updateError) {
      Logger.log(`❌ エラーステータスの更新に失敗しました: ${updateError.message}`);
    }
    return;
  }

  try {
    const copilot = new SalesCopilot(execUserEmail);
    // 非同期処理を呼び出し、エラーはcatchで補足
    copilot.executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, organizationId, referenceUrls)
      .catch(e => {
        Logger.log(`❌ executeAISalesActionの非同期実行中にエラー: ${e.message}\n${e.stack}`);
        // エラーが発生した場合も、ステータスを更新
        copilot._updateAppSheetRecord(recordId, { "execute_ai_status": "エラー", "suggest_ai_text": `処理エラー: ${e.message}` });
      });
  } catch (e) {
    Logger.log(`❌ executeAISalesActionで致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 【AppSheetから実行】完了したアクションに基づき、次のアクションを提案します。
 */
function suggestNextAction(completedActionId, execUserEmail) {
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.suggestNextAction(completedActionId)
      .catch(e => {
         Logger.log(`❌ suggestNextActionの非同期実行中にエラー: ${e.message}\n${e.stack}`);
      });
  } catch (e) {
    Logger.log(`❌ suggestNextActionで致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}


// =================================================================
// SalesCopilot クラス (メインのアプリケーションロジック)
// =================================================================

class SalesCopilot {
  constructor(execUserEmail) {
    if (!execUserEmail) {
      throw new Error("SalesCopilotの初期化に失敗: 実行ユーザーのメールアドレスは必須です。");
    }

    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);
    
    this.geminiModel = 'gemini-2.0-flash'; 

    const masterSheetId = this.props.MASTER_SHEET_ID;
    if (!masterSheetId) throw new Error("マスターシートのIDがスクリプトプロパティに設定されていません。");

    this.actionCategories = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.actionCategories);
    this.aiRoles = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.aiRoles);
    this.salesFlows = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.salesFlows);
  }

  /**
   * AIによる営業アクションの文章を生成します。
   */
  async executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, organizationId, referenceUrls) {
    try {
      // ★★★ 修正点: 処理開始時にレコードの存在を確認 ★★★
      const currentAction = await this._findRecordById('SalesAction', recordId);
      if (!currentAction) {
        throw new Error(`指定されたSalesActionレコードが見つかりません (ID: ${recordId})。AppSheet側でレコードが作成されているか、IDが正しいか確認してください。`);
      }
      
      const actionDetails = this._getActionDetails(actionName, contactMethod);
      if (!actionDetails) throw new Error(`アクション定義が見つかりません: ${actionName}/${contactMethod}`);

      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      
      const customerId = currentAction.取引先ID;

      const organizationRecord = organizationId ? await this._findRecordById('Organization', organizationId) : null;
      if (organizationId && !organizationRecord) {
        Logger.log(`警告: 組織ID [${organizationId}] に対応する組織情報が見つかりませんでした。`);
      }

      const accountRecord = customerId ? await this._findRecordById('Account', customerId) : null;
      if (customerId && !accountRecord) {
        Logger.log(`警告: 取引先ID [${customerId}] に対応するアカウント情報が見つかりませんでした。`);
      }

      const historySummary = customerId ? await this._summarizePastActions(customerId, recordId) : '';

      const { processedAddPrompt, referenceContent, markdownLinkList } = this._processUrlInputs(addPrompt, referenceUrls);
      
      const placeholders = {
        '[顧客の会社名]': companyName,
        '[取引先会社名]': companyName,
        '[企業名]': companyName,
        '[会社の住所]': companyAddress,
        '[取引先担当者名]': customerContactName,
        '[取引先氏名]': customerContactName,
        '[自社担当者名]': ourContactName,
        '[自社名]': organizationRecord ? organizationRecord.name : '株式会社ハロー！アルパカ',
        '[契約の確度]': probability,
        '[イベント名]': eventName,
        '[商談メモの内容を加味した、1言メッセージ]': processedAddPrompt,
        '[参考資料リンク]': markdownLinkList
      };
      
      const companyInfoFromSearch = companyName ? await this._getCompanyInfo(companyName) : '';
      
      const finalPrompt = this._buildFinalPrompt(mainPrompt || actionDetails.prompt, placeholders, contactMethod, probability, accountRecord, organizationRecord, companyInfoFromSearch, referenceContent, historySummary);
      Logger.log(`最終プロンプト: \n${finalPrompt}`);

      const geminiClient = new GeminiClient(this.geminiModel);
      geminiClient.setSystemInstructionText(aiRoleDescription);
      
      geminiClient.setPromptText(finalPrompt);

      const response = await geminiClient.generateCandidates();
      const generatedText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      if (!generatedText) throw new Error('Geminiからの応答が空でした。');

      const formattedData = this._formatResponse(generatedText, contactMethod);
      
      const updatePayload = {
        "suggest_ai_text": formattedData.suggest_ai_text,
        "subject": formattedData.subject,
        "body": formattedData.body,
        "execute_ai_status": "提案済み",
        "link_markdown": markdownLinkList
      };
      
      Logger.log(`更新ペイロード: ${JSON.stringify(updatePayload)}`);
      await this._updateAppSheetRecord(recordId, updatePayload);
      Logger.log(`処理完了 (AI提案生成): Record ID ${recordId}`);

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

      const nextActionFlow = this._getActionFlowDetails(completedAction['progress'], completedAction['action_name'], completedAction['result']);
      if (!nextActionFlow) {
        await this._updateAppSheetRecord(completedActionId, {"next_action_description": "営業フロー完了"});
        return;
      }

      const nextActionDetails = this._findNextActionInfo(nextActionFlow.next_action);
      const updatePayload = {
        "next_action_category_id": nextActionDetails.id,
        "next_action_description": nextActionDetails.description
      };
      await this._updateAppSheetRecord(completedActionId, updatePayload);
    } catch (e) {
      Logger.log(`❌ 次アクション提案エラー: ${e.message}\n${e.stack}`);
      throw e;
    }
  }

  /**
   * 過去のアクション履歴を要約します。
   */
  async _summarizePastActions(customerId, currentActionId) {
    try {
      Logger.log(`顧客ID [${customerId}] の過去の商談履歴の要約を開始します。`);
      const selector = `FILTER("SalesAction", AND([取引先ID] = "${customerId}", [ID] <> "${currentActionId}"))`;
      const pastActions = await this.appSheetClient.findData('SalesAction', this.execUserEmail, { "Selector": selector });

      if (!pastActions || pastActions.length === 0) {
        Logger.log("要約対象の過去のアクションはありませんでした。");
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
      const summary = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      
      Logger.log(`商談履歴の要約:\n${summary}`);
      return summary;

    } catch (e) {
      Logger.log(`商談履歴の要約中にエラーが発生しました: ${e.message}`);
      return "";
    }
  }
  
  /**
   * Google検索を使って企業情報を調査します。
   */
  async _getCompanyInfo(companyName) {
    try {
      const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_API_KEY');
      if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
        Logger.log('⚠️ 企業情報のリアルタイム検索はスキップされました。スクリプトプロパティに「GOOGLE_API_KEY」が設定されていません。');
        return '(リアルタイム企業情報の検索に失敗しました)';
      }

      const researchPrompt = `${companyName}の企業情報について、ウェブサイトや公開情報から以下の点を簡潔にまとめてください。\n- 事業内容\n- 主な製品やサービス\n- 最新のニュースやプレスリリース（1〜2件）`;
      const researchClient = new GeminiClient(this.geminiModel);
      
      researchClient.enableGoogleSearchTool(); 
      
      researchClient.setPromptText(researchPrompt);
      const response = await researchClient.generateCandidates();
      const info = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      Logger.log(`企業情報の調査結果:\n${info}`);
      return info;
    } catch (e) {
      Logger.log(`企業情報の調査中にエラーが発生しました: ${e.message}`);
      return '(リアルタイム企業情報の検索中にエラーが発生しました)';
    }
  }

  /**
   * 最終的なプロンプトを組み立てます。
   */
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

    let filledTemplate = template;
    for (const key in placeholders) {
      if (placeholders[key] !== undefined && placeholders[key] !== null) {
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(escapedKey, 'g');
        filledTemplate = filledTemplate.replace(regex, placeholders[key]);
      }
    }
    
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
      additionalInfo += `--- 企業情報（システムより取得） ---\n`;
      if (accountRecord.company_description) additionalInfo += `- 事業内容: ${accountRecord.company_description}\n`;
      if (accountRecord.main_service) additionalInfo += `- 主な製品/サービス: ${accountRecord.main_service}\n`;
      if (accountRecord.last_signal_summary) additionalInfo += `- 最新の動向: ${accountRecord.last_signal_summary}\n`;
      if (accountRecord.website_url) additionalInfo += `- ウェブサイト: ${accountRecord.website_url}\n`;
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

    if (companyInfoFromSearch) {
      additionalInfo += `\n--- 企業調査情報（Google検索） ---\n${companyInfoFromSearch}\n`;
      hasInfo = true;
    }
    if (referenceContent) {
      // ★★★ 修正点: 「引継ぎ資料」であることを明記 ★★★
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
      finalInstruction += `\n- 本文は、読みやすさを向上させるため、必要に応じて太字（**テキスト**）や箇条書き（- テキスト）などのMarkdown形式で記述してください。`;
      // ★★★ 修正点: 引継ぎ資料の活用を指示 ★★★
      finalInstruction += `\n- 【補足情報】にある「企業情報」「商談履歴の要約」「自社情報」「参考資料・引継ぎ資料の内容」を最優先で参考にし、本文の冒頭で相手が「おっ」と思うような、関心を持っていることが伝わる自然な一文を加えてください。`;
      finalInstruction += `\n- **テンプレート内のプレースホルダーは、補足情報を使って必ず具体的な内容に置き換えてください。** 最終的な文章に[]が残らないようにしてください。`;
      finalInstruction += `\n- 「利用可能な参考資料リンク」セクションに記載されているMarkdownリンクは、すべて本文中に自然な形で含めてください。`;
      finalInstruction += `\n- ★提供された情報以外のURL（例: https://example.com）は、絶対に生成しないでください。★`;
      finalInstruction += `\n- 件名は簡潔で分かりやすくしてください。`;
      finalInstruction += `\n- 生成する文章以外の解説や、確度に応じた文章の調整案などは一切含めないでください。`;
    }

    const finalPrompt = `【メール本文の骨子】\n${filledTemplate}` + (hasInfo ? additionalInfo : "") + finalInstruction;
    
    return finalPrompt;
  }

  /**
   * レコードを更新します。
   */
  async _updateAppSheetRecord(recordId, fieldsToUpdate) {
    const recordData = { "ID": recordId, ...fieldsToUpdate };
    return await this.appSheetClient.updateRecords('SalesAction', [recordData], this.execUserEmail);
  }

  /**
   * レコードをIDで検索します。
   */
  async _findRecordById(tableName, recordId) {
    const selector = `FILTER("${tableName}", [ID] = "${recordId}")`;
    const properties = { "Selector": selector };
    const result = await this.appSheetClient.findData(tableName, this.execUserEmail, properties);
    if (result && Array.isArray(result) && result.length > 0) {
      return result[0];
    }
    Logger.log(`テーブル[${tableName}]からID[${recordId}]のレコードが見つかりませんでした。応答: ${JSON.stringify(result)}`);
    return null;
  }

  // 他のヘルパー関数は変更ないため、元の実装を維持します。
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
