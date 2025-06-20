/**
 * =================================================================
 * AI Sales Action (RAG機能除外・レスポンス整形機能強化版 v19)
 * =================================================================
 * 既存の AISalesAction.gs からRAG (Retrieval-Augmented Generation)
 * に関連する機能をすべて削除し、リファクタリングしたバージョンです。
 *
 * 主な変更点:
 * - SalesCopilotクラスを、AIによる文章生成と次アクション提案のコア機能に特化。
 * - Google検索のロジックを「企業調査→本文生成」の2段階に変更。
 * - addPrompt内のGoogle Drive URLを自動で検出し、Markdownリンクに変換する機能を統合。
 * - 【v19での修正】ご指摘に基づき、execUserEmailが空の場合にセッションから自動取得する
 * フォールバック処理を削除しました。execUserEmailは必須の引数となります。
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
 * 【AppSheetから実行】AIによる文章生成を指示します。
 */
function executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName = '', companyAddress = '', customerContactName = '', ourContactName = '', probability = '', eventName = '', ourCompanyInfoText = '', ourCompanyInfoFileId = '', referenceUrls = '', execUserEmail) {
  
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
    return; // 処理を中断
  }

  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls);
  } catch (e) {
    Logger.log(`❌ executeAISalesActionで致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 【AppSheetから実行】アクションの結果に基づき、次のアクションを提案します。
 */
function suggestNextAction(completedActionId, execUserEmail) {
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.suggestNextAction(completedActionId);
  } catch (e) {
    Logger.log(`❌ suggestNextActionで致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

// =================================================================
// テスト関数 (GASエディタ実行用)
// =================================================================

/**
 * 【エディタ実行用】固定引数でexecuteAISalesActionをテストします。
 */
function test_executeAISalesAction() {
    const recordId = '7FBCF696-7397-49A3-BC8C-7E5E3AB3AAB4'; // テスト用のレコードID
    const AIRoleName = 'AI 営業マン';
    const actionName = 'あいさつ';
    const contactMethod = 'メール';
    const mainPrompt = `初めて連絡する顧客へ、丁寧な自己紹介と簡潔な挨拶のメール文面を作成してください。件名も提案してください。
Output Example
[会社名]
[氏名] 様

株式会社ペーパーカンパニーＡのAlpacaAppSheetです。
[イベント名]では、お忙しい中お名刺交換させていただき、誠にありがとうございました。

[商談メモの内容を加味した、1言メッセージ]

まだまだ小さな会社ではありますが、経営の効率化や、将来を見据えた体制づくりについて、何かお役に立てることがあるかもしれません。

まずは御礼まで。
貴重なご縁をありがとうございました。

今後ともどうぞよろしくお願いいたします。`;
    const addPrompt = `ドキュパカに興味ありとのこと。参考資料はこちらです。https://docs.google.com/document/d/1mCjPNOHvhKLohepguS3bt9E3NEKhNCNVPr7B9MDyPdQ/edit`;
    const companyName = '株式会社テスト';
    const companyAddress = '東京都千代田区1-1-1';
    const customerContactName = '山田 太郎';
    const ourContactName = '鈴木 一郎';
    const probability = 'A';
    const eventName = 'ものづくり産業交流展示会';
    const ourCompanyInfoText = '弊社はAIを活用したドキュメント管理ツール「ドキュパカ」を提供しており、製造業のDXを支援します。主な製品は「ドキュパカ-Lite」「ドキュパカ-Pro」です。';
    const ourCompanyInfoFileId = '';
    const referenceUrls = '';
    const execUserEmail = 'hello@al-pa-ca.com';

    Logger.log("以下のパラメータでテスト実行します:");
    Logger.log({recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls, execUserEmail});

    executeAISalesAction(recordId, '', '', AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls, execUserEmail);
}


// =================================================================
// SalesCopilot クラス (メインのアプリケーションロジック)
// =================================================================
class SalesCopilot {
  /**
   * @constructor
   * @param {string} execUserEmail - 実行ユーザーのメールアドレス
   */
  constructor(execUserEmail) {
    if (!execUserEmail) {
      // このコンストラクタは有効なメールアドレスが渡されることを前提とする
      throw new Error("SalesCopilotの初期化に失敗: 実行ユーザーのメールアドレスは必須です。");
    }

    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);

    const masterSheetId = this.props.MASTER_SHEET_ID;
    if (!masterSheetId) throw new Error("マスターシートのIDがスクリプトプロパティに設定されていません。");

    this.actionCategories = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.actionCategories);
    this.aiRoles = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.aiRoles);
    this.salesFlows = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.salesFlows);
  }

  /**
   * AIによる営業アクションの文章を生成し、AppSheetを更新します。
   */
  executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls) {
    try {
      const actionDetails = this._getActionDetails(actionName, contactMethod);
      if (!actionDetails) throw new Error(`アクション定義が見つかりません: ${actionName}/${contactMethod}`);

      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      
      const processedAddPrompt = this._processAddPromptWithMarkdownLinks(addPrompt);

      const placeholders = {
        '[商談メモの内容を加味した、1言メッセージ]': processedAddPrompt,
        '[具体的な課題]': processedAddPrompt, '[資料名]': processedAddPrompt,
        '[以前話した課題]': processedAddPrompt, '[推測される課題]': processedAddPrompt,
        '[提案書名]': processedAddPrompt, '[提案内容]': processedAddPrompt,
        '[議題]': processedAddPrompt, '[期間]': processedAddPrompt,
        '[顧客の会社名]': companyName, '[企業名]': companyName,
        '[会社の住所]': companyAddress,
        '[取引先担当者名]': customerContactName,
        '[自社担当者名]': ourContactName,
        '[契約の確度]': probability,
        '[イベント名]': eventName,
        '[自社情報]': ourCompanyInfoText
      };
      
      const useGoogleSearch = actionDetails.searchGoogle && companyName;
      let companyInfo = '';
      if (useGoogleSearch) {
        Logger.log(`Google検索を有効にして企業情報を調査します: ${companyName}`);
        companyInfo = this._getCompanyInfo(companyName);
      }
      
      const referenceContent = this._fetchContentFromDriveUrls(referenceUrls);

      const template = mainPrompt || actionDetails.prompt;
      const finalPrompt = this._buildFinalPrompt(template, placeholders, contactMethod, companyInfo, referenceContent);
      Logger.log(`最終プロンプト: \n${finalPrompt}`);

      const geminiClient = new GeminiClient('gemini-1.5-flash-latest');
      geminiClient.setSystemInstructionText(aiRoleDescription);
      
      if (ourCompanyInfoFileId) {
        try {
            const fileBlob = DriveApp.getFileById(ourCompanyInfoFileId).getBlob();
            geminiClient.attachFiles(fileBlob);
            Logger.log(`自社情報ファイルを添付しました: ${fileBlob.getName()}`);
        } catch (e) {
            Logger.log(`ファイルの添付に失敗しました。File ID: ${ourCompanyInfoFileId}, Error: ${e.message}`);
        }
      }
      
      geminiClient.setPromptText(finalPrompt);

      const response = geminiClient.generateCandidates();
      const generatedText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      if (!generatedText) throw new Error('Geminiからの応答が空でした。');

      const formattedData = this._formatResponse(generatedText, useGoogleSearch, contactMethod);
      
      const updatePayload = {
          "suggest_ai_text": formattedData.suggest_ai_text,
          "subject": formattedData.subject,
          "body": formattedData.body,
          "execute_ai_status": "提案済み"
      };
      
      Logger.log(`更新ペイロード: ${JSON.stringify(updatePayload)}`);
      this._updateAppSheetRecord(recordId, updatePayload);
      Logger.log(`処理完了 (AI提案生成): Record ID ${recordId}`);

    } catch (e) {
      Logger.log(`❌ AI提案生成エラー: ${e.message}\n${e.stack}`);
      this._updateAppSheetRecord(recordId, { "execute_ai_status": "エラー", "suggest_ai_text": `処理エラー: ${e.message}` });
    }
  }

  /**
   * 完了したアクションに基づき、次のアクションを提案します。
   */
  suggestNextAction(completedActionId) {
    try {
      const completedAction = this._findRecordById('SalesAction', completedActionId);
      if (!completedAction) throw new Error(`ID ${completedActionId} のアクションが見つかりません。`);

      const nextActionFlow = this._getActionFlowDetails(completedAction['progress'], completedAction['action_name'], completedAction['result']);
      if (!nextActionFlow) {
        this._updateAppSheetRecord(completedActionId, {"next_action_description": "営業フロー完了"});
        return;
      }

      const nextActionDetails = this._findNextActionInfo(nextActionFlow.next_action);
      const updatePayload = {
        "next_action_category_id": nextActionDetails.id,
        "next_action_description": nextActionDetails.description
      };
      this._updateAppSheetRecord(completedActionId, updatePayload);
    } catch (e) {
      Logger.log(`❌ 次アクション提案エラー: ${e.message}\n${e.stack}`);
    }
  }

  // --- プライベートヘルパーメソッド群 ---

  _processAddPromptWithMarkdownLinks(textBlock) {
    if (!textBlock) return '';

    const urlRegex = /https?:\/\/(?:drive|docs)\.google\.com\/(?:file|document|spreadsheets|presentation)\/d\/([a-zA-Z0-9_-]{28,})/g;
    const uniqueUrls = [...new Set(textBlock.match(urlRegex) || [])];
    
    if (uniqueUrls.length === 0) {
        return textBlock;
    }

    let processedText = textBlock;
    uniqueUrls.forEach(url => {
        try {
            const fileId = this._extractFileIdFromUrl(url);
            if (!fileId) return;
            const fileName = DriveApp.getFileById(fileId).getName();
            const markdownLink = `[${fileName}](${url})`;
            const urlPattern = new RegExp(url.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
            processedText = processedText.replace(urlPattern, markdownLink);
        } catch (e) {
            Logger.log(`Error processing URL [${url}] for markdown link: ${e.message}`);
        }
    });

    return processedText;
  }
  
  _fetchContentFromDriveUrls(urlsString) {
      if (!urlsString) return '';
      
      const urls = urlsString.split(',').map(url => url.trim()).filter(url => url);
      let combinedContent = '';
      
      for (const url of urls) {
          try {
              const fileId = this._extractFileIdFromUrl(url);
              if (!fileId) continue;
              
              const file = DriveApp.getFileById(fileId);
              const fileName = file.getName();
              const textContent = this._extractTextFromFile(file);
              
              if (textContent) {
                  combinedContent += `--- 参考資料: ${fileName} ---\n${textContent}\n\n`;
              }
          } catch (e) {
              Logger.log(`URLからのファイル読み込みに失敗しました: ${url}, Error: ${e.message}`);
          }
      }
      return combinedContent;
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
                return '';
        }
      } catch (e) {
        Logger.log(`ファイルからのテキスト抽出中にエラーが発生しました: ${fileName}, Error: ${e.message}`);
        return '';
      }
  }

  _getCompanyInfo(companyName) {
    try {
        const researchPrompt = `${companyName}の企業情報について、ウェブサイトや公開情報から以下の点を簡潔にまとめてください。\n- 事業内容\n- 主な製品やサービス\n- 最新のニュースやプレスリリース（1〜2件）`;
        const researchClient = new GeminiClient('gemini-1.5-flash-latest');
        researchClient.enableGoogleSearchTool();
        researchClient.setPromptText(researchPrompt);
        const response = researchClient.generateCandidates();
        const info = (response.candidates[0].content.parts || []).map(p => p.text).join('');
        Logger.log(`企業情報の調査結果:\n${info}`);
        return info;
    } catch (e) {
        Logger.log(`企業情報の調査中にエラーが発生しました: ${e.message}`);
        return '';
    }
  }

  _formatResponse(rawText, useGoogleSearch, contactMethod) {
    return this._splitSubjectAndBody(rawText, contactMethod);
  }

  _getAppSheetClient() {
    const { APPSHEET_APP_ID, APPSHEET_API_KEY } = this.props;
    if (!APPSHEET_APP_ID || !APPSHEET_API_KEY) throw new Error('AppSheet接続情報(ID/Key)がスクリプトプロパティに設定されていません。');
    return new AppSheetClient(APPSHEET_APP_ID, APPSHEET_API_KEY);
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

    if (contactMethod === 'メール') {
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
          subjectText = lines.find(line => line.trim() !== '') || '';
          const subjectLineIndex = lines.findIndex(line => line === subjectText);
          bodyText = lines.slice(subjectLineIndex + 1).join('\n').trim();
        }
        
        response.subject = subjectText.replace(/[\r\n]/g, ' ').trim();
        response.body = bodyText;

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
    const nextAction = this.actionCategories.find(row => row.action_name === nextActionName && row.contact_method === defaultContactMethod)
                      || this.actionCategories.find(row => row.action_name === nextActionName);
    
    if (nextAction && nextAction.id) {
      return { id: nextAction.id, description: `${nextAction.action_name} (${nextAction.contact_method}) を実施してください。` };
    }
    return { id: null, description: `推奨アクション: ${nextActionName}` };
  }

  _buildFinalPrompt(template, placeholders, contactMethod, companyInfo = '', referenceContent = '') {
    let finalPrompt = template;

    for (const key in placeholders) {
      if (placeholders[key] && finalPrompt.includes(key)) {
        const regex = new RegExp(`\\${key}`, 'g');
        finalPrompt = finalPrompt.replace(regex, placeholders[key]);
      }
    }
    
    let additionalInfo = '\n\n【補足情報】\n';
    let hasInfo = false;
    if (placeholders['[顧客の会社名]']) {
        additionalInfo += `- 宛先企業名: ${placeholders['[顧客の会社名]']}\n`;
        hasInfo = true;
    }
    if (placeholders['[会社の住所]']) {
        additionalInfo += `- 住所: ${placeholders['[会社の住所]']}\n`;
        hasInfo = true;
    }
    if (placeholders['[取引先担当者名]']) {
        additionalInfo += `- 宛先担当者名: ${placeholders['[取引先担当者名]']}\n`;
        hasInfo = true;
    }
    if (placeholders['[自社担当者名]']) {
        additionalInfo += `- 差出人担当者名: ${placeholders['[自社担当者名]']}\n`;
        hasInfo = true;
    }
    if (placeholders['[契約の確度]']) {
        additionalInfo += `- 契約の確度: ${placeholders['[契約の確度]']}\n`;
        hasInfo = true;
    }
    if (placeholders['[イベント名]']) {
        additionalInfo += `- 交換場所/イベント名: ${placeholders['[イベント名]']}\n`;
        hasInfo = true;
    }
    if (placeholders['[自社情報]']) {
        additionalInfo += `\n--- 自社情報 ---\n${placeholders['[自社情報]']}\n`;
        hasInfo = true;
    }
    if (companyInfo) {
        additionalInfo += `\n--- 企業調査情報 ---\n${companyInfo}\n`;
        hasInfo = true;
    }
    if (referenceContent) {
        additionalInfo += `\n--- 参考資料の内容 ---\n${referenceContent}\n`;
        hasInfo = true;
    }


    if(hasInfo){
        finalPrompt += additionalInfo;
    }

    // 空のプレースホルダーが残っている場合、それを含む文章ごと削除するなどのクリーンアップ
    finalPrompt = finalPrompt.replace(/\[[^\]]+\]/g, '');


    if (contactMethod === 'メール') {
        finalPrompt += `\n\n【重要】\n- 必ず【件名】【本文】の形式で、メールやスクリプトの文章だけを生成してください。\n- 件名は簡潔で分かりやすくしてください。\n- 生成する文章以外の解説や、確度に応じた文章の調整案などは一切含めないでください。`;
    }

    return finalPrompt;
  }

  _updateAppSheetRecord(recordId, fieldsToUpdate) {
    const recordData = { "ID": recordId, ...fieldsToUpdate };
    return this.appSheetClient.updateRecords('SalesAction', [recordData], this.execUserEmail);
  }

  _findRecordById(tableName, recordId) {
    const selector = `SELECT([ID], [ID] = "${recordId}")`;
    const properties = { "Selector": selector };
    const result = this.appSheetClient.findData(tableName, this.execUserEmail, properties);
    if (result && result.length > 0) return result[0];
    return null;
  }
}

/**
 * =================================================================
 * 外部ライブラリに関する注記
 * =================================================================
 * このスクリプトは、'AppSheetClient' および 'GeminiClient' クラスが
 * プロジェクト内の他のスクリプトファイルで定義されていることを
 * 前提としています。
 *
 * このファイルから重複するクラス定義を削除したため、
 * エラーが解消されるはずです。
 * =================================================================
 */
