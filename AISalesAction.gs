/**
 * =================================================================
 * AI Sales Action (完成版 v28)
 * =================================================================
 * これまでの機能に加え、URL処理と件名抽出ロジックを大幅に改善しました。
 *
 * 【v28での主な変更点】
 * - AIへの指示をより明確化し、参考資料のリンクが複数ある場合に、
 * そのすべてを生成する本文に含めるようにプロンプトを強化しました。
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
 * @param {string} recordId - AIによる提案を記録するSalesActionレコードのID。
 * @param {string} organizationId - 組織ID（現在は未使用ですが、将来の拡張用）。
 * @param {string} accountId - アカウントID（現在は未使用ですが、将来の拡張用）。
 * @param {string} AIRoleName - 使用するAIの役割名（例: 'AI 営業マン'）。
 * @param {string} actionName - 実行するアクションの名称（例: 'あいさつ'）。
 * @param {string} contactMethod - 接触方法（例: 'メール', '電話'）。
 * @param {string} mainPrompt - AIに渡すメインの指示やテンプレート。
 * @param {string} addPrompt - ユーザーが追記する自由記述の指示やメモ。
 * @param {string} [companyName=''] - 顧客の会社名。
 * @param {string} [companyAddress=''] - 顧客の住所。
 * @param {string} [customerContactName=''] - 顧客の担当者名。
 * @param {string} [ourContactName=''] - 自社の担当者名。
 * @param {string} [probability=''] - 契約の確度。
 * @param {string} [eventName=''] - 名刺交換などをしたイベント名。
 * @param {string} [ourCompanyInfoText=''] - テキスト形式の自社情報。
 * @param {string} [ourCompanyInfoFileId=''] - ファイル形式の自社情報のGoogle DriveファイルID。
 * @param {string} [referenceUrls=''] - 参考にするGoogle DriveのファイルURL（カンマ区切りで複数可）。
 * @param {string} execUserEmail - このスクリプトを実行するユーザーのメールアドレス。
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
 * 【AppSheetから実行】完了したアクションに基づき、次のアクションを提案します。
 * @param {string} completedActionId - 結果が記録されたSalesActionレコードのID。
 * @param {string} execUserEmail - このスクリプトを実行するユーザーのメールアドレス。
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
 * 【エディタ実行用】固定の引数を使ってexecuteAISalesActionをテストします。
 * 開発時にこの関数を実行して、メイン機能の動作確認を行います。
 */
function test_executeAISalesAction() {
    const recordId = '7FBCF696-7397-49A3-BC8C-7E5E3AB3AAB4'; // テスト用のレコードID
    const AIRoleName = 'AI 営業マン';
    const actionName = '事例紹介';
    const contactMethod = 'メール';
    const mainPrompt = `[顧客の会社名]の[取引先担当者名]様

いつもお世話になっております。
株式会社〇〇の[自社担当者名]です。

[商談メモの内容を加味した、1言メッセージ]

つきましては、貴社と同様の課題をお持ちだった企業の成功事例をご紹介する資料をお送りいたします。
添付の資料が、貴社の課題解決の一助となれば幸いです。

ご不明な点がございましたら、お気軽にお申し付けください。
`;
    const addPrompt = `先日お話しした件について、参考動画と資料をお送りします。https://drive.google.com/file/d/1kl8_Ly-lFB8pmIxrb6Yhi7oaJPns0gK1/view?usp=sharing`;
    const companyName = '株式会社テスト';
    const companyAddress = '東京都千代田区1-1-1';
    const customerContactName = '山田 太郎';
    const ourContactName = '鈴木 一郎';
    const probability = 'A';
    const eventName = '';
    const ourCompanyInfoText = '';
    const ourCompanyInfoFileId = '';
    const referenceUrls = 'https://docs.google.com/document/d/1mCjPNOHvhKLohepguS3bt9E3NEKhNCNVPr7B9MDyPdQ/edit'; // 追加の参考資料
    const execUserEmail = 'hello@al-pa-ca.com';

    Logger.log("以下のパラメータでテスト実行します:");
    Logger.log({recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls, execUserEmail});

    executeAISalesAction(recordId, '', '', AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls, execUserEmail);
}


// =================================================================
// SalesCopilot クラス (メインのアプリケーションロジック)
// =================================================================

/**
 * AI営業支援機能のメインロジックを管理するクラス。
 * @class
 */
class SalesCopilot {
  /**
   * SalesCopilotのインスタンスを初期化します。
   * 必要なクライアントの準備や、マスターデータの読み込みを行います。
   * @param {string} execUserEmail - スクリプトを実行するユーザーのメールアドレス。
   */
  constructor(execUserEmail) {
    if (!execUserEmail) {
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
   * AIによる営業アクションの文章を生成し、結果をAppSheetのレコードに書き込みます。
   * @param {string} recordId - AIによる提案を記録するSalesActionレコードのID。
   * @param {string} AIRoleName - 使用するAIの役割名。
   * @param {string} actionName - 実行するアクションの名称。
   * @param {string} contactMethod - 接触方法。
   * @param {string} mainPrompt - AIに渡すメインの指示やテンプレート。
   * @param {string} addPrompt - ユーザーが追記する自由記述の指示やメモ。
   * @param {string} companyName - 顧客の会社名。
   * @param {string} companyAddress - 顧客の住所。
   * @param {string} customerContactName - 顧客の担当者名。
   * @param {string} ourContactName - 自社の担当者名。
   * @param {string} probability - 契約の確度。
   * @param {string} eventName - 名刺交換などをしたイベント名。
   * @param {string} ourCompanyInfoText - テキスト形式の自社情報。
   * @param {string} ourCompanyInfoFileId - ファイル形式の自社情報のGoogle DriveファイルID。
   * @param {string} referenceUrls - 参考にするGoogle DriveのファイルURL（カンマ区切りで複数可）。
   */
  executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, eventName, ourCompanyInfoText, ourCompanyInfoFileId, referenceUrls) {
    try {
      const actionDetails = this._getActionDetails(actionName, contactMethod);
      if (!actionDetails) throw new Error(`アクション定義が見つかりません: ${actionName}/${contactMethod}`);

      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      
      const currentAction = this._findRecordById('SalesAction', recordId);
      const customerId = currentAction ? currentAction.取引先ID : null;

      let historySummary = '';
      if (customerId) {
        historySummary = this._summarizePastActions(customerId, recordId);
      }

      const { processedAddPrompt, referenceContent, markdownLinkList } = this._processUrlInputs(addPrompt, referenceUrls);

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
      
      const template = mainPrompt || actionDetails.prompt;
      const finalPrompt = this._buildFinalPrompt(template, placeholders, contactMethod, companyInfo, referenceContent, historySummary, markdownLinkList);
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
      
      // AppSheetに更新するペイロードを作成
      const updatePayload = {
          "suggest_ai_text": formattedData.suggest_ai_text,
          "subject": formattedData.subject,
          "body": formattedData.body,
          "execute_ai_status": "提案済み",
          "link_markdown": markdownLinkList
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
   * 完了したアクションの結果に基づき、次のアクションを提案します。
   * @param {string} completedActionId - 結果が記録されたSalesActionレコードのID。
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

  /**
   * 指定された顧客との過去のアクション履歴を取得し、AIに要約させます。
   * @private
   * @param {string} customerId - 顧客のID。
   * @param {string} currentActionId - 現在処理中のアクションのID（履歴から除外するため）。
   * @returns {string} - AIによって生成された履歴の要約テキスト。
   */
  _summarizePastActions(customerId, currentActionId) {
    try {
      Logger.log(`顧客ID [${customerId}] の過去の商談履歴の要約を開始します。`);
      const selector = `FILTER("SalesAction", AND([取引先ID] = "${customerId}", [ID] <> "${currentActionId}"))`;
      const pastActions = this.appSheetClient.findData('SalesAction', this.execUserEmail, { "Selector": selector });

      if (!pastActions || pastActions.length === 0) {
        Logger.log("要約対象の過去のアクションはありませんでした。");
        return "";
      }

      const historyText = pastActions
        .sort((a, b) => new Date(a.実施日時) - new Date(b.実施日時)) // 時系列にソート
        .map(action => `日時: ${action.実施日時}\nアクション: ${action.action_name}\nメモ: ${action.addPrompt || ''}\n結果: ${action.result || ''}\nAI提案: ${action.body || ''}`)
        .join('\n\n---\n\n');

      const summarizationPrompt = `以下の商談履歴の要点を、重要なポイントを3行程度でまとめてください。\n\n--- 履歴 ---\n${historyText}`;
      
      const summarizerClient = new GeminiClient('gemini-1.5-flash-latest');
      summarizerClient.setPromptText(summarizationPrompt);
      const response = summarizerClient.generateCandidates();
      const summary = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      
      Logger.log(`商談履歴の要約:\n${summary}`);
      return summary;

    } catch (e) {
      Logger.log(`商談履歴の要約中にエラーが発生しました: ${e.message}`);
      return "";
    }
  }

  /**
   * `addPrompt`と`referenceUrls`からURLを処理し、整形済みテキスト、ファイル内容、Markdownリンクリストを生成します。
   * @private
   * @param {string} addPrompt - ユーザーが追記する自由記述の指示やメモ。
   * @param {string} referenceUrls - 参考にするGoogle DriveのファイルURL（カンマ区切りで複数可）。
   * @returns {{processedAddPrompt: string, referenceContent: string, markdownLinkList: string}} - 処理結果のオブジェクト。
   */
  _processUrlInputs(addPrompt, referenceUrls) {
    const combinedUrlsString = [addPrompt, referenceUrls].filter(Boolean).join(',');
    const urlRegex = /https?:\/\/(?:drive|docs)\.google\.com\/(?:file|document|spreadsheets|presentation)\/d\/([a-zA-Z0-9_-]{28,})/g;
    const uniqueUrls = [...new Set(combinedUrlsString.match(urlRegex) || [])];

    if (uniqueUrls.length === 0) {
      return {
        processedAddPrompt: addPrompt,
        referenceContent: '',
        markdownLinkList: ''
      };
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

        // 1. Markdownリンクリストに追加
        markdownLinkArray.push(markdownLink);
        
        // 2. ファイル内容を抽出してAIへの参考情報に追加
        const textContent = this._extractTextFromFile(file);
        if (textContent) {
          referenceContentText += `--- 参考資料: ${fileName} ---\n${textContent}\n\n`;
        }

        // 3. 元のaddPrompt内のURLをMarkdownリンクに置換
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
  
  /**
   * Google DriveのURLからファイルIDを抽出します。
   * @private
   * @param {string} url - Google DriveのURL。
   * @returns {string|null} - 抽出されたファイルID。
   */
  _extractFileIdFromUrl(url) {
      if (!url) return null;
      const match = url.match(/\/d\/([a-zA-Z0-9_-]{28,})/);
      return match ? match[1] : null;
  }

  /**
   * Google Driveのファイルオブジェクトからテキスト内容を抽出します。動画ファイルにも対応。
   * @private
   * @param {GoogleAppsScript.Drive.File} file - テキストを抽出するファイルオブジェクト。
   * @returns {string} - 抽出されたテキスト、または動画ファイルの場合は説明文。
   */
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

  /**
   * Google検索を使って企業情報を調査します。
   * @private
   * @param {string} companyName - 調査対象の企業名。
   * @returns {string} - AIによって要約された企業情報。
   */
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

  /**
   * AIからの応答テキストを、件名と本文を含むオブジェクトに整形します。
   * @private
   * @param {string} rawText - AIからの生の応答テキスト。
   * @param {boolean} useGoogleSearch - Google検索が使用されたかどうか。
   * @param {string} contactMethod - 接触方法。
   * @returns {{suggest_ai_text: string, subject: string, body: string}} - 整形されたオブジェクト。
   */
  _formatResponse(rawText, useGoogleSearch, contactMethod) {
    return this._splitSubjectAndBody(rawText, contactMethod);
  }

  /**
   * AppSheet APIクライアントのインスタンスを取得します。
   * @private
   * @returns {AppSheetClient} - AppSheetClientのインスタンス。
   */
  _getAppSheetClient() {
    const { APPSHEET_APP_ID, APPSHEET_API_KEY } = this.props;
    if (!APPSHEET_APP_ID || !APPSHEET_API_KEY) throw new Error('AppSheet接続情報(ID/Key)がスクリプトプロパティに設定されていません。');
    return new AppSheetClient(APPSHEET_APP_ID, APPSHEET_API_KEY);
  }

  /**
   * マスターデータをスプレッドシートから読み込み、オブジェクトの配列に変換します。
   * @private
   * @param {string} sheetId - スプレッドシートのID。
   * @param {string} sheetName - シート名。
   * @returns {Object[]} - 読み込んだデータの配列。
   */
  _loadSheetData(sheetId, sheetName) {
    try {
      const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
      const [headers, ...rows] = sheet.getDataRange().getValues();
      return rows.map(row => headers.reduce((obj, header, i) => (obj[header] = row[i], obj), {}));
    } catch (e) {
      throw new Error(`マスターシート(ID: ${sheetId}, Name: ${sheetName})の読み込みに失敗しました。: ${e.message}`);
    }
  }

  /**
   * テキストを【件名】と【本文】マーカーを元に分割します。
   * @private
   * @param {string} text - 分割対象のテキスト。
   * @param {string} contactMethod - 接触方法。
   * @returns {{suggest_ai_text: string, subject: string, body: string}} - 分割されたオブジェクト。
   */
  _splitSubjectAndBody(text, contactMethod) {
    const response = { "suggest_ai_text": text, "subject": "", "body": text };

    if (contactMethod === 'メール') {
      const subjectMarker = '【件名】';
      const bodyMarker = '【本文】';
      
      const subjectIndex = text.indexOf(subjectMarker);
      
      if (subjectIndex !== -1) {
        // 【件名】マーカーがある場合
        const bodyIndex = text.indexOf(bodyMarker, subjectIndex);
        let subjectText = '';
        let bodyText = '';

        if (bodyIndex !== -1) {
          // 【本文】マーカーもある
          subjectText = text.substring(subjectIndex + subjectMarker.length, bodyIndex).trim();
          bodyText = text.substring(bodyIndex + bodyMarker.length).trim();
        } else {
          // 【件名】マーカーのみ
          const lines = text.substring(subjectIndex + subjectMarker.length).trim().split('\n');
          subjectText = lines.find(line => line.trim() !== '') || '';
          const subjectLineIndex = lines.findIndex(line => line === subjectText);
          bodyText = lines.slice(subjectLineIndex + 1).join('\n').trim();
        }
        
        response.subject = subjectText.replace(/[\r\n]/g, ' ').trim();
        response.body = bodyText;
      } else {
        // マーカーが全くない場合のフォールバック
        const lines = text.trim().split('\n');
        // 最初の行が50文字未満で、"様"や"こんにちは"を含まないなど、件名らしい場合
        if (lines.length > 1 && lines[0].length < 50 && !lines[0].includes('様') && !lines[0].includes('こんにちは')) {
            response.subject = lines[0].trim();
            response.body = lines.slice(1).join('\n').trim();
        }
      }
    }
    return response;
  }

  /**
   * ActionCategoryシートから、指定されたアクション名と接触方法に一致する定義を取得します。
   * @private
   * @param {string} actionName - アクション名。
   * @param {string} contactMethod - 接触方法。
   * @returns {Object|null} - 見つかったアクション定義オブジェクト。
   */
  _getActionDetails(actionName, contactMethod) {
    return this.actionCategories.find(row => row.action_name === actionName && row.contact_method === contactMethod) || null;
  }

  /**
   * AIRoleシートから、指定された役割名に一致する説明文（ペルソナ設定）を取得します。
   * @private
   * @param {string} roleName - AIの役割名。
   * @returns {string} - AIの役割説明文。
   */
  _getAIRoleDescription(roleName) {
    const role = this.aiRoles.find(row => row.name === roleName);
    return role ? role.description : `あなたは優秀な「${roleName}」です。`;
  }

  /**
   * ActionFlowシートから、現在の状況に一致する次のフロー定義を取得します。
   * @private
   * @param {string} currentProgress - 現在のステータス。
   * @param {string} currentActionName - 実行されたアクション名。
   * @param {string} currentResult - アクションの結果。
   * @returns {Object|null} - 見つかったフロー定義オブジェクト。
   */
  _getActionFlowDetails(currentProgress, currentActionName, currentResult) {
    return this.salesFlows.find(row => row.progress === currentProgress && row.action_id === currentActionName && row.result === currentResult) || null;
  }

  /**
   * 次のアクション名から、具体的なアクション情報を検索します。
   * @private
   * @param {string} nextActionName - 次に実行すべきアクションの名前。
   * @returns {{id: string|null, description: string}} - 次のアクションのIDと説明文。
   */
  _findNextActionInfo(nextActionName) {
    const defaultContactMethod = 'メール';
    const nextAction = this.actionCategories.find(row => row.action_name === nextActionName && row.contact_method === defaultContactMethod)
                      || this.actionCategories.find(row => row.action_name === nextActionName);
    
    if (nextAction && nextAction.id) {
      return { id: nextAction.id, description: `${nextAction.action_name} (${nextAction.contact_method}) を実施してください。` };
    }
    return { id: null, description: `推奨アクション: ${nextActionName}` };
  }

  /**
   * すべての情報を統合し、AIに渡す最終的なプロンプトを組み立てます。
   * @private
   * @param {string} template - プロンプトのテンプレート。
   * @param {Object} placeholders - テンプレートに埋め込むプレースホルダーの値。
   * @param {string} contactMethod - 接触方法。
   * @param {string} [companyInfo=''] - Google検索で得た企業情報。
   * @param {string} [referenceContent=''] - Google Driveファイルから読み込んだ参考情報。
   * @param {string} [historySummary=''] - AIが要約した過去の商談履歴。
   * @param {string} [markdownLinks=''] - 参考資料のMarkdownリンクリスト。
   * @returns {string} - 完成した最終プロンプト。
   */
  _buildFinalPrompt(template, placeholders, contactMethod, companyInfo = '', referenceContent = '', historySummary = '') {
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
    if (historySummary) {
        additionalInfo += `\n--- これまでの商談履歴の要約 ---\n${historySummary}\n`;
        hasInfo = true;
    }


    if(hasInfo){
        finalPrompt += additionalInfo;
    }

    // 空のプレースホルダーが残っている場合、それを含む文章ごと削除するなどのクリーンアップ
    finalPrompt = finalPrompt.replace(/\[[^\]]+\]/g, '');


    if (contactMethod === 'メール') {
        finalPrompt += `\n\n【重要】\n- 必ず【件名】【本文】の形式で、メールやスクリプトの文章だけを生成してください。\n- 【補足情報】にある「企業調査情報」や「商談履歴の要約」を参考に、本文の冒頭で相手が「おっ」と思うような、関心を持っていることが伝わる自然な一文を加えてください。（例：「貴社の〇〇のニュース、興味深く拝見しました」「前回の〇〇の件、その後いかがでしょうか」など）\n- 件名は簡潔で分かりやすくしてください。\n- 生成する文章以外の解説や、確度に応じた文章の調整案などは一切含めないでください。`;
    }

    return finalPrompt;
  }

  /**
   * 指定されたIDのAppSheetレコードを更新します。
   * @private
   * @param {string} recordId - 更新するレコードのID。
   * @param {Object} fieldsToUpdate - 更新するフィールドのキーと値。
   * @returns {Object} - AppSheet APIからの応答。
   */
  _updateAppSheetRecord(recordId, fieldsToUpdate) {
    const recordData = { "ID": recordId, ...fieldsToUpdate };
    return this.appSheetClient.updateRecords('SalesAction', [recordData], this.execUserEmail);
  }

  /**
   * 指定されたIDのレコードをAppSheetテーブルから検索します。
   * @private
   * @param {string} tableName - 検索対象のテーブル名。
   * @param {string} recordId - 検索するレコードのID。
   * @returns {Object|null} - 見つかったレコードオブジェクト。
   */
  _findRecordById(tableName, recordId) {
    const selector = `SELECT(${tableName}[ID], [ID] = "${recordId}")`;
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
