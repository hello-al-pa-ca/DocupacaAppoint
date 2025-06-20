/**
 * =================================================================
 * AI Sales Action (RAG機能除外・レスポンス整形機能強化版 v11)
 * =================================================================
 * 既存の AISalesAction.gs からRAG (Retrieval-Augmented Generation)
 * に関連する機能をすべて削除し、リファクタリングしたバージョンです。
 *
 * 主な変更点:
 * - SalesCopilotクラスを、AIによる文章生成と次アクション提案のコア機能に特化。
 * - placeholdersの役割を変更し、addPromptの内容を主要なプレースホルダーに割り当て。
 * - Google検索のロジックを「企業調査→本文生成」の2段階に変更し、調査結果をプロンプトに反映させるように強化。
 * - 無限ループの原因となっていた、処理開始時のステータス更新処理を削除。
 * - 【v11での修正】AIの応答に不要な解説が含まれないよう、プロンプトに制約を追加。
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
function executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName = '', companyAddress = '', customerContactName = '', ourContactName = '', probability = '', attachmentFileName = '', execUserEmail) {
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability);
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
ものづくり産業交流展示会では、お忙しい中お名刺交換させていただき、誠にありがとうございました。

[商談メモの内容を加味した、1言メッセージ]

まだまだ小さな会社ではありますが、経営の効率化や、将来を見据えた体制づくりについて、何かお役に立てることがあるかもしれません。

まずは御礼まで。
貴重なご縁をありがとうございました。

今後ともどうぞよろしくお願いいたします。`;
    const addPrompt = 'ドキュパカに興味あり';
    const companyName = '株式会社テスト';
    const companyAddress = '東京都千代田区1-1-1';
    const customerContactName = '山田 太郎'; // 取引先担当者名
    const ourContactName = '鈴木 一郎'; // 自社担当者名
    const probability = 'A'; // 契約の確度
    const execUserEmail = 'hello@al-pa-ca.com'; // 実行ユーザーのメールアドレス

    Logger.log("以下のパラメータでテスト実行します:");
    Logger.log({recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, execUserEmail});

    executeAISalesAction(recordId, '', '', AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability, '', execUserEmail);
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
    if (!execUserEmail) throw new Error("実行ユーザーのメールアドレス(execUserEmail)は必須です。");

    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    // AppSheetClientは外部ライブラリとして定義されていると仮定
    this.appSheetClient = new AppSheetClient(this.props.APPSHEET_APP_ID, this.props.APPSHEET_API_KEY);

    const masterSheetId = this.props.MASTER_SHEET_ID;
    if (!masterSheetId) throw new Error("マスターシートのIDがスクリプトプロパティに設定されていません。");

    // マスターデータをスプレッドシートから読み込む
    this.actionCategories = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.actionCategories);
    this.aiRoles = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.aiRoles);
    this.salesFlows = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.salesFlows);
  }

  /**
   * AIによる営業アクションの文章を生成し、AppSheetを更新します。
   */
  executeAISalesAction(recordId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, customerContactName, ourContactName, probability) {
    try {
      // 無限ループを防ぐため、処理開始時のステータス更新は削除

      const actionDetails = this._getActionDetails(actionName, contactMethod);
      if (!actionDetails) throw new Error(`アクション定義が見つかりません: ${actionName}/${contactMethod}`);

      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      
      const placeholders = {
        '[具体的な課題]': addPrompt, '[資料名]': addPrompt,
        '[以前話した課題]': addPrompt, '[推測される課題]': addPrompt,
        '[提案書名]': addPrompt, '[提案内容]': addPrompt,
        '[議題]': addPrompt, '[期間]': addPrompt,
        '[顧客の会社名]': companyName, '[企業名]': companyName,
        '[会社の住所]': companyAddress,
        '[取引先担当者名]': customerContactName,
        '[自社担当者名]': ourContactName,
        '[契約の確度]': probability
      };
      
      const useGoogleSearch = actionDetails.searchGoogle && companyName;
      let companyInfo = '';
      if (useGoogleSearch) {
        Logger.log(`Google検索を有効にして企業情報を調査します: ${companyName}`);
        companyInfo = this._getCompanyInfo(companyName);
      }

      const template = mainPrompt || actionDetails.prompt;
      const finalPrompt = this._buildFinalPrompt(template, placeholders, contactMethod, companyInfo);
      Logger.log(`最終プロンプト: \n${finalPrompt}`);

      const geminiClient = new GeminiClient('gemini-1.5-flash-latest');
      geminiClient.setSystemInstructionText(aiRoleDescription);
      
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
   * @param {string} completedActionId - 完了したアクションのレコードID
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
   * Google検索を利用して企業情報を取得します。
   * @private
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
        return ''; // エラー時は空文字を返す
    }
  }


  /**
   * AIからのレスポンスを{suggest_ai_text, subject, body}の形式に整形します。
   * @private
   */
  _formatResponse(rawText, useGoogleSearch, contactMethod) {
    // Google検索を利用したかどうかに関わらず、レスポンスの形式が統一されているため、
    // 整形ロジックは一つにまとめることができる。
    // ただし、検索時と非検索時でAIの応答の仕方が変わる可能性を考慮し、ロジックは分離しておく。
    if (useGoogleSearch) {
      Logger.log("Google検索結果を含むテキストを整形します...");
      // 二次的な整形は、検索結果がうまく反映されなかった場合にのみ有効。
      // 今回は`_buildFinalPrompt`で情報を渡しているため、通常の分割ロジックで対応可能。
      return this._splitSubjectAndBody(rawText, contactMethod);
    } else {
      return this._splitSubjectAndBody(rawText, contactMethod);
    }
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

  /**
   * テキストを件名と本文に分割します。
   * @private
   */
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

  _buildFinalPrompt(template, placeholders, contactMethod, companyInfo = '') {
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
    if (companyInfo) {
        additionalInfo += `\n--- 企業調査情報 ---\n${companyInfo}\n`;
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
