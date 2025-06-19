// =================================================================
// Test Functions for GAS Editor
// =================================================================

/**
 * 【エディタ実行用】提供された固定の引数でexecuteAISalesActionを直接実行するテスト関数
 */
function test_executeAISalesAction_with_hardcoded_args() {
    const recordId = 'D0DF2170-EC27-4906-A53F-131581C1FDF3';
    const organizationId = 'b7f7113f-771e-4d3d-bf76-2ade7d8f4cbe';
    const accountId = '9250CC98-C95A-43D9-B261-E7EFD163B5E3-b3ac984e';
    const AIRoleName = 'AI 営業マン';
    const actionName = 'あいさつ';
    const contactMethod = 'メール';
    const mainPrompt = '初めて連絡する顧客へ、丁寧な自己紹介と簡潔な挨拶のメール文面を作成してください。件名も提案してください。';
    const addPrompt = '';
    const companyName = '池田泉州銀行';
    const companyAddress = '吹田市豊津町9番1号 EDGE江坂19F';
    const attachmentFileName = '';
    const execUserEmail = 'hello@al-pa-ca.com';

    Logger.log("以下のハードコードされたパラメータで実行します:");
    Logger.log({recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, attachmentFileName, execUserEmail});

    executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, attachmentFileName, execUserEmail);
}

/**
 * =================================================================
 * AI Sales Cycle Co-pilot for AppSheet (v20 - Final JOIN Syntax Fix)
 * =================================================================
 * メインの実行ロジック(SalesCopilot)と、RAG機能(RAGClient, VectorDBManager)を
 * クラスとして分離し、コードの構造を全面的に再構成。
 * BigQueryのVECTOR_SEARCHでフィルタリングを行うための正しいSQL構文に修正。
 */

// =================================================================
// 定数宣言
// =================================================================
const MASTER_SHEET_NAMES = {
  actionCategories: 'ActionCategory',
  aiRoles: 'AIRole',
  salesFlows: 'ActionFlow'
};

const BIGQUERY_PROJECT_ID = '840992148496';
const BIGQUERY_DATASET_ID = 'rag_knowledge_base';
const BIGQUERY_TABLE_ID = 'knowledge_base';
const BIGQUERY_LOCATION = 'US'; // データセットのロケーションに合わせて変更してください

// =================================================================
// グローバル関数 (AppSheetまたは手動で実行)
// =================================================================

/**
 * 【AppSheetから実行】AIによる文章生成を指示する
 */
function executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName = '', companyAddress = '', attachmentFileName = '', execUserEmail) {
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, attachmentFileName);
  } catch (e) {
    Logger.log(`❌ ラッパー関数で致命的なエラー: ${e.message}`);
  }
}

/**
 * 【AppSheetから実行】アクションの結果に基づき、次のアクションを提案する
 */
function suggestNextAction(completedActionId, execUserEmail) {
  try {
    const copilot = new SalesCopilot(execUserEmail);
    copilot.suggestNextAction(completedActionId);
  } catch (e) {
    Logger.log(`❌ ラッパー関数で致命的なエラー: ${e.message}`);
  }
}

/**
 * 【手動・定時実行】管理用スプレッドシートを元に、全アカウントのナレッジベースを更新する
 */
function runDailyIndexUpdate() {
  try {
    const accountSheetId = PropertiesService.getScriptProperties().getProperty('ACCOUNT_SHEET_ID');
    if (!accountSheetId) throw new Error("アカウント管理シートのIDがスクリプトプロパティに設定されていません。");

    const sheet = SpreadsheetApp.openById(accountSheetId).getSheets()[0];
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const orgIdIndex = headers.indexOf('organization_id'), accountIdIndex = headers.indexOf('account_id'), folderIdIndex = headers.indexOf('rag_gdrive_folder_id');

    if (orgIdIndex === -1 || accountIdIndex === -1 || folderIdIndex === -1) {
      throw new Error("アカウント管理シートのヘッダーに 'organization_id', 'account_id', 'rag_gdrive_folder_id' のいずれかが見つかりません。");
    }

    const dbManager = new VectorDBManager(BIGQUERY_PROJECT_ID, BIGQUERY_DATASET_ID, BIGQUERY_TABLE_ID);

    rows.forEach(row => {
      const organizationId = row[orgIdIndex], accountId = row[accountIdIndex], folderId = row[folderIdIndex];
      if (organizationId && accountId && folderId) {
        Logger.log(`--- Indexing started for Account: ${accountId} ---`);
        dbManager.recreateIndexFromDriveFolder(organizationId, accountId, folderId);
        Logger.log(`--- Indexing finished for Account: ${accountId} ---`);
      }
    });
  } catch (e) {
    Logger.log(`❌ 一括インデックス更新中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

// =================================================================
// Test Functions for GAS Editor
// =================================================================

/**
 * 【エディタ実行用】提供された固定の引数でexecuteAISalesActionを直接実行するテスト関数
 */
function test_executeAISalesAction_with_hardcoded_args() {
    const recordId = 'D0DF2170-EC27-4906-A53F-131581C1FDF3';
    const organizationId = 'b7f7113f-771e-4d3d-bf76-2ade7d8f4cbe';
    const accountId = '9250CC98-C95A-43D9-B261-E7EFD163B5E3-b3ac984e';
    const AIRoleName = 'AI 営業マン';
    const actionName = 'あいさつ';
    const contactMethod = 'メール';
    const mainPrompt = '初めて連絡する顧客へ、丁寧な自己紹介と簡潔な挨拶のメール文面を作成してください。件名も提案してください。';
    const addPrompt = '';
    const companyName = '池田泉州銀行';
    const companyAddress = '吹田市豊津町9番1号 EDGE江坂19F';
    const attachmentFileName = '';
    const execUserEmail = 'hello@al-pa-ca.com';

    Logger.log("以下のハードコードされたパラメータで実行します:");
    Logger.log({recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, attachmentFileName, execUserEmail});

    executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName, companyAddress, attachmentFileName, execUserEmail);
}


// =================================================================
// SalesCopilot クラス (メインのアプリケーションロジック)
// =================================================================
class SalesCopilot {
  constructor(execUserEmail) {
    if (!execUserEmail) throw new Error("実行ユーザーのメールアドレス(execUserEmail)は必須です。");
    
    this.props = PropertiesService.getScriptProperties().getProperties();
    this.execUserEmail = execUserEmail;
    this.appSheetClient = this._getAppSheetClient();
    
    const masterSheetId = this.props.MASTER_SHEET_ID;
    if (!masterSheetId) throw new Error("マスターシートIDが設定されていません。");
    
    // ★★★ 変更点1: アカウント管理シートの情報を読み込む ★★★
    this.accountData = this._loadAccountData(); 
    this.actionCategories = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.actionCategories);
    this.aiRoles = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.aiRoles);
    this.salesFlows = this._loadSheetData(masterSheetId, MASTER_SHEET_NAMES.salesFlows);

    // RAGClientの初期化は変更なし
    this.ragClient = new RAGClient(this.execUserEmail); 
  }

  executeAISalesAction(recordId, organizationId, accountId, AIRoleName, actionName, contactMethod, mainPrompt, addPrompt, companyName = '', companyAddress = '', attachmentFileName = '') {
    try {
      this._updateAppSheetRecord(recordId, { "execute_ai_status": "AI処理中" });

      const actionDetails = this._getActionDetails(actionName, contactMethod);
      if (!actionDetails) throw new Error(`アクション定義が見つかりません: ${actionName}/${contactMethod}`);
      
      const aiRoleDescription = this._getAIRoleDescription(AIRoleName);
      if (!aiRoleDescription) throw new Error(`AI役割定義が見つかりません: ${AIRoleName}`);
      
      const userQueryForRAG = `${companyName || ''} ${companyAddress || ''} ${mainPrompt} ${addPrompt}`;
      const ragResult = this.ragClient.getContext(userQueryForRAG, organizationId, accountId);

      let fileBlob = null, identifiedFileName = '', fileIdToUse = null;
      
      // ★★★ 変更点2: アカウントに対応するフォルダIDを取得してファイル検索を行う ★★★
      if (attachmentFileName) {
        const ragFolderId = this._getRagFolderId(accountId);
        if (ragFolderId) {
          fileIdToUse = this._getFileIdByName(attachmentFileName, ragFolderId);
        } else {
          Logger.log(`アカウントID ${accountId} に対応するRAGフォルダIDが見つかりませんでした。`);
        }
      }
      
      if (!fileIdToUse) fileIdToUse = this._determineBestAttachment(ragResult.potentialFiles);
      
      fileIdToUse = this._getFileIdByName(fileIdToUse, "1T1SG1hCXU3SiQ3xos_-WMayy3gqW6AHg");
      if (fileIdToUse) {
        // IDの形式を簡易的にチェック (例: '.' が含まれていない、一定の長さ以上)
        if (typeof fileIdToUse !== 'string' || fileIdToUse.includes('.') || fileIdToUse.length < 20) {
            Logger.log(`無効なファイルID形式のためスキップ: ${fileIdToUse}`);
        } else {
            try {

                fileBlob = DriveApp.getFileById(fileIdToUse).getBlob();
                identifiedFileName = fileBlob.getName();
            } catch(fileError) {
                Logger.log(`ファイル(ID: ${fileIdToUse})の読み込みに失敗: ${fileError.message}`);
            }
        }
      }

      const placeholders = {
        '[具体的な課題]': mainPrompt, '[資料名]': identifiedFileName || '関連資料',
        '[以前話した課題]': addPrompt, '[推測される課題]': mainPrompt,
        '[提案書名]': identifiedFileName || '先日お送りした提案書', '[提案内容]': mainPrompt,
        '[議題]': mainPrompt, '[期間]': addPrompt,
        '[顧客の会社名]': companyName, '[企業名]': companyName,
        '[会社の住所]': companyAddress 
      };
      const finalPrompt = this._buildFinalPrompt(actionDetails.prompt, placeholders, ragResult.context);
      
      console.log(finalPrompt);

      const geminiClient = new GeminiClient('gemini-2.0-flash');
      geminiClient.setSystemInstructionText(aiRoleDescription);
      if (actionDetails.searchGoogle && companyName) geminiClient.enableGoogleSearchTool();
      if (fileBlob) geminiClient.attachFiles(fileBlob);
      geminiClient.setPromptText(finalPrompt);

      const response = geminiClient.generateCandidates();
      const generatedText = (response.candidates[0].content.parts || []).map(p => p.text).join('');
      if (!generatedText) throw new Error('Geminiからの応答が空でした。');

      const updatePayload = this._splitSubjectAndBody(generatedText);
      updatePayload["execute_ai_status"] = "提案済み";
      
      console.log(updatePayload);

      this._updateAppSheetRecord(recordId, updatePayload);
      Logger.log(`処理完了 (AI提案生成): Record ID ${recordId}`);

    } catch (e) {
      Logger.log(`❌ AI提案生成エラー: ${e.message}\n${e.stack}`);
      this._updateAppSheetRecord(recordId, { "execute_ai_status": "エラー", "suggest_ai_text": `処理エラー: ${e.message}` });
    }
  }
  
  suggestNextAction(completedActionId) {
    try {
      const completedAction = this._findRecordById('SalesActions', completedActionId);
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
  _getAppSheetClient() { const { APPSHEET_APP_ID, APPSHEET_API_KEY } = this.props; if (!APPSHEET_APP_ID || !APPSHEET_API_KEY) throw new Error('AppSheet接続情報(ID/Key)が不足しています。'); return new AppSheetClient(APPSHEET_APP_ID, APPSHEET_API_KEY); }
  _loadSheetData(sheetId, sheetName) { try { const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName); const [headers, ...rows] = sheet.getDataRange().getValues(); return rows.map(row => headers.reduce((obj, header, i) => (obj[header] = row[i], obj), {})); } catch(e) { throw new Error(`マスターシート(ID: ${sheetId}, Name: ${sheetName})の読み込みに失敗しました。`); } }
  
  // ★★★ 追加メソッド: アカウント情報を読み込む ★★★
  _loadAccountData() {
    const accountSheetId = this.props.ACCOUNT_SHEET_ID;
    if (!accountSheetId) throw new Error("アカウント管理シートのIDがスクリプトプロパティに設定されていません。");
    const sheet = SpreadsheetApp.openById(accountSheetId).getSheets()[0];
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const accountIdIndex = headers.indexOf('id');
    const folderIdIndex = headers.indexOf('rag_gdrive_folder_id');
    if (accountIdIndex === -1 || folderIdIndex === -1) {
      throw new Error("アカウント管理シートのヘッダーに 'account_id', 'rag_gdrive_folder_id' のいずれかが見つかりません。");
    }
    // account_id をキー、folder_id を値とするマップを作成
    return rows.reduce((map, row) => {
      const accountId = row[accountIdIndex];
      const folderId = row[folderIdIndex];
      if (accountId && folderId) {
        map[accountId] = folderId;
      }
      return map;
    }, {});
  }
  
  // ★★★ 追加メソッド: accountIdからRAGフォルダIDを取得する ★★★
  _getRagFolderId(accountId) {
    return this.accountData[accountId] || null;
  }

  // ★★★ 修正メソッド: _getFileIdByName に folderId を引数として追加 ★★★
  _getFileIdByName(fileName, folderId) {
    if (!fileName || !folderId) return null;
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFilesByName(fileName);
      if (files.hasNext()) {
        return files.next().getId();
      }
      return null;
    } catch (e) {
      Logger.log(`ファイル検索エラー (FolderID: ${folderId}, FileName: ${fileName}): ${e.message}`);
      return null;
    }
  }
  
  _splitSubjectAndBody(text, contactMethod) {
  // 基本のレスポンス構造
  const response = {
    "suggest_ai_text": text,
    "subject": "",
    "body": ""
  };

  // 接触方法がメールの場合のみ、件名と本文を抽出
  if (contactMethod === 'メール') {
    const subjectMarker = '【件名】';
    const bodyMarker = '【本文】';
    
    const subjectIndex = text.indexOf(subjectMarker);
    const bodyIndex = text.indexOf(bodyMarker);
    
    if (subjectIndex !== -1 && bodyIndex !== -1) {
      // 件名と本文を抽出
      let subject = text.substring(subjectIndex + subjectMarker.length, bodyIndex).trim();
      let body = text.substring(bodyIndex + bodyMarker.length).trim();
      
      // 件名から改行を除去
      subject = subject.replace(/\n/g, ' ').trim();
      
      response.subject = subject;
      response.body = body;
    } else if (subjectIndex !== -1) {
      // 件名のみ見つかった場合
      let subject = text.substring(subjectIndex + subjectMarker.length).trim();
      subject = subject.split('\n')[0].trim(); // 最初の行のみを件名とする
      response.subject = subject;
      response.body = text;
    } else {
      // マーカーが見つからない場合、全文を本文として扱う
      response.body = text;
      
      // 件名を推測で抽出する試み（最初の行が件名の可能性）
      const lines = text.split('\n');
      if (lines.length > 0 && lines[0].length < 50 && !lines[0].includes('様')) {
        response.subject = lines[0].trim();
        response.body = lines.slice(1).join('\n').trim();
      }
    }
  }
  // メール以外の接触方法の場合、subjectとbodyは空文字列のまま

  Logger.log(`件名抽出結果: "${response.subject}"`);
  Logger.log(`本文抽出結果: "${response.body.substring(0, 100)}..."`);
  
  return response;
}

  _determineBestAttachment(files) { if (!files || files.length === 0) return null; const counts = files.reduce((acc, file) => { acc[file] = (acc[file] || 0) + 1; return acc; }, {}); return Object.keys(counts).reduce((a, b) => counts[a] > counts[b] ? a : b); }

  _getActionDetails(actionName, contactMethod) { return this.actionCategories.find(row => row.action_name === actionName && row.contact_method === contactMethod) || null; }
  _getAIRoleDescription(roleName) { const role = this.aiRoles.find(row => row.name === roleName); return role ? role.description : `あなたは優秀な「${roleName}」です。`; }
  _getActionFlowDetails(currentProgress, currentActionName, currentResult) { return this.salesFlows.find(row => row.progress === currentProgress && row.action_id === currentActionName && row.result === currentResult) || null; }
  _findNextActionInfo(nextActionName) { const defaultContactMethod = 'メール'; const nextAction = this.actionCategories.find(row => row.action_name === nextActionName && row.contact_method === defaultContactMethod) || this.actionCategories.find(row => row.action_name === nextActionName); if (nextAction && nextAction.id) { return { id: nextAction.id, description: `${nextAction.action_name} (${nextAction.contact_method}) を実施してください。` }; } return { id: null, description: `推奨アクション: ${nextActionName}` }; }

  _buildFinalPrompt(template, placeholders, ragContext) {
  let finalPrompt = template;
  
  // プレースホルダーの置換
  for (const key in placeholders) {
    if (placeholders[key]) {
      // より確実な置換のため、グローバル置換を使用
      const regex = new RegExp(`\\${key}`, 'g');
      finalPrompt = finalPrompt.replace(regex, placeholders[key]);
    }
  }
  
  // 空の住所プレースホルダーのクリーンアップ
  if (finalPrompt.includes('[会社の住所]')) {
    finalPrompt = finalPrompt.replace(/（所在地: \[会社の住所\]）/g, '');
    finalPrompt = finalPrompt.replace(/所在地が\[会社の住所\]とのことで、[^。]*。/g, '');
    finalPrompt = finalPrompt.replace(/\[会社の住所\]/g, '');
  }
  
  // RAGコンテキストの追加
  if (ragContext && ragContext !== "（参考情報なし）") {
    finalPrompt += `\n\n【参考情報】\n${ragContext}`;
  }
  
  // 最終的なプロンプトの調整指示を追加
  finalPrompt += `\n\n【重要】
- 必ず【件名】【本文】の形式で回答してください
- プレースホルダー（[○○]）がある場合は、実際の情報に置き換えてください
- 件名は簡潔で魅力的にしてください`;

  Logger.log(`最終プロンプト（最初の200文字）: ${finalPrompt.substring(0, 200)}...`);
  
  return finalPrompt;
}
  
  _updateAppSheetRecord(recordId, fieldsToUpdate) { const recordData = { "ID": recordId, ...fieldsToUpdate }; return this.appSheetClient.updateRecords('SalesActions', [recordData], this.execUserEmail); }
  _findRecordById(tableName, recordId){ const selector = `SELECT([ID], [ID] = "${recordId}")`; const properties = { "Selector": selector }; const result = this.appSheetClient.findData(tableName, this.execUserEmail, properties); if(result && result.length > 0) return result[0]; return null; }
}


// =================================================================
// RAGClient クラス (ナレッジ検索ロジック)
// =================================================================
class RAGClient {
  constructor(execUserEmail) {
    this.execUserEmail = execUserEmail;
    this.embeddingClient = new EmbeddingClient('text-embedding-004');
    this.bigQueryAuthToken = this._getBigQueryAuthToken();
  }

  _getBigQueryAuthToken() {
    const service = OAuth2.createService('BigQueryRAG')
      .setTokenUrl('https://accounts.google.com/o/oauth2/token')
      .setPrivateKey(SERVICE_ACCOUNT_KEY.private_key)
      .setIssuer(SERVICE_ACCOUNT_KEY.client_email)
      .setSubject(this.execUserEmail)
      .setPropertyStore(PropertiesService.getScriptProperties())
      .setScope('https://www.googleapis.com/auth/bigquery');
    if (service.hasAccess()) return service.getAccessToken();
    else {
      Logger.log(`OAuth2 Error: ${service.getLastError()}`);
      throw new Error("BigQueryへの認証に失敗しました。");
    }
  }

  getContext(userQuery, organizationId, accountId) {
    const defaultReturn = { context: "（参考情報なし）", potentialFiles: [] };
    if (!userQuery || !organizationId || !accountId) return defaultReturn;
    
    try {
      const queryVector = this.embeddingClient.generate(userQuery, 'RETRIEVAL_QUERY');
      
      const sql = `
        SELECT
          *
        FROM
          VECTOR_SEARCH(
            TABLE \`${BIGQUERY_PROJECT_ID}.${BIGQUERY_DATASET_ID}.${BIGQUERY_TABLE_ID}\`,
            'embedding',
            (SELECT @queryVector AS embedding),
            top_k => 20,
            distance_type => 'COSINE'
          )`;
      
      const requestBody = {
        query: sql,
        useLegacySql: false,
        queryParameters: [
          { name: 'queryVector', parameterType: { type: 'ARRAY', arrayType: { type: 'FLOAT64' } }, 
            parameterValue: { arrayValues: queryVector.map(v => ({ value: v.toString() })) } }
        ]
      };
      
      const url = `https://bigquery.googleapis.com/bigquery/v2/projects/${BIGQUERY_PROJECT_ID}/queries`;
      const options = {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': `Bearer ${this.bigQueryAuthToken}` },
        payload: JSON.stringify(requestBody), muteHttpExceptions: true
      };
      const response = UrlFetchApp.fetch(url, options);
      const result = JSON.parse(response.getContentText());

      if (result.error || response.getResponseCode() >= 400) {
        throw new Error(`BigQuery API Error: ${result.error ? JSON.stringify(result.error.errors) : response.getContentText()}`);
      }
      if (!result.rows || result.rows.length === 0) return defaultReturn;
      
      // ★★★ 最終修正箇所 ★★★
      // ログから判明したネスト構造(baseフィールド)に直接アクセスするように修正。
      // baseレコード内のフィールド順序は固定と仮定し、インデックスで指定する。
      // baseフィールドのスキーマ: [chunk_id, organization_id, account_id, source_document, chunk_text, embedding]
      const filteredRows = result.rows
        .map(row => {
          // row.f[1] が 'base' レコードに対応する
          const baseRecord = row.f[1].v.f; 
          return {
            organization_id: baseRecord[1].v,
            account_id:      baseRecord[2].v,
            source_document: baseRecord[3].v,
            chunk_text:      baseRecord[4].v
          };
        })
        .filter(record => record.organization_id === organizationId && record.account_id === accountId)
        .slice(0, 5); // 上位5件に制限

      if (filteredRows.length === 0) return defaultReturn;
      
      const context = `--- 参考：過去の類似データ ---\n${filteredRows.map(record => record.chunk_text).join('\n---\n')}`;
      const potentialFiles = filteredRows.map(record => record.source_document).filter(v => v);
      return { context, potentialFiles };
    } catch (e) {
      // エラーメッセージに詳細を追加してデバッグしやすくする
      const errorMessage = e.stack || e.message;
      Logger.log(`BigQuery RAG検索エラー: ${errorMessage}`); 
      throw new Error(`ナレッジベースの検索に失敗しました: ${e.message}`);
    }
  }
}




// =================================================================
// VectorDBManager クラス (ナレッジ構築ロジック)
// =================================================================
class VectorDBManager {
  constructor(projectId, datasetId, tableId) {
    this.projectId = projectId;
    this.datasetId = datasetId;
    this.tableId = tableId;
    this.embeddingClient = new EmbeddingClient('text-embedding-004');
    this.bigQueryAuthToken = this._getBigQueryAuthTokenForManager();
  }
  
  _getBigQueryAuthTokenForManager() {
    const service = OAuth2.createService('BigQueryManager')
        .setTokenUrl('https://accounts.google.com/o/oauth2/token')
        .setPrivateKey(SERVICE_ACCOUNT_KEY.private_key)
        .setIssuer(SERVICE_ACCOUNT_KEY.client_email)
        .setPropertyStore(PropertiesService.getScriptProperties())
        .setScope('https://www.googleapis.com/auth/bigquery');
    if (service.hasAccess()) return service.getAccessToken();
    else throw new Error("BigQuery Managerの認証に失敗しました。");
  }

  recreateIndexFromDriveFolder(organizationId, accountId, folderId) {
    try {
      this._deleteRowsByAccountId(accountId);
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      const rowsToInsert = [];
      while (files.hasNext()) {
        const file = files.next();
        try {
          const text = this._extractTextFromFile(file);
          if (!text || text.trim() === '') continue;
          const chunks = this._splitTextIntoChunks(text, 500);
          chunks.forEach(chunk => {
            const vector = this.embeddingClient.generate(chunk, 'RETRIEVAL_DOCUMENT');
            rowsToInsert.push({
              json: {
                chunk_id: Utilities.getUuid(), organization_id: organizationId, account_id: accountId,
                source_document: file.getId(), chunk_text: chunk, embedding: vector
              }
            });
          });
        } catch (fileError) {
          Logger.log(`❌ ファイル処理中にエラー: ${file.getName()}, Error: ${fileError.message}`);
          continue;
        }
      }
      if (rowsToInsert.length > 0) {
        const url = `https://bigquery.googleapis.com/bigquery/v2/projects/${this.projectId}/datasets/${this.datasetId}/tables/${this.tableId}/insertAll`;
        const request = { "rows": rowsToInsert };
        const options = {
          method: 'post', contentType: 'application/json',
          headers: { 'Authorization': `Bearer ${this.bigQueryAuthToken}` },
          payload: JSON.stringify(request), muteHttpExceptions: true
        };
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() >= 400) {
          throw new Error(`BigQuery InsertAll API Error: ${response.getContentText()}`);
        }
        Logger.log(`  ${rowsToInsert.length}件のチャンクをBigQueryに正常に保存しました。`);
      }
    } catch (e) {
      Logger.log(`❌ recreateIndexFromDriveFolderでエラーが発生しました: ${e.message}`);
      throw e;
    }
  }

  _extractTextFromFile(file) {
    const mimeType = file.getMimeType();
    if (mimeType === MimeType.PLAIN_TEXT || mimeType === MimeType.GOOGLE_DOCS) {
      return file.getBlob().getDataAsString('UTF-8');
    } else if (mimeType === 'application/pdf') {
      const tempDoc = Drive.Files.insert({ title: `temp_ocr_${file.getId()}`, mimeType: MimeType.GOOGLE_DOCS }, file.getBlob(), { ocr: true });
      const text = DocumentApp.openById(tempDoc.id).getBody().getText();
      Drive.Files.remove(tempDoc.id);
      return text;
    }
    return '';
  }
  
  _deleteRowsByAccountId(accountId) {
    const sql = `DELETE FROM \`${this.projectId}.${this.datasetId}.${this.tableId}\` WHERE account_id = @accountId`;
    const requestBody = { query: sql, useLegacySql: false, queryParameters: [ { name: 'accountId', parameterType: { type: 'STRING' }, parameterValue: { value: accountId } } ] };
    const url = `https://bigquery.googleapis.com/bigquery/v2/projects/${this.projectId}/queries`;
    const options = {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': `Bearer ${this.bigQueryAuthToken}` },
      payload: JSON.stringify(requestBody), muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if(response.getResponseCode() >= 400) {
        throw new Error(`BigQuery Delete API Error: ${response.getContentText()}`);
    }
  }

  _splitTextIntoChunks(text, chunkSize) {
    const chunks = []; let i = 0;
    while (i < text.length) { chunks.push(text.substring(i, i + chunkSize)); i += chunkSize; }
    return chunks;
  }
}
