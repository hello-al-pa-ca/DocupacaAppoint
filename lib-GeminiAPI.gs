/**
 * ===================================================================================
 * Gemini & Generative AI Client Library (v1.2 - API仕様変更対応版)
 * ===================================================================================
 *
 * 【v1.2での主な変更点】
 * - Gemini APIの仕様変更に伴い、Google検索ツールの内部名を古い'google_search_retrieval'から
 * 新しい'google_search'に修正しました。これにより、企業調査時のエラーを解消します。
 *
 * 【重要】利用前の設定:
 * 1. GCPプロジェクトの関連付け: Apps Scriptの「プロジェクトの設定」でGCPプロジェクト番号を設定します。
 * 2. APIの有効化: 関連付けたGCPプロジェクトで "Generative Language API" を有効化します。
 * 3. スクリプトプロパティの設定:
 * - GOOGLE_API_KEY: あなたのGoogle APIキー
 * ===================================================================================
 */

/**
 * スクリプトプロパティまたはコード内からAPIキーを取得します。
 * @returns {string | null} Google APIキー
 * @private
 */
function getApiKey_() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_API_KEY');
    if (apiKey) {
      return apiKey;
    }
  } catch (e) {
    Logger.log('スクリプトプロパティからAPIキーを取得できませんでした。コード内のプレースホルダーを確認します。');
  }
  // 重要：本番環境では、APIキーをコードに直接記述しないでください。
  // スクリプトプロパティに設定することを強く推奨します。
  return 'YOUR_API_KEY_HERE'; 
}


// --- 実行用サンプル関数群 ---

/**
 * テキスト生成のサンプル実行関数
 */
function runGeminiClientExample() {
  try {
    const model = 'gemini-1.5-flash-latest';
    const client = new GeminiClient(model);
    client.setPromptText('Google Apps Scriptの便利な活用方法を3つ、箇条書きで分かりやすく説明してください。');
    const response = client.generateCandidates();
    const text = response.candidates[0].content.parts[0].text;
    Logger.log('✅ APIからの応答:\n' + text);
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 画像とテキストをプロンプトとして、画像の内容を説明させるテスト関数
 */
function runVisionDescriptionExample() {
  try {
    const imageFileId = '1lrKq99FOJ1-JycyvwGiVH0gTICHTzmVW'; 
    if (imageFileId === 'YOUR_IMAGE_FILE_ID_HERE') {
      throw new Error('Googleドライブ上の画像ファイルIDを指定してください。');
    }
    const imageBlob = DriveApp.getFileById(imageFileId).getBlob();
    const model = 'gemini-1.5-flash-latest'; 
    const client = new GeminiClient(model);
    client.attachFiles(imageBlob); 
    client.setPromptText("この画像に写っているものを、詳しく説明してください。");
    const response = client.generateCandidates();
    const text = response.candidates[0].content.parts[0].text;
    Logger.log('✅ 画像の説明:\n' + text);
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * Google検索とURLによるグラウンディングを使用した実行サンプル関数
 */
function runGroundingExample() {
  try {
    const model = 'gemini-1.5-pro-latest';
    const client = new GeminiClient(model);
    
    client.enableGoogleSearchTool();
    
    const promptText = `神戸で明日から3日間のイベントスケジュールを提案してください。`;
    client.setPromptText(promptText);

    const response = client.generateCandidates();
    const text = (response.candidates[0].content.parts || []).map(p => p.text).join('');
     if (text) {
       Logger.log('✅ グラウンディングに基づいた回答:\n' + text);
    } else {
       Logger.log('⚠️ モデルからテキスト応答を取得できませんでした。');
    }
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}


/**
 * ヘルパー関数: 2つのベクトル間のコサイン類似度を計算する
 */
function calculateCosineSimilarity(vecA, vecB) {
  if (vecA.length !== vecB.length) {
    throw new Error("ベクトルの次元が一致しません。");
  }
  let dotProduct = 0;
  let normA = 0;
  let normB = 0;
  for (let i = 0; i < vecA.length; i++) {
    dotProduct += vecA[i] * vecB[i];
    normA += vecA[i] * vecA[i];
    normB += vecB[i] * vecB[i];
  }
  if (normA === 0 || normB === 0) return 0;
  return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
}


// --- 以下、クラス定義 ---

/**
 * @class RequestAPI
 * @classdesc APIリクエストの実行と再試行ロジックを管理する基本クラス。
 */
class RequestAPI {
  /**
   * @constructor
   */
  constructor() {
    /** @private */
    this.headers = { 'Content-Type': 'application/json' };
    /** @private */
    this.retryCount = 1;
    /** @private */
    this.retryDelay = 1000;
  }

  /**
   * リクエストヘッダーを設定します。
   * @param {string} key - ヘッダーのキー。
   * @param {string} val - ヘッダーの値。
   */
  setHeaders(key, val) { this.headers[key] = val; }

  /**
   * リクエスト失敗時の再試行設定を行います。
   * @param {number} count - 再試行の回数。
   * @param {number} delay - 再試行までの待機時間（ミリ秒）。
   */
  setRetryConfig(count, delay) { this.retryCount = count; this.retryDelay = delay; }

  /**
   * APIリクエストを実際に実行します。OAuthトークンを取得し、ヘッダーに付与します。
   * @param {string} url - リクエスト先のURL。
   * @param {object} options - UrlFetchApp.fetch()に渡すオプション。
   * @returns {object} - APIからのJSONレスポンス。
   * @protected
   */
  executeRequest(url, options) {
    let lastError;
    for (let i = 0; i < this.retryCount; i++) {
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        if (responseCode >= 200 && responseCode < 300) {
          try {
            return responseText ? JSON.parse(responseText) : {};
          } catch (jsonError) {
             throw new Error(`Failed to parse JSON response: ${jsonError.message}. Response text: ${responseText}`);
          }
        }
        
        throw new Error(`Request failed with status ${responseCode}: ${responseText}`);
      } catch (error) {
        lastError = error;
        console.error("Error:", error.message, "\nStack:", error.stack);
        if (i < this.retryCount - 1) {
          Utilities.sleep(this.retryDelay);
        }
      }
    }
    throw lastError;
  }

  /**
   * POSTリクエストを実行するためのヘルパーメソッド。
   * @param {string} url - リクエスト先のURL。
   * @param {object} payload - 送信するJSONペイロード。
   * @returns {object} - APIからのJSONレスポンス。
   * @protected
   */
  requestPostAPI(url, payload) {
    const options = {
      method: 'POST',
      headers: this.headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    return this.executeRequest(url, options);
  }
}

/**
 * @class AIComposer
 * @classdesc ファイル形式のサポート確認やBase64変換など、AIモデルへの入力作成に関連する機能を持つクラス。
 * @extends RequestAPI
 */
class AIComposer extends RequestAPI {
  /**
   * @constructor
   */
  constructor() {
    super();
    /** @private */
    this.supportedFormats = {
      application: ['application/pdf', 'application/x-javascript', 'application/x-python'],
      text: ['text/javascript', 'text/x-python', 'text/plain', 'text/html', 'text/css', 'text/md', 'text/csv', 'text/xml', 'text/rtf'],
      image: ['image/png', 'image/jpeg', 'image/heic', 'image/heif', 'image/webp'],
      audio: ['audio/wav', 'audio/mp3', 'audio/aiff', 'audio/aac', 'audio/ogg', 'audio/flac']
    };
    /** @private */
    this.allSupportedFormats = Object.values(this.supportedFormats).flat();
  }

  /**
   * 指定されたBlobがサポートされているファイル形式か確認します。
   * @param {GoogleAppsScript.Base.Blob} blob - 確認するBlobオブジェクト。
   * @returns {boolean} - サポートされていればtrue。
   * @private
   */
  _isSupportedFileFormat(blob) {
    return blob && typeof blob.getContentType === 'function' && this.allSupportedFormats.includes(blob.getContentType());
  }
  
  /**
   * BlobオブジェクトをBase64文字列に変換します。
   * @param {GoogleAppsScript.Base.Blob} blob - 変換するBlobオブジェクト。
   * @param {object} [options] - オプション。
   * @param {boolean} [options.includeDataUri=false] - Data URI形式で返すかどうか。
   * @returns {string} - Base64エンコードされた文字列。
   */
  convertBlobToBase64(blob, options = {}) {
    if (!this._isSupportedFileFormat(blob)) {
      throw new Error(`Unsupported file format: ${blob.getContentType()}`);
    }
    const base64 = Utilities.base64Encode(blob.getBytes());
    if (options.includeDataUri) {
      const mimeType = blob.getContentType();
      return `data:${mimeType};base64,${base64}`;
    }
    return base64;
  }
}

/**
 * @class GeminiClient
 * @classdesc Geminiモデルとの対話を行うためのクライアントクラス。
 * @extends AIComposer
 */
class GeminiClient extends AIComposer {
  /**
   * @param {string} model - 使用するモデル名 (例: 'gemini-1.5-pro-latest')。
   * @constructor
   */
  constructor(model) {
    super();
    if (!model) {
      throw new Error('モデル名は必須です。');
    }
    this.model = model;
    this.baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}`;
    this.uploadUrl = 'https://generativelanguage.googleapis.com/upload/v1beta/files';
    this.promptContents = {
      "contents": [],
      "generationConfig": {
        "temperature": 0.5, "topK": 40, "topP": 0.95, "maxOutputTokens": 8192
      }
    };
  }

  /**
   * ファイルをインラインデータ（Base64）としてプロンプトに添付します。
   * @param {GoogleAppsScript.Base.Blob} blob - 添付するBlobオブジェクト。
   * @private
   */
  _attachInline(blob) {
    let userContent = this.promptContents.contents.find(c => c.role === 'user');
    if (!userContent) {
      userContent = { role: 'user', parts: [] };
      this.promptContents.contents.push(userContent);
    }
    userContent.parts.push({
      "inlineData": { "mimeType": blob.getContentType(), "data": this.convertBlobToBase64(blob) }
    });
  }

  /**
   * ファイルをプロンプトに添付します。
   */
  attachFiles(blobs) {
    if (!Array.isArray(blobs)) {
      blobs = [blobs]; 
    }
    for (const blob of blobs) {
      if (!blob || typeof blob.getBytes !== 'function') {
        Logger.log('無効なBlobオブジェクトがスキップされました。');
        continue;
      }
      this._attachInline(blob);
    }
    return this;
  }
  
  /**
   * 現在のプロンプト内容でモデルにリクエストを送信し、応答を生成します。
   */
  generateCandidates() {
    const url = `${this.baseUrl}:generateContent`;
    // =================================================================
    // ▼▼▼【v1.2 修正点】古いツール名を修正 ▼▼▼
    // =================================================================
    const useApiKey = this.promptContents.tools && 
                      this.promptContents.tools.some(tool => tool.hasOwnProperty('google_search'));

    let options;
    let fullUrl;
    
    if (useApiKey) {
      const apiKey = getApiKey_();
      fullUrl = `${url}?key=${apiKey}`;
      options = {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify(this.promptContents),
        muteHttpExceptions: true
      };
    } else {
      fullUrl = url;
      const accessToken = ScriptApp.getOAuthToken();
      if (!accessToken) {
        throw new Error('OAuthトークンが取得できませんでした。承認フローを確認してください。');
      }
      options = {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + accessToken
        },
        payload: JSON.stringify(this.promptContents),
        muteHttpExceptions: true
      };
    }

    return this.executeRequest(fullUrl, options);
  }

  /**
   * プロンプトにテキスト部分を追加します。
   * @param {string} promptText - 追加するテキスト。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setPromptText(promptText) {
    let userContent = this.promptContents.contents.find(c => c.role === 'user');
    if (userContent) {
      const textPart = userContent.parts.find(p => "text" in p);
      if(textPart) {
        textPart.text += `\n${promptText}`;
      } else {
        userContent.parts.push({ "text": promptText });
      }
    } else {
      this.promptContents.contents.push({ "role": "user", "parts": [{ "text": promptText }] });
    }
    return this;
  }
  
  /**
   * モデルに対するシステムレベルの指示を設定します。
   * @param {string} systemInstruction - 設定するシステム指示テキスト。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setSystemInstructionText(systemInstruction) {
    this.promptContents["systemInstruction"] = { "parts": [{ "text": systemInstruction }] };
    return this;
  }
  
  /**
   * ツールを有効/無効にするための内部ヘルパーメソッド。
   * @param {string} toolName - ツールの名前。
   * @param {boolean} [enable=true] - 有効にするか無効にするか。
   * @private
   */
  _enableTool(toolName, enable = true) {
    if (!this.promptContents.tools) {
      this.promptContents.tools = [];
    }
    const toolIndex = this.promptContents.tools.findIndex(tool => tool.hasOwnProperty(toolName));
    if (enable) {
      if (toolIndex === -1) {
        this.promptContents.tools.push({ [toolName]: {} });
      }
    } else {
      if (toolIndex > -1) {
        this.promptContents.tools.splice(toolIndex, 1);
      }
    }
  }

  /**
   * Google検索ツールを有効化します。モデルがWeb検索を行えるようになります。
   * @param {boolean} [enable=true] - 有効にするか無効にするか。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  enableGoogleSearchTool(enable = true) {
    this._enableTool('google_search', enable);
    return this;
  }

  // 他のセッター関数は変更がないため省略
}
