/**
 * ===================================================================================
 * Gemini & Generative AI Client Library (v1.2 - API仕様変更/認証ロジック対応版)
 * ===================================================================================
 *
 * 【v1.2での主な変更点】
 * - Google検索ツールの内部名を 'google_search_retrieval' から 'google_search' に修正。
 * - API認証ロジックを改善し、以下のように切り分けました:
 * - `GeminiClient`のコンストラクタにAPIキーを渡すことで、そのインスタンスでの**デフォルト認証**をAPIキーにする。
 * - `google_search`ツールが有効な場合: APIキー認証を**強制**（コンストラクタに渡されていない場合はエラー）。
 * - 上記以外（APIキーが渡されていない、かつgoogle_searchツールが有効でない）の場合: OAuthトークン認証を利用。
 * OAuthが取得できない場合はエラー。
 * - `EmbeddingClient`と`ImagenClient`は、引き続き`getApiKey_()`から取得したAPIキーを必須として利用。
 *
 * 【重要】利用前の設定:
 * 1. GCPプロジェクトの関連付け: Apps Scriptの「プロジェクトの設定」でGCPプロジェクト番号を設定します。
 * 2. APIの有効化: 関連付けたGCPプロジェクトで "Generative Language API" を有効化します。
 * 3. スクリプトプロパティの設定:
 * - GOOGLE_API_KEY: あなたのGoogle APIキー (google_searchツール、File API、Embedding/Imagen APIを利用する場合に必須)
 * - [ファイル] > [プロジェクトのプロパティ] > [スクリプトのプロパティ] から追加してください。
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
 * テキスト生成のサンプル実行関数 (OAuth認証とAPIキー認証のデモンストレーション)
 */
function runGeminiClientExample() {
  try {
    const model = 'gemini-1.5-flash-latest';

    // --- OAuth認証のデモンストレーション ---
    Logger.log('--- OAuth認証でテキスト生成 (GCPプロジェクト関連付けと承認が必要) ---');
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const clientOAuth = new GeminiClient(model); 
    clientOAuth.setPromptText('Google Apps Scriptの便利な活用方法を3つ、箇条書きで分かりやすく説明してください。');
    let responseOAuth = clientOAuth.generateCandidates();
    Logger.log('✅ OAuth認証でのAPI応答:\n' + (responseOAuth.candidates[0].content.parts || []).map(p => p.text).join(''));

    // --- APIキー認証のデモンストレーション ---
    Logger.log('\n--- APIキー認証でテキスト生成 (GOOGLE_API_KEYの設定が必要) ---');
    // コンストラクタにAPIキーを明示的に渡すことで、OAuthより優先されます。
    const clientApiKey = new GeminiClient(model, getApiKey_()); 
    clientApiKey.setPromptText('Gemini APIでできることを3つ、箇条書きで分かりやすく説明してください。');
    let responseApiKey = clientApiKey.generateCandidates();
    Logger.log('✅ APIキー認証でのAPI応答:\n' + (responseApiKey.candidates[0].content.parts || []).map(p => p.text).join(''));

  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * 画像とテキストをプロンプトとして、画像の内容を説明させるテスト関数 (OAuth認証で実行)
 */
function runVisionDescriptionExample() {
  try {
    const imageFileId = '1lrKq99FOJ1-JycyvwGiVH0gTICHTzmVW'; 
    if (imageFileId === 'YOUR_IMAGE_FILE_ID_HERE') {
      throw new Error('Googleドライブ上の画像ファイルIDを指定してください。');
    }
    const imageBlob = DriveApp.getFileById(imageFileId).getBlob();
    const model = 'gemini-1.5-pro-latest'; 
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const client = new GeminiClient(model); 
    client.attachFiles(imageBlob); 
    client.setPromptText("この画像に写っているものを、詳しく説明してください。");
    const response = client.generateCandidates();
    const text = (response.candidates[0].content.parts || []).map(p => p.text).join('');
    Logger.log('✅ 画像の説明:\n' + text);
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * File APIでアップロードしたファイルを再利用するサンプル関数 (File APIはAPIキー必須)
 */
function runReusableFileApiExample() {
  try {
    const fileId = '1D8QXlRMt-rTZmzuHF3zX4NdLqjEAiB9F';
    if (fileId === 'YOUR_REUSABLE_FILE_ID_HERE') {
      throw new Error('テスト用のファイルのIDを指定してください。');
    }
    const fileBlob = DriveApp.getFileById(fileId).getBlob();
    const model = 'gemini-1.5-pro-latest';
    // File APIはGeminiClientのuploadFileメソッド内でAPIキーを直接使用するため、
    // ここでGeminiClientにAPIキーを渡す必要はありません（OAuthで動作）。
    const client = new GeminiClient(model); 

    Logger.log("ファイルを一度アップロードします...");
    const uploadedFileInfo = client.uploadFile(fileBlob); 
    Logger.log(`アップロード完了。File URI: ${uploadedFileInfo.uri}`);

    client.attachUploadedFile(uploadedFileInfo);

    Logger.log("\n--- 1回目の質問 ---");
    client.setPromptText("このドキュメントの内容を3つの箇条書きで要約してください。");
    const response1 = client.generateCandidates();
    const summaryText = (response1.candidates[0].content.parts || []).map(p => p.text).join('');
     if (!summaryText) {
      throw new Error(`APIからの1回目の応答が不正です: ${JSON.stringify(response1)}`);
    }
    Logger.log('✅ 1回目の回答:\n' + summaryText);

    client.promptContents.contents.push(response1.candidates[0].content);

    Logger.log("\n--- 2回目の質問 ---");
    client.setPromptText("先ほどの要約について、2番目の項目をさらに詳しく説明してください。");
    const response2 = client.generateCandidates();
    const detailedText = (response2.candidates[0].content.parts || []).map(p => p.text).join('');
    if (!detailedText) {
      throw new Error(`APIからの2回目の応答が不正です: ${JSON.stringify(response2)}`);
    }
    Logger.log('✅ 2回目の回答:\n' + detailedText);

  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}


/**
 * System Instruction を使った実行サンプル関数 (OAuth認証で実行)
 */
function runGeminiWithSystemInstructionExample() {
  try {
    const model = 'gemini-1.5-flash-latest';
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const client = new GeminiClient(model); 
    client.setSystemInstructionText('あなたは優秀なコピーライターです。簡潔で、キャッチーな文章を作成してください。');
    client.setPromptText('新しいプログラミング学習サービスの名前を5つ提案してください。');
    const response = client.generateCandidates();
    const text = (response.candidates[0].content.parts || []).map(p => p.text).join('');
    Logger.log('✅ コピーライターからの提案:\n' + text);
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * エンベディング生成のサンプル実行関数 (Embedding APIはAPIキー必須)
 */
function runEmbeddingExample() {
  try {
    const model = 'text-embedding-004';
    // EmbeddingClientは内部でAPIキーを使用します。
    const embeddingClient = new EmbeddingClient(model); 
    const text1 = "犬を動物病院に連れて行った";
    const vector1 = embeddingClient.generate(text1, 'RETRIEVAL_DOCUMENT');
    Logger.log(`「${text1}」のベクトル (最初の5次元): ${vector1.slice(0, 5)}...`);
    Logger.log(`ベクトル次元数: ${vector1.length}`);
    
    const text2 = "猫を獣医に診てもらった";
    const vector2 = embeddingClient.generate(text2, 'RETRIEVAL_DOCUMENT');
    
    const similarity = calculateCosineSimilarity(vector1, vector2);
    Logger.log(`テキスト間のコサイン類似度: ${similarity}`);
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * Imagen 3 を使った高品質な画像生成のサンプル (Imagen APIはAPIキー必須)
 */
function runImagenExample() {
  try {
    // ImagenClientは内部でAPIキーを使用します。
    const client = new ImagenClient(); 
    const prompt = "A photorealistic image of a golden retriever puppy playing in a field of flowers, with a shallow depth of field.";
    const imageBase64Array = client.generate(prompt, { sampleCount: 2, aspectRatio: "16:9" });
    Logger.log(`${imageBase64Array.length}枚の画像が生成されました。`);
    imageBase64Array.forEach((base64, index) => {
      const decoded = Utilities.base64Decode(base64, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decoded, 'image/png', `imagen-image-${index + 1}.png`);
      DriveApp.createFile(blob);
      Logger.log(`'imagen-image-${index + 1}.png' をドライブに保存しました。`);
    });
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * GeminiとImagenを連携させ、画像とテキストを別々に生成するサンプル (GeminiClientはOAuth、ImagenClientはAPIキー)
 */
function runCombinedImageAndTextGenerationExample() {
  try {
    const imagePrompt = "空飛ぶブタの楽しいイラスト";
    const storyPrompt = `「${imagePrompt}」というテーマで、そのブタが冒険に出る短い物語を考えてください。`;

    Logger.log("🎨 Imagen を使って画像を生成します...");
    // ImagenClientは内部でAPIキーを使用します。
    const imagenClient = new ImagenClient(); 
    const imageBase64Array = imagenClient.generate(imagePrompt, { sampleCount: 1 });
    
    if (imageBase64Array.length > 0) {
      Logger.log(`✅ 1枚の画像が生成されました。ドライブに保存します...`);
      const imageBase64 = imageBase64Array[0];
      const decoded = Utilities.base64Decode(imageBase64, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decoded, 'image/png', 'flying-pig-from-imagen.png');
      DriveApp.createFile(blob);
      Logger.log(`'flying-pig-from-imagen.png' をドライブに保存しました。`);
    } else {
      Logger.log('⚠️ 画像の生成に失敗しました。');
    }

    Logger.log("\n📚 Gemini を使って物語を生成します...");
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const geminiClient = new GeminiClient('gemini-1.5-pro-latest'); 
    geminiClient.setPromptText(storyPrompt);
    const response = geminiClient.generateCandidates();
    const story = (response.candidates[0].content.parts || []).map(p => p.text).join('');
    if (story) {
      Logger.log('✅ 生成された物語:\n' + story);
    } else {
       Logger.log('⚠️ 物語の生成に失敗しました。');
    }

  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * URLコンテキストのみを使用したグラウンディングのサンプル (OAuth認証で実行)
 */
function runUrlContextOnlyExample() {
  try {
    const model = 'gemini-1.5-pro-latest';
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const client = new GeminiClient(model); 
    
    client.enableUrlContextTool();
    
    const url = "https://www.nasa.gov/missions/artemis/";
    const promptText = `このURL (${url}) の内容を3つのポイントで要約してください。`;
    client.setPromptText(promptText);

    const response = client.generateCandidates();
    const text = (response.candidates[0].content.parts || []).map(p => p.text).join('');
    if (text) {
      Logger.log('✅ URLコンテキストのみに基づいた回答:\n' + text);
    } else {
      Logger.log('⚠️ モデルからテキスト応答を取得できませんでした。');
    }
  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * Google検索とURLによるグラウンディングを使用した実行サンプル関数 (APIキー認証で実行)
 */
function runGroundingExample() {
  try {
    const model = 'gemini-1.5-pro-latest';
    // Google検索ツールを有効化するため、コンストラクタでAPIキーを渡します。
    // getApiKey_()はスクリプトプロパティからAPIキーを読み込みます。
    const client = new GeminiClient(model, getApiKey_()); 
    
    client.enableGoogleSearchTool();
    // URLコンテキストツールも有効化 (必要に応じて)
    // client.enableUrlContextTool(); 
    
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
 * 「思考(Thinking)」機能を使用した実行サンプル関数 (OAuth認証で実行)
 */
function runThinkingExample() {
  try {
    const model = 'gemini-2.5-pro-preview-06-05'; 
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const client = new GeminiClient(model); 
    
    const prompt = `Alice, Bob, Carolはそれぞれ赤、緑、青の家に住んでいます。
    - 赤い家に住んでいる人は猫を飼っています。
    - Bobは緑の家には住んでいません。
    - Carolは犬を飼っています。
    - 緑の家は赤い家の左隣にあります。
    - Aliceは猫を飼っていません。
    それぞれ誰がどの家に住んでいて、何のペットを飼っていますか？`;

    client.setPromptText(prompt)
          .enableThinkingSummary(true) // 思考の要約を有効化
          .setThinkingBudget(8192);   // 思考のトークン上限を設定

    const response = client.generateCandidates();

    Logger.log('--- APIからの応答 ---');
    if (response.candidates && response.candidates[0].content.parts) {
      response.candidates[0].content.parts.forEach(part => {
        if (part.thought) {
          Logger.log('🧠 思考の要約:\n' + part.text);
        } else {
          Logger.log('✅ 最終的な回答:\n' + part.text);
        }
      });
    }

    // トークン使用量のログ出力
    if (response.usageMetadata) {
       Logger.log('\n--- トークン使用量 ---');
       Logger.log(`思考トークン: ${response.usageMetadata.thoughtsTokenCount || 0}`);
       Logger.log(`出力トークン: ${response.usageMetadata.candidatesTokenCount || 0}`);
       Logger.log(`合計トークン: ${response.usageMetadata.totalTokenCount || 0}`);
    }

  } catch (e) {
    Logger.log(`❌ エラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

/**
 * コード実行ツールを使用した実行サンプル関数 (OAuth認証で実行)
 */
function runCodeExecutionExample() {
  try {
    const model = 'gemini-1.5-pro-latest';
    // APIキーを渡さないことで、OAuth認証が試みられます。
    const client = new GeminiClient(model); 
    client.setCodeExecutionTool(true);
    client.setPromptText("Pythonコードを実行して、2の16乗を計算し、その結果だけを答えてください。");
    const response = client.generateCandidates();
    const codePart = response.candidates[0].content.parts.find(p => p.executable_code);
    if(codePart){
      Logger.log('✅ モデルが提案した実行コード:\n' + codePart.executable_code.code);
    }
    const text = (response.candidates[0].content.parts || []).filter(p=>p.text).map(p => p.text).join('');
    if(text){
       Logger.log('✅ 実行結果を含むモデルの回答:\n' + text);
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
   * APIリクエストを実際に実行します。
   * このメソッドは、認証（AuthorizationヘッダーやAPIキーURLパラメータ）は行いません。
   * 認証情報は、呼び出し元（例: GeminiClient.generateCandidates）で設定してください。
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
   * このメソッドは、認証（AuthorizationヘッダーやAPIキーURLパラメータ）は行いません。
   * 認証情報は、呼び出し元で設定されたoptions.headersを使用します。
   * @param {string} url - リクエスト先のURL。
   * @param {object} payload - 送信するJSONペイロード。
   * @returns {object} - APIからのJSONレスポンス。
   * @protected
   */
  requestPostAPI(url, payload) {
    const options = {
      method: 'POST',
      headers: this.headers, // headersはconstructorで初期化されたもの、またはsetHeadersで設定されたものを使用
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
   * @param {string | null} [apiKey=null] - オプション。このインスタンスで使用するAPIキー。
   * 指定された場合、OAuthトークンより優先して利用されます。
   * @constructor
   */
  constructor(model, apiKey = null) {
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
    /**
     * @private
     * コンストラクタで提供されたAPIキー。これが存在する場合、特定のケースでOAuthより優先されます。
     */
    this.userProvidedApiKey = apiKey;
    /**
     * @private
     * JSON形式の出力が有効になっているかを示すフラグ。
     */
    this.jsonOutputEnabled = false;
  }
  
  /**
   * File APIを使用してファイルをアップロードします。認証にはAPIキーが必要です。
   * File APIは常にAPIキーを使用します。
   * @param {GoogleAppsScript.Base.Blob} blob - アップロードするBlobオブジェクト。
   * @returns {object} - アップロードされたファイル情報。
   */
  uploadFile(blob) {
    if (!blob || typeof blob.getBytes !== 'function') {
      throw new Error('無効なBlobオブジェクトです。');
    }
    
    const apiKey = getApiKey_();
    if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
      throw new Error('APIキーが設定されていません。「getApiKey_」関数を編集するか、スクリプトプロパティに「GOOGLE_API_KEY」として設定してください。');
    }

    const fileBytes = blob.getBytes();
    
    // 1. アップロードセッションを開始し、アップロード用URLを取得
    const startSessionUrl = `${this.uploadUrl}?key=${apiKey}`; 

    const startOptions = {
      method: 'POST',
      headers: {
        'X-Goog-Upload-Protocol': 'resumable',
        'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': fileBytes.length.toString(), 
        'X-Goog-Upload-Header-Content-Type': blob.getContentType(),
        'Content-Type': 'application/json; charset=UTF-8'
      },
      payload: JSON.stringify({
        'file': { 'displayName': blob.getName() || 'untitled' }
      }),
      muteHttpExceptions: true
    };
    
    const startResponse = UrlFetchApp.fetch(startSessionUrl, startOptions);
    const responseCode = startResponse.getResponseCode();
    if (responseCode >= 200 && responseCode < 300) {
      // 応答ヘッダーの取得
      const responseHeaders = startResponse.getHeaders();
      const uploadUrl = responseHeaders['x-goog-upload-url']; // ヘッダー名は小文字になる

      if (!uploadUrl) {
        Logger.log('アップロードURLの取得に失敗しました。');
        Logger.log('レスポンスヘッダー: ' + JSON.stringify(responseHeaders, null, 2));
        Logger.log('レスポンスボディ: ' + startResponse.getContentText());
        throw new Error('Could not get upload URL from server response.');
      }

      // 2. 取得したURLにファイル本体をアップロードし、完了させる
      const uploadOptions = {
        method: 'POST',
        headers: {
          'X-Goog-Upload-Command': 'upload, finalize',
          'X-Goog-Upload-Offset': '0'
        },
        contentType: blob.getContentType(),
        payload: fileBytes,
        muteHttpExceptions: true
      };
      
      const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
      const uploadResponseCode = uploadResponse.getResponseCode();

      if (uploadResponseCode >= 200 && uploadResponseCode < 300) {
        const responseJson = JSON.parse(uploadResponse.getContentText());
        const fileInfo = responseJson.file || responseJson; 
        if (!fileInfo.uri) {
          throw new Error(`File upload succeeded, but the response is invalid: ${uploadResponse.getContentText()}`);
        }
        return fileInfo;
      } else {
        throw new Error(`File upload failed. Status: ${uploadResponseCode}, Body: ${uploadResponse.getContentText()}`);
      }
    } else {
      throw new Error(`Failed to start upload session. Status: ${responseCode}, Body: ${startResponse.getContentText()}`);
    }
  }

  /**
   * アップロード済みのファイルをプロンプトに添付します。
   * @param {object} fileInfo - `uploadFile`から返されたファイル情報オブジェクト。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  attachUploadedFile(fileInfo) {
    if (!fileInfo || !fileInfo.uri || !fileInfo.mimeType) {
        throw new Error('無効なファイル情報オブジェクトです。uriとmimeTypeプロパティが必要です。');
    }
    let userContent = this.promptContents.contents.find(c => c.role === 'user');
    if (!userContent) {
      userContent = { role: 'user', parts: [] };
      this.promptContents.contents.push(userContent);
    }
    userContent.parts.push({
      "fileData": {
        "mimeType": fileInfo.mimeType,
        "fileUri": fileInfo.uri
      }
    });
    return this;
  }

  /**
   * ファイルをプロンプトに添付します。File APIではなくインラインデータとして添付します。
   * 大容量ファイルには適していません（最大 20MB 程度の制限）。
   * @param {GoogleAppsScript.Base.Blob | GoogleAppsScript.Base.Blob[]} blobs - 添付するBlobまたはBlobの配列。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
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
   * 現在のプロンプト内容でモデルにリクエストを送信し、応答を生成します。
   * コンストラクタでAPIキーが設定されているか、Google検索ツールが有効かによって認証方法を決定します。
   * @returns {object} - APIからの完全なJSONレスポンス。
   */
  generateCandidates() {
    const url = `${this.baseUrl}:generateContent`;
    let finalUrl = url;
    let options = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(this.promptContents),
      muteHttpExceptions: true
    };

    // --- 認証ロジック ---
    const useGoogleSearchTool = this.promptContents.tools && 
                                this.promptContents.tools.some(tool => tool.hasOwnProperty('google_search'));

    if (useGoogleSearchTool) {
      // Google Search Toolが有効な場合はAPIキーを強制的に使用
      const apiKey = getApiKey_(); // getApiKey_()はスクリプトプロパティから取得を試みる
      if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
        throw new Error('Google検索ツールを利用するにはAPIキーが必要です。スクリプトプロパティ「GOOGLE_API_KEY」を設定してください。');
      }
      finalUrl = `${url}?key=${apiKey}`;
    } else if (this.userProvidedApiKey) {
      // コンストラクタでAPIキーが提供されている場合、それを認証に使う
      finalUrl = `${url}?key=${this.userProvidedApiKey}`;
    } else {
      // Google Search Toolも有効でなく、APIキーも提供されていない場合、OAuthトークン認証を試みる
      let accessToken = null;
      try {
        accessToken = ScriptApp.getOAuthToken();
      } catch (e) {
        // OAuthトークンの取得に失敗した場合、ログに出力
        Logger.log("OAuthトークンの取得に失敗しました (GCPプロジェクト未関連付け、または未承認の可能性): " + e.message);
      }

      if (accessToken) {
        options.headers['Authorization'] = 'Bearer ' + accessToken;
      } else {
        throw new Error('APIにアクセスできません。GCPプロジェクトを関連付けて承認を行うか、Google検索ツールを利用するか、またはコンストラクタでAPIキーを渡してください。');
      }
    }

    // --- 応答モダリティ/MIMEタイプ ロジック ---
    // generationConfigはpromptContentsのコピーから始める
    const currentGenerationConfig = { ...this.promptContents.generationConfig };

    const useGrounding = this.promptContents.tools && (
        this.promptContents.tools.some(tool => tool.hasOwnProperty('google_search')) ||
        this.promptContents.tools.some(tool => tool.hasOwnProperty('url_context'))
    );

    if (useGrounding) {
        // グラウンディングツールが有効な場合、応答MIMEタイプを text/plain に強制し、他のモダリティ設定をクリア
        currentGenerationConfig.responseMimeType = "text/plain";
        delete currentGenerationConfig.responseModalities;
        Logger.log('グラウンディングツールが有効なため、応答MIMEタイプを text/plain に設定しました。');
    } else if (this.jsonOutputEnabled) {
        // JSON出力が有効な場合、応答MIMEタイプを application/json に設定
        currentGenerationConfig.responseMimeType = "application/json";
        delete currentGenerationConfig.responseModalities; // 他のモダリティ指定と競合するためクリア
        Logger.log('JSON出力が有効なため、応答MIMEタイプを application/json に設定しました。');
    } else if (currentGenerationConfig.responseModalities && currentGenerationConfig.responseModalities.includes("IMAGE")) {
        // 画像応答が有効な場合、responseModalitiesを維持し、responseMimeTypeをクリア
        delete currentGenerationConfig.responseMimeType;
        Logger.log('画像応答が有効なため、responseModalitiesが使用されます。');
    } else {
        // デフォルトのテキスト応答設定
        delete currentGenerationConfig.responseMimeType;
        delete currentGenerationConfig.responseModalities;
        Logger.log('デフォルトのテキスト応答設定を使用します。');
    }

    // 更新されたgenerationConfigをペイロードに適用
    options.payload = JSON.stringify({
        ...this.promptContents,
        generationConfig: currentGenerationConfig
    });

    return this.executeRequest(finalUrl, options);
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
   * URLコンテキストツールを有効化します。プロンプト内のURLをモデルが参照できるようになります。
   * @param {boolean} [enable=true] - 有効にするか無効にするか。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  enableUrlContextTool(enable = true) {
    this._enableTool('url_context', enable);
    return this;
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

  /**
   * コード実行ツールを有効化します。モデルがPythonコードを実行できるようになります。
   * @param {boolean} [enable=true] - 有効にするか無効にするか。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setCodeExecutionTool(enable = true) {
    this._enableTool('code_execution', enable);
    return this;
  }
  
  /**
   * モデルからの応答に画像を含めるように要求します。
   * @param {boolean} [enable=true] - 有効にするか無効にするか。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  enableImageResponse(enable = true) {
    if (enable) {
      this.promptContents.generationConfig.responseModalities = ["TEXT", "IMAGE"];
      // 画像応答が有効になったら、JSON出力フラグを無効にする
      if (this.jsonOutputEnabled) {
        Logger.log('警告: 画像応答とJSON出力は同時に有効にできません。画像応答が優先されます。');
        this.jsonOutputEnabled = false;
      }
    } else {
      delete this.promptContents.generationConfig.responseModalities;
    }
    return this;
  }

  /**
   * JSON形式での構造化出力を有効にします。
   * グラウンディングツール（Google検索、URLコンテキスト）が有効な場合、この設定は上書きされ、
   * 応答は強制的に `text/plain` になります。
   * 画像応答 (`enableImageResponse`) と同時に有効にすることはできません。
   * @param {boolean} [enable=true] - JSON出力を有効にするか。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  enableJsonOutput(enable = true) {
    this.jsonOutputEnabled = enable;
    if (enable && this.promptContents.generationConfig && this.promptContents.generationConfig.responseModalities && this.promptContents.generationConfig.responseModalities.includes("IMAGE")) {
      Logger.log('警告: JSON出力と画像応答は同時に有効にできません。JSON出力が優先されます。');
      // JSON出力が優先されるので、画像応答のモダリティをクリア
      delete this.promptContents.generationConfig.responseModalities; 
    }
    return this;
  }

  /**
   * (試験運用版) 思考の要約を有効にします。
   * モデルの内部推論プロセスに関する要約をレスポンスに含めるように要求します。
   * この機能は Gemini 2.5 シリーズのモデルでサポートされています。
   * @param {boolean} [enable=true] - 思考の要約を有効にするか。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  enableThinkingSummary(enable = true) {
    if (!this.promptContents.generationConfig) {
      this.promptContents.generationConfig = {};
    }
    if (!this.promptContents.generationConfig.thinkingConfig) {
      this.promptContents.generationConfig.thinkingConfig = {};
    }
    this.promptContents.generationConfig.thinkingConfig.includeThoughts = enable;
    return this;
  }

  /**
   * (試験運用版) 思考予算を設定します。
   * 回答生成時に使用できる思考トークンの上限をモデルに指示します。
   * この機能は Gemini 2.5 シリーズのモデルでサポートされています。
   * @param {number} budget - 思考に使用するトークンの上限 (例: 1024, 8192)。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setThinkingBudget(budget) {
    if (typeof budget !== 'number' || budget < 0) {
      throw new Error('思考予算は0以上の数値を指定してください。');
    }
    if (!this.promptContents.generationConfig) {
      this.promptContents.generationConfig = {};
    }
    if (!this.promptContents.generationConfig.thinkingConfig) {
      this.promptContents.generationConfig.thinkingConfig = {};
    }
    this.promptContents.generationConfig.thinkingConfig.thinkingBudget = budget;
    return this;
  }

  /**
   * 生成されるテキストのランダム性を設定します (0.0〜1.0)。
   * @param {number} temperature - 温度。値が高いほど創造的になります。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setTemperature(temperature){ this.promptContents.generationConfig.temperature = temperature; return this }
  
  /**
   * Top-Pサンプリングの累積確率を設定します。
   * @param {number} topP - Top-P値。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setTopP(topP){ this.promptContents.generationConfig.topP = topP; return this }

  /**
   * Top-Kサンプリングで考慮するトークンの数を設定します。
   * @param {number} topK - Top-K値。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setTopK(topK){ this.promptContents.generationConfig.topK = topK; return this }

  /**
   * 生成される最大トークン数を設定します。
   * @param {number} maxOutputTokens - 最大トークン数。
   * @returns {GeminiClient} - メソッドチェーンのための自身。
   */
  setMaxOutputTokens(maxOutputTokens){ this.promptContents.generationConfig.maxOutputTokens = maxOutputTokens; return this }
}

/**
 * @class ImagenClient
 * @classdesc Imagenモデルを使用して画像を生成するためのクライアントクラス。
 * @extends RequestAPI
 */
class ImagenClient extends RequestAPI {
  /**
   * @constructor
   */
  constructor() {
    super();
    this.baseUrl = 'https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict';
  }

  /**
   * 指定されたプロンプトから画像を生成します。
   * @param {string} prompt - 画像生成のプロンプト。
   * @param {object} [options] - オプション。
   * @param {number} [options.sampleCount=1] - 生成する画像の数。
   * @param {string} [options.aspectRatio="1:1"] - 画像のアスペクト比。
   * @returns {string[]} - Base64エンコードされた画像の文字列の配列。
   */
  generate(prompt, options = {}) {
    const payload = {
      instances: [{ "prompt": prompt }],
      parameters: {
        sampleCount: options.sampleCount || 1,
        aspectRatio: options.aspectRatio || "1:1"
      }
    };
    // Imagen APIはAPIキー認証を必要とします
    const apiKey = getApiKey_();
    if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
      throw new Error('Imagen APIを利用するにはAPIキーが必要です。スクリプトプロパティ「GOOGLE_API_KEY」を設定してください。');
    }
    const urlWithKey = `${this.baseUrl}?key=${apiKey}`;

    const response = this.requestPostAPI(urlWithKey, payload); // requestPostAPIはheadersを内部で設定
    
    if (!response.predictions || !Array.isArray(response.predictions)) {
      throw new Error('画像データの取得に失敗しました。');
    }
    return response.predictions.map(pred => pred.bytesBase64Encoded);
  }
}

/**
 * @class EmbeddingClient
 * @classdesc テキストからエンベディング（ベクトル表現）を生成するためのクライアントクラス。
 * @extends RequestAPI
 */
class EmbeddingClient extends RequestAPI {
  /**
   * @param {string} model - 使用するエンベディングモデル名 (例: 'text-embedding-004')。
   * @constructor
   */
  constructor(model) {
    super();
    if (!model) {
      throw new Error('エンベディングモデル名は必須です。');
    }
    this.model = model;
    this.baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${this.model}`;
  }

  /**
   * テキストのエンベディング（ベクトル）を生成します。
   * @param {string} text - エンベディングするテキスト。
   * @param {string} [taskType] - タスクの種類 (例: 'RETRIEVAL_DOCUMENT', 'RETRIEVAL_QUERY')。
   * @returns {number[]} - 生成されたエンベディング（ベクトル）。
   */
  generate(text, taskType = null) {
    const url = `${this.baseUrl}:embedContent`;
    
    if (Array.isArray(text)) {
      console.warn("EmbeddingClient.generateに配列が渡されましたが、最初の要素のみ処理されます。バッチ処理にはbatchEmbedContentsエンドポイントを実装したクライアントを使用してください。");
      text = text[0];
    }

    const payload = {
      "content": { "parts": [{ "text": text }] }
    };

    // Embedding APIもAPIキー認証を必要とします
    const apiKey = getApiKey_();
    if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
      throw new Error('Embedding APIを利用するにはAPIキーが必要です。スクリプトプロパティ「GOOGLE_API_KEY」を設定してください。');
    }
    const urlWithKey = `${url}?key=${apiKey}`;

    const response = this.requestPostAPI(urlWithKey, payload); // requestPostAPIはheadersを内部で設定
    
    if (response.embedding && response.embedding.values) {
        return response.embedding.values;
    } else {
      throw new Error("生成されたエンベディングの形式が不正です。");
    }
  }
}
