/**
 * ===================================================================================
 * 【RAG with BigQuery & Spreadsheet】事前準備
 * ===================================================================================
 *
 * 1. BigQueryのテーブル準備:
 * a) ナレッジベース用テーブル:
 * - 以下のSQLをBigQueryコンソールで実行して、ベクトル化されたドキュメントチャンクを保存するテーブルを作成します。
 * `your_project_id`と`your_dataset`はご自身の環境に合わせてください。
 *
 * CREATE TABLE `your_project_id.your_dataset.knowledge_base` (
 * chunk_id STRING,
 * organization_id STRING,
 * account_id STRING,
 * source_document STRING,
 * chunk_text STRING,
 * embedding ARRAY<FLOAT64>
 * );
 *
 * b) アカウント管理用スプレッドシートの準備:
 * - このシステムが参照するアカウント情報を管理するスプレッドシートを用意します。
 * - 1行目にヘッダーとして、少なくとも以下の3つの列名を含めてください。
 * `id`, `organization_id`, `rag_gdrive_folder_id`
 * - 2行目以降に、管理したい企業（取引先）のアカウントID、所属する組織ID、
 * ナレッジが格納されているGoogle DriveフォルダのIDをそれぞれ入力します。
 *
 * 2. Apps Scriptの高度なサービスと権限の設定 (★★★重要★★★):
 * a) BigQuery APIの有効化:
 * - Apps Scriptエディタで、「サービス」>「＋」>「BigQuery API」を追加します。
 *
 * b) マニフェストファイルにスコープを追加:
 * - 「表示」>「マニフェスト ファイルを表示」を選択し、`appsscript.json`を開きます。
 * - `oauthScopes` のリストに、以下のスコープが含まれていることを確認してください。
 *
 * "oauthScopes": [
 * "https://www.googleapis.com/auth/script.external_request",
 * "https://www.googleapis.com/auth/drive.readonly",
 * "https://www.googleapis.com/auth/bigquery"
 * ],
 *
 * 3. 定時実行トリガーの設定:
 * - Apps Scriptエディタの「トリガー」から、`runDailyIndexUpdate` 関数を
 * 毎日深夜などに実行するように設定します。
 *
 * ===================================================================================
 */