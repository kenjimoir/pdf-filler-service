# PDF Filler Service

PDFテンプレート（AcroForm）にJSONの回答を流し込み、外観ストリームを正しく生成し、日本語・英語・数字・チェックボックスを安定して表示させた上で、編集可能（非フラット）のままGoogle Driveに保存するサービスです。

## 技術要件

- Node.js LTS (>=18)
- Express
- pdf-lib + @pdf-lib/fontkit（フォント埋め込み & 外観生成）
- チェックマーク: StandardFonts.ZapfDingbats を使用（Adobe Pi 依存を排除）
- Google Drive API: googleapis (Drive v3)、サービスアカウント認証
- デプロイ先: Render
- フォント: 日本語用 NotoSansCJKjp-Regular.otf（必須）＋英数字用 NotoSans-Regular.ttf（任意）

## セットアップ手順

### 1. 依存パッケージのインストール

```bash
npm install
```

### 2. フォントファイルの配置

以下のフォントファイルを `fonts/` ディレクトリに配置してください：

#### 必須フォント
- **NotoSansCJKjp-Regular.otf**: 日本語表示用
  - ダウンロード: [Google Fonts - Noto Sans CJK JP](https://fonts.google.com/noto/specimen/Noto+Sans+JP)
  - または: [GitHub - google/fonts](https://github.com/google/fonts/tree/main/ofl/notosanscjktc)

#### 任意フォント（推奨）
- **NotoSans-Regular.ttf**: 英数字表示用
  - ダウンロード: [Google Fonts - Noto Sans](https://fonts.google.com/noto/specimen/Noto+Sans)
  - または: [GitHub - google/fonts](https://github.com/google/fonts/tree/main/ofl/notosans)

フォントファイルは以下のライセンスに従います：
- Noto フォント: OFL (SIL Open Font License)
- pdf-lib: MIT License

### 3. 環境変数の設定

`.env.example` を `.env` にコピーして、必要な値を設定してください：

```bash
cp .env.example .env
```

#### 環境変数の説明

- **API_BEARER_TOKEN**: Apps Scriptから呼び出す際の認証トークン（任意の長いランダム文字列を推奨）
- **GOOGLE_SERVICE_ACCOUNT_BASE64**: GoogleサービスアカウントのJSONキーをBase64エンコードしたもの
- **DEFAULT_OUTPUT_FOLDER_ID**: デフォルトの出力フォルダID（任意）
- **PORT**: サーバーのポート番号（デフォルト: 3000）

#### Googleサービスアカウントの設定

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. 「APIとサービス」→「ライブラリ」から「Google Drive API」を有効化
3. 「認証情報」→「サービスアカウントを作成」でサービスアカウントを作成
4. サービスアカウントのJSONキーをダウンロード
5. JSONキーをBase64エンコード:
   ```bash
   base64 -i service-account-key.json
   ```
6. エンコードされた文字列を `GOOGLE_SERVICE_ACCOUNT_BASE64` に設定

#### Google Drive APIの権限設定

サービスアカウントのメールアドレスに対して、テンプレートPDFと出力フォルダに対する以下の権限を付与してください：
- **閲覧者**: テンプレートPDFの読み取り権限
- **編集者**: 出力フォルダへの書き込み権限

### 4. ローカルでの実行

```bash
npm start
```

サーバーが起動したら、`http://localhost:3000/health` にアクセスして動作確認できます。

## Render へのデプロイ

### 1. Render アカウントの作成

[Render](https://render.com/) でアカウントを作成し、GitHubリポジトリと連携します。

### 2. 新しいWebサービスの作成

1. Renderダッシュボードで「New +」→「Web Service」を選択
2. GitHubリポジトリを接続
3. 以下の設定を行います：
   - **Name**: `pdf-filler-service` (任意)
   - **Environment**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Plan**: Free または Starter

### 3. 環境変数の設定

Renderダッシュボードの「Environment」セクションで、以下の環境変数を設定します：

- `API_BEARER_TOKEN`: Apps Scriptから呼び出す際の認証トークン
- `GOOGLE_SERVICE_ACCOUNT_BASE64`: Base64エンコードされたサービスアカウントJSONキー
- `DEFAULT_OUTPUT_FOLDER_ID`: デフォルトの出力フォルダID（任意）
- `PORT`: Renderが自動設定するため、設定不要（または `3000`）

### 4. フォントファイルの配置

Renderでは、フォントファイルもGitリポジトリに含める必要があります。`fonts/` ディレクトリに必要なフォントファイルを配置し、Gitにコミットしてください。

**注意**: フォントファイルは大きいため、Git LFSの使用を検討してください。

### 5. デプロイ

設定を保存すると、Renderが自動的にデプロイを開始します。デプロイが完了したら、`https://your-service.onrender.com/health` にアクセスして動作確認してください。

## API仕様

### GET /health

ヘルスチェックエンドポイント

**レスポンス**
```json
{
  "ok": true
}
```

### POST /fill

PDFテンプレートにデータを流し込んでGoogle Driveに保存

**リクエストヘッダー**
```
Authorization: Bearer {API_BEARER_TOKEN}
Content-Type: application/json
```

**リクエストボディ（新形式）**
```json
{
  "templateId": "1OB7Ep8JlC32tkWt93uiYPQjpCiDxb1yp",
  "output": {
    "name": "申込書_山田太郎_2025-11-03.pdf",
    "folderId": "1Q2DuIctvy1n4YGSixA15BRQ-9gdWtQ"
  },
  "fields": {
    "ApplicantNameJP": "山田 太郎",
    "ApplicantNameEN": "Taro Yamada",
    "BirthDate": "1990-01-01",
    "Phone": "03-1234-5678",
    "AgreeTerms": true
  }
}
```

**リクエストボディ（旧形式 - Code.gs互換）**
```json
{
  "templateFileId": "1OB7Ep8JlC32tkWt93uiYPQjpCiDxb1yp",
  "outputName": "申込書_山田太郎_2025-11-03.pdf",
  "folderId": "1Q2DuIctvy1n4YGSixA15BRQ-9gdWtQ",
  "fields": {
    "ApplicantNameJP": "山田 太郎",
    "ApplicantNameEN": "Taro Yamada",
    "BirthDate": "1990-01-01",
    "Phone": "03-1234-5678",
    "AgreeTerms": true
  }
}
```

**レスポンス（成功時）**
```json
{
  "ok": true,
  "file": {
    "id": "1abc123...",
    "name": "申込書_山田太郎_2025-11-03.pdf",
    "webViewLink": "https://drive.google.com/file/d/1abc123.../view"
  },
  "driveFile": {
    "id": "1abc123...",
    "name": "申込書_山田太郎_2025-11-03.pdf",
    "parents": ["1Q2DuIctvy1n4YGSixA15BRQ-9gdWtQ"],
    "webViewLink": "https://drive.google.com/file/d/1abc123.../view"
  }
}
```

**レスポンス（エラー時）**
```json
{
  "ok": false,
  "error": "エラーメッセージ"
}
```

## テンプレートPDF要件

### AcroFormフィールド

テンプレートPDFは以下の要件を満たす必要があります：

1. **AcroForm形式**: Adobe Acrobatで作成されたフォームフィールドを含むPDF
2. **フィールド名**: `fields` オブジェクトのキーと一致するフィールド名を使用
3. **チェックボックスのexport値**: 
   - 一般的な値: `"On"`, `"Yes"`, `"Off"`, `"No"`
   - 小文字/大文字の違いに対応
   - export値が設定されていない場合は、デフォルトの `check()` を使用

### フィールドタイプの対応

- **テキストフィールド**: 日本語・英語・数字に対応（CJK判定により自動でフォントを選択）
- **チェックボックス**: ZapfDingbatsフォントを使用して外観を生成
- **ラジオボタン**: ZapfDingbatsフォントを使用して外観を生成
- **ドロップダウン**: テキストフィールドと同様に処理

## トラブルシュート

### 日本語が豆腐（□）で表示される

**原因**: CJKフォントが正しく読み込まれていない、または埋め込まれていない

**対処法**:
1. `fonts/NotoSansCJKjp-Regular.otf` が正しく配置されているか確認
2. フォントファイルのパスが正しいか確認
3. サーバーログでフォント読み込みエラーがないか確認

### 権限エラー（403 Forbidden）

**原因**: サービスアカウントに適切な権限が付与されていない

**対処法**:
1. テンプレートPDFにサービスアカウントのメールアドレスに「閲覧者」権限を付与
2. 出力フォルダにサービスアカウントのメールアドレスに「編集者」権限を付与
3. Google Cloud ConsoleでDrive APIが有効になっているか確認

### チェックボックスが表示されない

**原因**: 外観ストリームが正しく生成されていない

**対処法**:
1. テンプレートPDFのチェックボックスのexport値が正しいか確認
2. `fields` オブジェクトの値が `true`, `"true"`, `"on"`, `"yes"` などの正しい形式か確認
3. サーバーログでエラーメッセージを確認

### 401 Unauthorized

**原因**: Bearerトークンが正しく設定されていない、または一致していない

**対処法**:
1. `API_BEARER_TOKEN` 環境変数が正しく設定されているか確認
2. Apps Script側のリクエストヘッダーに `Authorization: Bearer {TOKEN}` が含まれているか確認
3. トークンの前後の空白文字がないか確認

### PDFがフラット化されて編集できない

**原因**: `pdfDoc.save()` のオプションが正しく設定されていない

**対処法**:
- 現在の実装では `updateFieldAppearances: false`, `useObjectStreams: false` で保存しているため、編集可能なPDFが生成されるはずです
- 問題が続く場合は、サーバーログでエラーメッセージを確認

### フォントファイルが見つからない（Render）

**原因**: Gitリポジトリにフォントファイルが含まれていない

**対処法**:
1. `fonts/` ディレクトリがGitリポジトリに含まれているか確認
2. `.gitignore` でフォントファイルが除外されていないか確認
3. フォントファイルが大きすぎる場合は、Git LFSの使用を検討

## ライセンス

- **pdf-lib**: MIT License
- **Noto フォント**: OFL (SIL Open Font License)
- **本プロジェクト**: MIT License（想定）

## 今後の拡張予定

- 数値/日付フォーマット機能
- フラット化オプション（編集不可PDFの生成）
- バッチ処理対応
- エラーログの詳細化

