# GPT Image ロゴ編集・SharePoint画像参照 作業再開メモ

更新日: 2026-07-21
状態: 実装途中の検証段階で中断。コード変更はコミットされていない。

## 再開後の追記: 「差し替えて」が編集扱いされない問題

再開後の実環境ログでは、添付Base64は `imageAttachmentCount: 1` として正常に到着していたが、`[edit_existing_image] input images` まで到達しなかった。

確定原因:

- 画像編集意図を判定する動詞一覧に「差し替え」がなかった。
- 「荷台のMidacを添付ロゴに差し替えて」が画像編集と判定されなかった。
- 添付画像があるため、画像編集ではなく画像説明用の `multimodal` 経路へ流れていた。

対応:

- `差し替`、`置き換`、`置換`、`交換`、`反映`、`適用`、`使用`、`使って`、`にして`、`貼り`、`貼って` を画像編集動詞へ追加。
- Chat APIログへ `imageAttachmentCountForRouting` と `requiredImageToolName` を追加。
- 「添付のロゴに差し替えて」「添付ロゴに置き換えて」の回帰テストを追加。
- 回帰テストは49項目合格、TypeScript型検査・ESLintも合格。

再テスト時、次のログを先に確認すること:

```text
imageAttachmentCount: 1
imageAttachmentCountForRouting: 1
requiredImageToolName: 'edit_existing_image'
```

その後に次が出れば、ルーティングと参照画像受け渡しの両方が成功している。

```text
[edit_existing_image] input images: {
  referencesRequested: 1,
  referencesLoaded: 1
}
```

## 再開後の追記2: GPT Image完了後も画面が終了しない問題

実環境ログで次を確認した。

```text
referencesRequested: 1
referencesLoaded: 1
[edit_existing_image] gpt-image edit completed: { elapsedMs: 56581 }
```

この時点で、ロゴ参照とGPT Image編集は成功している。停止原因は画像APIではなく、その後のSSEストリーム終了通知だった。

確定原因:

- 画像編集ではOpenAI SDKの `runTools` に特定ツールを強制指定している。
- 使用中のSDKは強制ツールの実行後にrunnerを終了するが、assistant本文がないため `finalContent` イベントを発生させない。
- `open-ai-stream.ts` は `finalContent`、エラー、キャンセル時だけブラウザのストリームを閉じていた。
- ツールは完了してもブラウザへ終了通知が届かず、画面が無期限に処理中になった。

対応:

- runnerの `end` イベントを処理。
- `finalContent` がない場合は、最後のツール結果に含まれる画像URLからMarkdown画像付き最終回答を生成。
- ツール結果と最終回答を保存して、SSEへ `finalContent` を送信してからストリームを閉じる。
- DB保存に失敗してもブラウザストリームは必ず閉じる。
- 成功・エラー・キャンセルの終了処理が重複しないよう `finalized` ガードを追加。
- 画像保存完了ログ `[edit_existing_image] output saved` を追加。

再テスト時の期待ログ:

```text
[edit_existing_image] gpt-image edit completed: ...
[edit_existing_image] output saved: ...
[open-ai-stream] runner ended without finalContent; using fallback
```

その後、画面のぐるぐるが終了し、編集画像が表示されること。

## 再開後の追記3: 終了したが修正画像が画面に出ない問題

追記2の修正後、次のログまで正常に出て、ぐるぐるも終了したが、実行中の画面へ修正画像が追加されなかった。

```text
[edit_existing_image] output saved: ...
[open-ai-stream] runner ended without finalContent; using fallback
```

確定原因:

- サーバーはツール結果から作ったフォールバック本文を `finalContent` として送信していた。
- クライアントの `chat-store.tsx` は通常の `content` イベントでのみassistantメッセージを画面へ追加していた。
- `finalContent` の処理はloading解除・読み上げ・タイトル更新だけで、メッセージ追加を行っていなかった。
- 強制ツール終了では通常の `content` が存在しないため、最終本文とMarkdown画像が画面へ表示されなかった。

対応:

- `finalContent` 受信時にもassistantメッセージを追加または更新。
- 通常のストリーミング本文がある場合は同じmessage IDを更新し、二重表示を防止。
- `content` がないツール終了の場合は新しいmessage IDを作り、画像付きフォールバック本文を表示。
- 画像回帰テストは55項目合格、TypeScript型検査・ESLintも合格。

既に `output saved` が出た画像はBlob保存済み。サーバー側の最終メッセージ保存も成功していれば、ブラウザ再読み込みだけで履歴に表示される可能性がある。

## 実環境テスト成功（2026-07-21）

追記3のクライアント表示修正後、ユーザーが実環境で再確認し、修正画像がチャット画面へ正常に表示された。

成功した処理経路:

1. 生成済みパッカー車をスレッドの最新画像として取得。
2. 紙クリップで添付した `midac_logo.png` を参照画像として取得。
3. `requiredImageToolName: 'edit_existing_image'` で画像編集へルーティング。
4. GPT Imageへベース画像1枚とロゴ参照画像1枚を送信。
5. GPT Image編集完了。
6. 完成画像をBlobへ保存。
7. 強制ツール終了時のフォールバック最終メッセージを生成。
8. SSEを正常終了。
9. クライアントが `finalContent` をassistantメッセージとして追加。
10. 修正画像をチャット画面へ表示。

実環境で確認した主要ログ:

```text
requiredImageToolName: 'edit_existing_image'

[edit_existing_image] input images: {
  referencesRequested: 1,
  referencesLoaded: 1,
  inputBytes: [1465884, 37378]
}

[edit_existing_image] gpt-image edit completed: {
  elapsedMs: 59244
}

[edit_existing_image] output saved: {
  imageName: '4YDLTfqN0mX5e61kKTvrImxnOpQrgpbqwPyc.png',
  bytes: 2006346,
  elapsedMs: 63627
}

[open-ai-stream] runner ended without finalContent; using fallback
```

実環境テスト結果:

- 添付ロゴの参照入力: 成功
- 既存画像の編集ルーティング: 成功
- GPT Image編集: 成功
- Blob保存: 成功
- 無限ぐるぐるの解消: 成功
- 修正画像のチャット表示: 成功

紙クリップ添付ロゴを生成済み画像へ差し替える基本経路は、実環境で動作確認済みとなった。

残る確認項目:

- SharePoint／AI Search上の画像をファイル名指定して利用する経路の実環境テスト。
- ロゴの形状・文字・色の再現品質確認。
- 正確なピクセル一致が必要な場合の決定的PNG合成方式の検討。

## 調査中の誤認記録（2026-07-21）

調査中にClaude（Sonnet）が2点の誤認を行った。後続の作業者が同じ誤りを繰り返さないよう記録する。

### 誤認①: `storedImageAttachment.buffer` が `executeEditExistingImage` へ渡らない

**誤認した内容:**
`chat-api.ts` の `LoadLatestImageAttachment` は `storedImageAttachment` をルーティング判定のために取得するが、その `.buffer` は `executeEditExistingImage` へ引き渡されていないと主張した。「ロゴバッファがルーティング後に失われている」という誤った前提のもと、data URLへ変換してから渡す修正案を提案した。

**正しい動作:**
`executeEditExistingImage` は `chat-api.ts` とは独立して、内部で `LoadLatestImageAttachment(chatThread.id)` を再呼び出しする（`chat-api-default-extensions.ts` 行7881付近）。取得したバッファを `normalizeAzureEditImage` で正規化し、`loadedReferenceBuffers` へ追加している（行7997付近）。
つまり、ロゴは`chat-api.ts`で捨てられているのではなく、ツール実行側が自分で取りに行く設計になっている。提案した修正を適用していたら、ロゴが2重送信されてAzureから `invalid_image_file` エラーが返る可能性があった。

**正しいフロー:**

```
chat-api.ts
  └─ LoadLatestImageAttachment()  ← ルーティング用（カウント目的のみ）
  └─ resolveRequiredImageToolName() → 'edit_existing_image'
  └─ ChatApiExtensions() へ渡す（bufferは渡さない）

executeEditExistingImage (chat-api-default-extensions.ts)
  └─ LoadLatestImageAttachment()  ← 内部で独自に再取得
  └─ normalizeAzureEditImage()
  └─ loadedReferenceBuffers.push()
  └─ GPT Image API へ送信
```

### 誤認②: 2回呼ばれるのは「1回目=アップロード、2回目=AI処理」

**誤認した内容:**
`ChatAPIEntry` が2回呼ばれるログを見て、「1回目がファイルアップロード専用、2回目がAI処理」と解釈した。

**正しい動作:**
両方の呼び出しとも `ChatAPIEntry`（AI処理）であり、内容が異なる2回のメッセージに対応しているだけ。ファイルのアップロードは `/api/sl/publish` が別経路で処理しており、`ChatAPIEntry` とは無関係に走る。2回のログが重なって見えた理由は、アップロード完了通知メッセージとロゴ差し替え指示の2連続送信によるもの。

## 目的

次の2つの利用方法を成立させる。

1. 紙クリップで添付したロゴを、生成済み画像の一部へ組み込む。
2. 「SharePointにある `midac_logo.png` を使って」のように指定し、AI Searchで画像を特定してGPT Imageの参照画像へ渡す。

主なテスト例:

- 最初に「白基調で、荷台にMidacと入れたパッカー車を描いて」
- 次に `midac_logo.png` を添付して「荷台のMidacを添付ロゴに差し替えて」
- または「SharePointにあるmidac_logo.pngを使って、荷台のMidacをそのロゴに差し替えて」
- 新規生成時に「添付ロゴを使って白いパッカー車をデザインして」

## 発生していた現象と確定原因

### 添付ロゴと無関係な青いマークに置き換わった

問題発生時のログ:

```text
[edit_existing_image] input images: {
  base: 'thread:__latest__.png',
  referencesRequested: 0,
  referencesLoaded: 0
}
```

GPT Imageへ渡っていたのはパッカー車のベース画像だけで、`midac_logo.png` の画素は渡っていなかった。モデルは文章だけから架空の青いロゴを生成した。

これはモデルのロゴ再現精度以前の問題で、参照画像の受け渡し欠落が直接原因。

### 以前に発生した関連問題

- `create_img` が選ばれ、添付画像を受け取らず新規画像を生成していた。
- 画像URLを取得したつもりでもHTMLページが返り、画像として読み取れなかった。
- 同じ画像が3枚目として重複送信され、Azureから `invalid_image_file` が返った。
- 紙クリップ画像がブラウザ内Base64だけで処理され、通常のSharePointアップロード表示・保存経路を通っていなかった。
- 通常添付画像のプレビューへ、依頼文をCanvasで重ねて表示していた。

## 現在までの実装

### 紙クリップ画像

- PNG/JPEG/WebPを画像素材として判定。
- BlobとSharePointへ保存。
- 文書用の `CrackDocument`、RAG本文解析、チャット文書登録は実行しない。
- GPT Image用のBase64はアップロード開始前に画面側へ保持。
- アップロード中の送信を防止。
- 通常添付画像へのCanvas文字合成を停止。

### 画像編集ルーティング

- 画像添付を伴う生成・合成・差し替え依頼は `edit_existing_image` を選択。
- 明確な文字入れ依頼だけ `add_text_to_existing_image` を使用。
- スレッドの `__latest__.png` を既存画像編集のベースとして優先。
- 画像実体の形式検査、WebPからPNGへの正規化、SHA-256による重複排除を追加。
- GPT Image編集はタイムアウト3分、SDK自動再試行なし。

### 添付ロゴの短期サーバー保持

対象: `src/features/chat-page/chat-services/chat-image-service.ts`

- 紙クリップ画像を `__latest_attachment__.img` として画像コンテナへ一時保存。
- ファイル名、MIME、保存日時、サイズをJSONで保存。
- 既定の有効期限は30分。
- 環境変数 `IMAGE_ATTACHMENT_TTL_MS` で変更可能。
- チャットのBase64が欠落しても、「添付」「ロゴ」等を明示した依頼なら直近素材を復元。
- 編集成功後は使用済みとして記録し、別の依頼へ古いロゴを持ち越さない。
- APIエラー時は未使用のまま残るため再試行可能。

### SharePoint画像をAI Searchから使う経路

対象:

- `src/features/chat-page/chat-input/file/file-store.ts`
- `src/features/chat-page/chat-services/chat-api/image/image-intent.ts`
- `src/features/chat-page/chat-services/chat-api/chat-api.ts`
- `src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts`

実装内容:

1. SharePointへ画像をアップロードした後、本文解析はせず、ファイル名・URL・部署・公開範囲・所有者・SharePoint item IDをAI Searchへ登録。
2. 「SPにあるLogoを使って」「SharePointにあるmidac_logo.pngを使って」を画像参照依頼として認識。
3. 既存のAI Search ACLフィルターを使い、ユーザーがアクセス可能なSharePoint画像だけ検索。
4. 対象ファイルをGraph APIで取得し、スレッド用Blobへ一時キャッシュ。
5. PNG/JPEG/WebPとして検証後、GPT Imageの参照画像へ追加。
6. 複数ファイルが該当する場合は勝手に選択せず、拡張子を含む正確なファイル名を要求。
7. 画像メタデータの索引だけ失敗した場合、SharePointアップロード自体は成功扱いを維持し、エラー通知だけ表示。

## ローカル検証結果

最後に実行したコマンド:

```powershell
cd C:\Users\021213\azurechat-office-work\src
node scripts/test-image-upload-routing.cjs
npx tsc --noEmit --pretty false
npm run lint

cd C:\Users\021213\azurechat-office-work
git diff --check
```

結果:

- 画像アップロード・編集回帰テスト: 47項目合格
- TypeScript型検査: 合格
- ESLint: 警告・エラーなし
- `git diff --check`: エラーなし（改行コード警告のみ）

## 未実施・再開時に必ず確認すること

### 1. 実環境E2Eテスト

SharePoint、Azure AI Search、Graph API、GPT Imageを通した実環境テストは未実施。会社データへのアクセスと画像API課金を伴うため、ローカルの静的・回帰テストまでで中断した。

開発サーバーを再起動してからテストする。

```powershell
cd C:\Users\021213\azurechat-office-work\src
npm run dev
```

### 2. 紙クリップ添付による差し替え

期待ログ:

```text
[edit_existing_image] input images: {
  referencesRequested: 1,
  referencesLoaded: 1,
  storedAttachment: 'midac_logo.png'
}
```

`referencesLoaded: 0` のままなら、ロゴはGPT Imageへ渡っていないため生成を続けてはいけない。

### 3. SharePoint画像による差し替え

推奨テスト文:

```text
SharePointにあるmidac_logo.pngを使って、荷台のMidac表記をそのロゴに差し替えて。他の部分は変更しないで。
```

期待ログ:

```text
referencesRequested: 1
referencesLoaded: 1
sharePointReference: 'midac_logo.png'
```

確認ポイント:

- 画像アップロード時に `Indexing image metadata` が表示されること。
- AI Searchに `metadata = midac_logo.png` の画像レコードが存在すること。
- ACLにより個人・部署・全社共通の範囲が正しく制御されること。
- Graph取得後のURLがHTMLではなく画像バイナリを返すこと。
- GPT Image呼び出し直前に `referencesLoaded: 1` 以上になること。

### 4. 既にSharePointに存在する画像

今回のアップロード経路を通っていない既存画像は、AI Searchに画像メタデータがなければ検索できない。

再開時の選択肢:

- 今回の画面から画像を再アップロードして索引する。
- SharePoint同期処理を拡張し、既存PNG/JPEG/WebPもメタデータ索引対象にする。

後者はまだ実装していない。

### 5. ロゴの完全一致

`referencesLoaded: 1` になれば参照画像は正しく渡っているが、GPT Imageによる再描画ではロゴの文字・輪郭・縦横比が完全一致する保証はない。

正確な企業ロゴが必須の場合は、次の方式を別実装として検討する。

1. GPT Imageでロゴなしの車両画像を生成。
2. 荷台の貼付位置を決定。
3. 元の透過PNGを通常の画像合成処理で変形・貼付。

この決定的合成方式は未実装。

## 作業ツリー上の注意

- 作業ツリーには今回以前からの変更や未追跡ファイルが多数ある。
- 今回の画像修正だけを理由に、他の変更をリセット・削除しないこと。
- `git reset --hard`、`git checkout --` 等を使用しないこと。
- コミット前に今回分と既存変更を切り分けてレビューすること。

主な関連ファイル:

- `src/features/chat-page/chat-input/chat-input.tsx`
- `src/features/chat-page/chat-input/file/file-store.ts`
- `src/features/ui/chat/chat-input-area/input-image-store.ts`
- `src/features/chat-page/chat-services/chat-image-service.ts`
- `src/features/chat-page/chat-services/chat-api/image/image-intent.ts`
- `src/features/chat-page/chat-services/chat-api/chat-api.ts`
- `src/features/chat-page/chat-services/chat-api/chat-api-extension.ts`
- `src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts`
- `src/features/chat-page/message-content.tsx`
- `src/scripts/test-image-upload-routing.cjs`

## 2026-07-21 TestSiteのSharePointロゴ読込失敗への追加修正

TestSiteではAI Searchによる `midac_logo.png` の特定、Graph経由のBlob保存、SAS URLのHEAD確認までは成功したが、GPT Image呼び出し前に「参照画像の一部をPNG/JPEGとして読み取れない」と返った。

原因候補は、SharePoint画像を一時BlobのSAS URLから再取得する経路と、モデルが `referenceImageUrls` に `midac_logo.png` のようなURLではない値を追加する経路が、同じ参照配列に混在していたこと。1件でも読込失敗があると、正しいSharePoint画像を取得できていても処理全体が失敗する実装だった。

追加修正:

- SharePoint画像は一時BlobのSAS URLを再取得せず、Azure Storage SDKの `downloadToBuffer()` で直接読む。
- SDK直接読込が利用できない場合だけ、SAS URLのHTTP取得へフォールバックする。
- HTTP取得が非成功の場合、SASクエリを除外してHTTPステータス、Content-Type、ホストとパスを記録する。
- GPT Image実行前の `input images.base` もSASクエリを除外し、ホストとパスだけを記録する。`sig`、`se` 等はログへ出さない。
- モデル指定の `referenceImageUrls` はHTTP(S) URLまたは対応画像data URLだけを許可し、裸のファイル名を除外する。
- SharePoint画像をコード側で解決した場合、モデル生成の参照URLは混在させず、SDKで取得した信頼済みバッファを参照画像として使用する。
- 通常の紙クリップ添付画像と既存生成画像の編集経路は変更しない。

期待ログ:

```text
[edit_existing_image] SP image loaded via Blob SDK: {
  container: 'dl-link',
  blobPath: '<thread-id>/midac_logo.png',
  bytes: <positive number>,
  format: 'png'
}
[edit_existing_image] input images: {
  referencesLoaded: 1,
  sharePointReference: 'midac_logo.png'
}
```

検証結果:

- `node scripts/test-image-upload-routing.cjs`: 71 assertions passed
- `npx tsc --noEmit`: 成功
- `npm run lint`: warnings/errorsなし
- `git diff --check`: 問題なし
- `npm run build`: コードエラーではなく、起動中のNext.js開発サーバーが `src/.next/trace` を使用していたため `EPERM`。ビルドを再確認する場合は開発サーバー停止後に実行する。

## 再開時の最短手順

1. このメモと `git status --short` を確認。
2. `npm run dev` で開発サーバーを再起動。
3. 新しいチャットで `midac_logo.png` を紙クリップからアップロード。
4. SharePoint保存とAI Searchメタデータ索引の完了をログで確認。
5. 紙クリップ添付差し替えを1回テスト。
6. SharePointファイル名指定差し替えを1回テスト。
7. `referencesLoaded`、`storedAttachment`、`sharePointReference` を確認。
8. 出力ロゴが変形する場合、参照欠落かモデル再現限界かをログで切り分ける。
9. 必要なら決定的PNG合成方式を次の作業として設計する。
