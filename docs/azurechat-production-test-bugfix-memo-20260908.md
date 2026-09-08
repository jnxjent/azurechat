# AzureChat 本番テスト・バグ修正メモ（2026-09-08）

## 目的

AzureChatの本番テストで確認された、モデル設定、PPTX処理、Word変更履歴、生成ファイル回答後の余分な引用・内部リンク表示について、原因調査と最小修正を実施した。

このメモ作成時点では、下記コード修正はLocalテスト完了。本番環境では、デプロイ後に同一操作で最終確認する。

## 本番モデル設定の確認

本番App ServiceのApp Settingsを確認し、ユーザーがAzure Portalで次の値へ変更した。

| 環境変数 | 設定値 | 用途 |
| --- | --- | --- |
| `AZURE_OPENAI_API_DEPLOYMENT_NAME` | `gpt-5.6-luna` | 通常チャット・ツール実行 |
| `AZURE_OPENAI_PPT_DEPLOYMENT_NAME` | `gpt-5.6-terra` | PPT生成 |
| `AZURE_OPENAI_PPT_VISION_DEPLOYMENT_NAME` | `gpt-5.6-terra` | PPT Vision処理 |
| `AZURE_OPENAI_SOQL_MODEL` | `gpt-5.6-luna` | SOQL関連 |

`AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME=gpt-5.4-nano` は汎用Vision用として変更対象外とした。

本番ログでは、PPT処理の `gpt-5.6-terra` と通常ツール処理の `gpt-5.6-luna` 使用を確認済み。

## 修正1: PPTXからPNGを生成できない場合がある

### 原因・対策

LibreOffice変換時の一時プロファイル競合や、想定と異なるPDF出力名などでPNGが生成されない場合に備え、変換処理を強化した。

- LibreOfficeごとに一意な一時プロファイルを使用
- `pdf:impress_pdf_Export` を明示
- PDF出力名が想定外でも、出力が1件なら採用
- PDFが見つからない場合のstdout、stderr、出力一覧を詳細化
- 変換を1回再試行
- 縮小済みPPTXの変換に失敗した場合、元PPTXで再試行
- API側のstderrログ上限を300文字から2,000文字へ拡張

### 変更ファイル

- `src/scripts/pptx_to_png.py`
- `src/app/api/vision-review-pptx/route.ts`
- `src/tests/test_pptx_to_png.py`

### Local検証

- Windows COMによる実ファイル変換: 12/12スライドのPNG生成成功
- Linux/LibreOffice経路のモックテスト: 3件成功
- TypeScript検査成功

実際のLinux App Service上のLibreOffice動作は、本番デプロイ後にログで最終確認する。

## 修正2: AzureChat WebのWord校正で変更履歴が付かない

### 原因

文書検索からGPTが作った具体的な置換指示（A→B）を、`executeEditWord` が元の抽象的なユーザー依頼で上書きしていた。このため、Web版では具体的な修正候補が失われ、変更履歴付きの置換が実行されない場合があった。

### 対策

- ユーザーが明示したA→B指示を最優先
- ユーザー依頼が抽象的な場合、文書レビューから得た具体的なツール指示を保持
- 具体的な指示がどちらにもない場合のみ、元のユーザー依頼を使用

### 変更ファイル

- `src/features/chat-page/chat-services/chat-api/word-edit-instruction.ts`
- `src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts`
- `src/tests/test-word-edit-instruction.cjs`
- `src/tests/test_edit_word_track_changes.py`

### Local検証

- 指示選択テスト: 6件成功
- 実DOCXへの2件の置換テスト成功
- DOCX内部に `<w:del>` と `<w:ins>` が生成されることを確認
- Local E2Eログで以下を確認
  - `preserving explicit tool instruction derived from document review`
  - 具体的な修正候補4件を適用
  - `trackChanges: true`
  - 更新後Wordのポインタ保存成功

## 修正3: 生成ファイル回答後に大量の余分な引用が表示される

### 原因

Word、Excel、PDF、PPTXなどの生成・編集処理でも検索結果がツール結果に含まれ、短いファイル完了回答の後ろへ検索引用が付加されていた。

### 対策

- 成功した `downloadUrl` を持つ生成ファイル結果を判定
- 生成ファイル結果の場合は検索引用を最終回答へ付けない
- 検索結果だけ、またはエラー結果の場合は通常の引用処理を維持

### 変更ファイル

- `src/features/chat-page/chat-services/chat-api/generated-file-result.ts`
- `src/features/chat-page/chat-services/chat-api/open-ai-stream.ts`
- `src/tests/test-generated-file-result.cjs`
- `src/scripts/test-image-upload-routing.cjs`

### Local検証

- 生成ファイル判定テスト: 5件成功
- 既存の画像・編集回帰テスト: 96件成功
- Local E2Eログで次を確認

```text
[open-ai-stream] suppressed citations for generated file result
```

## 修正4: `turn0file0` などの内部参照記号が表示される

### 症状

正常なExcelダウンロードボタンとは別に、回答本文末尾へ次のような内部参照記号が表示された。

```text
fileciteturn0file0
```

### 原因

GPTがChatGPT形式の内部ファイル引用を出力する場合があるが、従来の除去処理はAzureChat独自形式の `{% citation ... /%}` のみを対象としていた。そのため、未対応の内部参照記号がそのままMarkdown画面へ渡っていた。

### 対策

- `fileciteturn...file...` を除去
- `citeturn...search...` と複数参照も除去
- サーバーの最終応答保存前に除去し、新規混入を防止
- Markdown表示時にも除去し、すでにDBへ保存された異常表示にも対応
- AzureChat独自の正常な引用は維持

### 変更ファイル

- `src/features/ui/markdown/citation-markup.ts`
- `src/features/ui/markdown/markdown.tsx`
- `src/tests/test-citation-markup.cjs`

### Local検証

- 内部引用除去テスト: 5件成功
- TypeScript検査成功
- `git diff --check` 成功
- 実操作ログで生成ファイル引用の抑止を確認

## 最終Localテスト結果

```text
Citation markup regression tests: 5 assertions passed
Generated file result regression tests: 5 assertions passed
Image upload/edit regression tests: 96 assertions passed
TypeScript check: passed
git diff --check: passed
```

PDFからExcelへの実操作でも以下を確認した。

- Azure Document Intelligenceによる4ページ処理成功
- 5シート生成成功
- `.xlsx` のBlob保存成功
- 最新Excelポインタ保存成功
- 生成ファイル回答の引用抑止成功

ログの `correction time limit reached, skipping remaining columns` は、勘定科目名の追加補正が時間制限に達した警告。Excel生成や内部リンク修正の失敗ではない。

## 本番デプロイ後の確認項目

1. PDF→Excel変換後、正常なExcelダウンロードボタンが1件表示されること
2. 回答本文に `turn0file0`、`filecite`、壊れたリンクが表示されないこと
3. 既存の異常表示を含む会話を再読み込みし、内部参照記号が非表示になること
4. SharePoint文書への通常の引用は従来どおり表示・クリックできること
5. Word校正で具体的な修正が適用され、変更履歴がWord内に残ること
6. PPTX添付処理でスライド画像が生成され、Vision処理まで完了すること

## 未対応・別件

- PPTXの表紙とコンテンツスライドの対応で、ログ上 `page[0] -> content slide[-1]` となる可能性がある件は別問題として未修正。
- ビルドがNode.js 20、App Service実行環境がNode.js 22となっている場合のネイティブモジュールABI警告は別問題として未修正。
- Linux App Service上でのLibreOffice実変換は、本番デプロイ後の確認が必要。
