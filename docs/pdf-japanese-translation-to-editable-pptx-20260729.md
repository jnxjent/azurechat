# PDF日本語多言語翻訳 → 編集可能PPTX機能 実装メモ

更新日: 2026-07-29

## 1. 目的

ユーザーがPDFを添付し、次のように依頼した場合に、日本語部分だけを指定言語へ翻訳した編集可能なPowerPointを出力する。

> 添付ファイルの日本語部分のみを英訳して。その他、絵はそのままで。編集可能なPPTで出力して。

英語以外の場合は、例えば次のように依頼する。

> 添付ファイルの日本語部分のみをベトナム語に翻訳して。その他、絵はそのままで。編集可能なPPTで出力して。

想定用途は、文字レイヤーを持たないパンフレットやチラシを多言語化し、出力後に利用者が翻訳文、フォントサイズ、改行、位置などをPowerPoint上で手修正するケース。

## 2. 確定仕様

- ツール名: `translate_pdf_to_pptx`
- 入力: PDF
- 出力: PowerPoint（`.pptx`）
- 翻訳先: 英語、ポルトガル語、ベトナム語、インドネシア語、中国語（簡体字）、韓国語、スペイン語、タガログ語
- 翻訳先コード: `en`、`pt`、`vi`、`id`、`zh-CN`、`ko`、`es`、`fil`
- 翻訳先を省略した場合は英語（`en`）として処理
- PDFの1ページをPowerPointの1スライドとして出力
- 元の写真、イラスト、罫線、背景、配色はページ背景として維持
- OCRで検出した日本語を指定言語へ翻訳
- 翻訳文は編集可能なPowerPointテキストボックスとして配置
- 日本語以外の文字は原則として翻訳対象外
- 縦長の日本語領域は、翻訳文を回転テキストとして配置
- 最小フォントサイズ: **7pt**
- 最大フォントサイズ: 30pt
- 長文はPowerPointの自動縮小を使用し、テキストボックス内へ収める
- 現在の処理上限: 10ページ
- 生成したPPTXをスレッドの最新PPTXとして保存し、その後の `edit_pptx` で編集可能

## 3. 処理フロー

1. スレッド内の添付PDF、またはSharePoint/SLで指定されたPDFを解決する。
2. `/api/edit-pptx` を `action: "translate_pdf_to_pptx"` で呼び出す。
3. Azure Document Intelligenceの `prebuilt-read` で文字と座標を取得する。
4. 日本語を含む行だけを抽出する。
5. 抽出行をAzure OpenAIで日本語から指定言語へ翻訳する。
6. PyMuPDFで各PDFページを画像化する。
7. 元の日本語領域を周辺から推定した背景色で消去する。
8. ページ画像をスライド背景として配置する。
9. 元の座標へ、透明で編集可能な翻訳文テキストボックスを配置する。
10. PPTXをBlob Storageへ保存し、UIにダウンロードボタンを表示する。

翻訳APIが一部の行を返さなかった場合は、欠落行だけを指定言語で自動的に再翻訳する。

出力ファイル名は、チャット表示とダウンロード後の両方で翻訳元・翻訳先を判別できる名前にする。

- 表示名の例: `panhu_ポルトガル語版.pptx`
- Blob物理名の例: `panhu_ポルトガル語版_<一意ID>.pptx`
- ブラウザーが `Content-Disposition` の表示名を使用する場合は、一意IDなしの表示名で保存される
- ブラウザーがヘッダーを使用せずURL末尾をファイル名にする場合も、物理名に元ファイル名と言語名が含まれる
- この物理名変更は `translate_pdf_to_pptx` の出力だけに限定し、通常のPPT生成・編集には適用しない

同じスレッドで翻訳先だけを変更した場合は、直前に生成したPPTXを編集・再翻訳せず、スレッド内の元PDFから新しい言語版を生成する。例えば、中国語版の生成後に「添付をポルトガル語に変換して」と依頼すると、同じ元PDFを `targetLanguage: "pt"` で再処理する。

LLMがこの依頼を誤って `create_pptx` または `edit_pptx` に振り分けた場合も、実行時ガードが次の条件をすべて満たす場合だけ `translate_pdf_to_pptx` へ処理を戻す。

- 直近PPTXの表示名が本機能の翻訳版命名（`英訳版`、`中国語版`、`ポルトガル語版`など）である
- 依頼に対応言語と、翻訳先変更を示す表現がある
- スレッド内に元PDFが残っている
- 新規PPT作成や、タイトル・特定ページなど一部分だけの編集依頼ではない

この限定条件により、通常の `create_pptx` と `edit_pptx` の処理には介入しない。

## 4. 変更ファイル

### 新規

- `src/scripts/pdf_translate_to_pptx.py`
  - OCR、指定言語への翻訳、ページ画像化、日本語領域の消去、PPTX生成を担当する。
  - CLI引数 `--target-language` を受け取り、省略時は英語を使用する。
  - 中国語にはMicrosoft YaHei、韓国語にはMalgun Gothic、その他にはArialを指定する。

### 変更

- `src/app/api/edit-pptx/route.ts`
  - `translate_pdf_to_pptx` アクションを追加。
  - `targetLanguage` の検証とPythonスクリプトへの引き渡しを追加。
  - Pythonスクリプトの実行、PPTXのBlob保存、最新PPTXポインター保存を追加。

- `src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts`
  - `translate_pdf_to_pptx` ツール定義と実行処理を追加。
  - `targetLanguage` の8言語指定を追加。
  - `fileUrl` 省略時はスレッド内の最新PDFを使用。
  - `fileQuery` 指定時はSharePoint/SLのPDFを解決。

- `src/features/chat-page/chat-services/chat-api/chat-api-extension.ts`
  - 「日本語部分のみ英訳」「ベトナム語に差し替え」「絵はそのまま」「編集可能なPPT」などの依頼を新ツールへ振り分けるルールを追加。

## 5. 使用する既存環境

新しい環境変数やパッケージは追加していない。既存の以下を使用する。

- `AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT`
- `AZURE_DOCUMENT_INTELLIGENCE_KEY`
- `AZURE_OPENAI_API_KEY`
- `AZURE_OPENAI_API_INSTANCE_NAME`
- `AZURE_OPENAI_API_DEPLOYMENT_NAME`
- `AZURE_OPENAI_API_VERSION`
- `AZURE_STORAGE_ACCOUNT_NAME`
- `AZURE_STORAGE_ACCOUNT_KEY`
- PyMuPDF
- Pillow
- python-pptx
- Azure Document Intelligence SDK

通常のAzure OpenAI設定がない場合、翻訳処理では既存のVision用Azure OpenAI設定をフォールバックとして利用する。

## 6. 検証内容

検証ファイル:

- `panhu.pdf`
- 全2ページ
- Adobe Illustrator出力
- PDF内に抽出可能な文字レイヤーなし

英語版の初期検証結果:

- Azure Document IntelligenceによるOCR成功
- 日本語195行を検出・英訳
- 1ページ目に編集可能テキスト105個を生成
- 2ページ目に編集可能テキスト90個を生成
- 編集可能テキスト内に日本語が残っていないことを確認
- スライドサイズを元PDFと同じA3横相当に設定
- PowerPointで画像化し、写真・イラスト・配色と英文配置を目視確認
- ESLint成功
- Next.js本番ビルド成功
- Python構文チェック成功

フォントサイズは実出力の確認を経て、当初の5ptから6.5ptへ拡大し、最終的に最小7ptへ確定した。

多言語化後の検証結果:

- 同じ `panhu.pdf` をベトナム語（`vi`）と中国語・簡体字（`zh-CN`）へ実変換
- 各言語とも全2ページ、日本語195行を翻訳し、編集可能テキスト195個を生成
- ベトナム語版のテキストボックス内に日本語文字が残っていないことを確認
- 中国語版のテキストボックス内にひらがな・カタカナが残っていないことを確認
- ベトナム語版でArial、中国語版でMicrosoft YaHeiがラテン文字・東アジア文字の両方に設定されていることを確認
- PowerPointで両言語版を画像化し、写真・イラスト・配色が維持され、翻訳文が描画されることを確認
- `targetLanguage` 省略時に英語へ変換される後方互換を再確認
- 同一スレッド内の翻訳先変更を `create_pptx` / `edit_pptx` ではなく元PDFの再翻訳へ戻すガードを追加
- ESLint、Next.js本番ビルド、Python構文チェック成功

## 7. Localでの確認方法

```powershell
cd C:\Users\021213\azurechat-office-work\src
npm run dev
```

ブラウザーで `http://localhost:3000` を開き、PDFを添付して次のように依頼する。

> 添付ファイルの日本語部分のみを英訳して。その他、絵はそのままで。編集可能なPPTで出力して。

確認項目:

- `.pptx` のダウンロードボタンが表示されること
- PDFページ数とスライド数が一致すること
- 日本語部分が指定言語へ置き換わっていること
- 写真とイラストが維持されていること
- 翻訳文をPowerPoint上で選択・編集できること
- 本文の基本フォントが7pt以上になっていること
- 長文が枠外へ大きくはみ出していないこと

開発サーバー起動中にPythonスクリプトだけを変更した場合、通常はサーバー再起動不要。ブラウザーを更新し、新しく変換したPPTXで確認する。既に生成済みのPPTXには変更は反映されない。

## 8. 制約・注意点

- 元PDFの絵や文字はスライド背景画像になるため、背景側のイラストはPowerPoint上で個別編集できない。
- 編集可能なのは、新たに配置した翻訳文テキストボックス。
- OCR単位が行単位のため、複雑な段組みでは翻訳文の改行や位置を手修正する場合がある。
- 英語やラテン文字系言語は日本語より長くなる場合があり、密集したレイアウトでは自動縮小される。
- 中国語は簡体字を標準とする。繁体字は現在の対象外。
- PowerPointを開く環境に指定フォントがない場合は、PowerPoint側で代替フォントが使用される。
- 写真や複雑な模様の上に日本語がある場合、背景色による消去跡が出る可能性がある。
- 小さな文字やイラスト内の装飾文字は、OCR結果によって位置精度が下がる可能性がある。
- 異なる用紙サイズが混在するPDFは未最適化。現在は先頭ページのサイズをPowerPoint全体へ使用する。
- PDFデータはAzure Document Intelligenceへ送信され、抽出した日本語テキストはAzure OpenAIへ送信される。
- 最大10ページを超えるPDFは現在エラーにする。

## 9. 既存機能への影響

既存機能への影響を限定するため、処理は専用ツール・専用アクション・専用Pythonスクリプトとして追加している。

以下の既存処理本体は変更していない。

- 通常のPPT作成
- 既存PPT編集
- PDFからWordへの変換
- PDF/WordからExcelへの変換
- Word編集
- Excel編集
- 画像生成・画像編集

既存APIの共通部分では、命令文なしで実行可能なアクションの許可リストに `translate_pdf_to_pptx` を追加しただけであり、ほかのアクションの条件は維持している。

## 10. 現在の状態

- Local実装完了
- サンプルPDFでの実処理確認完了
- 8言語の選択処理を実装完了
- 英語、ベトナム語、中国語（簡体字）の実変換確認完了
- フォント最小7ptへの調整完了
- 本機能はひとまず完成扱い
- 本番環境へのデプロイは、この実装作業には含めていない
