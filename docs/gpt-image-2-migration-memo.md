# GPT Image 2 移行メモ

## 状態

一旦棚上げ。実施時に再確認する。

## 結論

環境変数2行の変更だけでは移行できない。

- Azure上のデプロイ名は `gpt-image-2`（`gpt-image-2.0`ではない）
- `2026-04-21`はAPIバージョンではなくモデルバージョン
- 現行のAzure deployment形式を維持する場合、APIバージョンはまず `2025-04-01-preview`を使用する

## 環境変数の変更候補

```env
AZURE_OPENAI_IMAGE_DEPLOYMENT=gpt-image-2
AZURE_OPENAI_IMAGE_API_VERSION=2025-04-01-preview
AZURE_OPENAI_DALLE_API_DEPLOYMENT_NAME=gpt-image-2
AZURE_OPENAI_DALLE_API_VERSION=2025-04-01-preview
```

実際のデプロイ名がAzure Portal上の名前と完全一致することを確認する。

## 必要なコード修正

1. `src/app/api/gen-image/route.ts`と`src/features/chat-page/chat-services/chat-api/image/create.ts`で、画像生成用APIバージョンとして`AZURE_OPENAI_IMAGE_API_VERSION`を参照する。
2. `model: "gpt-image-1.5"`のハードコードを削除する。
3. model指定が必要な箇所は、画像用またはDALLE用のデプロイ環境変数を使用する。
4. `OpenAIDALLEInstance`を使うPPT表紙・PPT編集系も`gpt-image-2`へ統一する。
5. `response_format`、`reasoning_effort`、`temperature`など、`gpt-image-2`で未対応のパラメーターを送っていないか確認する。400になるパラメーターは削除する。
6. 将来的にAzure OpenAI v1 APIへ移行する場合は、deployment形式から`/openai/v1/images/generations?api-version=preview`形式への変更を別作業として扱う。

## 回帰テスト

- 通常のチャット画像生成
- 文字入れ画像生成
- PPT表紙イラスト生成
- PPT編集時の画像生成
- Base64画像レスポンスの保存
- Azure Blobへのアップロード
- APIエラー時のユーザー向けメッセージ

## セキュリティ

メモやコミットへAPIキー、ストレージキー、クライアントシークレットを記載しない。共有済みの認証情報は実施作業とは別にローテーションする。
