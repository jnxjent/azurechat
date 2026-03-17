// src/features/common/services/key-vault.ts
import "server-only";
import { DefaultAzureCredential } from "@azure/identity";
import { SecretClient } from "@azure/keyvault-secrets";

// ---- Hot reload 対策のシングルトンキャッシュ ----
const g = globalThis as any;
let _kvClient: SecretClient | null = g.__kvClient ?? null;

// ---- KV URI を解決（よくある別名も全部見る）----
function resolveKeyVaultUri(): string {
  const name = (process.env.AZURE_KEY_VAULT_NAME || "").trim();

  const candidates = [
    process.env.AZURE_KEY_VAULT_URI,
    process.env.AZURE_KEY_VAULT_URL,
    process.env.KEY_VAULT_URI,
    process.env.KEYVAULT_URI,
    name ? `https://${name}.vault.azure.net` : "",
  ]
    .map((v) => (v || "").trim())
    .filter(Boolean)
    .map((v) => v.replace(/\/+$/, "")); // 末尾スラッシュ除去

  const uri = candidates[0] || "";
  const runtime = (global as any).EdgeRuntime ? "edge" : "node";

  if (!uri) {
    // 既存メッセージ文言は維持（依存コードがある可能性を考慮）
    throw new Error(
      "Azure Key vault is not configured correctly, check environment variables."
    );
  }

  // 最低限の形式チェック
  if (!/^https:\/\/[a-z0-9-]+\.vault\.azure\.net$/i.test(uri)) {
    throw new Error(
      `Azure Key Vault URI looks invalid: "${uri}". Expected "https://<name>.vault.azure.net"`
    );
  }

  // Edge Runtime では SDK が動かないため注意喚起（呼び出し元で runtime を nodejs に）
  if (runtime === "edge") {
    console.warn(
      '[KeyVault] Running on Edge runtime. Move the caller to a Node.js route and add `export const runtime = "nodejs";`'
    );
  }

  return uri;
}

// ---- SecretClient を取得（シングルトン）----
export const AzureKeyVaultInstance = (): SecretClient => {
  if (_kvClient) return _kvClient;

  const uri = resolveKeyVaultUri();
  const credential = new DefaultAzureCredential(); // CLI / Managed ID / SP / Env を自動解決
  _kvClient = new SecretClient(uri, credential);
  g.__kvClient = _kvClient;

  return _kvClient;
};

// （必要ならユーティリティも用意しておくと便利）
export async function getSecretValue(name: string): Promise<string> {
  const client = AzureKeyVaultInstance();
  const s = await client.getSecret(name);
  return s.value ?? "";
}
export async function setSecretValue(name: string, value: string): Promise<void> {
  const client = AzureKeyVaultInstance();
  await client.setSecret(name, value);
}
