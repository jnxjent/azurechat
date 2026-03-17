"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { userHashedId } from "@/features/auth-page/helpers";
import {
  FindAllExtensionForCurrentUser,
  FindSecureHeaderValue,
} from "@/features/extensions-page/extension-services/extension-service";
import {
  ExtensionFunctionModel,
  ExtensionModel,
} from "@/features/extensions-page/extension-services/models";
import { RunnableToolFunction } from "openai/lib/RunnableFunction";
import { ToolsInterface } from "../models";

/** --- ユーティリティ --- */
function looksJsonContentType(ct?: string | null) {
  if (!ct) return false;
  const lower = ct.toLowerCase();
  return lower.includes("application/json") || lower.endsWith("+json");
}

async function parseJsonSafe(res: Response): Promise<any> {
  const ct = res.headers.get("content-type");
  const raw = await res.text(); // 先にテキストで読む（デバッグしやすさ優先）
  if (looksJsonContentType(ct)) {
    try {
      return JSON.parse(raw);
    } catch (e) {
      // JSON宣言だが壊れている
      throw new Error(
        `Failed to parse JSON (declared as JSON). ParseError=${(e as Error).message}. BodySnippet=${raw.slice(0, 500)}`
      );
    }
  }
  // JSON以外（HTMLなど）
  throw new Error(
    `Non-JSON response. Content-Type="${ct ?? "unknown"}". BodySnippet=${raw
      .replace(/\s+/g, " ")
      .slice(0, 500)}`
  );
}

export const GetDynamicExtensions = async (props: {
  extensionIds: string[];
}): Promise<ServerActionResponse<Array<any>>> => {
  const extensionResponse = await FindAllExtensionForCurrentUser();

  if (extensionResponse.status === "OK") {
    const extensionToReturn = extensionResponse.response.filter((e) =>
      props.extensionIds.includes(e.id)
    );

    const dynamicExtensions: Array<RunnableToolFunction<any>> = [];

    extensionToReturn.forEach((e) => {
      e.functions.forEach((f) => {
        const extension = JSON.parse(f.code) as ToolsInterface;
        dynamicExtensions.push({
          type: "function",
          function: {
            function: (args: any) =>
              executeFunction({
                functionModel: f,
                extensionModel: e,
                args,
              }),
            parse: JSON.parse,
            parameters: extension.parameters,
            description: extension.description,
            name: extension.name,
          },
        });
      });
    });

    return { status: "OK", response: dynamicExtensions };
  }

  return extensionResponse;
};

async function executeFunction(props: {
  functionModel: ExtensionFunctionModel;
  extensionModel: ExtensionModel;
  args: any;
}) {
  try {
    const { functionModel, args, extensionModel } = props;

    // 1) セキュアヘッダ解決
    const headerItems = await Promise.all(
      extensionModel.headers.map(async (h) => {
        const hv = await FindSecureHeaderValue(h.id);
        return {
          id: h.id,
          key: h.key,
          value: hv.status === "OK" ? hv.response : "***",
        };
      })
    );

    // 2) user id を auth として付与（アプリ仕様）
    headerItems.push({
      id: "authorization",
      key: "authorization",
      value: await userHashedId(),
    });

    // 3) ヘッダ辞書化
    const headers: Record<string, string> = headerItems.reduce(
      (acc, h) => ((acc[h.key] = h.value), acc),
      {} as Record<string, string>
    );

    // 4) クエリ置換（エンドポイント文字列をコピーしてから置換）
    let endpoint = functionModel.endpoint;
    if (args?.query && typeof args.query === "object") {
      for (const key of Object.keys(args.query)) {
        // {city} のようなテンプレートを想定： "…?q={city}" -> 値に置換
        const val = String(args.query[key] ?? "");
        const safe = encodeURIComponent(val);
        endpoint = endpoint.replace(new RegExp(`{${key}}`, "g"), safe);
        // 後方互換：単純な key マッチにも対応（既存実装踏襲）
        endpoint = endpoint.replace(new RegExp(`${key}`, "g"), safe);
      }
    }

    // 5) リクエスト構築
    const requestInit: RequestInit = {
      method: functionModel.endpointType,
      headers,
      cache: "no-store",
    };
    if (args?.body) {
      requestInit.body = JSON.stringify(args.body);
      // JSON 送信であることを明示（既に付いているなら上書きしない）
      if (!Object.keys(headers).some((k) => k.toLowerCase() === "content-type")) {
        (requestInit.headers as Record<string, string>)["content-type"] =
          "application/json";
      }
    }

    // 6) 呼び出し
    const response = await fetch(endpoint, requestInit);

    // 7) ステータスエラーは本文をスニペットで返す（HTMLでも可視化）
    if (!response.ok) {
      const ct = response.headers.get("content-type");
      const body = await response.text();
      const hint =
        looksJsonContentType(ct) && body.trim().startsWith("{")
          ? (() => {
              try {
                const j = JSON.parse(body);
                return JSON.stringify(j).slice(0, 500);
              } catch {
                return body.slice(0, 500);
              }
            })()
          : body.slice(0, 500);
      return `There was an error calling the api: ${response.status} ${response.statusText}. URL=${endpoint}. Snippet=${hint}`;
    }

    // 8) JSON以外の本文（HTMLなど）に対する保護：JSONとして読めなければ詳細を投げる
    const result = await parseJsonSafe(response);

    return {
      id: functionModel.id,
      result,
    };
  } catch (e: any) {
    // 9) ランタイム例外（<!DOCTYPE…>含む）を安全に文字列化
    const msg = e?.message || String(e);
    console.error("🔴 executeFunction error:", msg);
    return `There was an error calling the api: ${msg}`;
  }
}
