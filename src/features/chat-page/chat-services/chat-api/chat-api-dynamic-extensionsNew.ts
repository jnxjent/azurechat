// src/features/chat-page/chat-services/chat-api/chat-api-dynamic-extensions.ts
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

    // secure headers
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

    // user id header（拡張がユーザー文脈を期待するため）
    headerItems.push({
      id: "authorization",
      key: "authorization",
      value: await userHashedId(),
    });

    // ヘッダー辞書化
    const headers: Record<string, string> = headerItems.reduce(
      (acc, h) => ((acc[h.key] = h.value), acc),
      {} as Record<string, string>
    );

    // Content-Type は必ず JSON
    if (!headers["content-type"]) {
      headers["content-type"] = "application/json";
    }

    // query 文字列置換（必要なら）
    if (args?.query) {
      for (const key in args.query) {
        const val = args.query[key];
        functionModel.endpoint = functionModel.endpoint.replace(`${key}`, val);
      }
    }

    // ★ Body が空でも必ず JSON を送る（デフォルトペイロード）
    const defaultBody = { query: "*", top: 3 };
    const payload = args?.body ?? defaultBody;

    const requestInit: RequestInit = {
      method: functionModel.endpointType,
      headers,
      cache: "no-store",
      body: JSON.stringify(payload),
    };

    const resp = await fetch(functionModel.endpoint, requestInit);

    // JSONでない場合のボディ先頭をログ（あなたのログに出る）
    const text = await resp.text();
    const ct = resp.headers.get("content-type") ?? "";
    if (!ct.includes("application/json")) {
      console.error(
        `🔴 executeFunction error: Non-JSON response. Content-Type="${ct}". BodySnippet=${text.slice(
          0,
          200
        )}`
      );
      return `There was an error calling the api: Non-JSON response.`;
    }

    if (!resp.ok) {
      return `There was an error calling the api: ${resp.status} ${resp.statusText}`;
    }

    const data = JSON.parse(text);
    return { id: functionModel.id, result: data };
  } catch (e) {
    console.error("🔴 executeFunction", e);
    return `There was an error calling the api: ${e}`;
  }
}
