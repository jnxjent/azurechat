const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");
const ts = require("typescript");

function loadTypeScriptModule(relativePath, mocks = {}) {
  const fileName = path.resolve(relativePath);
  const source = fs.readFileSync(fileName, "utf8");
  const javascript = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true,
    },
  }).outputText;
  const loaded = new Module(fileName);
  loaded.filename = fileName;
  loaded.paths = Module._nodeModulePaths(path.dirname(fileName));
  const originalLoad = Module._load;
  Module._load = function (request, parent, isMain) {
    if (Object.prototype.hasOwnProperty.call(mocks, request)) {
      return mocks[request];
    }
    return originalLoad.call(this, request, parent, isMain);
  };
  try {
    loaded._compile(javascript, fileName);
  } finally {
    Module._load = originalLoad;
  }
  return loaded.exports;
}

const searchTool = loadTypeScriptModule("features/teams/teams-search-tool.ts");

assert.deepEqual(
  searchTool.parseTeamsInternalSearchToolArguments(
    JSON.stringify({
      query: "各候補地 メリット デメリット",
      scope: "dept_common",
      folder: "可能性調査",
    })
  ),
  {
    query: "各候補地 メリット デメリット",
    scope: "dept_common",
    folder: "可能性調査",
  }
);
assert.throws(
  () =>
    searchTool.parseTeamsInternalSearchToolArguments(
      JSON.stringify({ query: "候補地", scope: "unknown" })
    ),
  /Unsupported Teams internal search scope/
);
assert.match(
  searchTool.TEAMS_INTERNAL_SEARCH_INSTRUCTIONS,
  /「『可能性調査』のフォルダーで」/
);
assert.equal(
  searchTool.TEAMS_INTERNAL_SEARCH_TOOL.function.parameters.additionalProperties,
  false
);

const completionQueue = [];
const searchCalls = [];
const completionCalls = [];
const openai = {
  chat: {
    completions: {
      create: async (request) => {
        completionCalls.push(request);
        const next = completionQueue.shift();
        if (!next) throw new Error("Unexpected completion call");
        return next;
      },
    },
  },
};

const chatService = loadTypeScriptModule("features/teams/teams-chat-service.ts", {
  "server-only": {},
  "@/features/common/services/openai": {
    OpenAIInstance: () => openai,
  },
  "./teams-search-service": {
    searchTeamsKnowledgeWithTarget: async (args) => {
      searchCalls.push(args);
      return {
        context: "[1] 候補地資料\nA候補地のメリットとデメリット",
        sources: [
          {
            index: 1,
            name: "候補地資料.docx",
            url: "https://example.test/candidate.docx",
            kind: "internal",
          },
        ],
        userEmail: args.userEmail || "",
        dept: "bm",
      };
    },
  },
  "./teams-brave-search-service": {
    resolveBraveSearchRequest: (message) => ({
      enabled: false,
      query: message,
      skipInternalSearch: false,
    }),
    searchBraveWeb: async () => ({ context: "", sources: [] }),
  },
  "./teams-instant-reply": {
    resolveTeamsInstantReply: () => null,
  },
  "./teams-web-query": {
    buildTeamsWebQuery: (query) => ({ query, enriched: false }),
  },
  "./teams-search-tool": searchTool,
});

const slSearchTarget = loadTypeScriptModule("lib/sl-search-target.ts");
const vectorSearchCalls = [];
const searchService = loadTypeScriptModule(
  "features/teams/teams-search-service.ts",
  {
    "server-only": {},
    "@/features/chat-page/chat-services/azure-ai-search/azure-ai-search": {
      SimpleSearch: async () => ({ status: "OK", response: [] }),
      ExtensionSimilaritySearch: async (args) => {
        vectorSearchCalls.push(args);
        return {
          status: "OK",
          response: [
            {
              document: {
                metadata: "候補地資料.docx",
                pageContent: "A候補地のメリットとデメリット",
                fileUrl: "https://example.test/candidate.docx",
              },
            },
          ],
        };
      },
    },
    "@/features/auth-page/helpers": {
      hashValue: (value) => `hash:${value}`,
    },
    "@/lib/sl-dept": {
      resolveSlAccess: () => ({ dept: "bm", role: "dept_admin" }),
    },
    "@/lib/sl-search-target": slSearchTarget,
  }
);

process.env.AZURE_OPENAI_API_KEY = "test-key";
process.env.AZURE_OPENAI_API_INSTANCE_NAME = "test-instance";
process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME = "test-deployment";
process.env.AZURE_OPENAI_API_EMBEDDINGS_DEPLOYMENT_NAME = "test-embeddings";
process.env.AZURE_OPENAI_API_VERSION = "test-version";
process.env.AZURE_SEARCH_API_KEY = "test-search-key";
process.env.AZURE_SEARCH_NAME = "test-search-name";
process.env.AZURE_SEARCH_INDEX_NAME = "test-search-index";

async function run() {
  completionQueue.push(
    {
      choices: [
        {
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call-folder-search",
                type: "function",
                function: {
                  name: searchTool.TEAMS_INTERNAL_SEARCH_TOOL_NAME,
                  arguments: JSON.stringify({
                    query: "各候補地 メリット デメリット",
                    scope: "dept_common",
                    folder: "可能性調査",
                  }),
                },
              },
            ],
          },
        },
      ],
    },
    {
      choices: [
        {
          message: {
            role: "assistant",
            content: "比較結果です。[1]",
          },
        },
      ],
    }
  );

  const scopedReply = await chatService.createTeamsChatReply({
    conversationId: "folder-variation-test",
    message:
      "部署共通フォルダーの「可能性調査」のフォルダーで、各候補地のメリットデメリットを比較表にして",
    userEmail: "USER@example.com",
  });
  assert.equal(scopedReply.type, "reply");
  assert.match(scopedReply.text, /比較結果です。\[1\]/);
  assert.match(scopedReply.text, /候補地資料\.docx/);
  assert.deepEqual(searchCalls[0], {
    query: "各候補地 メリット デメリット",
    scope: "dept_common",
    folder: "可能性調査",
    userEmail: "USER@example.com",
  });
  assert.equal(completionCalls[0].tool_choice, "auto");
  assert.equal(completionCalls[0].tools[0].function.name, "teams_internal_search");
  assert.equal(
    completionCalls[1].messages.at(-1).tool_call_id,
    "call-folder-search"
  );

  completionQueue.push({
    choices: [
      {
        message: {
          role: "assistant",
          content: "東京は日本の首都です。",
        },
      },
    ],
  });
  const generalReply = await chatService.createTeamsChatReply({
    conversationId: "general-question-test",
    message: "日本の首都は？",
    userEmail: "USER@example.com",
  });
  assert.equal(generalReply.text, "東京は日本の首都です。");
  assert.equal(searchCalls.length, 1);

  await searchService.searchTeamsKnowledgeWithTarget({
    query: "各候補地 メリット デメリット",
    scope: "dept_common",
    folder: "可能性調査",
    userEmail: "user@example.com",
  });
  assert.equal(vectorSearchCalls.length, 1);
  assert.equal(vectorSearchCalls[0].searchText, "各候補地 メリット デメリット");
  assert.match(vectorSearchCalls[0].filter, /slScope eq 'dept_common'/);
  assert.match(
    vectorSearchCalls[0].filter,
    /search\.ismatch\('可能性調査', 'relativePath'/
  );
  assert.equal(vectorSearchCalls[0].deptLower, "bm");
  assert.equal(vectorSearchCalls[0].userHash, "hash:user@example.com");

  console.log("Teams LLM search tool tests passed.");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
