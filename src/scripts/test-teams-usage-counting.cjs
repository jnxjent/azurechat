const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");
const ts = require("typescript");

const records = new Map();
const mockContainer = {
  items: {
    query() {
      return {
        async fetchAll() {
          return { resources: [] };
        },
      };
    },
    async upsert(record) {
      records.set(record.id, record);
      return { resource: record };
    },
    async create(record) {
      if (records.has(record.id)) {
        const conflict = new Error("Conflict");
        conflict.code = 409;
        throw conflict;
      }
      records.set(record.id, record);
      return { resource: record };
    },
  },
};

function loadUsageService() {
  const fileName = path.resolve("features/teams/teams-usage-service.ts");
  const source = fs.readFileSync(fileName, "utf8");
  const javascript = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  }).outputText;
  const loaded = new Module(fileName);
  loaded.filename = fileName;
  loaded.paths = Module._nodeModulePaths(path.dirname(fileName));
  loaded.require = (request) => {
    if (request === "server-only") return {};
    if (request === "@/features/common/services/cosmos") {
      return { HistoryContainer: () => mockContainer };
    }
    return Module.prototype.require.call(loaded, request);
  };
  loaded._compile(javascript, fileName);
  return loaded.exports;
}

async function main() {
  const { recordTeamsChatTurn } = loadUsageService();
  const props = {
    conversationId: "conversation-1",
    activityId: "activity-1",
    teamsUserId: "teams-user-1",
  };

  await recordTeamsChatTurn(props);
  await recordTeamsChatTurn(props);

  assert.equal(records.size, 1, "redelivered activity must be idempotent");
  const record = [...records.values()][0];
  assert.equal(record.type, "TEAMS_CHAT_TURN");
  assert.equal(record.channel, "teams");
  assert.equal(record.isDeleted, false);
  assert.match(record.id, /^teams-turn-[a-f0-9]{64}$/);
  assert.match(record.userId, /^[a-f0-9]{64}$/);
  assert.notEqual(record.userId, props.teamsUserId);
  assert.notEqual(record.conversationIdHash, props.conversationId);
  assert.notEqual(record.activityIdHash, props.activityId);

  const appSource = fs.readFileSync("features/teams/teams-app.ts", "utf8");
  assert.match(appSource, /if \(result\.type === "reply"\)/);
  assert.match(appSource, /recordCompletedTeamsTurn\(/);

  const reportSource = fs.readFileSync("cosmosdb.py", "utf8");
  assert.match(reportSource, /web_active_df\["type"\] == "CHAT_MESSAGE"/);
  assert.match(reportSource, /web_active_df\["role"\] == "user"/);
  assert.match(reportSource, /c\.type = "TEAMS_CHAT_TURN"/);
  assert.match(reportSource, /args\.container or os\.getenv\(/);
  assert.match(reportSource, /AZURE_COSMOSDB_CONTAINER_NAME/);
  assert.match(reportSource, /max_item_count=100/);
  assert.match(reportSource, /max_attempts = 3/);
  assert.match(reportSource, /choices=\["all", "web", "teams"\]/);
  assert.match(reportSource, /if args\.channel in \["all", "teams"\]/);
  assert.match(reportSource, /legacy_records=/);
  assert.match(reportSource, /display\.max_rows/);

  console.log("Teams/Web comparable chat usage tests passed.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
