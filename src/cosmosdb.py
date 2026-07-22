import argparse
import os
from datetime import datetime
from time import sleep

import pandas as pd
from azure.cosmos import CosmosClient
from dotenv import load_dotenv


def parse_args():
    parser = argparse.ArgumentParser(
        description="Aggregate Web and Teams chat usage from Cosmos DB."
    )
    parser.add_argument(
        "--container",
        help=(
            "Cosmos container to aggregate. Defaults to "
            "AZURE_COSMOSDB_CONTAINER_NAME from .env.local."
        ),
    )
    parser.add_argument(
        "--channel",
        choices=["all", "web", "teams"],
        default="all",
        help="Channel to aggregate. Defaults to all.",
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help="Show every weekly/monthly row instead of Pandas truncation.",
    )
    return parser.parse_args()


args = parse_args()
if args.full:
    pd.set_option("display.max_rows", None)
env_path = os.path.join(os.path.dirname(__file__), ".env.local")
load_dotenv(dotenv_path=env_path)

COSMOSDB_URI = os.getenv("AZURE_COSMOSDB_URI")
COSMOSDB_KEY = os.getenv("AZURE_COSMOSDB_KEY")
DATABASE_NAME = os.getenv("AZURE_COSMOSDB_DB_NAME", "chat")
CONTAINER_NAME = args.container or os.getenv(
    "AZURE_COSMOSDB_CONTAINER_NAME", "history"
)

if not COSMOSDB_URI or not COSMOSDB_KEY:
    raise RuntimeError("AZURE_COSMOSDB_URI and AZURE_COSMOSDB_KEY are required")

print(f"Cosmos target: database={DATABASE_NAME}, container={CONTAINER_NAME}")

client = CosmosClient(COSMOSDB_URI, COSMOSDB_KEY)
database = client.get_database_client(DATABASE_NAME)
container = database.get_container_client(CONTAINER_NAME)

start_date = datetime(2024, 10, 1)

# Fetch the same Web records used by the former script so legacy_records can
# be compared directly with historical reports. The data is then split into
# real chat turns and threads in memory.
web_usage_query = f'''
SELECT c.createdAt, c.userId, c.type, c.role, c.isDeleted
FROM c
WHERE NOT CONTAINS(c.userId, "j.nomoto@midac.jp")
  AND c.createdAt >= "{start_date.isoformat()}"
  AND (NOT IS_DEFINED(c.channel) OR c.channel != "teams")
'''

# Fetch Teams turns and conversations together for the same reason. A turn is
# written only after the bot reply has been sent successfully.
teams_usage_query = f'''
SELECT c.createdAt, c.userId, c.type
FROM c
WHERE c.channel = "teams"
  AND (c.type = "TEAMS_CHAT_TURN" OR c.type = "TEAMS_CHAT_THREAD")
  AND c.createdAt >= "{start_date.isoformat()}"
  AND (NOT IS_DEFINED(c.isDeleted) OR c.isDeleted = false)
'''


def fetch_dataframe(query, columns, label):
    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        try:
            print(f"Loading {label} usage (attempt {attempt}/{max_attempts})...")
            items = container.query_items(
                query=query,
                enable_cross_partition_query=True,
                max_item_count=100,
            )
            data = []
            for item in items:
                data.append(
                    {column: item.get(column) for column in columns}
                )
                if len(data) % 1000 == 0:
                    print(f"  {label}: loaded {len(data)} records")
            print(f"Loaded {label} usage: {len(data)} records")
            return pd.DataFrame(data, columns=columns)
        except Exception as error:
            if attempt == max_attempts:
                raise
            wait_seconds = 2 ** attempt
            print(
                "Cosmos query interrupted "
                f"(attempt {attempt}/{max_attempts}); "
                f"retrying in {wait_seconds}s: {error}"
            )
            sleep(wait_seconds)


def add_periods(dataframe):
    df = dataframe.copy()
    if df.empty:
        return df
    df["createdAt"] = pd.to_datetime(df["createdAt"], utc=True).dt.tz_convert(
        None
    )
    df["month"] = df["createdAt"].dt.to_period("M").apply(
        lambda period: period.start_time
    )
    df["week_start"] = df["createdAt"].dt.to_period("W").apply(
        lambda period: period.start_time
    )
    return df


def aggregate_chats(chat_dataframe, period_column):
    df = add_periods(chat_dataframe)
    if df.empty:
        return pd.DataFrame(columns=[period_column, "chats", "users"])
    return df.groupby(period_column).agg(
        chats=("createdAt", "count"),
        users=("userId", "nunique"),
    ).reset_index()


def aggregate_threads(thread_dataframe, period_column):
    df = add_periods(thread_dataframe)
    if df.empty:
        return pd.DataFrame(columns=[period_column, "threads"])
    return df.groupby(period_column).agg(
        threads=("createdAt", "count"),
    ).reset_index()


def aggregate_legacy_records(legacy_dataframe, period_column):
    df = add_periods(legacy_dataframe)
    if df.empty:
        return pd.DataFrame(columns=[period_column, "legacy_records"])
    return df.groupby(period_column).agg(
        legacy_records=("createdAt", "count"),
    ).reset_index()


def combine_summary(
    chat_dataframe,
    thread_dataframe,
    period_column,
    legacy_dataframe=None,
):
    chats = aggregate_chats(chat_dataframe, period_column)
    threads = aggregate_threads(thread_dataframe, period_column)
    summary = chats.merge(threads, on=period_column, how="outer")
    if legacy_dataframe is not None:
        legacy = aggregate_legacy_records(legacy_dataframe, period_column)
        summary = summary.merge(legacy, on=period_column, how="outer")
    if summary.empty:
        columns = [period_column, "chats", "threads"]
        if legacy_dataframe is not None:
            columns.append("legacy_records")
        columns.append("users")
        return pd.DataFrame(columns=columns)
    count_columns = ["chats", "threads", "users"]
    if legacy_dataframe is not None:
        count_columns.append("legacy_records")
    for column in count_columns:
        summary[column] = pd.to_numeric(
            summary[column], errors="coerce"
        ).fillna(0).astype(int)
    columns = [period_column, "chats", "threads"]
    if legacy_dataframe is not None:
        columns.append("legacy_records")
    columns.append("users")
    return summary[columns].sort_values(period_column)


web_usage_df = pd.DataFrame(
    columns=["createdAt", "userId", "type", "role", "isDeleted"]
)
teams_usage_df = pd.DataFrame(columns=["createdAt", "userId", "type"])

# Load Teams first because it is normally much smaller and provides immediate
# feedback even when the production Web history takes several minutes.
if args.channel in ["all", "teams"]:
    teams_usage_df = fetch_dataframe(
        teams_usage_query,
        ["createdAt", "userId", "type"],
        "Teams",
    )
if args.channel in ["all", "web"]:
    web_usage_df = fetch_dataframe(
        web_usage_query,
        ["createdAt", "userId", "type", "role", "isDeleted"],
        "Web",
    )
web_active_df = web_usage_df[
    web_usage_df["isDeleted"].isna()
    | (web_usage_df["isDeleted"] == False)
].copy()
web_chat_df = web_active_df[
    (web_active_df["type"] == "CHAT_MESSAGE")
    & (web_active_df["role"] == "user")
].copy()
web_thread_df = web_active_df[
    web_active_df["type"] == "CHAT_THREAD"
].copy()
teams_chat_df = teams_usage_df[
    teams_usage_df["type"] == "TEAMS_CHAT_TURN"
].copy()
teams_thread_df = teams_usage_df[
    teams_usage_df["type"] == "TEAMS_CHAT_THREAD"
].copy()

# Preserve the existing safeguard for old records that may contain a raw ID.
web_chat_df = web_chat_df[
    ~web_chat_df["userId"].str.contains("j.nomoto@midac.jp", na=False)
]
web_thread_df = web_thread_df[
    ~web_thread_df["userId"].str.contains("j.nomoto@midac.jp", na=False)
]

web_weekly_summary = combine_summary(
    web_chat_df, web_thread_df, "week_start", web_usage_df
)
web_monthly_summary = combine_summary(
    web_chat_df, web_thread_df, "month", web_usage_df
)
teams_weekly_summary = combine_summary(
    teams_chat_df, teams_thread_df, "week_start"
)
teams_monthly_summary = combine_summary(
    teams_chat_df, teams_thread_df, "month"
)

if args.channel in ["all", "web"]:
    print("\nWeb Weekly Summary (legacy_records matches the former count):")
    print(web_weekly_summary)
    print("\nWeb Monthly Summary (legacy_records matches the former count):")
    print(web_monthly_summary)
if args.channel in ["all", "teams"]:
    print("\nTeams Weekly Summary (chats primary, threads secondary):")
    print(teams_weekly_summary)
    print("\nTeams Monthly Summary (chats primary, threads secondary):")
    print(teams_monthly_summary)

# Keep the established CSV filenames. Each now contains comparable chat-turn
# counts plus the previous thread/conversation count as a secondary metric.
if args.channel in ["all", "web"]:
    web_weekly_summary.to_csv("weekly_summary.csv", index=False)
    web_monthly_summary.to_csv("monthly_summary.csv", index=False)
if args.channel in ["all", "teams"]:
    teams_weekly_summary.to_csv("teams_weekly_summary.csv", index=False)
    teams_monthly_summary.to_csv("teams_monthly_summary.csv", index=False)
