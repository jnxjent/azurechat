import os
from datetime import datetime
from azure.cosmos import CosmosClient
import pandas as pd
from dotenv import load_dotenv

# .env.local のフルパスを指定
env_path = os.path.join(os.path.dirname(__file__), ".env.local")
load_dotenv(dotenv_path=env_path)

# 環境変数の取得
COSMOSDB_URI = os.getenv("AZURE_COSMOSDB_URI")
COSMOSDB_KEY = os.getenv("AZURE_COSMOSDB_KEY")
DATABASE_NAME = "chat"
CONTAINER_NAME ="history"

if not COSMOSDB_URI or not COSMOSDB_KEY:
    raise RuntimeError("AZURE_COSMOSDB_URI and AZURE_COSMOSDB_KEY are required")

# Initialize Cosmos DB client
client = CosmosClient(COSMOSDB_URI, COSMOSDB_KEY)
database = client.get_database_client(DATABASE_NAME)
container = database.get_container_client(CONTAINER_NAME)

# Define start and end dates
start_date = datetime(2024, 10, 1)
end_date = datetime.utcnow()

# Query existing Web data. Records without a channel are legacy Web records.
# The original counting method is intentionally preserved.
web_query = f'''
SELECT c.createdAt, c.userId
FROM c
WHERE NOT CONTAINS(c.userId, "j.nomoto@midac.jp")
  AND c.createdAt >= "{start_date.isoformat()}"
  AND (NOT IS_DEFINED(c.channel) OR c.channel != "teams")
'''

# Query Teams thread records written by the Teams module.
teams_query = f'''
SELECT c.createdAt, c.userId
FROM c
WHERE c.channel = "teams"
  AND c.type = "TEAMS_CHAT_THREAD"
  AND c.createdAt >= "{start_date.isoformat()}"
'''


def fetch_dataframe(query):
    items = container.query_items(
        query=query,
        enable_cross_partition_query=True,
    )
    data = [
        {"createdAt": item["createdAt"], "userId": item["userId"]}
        for item in items
    ]
    return pd.DataFrame(data, columns=["createdAt", "userId"])


def aggregate(dataframe):
    df = dataframe.copy()
    if df.empty:
        weekly = pd.DataFrame(columns=["week_start", "threads", "users"])
        monthly = pd.DataFrame(columns=["month", "threads", "users"])
        return weekly, monthly

    df["createdAt"] = pd.to_datetime(df["createdAt"])
    df["month"] = df["createdAt"].dt.to_period("M").apply(
        lambda period: period.start_time
    )
    df["week_start"] = df["createdAt"].dt.to_period("W").apply(
        lambda period: period.start_time
    )

    weekly = df.groupby("week_start").agg(
        threads=("createdAt", "count"),
        users=("userId", "nunique"),
    ).reset_index()

    monthly = df.groupby("month").agg(
        threads=("createdAt", "count"),
        users=("userId", "nunique"),
    ).reset_index()
    monthly["month"] = pd.to_datetime(monthly["month"])
    return weekly, monthly


web_df = fetch_dataframe(web_query)
teams_df = fetch_dataframe(teams_query)

# Additional safety net retained for the existing Web statistics.
web_df = web_df[~web_df["userId"].str.contains("j.nomoto@midac.jp", na=False)]

web_weekly_summary, web_monthly_summary = aggregate(web_df)
teams_weekly_summary, teams_monthly_summary = aggregate(teams_df)

# Print the results
print("Web Weekly Summary:")
print(web_weekly_summary)
print("\nWeb Monthly Summary:")
print(web_monthly_summary)
print("\nTeams Weekly Summary:")
print(teams_weekly_summary)
print("\nTeams Monthly Summary:")
print(teams_monthly_summary)

# Keep the existing Web CSV names unchanged and write Teams separately.
web_weekly_summary.to_csv("weekly_summary.csv", index=False)
web_monthly_summary.to_csv("monthly_summary.csv", index=False)
teams_weekly_summary.to_csv("teams_weekly_summary.csv", index=False)
teams_monthly_summary.to_csv("teams_monthly_summary.csv", index=False)
