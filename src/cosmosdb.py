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

print(f"URI: {COSMOSDB_URI}")
print(f"KEY: {COSMOSDB_KEY[:8]}...")  # キーの先頭だけ確認用に

# Initialize Cosmos DB client
client = CosmosClient(COSMOSDB_URI, COSMOSDB_KEY)
database = client.get_database_client(DATABASE_NAME)
container = database.get_container_client(CONTAINER_NAME)

# Define start and end dates
start_date = datetime(2024, 10, 1)
end_date = datetime.utcnow()

# Query data excluding admin usage
query = f'''
SELECT c.createdAt, c.userId
FROM c
WHERE NOT CONTAINS(c.userId, "j.nomoto@midac.jp") AND c.createdAt >= "{start_date.isoformat()}"
'''

# Fetch data from Cosmos DB
items = container.query_items(query=query, enable_cross_partition_query=True)
data = [{"createdAt": item["createdAt"], "userId": item["userId"]} for item in items]

# Create a DataFrame
df = pd.DataFrame(data)

# Convert createdAt to datetime
df["createdAt"] = pd.to_datetime(df["createdAt"])

# Filter out admin user (additional safety net)
df = df[~df["userId"].str.contains("j.nomoto@midac.jp", na=False)]

# Add a "month" column for monthly aggregation
df["month"] = df["createdAt"].dt.to_period("M").apply(lambda r: r.start_time)

# Add a "week_start" column for weekly aggregation
df["week_start"] = df["createdAt"].dt.to_period("W").apply(lambda r: r.start_time)

# Weekly aggregation
weekly_summary = df.groupby("week_start").agg(
    threads=("createdAt", "count"),
    users=("userId", "nunique")
).reset_index()

# Monthly aggregation
monthly_summary = df.groupby("month").agg(
    threads=("createdAt", "count"),
    users=("userId", "nunique")
).reset_index()

# Ensure 'month' is a datetime for proper display
monthly_summary["month"] = pd.to_datetime(monthly_summary["month"])

# Print the results
print("Weekly Summary:")
print(weekly_summary)
print("\nMonthly Summary:")
print(monthly_summary)

# Save the summaries to CSV
weekly_summary.to_csv("weekly_summary.csv", index=False)
monthly_summary.to_csv("monthly_summary.csv", index=False)
