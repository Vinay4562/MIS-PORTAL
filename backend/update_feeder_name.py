import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

mongo_url = os.environ['MONGO_URL']
client = MongoClient(mongo_url)
db = client[os.environ['DB_NAME']]

# Update max_min_feeders
result = db.max_min_feeders.update_one(
    {"name": "Bus + Station"},
    {"$set": {"name": "Bus Voltages & Station Load"}}
)

print(f"Matched: {result.matched_count}, Modified: {result.modified_count}")
