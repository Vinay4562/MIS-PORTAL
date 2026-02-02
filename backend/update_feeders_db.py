import asyncio
import os
from pathlib import Path
from motor.motor_asyncio import AsyncIOMotorClient
from dotenv import load_dotenv
from datetime import datetime, timezone
import uuid
from pydantic import BaseModel, Field, ConfigDict

# Load env
ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

class Feeder(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    end1_name: str
    end2_name: str
    end1_import_mf: float
    end1_export_mf: float
    end2_import_mf: float
    end2_export_mf: float
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

feeders_data = [
    {"name": "400 KV Shanakrapally-Narsapur-1", "end1": "Shankarapally", "end2": "Narsapur", "mf": [0.2, 0.2, 1, 1]},
    {"name": "400 KV Shanakrapally-MHRM-2", "end1": "Shankarapally", "end2": "Maheshwaram-2", "mf": [0.2, 0.2, 1, 1]},
    {"name": "400 KV Shanakrapally-MHRM-1", "end1": "Shankarapally", "end2": "Maheshwaram-1", "mf": [0.2, 0.2, 1, 1]},
    {"name": "220 KV Sadasivapet-1", "end1": "Shankarapally", "end2": "Sadasivapet", "mf": [1, 1, 1.6, 1.6]},
    {"name": "220 KV Sadasivapet-2", "end1": "Shankarapally", "end2": "Sadasivapet", "mf": [1, 1, 1.6, 1.6]},
    {"name": "220 KV Parigi-1", "end1": "Shankarapally", "end2": "Parigi", "mf": [1, 1, 1, 1]},
    {"name": "220 KV Parigi-2", "end1": "Shankarapally", "end2": "Parigi", "mf": [1, 1, 1, 1]},
    {"name": "220 KV Tandur", "end1": "Shankarapally", "end2": "Tandur", "mf": [1, 1, 1, 1]},
    {"name": "220 KV Gachibowli-1", "end1": "Shankarapally", "end2": "Gachibowli", "mf": [2, 2, 2, 2]},
    {"name": "220 KV Gachibowli-2", "end1": "Shankarapally", "end2": "Gachibowli", "mf": [2, 2, 2, 2]},
    {"name": "220 KV KethiReddyPally", "end1": "Shankarapally", "end2": "KethiReddyPally", "mf": [1, 1, 1, 1]},
    {"name": "400 KV KethiReddyPally-1", "end1": "Shankarapally", "end2": "KethiReddyPally", "mf": [1, 1, 1, 1]},
    {"name": "220 KV Yeddumailaram-1", "end1": "Shankarapally", "end2": "Yeddumailaram", "mf": [1, 1, 1, 1]},
    {"name": "220 KV Yeddumailaram-2", "end1": "Shankarapally", "end2": "Yeddumailaram", "mf": [1, 1, 1, 1]},
    {"name": "400 KV Shanakrapally-Narsapur-2", "end1": "Shankarapally", "end2": "Narsapur", "mf": [0.2, 0.2, 1, 1]},
    {"name": "400 KV Nizamabad-1&2", "end1": "Shankarapally", "end2": "Nizamabad", "mf": [1000, 1000, 1000, 1000]},
    {"name": "400 KV KethiReddyPally-2", "end1": "Shankarapally", "end2": "KethiReddyPally", "mf": [1, 1, 1, 1]}
]

async def update_feeders():
    print("Connecting to DB...")
    # Verify connection
    try:
        await client.admin.command('ping')
        print("Connected.")
    except Exception as e:
        print(f"Failed to connect: {e}")
        return

    for fd in feeders_data:
        existing = await db.feeders.find_one({"name": fd['name']})
        
        if existing:
            # Update MFs
            update_fields = {
                "end1_import_mf": fd['mf'][0],
                "end1_export_mf": fd['mf'][1],
                "end2_import_mf": fd['mf'][2],
                "end2_export_mf": fd['mf'][3]
            }
            await db.feeders.update_one({"_id": existing['_id']}, {"$set": update_fields})
            print(f"Updated: {fd['name']}")
        else:
            # Insert new
            feeder_obj = Feeder(
                name=fd['name'],
                end1_name=fd['end1'],
                end2_name=fd['end2'],
                end1_import_mf=fd['mf'][0],
                end1_export_mf=fd['mf'][1],
                end2_import_mf=fd['mf'][2],
                end2_export_mf=fd['mf'][3]
            )
            doc = feeder_obj.model_dump()
            doc['created_at'] = doc['created_at'].isoformat()
            await db.feeders.insert_one(doc)
            print(f"Created: {fd['name']}")

    print("Done.")
    client.close()

if __name__ == "__main__":
    asyncio.run(update_feeders())
