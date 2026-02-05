
import asyncio
import os
from motor.motor_asyncio import AsyncIOMotorClient
from dotenv import load_dotenv

load_dotenv()
mongo_url = os.environ.get('MONGO_URL')
db_name = os.environ.get('DB_NAME')

async def check_full_data():
    print(f"Connecting to {db_name} at {mongo_url.split('@')[1] if '@' in mongo_url else 'localhost'}...")
    client = AsyncIOMotorClient(mongo_url)
    db = client[db_name]
    
    # Check Sheets
    sheets = await db.energy_sheets.find().to_list(100)
    print(f"\n--- Energy Sheets ({len(sheets)}) ---")
    for s in sheets:
        print(f"ID: {s.get('id', s.get('_id'))}, Name: {s.get('name')}")
        if 'meters' in s:
            print(f"  Has embedded meters: {len(s['meters'])}")
        else:
            print("  No embedded meters")

    # Check Meters Collection
    meters = await db.energy_meters.find().to_list(100)
    print(f"\n--- Energy Meters Collection ({len(meters)}) ---")
    for m in meters:
        print(f"ID: {m.get('id', m.get('_id'))}, SheetID: {m.get('sheet_id')}, Name: {m.get('name')}")

    # Check Entries
    entries = await db.energy_entries.find().to_list(10)
    print(f"\n--- Energy Entries ({len(entries)}) ---")
    for e in entries:
        print(f"Date: {e.get('date')}, SheetID: {e.get('sheet_id')}, Readings: {len(e.get('readings', []))}")

if __name__ == "__main__":
    asyncio.run(check_full_data())
