import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List

import requests
from dotenv import load_dotenv
from pymongo import MongoClient

from server import (
    _parse_whatsapp_messages,
    _classify_interruption_message,
    _extract_time_from_text,
    _build_feeder_aliases,
    InterruptionEntry,
)


BASE_URL = os.environ.get("IMPORT_BASE_URL", "http://127.0.0.1:8000/api")
ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / ".env")
mongo_url = os.environ.get("MONGO_URL")
db_name = os.environ.get("DB_NAME")
mongo_client = MongoClient(mongo_url) if mongo_url and db_name else None
db = mongo_client[db_name] if mongo_client else None


def main():
    email = f"whatsapp_import_{datetime.now().strftime('%Y%m%d%H%M%S')}@example.com"
    password = "TestPass123!"
    full_name = "WhatsApp Import"

    print(f"Registering user {email}")
    reg_resp = requests.post(
        f"{BASE_URL}/auth/register",
        json={"email": email, "password": password, "full_name": full_name},
    )
    reg_resp.raise_for_status()
    reg_data = reg_resp.json()
    token = reg_data["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    try:
        init_resp = requests.post(f"{BASE_URL}/init-feeders", headers=headers)
        if init_resp.status_code == 200:
            print("Initialized feeders (or already initialized)")
        else:
            print(f"Init feeders returned status {init_resp.status_code}: {init_resp.text}")
    except Exception as e:
        print(f"Warning: init-feeders call failed: {e}")

    try:
        mm_init = requests.post(f"{BASE_URL}/max-min/init", headers=headers)
        if mm_init.status_code == 200:
            print("Max-Min feeders initialized or already present")
        else:
            print(f"Max-Min init returned status {mm_init.status_code}: {mm_init.text}")
    except Exception as e:
        print(f"Warning: max-min init failed: {e}")

    feeders_resp = requests.get(f"{BASE_URL}/max-min/feeders", headers=headers)
    feeders_resp.raise_for_status()
    all_feeders: List[Dict[str, Any]] = feeders_resp.json()
    feeders = [
        f
        for f in all_feeders
        if f.get("type") in ["feeder_400kv", "feeder_220kv", "ict_feeder", "reactor_feeder"]
    ]
    aliases_map: Dict[str, List[str]] = {}
    for f in feeders:
        fid = f["id"]
        aliases_map[fid] = _build_feeder_aliases(f["name"])

    preferred_names = ["chat.txt", "_chat.txt"]
    chat_path = None
    for name in preferred_names:
        candidate = ROOT_DIR.parent / name
        if candidate.exists():
            chat_path = candidate
            break
    if chat_path is None:
        raise FileNotFoundError(
            f"No WhatsApp chat export found. Expected one of: {', '.join(preferred_names)} in {ROOT_DIR.parent}"
        )

    year = 2026
    month = 1

    content = chat_path.read_text(encoding="utf-8", errors="ignore")
    messages = _parse_whatsapp_messages(content)
    print(f"Total WhatsApp messages parsed: {len(messages)}")

    events_by_feeder: Dict[str, List[Dict[str, Any]]] = {}
    for msg in messages:
        text = msg.get("text") or ""
        if not text:
            continue
        lower = text.lower()
        kind = _classify_interruption_message(text)
        matched_feeder_id = None
        matched_feeder_name = None
        for f in feeders:
            fid = f["id"]
            aliases = aliases_map.get(fid) or []
            if any(a in lower for a in aliases):
                matched_feeder_id = fid
                matched_feeder_name = f["name"]
                break
        if "thandur" in lower or "bus reactor" in lower:
            print(
                "DEBUG_MSG",
                msg["timestamp"],
                "KIND",
                kind,
                "FEEDER",
                matched_feeder_name,
                "TEXT",
                text.replace("\n", " ")[:200],
            )
        if not matched_feeder_id or not kind:
            continue
        ts = _extract_time_from_text(text, msg["timestamp"])
        if "thandur" in lower:
            print("DEBUG_TS_THANDUR", msg["timestamp"], "->", ts, "KIND", kind)
        if matched_feeder_id not in events_by_feeder:
            events_by_feeder[matched_feeder_id] = []
        events_by_feeder[matched_feeder_id].append(
            {"timestamp": ts, "kind": kind, "text": text, "feeder_name": matched_feeder_name}
        )

    entries: List[Dict[str, Any]] = []
    for f in feeders:
        fid = f["id"]
        events = events_by_feeder.get(fid, [])
        if not events:
            continue
        events = [e for e in events if e["timestamp"].year == year]
        if not events:
            continue
        events.sort(key=lambda e: e["timestamp"])
        open_outages: List[Dict[str, Any]] = []
        for ev in events:
            if ev["kind"] == "outage":
                open_outages.append(ev)
            elif ev["kind"] == "restore":
                if not open_outages:
                    continue
                outage = open_outages.pop(0)
                start_ts = outage["timestamp"]
                end_ts = ev["timestamp"]
                if start_ts.month != month:
                    continue
                if end_ts <= start_ts:
                    continue
                duration_minutes = (end_ts - start_ts).total_seconds() / 60.0
                date_str = start_ts.date().isoformat()
                start_time = start_ts.strftime("%H:%M")
                end_time = end_ts.strftime("%H:%M")
                entries.append(
                    {
                        "feeder_id": fid,
                        "feeder_name": ev["feeder_name"],
                        "date": date_str,
                        "start_time": start_time,
                        "end_time": end_time,
                        "duration_minutes": round(duration_minutes, 2),
                        "description": outage["text"],
                    }
                )

    print(f"Computed {len(entries)} interruption events for {year}-{month:02d}")

    if not entries:
        print("No interruption events found to import.")
        return

    print("Entries computed for import:")
    for e in entries:
        print(e)

    if db is None:
        print("MongoDB not configured; cannot import interruption events.")
        return

    print("Importing interruption events into MongoDB...")
    coll = db.interruption_entries
    inserted = 0
    skipped_existing = 0
    for e in entries:
        feeder_id = e.get("feeder_id")
        date_str = e.get("date")
        start_time = e.get("start_time")
        if not feeder_id or not date_str or not start_time:
            continue
        existing = coll.find_one(
            {"feeder_id": feeder_id, "date": date_str, "data.start_time": start_time}
        )
        if existing:
            skipped_existing += 1
            continue
        data = {
            "start_time": start_time,
            "end_time": e.get("end_time"),
            "duration_minutes": e.get("duration_minutes"),
            "description": e.get("description"),
        }
        obj = InterruptionEntry(feeder_id=feeder_id, date=date_str, data=data).model_dump()
        created_at = obj.get("created_at")
        updated_at = obj.get("updated_at")
        if isinstance(created_at, datetime):
            obj["created_at"] = created_at.isoformat()
        if isinstance(updated_at, datetime):
            obj["updated_at"] = updated_at.isoformat()
        coll.insert_one(obj)
        inserted += 1

    print(f"Inserted {inserted} interruption events, skipped {skipped_existing} existing.")

    print(f"\nVerifying interruption entries for {year}-{month:02d}:")
    total_found = 0
    for f in feeders:
        fid = f["id"]
        name = f["name"]
        resp = requests.get(
            f"{BASE_URL}/interruptions/entries/{fid}",
            params={"year": year, "month": month},
            headers=headers,
        )
        if resp.status_code != 200:
            print(f"- {name}: failed to fetch entries (status {resp.status_code})")
            continue
        data = resp.json()
        count = len(data)
        total_found += count
        print(f"- {name}: {count} entries")
    print(f"Total interruption entries returned by API for {year}-{month:02d}: {total_found}")

    print("\nSample documents from MongoDB for debug:")
    for e in entries:
        feeder_id = e.get("feeder_id")
        date_str = e.get("date")
        start_time = e.get("start_time")
        doc = coll.find_one(
            {"feeder_id": feeder_id, "date": date_str, "data.start_time": start_time},
            {"_id": 0},
        )
        if doc:
            print(doc)


if __name__ == "__main__":
    main()
