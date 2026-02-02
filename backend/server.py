from fastapi import FastAPI, APIRouter, HTTPException, Depends, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.responses import StreamingResponse
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List, Optional
import uuid
from datetime import datetime, timezone, timedelta
from passlib.context import CryptContext
import jwt
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

app = FastAPI()
api_router = APIRouter(prefix="/api")

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
security = HTTPBearer()

SECRET_KEY = os.environ.get('JWT_SECRET_KEY', 'your-secret-key-change-in-production')
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24 * 7

class UserRegister(BaseModel):
    email: str
    password: str
    full_name: str

class UserLogin(BaseModel):
    email: str
    password: str

class User(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    email: str
    full_name: str
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class UserInDB(User):
    hashed_password: str

class Token(BaseModel):
    access_token: str
    token_type: str
    user: User

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

class DailyEntry(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    feeder_id: str
    date: str
    end1_import_initial: float
    end1_import_final: float
    end1_export_initial: float
    end1_export_final: float
    end2_import_initial: float
    end2_import_final: float
    end2_export_initial: float
    end2_export_final: float
    end1_import_consumption: float = 0
    end1_export_consumption: float = 0
    end2_import_consumption: float = 0
    end2_export_consumption: float = 0
    loss_percent: float = 0
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    updated_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class DailyEntryCreate(BaseModel):
    feeder_id: str
    date: str
    end1_import_final: float
    end1_export_final: float
    end2_import_final: float
    end2_export_final: float

class DailyEntryUpdate(BaseModel):
    end1_import_initial: Optional[float] = None
    end1_import_final: Optional[float] = None
    end1_export_initial: Optional[float] = None
    end1_export_final: Optional[float] = None
    end2_import_initial: Optional[float] = None
    end2_import_final: Optional[float] = None
    end2_export_initial: Optional[float] = None
    end2_export_final: Optional[float] = None

class EnergySheet(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    meters: Optional[List['EnergyMeter']] = None

class EnergyMeter(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    sheet_id: str
    name: str
    mf: float
    unit: str
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class EnergyReading(BaseModel):
    model_config = ConfigDict(extra="ignore")
    meter_id: str
    initial: float
    final: float
    consumption: float

class EnergyEntry(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    sheet_id: str
    date: str
    readings: List[EnergyReading]
    total_consumption: float
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    updated_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class EnergyReadingInput(BaseModel):
    meter_id: str
    final: float

class EnergyEntryInput(BaseModel):
    sheet_id: str
    date: str
    readings: List[EnergyReadingInput]

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def get_password_hash(password):
    return pwd_context.hash(password)

def create_access_token(data: dict):
    to_encode = data.copy()
    expire = datetime.now(timezone.utc) + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)
    return encoded_jwt

async def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    try:
        token = credentials.credentials
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        user_id: str = payload.get("sub")
        if user_id is None:
            raise HTTPException(status_code=401, detail="Invalid token")
        user = await db.users.find_one({"id": user_id}, {"_id": 0})
        if user is None:
            raise HTTPException(status_code=401, detail="User not found")
        return User(**user)
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except Exception:
        raise HTTPException(status_code=401, detail="Invalid token")

@api_router.post("/auth/register", response_model=Token)
async def register(user_data: UserRegister):
    existing_user = await db.users.find_one({"email": user_data.email}, {"_id": 0})
    if existing_user:
        raise HTTPException(status_code=400, detail="Email already registered")
    
    hashed_password = get_password_hash(user_data.password)
    user_obj = UserInDB(
        email=user_data.email,
        full_name=user_data.full_name,
        hashed_password=hashed_password
    )
    
    doc = user_obj.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    await db.users.insert_one(doc)
    
    user_response = User(id=user_obj.id, email=user_obj.email, full_name=user_obj.full_name, created_at=user_obj.created_at)
    access_token = create_access_token(data={"sub": user_obj.id})
    
    return Token(access_token=access_token, token_type="bearer", user=user_response)

@api_router.post("/auth/login", response_model=Token)
async def login(user_data: UserLogin):
    user = await db.users.find_one({"email": user_data.email}, {"_id": 0})
    if not user:
        raise HTTPException(status_code=401, detail="Invalid credentials")
    
    if not verify_password(user_data.password, user['hashed_password']):
        raise HTTPException(status_code=401, detail="Invalid credentials")
    
    user_response = User(**user)
    access_token = create_access_token(data={"sub": user['id']})
    
    return Token(access_token=access_token, token_type="bearer", user=user_response)

@api_router.get("/auth/me", response_model=User)
async def get_me(current_user: User = Depends(get_current_user)):
    return current_user

@api_router.get("/feeders", response_model=List[Feeder])
async def get_feeders(current_user: User = Depends(get_current_user)):
    feeders = await db.feeders.find({}, {"_id": 0}).to_list(100)
    for feeder in feeders:
        if isinstance(feeder.get('created_at'), str):
            feeder['created_at'] = datetime.fromisoformat(feeder['created_at'])
    return feeders

@api_router.get("/feeders/{feeder_id}", response_model=Feeder)
async def get_feeder(feeder_id: str, current_user: User = Depends(get_current_user)):
    feeder = await db.feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    if isinstance(feeder.get('created_at'), str):
        feeder['created_at'] = datetime.fromisoformat(feeder['created_at'])
    return Feeder(**feeder)

@api_router.post("/entries", response_model=DailyEntry)
async def create_entry(entry_data: DailyEntryCreate, current_user: User = Depends(get_current_user)):
    feeder = await db.feeders.find_one({"id": entry_data.feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    
    existing_entry = await db.entries.find_one({"feeder_id": entry_data.feeder_id, "date": entry_data.date}, {"_id": 0})
    if existing_entry:
        raise HTTPException(status_code=400, detail="Entry already exists for this date")
    
    date_obj = datetime.strptime(entry_data.date, "%Y-%m-%d")
    prev_date = (date_obj - timedelta(days=1)).strftime("%Y-%m-%d")
    prev_entry = await db.entries.find_one({"feeder_id": entry_data.feeder_id, "date": prev_date}, {"_id": 0})
    
    if prev_entry:
        end1_import_initial = prev_entry['end1_import_final']
        end1_export_initial = prev_entry['end1_export_final']
        end2_import_initial = prev_entry['end2_import_final']
        end2_export_initial = prev_entry['end2_export_final']
    else:
        end1_import_initial = 0
        end1_export_initial = 0
        end2_import_initial = 0
        end2_export_initial = 0
    
    end1_import_consumption = (entry_data.end1_import_final - end1_import_initial) * feeder['end1_import_mf']
    end1_export_consumption = (entry_data.end1_export_final - end1_export_initial) * feeder['end1_export_mf']
    end2_import_consumption = (entry_data.end2_import_final - end2_import_initial) * feeder['end2_import_mf']
    end2_export_consumption = (entry_data.end2_export_final - end2_export_initial) * feeder['end2_export_mf']
    
    total_import = end1_import_consumption + end2_import_consumption
    if total_import == 0:
        loss_percent = 0
    else:
        loss_percent = ((end1_import_consumption - end1_export_consumption + end2_import_consumption - end2_export_consumption) / total_import) * 100
    
    entry_obj = DailyEntry(
        feeder_id=entry_data.feeder_id,
        date=entry_data.date,
        end1_import_initial=end1_import_initial,
        end1_import_final=entry_data.end1_import_final,
        end1_export_initial=end1_export_initial,
        end1_export_final=entry_data.end1_export_final,
        end2_import_initial=end2_import_initial,
        end2_import_final=entry_data.end2_import_final,
        end2_export_initial=end2_export_initial,
        end2_export_final=entry_data.end2_export_final,
        end1_import_consumption=end1_import_consumption,
        end1_export_consumption=end1_export_consumption,
        end2_import_consumption=end2_import_consumption,
        end2_export_consumption=end2_export_consumption,
        loss_percent=loss_percent
    )
    
    doc = entry_obj.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    doc['updated_at'] = doc['updated_at'].isoformat()
    await db.entries.insert_one(doc)
    
    return entry_obj

@api_router.get("/entries", response_model=List[DailyEntry])
async def get_entries(
    feeder_id: Optional[str] = None,
    year: Optional[int] = None,
    month: Optional[int] = None,
    current_user: User = Depends(get_current_user)
):
    query = {}
    if feeder_id:
        query['feeder_id'] = feeder_id
    if year and month:
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"
        query['date'] = {"$gte": start_date, "$lt": end_date}
    
    entries = await db.entries.find(query, {"_id": 0}).sort("date", 1).to_list(1000)
    for entry in entries:
        if isinstance(entry.get('created_at'), str):
            entry['created_at'] = datetime.fromisoformat(entry['created_at'])
        if isinstance(entry.get('updated_at'), str):
            entry['updated_at'] = datetime.fromisoformat(entry['updated_at'])
    return entries

@api_router.put("/entries/{entry_id}", response_model=DailyEntry)
async def update_entry(entry_id: str, update_data: DailyEntryUpdate, current_user: User = Depends(get_current_user)):
    entry = await db.entries.find_one({"id": entry_id}, {"_id": 0})
    if not entry:
        raise HTTPException(status_code=404, detail="Entry not found")
    
    feeder = await db.feeders.find_one({"id": entry['feeder_id']}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    
    update_dict = {k: v for k, v in update_data.model_dump().items() if v is not None}
    
    for key, value in update_dict.items():
        entry[key] = value
    
    end1_import_consumption = (entry['end1_import_final'] - entry['end1_import_initial']) * feeder['end1_import_mf']
    end1_export_consumption = (entry['end1_export_final'] - entry['end1_export_initial']) * feeder['end1_export_mf']
    end2_import_consumption = (entry['end2_import_final'] - entry['end2_import_initial']) * feeder['end2_import_mf']
    end2_export_consumption = (entry['end2_export_final'] - entry['end2_export_initial']) * feeder['end2_export_mf']
    
    total_import = end1_import_consumption + end2_import_consumption
    if total_import == 0:
        loss_percent = 0
    else:
        loss_percent = ((end1_import_consumption - end1_export_consumption + end2_import_consumption - end2_export_consumption) / total_import) * 100
    
    entry['end1_import_consumption'] = end1_import_consumption
    entry['end1_export_consumption'] = end1_export_consumption
    entry['end2_import_consumption'] = end2_import_consumption
    entry['end2_export_consumption'] = end2_export_consumption
    entry['loss_percent'] = loss_percent
    entry['updated_at'] = datetime.now(timezone.utc).isoformat()
    
    await db.entries.update_one({"id": entry_id}, {"$set": entry})
    
    if isinstance(entry.get('created_at'), str):
        entry['created_at'] = datetime.fromisoformat(entry['created_at'])
    if isinstance(entry.get('updated_at'), str):
        entry['updated_at'] = datetime.fromisoformat(entry['updated_at'])
    
    return DailyEntry(**entry)

@api_router.delete("/entries/{entry_id}")
async def delete_entry(entry_id: str, current_user: User = Depends(get_current_user)):
    result = await db.entries.delete_one({"id": entry_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Entry not found")
    return {"message": "Entry deleted successfully"}

@api_router.get("/export/{feeder_id}/{year}/{month}")
async def export_feeder_data(feeder_id: str, year: int, month: int, current_user: User = Depends(get_current_user)):
    feeder = await db.feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
    
    entries = await db.entries.find(
        {"feeder_id": feeder_id, "date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).sort("date", 1).to_list(1000)
    
    wb = Workbook()
    ws = wb.active
    ws.title = feeder['name'][:31]
    
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    headers = [
        "Date",
        f"{feeder['end1_name']} Import Initial",
        f"{feeder['end1_name']} Import Final",
        f"{feeder['end1_name']} Import MF",
        f"{feeder['end1_name']} Import Consumption",
        f"{feeder['end1_name']} Export Initial",
        f"{feeder['end1_name']} Export Final",
        f"{feeder['end1_name']} Export MF",
        f"{feeder['end1_name']} Export Consumption",
        f"{feeder['end2_name']} Import Initial",
        f"{feeder['end2_name']} Import Final",
        f"{feeder['end2_name']} Import MF",
        f"{feeder['end2_name']} Import Consumption",
        f"{feeder['end2_name']} Export Initial",
        f"{feeder['end2_name']} Export Final",
        f"{feeder['end2_name']} Export MF",
        f"{feeder['end2_name']} Export Consumption",
        "% Loss"
    ]
    
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    
    for entry in entries:
        ws.append([
            entry['date'],
            entry['end1_import_initial'],
            entry['end1_import_final'],
            feeder['end1_import_mf'],
            entry['end1_import_consumption'],
            entry['end1_export_initial'],
            entry['end1_export_final'],
            feeder['end1_export_mf'],
            entry['end1_export_consumption'],
            entry['end2_import_initial'],
            entry['end2_import_final'],
            feeder['end2_import_mf'],
            entry['end2_import_consumption'],
            entry['end2_export_initial'],
            entry['end2_export_final'],
            feeder['end2_export_mf'],
            entry['end2_export_consumption'],
            entry['loss_percent']
        ])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"{feeder['name']}_{year}_{month:02d}.xlsx"
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@api_router.post("/init-feeders")
async def initialize_feeders(current_user: User = Depends(get_current_user)):
    existing_count = await db.feeders.count_documents({})
    if existing_count > 0:
        return {"message": "Feeders already initialized", "count": existing_count}
    
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
    
    feeders = []
    for fd in feeders_data:
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
        feeders.append(doc)
    
    await db.feeders.insert_many(feeders)
    return {"message": "Feeders initialized successfully", "count": len(feeders)}

@api_router.post("/energy/init")
async def initialize_energy_module(current_user: User = Depends(get_current_user)):
    # Check if already initialized
    if await db.energy_sheets.count_documents({}) > 0:
        return {"message": "Energy module already initialized"}

    # Create Sheets
    sheets_data = ["ICT-1", "ICT-2", "ICT-3", "ICT-4", "33KV"]
    sheet_ids = {}
    
    for name in sheets_data:
        sheet = EnergySheet(name=name)
        doc = sheet.model_dump()
        await db.energy_sheets.insert_one(doc)
        sheet_ids[name] = sheet.id

    # Create Meters
    meters = []
    
    # ICT Meters
    for i in range(1, 5):
        sheet_name = f"ICT-{i}"
        s_id = sheet_ids[sheet_name]
        meters.append(EnergyMeter(sheet_id=s_id, name="HV", mf=0.1 if i==1 else 1.0, unit="MWH")) # Example MF, user said fixed but didn't specify all. using 1.0 default or from screenshot
        meters.append(EnergyMeter(sheet_id=s_id, name="IV", mf=1.5 if i==1 else 1.0, unit="MWH"))
        # Note: Screenshot for ICT-1 shows HV MF=0.1, IV MF=1.5. I'll use those for ICT-1. Others 1.0 for now.

    # 33KV Meters
    s33_id = sheet_ids["33KV"]
    meters.append(EnergyMeter(sheet_id=s33_id, name="Donthapally", mf=1000, unit="KWH")) # Screenshot says MF(0.2) but column says 1000? Wait.
    # Screenshot 33KV: "MF(0.2)" in header? No, "MF(0.2)" might be the CT/PT ratio or something.
    # But column C says "1000". And formula D=(B-A)*C. So MF is 1000.
    # Another meter: "33KV Kandi", MF column says "4".
    meters.append(EnergyMeter(sheet_id=s33_id, name="Kandi", mf=4, unit="KWH"))

    for m in meters:
        doc = m.model_dump()
        await db.energy_meters.insert_one(doc)

    return {"message": "Energy module initialized", "sheets": len(sheets_data), "meters": len(meters)}

@api_router.get("/energy/sheets")
async def get_energy_sheets(current_user: User = Depends(get_current_user)):
    sheets = await db.energy_sheets.find({}, {"_id": 0}).to_list(100)
    result = []
    for sheet in sheets:
        meters = await db.energy_meters.find({"sheet_id": sheet['id']}, {"_id": 0}).to_list(100)
        sheet['meters'] = meters
        result.append(sheet)
    return result

@api_router.get("/energy/entries/{sheet_id}")
async def get_energy_entries(
    sheet_id: str,
    year: Optional[int] = None,
    month: Optional[int] = None,
    current_user: User = Depends(get_current_user)
):
    query = {"sheet_id": sheet_id}
    if year and month:
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"
        query['date'] = {"$gte": start_date, "$lt": end_date}
    
    entries = await db.energy_entries.find(query, {"_id": 0}).sort("date", 1).to_list(1000)
    for entry in entries:
        if isinstance(entry.get('created_at'), str):
            entry['created_at'] = datetime.fromisoformat(entry['created_at'])
        if isinstance(entry.get('updated_at'), str):
            entry['updated_at'] = datetime.fromisoformat(entry['updated_at'])
    return entries

@api_router.post("/energy/entries")
async def save_energy_entry(
    entry_input: EnergyEntryInput,
    current_user: User = Depends(get_current_user)
):
    # Get previous day's entry for initials
    date_obj = datetime.strptime(entry_input.date, "%Y-%m-%d")
    prev_date = (date_obj - timedelta(days=1)).strftime("%Y-%m-%d")
    prev_entry = await db.energy_entries.find_one(
        {"sheet_id": entry_input.sheet_id, "date": prev_date},
        {"_id": 0}
    )
    
    prev_finals = {}
    if prev_entry:
        for r in prev_entry['readings']:
            prev_finals[r['meter_id']] = r['final']
            
    # Calculate readings
    readings = []
    total_consumption = 0
    
    for r_in in entry_input.readings:
        meter = await db.energy_meters.find_one({"id": r_in.meter_id})
        if not meter:
            continue
            
        initial = prev_finals.get(r_in.meter_id, 0.0)
        consumption = (r_in.final - initial) * meter['mf']
        total_consumption += consumption
        
        readings.append(EnergyReading(
            meter_id=r_in.meter_id,
            initial=initial,
            final=r_in.final,
            consumption=consumption
        ))
        
    # Upsert entry
    existing = await db.energy_entries.find_one(
        {"sheet_id": entry_input.sheet_id, "date": entry_input.date}
    )
    
    entry_data = EnergyEntry(
        sheet_id=entry_input.sheet_id,
        date=entry_input.date,
        readings=readings,
        total_consumption=total_consumption
    )
    
    doc = entry_data.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    doc['updated_at'] = doc['updated_at'].isoformat()
    
    if existing:
        # Preserve original created_at
        doc['created_at'] = existing['created_at']
        await db.energy_entries.replace_one({"_id": existing['_id']}, doc)
    else:
        await db.energy_entries.insert_one(doc)
        
    return entry_data

@api_router.get("/energy/export/{sheet_id}/{year}/{month}")
async def export_energy_sheet(
    sheet_id: str,
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    sheet = await db.energy_sheets.find_one({"id": sheet_id}, {"_id": 0})
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
        
    meters = await db.energy_meters.find({"sheet_id": sheet_id}, {"_id": 0}).to_list(100)
    
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
        
    entries = await db.energy_entries.find(
        {"sheet_id": sheet_id, "date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).sort("date", 1).to_list(1000)
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"{sheet['name']} - {month}-{year}"
    
    # Headers
    headers = ["Date"]
    for m in meters:
        headers.extend([
            f"{m['name']} Initial",
            f"{m['name']} Final",
            f"{m['name']} MF",
            f"{m['name']} Consumption"
        ])
    headers.append("Total Consumption")
    
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        
    # Data
    for entry in entries:
        row = [entry['date']]
        readings_map = {r['meter_id']: r for r in entry['readings']}
        
        for m in meters:
            r = readings_map.get(m['id'])
            if r:
                row.extend([r['initial'], r['final'], m['mf'], r['consumption']])
            else:
                row.extend([0, 0, m['mf'], 0])
                
        row.append(entry['total_consumption'])
        ws.append(row)
        
    # Auto-width
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"{sheet['name']}_{year}_{month:02d}.xlsx"
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@api_router.put("/energy/entries/{entry_id}", response_model=EnergyEntry)
async def update_energy_entry(
    entry_id: str,
    entry_input: EnergyEntryInput,
    current_user: User = Depends(get_current_user)
):
    # Verify entry exists
    existing = await db.energy_entries.find_one({"id": entry_id})
    if not existing:
        raise HTTPException(status_code=404, detail="Entry not found")

    # Get previous day's entry for initials
    date_obj = datetime.strptime(entry_input.date, "%Y-%m-%d")
    prev_date = (date_obj - timedelta(days=1)).strftime("%Y-%m-%d")
    prev_entry = await db.energy_entries.find_one(
        {"sheet_id": entry_input.sheet_id, "date": prev_date},
        {"_id": 0}
    )
    
    prev_finals = {}
    if prev_entry:
        for r in prev_entry['readings']:
            prev_finals[r['meter_id']] = r['final']
            
    # Calculate readings
    readings = []
    total_consumption = 0
    
    for r_in in entry_input.readings:
        meter = await db.energy_meters.find_one({"id": r_in.meter_id})
        if not meter:
            continue
            
        initial = prev_finals.get(r_in.meter_id, 0.0)
        consumption = (r_in.final - initial) * meter['mf']
        total_consumption += consumption
        
        readings.append(EnergyReading(
            meter_id=r_in.meter_id,
            initial=initial,
            final=r_in.final,
            consumption=consumption
        ))
        
    # Update entry
    entry_data = EnergyEntry(
        id=entry_id, # Preserve ID
        sheet_id=entry_input.sheet_id,
        date=entry_input.date,
        readings=readings,
        total_consumption=total_consumption,
        created_at=existing['created_at'], # Preserve created_at
        updated_at=datetime.now(timezone.utc)
    )
    
    doc = entry_data.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    doc['updated_at'] = doc['updated_at'].isoformat()
    
    await db.energy_entries.replace_one({"id": entry_id}, doc)
        
    return entry_data

@api_router.delete("/energy/entries/{entry_id}")
async def delete_energy_entry(entry_id: str, current_user: User = Depends(get_current_user)):
    result = await db.energy_entries.delete_one({"id": entry_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Entry not found")
    return {"message": "Entry deleted successfully"}


app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()