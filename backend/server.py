from fastapi import FastAPI, APIRouter, HTTPException, Depends, status, UploadFile, File
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
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random
import string
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import load_workbook

def format_date(date_str):
    if not date_str:
        return "-"
    try:
        parts = date_str.split("-")
        if len(parts) == 3:
            return f"{parts[2]}-{parts[1]}-{parts[0]}"
    except:
        pass
    return date_str

def format_time(time_str):
    if not time_str:
        return ""
    time_str = str(time_str).strip()
    try:
        if ':' in time_str:
            parts = time_str.split(':')
            if len(parts) >= 2:
                return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    except:
        pass
    return time_str

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

def send_otp_email(user_email: str, otp: str, reason: str = "reset"):
    sender_email = os.environ.get("SMTP_EMAIL")
    sender_password = os.environ.get("SMTP_PASSWORD")
    # Admin email is same as sender in this requirement
    admin_email = sender_email 
    
    message = MIMEMultipart("alternative")
    message["From"] = sender_email
    message["To"] = admin_email

    if reason == "signup":
        message["Subject"] = "New Account Signup OTP Request"
        action_text = "new account signup"
    else:
        message["Subject"] = "Password Reset OTP Request"
        action_text = "password reset"

    text = f"""\
    A {action_text} was requested for user: {user_email}.
    Your OTP is: {otp}
    """
    html = f"""\
    <html>
      <body>
        <h2>{message["Subject"]}</h2>
        <p>A {action_text} was requested for user: <b>{user_email}</b>.</p>
        <p>Your OTP is: <b style="font-size: 24px;">{otp}</b></p>
        <p>Please approve this request by sharing the OTP with the user if appropriate.</p>
      </body>
    </html>
    """

    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")
    message.attach(part1)
    message.attach(part2)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, admin_email, message.as_string())
    except Exception as e:
        print(f"Error sending email: {e}")
        raise HTTPException(status_code=500, detail="Failed to send OTP email")

mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

app = FastAPI()

origins = [
    "http://localhost:3000",
    "https://mis-portal-liard.vercel.app",
]

env_origins = os.environ.get('CORS_ORIGINS', '')
if env_origins:
    origins.extend(env_origins.split(','))

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=origins,
    allow_methods=["*"],
    allow_headers=["*"],
)

api_router = APIRouter(prefix="/api")

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
security = HTTPBearer()

SECRET_KEY = os.environ['JWT_SECRET_KEY']
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24 * 7

class UserRegister(BaseModel):
    email: str
    password: str
    full_name: str

class UserLogin(BaseModel):
    email: str
    password: str

class ForgotPasswordRequest(BaseModel):
    email: str

class ResetPasswordRequest(BaseModel):
    email: str
    otp: str
    new_password: str

class SignupVerifyRequest(BaseModel):
    email: str
    otp: str

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

class MaxMinFeeder(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    type: str  # bus_station, feeder_400kv, feeder_220kv, ict_feeder
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class MaxMinEntry(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    feeder_id: str
    date: str
    data: dict
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    updated_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class MaxMinEntryCreate(BaseModel):
    feeder_id: str
    date: str
    data: dict

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

@api_router.post("/auth/forgot-password")
async def forgot_password(request: ForgotPasswordRequest):
    user = await db.users.find_one({"email": request.email})
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    
    otp = ''.join(random.choices(string.digits, k=6))
    # Save OTP with expiry (15 mins)
    await db.password_resets.update_one(
        {"email": request.email},
        {"$set": {"otp": otp, "expires_at": datetime.now(timezone.utc) + timedelta(minutes=15)}},
        upsert=True
    )
    
    # Send to admin
    send_otp_email(request.email, otp)
    
    return {"message": "OTP sent to admin for approval"}

@api_router.post("/auth/reset-password")
async def reset_password(request: ResetPasswordRequest):
    record = await db.password_resets.find_one({"email": request.email})
    if not record:
        raise HTTPException(status_code=400, detail="Invalid request")
        
    if record["otp"] != request.otp:
        raise HTTPException(status_code=400, detail="Invalid OTP")
        
    # Handle both naive and aware datetime for expires_at
    expires_at = record["expires_at"]
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
        
    if expires_at < datetime.now(timezone.utc):
         raise HTTPException(status_code=400, detail="OTP expired")
         
    # Update password
    hashed_password = get_password_hash(request.new_password)
    await db.users.update_one(
        {"email": request.email},
        {"$set": {"hashed_password": hashed_password}}
    )
    
    # Delete OTP
    await db.password_resets.delete_one({"email": request.email})
    
    return {"message": "Password reset successful"}

@api_router.post("/auth/signup-request")
async def signup_request(user_data: UserRegister):
    existing_user = await db.users.find_one({"email": user_data.email})
    if existing_user:
        raise HTTPException(status_code=400, detail="Email already registered")
    
    otp = ''.join(random.choices(string.digits, k=6))
    hashed_password = get_password_hash(user_data.password)
    
    request_data = {
        "email": user_data.email,
        "full_name": user_data.full_name,
        "hashed_password": hashed_password,
        "otp": otp,
        "expires_at": datetime.now(timezone.utc) + timedelta(minutes=15),
        "created_at": datetime.now(timezone.utc)
    }
    
    await db.signup_requests.update_one(
        {"email": user_data.email},
        {"$set": request_data},
        upsert=True
    )
    
    send_otp_email(user_data.email, otp, reason="signup")
    
    return {"message": "OTP sent to admin for approval"}

@api_router.post("/auth/signup-verify", response_model=Token)
async def signup_verify(request: SignupVerifyRequest):
    record = await db.signup_requests.find_one({"email": request.email})
    if not record:
        raise HTTPException(status_code=400, detail="Invalid request or expired")
        
    if record["otp"] != request.otp:
        raise HTTPException(status_code=400, detail="Invalid OTP")
        
    expires_at = record["expires_at"]
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
        
    if expires_at < datetime.now(timezone.utc):
         raise HTTPException(status_code=400, detail="OTP expired")
    
    # Create user
    user_obj = UserInDB(
        email=record["email"],
        full_name=record["full_name"],
        hashed_password=record["hashed_password"]
    )
    
    doc = user_obj.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    await db.users.insert_one(doc)
    
    # Delete request
    await db.signup_requests.delete_one({"email": request.email})
    
    user_response = User(id=user_obj.id, email=user_obj.email, full_name=user_obj.full_name, created_at=user_obj.created_at)
    access_token = create_access_token(data={"sub": user_obj.id})
    
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

@api_router.post("/max-min/preview-import/{feeder_id}")
async def preview_max_min_import(
    feeder_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user)
):
    feeder = await db.max_min_feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only .xlsx or .xls files are supported")
    content = await file.read()
    wb = load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    headers = [str(ws.cell(row=1, column=col).value or "").strip() for col in range(1, ws.max_column + 1)]
    def tokens(h):
        return h.lower().replace("_", " ").replace("-", " ").replace(".", " ").replace("(", " ").replace(")", " ").split()
    def find_col(required):
        for idx, h in enumerate(headers):
            t = tokens(h)
            if all(r in t for r in required):
                return idx + 1
        return None
    preview = []
    ict_ids = []
    if feeder['type'] == 'bus_station':
        ict_feeders = await db.max_min_feeders.find({"type": "ict_feeder"}, {"id": 1}).to_list(None)
        ict_ids = [f['id'] for f in ict_feeders]

    for row in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row, column=1).value
        if not date_cell:
            continue
        if isinstance(date_cell, datetime):
            date_str = date_cell.date().isoformat()
        else:
            val_str = str(date_cell).strip()
            fmts = ["%Y-%m-%d", "%d-%m-%Y", "%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"]
            parsed = False
            for fmt in fmts:
                try:
                    date_str = datetime.strptime(val_str, fmt).date().isoformat()
                    parsed = True
                    break
                except:
                    continue
            if not parsed:
                continue
        data = {}
        if feeder['type'] == 'bus_station':
            v400_max_col = find_col({"max", "bus", "voltage", "400kv"}) or find_col({"max", "400kv"})
            v220_max_col = find_col({"max", "bus", "voltage", "220kv"}) or find_col({"max", "220kv"})
            v400_min_col = find_col({"min", "bus", "voltage", "400kv"}) or find_col({"min", "400kv"})
            v220_min_col = find_col({"min", "bus", "voltage", "220kv"}) or find_col({"min", "220kv"})
            max_time_col = find_col({"max", "time"})
            min_time_col = find_col({"min", "time"})
            load_max_mw_col = find_col({"station", "load", "max", "mw"}) or find_col({"max", "mw"})
            load_mvar_col = find_col({"station", "load", "mvar"}) or find_col({"mvar"})
            load_time_col = find_col({"station", "load", "time"})

            def getf(c):
                try:
                    v = ws.cell(row=row, column=c).value if c else None
                    return float(v) if v is not None and str(v).strip() != "" else None
                except:
                    return None
            def gets(c):
                v = ws.cell(row=row, column=c).value if c else None
                return str(v).strip() if v is not None else None
            max_time = gets(max_time_col)
            min_time = gets(min_time_col)
            
            station_time = gets(load_time_col)
            station_max_mw = getf(load_max_mw_col)
            station_mvar = getf(load_mvar_col)
            
            if ict_ids:
                ict_entries = await db.max_min_entries.find(
                    {"feeder_id": {"$in": ict_ids}, "date": date_str},
                    {"data.max": 1}
                ).to_list(None)
                
                times = []
                total_mw = 0.0
                total_mvar = 0.0
                has_ict_data = False
                
                for e in ict_entries:
                    d = e.get('data', {}).get('max', {})
                    t = d.get('time')
                    mw = d.get('mw')
                    mvar = d.get('mvar')
                    
                    if t: times.append(t)
                    if mw is not None: 
                        try: total_mw += float(mw)
                        except: pass
                        has_ict_data = True
                    if mvar is not None:
                        try: total_mvar += float(mvar)
                        except: pass
                
                if has_ict_data:
                    station_max_mw = total_mw
                    station_mvar = total_mvar
                
                times = [t for t in times if t]
                if times:
                    from collections import Counter
                    station_time = Counter(times).most_common(1)[0][0]

            data = {
                "max_bus_voltage_400kv": {"value": getf(v400_max_col), "time": max_time},
                "max_bus_voltage_220kv": {"value": getf(v220_max_col), "time": max_time},
                "min_bus_voltage_400kv": {"value": getf(v400_min_col), "time": min_time},
                "min_bus_voltage_220kv": {"value": getf(v220_min_col), "time": min_time},
                "station_load": {"max_mw": station_max_mw, "mvar": station_mvar, "time": station_time}
            }
            if all(
                x is None for x in [
                    data["max_bus_voltage_400kv"]["value"],
                    data["max_bus_voltage_220kv"]["value"],
                    data["min_bus_voltage_400kv"]["value"],
                    data["min_bus_voltage_220kv"]["value"],
                    data["station_load"]["max_mw"],
                    data["station_load"]["mvar"]
                ]
            ):
                continue
        elif feeder['type'] == 'ict_feeder':
            max_amps_col = find_col({"max", "amps"})
            max_mw_col = find_col({"max", "mw"})
            max_mvar_col = find_col({"max", "mvar"})
            max_time_col = find_col({"max", "time"})
            min_amps_col = find_col({"min", "amps"})
            min_mw_col = find_col({"min", "mw"})
            min_mvar_col = find_col({"min", "mvar"})
            min_time_col = find_col({"min", "time"})
            avg_amps_col = find_col({"avg", "amps"})
            avg_mw_col = find_col({"avg", "mw"})
            def getf(c):
                try:
                    v = ws.cell(row=row, column=c).value if c else None
                    return float(v) if v is not None and str(v).strip() != "" else None
                except:
                    return None
            def gets(c):
                v = ws.cell(row=row, column=c).value if c else None
                return str(v).strip() if v is not None else None
            
            # Calculate Avg if missing
            max_amps_val = getf(max_amps_col)
            min_amps_val = getf(min_amps_col)
            avg_amps_val = getf(avg_amps_col)
            
            if avg_amps_val is None and max_amps_val is not None and min_amps_val is not None:
                avg_amps_val = (max_amps_val + min_amps_val) / 2
                
            max_mw_val = getf(max_mw_col)
            min_mw_val = getf(min_mw_col)
            avg_mw_val = getf(avg_mw_col)
            
            if avg_mw_val is None and max_mw_val is not None and min_mw_val is not None:
                avg_mw_val = (max_mw_val + min_mw_val) / 2

            data = {
                "max": {"amps": max_amps_val, "mw": max_mw_val, "mvar": getf(max_mvar_col), "time": gets(max_time_col)},
                "min": {"amps": min_amps_val, "mw": min_mw_val, "mvar": getf(min_mvar_col), "time": gets(min_time_col)},
                "avg": {"amps": avg_amps_val, "mw": avg_mw_val}
            }
            if all(x is None for x in [data["max"]["amps"], data["max"]["mw"], data["min"]["amps"], data["min"]["mw"], data["avg"]["amps"], data["avg"]["mw"]]):
                continue
        else:
            max_amps_col = find_col({"max", "amps"})
            max_mw_col = find_col({"max", "mw"})
            max_time_col = find_col({"max", "time"})
            min_amps_col = find_col({"min", "amps"})
            min_mw_col = find_col({"min", "mw"})
            min_time_col = find_col({"min", "time"})
            avg_amps_col = find_col({"avg", "amps"})
            avg_mw_col = find_col({"avg", "mw"})
            def getf(c):
                try:
                    v = ws.cell(row=row, column=c).value if c else None
                    return float(v) if v is not None and str(v).strip() != "" else None
                except:
                    return None
            def gets(c):
                v = ws.cell(row=row, column=c).value if c else None
                return str(v).strip() if v is not None else None
            
            # Calculate Avg if missing
            max_amps_val = getf(max_amps_col)
            min_amps_val = getf(min_amps_col)
            avg_amps_val = getf(avg_amps_col)
            
            if avg_amps_val is None and max_amps_val is not None and min_amps_val is not None:
                avg_amps_val = (max_amps_val + min_amps_val) / 2
                
            max_mw_val = getf(max_mw_col)
            min_mw_val = getf(min_mw_col)
            avg_mw_val = getf(avg_mw_col)
            
            if avg_mw_val is None and max_mw_val is not None and min_mw_val is not None:
                avg_mw_val = (max_mw_val + min_mw_val) / 2

            data = {
                "max": {"amps": max_amps_val, "mw": max_mw_val, "time": gets(max_time_col)},
                "min": {"amps": min_amps_val, "mw": min_mw_val, "time": gets(min_time_col)},
                "avg": {"amps": avg_amps_val, "mw": avg_mw_val}
            }
            if all(x is None for x in [data["max"]["amps"], data["max"]["mw"], data["min"]["amps"], data["min"]["mw"], data["avg"]["amps"], data["avg"]["mw"]]):
                continue
        exists = await db.max_min_entries.find_one({"feeder_id": feeder_id, "date": date_str})
        preview.append({"date": date_str, "data": data, "exists": bool(exists)})
    return preview

class MaxMinImportPayload(BaseModel):
    feeder_id: str
    entries: List[dict]

@api_router.post("/max-min/import-entries")
async def import_max_min_entries(payload: MaxMinImportPayload, current_user: User = Depends(get_current_user)):
    feeder_id = payload.feeder_id
    feeder = await db.max_min_feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    entries_sorted = sorted(payload.entries, key=lambda x: x['date'])
    imported = 0
    for e in entries_sorted:
        date_str = e['date']
        existing = await db.max_min_entries.find_one({"feeder_id": feeder_id, "date": date_str})
        if existing:
            await db.max_min_entries.update_one(
                {"_id": existing["_id"]},
                {"$set": {
                    "data": e.get("data", {}),
                    "updated_at": datetime.now(timezone.utc).isoformat()
                }}
            )
            imported += 1
            continue
        entry_obj = MaxMinEntry(feeder_id=feeder_id, date=date_str, data=e.get("data", {}))
        doc = entry_obj.model_dump()
        doc['created_at'] = doc['created_at'].isoformat()
        doc['updated_at'] = doc['updated_at'].isoformat()
        await db.max_min_entries.insert_one(doc)
        imported += 1
    return {"imported": imported}

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
    
    # Update next day's initial values if exists
    date_obj = datetime.strptime(entry['date'], "%Y-%m-%d")
    next_date = (date_obj + timedelta(days=1)).strftime("%Y-%m-%d")
    next_entry = await db.entries.find_one({"feeder_id": entry['feeder_id'], "date": next_date})
    
    if next_entry:
        next_entry['end1_import_initial'] = entry['end1_import_final']
        next_entry['end1_export_initial'] = entry['end1_export_final']
        next_entry['end2_import_initial'] = entry['end2_import_final']
        next_entry['end2_export_initial'] = entry['end2_export_final']
        
        # Recalculate next entry consumption
        next_entry['end1_import_consumption'] = (next_entry['end1_import_final'] - next_entry['end1_import_initial']) * feeder['end1_import_mf']
        next_entry['end1_export_consumption'] = (next_entry['end1_export_final'] - next_entry['end1_export_initial']) * feeder['end1_export_mf']
        next_entry['end2_import_consumption'] = (next_entry['end2_import_final'] - next_entry['end2_import_initial']) * feeder['end2_import_mf']
        next_entry['end2_export_consumption'] = (next_entry['end2_export_final'] - next_entry['end2_export_initial']) * feeder['end2_export_mf']
        
        total_import = next_entry['end1_import_consumption'] + next_entry['end2_import_consumption']
        if total_import == 0:
            next_entry['loss_percent'] = 0
        else:
            next_entry['loss_percent'] = ((next_entry['end1_import_consumption'] - next_entry['end1_export_consumption'] + next_entry['end2_import_consumption'] - next_entry['end2_export_consumption']) / total_import) * 100
            
        next_entry['updated_at'] = datetime.now(timezone.utc).isoformat()
        await db.entries.replace_one({"_id": next_entry['_id']}, next_entry)
    
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

# Helper to parse float safely
def get_float(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

# Helper to calculate stats for a period
def calculate_period_stats(period_entries, feeder_type):
    stats = {}
    
    if feeder_type == 'bus_station':
        # Initialize stats with defaults
        stats = {
            'max_400kv': -float('inf'), 'max_400kv_date': '-', 'max_400kv_time': '-',
            'min_400kv': float('inf'), 'min_400kv_date': '-', 'min_400kv_time': '-',
            'max_220kv': -float('inf'), 'max_220kv_date': '-', 'max_220kv_time': '-',
            'min_220kv': float('inf'), 'min_220kv_date': '-', 'min_220kv_time': '-',
            'max_load': -float('inf'), 'max_load_date': '-', 'max_load_time': '-', 'max_load_mvar': '-'
        }

        # 1. Find Max/Min Values
        for e in period_entries:
            d = e.get('data', {})
            
            # 400KV
            v = get_float(d.get('max_bus_voltage_400kv', {}).get('value'))
            if v is not None and v > stats['max_400kv']: stats['max_400kv'] = v
            v = get_float(d.get('min_bus_voltage_400kv', {}).get('value'))
            if v is not None and v < stats['min_400kv']: stats['min_400kv'] = v
            
            # 220KV
            v = get_float(d.get('max_bus_voltage_220kv', {}).get('value'))
            if v is not None and v > stats['max_220kv']: stats['max_220kv'] = v
            v = get_float(d.get('min_bus_voltage_220kv', {}).get('value'))
            if v is not None and v < stats['min_220kv']: stats['min_220kv'] = v

            # Load (Independent)
            v = get_float(d.get('station_load', {}).get('max_mw'))
            if v is not None and v > stats['max_load']:
                stats['max_load'] = v
                stats['max_load_date'] = e['date']
                stats['max_load_time'] = d.get('station_load', {}).get('time', '-')
                stats['max_load_mvar'] = d.get('station_load', {}).get('mvar', '-')

        # 2. Collect Candidates
        cands_max_400 = []
        cands_min_400 = []
        cands_max_220 = []
        cands_min_220 = []

        for e in period_entries:
            d = e.get('data', {})
            date = e['date']
            
            # Max 400
            v = get_float(d.get('max_bus_voltage_400kv', {}).get('value'))
            if v == stats['max_400kv'] and v is not None:
                cands_max_400.append({'date': date, 'time': str(d.get('max_bus_voltage_400kv', {}).get('time', '')).strip()})
            
            # Min 400
            v = get_float(d.get('min_bus_voltage_400kv', {}).get('value'))
            if v == stats['min_400kv'] and v is not None:
                cands_min_400.append({'date': date, 'time': str(d.get('min_bus_voltage_400kv', {}).get('time', '')).strip()})
            
            # Max 220
            v = get_float(d.get('max_bus_voltage_220kv', {}).get('value'))
            if v == stats['max_220kv'] and v is not None:
                cands_max_220.append({'date': date, 'time': str(d.get('max_bus_voltage_220kv', {}).get('time', '')).strip()})
            
            # Min 220
            v = get_float(d.get('min_bus_voltage_220kv', {}).get('value'))
            if v == stats['min_220kv'] and v is not None:
                cands_min_220.append({'date': date, 'time': str(d.get('min_bus_voltage_220kv', {}).get('time', '')).strip()})

        # 3. Find Best Matches (Common Time Priority)
        def find_best(list_a, list_b):
            # Try to find intersection
            for a in list_a:
                for b in list_b:
                    if a['date'] == b['date'] and a['time'] == b['time']:
                        return a, b
            # No intersection, return first available
            res_a = list_a[0] if list_a else {'date': '-', 'time': '-'}
            res_b = list_b[0] if list_b else {'date': '-', 'time': '-'}
            return res_a, res_b

        final_max_400, final_max_220 = find_best(cands_max_400, cands_max_220)
        final_min_400, final_min_220 = find_best(cands_min_400, cands_min_220)

        stats['max_400kv_date'] = final_max_400['date']
        stats['max_400kv_time'] = final_max_400['time']
        stats['max_220kv_date'] = final_max_220['date']
        stats['max_220kv_time'] = final_max_220['time']

        stats['min_400kv_date'] = final_min_400['date']
        stats['min_400kv_time'] = final_min_400['time']
        stats['min_220kv_date'] = final_min_220['date']
        stats['min_220kv_time'] = final_min_220['time']
        
        # Cleanup infinities
        if stats['max_400kv'] == -float('inf'): stats['max_400kv'] = '-'
        if stats['min_400kv'] == float('inf'): stats['min_400kv'] = '-'
        if stats['max_220kv'] == -float('inf'): stats['max_220kv'] = '-'
        if stats['min_220kv'] == float('inf'): stats['min_220kv'] = '-'
        if stats['max_load'] == -float('inf'): stats['max_load'] = '-'
        
    else:
        # Feeder / ICT
        # Initialize stats
        stats['max_amps'] = -float('inf'); stats['max_amps_date'] = '-'; stats['max_amps_time'] = '-'
        stats['min_amps'] = float('inf'); stats['min_amps_date'] = '-'; stats['min_amps_time'] = '-'
        stats['max_mw'] = -float('inf'); stats['max_mw_date'] = '-'; stats['max_mw_time'] = '-'
        stats['min_mw'] = float('inf'); stats['min_mw_date'] = '-'; stats['min_mw_time'] = '-'
        
        total_amps = 0; count_amps = 0
        total_mw = 0; count_mw = 0
        
        max_mw_entry = None
        min_mw_entry = None
        
        for e in period_entries:
            d = e.get('data', {})
            date = e['date']
            
            # Collect Averages (Amps)
            max_amps = get_float(d.get('max', {}).get('amps'))
            min_amps = get_float(d.get('min', {}).get('amps'))
            avg_amps = get_float(d.get('avg', {}).get('amps'))
            
            # Auto-calc avg if missing
            if avg_amps is None and max_amps is not None and min_amps is not None:
                avg_amps = (max_amps + min_amps) / 2
            
            if avg_amps is not None:
                total_amps += avg_amps
                count_amps += 1
                
            # Collect Averages (MW) and Find Max/Min MW
            max_mw = get_float(d.get('max', {}).get('mw'))
            min_mw = get_float(d.get('min', {}).get('mw'))
            avg_mw = get_float(d.get('avg', {}).get('mw'))
            
            # Auto-calc avg if missing
            if avg_mw is None and max_mw is not None and min_mw is not None:
                avg_mw = (max_mw + min_mw) / 2
                
            if avg_mw is not None:
                total_mw += avg_mw
                count_mw += 1
            
            # Find Max MW Entry
            if max_mw is not None and max_mw > stats['max_mw']:
                stats['max_mw'] = max_mw
                max_mw_entry = e
            
            # Find Min MW Entry
            if min_mw is not None and min_mw < stats['min_mw']:
                stats['min_mw'] = min_mw
                min_mw_entry = e
        
        # Apply Max MW Logic: Get corresponding Amps/Time from Max MW entry
        if max_mw_entry:
            d = max_mw_entry.get('data', {})
            stats['max_mw_date'] = max_mw_entry['date']
            stats['max_mw_time'] = d.get('max', {}).get('time', '-')
            
            stats['max_amps'] = get_float(d.get('max', {}).get('amps'))
            stats['max_amps_date'] = max_mw_entry['date']
            stats['max_amps_time'] = d.get('max', {}).get('time', '-')
        
        # Apply Min MW Logic: Get corresponding Amps/Time from Min MW entry
        if min_mw_entry:
            d = min_mw_entry.get('data', {})
            stats['min_mw_date'] = min_mw_entry['date']
            stats['min_mw_time'] = d.get('min', {}).get('time', '-')
            
            stats['min_amps'] = get_float(d.get('min', {}).get('amps'))
            stats['min_amps_date'] = min_mw_entry['date']
            stats['min_amps_time'] = d.get('min', {}).get('time', '-')
        
        # Cleanup and averages
        if stats['max_amps'] is None or stats['max_amps'] == -float('inf'): stats['max_amps'] = '-'
        if stats['min_amps'] is None or stats['min_amps'] == float('inf'): stats['min_amps'] = '-'
        stats['avg_amps'] = round(total_amps / count_amps, 2) if count_amps > 0 else '-'
        
        if stats['max_mw'] == -float('inf'): stats['max_mw'] = '-'
        if stats['min_mw'] == float('inf'): stats['min_mw'] = '-'
        stats['avg_mw'] = round(total_mw / count_mw, 2) if count_mw > 0 else '-'
        
    return stats

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
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for entry in entries:
        ws.append([
            format_date(entry['date']),
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
    
@api_router.post("/preview-import/{feeder_id}")
async def preview_import(
    feeder_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user)
):
    feeder = await db.feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only .xlsx or .xls files are supported")
    content = await file.read()
    wb = load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    headers = [str(ws.cell(row=1, column=col).value or "").strip() for col in range(1, ws.max_column + 1)]
    def find_col(end_name: str, kind: str):
        target_tokens = end_name.strip().lower().replace("_", " ").replace("-", " ").replace(".", " ").split()
        for idx, h in enumerate(headers):
            h_low = str(h).strip().lower()
            normalized = h_low.replace("_", " ").replace("-", " ").replace(".", " ")
            tokens = normalized.split()
            if all(t in tokens for t in target_tokens) and kind in tokens and "final" in tokens:
                return idx + 1
        return None
    def find_generic_col(end_label: str, kind: str):
        end_label = end_label.strip().lower()
        for idx, h in enumerate(headers):
            h_low = str(h).strip().lower()
            normalized = h_low.replace("_", " ").replace("-", " ").replace(".", " ")
            tokens = normalized.split()
            if end_label in tokens and kind in tokens and "final" in tokens:
                return idx + 1
        return None
    end1_imp_col = find_col(feeder['end1_name'], "import") or find_generic_col("end1", "import")
    end1_exp_col = find_col(feeder['end1_name'], "export") or find_generic_col("end1", "export")
    end2_imp_col = find_col(feeder['end2_name'], "import") or find_generic_col("end2", "import")
    end2_exp_col = find_col(feeder['end2_name'], "export") or find_generic_col("end2", "export")
    preview = []
    for row in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row, column=1).value
        if not date_cell:
            continue
        if isinstance(date_cell, datetime):
            date_str = date_cell.date().isoformat()
        else:
            val_str = str(date_cell).strip()
            formats = [
                "%Y-%m-%d",
                "%d-%m-%Y",
                "%d-%b-%y",
                "%d-%b-%Y",
                "%d/%m/%Y",
                "%m/%d/%Y",
                "%d.%m.%Y"
            ]
            parsed = False
            for fmt in formats:
                try:
                    date_str = datetime.strptime(val_str, fmt).date().isoformat()
                    parsed = True
                    break
                except:
                    continue
            if not parsed:
                continue
        def get_float(v):
            try:
                return float(v) if v is not None and str(v).strip() != "" else None
            except:
                return None
        e1i = get_float(ws.cell(row=row, column=end1_imp_col).value) if end1_imp_col else None
        e1e = get_float(ws.cell(row=row, column=end1_exp_col).value) if end1_exp_col else None
        e2i = get_float(ws.cell(row=row, column=end2_imp_col).value) if end2_imp_col else None
        e2e = get_float(ws.cell(row=row, column=end2_exp_col).value) if end2_exp_col else None
        if e1i is None and e1e is None and e2i is None and e2e is None:
            continue
        exists = await db.entries.find_one({"feeder_id": feeder_id, "date": date_str})
        preview.append({
            "date": date_str,
            "end1_import_final": e1i,
            "end1_export_final": e1e,
            "end2_import_final": e2i,
            "end2_export_final": e2e,
            "exists": bool(exists)
        })
    return preview

class LineLossesImportPayload(BaseModel):
    feeder_id: str
    entries: List[dict]

@api_router.post("/import-entries")
async def import_entries(payload: LineLossesImportPayload, current_user: User = Depends(get_current_user)):
    feeder = await db.feeders.find_one({"id": payload.feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    entries_sorted = sorted(payload.entries, key=lambda x: x['date'])
    imported = 0
    for e in entries_sorted:
        date_str = e['date']
        existing = await db.entries.find_one({"feeder_id": payload.feeder_id, "date": date_str})
        if existing:
            continue
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        prev_date = (date_obj - timedelta(days=1)).strftime("%Y-%m-%d")
        prev_entry = await db.entries.find_one({"feeder_id": payload.feeder_id, "date": prev_date}, {"_id": 0})
        end1_import_initial = prev_entry['end1_import_final'] if prev_entry else 0
        end1_export_initial = prev_entry['end1_export_final'] if prev_entry else 0
        end2_import_initial = prev_entry['end2_import_final'] if prev_entry else 0
        end2_export_initial = prev_entry['end2_export_final'] if prev_entry else 0
        end1_import_final = float(e.get('end1_import_final', 0) or 0)
        end1_export_final = float(e.get('end1_export_final', 0) or 0)
        end2_import_final = float(e.get('end2_import_final', 0) or 0)
        end2_export_final = float(e.get('end2_export_final', 0) or 0)
        end1_import_consumption = (end1_import_final - end1_import_initial) * feeder['end1_import_mf']
        end1_export_consumption = (end1_export_final - end1_export_initial) * feeder['end1_export_mf']
        end2_import_consumption = (end2_import_final - end2_import_initial) * feeder['end2_import_mf']
        end2_export_consumption = (end2_export_final - end2_export_initial) * feeder['end2_export_mf']
        total_import = end1_import_consumption + end2_import_consumption
        loss_percent = 0 if total_import == 0 else ((end1_import_consumption - end1_export_consumption + end2_import_consumption - end2_export_consumption) / total_import) * 100
        entry_obj = DailyEntry(
            feeder_id=payload.feeder_id,
            date=date_str,
            end1_import_initial=end1_import_initial,
            end1_import_final=end1_import_final,
            end1_export_initial=end1_export_initial,
            end1_export_final=end1_export_final,
            end2_import_initial=end2_import_initial,
            end2_import_final=end2_import_final,
            end2_export_initial=end2_export_initial,
            end2_export_final=end2_export_final,
            end1_import_consumption=end1_import_consumption,
            end1_export_consumption=end1_export_consumption,
            end2_import_consumption=end2_import_consumption,
            end2_export_consumption=end2_export_consumption,
            loss_percent=loss_percent
        ).model_dump()
        entry_obj['created_at'] = entry_obj['created_at'].isoformat() if isinstance(entry_obj['created_at'], datetime) else entry_obj['created_at']
        entry_obj['updated_at'] = entry_obj['updated_at'].isoformat() if isinstance(entry_obj['updated_at'], datetime) else entry_obj['updated_at']
        await db.entries.insert_one(entry_obj)
        imported += 1
    return {"imported": imported}

@api_router.post("/energy/preview-import/{sheet_id}")
async def preview_energy_import(
    sheet_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user)
):
    sheet = await db.energy_sheets.find_one({"id": sheet_id}, {"_id": 0})
    if not sheet:
        raise HTTPException(status_code=404, detail="Sheet not found")
    meters = await db.energy_meters.find({"sheet_id": sheet_id}, {"_id": 0}).to_list(100)
    if not meters:
        raise HTTPException(status_code=400, detail="No meters configured for this sheet")
    
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only .xlsx or .xls files are supported")
    
    content = await file.read()
    wb = load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    
    headers = [str(ws.cell(row=1, column=col).value or "").strip() for col in range(1, ws.max_column + 1)]
    
    def find_final_col(meter_name: str):
        target = meter_name.strip().lower()
        for idx, h in enumerate(headers):
            h_low = h.strip().lower()
            
            # 1. Exact phrase match
            if h_low == f"{target} final":
                return idx + 1
                
            # 2. Token-based match (safer for short names like "IV")
            # Normalize separators to spaces
            normalized = h_low.replace("_", " ").replace("-", " ").replace(".", " ")
            tokens = normalized.split()
            
            if target in tokens and "final" in tokens:
                return idx + 1
                
        return None
    
    meter_final_cols = {}
    for m in meters:
        col = find_final_col(m['name'])
        if col:
            meter_final_cols[m['id']] = col
    
    preview = []
    for row in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row, column=1).value
        if not date_cell:
            continue
        if isinstance(date_cell, datetime):
            date_str = date_cell.date().isoformat()
        else:
            val_str = str(date_cell).strip()
            # Try multiple common date formats
            formats = [
                "%Y-%m-%d",      # 2026-01-01
                "%d-%m-%Y",      # 01-01-2026
                "%d-%b-%y",      # 01-Jan-26
                "%d-%b-%Y",      # 01-Jan-2026
                "%d/%m/%Y",      # 01/01/2026
                "%m/%d/%Y",      # 01/01/2026 (US)
                "%d.%m.%Y"       # 01.01.2026
            ]
            parsed = False
            for fmt in formats:
                try:
                    date_str = datetime.strptime(val_str, fmt).date().isoformat()
                    parsed = True
                    break
                except:
                    continue
            
            if not parsed:
                continue
        
        readings = []
        for meter_id, col in meter_final_cols.items():
            val = ws.cell(row=row, column=col).value
            try:
                final = float(val) if val is not None else None
            except:
                final = None
            if final is not None:
                readings.append({"meter_id": meter_id, "final": final})
        
        if not readings:
            continue
        
        exists = await db.energy_entries.find_one({"sheet_id": sheet_id, "date": date_str})
        preview.append({
            "date": date_str,
            "readings": readings,
            "exists": bool(exists)
        })
    
    return preview

class EnergyImportPayload(BaseModel):
    sheet_id: str
    entries: List[dict]

@api_router.post("/energy/import-entries")
async def import_energy_entries(payload: EnergyImportPayload, current_user: User = Depends(get_current_user)):
    sheet_id = payload.sheet_id
    meters = await db.energy_meters.find({"sheet_id": sheet_id}, {"_id": 0}).to_list(100)
    meter_map = {m['id']: m for m in meters}
    
    entries_sorted = sorted(payload.entries, key=lambda x: x['date'])
    imported = 0
    for e in entries_sorted:
        date_str = e['date']
        existing = await db.energy_entries.find_one({"sheet_id": sheet_id, "date": date_str})
        if existing:
            continue
        
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        prev_date = (date_obj - timedelta(days=1)).strftime("%Y-%m-%d")
        prev_entry = await db.energy_entries.find_one({"sheet_id": sheet_id, "date": prev_date}, {"_id": 0})
        prev_finals = {}
        if prev_entry:
            for r in prev_entry.get('readings', []):
                prev_finals[r['meter_id']] = r['final']
        
        readings = []
        total_consumption = 0.0
        for r in e.get('readings', []):
            meter = meter_map.get(r['meter_id'])
            if not meter:
                continue
            initial = prev_finals.get(r['meter_id'], 0.0)
            final = float(r['final'])
            consumption = (final - initial) * meter['mf']
            readings.append({
                "meter_id": r['meter_id'],
                "initial": initial,
                "final": final,
                "consumption": consumption
            })
            total_consumption += consumption
        
        entry_doc = EnergyEntry(
            sheet_id=sheet_id,
            date=date_str,
            readings=[EnergyReading(**rr) for rr in readings],
            total_consumption=total_consumption
        ).model_dump()
        
        if isinstance(entry_doc.get('created_at'), datetime):
            entry_doc['created_at'] = entry_doc['created_at'].isoformat()
        if isinstance(entry_doc.get('updated_at'), datetime):
            entry_doc['updated_at'] = entry_doc['updated_at'].isoformat()
        
        await db.energy_entries.insert_one(entry_doc)
        imported += 1
    
    return {"imported": imported}

# Max-Min Data Module Endpoints

@api_router.post("/max-min/init")
async def init_max_min_feeders(current_user: User = Depends(get_current_user)):
    existing_count = await db.max_min_feeders.count_documents({})
    if existing_count > 0:
        return {"message": "Max-Min feeders already initialized", "count": existing_count}
    
    feeders_data = [
        {"name": "Bus + Station", "type": "bus_station"},
        {"name": "400KV MAHESHWARAM-1", "type": "feeder_400kv"},
        {"name": "400KV MAHESHWARAM-2", "type": "feeder_400kv"},
        {"name": "400KV NARSAPUR-1", "type": "feeder_400kv"},
        {"name": "400KV NARSAPUR-2", "type": "feeder_400kv"},
        {"name": "400KV KETHIREDDYPALLY-1", "type": "feeder_400kv"},
        {"name": "400KV KETHIREDDYPALLY-2", "type": "feeder_400kv"},
        {"name": "400KV NIZAMABAD-1", "type": "feeder_400kv"},
        {"name": "400KV NIZAMABAD-2", "type": "feeder_400kv"},
        {"name": "220KV PARIGI-1", "type": "feeder_220kv"},
        {"name": "220KV PARIGI-2", "type": "feeder_220kv"},
        {"name": "220KV THANDUR", "type": "feeder_220kv"},
        {"name": "220KV GACHIBOWLI-1", "type": "feeder_220kv"},
        {"name": "220KV GACHIBOWLI-2", "type": "feeder_220kv"},
        {"name": "220KV KETHIREDDYPALLY", "type": "feeder_220kv"},
        {"name": "220KV YEDDUMAILARAM-1", "type": "feeder_220kv"},
        {"name": "220KV YEDDUMAILARAM-2", "type": "feeder_220kv"},
        {"name": "220KV SADASIVAPET-1", "type": "feeder_220kv"},
        {"name": "220KV SADASIVAPET-2", "type": "feeder_220kv"},
        {"name": "ICT-1 (315MVA)", "type": "ict_feeder"},
        {"name": "ICT-2 (315MVA)", "type": "ict_feeder"},
        {"name": "ICT-3 (315MVA)", "type": "ict_feeder"},
        {"name": "ICT-4 (500MVA)", "type": "ict_feeder"},
    ]
    
    for feeder in feeders_data:
        feeder_obj = MaxMinFeeder(**feeder)
        doc = feeder_obj.model_dump()
        doc['created_at'] = doc['created_at'].isoformat()
        await db.max_min_feeders.insert_one(doc)
    
    return {"message": "Max-Min feeders initialized successfully", "count": len(feeders_data)}

@api_router.get("/max-min/feeders", response_model=List[MaxMinFeeder])
async def get_max_min_feeders(current_user: User = Depends(get_current_user)):
    feeders = await db.max_min_feeders.find({}, {"_id": 0}).to_list(100)
    for feeder in feeders:
        if isinstance(feeder.get('created_at'), str):
            feeder['created_at'] = datetime.fromisoformat(feeder['created_at'])
    return feeders

@api_router.get("/max-min/entries/{feeder_id}", response_model=List[MaxMinEntry])
async def get_max_min_entries(
    feeder_id: str,
    year: Optional[int] = None,
    month: Optional[int] = None,
    current_user: User = Depends(get_current_user)
):
    query = {"feeder_id": feeder_id}
    if year and month:
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"
        query['date'] = {"$gte": start_date, "$lt": end_date}
    
    entries = await db.max_min_entries.find(query, {"_id": 0}).sort("date", 1).to_list(1000)
    
    # If it's the Bus + Station page, we might need to calculate Station Load on the fly if it's missing or outdated?
    # But user said "Auto-calculate from entered daily data".
    # Let's trust the stored data for now, but we might need a way to refresh it.
    
    for entry in entries:
        if isinstance(entry.get('created_at'), str):
            entry['created_at'] = datetime.fromisoformat(entry['created_at'])
        if isinstance(entry.get('updated_at'), str):
            entry['updated_at'] = datetime.fromisoformat(entry['updated_at'])
    return entries

@api_router.post("/max-min/entries", response_model=MaxMinEntry)
async def create_max_min_entry(entry_data: MaxMinEntryCreate, current_user: User = Depends(get_current_user)):
    # Check if entry exists
    existing_entry = await db.max_min_entries.find_one(
        {"feeder_id": entry_data.feeder_id, "date": entry_data.date},
        {"_id": 0}
    )
    
    if existing_entry:
        # Update existing
        update_data = {
            "data": entry_data.data,
            "updated_at": datetime.now(timezone.utc).isoformat()
        }
        await db.max_min_entries.update_one(
            {"id": existing_entry['id']},
            {"$set": update_data}
        )
        existing_entry['data'] = entry_data.data
        existing_entry['updated_at'] = update_data['updated_at']
        if isinstance(existing_entry.get('created_at'), str):
            existing_entry['created_at'] = datetime.fromisoformat(existing_entry['created_at'])
        if isinstance(existing_entry.get('updated_at'), str):
            existing_entry['updated_at'] = datetime.fromisoformat(existing_entry['updated_at'])
        return MaxMinEntry(**existing_entry)
    else:
        # Create new
        entry_obj = MaxMinEntry(
            feeder_id=entry_data.feeder_id,
            date=entry_data.date,
            data=entry_data.data
        )
        doc = entry_obj.model_dump()
        doc['created_at'] = doc['created_at'].isoformat()
        doc['updated_at'] = doc['updated_at'].isoformat()
        await db.max_min_entries.insert_one(doc)
        return entry_obj

@api_router.put("/max-min/entries/{entry_id}", response_model=MaxMinEntry)
async def update_max_min_entry(entry_id: str, entry_data: MaxMinEntryCreate, current_user: User = Depends(get_current_user)):
    existing_entry = await db.max_min_entries.find_one({"id": entry_id}, {"_id": 0})
    if not existing_entry:
        raise HTTPException(status_code=404, detail="Entry not found")
    
    update_data = {
        "data": entry_data.data,
        "updated_at": datetime.now(timezone.utc).isoformat()
    }
    
    # Also verify feeder_id and date if needed, but usually we trust the ID
    # However, user might change date in the modal, which is tricky.
    # Frontend sends date in body.
    if existing_entry['date'] != entry_data.date:
        # Check for conflict
        conflict = await db.max_min_entries.find_one(
            {"feeder_id": existing_entry['feeder_id'], "date": entry_data.date},
            {"_id": 0}
        )
        if conflict:
             raise HTTPException(status_code=400, detail="Entry for this date already exists")
        update_data['date'] = entry_data.date

    await db.max_min_entries.update_one(
        {"id": entry_id},
        {"$set": update_data}
    )
    
    updated_entry = await db.max_min_entries.find_one({"id": entry_id}, {"_id": 0})
    if isinstance(updated_entry.get('created_at'), str):
        updated_entry['created_at'] = datetime.fromisoformat(updated_entry['created_at'])
    if isinstance(updated_entry.get('updated_at'), str):
        updated_entry['updated_at'] = datetime.fromisoformat(updated_entry['updated_at'])
    return MaxMinEntry(**updated_entry)

@api_router.delete("/max-min/entries/{entry_id}")
async def delete_max_min_entry(entry_id: str, current_user: User = Depends(get_current_user)):
    result = await db.max_min_entries.delete_one({"id": entry_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Entry not found")
    return {"message": "Entry deleted successfully"}

@api_router.get("/max-min/export/{feeder_id}/{year}/{month}")
async def export_max_min_data(
    feeder_id: str,
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    feeder = await db.max_min_feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        raise HTTPException(status_code=404, detail="Feeder not found")
    
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
        
    entries = await db.max_min_entries.find(
        {"feeder_id": feeder_id, "date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).sort("date", 1).to_list(1000)
    
    wb = Workbook()
    ws = wb.active
    ws.title = feeder['name'][:31]
    
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    # Define headers based on feeder type
    if feeder['type'] == 'bus_station':
        headers = [
            "Date",
            "Max Bus Voltage 400KV Value", "Time",
            "Max Bus Voltage 220KV Value", "Time",
            "Min Bus Voltage 400KV Value", "Time",
            "Min Bus Voltage 220KV Value", "Time",
            "Station Load Max MW", "Station Load MVAR", "Station Load Time"
        ]
    elif feeder['type'] == 'ict_feeder':
        headers = [
            "Date",
            "Max Amps", "Max MW", "Max MVAR", "Max Time",
            "Min Amps", "Min MW", "Min MVAR", "Min Time",
            "Avg Amps", "Avg MW"
        ]
    else: # non-ICT feeders
        headers = [
            "Date",
            "Max Amps", "Max MW", "Max Time",
            "Min Amps", "Min MW", "Min Time",
            "Avg Amps", "Avg MW"
        ]
        
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        
    for entry in entries:
        d = entry.get('data', {})
        row = []
        row.append(format_date(entry['date']))
        
        if feeder['type'] == 'bus_station':
            row.extend([
                d.get('max_bus_voltage_400kv', {}).get('value', ''), format_time(d.get('max_bus_voltage_400kv', {}).get('time', '')),
                d.get('max_bus_voltage_220kv', {}).get('value', ''), format_time(d.get('max_bus_voltage_220kv', {}).get('time', '')),
                d.get('min_bus_voltage_400kv', {}).get('value', ''), format_time(d.get('min_bus_voltage_400kv', {}).get('time', '')),
                d.get('min_bus_voltage_220kv', {}).get('value', ''), format_time(d.get('min_bus_voltage_220kv', {}).get('time', '')),
                d.get('station_load', {}).get('max_mw', ''), d.get('station_load', {}).get('mvar', ''), format_time(d.get('station_load', {}).get('time', ''))
            ])
        elif feeder['type'] == 'ict_feeder':
            row.extend([
                d.get('max', {}).get('amps', ''), d.get('max', {}).get('mw', ''), d.get('max', {}).get('mvar', ''), format_time(d.get('max', {}).get('time', '')),
                d.get('min', {}).get('amps', ''), d.get('min', {}).get('mw', ''), d.get('min', {}).get('mvar', ''), format_time(d.get('min', {}).get('time', '')),
                d.get('avg', {}).get('amps', ''), d.get('avg', {}).get('mw', '')
            ])
        else:
            row.extend([
                d.get('max', {}).get('amps', ''), d.get('max', {}).get('mw', ''), format_time(d.get('max', {}).get('time', '')),
                d.get('min', {}).get('amps', ''), d.get('min', {}).get('mw', ''), format_time(d.get('min', {}).get('time', '')),
                d.get('avg', {}).get('amps', ''), d.get('avg', {}).get('mw', '')
            ])
        ws.append(row)
        
    
    # ---------------------------------------------------------
    # Monthly Summary Section
    # ---------------------------------------------------------
    
    # Define periods
    import calendar
    from openpyxl.styles import Border, Side  # Import here
    
    last_day = calendar.monthrange(year, month)[1]
    
    p1_start = f"{year}-{month:02d}-01"
    p1_end = f"{year}-{month:02d}-15"
    
    p2_start = f"{year}-{month:02d}-16"
    p2_end = f"{year}-{month:02d}-{last_day}"
    
    full_start = f"{year}-{month:02d}-01"
    full_end = f"{year}-{month:02d}-{last_day}"
    
    periods = [
        {"name": "1st to 15th", "start": p1_start, "end": p1_end},
        {"name": "16th to End", "start": p2_start, "end": p2_end},
        {"name": "Full Month", "start": full_start, "end": full_end}
    ]
    
    # Helper to parse float safely
    def get_float(val):
        try:
            return float(val)
        except (ValueError, TypeError):
            return None
            
    # Helper to format time (slice first 5 chars)
    def format_time(t):
        if not t:
            return ''
        s = str(t).strip()
        # Truncate to hh:mm if it looks like a time string
        if len(s) >= 5 and ':' in s:
            return s[:5]
        return s

    # Helper to calculate stats for a period
    def calculate_period_stats(period_entries, feeder_type):
        stats = {}
        
        if feeder_type == 'bus_station':
            # Initialize stats with defaults
            stats = {
                'max_400kv': -float('inf'), 'max_400kv_date': '-', 'max_400kv_time': '-',
                'min_400kv': float('inf'), 'min_400kv_date': '-', 'min_400kv_time': '-',
                'max_220kv': -float('inf'), 'max_220kv_date': '-', 'max_220kv_time': '-',
                'min_220kv': float('inf'), 'min_220kv_date': '-', 'min_220kv_time': '-',
                'max_load': -float('inf'), 'max_load_date': '-', 'max_load_time': '-'
            }

            # 1. Find Max/Min Values
            for e in period_entries:
                d = e.get('data', {})
                
                # 400KV
                v = get_float(d.get('max_bus_voltage_400kv', {}).get('value'))
                if v is not None and v > stats['max_400kv']: stats['max_400kv'] = v
                v = get_float(d.get('min_bus_voltage_400kv', {}).get('value'))
                if v is not None and v < stats['min_400kv']: stats['min_400kv'] = v
                
                # 220KV
                v = get_float(d.get('max_bus_voltage_220kv', {}).get('value'))
                if v is not None and v > stats['max_220kv']: stats['max_220kv'] = v
                v = get_float(d.get('min_bus_voltage_220kv', {}).get('value'))
                if v is not None and v < stats['min_220kv']: stats['min_220kv'] = v

                # Load (Independent)
                v = get_float(d.get('station_load', {}).get('max_mw'))
                if v is not None and v > stats['max_load']:
                    stats['max_load'] = v
                    stats['max_load_date'] = e['date']
                    stats['max_load_time'] = d.get('station_load', {}).get('time', '-')

            # 2. Collect Candidates
            cands_max_400 = []
            cands_min_400 = []
            cands_max_220 = []
            cands_min_220 = []

            for e in period_entries:
                d = e.get('data', {})
                date = e['date']
                
                # Max 400
                v = get_float(d.get('max_bus_voltage_400kv', {}).get('value'))
                if v == stats['max_400kv'] and v is not None:
                    cands_max_400.append({'date': date, 'time': str(d.get('max_bus_voltage_400kv', {}).get('time', '')).strip()})
                
                # Min 400
                v = get_float(d.get('min_bus_voltage_400kv', {}).get('value'))
                if v == stats['min_400kv'] and v is not None:
                    cands_min_400.append({'date': date, 'time': str(d.get('min_bus_voltage_400kv', {}).get('time', '')).strip()})
                
                # Max 220
                v = get_float(d.get('max_bus_voltage_220kv', {}).get('value'))
                if v == stats['max_220kv'] and v is not None:
                    cands_max_220.append({'date': date, 'time': str(d.get('max_bus_voltage_220kv', {}).get('time', '')).strip()})
                
                # Min 220
                v = get_float(d.get('min_bus_voltage_220kv', {}).get('value'))
                if v == stats['min_220kv'] and v is not None:
                    cands_min_220.append({'date': date, 'time': str(d.get('min_bus_voltage_220kv', {}).get('time', '')).strip()})

            # 3. Find Best Matches (Common Time Priority)
            def find_best(list_a, list_b):
                # Try to find intersection
                for a in list_a:
                    for b in list_b:
                        if a['date'] == b['date'] and a['time'] == b['time']:
                            return a, b
                # No intersection, return first available
                res_a = list_a[0] if list_a else {'date': '-', 'time': '-'}
                res_b = list_b[0] if list_b else {'date': '-', 'time': '-'}
                return res_a, res_b

            final_max_400, final_max_220 = find_best(cands_max_400, cands_max_220)
            final_min_400, final_min_220 = find_best(cands_min_400, cands_min_220)

            stats['max_400kv_date'] = final_max_400['date']
            stats['max_400kv_time'] = final_max_400['time']
            stats['max_220kv_date'] = final_max_220['date']
            stats['max_220kv_time'] = final_max_220['time']

            stats['min_400kv_date'] = final_min_400['date']
            stats['min_400kv_time'] = final_min_400['time']
            stats['min_220kv_date'] = final_min_220['date']
            stats['min_220kv_time'] = final_min_220['time']
            
            # Cleanup infinities
            if stats['max_400kv'] == -float('inf'): stats['max_400kv'] = '-'
            if stats['min_400kv'] == float('inf'): stats['min_400kv'] = '-'
            if stats['max_220kv'] == -float('inf'): stats['max_220kv'] = '-'
            if stats['min_220kv'] == float('inf'): stats['min_220kv'] = '-'
            if stats['max_load'] == -float('inf'): stats['max_load'] = '-'
            
        else:
            # Feeder / ICT
            # Initialize stats
            stats['max_amps'] = -float('inf'); stats['max_amps_date'] = '-'; stats['max_amps_time'] = '-'
            stats['min_amps'] = float('inf'); stats['min_amps_date'] = '-'; stats['min_amps_time'] = '-'
            stats['max_mw'] = -float('inf'); stats['max_mw_date'] = '-'; stats['max_mw_time'] = '-'
            stats['min_mw'] = float('inf'); stats['min_mw_date'] = '-'; stats['min_mw_time'] = '-'
            
            total_amps = 0; count_amps = 0
            total_mw = 0; count_mw = 0
            
            max_mw_entry = None
            min_mw_entry = None
            
            for e in period_entries:
                d = e.get('data', {})
                date = e['date']
                
                # Collect Averages (Amps)
                max_amps = get_float(d.get('max', {}).get('amps'))
                min_amps = get_float(d.get('min', {}).get('amps'))
                avg_amps = get_float(d.get('avg', {}).get('amps'))
                
                # Auto-calc avg if missing
                if avg_amps is None and max_amps is not None and min_amps is not None:
                    avg_amps = (max_amps + min_amps) / 2
                
                if avg_amps is not None:
                    total_amps += avg_amps
                    count_amps += 1
                    
                # Collect Averages (MW) and Find Max/Min MW
                max_mw = get_float(d.get('max', {}).get('mw'))
                min_mw = get_float(d.get('min', {}).get('mw'))
                avg_mw = get_float(d.get('avg', {}).get('mw'))
                
                # Auto-calc avg if missing
                if avg_mw is None and max_mw is not None and min_mw is not None:
                    avg_mw = (max_mw + min_mw) / 2
                    
                if avg_mw is not None:
                    total_mw += avg_mw
                    count_mw += 1
                
                # Find Max MW Entry
                if max_mw is not None and max_mw > stats['max_mw']:
                    stats['max_mw'] = max_mw
                    max_mw_entry = e
                
                # Find Min MW Entry
                if min_mw is not None and min_mw < stats['min_mw']:
                    stats['min_mw'] = min_mw
                    min_mw_entry = e
            
            # Apply Max MW Logic: Get corresponding Amps/Time from Max MW entry
            if max_mw_entry:
                d = max_mw_entry.get('data', {})
                stats['max_mw_date'] = max_mw_entry['date']
                stats['max_mw_time'] = d.get('max', {}).get('time', '-')
                
                stats['max_amps'] = get_float(d.get('max', {}).get('amps'))
                stats['max_amps_date'] = max_mw_entry['date']
                stats['max_amps_time'] = d.get('max', {}).get('time', '-')
            
            # Apply Min MW Logic: Get corresponding Amps/Time from Min MW entry
            if min_mw_entry:
                d = min_mw_entry.get('data', {})
                stats['min_mw_date'] = min_mw_entry['date']
                stats['min_mw_time'] = d.get('min', {}).get('time', '-')
                
                stats['min_amps'] = get_float(d.get('min', {}).get('amps'))
                stats['min_amps_date'] = min_mw_entry['date']
                stats['min_amps_time'] = d.get('min', {}).get('time', '-')
            
            # Cleanup and averages
            if stats['max_amps'] is None or stats['max_amps'] == -float('inf'): stats['max_amps'] = '-'
            if stats['min_amps'] is None or stats['min_amps'] == float('inf'): stats['min_amps'] = '-'
            stats['avg_amps'] = round(total_amps / count_amps, 2) if count_amps > 0 else '-'
            
            if stats['max_mw'] == -float('inf'): stats['max_mw'] = '-'
            if stats['min_mw'] == float('inf'): stats['min_mw'] = '-'
            stats['avg_mw'] = round(total_mw / count_mw, 2) if count_mw > 0 else '-'
            
        return stats

    # Add spacing
    ws.append([])
    ws.append([])
    
    # Summary Header
    summary_header_row = ws.max_row + 1
    ws.cell(row=summary_header_row, column=1, value="Monthly Summary Report")
    ws.cell(row=summary_header_row, column=1).font = Font(bold=True, size=14)
    
    # Summary Table Headers
    header_row = ws.max_row + 1
    
    if feeder['type'] == 'bus_station':
        summary_headers = [
            "Period", "Parameter", 
            "Maximum Value", "Date", "Time",
            "Minimum Value", "Date", "Time"
        ]
    else:
        summary_headers = [
            "Period", "Parameter", 
            "Maximum Value", "Date", "Time",
            "Minimum Value", "Date", "Time",
            "Average Value"
        ]
    
    for col_idx, header in enumerate(summary_headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # Add border
        cell.border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )

    # Populate Summary Data
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    
    for p in periods:
        # Filter entries for period
        p_entries = [e for e in entries if p['start'] <= e['date'] <= p['end']]
        stats = calculate_period_stats(p_entries, feeder['type'])
        
        start_row = ws.max_row + 1
        
        if feeder['type'] == 'bus_station':
            # Rows: 400KV, 220KV, Station Load
            rows_data = [
                ("400KV Bus Voltage", stats['max_400kv'], format_date(stats['max_400kv_date']), format_time(stats['max_400kv_time']), stats['min_400kv'], format_date(stats['min_400kv_date']), format_time(stats['min_400kv_time'])),
                ("220KV Bus Voltage", stats['max_220kv'], format_date(stats['max_220kv_date']), format_time(stats['max_220kv_time']), stats['min_220kv'], format_date(stats['min_220kv_date']), format_time(stats['min_220kv_time'])),
                ("Station Load (MW)", stats['max_load'], format_date(stats['max_load_date']), format_time(stats['max_load_time']), "-", "-", "-")
            ]
            
            # Merge Period Cell
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+2, end_column=1)
            p_cell = ws.cell(row=start_row, column=1, value=p['name'])
            p_cell.alignment = center_align
            p_cell.border = thin_border
            # Apply border to merged cells (needs to be applied to all cells in range or use style)
            # Simplest is to apply to top-left and ensure others are handled, but openpyxl requires iterating range for borders
            for r in range(start_row, start_row+3):
                ws.cell(row=r, column=1).border = thin_border
            
            for i, (param, max_v, max_d, max_t, min_v, min_d, min_t) in enumerate(rows_data):
                r = start_row + i
                ws.cell(row=r, column=2, value=param).border = thin_border
                ws.cell(row=r, column=3, value=max_v).border = thin_border
                ws.cell(row=r, column=3).alignment = center_align
                ws.cell(row=r, column=4, value=max_d).border = thin_border
                ws.cell(row=r, column=4).alignment = center_align
                ws.cell(row=r, column=5, value=max_t).border = thin_border
                ws.cell(row=r, column=5).alignment = center_align
                ws.cell(row=r, column=6, value=min_v).border = thin_border
                ws.cell(row=r, column=6).alignment = center_align
                ws.cell(row=r, column=7, value=min_d).border = thin_border
                ws.cell(row=r, column=7).alignment = center_align
                ws.cell(row=r, column=8, value=min_t).border = thin_border
                ws.cell(row=r, column=8).alignment = center_align

        else:
            # Rows: Amps, MW
            rows_data = [
                ("Amps", stats['max_amps'], format_date(stats['max_amps_date']), format_time(stats['max_amps_time']), stats['min_amps'], format_date(stats['min_amps_date']), format_time(stats['min_amps_time']), stats['avg_amps']),
                ("MW", stats['max_mw'], format_date(stats['max_mw_date']), format_time(stats['max_mw_time']), stats['min_mw'], format_date(stats['min_mw_date']), format_time(stats['min_mw_time']), stats['avg_mw'])
            ]
            
            # Merge Period Cell
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+1, end_column=1)
            p_cell = ws.cell(row=start_row, column=1, value=p['name'])
            p_cell.alignment = center_align
            p_cell.border = thin_border
            for r in range(start_row, start_row+2):
                ws.cell(row=r, column=1).border = thin_border
                
            for i, (param, max_v, max_d, max_t, min_v, min_d, min_t, avg_v) in enumerate(rows_data):
                r = start_row + i
                ws.cell(row=r, column=2, value=param).border = thin_border
                ws.cell(row=r, column=3, value=max_v).border = thin_border
                ws.cell(row=r, column=3).alignment = center_align
                ws.cell(row=r, column=4, value=max_d).border = thin_border
                ws.cell(row=r, column=4).alignment = center_align
                ws.cell(row=r, column=5, value=max_t).border = thin_border
                ws.cell(row=r, column=5).alignment = center_align
                ws.cell(row=r, column=6, value=min_v).border = thin_border
                ws.cell(row=r, column=6).alignment = center_align
                ws.cell(row=r, column=7, value=min_d).border = thin_border
                ws.cell(row=r, column=7).alignment = center_align
                ws.cell(row=r, column=8, value=min_t).border = thin_border
                ws.cell(row=r, column=8).alignment = center_align
                ws.cell(row=r, column=9, value=avg_v).border = thin_border
                ws.cell(row=r, column=9).alignment = center_align
    
    # Auto-width columns with wrapping support
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        
        # Calculate max length based on DATA only (skip header at row 1)
        for cell in col[1:]:
            try:
                val = str(cell.value or "")
                if len(val) > max_length:
                    max_length = len(val)
            except:
                pass
        
        # Determine width: minimal 8, maximal 20, or fit data
        # We ignore header length because we enabled wrap_text for headers
        adjusted_width = max(8, min(max_length + 2, 20))
        ws.column_dimensions[column].width = adjusted_width
                
    # Adjust widths for summary section if needed
    # We already adjusted for data, but summary might be wider?
    # Actually, headers are usually wide enough.
    
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

@api_router.get("/line-losses/export-all/{year}/{month}")
async def export_all_line_losses(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    feeders = await db.feeders.find({}, {"_id": 0}).to_list(100)
    
    # Sort feeders based on predefined order
    FEEDER_ORDER = [
        "400KV MAHESHWARAM-2",
        "400KV MAHESHWARAM-1",
        "400KV NARSAPUR-1",
        "400KV NARSAPUR-2",
        "400KV KETHIREDDYPALLY-1",
        "400KV KETHIREDDYPALLY-2",
        "400KV NIZAMABAD-1",
        "400KV NIZAMABAD-2",
        "220KV PARIGI-1",
        "220KV PARIGI-2",
        "220KV THANDUR",
        "220KV GACHIBOWLI-1",
        "220KV GACHIBOWLI-2",
        "220KV KETHIREDDYPALLY",
        "220KV YEDDUMAILARAM-1",
        "220KV YEDDUMAILARAM-2",
        "220KV SADASIVAPET-1",
        "220KV SADASIVAPET-2"
    ]
    
    feeders.sort(key=lambda x: FEEDER_ORDER.index(x['name']) if x['name'] in FEEDER_ORDER else 999)
    
    wb = Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
        
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
        
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    if not feeders:
        ws = wb.create_sheet("No Data")
        ws.cell(row=1, column=1, value="No feeders found")
        
    for feeder in feeders:
        ws = wb.create_sheet(title=feeder['name'][:30]) # Sheet name limit 31 chars
        
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
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        entries = await db.entries.find(
            {"feeder_id": feeder['id'], "date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).sort("date", 1).to_list(1000)
        
        for entry in entries:
            ws.append([
                format_date(entry['date']),
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
            
        # Auto-fit columns with wrap text logic
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            # Skip header row for width calculation to avoid overly wide columns
            for cell in col[1:]:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            # Set width based on data content with a minimum and maximum limit
            # This allows headers to wrap if they are long
            adjusted_width = max(max_length + 2, 12)
            ws.column_dimensions[column].width = adjusted_width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Line_Losses_{month}-{year}.xlsx"}
    )


@api_router.get("/max-min/export-all/{year}/{month}")
async def export_all_max_min_data(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    # Fetch all feeders
    feeders = await db.max_min_feeders.find({}, {"_id": 0}).to_list(100)
    
    # Sort feeders based on predefined order
    FEEDER_ORDER = [
        "Bus Voltages & Station Load",
        "400KV MAHESHWARAM-2",
        "400KV MAHESHWARAM-1",
        "400KV NARSAPUR-1",
        "400KV NARSAPUR-2",
        "400KV KETHIREDDYPALLY-1",
        "400KV KETHIREDDYPALLY-2",
        "400KV NIZAMABAD-1",
        "400KV NIZAMABAD-2",
        "ICT-1 (315MVA)",
        "ICT-2 (315MVA)",
        "ICT-3 (315MVA)",
        "ICT-4 (500MVA)",
        "220KV PARIGI-1",
        "220KV PARIGI-2",
        "220KV THANDUR",
        "220KV GACHIBOWLI-1",
        "220KV GACHIBOWLI-2",
        "220KV KETHIREDDYPALLY",
        "220KV YEDDUMAILARAM-1",
        "220KV YEDDUMAILARAM-2",
        "220KV SADASIVAPET-1",
        "220KV SADASIVAPET-2"
    ]
    
    feeders.sort(key=lambda x: FEEDER_ORDER.index(x['name']) if x['name'] in FEEDER_ORDER else 999)
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"All Feeders {month}-{year}"
    
    # Styles
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    current_col = 1
    
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
    
    import calendar
    last_day = calendar.monthrange(year, month)[1]

    for feeder in feeders:
        entries = await db.max_min_entries.find(
            {"feeder_id": feeder['id'], "date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).sort("date", 1).to_list(1000)
        
        # Headers
        if feeder['type'] == 'bus_station':
            headers = [
                "Date",
                "Max Bus Voltage 400KV Value", "Time",
                "Max Bus Voltage 220KV Value", "Time",
                "Min Bus Voltage 400KV Value", "Time",
                "Min Bus Voltage 220KV Value", "Time",
                "Station Load Max MW", "Station Load MVAR", "Station Load Time"
            ]
        elif feeder['type'] == 'ict_feeder':
            headers = ["Date", "Max Amps", "Max MW", "Max MVAR", "Max Time", "Min Amps", "Min MW", "Min MVAR", "Min Time", "Avg Amps", "Avg MW"]
        else:
            headers = ["Date", "Max Amps", "Max MW", "Max Time", "Min Amps", "Min MW", "Min Time", "Avg Amps", "Avg MW"]
            
        # Write Feeder Name
        ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + len(headers) - 1)
        cell = ws.cell(row=1, column=current_col, value=feeder['name'])
        cell.alignment = center_align
        cell.font = Font(bold=True, size=12)
        cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        for i in range(len(headers)):
            ws.cell(row=1, column=current_col+i).border = thin_border
        
        # Write Headers
        for i, h in enumerate(headers):
            cell = ws.cell(row=2, column=current_col + i, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            
        # Write Data
        row_idx = 3
        for entry in entries:
            d = entry.get('data', {})
            row_data = [format_date(entry['date'])]
            if feeder['type'] == 'bus_station':
                row_data.extend([
                    d.get('max_bus_voltage_400kv', {}).get('value', ''), format_time(d.get('max_bus_voltage_400kv', {}).get('time', '')),
                    d.get('max_bus_voltage_220kv', {}).get('value', ''), format_time(d.get('max_bus_voltage_220kv', {}).get('time', '')),
                    d.get('min_bus_voltage_400kv', {}).get('value', ''), format_time(d.get('min_bus_voltage_400kv', {}).get('time', '')),
                    d.get('min_bus_voltage_220kv', {}).get('value', ''), format_time(d.get('min_bus_voltage_220kv', {}).get('time', '')),
                    d.get('station_load', {}).get('max_mw', ''), d.get('station_load', {}).get('mvar', ''), format_time(d.get('station_load', {}).get('time', ''))
                ])
            elif feeder['type'] == 'ict_feeder':
                row_data.extend([
                    d.get('max', {}).get('amps', ''), d.get('max', {}).get('mw', ''), d.get('max', {}).get('mvar', ''), format_time(d.get('max', {}).get('time', '')),
                    d.get('min', {}).get('amps', ''), d.get('min', {}).get('mw', ''), d.get('min', {}).get('mvar', ''), format_time(d.get('min', {}).get('time', '')),
                    d.get('avg', {}).get('amps', ''), d.get('avg', {}).get('mw', '')
                ])
            else:
                row_data.extend([
                    d.get('max', {}).get('amps', ''), d.get('max', {}).get('mw', ''), format_time(d.get('max', {}).get('time', '')),
                    d.get('min', {}).get('amps', ''), d.get('min', {}).get('mw', ''), format_time(d.get('min', {}).get('time', '')),
                    d.get('avg', {}).get('amps', ''), d.get('avg', {}).get('mw', '')
                ])
            
            for i, val in enumerate(row_data):
                cell = ws.cell(row=row_idx, column=current_col + i, value=val)
                cell.alignment = center_align
                cell.border = thin_border
            row_idx += 1
            
        # Write Stats (Summary)
        row_idx += 1
        periods = [
            {"name": "1st to 15th", "start": f"{year}-{month:02d}-01", "end": f"{year}-{month:02d}-15"},
            {"name": "16th to End", "start": f"{year}-{month:02d}-16", "end": f"{year}-{month:02d}-{last_day}"},
            {"name": "Full Month", "start": f"{year}-{month:02d}-01", "end": f"{year}-{month:02d}-{last_day}"}
        ]
        
        for p in periods:
            p_entries = [e for e in entries if p['start'] <= e['date'] <= p['end']]
            stats = calculate_period_stats(p_entries, feeder['type'])
            
            # Header for Period
            ws.merge_cells(start_row=row_idx, start_column=current_col, end_row=row_idx, end_column=current_col + len(headers) - 1)
            cell = ws.cell(row=row_idx, column=current_col, value=f"Summary: {p['name']}")
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            for i in range(len(headers)): ws.cell(row=row_idx, column=current_col+i).border = thin_border
            row_idx += 1
            
            if feeder['type'] == 'bus_station':
                # Write stats rows
                labels = ["Max 400KV", "Min 400KV", "Max 220KV", "Min 220KV", "Max Load"]
                vals = [
                    (stats['max_400kv'], format_date(stats['max_400kv_date']), format_time(stats['max_400kv_time'])),
                    (stats['min_400kv'], format_date(stats['min_400kv_date']), format_time(stats['min_400kv_time'])),
                    (stats['max_220kv'], format_date(stats['max_220kv_date']), format_time(stats['max_220kv_time'])),
                    (stats['min_220kv'], format_date(stats['min_220kv_date']), format_time(stats['min_220kv_time'])),
                    (stats['max_load'], format_date(stats['max_load_date']), format_time(stats['max_load_time']))
                ]
                for label, val_tuple in zip(labels, vals):
                    ws.cell(row=row_idx, column=current_col, value=label).border = thin_border
                    ws.cell(row=row_idx, column=current_col+1, value=val_tuple[0]).border = thin_border
                    ws.cell(row=row_idx, column=current_col+2, value=val_tuple[1]).border = thin_border
                    ws.cell(row=row_idx, column=current_col+3, value=val_tuple[2]).border = thin_border
                    row_idx += 1
            else:
                labels = ["Max Amps", "Min Amps", "Max MW", "Min MW"]
                vals = [
                    (stats['max_amps'], format_date(stats['max_amps_date']), format_time(stats['max_amps_time'])),
                    (stats['min_amps'], format_date(stats['min_amps_date']), format_time(stats['min_amps_time'])),
                    (stats['max_mw'], format_date(stats['max_mw_date']), format_time(stats['max_mw_time'])),
                    (stats['min_mw'], format_date(stats['min_mw_date']), format_time(stats['min_mw_time']))
                ]
                for label, val_tuple in zip(labels, vals):
                    ws.cell(row=row_idx, column=current_col, value=label).border = thin_border
                    ws.cell(row=row_idx, column=current_col+1, value=val_tuple[0]).border = thin_border
                    ws.cell(row=row_idx, column=current_col+2, value=val_tuple[1]).border = thin_border
                    ws.cell(row=row_idx, column=current_col+3, value=val_tuple[2]).border = thin_border
                    row_idx += 1
            row_idx += 1 # Gap between periods
            
        # Auto-fit columns
        for i in range(len(headers)):
            col_letter = ws.cell(row=2, column=current_col+i).column_letter
            max_len = 0
            for row in range(1, row_idx):
                val = ws.cell(row=row, column=current_col+i).value
                if val:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max_len + 2
            
        current_col += len(headers) + 1 # Gap between feeders

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=MaxMin_All_{year}_{month:02d}.xlsx"}
    )

@api_router.get("/energy/export-all/{year}/{month}")
async def export_all_energy_sheets(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    sheets = await db.energy_sheets.find({}, {"_id": 0}).to_list(100)
    sheets.sort(key=lambda x: x['name'])
    
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
        
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
        
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for sheet in sheets:
        ws = wb.create_sheet(title=sheet['name'])
        
        meters = await db.energy_meters.find({"sheet_id": sheet['id']}, {"_id": 0}).to_list(100)
        
        entries = await db.energy_entries.find(
            {"sheet_id": sheet['id'], "date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).sort("date", 1).to_list(1000)
        
        headers = ["Date"]
        for m in meters:
            headers.extend([
                f"{m['name']} Initial",
                f"{m['name']} Final",
                f"{m['name']} MF",
                f"{m['name']} Consumption"
            ])
        headers.append("Total Consumption")
        
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        for entry in entries:
            row = [format_date(entry['date'])]
            readings_map = {r['meter_id']: r for r in entry['readings']}
            
            for m in meters:
                r = readings_map.get(m['id'])
                if r:
                    row.extend([r['initial'], r['final'], m['mf'], r['consumption']])
                else:
                    row.extend(['-', '-', m['mf'], '-'])
            
            row.append(entry['total_consumption'])
            ws.append(row)
            
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Energy_Consumption_{month}-{year}.xlsx"}
    )

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
        row = [format_date(entry['date'])]
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

    # Update next day's initial values if exists
    date_obj = datetime.strptime(entry_input.date, "%Y-%m-%d")
    next_date = (date_obj + timedelta(days=1)).strftime("%Y-%m-%d")
    next_entry = await db.energy_entries.find_one(
        {"sheet_id": entry_input.sheet_id, "date": next_date}
    )
    
    if next_entry:
        # Create a map of current finals
        current_finals = {r.meter_id: r.final for r in readings}
        
        next_readings = []
        next_total_consumption = 0
        
        for r in next_entry['readings']:
            if r['meter_id'] in current_finals:
                new_initial = current_finals[r['meter_id']]
                
                # Fetch meter to get MF
                meter = await db.energy_meters.find_one({"id": r['meter_id']})
                if meter:
                    new_consumption = (r['final'] - new_initial) * meter['mf']
                    r['initial'] = new_initial
                    r['consumption'] = new_consumption
            
            next_readings.append(r)
            next_total_consumption += r.get('consumption', 0)
            
        next_entry['readings'] = next_readings
        next_entry['total_consumption'] = next_total_consumption
        next_entry['updated_at'] = datetime.now(timezone.utc).isoformat()
        
        await db.energy_entries.replace_one({"_id": next_entry['_id']}, next_entry)
        
    return entry_data

@api_router.delete("/energy/entries/{entry_id}")
async def delete_energy_entry(entry_id: str, current_user: User = Depends(get_current_user)):
    result = await db.energy_entries.delete_one({"id": entry_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Entry not found")
    return {"message": "Entry deleted successfully"}


app.include_router(api_router)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()
