from fastapi import FastAPI, APIRouter, HTTPException, Depends, status, UploadFile, File
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.responses import StreamingResponse, JSONResponse
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
import calendar
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import random
import string
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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

def send_reports_email(recipient_email: str, attachments: list):
    try:
        sender_email = os.environ.get("SMTP_EMAIL")
        sender_password = os.environ.get("SMTP_PASSWORD")

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = recipient_email
        message["Subject"] = "MIS Portal Reports"

        body = "Please find the attached reports for the selected period."
        message.attach(MIMEText(body, "plain"))

        for filename, content in attachments:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(content)
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {filename}",
            )
            message.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
    except Exception as e:
        print(f"Error sending reports email: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to send email: {str(e)}")

mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

app = FastAPI()

@app.on_event("startup")
async def startup_db_client():
    # Create indexes for performance
    try:
        # Users
        await db.users.create_index("email", unique=True)
        await db.users.create_index("id", unique=True)
        
        # Feeders
        await db.feeders.create_index("id", unique=True)
        
        # Entries
        await db.entries.create_index("id", unique=True)
        await db.entries.create_index([("feeder_id", 1), ("date", 1)])
        await db.entries.create_index("date")
        
        # Max Min Feeders
        await db.max_min_feeders.create_index("id", unique=True)
        
        # Max Min Entries
        await db.max_min_entries.create_index("id", unique=True)
        await db.max_min_entries.create_index([("feeder_id", 1), ("date", 1)])
        
        print("Database indexes created successfully")
    except Exception as e:
        print(f"Error creating indexes: {e}")


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

class EmailReportRequest(BaseModel):
    email: str
    year: int
    month: int
    report_ids: Optional[List[str]] = None

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
            d = e.get('data') or {}
            
            # 400KV
            v = get_float((d.get('max_bus_voltage_400kv') or {}).get('value'))
            if v is not None and v > stats['max_400kv']: stats['max_400kv'] = v
            v = get_float((d.get('min_bus_voltage_400kv') or {}).get('value'))
            if v is not None and v < stats['min_400kv']: stats['min_400kv'] = v
            
            # 220KV
            v = get_float((d.get('max_bus_voltage_220kv') or {}).get('value'))
            if v is not None and v > stats['max_220kv']: stats['max_220kv'] = v
            v = get_float((d.get('min_bus_voltage_220kv') or {}).get('value'))
            if v is not None and v < stats['min_220kv']: stats['min_220kv'] = v

            # Load (Independent)
            v = get_float((d.get('station_load') or {}).get('max_mw'))
            if v is not None and v > stats['max_load']:
                stats['max_load'] = v
                stats['max_load_date'] = e['date']
                stats['max_load_time'] = (d.get('station_load') or {}).get('time', '-')
                stats['max_load_mvar'] = (d.get('station_load') or {}).get('mvar', '-')

        # 2. Collect Candidates
        cands_max_400 = []
        cands_min_400 = []
        cands_max_220 = []
        cands_min_220 = []

        for e in period_entries:
            d = e.get('data') or {}
            date = e['date']
            
            # Max 400
            v = get_float((d.get('max_bus_voltage_400kv') or {}).get('value'))
            if v == stats['max_400kv'] and v is not None:
                cands_max_400.append({'date': date, 'time': str((d.get('max_bus_voltage_400kv') or {}).get('time', '')).strip()})
            
            # Min 400
            v = get_float((d.get('min_bus_voltage_400kv') or {}).get('value'))
            if v == stats['min_400kv'] and v is not None:
                cands_min_400.append({'date': date, 'time': str((d.get('min_bus_voltage_400kv') or {}).get('time', '')).strip()})
            
            # Max 220
            v = get_float((d.get('max_bus_voltage_220kv') or {}).get('value'))
            if v == stats['max_220kv'] and v is not None:
                cands_max_220.append({'date': date, 'time': str((d.get('max_bus_voltage_220kv') or {}).get('time', '')).strip()})
            
            # Min 220
            v = get_float((d.get('min_bus_voltage_220kv') or {}).get('value'))
            if v == stats['min_220kv'] and v is not None:
                cands_min_220.append({'date': date, 'time': str((d.get('min_bus_voltage_220kv') or {}).get('time', '')).strip()})

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
            d = e.get('data') or {}
            date = e['date']
            
            # Collect Averages (Amps)
            max_amps = get_float((d.get('max') or {}).get('amps'))
            min_amps = get_float((d.get('min') or {}).get('amps'))
            avg_amps = get_float((d.get('avg') or {}).get('amps'))
            
            # Auto-calc avg if missing
            if avg_amps is None and max_amps is not None and min_amps is not None:
                avg_amps = (max_amps + min_amps) / 2
            
            if avg_amps is not None:
                total_amps += avg_amps
                count_amps += 1
                
            # Collect Averages (MW) and Find Max/Min MW
            max_mw = get_float((d.get('max') or {}).get('mw'))
            min_mw = get_float((d.get('min') or {}).get('mw'))
            avg_mw = get_float((d.get('avg') or {}).get('mw'))
            
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
            d = max_mw_entry.get('data') or {}
            stats['max_mw_date'] = max_mw_entry['date']
            stats['max_mw_time'] = (d.get('max') or {}).get('time', '-')
            
            stats['max_amps'] = get_float((d.get('max') or {}).get('amps'))
            stats['max_amps_date'] = max_mw_entry['date']
            stats['max_amps_time'] = (d.get('max') or {}).get('time', '-')
        
        # Apply Min MW Logic: Get corresponding Amps/Time from Min MW entry
        if min_mw_entry:
            d = min_mw_entry.get('data') or {}
            stats['min_mw_date'] = min_mw_entry['date']
            stats['min_mw_time'] = (d.get('min') or {}).get('time', '-')
            
            stats['min_amps'] = get_float((d.get('min') or {}).get('amps'))
            stats['min_amps_date'] = min_mw_entry['date']
            stats['min_amps_time'] = (d.get('min') or {}).get('time', '-')
        
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

# ---------------------------------------------------------
# Max-Min Summary Logic Helpers
# ---------------------------------------------------------

DOUBLE_CIRCUIT_PAIRS = [
    {"names": ["400KV MAHESHWARAM-2", "400KV MAHESHWARAM-1"]},
    {"names": ["400KV NARSAPUR-1", "400KV NARSAPUR-2"]},
    {"names": ["400KV KETHIREDDYPALLY-1", "400KV KETHIREDDYPALLY-2"]},
    {"names": ["400KV NIZAMABAD-1", "400KV NIZAMABAD-2"]},
    {"names": ["220KV PARIGI-1", "220KV PARIGI-2"]},
    {"names": ["220KV GACHIBOWLI-1", "220KV GACHIBOWLI-2"]},
    {"names": ["220KV YEDDUMAILARAM-1", "220KV YEDDUMAILARAM-2"]},
    {"names": ["220KV SADASIVAPET-1", "220KV SADASIVAPET-2"]}
]
ICT_GROUP_NAMES = ["ICT-1 (315MVA)", "ICT-2 (315MVA)", "ICT-3 (315MVA)", "ICT-4 (500MVA)"]

def normalize_time(t):
    if not t: return ""
    s = str(t).strip()
    
    # Handle "YYYY-MM-DD HH:MM:SS" or similar
    if ' ' in s:
        s = s.split(' ')[-1]
        
    # Handle HH:MM:SS vs HH:MM
    if ':' in s:
        parts = s.split(':')
        if len(parts) >= 2:
            try:
                return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
            except ValueError:
                pass
    return s

def get_float_safe(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def get_feeder_group_info(feeder_name):
    # Check Double Circuit
    for pair in DOUBLE_CIRCUIT_PAIRS:
        if feeder_name in pair['names']:
            return True, [n for n in pair['names'] if n != feeder_name]
    
    # Check ICT
    if feeder_name in ICT_GROUP_NAMES:
        return True, [n for n in ICT_GROUP_NAMES if n != feeder_name]
        
    return False, []

def determine_leader(p_group_map):
    global_max_mw = -float('inf')
    candidate_leaders = [] # List of {'fid': fid, 'mw': mw, 'entry': entry}

    for fid, entries in p_group_map.items():
        local_max = -float('inf')
        local_entry = None
        for e in entries:
            mw = get_float_safe(e.get('data', {}).get('max', {}).get('mw'))
            if mw is not None and mw > local_max:
                local_max = mw
                local_entry = e
        
        if local_max > global_max_mw:
            global_max_mw = local_max
            candidate_leaders = [{'fid': fid, 'mw': local_max, 'entry': local_entry}]
        elif local_max == global_max_mw and local_max > -float('inf'):
            candidate_leaders.append({'fid': fid, 'mw': local_max, 'entry': local_entry})
    
    leader_id = None
    
    if not candidate_leaders:
         # Fallback if no data
         return list(p_group_map.keys())[0] if p_group_map else None

    if len(candidate_leaders) == 1:
        leader_id = candidate_leaders[0]['fid']
    elif len(candidate_leaders) > 1:
        # Sort candidates by ID for deterministic processing order (crucial for ties)
        candidate_leaders.sort(key=lambda x: str(x['fid']))

        # Tie Breaker: Calculate Sum of MW at the candidate's timestamp
        best_sum = -float('inf')
        best_leader = None
        
        for cand in candidate_leaders:
            c_fid = cand['fid']
            c_entry = cand['entry']
            c_date = c_entry['date']
            c_time = normalize_time(c_entry.get('data', {}).get('max', {}).get('time'))
            
            current_sum = 0
            valid_timestamp = True
            
            # Sum all partners at this timestamp
            for pid, pentries in p_group_map.items():
                p_entry = next((e for e in pentries if e['date'] == c_date), None)
                if not p_entry:
                    valid_timestamp = False
                    break
                
                p_time = normalize_time(p_entry.get('data', {}).get('max', {}).get('time'))
                if p_time != c_time:
                    valid_timestamp = False
                    break
                    
                val = get_float_safe(p_entry.get('data', {}).get('max', {}).get('mw'))
                if val is not None: current_sum += val
            
            if valid_timestamp and current_sum > best_sum:
                best_sum = current_sum
                best_leader = c_fid
        
        if best_leader:
            leader_id = best_leader
        else:
            # Fallback: Sort by ID for determinism
            candidate_leaders.sort(key=lambda x: x['fid'])
            leader_id = candidate_leaders[0]['fid']
            
    return leader_id

def calculate_standard_stats(period_entries, feeder_type):
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
            v = get_float_safe(d.get('max_bus_voltage_400kv', {}).get('value'))
            if v is not None and v > stats['max_400kv']: stats['max_400kv'] = v
            v = get_float_safe(d.get('min_bus_voltage_400kv', {}).get('value'))
            if v is not None and v < stats['min_400kv']: stats['min_400kv'] = v
            
            # 220KV
            v = get_float_safe(d.get('max_bus_voltage_220kv', {}).get('value'))
            if v is not None and v > stats['max_220kv']: stats['max_220kv'] = v
            v = get_float_safe(d.get('min_bus_voltage_220kv', {}).get('value'))
            if v is not None and v < stats['min_220kv']: stats['min_220kv'] = v

            # Load (Independent)
            v = get_float_safe(d.get('station_load', {}).get('max_mw'))
            if v is not None and v > stats['max_load']:
                stats['max_load'] = v
                stats['max_load_date'] = e['date']
                stats['max_load_time'] = d.get('station_load', {}).get('time', '-')
                stats['max_load_mvar'] = d.get('station_load', {}).get('mvar', '-')

        # 2. Collect Candidates & 3. Find Best Matches (Simplified for brevity as per original logic)
        # For simplicity in this refactor, we will just use the FIRST occurrence for Max/Min if simple comparison
        # But original logic had "Common Time Priority" for voltages.
        # Let's keep it simple or copy full logic? 
        # Copying full logic is safer.
        
        cands_max_400 = []
        cands_min_400 = []
        cands_max_220 = []
        cands_min_220 = []

        for e in period_entries:
            d = e.get('data', {})
            date = e['date']
            
            # Max 400
            v = get_float_safe(d.get('max_bus_voltage_400kv', {}).get('value'))
            if v == stats['max_400kv'] and v is not None:
                cands_max_400.append({'date': date, 'time': str(d.get('max_bus_voltage_400kv', {}).get('time', '')).strip()})
        
            # Min 400
            v = get_float_safe(d.get('min_bus_voltage_400kv', {}).get('value'))
            if v == stats['min_400kv'] and v is not None:
                cands_min_400.append({'date': date, 'time': str(d.get('min_bus_voltage_400kv', {}).get('time', '')).strip()})
        
            # Max 220
            v = get_float_safe(d.get('max_bus_voltage_220kv', {}).get('value'))
            if v == stats['max_220kv'] and v is not None:
                cands_max_220.append({'date': date, 'time': str(d.get('max_bus_voltage_220kv', {}).get('time', '')).strip()})
        
            # Min 220
            v = get_float_safe(d.get('min_bus_voltage_220kv', {}).get('value'))
            if v == stats['min_220kv'] and v is not None:
                cands_min_220.append({'date': date, 'time': str(d.get('min_bus_voltage_220kv', {}).get('time', '')).strip()})

        def find_best(list_a, list_b):
            for a in list_a:
                for b in list_b:
                    if a['date'] == b['date'] and a['time'] == b['time']:
                        return a, b
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
        
        # Cleanup
        if stats['max_400kv'] == -float('inf'): stats['max_400kv'] = '-'
        if stats['min_400kv'] == float('inf'): stats['min_400kv'] = '-'
        if stats['max_220kv'] == -float('inf'): stats['max_220kv'] = '-'
        if stats['min_220kv'] == float('inf'): stats['min_220kv'] = '-'
        if stats['max_load'] == -float('inf'): stats['max_load'] = '-'
        
    else:
        # Feeder / ICT Standard Logic
        stats = {
            'max_amps': -float('inf'), 'max_amps_date': '-', 'max_amps_time': '-',
            'min_amps': float('inf'), 'min_amps_date': '-', 'min_amps_time': '-',
            'max_mw': -float('inf'), 'max_mw_date': '-', 'max_mw_time': '-',
            'min_mw': float('inf'), 'min_mw_date': '-', 'min_mw_time': '-',
            'avg_amps': 0, 'avg_mw': 0
        }
        
        total_amps = 0; count_amps = 0
        total_mw = 0; count_mw = 0
        
        max_mw_entry = None
        min_mw_entry = None
        
        for e in period_entries:
            d = e.get('data', {})
            
            # Avg Amps
            max_a = get_float_safe(d.get('max', {}).get('amps'))
            min_a = get_float_safe(d.get('min', {}).get('amps'))
            avg_a = get_float_safe(d.get('avg', {}).get('amps'))
            if avg_a is None and max_a is not None and min_a is not None: avg_a = (max_a + min_a) / 2
            if avg_a is not None: total_amps += avg_a; count_amps += 1
                
            # Avg MW
            max_m = get_float_safe(d.get('max', {}).get('mw'))
            min_m = get_float_safe(d.get('min', {}).get('mw'))
            avg_m = get_float_safe(d.get('avg', {}).get('mw'))
            if avg_m is None and max_m is not None and min_m is not None: avg_m = (max_m + min_m) / 2
            if avg_m is not None: total_mw += avg_m; count_mw += 1
            
            # Find Max/Min MW Candidates
            if max_m is not None and max_m > stats['max_mw']:
                stats['max_mw'] = max_m
                max_mw_entry = e
            if min_m is not None and min_m < stats['min_mw']:
                stats['min_mw'] = min_m
                min_mw_entry = e
                
        # Fill Stats from Max/Min MW Entries
        if max_mw_entry:
            d = max_mw_entry.get('data', {})
            stats['max_mw_date'] = max_mw_entry['date']
            stats['max_mw_time'] = d.get('max', {}).get('time', '-')
            stats['max_amps'] = get_float_safe(d.get('max', {}).get('amps'))
            stats['max_amps_date'] = max_mw_entry['date']
            stats['max_amps_time'] = d.get('max', {}).get('time', '-')

        if min_mw_entry:
            d = min_mw_entry.get('data', {})
            stats['min_mw_date'] = min_mw_entry['date']
            stats['min_mw_time'] = d.get('min', {}).get('time', '-')
            stats['min_amps'] = get_float_safe(d.get('min', {}).get('amps'))
            stats['min_amps_date'] = min_mw_entry['date']
            stats['min_amps_time'] = d.get('min', {}).get('time', '-')
            
        # Cleanup
        if stats['max_amps'] is None or stats['max_amps'] == -float('inf'): stats['max_amps'] = '-'
        if stats['min_amps'] is None or stats['min_amps'] == float('inf'): stats['min_amps'] = '-'
        stats['avg_amps'] = round(total_amps / count_amps, 2) if count_amps > 0 else '-'
        if stats['max_mw'] == -float('inf'): stats['max_mw'] = '-'
        if stats['min_mw'] == float('inf'): stats['min_mw'] = '-'
        stats['avg_mw'] = round(total_mw / count_mw, 2) if count_mw > 0 else '-'
        
    return stats

def calculate_coincident_stats(leader_entries, current_feeder_id, group_entries_map, feeder_type):
    # This function uses leader_entries to find the "Best Coincident Timestamp"
    # And then returns the stats for the CURRENT feeder at that timestamp.
    
    # Base calculation for Averages (independent of coincidence)
    # We must calculate averages from the CURRENT feeder's entries, not Leader's.
    current_feeder_entries = group_entries_map.get(current_feeder_id, [])
    base_stats = calculate_standard_stats(current_feeder_entries, feeder_type)
    
    # Reset Max/Min fields to be populated by Coincident Logic
    base_stats.update({
        'max_mw': -float('inf'), 'max_mw_date': '-', 'max_mw_time': '-',
        'max_amps': -float('inf'), 'max_amps_date': '-', 'max_amps_time': '-',
        'min_mw': float('inf'), 'min_mw_date': '-', 'min_mw_time': '-',
        'min_amps': float('inf'), 'min_amps_date': '-', 'min_amps_time': '-'
    })

    # 1. Gather Candidates from LEADER
    max_candidates = []
    min_candidates = []
    
    for e in leader_entries:
        d = e.get('data', {})
        mw_max = get_float_safe(d.get('max', {}).get('mw'))
        time_max = normalize_time(d.get('max', {}).get('time'))
        if mw_max is not None:
            max_candidates.append({'entry': e, 'val': mw_max, 'date': e['date'], 'time': time_max})
            
        mw_min = get_float_safe(d.get('min', {}).get('mw'))
        time_min = normalize_time(d.get('min', {}).get('time'))
        if mw_min is not None:
            min_candidates.append({'entry': e, 'val': mw_min, 'date': e['date'], 'time': time_min})
            
    # Sort
    max_candidates.sort(key=lambda x: x['val'], reverse=True)
    min_candidates.sort(key=lambda x: x['val']) # Ascending
    
    # Find Coincident Max Timestamp
    found_max_timestamp = None # {date, time}
    for cand in max_candidates:
        c_date = cand['date']
        c_time = cand['time']
        
        # Check all partners
        all_match = True
        for fid, entries in group_entries_map.items():
            # Find entry for c_date
            partner_entry = next((p for p in entries if p['date'] == c_date), None)
            if not partner_entry:
                all_match = False
                break
            
            # Check time matches
            p_time = normalize_time(partner_entry.get('data', {}).get('max', {}).get('time'))
            if p_time != c_time:
                all_match = False
                break
        
        if all_match:
            found_max_timestamp = {'date': c_date, 'time': c_time}
            break
            
    # Find Coincident Min Timestamp
    found_min_timestamp = None
    for cand in min_candidates:
        c_date = cand['date']
        c_time = cand['time']
        
        all_match = True
        for fid, entries in group_entries_map.items():
            partner_entry = next((p for p in entries if p['date'] == c_date), None)
            if not partner_entry:
                all_match = False
                break
            p_time = normalize_time(partner_entry.get('data', {}).get('min', {}).get('time'))
            if p_time != c_time:
                all_match = False
                break
        
        if all_match:
            found_min_timestamp = {'date': c_date, 'time': c_time}
            break
            
    # Update Stats with CURRENT FEEDER values at found timestamps
    if found_max_timestamp:
        # Find entry for current feeder at this date
        entry = next((e for e in current_feeder_entries if e['date'] == found_max_timestamp['date']), None)
        if entry:
            d = entry.get('data', {})
            base_stats['max_mw'] = get_float_safe(d.get('max', {}).get('mw'))
            base_stats['max_mw_date'] = found_max_timestamp['date']
            base_stats['max_mw_time'] = found_max_timestamp['time'] # Use common time
            
            base_stats['max_amps'] = get_float_safe(d.get('max', {}).get('amps'))
            base_stats['max_amps_date'] = found_max_timestamp['date']
            base_stats['max_amps_time'] = found_max_timestamp['time']
            
            if base_stats['max_mw'] is None: base_stats['max_mw'] = '-'
            if base_stats['max_amps'] is None: base_stats['max_amps'] = '-'
    else:
        base_stats['max_mw'] = '-'
        base_stats['max_mw_date'] = '-'
        base_stats['max_mw_time'] = '-'
        base_stats['max_amps'] = '-'
        base_stats['max_amps_date'] = '-'
        base_stats['max_amps_time'] = '-'

    if found_min_timestamp:
        entry = next((e for e in current_feeder_entries if e['date'] == found_min_timestamp['date']), None)
        if entry:
            d = entry.get('data', {})
            base_stats['min_mw'] = get_float_safe(d.get('min', {}).get('mw'))
            base_stats['min_mw_date'] = found_min_timestamp['date']
            base_stats['min_mw_time'] = found_min_timestamp['time']
            
            base_stats['min_amps'] = get_float_safe(d.get('min', {}).get('amps'))
            base_stats['min_amps_date'] = found_min_timestamp['date']
            base_stats['min_amps_time'] = found_min_timestamp['time']

            if base_stats['min_mw'] is None: base_stats['min_mw'] = '-'
            if base_stats['min_amps'] is None: base_stats['min_amps'] = '-'
    else:
        base_stats['min_mw'] = '-'
        base_stats['min_mw_date'] = '-'
        base_stats['min_mw_time'] = '-'
        base_stats['min_amps'] = '-'
        base_stats['min_amps_date'] = '-'
        base_stats['min_amps_time'] = '-'
        
    return base_stats

@api_router.get("/max-min/summary/{feeder_id}/{year}/{month}")
async def get_monthly_summary(
    feeder_id: str,
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    # 1. Fetch Feeder
    feeder = await db.max_min_feeders.find_one({"id": feeder_id}, {"_id": 0})
    if not feeder:
        return JSONResponse(status_code=404, content={"message": "Feeder not found"})

    # 2. Check if Special Logic Applies
    # is_special = False
    # partner_names = []
    
    # # Check Double Circuit
    # for pair in DOUBLE_CIRCUIT_PAIRS:
    #     if feeder['name'] in pair['names']:
    #         is_special = True
    #         partner_names = [n for n in pair['names'] if n != feeder['name']]
    #         break
            
    # # Check ICT
    # if not is_special and feeder['name'] in ICT_GROUP_NAMES:
    #     is_special = True
    #     partner_names = [n for n in ICT_GROUP_NAMES if n != feeder['name']]

    # if not is_special:
    #     return None

    # 3. Fetch Partner Feeders & Entries
    # partner_feeders = await db.max_min_feeders.find({"name": {"$in": partner_names}}).to_list(100)
    # partner_ids = [p['id'] for p in partner_feeders]
    
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"

    # Fetch Target Entries
    target_entries = await db.max_min_entries.find(
        {"feeder_id": feeder_id, "date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).to_list(1000)
    
    # Fetch Partner Entries
    # group_entries_map = {}
    # if partner_ids:
    #     p_entries = await db.max_min_entries.find(
    #         {"feeder_id": {"$in": partner_ids}, "date": {"$gte": start_date, "$lt": end_date}},
    #         {"_id": 0}
    #     ).to_list(5000)
        
    #     for e in p_entries:
    #         fid = e['feeder_id']
    #         if fid not in group_entries_map: group_entries_map[fid] = []
    #         group_entries_map[fid].append(e)

    # 4. Calculate Stats for Periods
    periods = [
        {"name": "1st to 15th", "start": 1, "end": 15},
        {"name": "16th to End", "start": 16, "end": 31},
        {"name": "Full Month", "start": 1, "end": 31}
    ]
    
    results = []
    
    for p in periods:
        p_target = [
            e for e in target_entries 
            if p["start"] <= int(e['date'].split('-')[2]) <= p["end"]
        ]
        
        # p_group_map = {}
        # # Add current feeder to map
        # p_group_map[feeder_id] = p_target

        # for fid, entries in group_entries_map.items():
        #     p_group_map[fid] = [
        #         e for e in entries
        #         if p["start"] <= int(e['date'].split('-')[2]) <= p["end"]
        #     ]
        
        # # Determine Leader (Feeder with highest Max MW in this period)
        # leader_id = determine_leader(p_group_map)
        # if not leader_id: leader_id = feeder_id

        # leader_entries = p_group_map[leader_id]
            
        # stats = calculate_coincident_stats(leader_entries, feeder_id, p_group_map, feeder['type'])
        stats = calculate_standard_stats(p_target, feeder['type'])
        stats['name'] = p['name']
        results.append(stats)
        
    return results

@api_router.get("/max-min/export/{feeder_id}/{year}/{month}")
async def export_max_min_data(
    feeder_id: str,
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        import io
        import traceback
        from fastapi.responses import JSONResponse

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
                "Max Bus Voltage 400KV", "Max Bus Voltage 220KV", "Max Value Time",
                "Min Bus Voltage 400KV", "Min Bus Voltage 220KV", "Min Value Time",
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
                # Determine common time for Max Voltages
                max_time = d.get('max_bus_voltage', {}).get('time') or d.get('max_bus_voltage_400kv', {}).get('time') or d.get('max_bus_voltage_220kv', {}).get('time')
                
                # Determine common time for Min Voltages
                min_time = d.get('min_bus_voltage', {}).get('time') or d.get('min_bus_voltage_400kv', {}).get('time') or d.get('min_bus_voltage_220kv', {}).get('time')
                
                row.extend([
                    d.get('max_bus_voltage_400kv', {}).get('value', ''), 
                    d.get('max_bus_voltage_220kv', {}).get('value', ''), 
                    format_time(max_time),
                    
                    d.get('min_bus_voltage_400kv', {}).get('value', ''), 
                    d.get('min_bus_voltage_220kv', {}).get('value', ''), 
                    format_time(min_time),
                    
                    d.get('station_load', {}).get('max_mw', ''), 
                    d.get('station_load', {}).get('mvar', ''), 
                    format_time(d.get('station_load', {}).get('time', ''))
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
        def format_time_local(t):
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
                    ("400KV Bus Voltage", stats['max_400kv'], format_date(stats['max_400kv_date']), format_time_local(stats['max_400kv_time']), stats['min_400kv'], format_date(stats['min_400kv_date']), format_time_local(stats['min_400kv_time'])),
                    ("220KV Bus Voltage", stats['max_220kv'], format_date(stats['max_220kv_date']), format_time_local(stats['max_220kv_time']), stats['min_220kv'], format_date(stats['min_220kv_date']), format_time_local(stats['min_220kv_time'])),
                    ("Station Load (MW)", stats['max_load'], format_date(stats['max_load_date']), format_time_local(stats['max_load_time']), "-", "-", "-")
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
                    ("Amps", stats['max_amps'], format_date(stats['max_amps_date']), format_time_local(stats['max_amps_time']), stats['min_amps'], format_date(stats['min_amps_date']), format_time_local(stats['min_amps_time']), stats['avg_amps']),
                    ("MW", stats['max_mw'], format_date(stats['max_mw_date']), format_time_local(stats['max_mw_time']), stats['min_mw'], format_date(stats['min_mw_date']), format_time_local(stats['min_mw_time']), stats['avg_mw'])
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
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)},
            headers={
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Credentials": "true",
            }
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

async def _generate_fortnight_report_wb(year: int, month: int):
    try:
        import io
        import calendar
        import traceback
        from openpyxl.utils import get_column_letter

        # Fetch all feeders
        feeders = await db.max_min_feeders.find({}, {"_id": 0}).to_list(100)
        
        # Sort feeders based on predefined order (INCLUDING ICTs)
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
            "220KV PARIGI-1",
            "220KV PARIGI-2",
            "220KV THANDUR",
            "220KV GACHIBOWLI-1",
            "220KV GACHIBOWLI-2",
            "220KV KETHIREDDYPALLY",
            "220KV YEDDUMAILARAM-1",
            "220KV YEDDUMAILARAM-2",
            "220KV SADASIVAPET-1",
            "220KV SADASIVAPET-2",
            "ICT-1 (315MVA)",
            "ICT-2 (315MVA)",
            "ICT-3 (315MVA)",
            "ICT-4 (500MVA)"
        ]
        
        ordered_feeders = []
        # Add Bus Station first
        bus_feeder = next((x for x in feeders if x.get('type') == "bus_station"), None)
        if bus_feeder:
            ordered_feeders.append(bus_feeder)
            
        # Add others in order
        for name in FEEDER_ORDER:
            if bus_feeder and name == bus_feeder.get('name'): continue
            f = next((x for x in feeders if x.get('name') == name), None)
            if f:
                ordered_feeders.append(f)
                
        
        # Helper to determine rating
        def get_rating(name):
            if "400KV" in name: return "400KV"
            if "220KV" in name: return "220KV"
            if "315MVA" in name: return "315MVA"
            if "500MVA" in name: return "500MVA"
            return "-"

        # Split feeders
        main_feeders = [f for f in ordered_feeders if f['type'] != 'ict_feeder' and f['type'] != 'bus_station']
        ict_feeders = [f for f in ordered_feeders if f['type'] == 'ict_feeder']
        bus_station_feeder = next((f for f in ordered_feeders if f['type'] == 'bus_station'), None)

        wb = Workbook()
        if wb.sheetnames:
            wb.remove(wb.active)
            
        last_day = calendar.monthrange(year, month)[1]
        month_name = calendar.month_name[month]
        
        periods = [
            {"name": "1-15", "start": f"{year}-{month:02d}-01", "end": f"{year}-{month:02d}-15"},
            {"name": "16-End", "start": f"{year}-{month:02d}-16", "end": f"{year}-{month:02d}-{last_day}"},
            {"name": "Full Month", "start": f"{year}-{month:02d}-01", "end": f"{year}-{month:02d}-{last_day}"}
        ]
        
        header_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        # header_font = Font(bold=True, color="000000") # Black text
        title_font = Font(bold=True, size=12)
        bold_font = Font(bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Fetch all entries for the month
        all_entries = await db.max_min_entries.find(
            {"date": {"$gte": f"{year}-{month:02d}-01", "$lte": f"{year}-{month:02d}-{last_day}"}},
            {"_id": 0}
        ).to_list(10000)

        entries_by_feeder = {}
        for e in all_entries:
            fid = e.get('feeder_id')
            if fid:
                if fid not in entries_by_feeder:
                    entries_by_feeder[fid] = []
                entries_by_feeder[fid].append(e)

        for p in periods:
            ws = wb.create_sheet(title=p['name'])
            
            # Row 1: Title
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
            title_cell = ws.cell(row=1, column=1, value=f"MAX DEMAND OF EACH FEEDER and ICT DURING THE MONTH OF {month_name}-{year} ({p['name']})")
            title_cell.font = title_font
            title_cell.alignment = center_align
            
            # Row 2: Main Headers
            # Col 1: Sl.No
            ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
            c = ws.cell(row=2, column=1, value="Sl.No")
            c.font = bold_font; c.alignment = center_align; c.border = thin_border
            ws.cell(row=3, column=1).border = thin_border
            
            # Col 2: Name of the feeder
            ws.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)
            c = ws.cell(row=2, column=2, value="Name of the feeder")
            c.font = bold_font; c.alignment = center_align; c.border = thin_border
            ws.cell(row=3, column=2).border = thin_border
            
            # Col 3: Rating
            ws.merge_cells(start_row=2, start_column=3, end_row=3, end_column=3)
            c = ws.cell(row=2, column=3, value="Rating")
            c.font = bold_font; c.alignment = center_align; c.border = thin_border
            ws.cell(row=3, column=3).border = thin_border
            
            # Col 4-7: Max Demand reached during
            ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=7)
            c = ws.cell(row=2, column=4, value="Max Demand reached during")
            c.font = bold_font; c.alignment = center_align; c.border = thin_border
            for i in range(5, 8): ws.cell(row=2, column=i).border = thin_border
            
            # Col 8-11: Min Demand reached
            ws.merge_cells(start_row=2, start_column=8, end_row=2, end_column=11)
            c = ws.cell(row=2, column=8, value="Min Demand reached")
            c.font = bold_font; c.alignment = center_align; c.border = thin_border
            for i in range(9, 12): ws.cell(row=2, column=i).border = thin_border
            
            # Col 12: Remarks
            ws.merge_cells(start_row=2, start_column=12, end_row=3, end_column=12)
            c = ws.cell(row=2, column=12, value="Remarks")
            c.font = bold_font; c.alignment = center_align; c.border = thin_border
            ws.cell(row=3, column=12).border = thin_border
            
            # Row 3: Sub Headers
            sub_headers = ["AMPS", "MW", "Date", "Time", "AMPS", "MW", "Date", "Time"]
            for i, h in enumerate(sub_headers):
                c = ws.cell(row=3, column=4+i, value=h)
                c.font = bold_font; c.alignment = center_align; c.border = thin_border
                
            row_idx = 4
            
            # 1. Main Feeders
            for i, feeder in enumerate(main_feeders, 1):
                f_entries = entries_by_feeder.get(feeder['id'], [])
                p_entries = [e for e in f_entries if p['start'] <= e['date'] <= p['end']]
                
                # Check for Grouping Logic
                is_special, partner_names = get_feeder_group_info(feeder['name'])
                if is_special:
                    p_group_map = {}
                    p_group_map[feeder['id']] = p_entries
                    
                    # Gather partner entries
                    for pname in partner_names:
                         p_feeder = next((x for x in feeders if x['name'] == pname), None)
                         if p_feeder:
                             pf_entries = entries_by_feeder.get(p_feeder['id'], [])
                             pp_entries = [e for e in pf_entries if p['start'] <= e['date'] <= p['end']]
                             p_group_map[p_feeder['id']] = pp_entries
                    
                    leader_id = determine_leader(p_group_map)
                    if not leader_id: leader_id = feeder['id']
                    
                    leader_entries = p_group_map.get(leader_id, [])
                    stats = calculate_coincident_stats(leader_entries, feeder['id'], p_group_map, feeder['type'])
                else:
                    stats = calculate_period_stats(p_entries, feeder['type'])
                
                ws.cell(row=row_idx, column=1, value=i).border = thin_border
                ws.cell(row=row_idx, column=1).alignment = center_align
                
                ws.cell(row=row_idx, column=2, value=feeder['name']).border = thin_border
                
                ws.cell(row=row_idx, column=3, value=get_rating(feeder['name'])).border = thin_border
                ws.cell(row=row_idx, column=3).alignment = center_align
                
                # Max
                ws.cell(row=row_idx, column=4, value=stats.get('max_amps', '-')).border = thin_border
                ws.cell(row=row_idx, column=5, value=stats.get('max_mw', '-')).border = thin_border
                ws.cell(row=row_idx, column=6, value=format_date(stats.get('max_mw_date'))).border = thin_border
                ws.cell(row=row_idx, column=7, value=format_time(stats.get('max_mw_time'))).border = thin_border
                
                # Min
                ws.cell(row=row_idx, column=8, value=stats.get('min_amps', '-')).border = thin_border
                ws.cell(row=row_idx, column=9, value=stats.get('min_mw', '-')).border = thin_border
                ws.cell(row=row_idx, column=10, value=format_date(stats.get('min_mw_date'))).border = thin_border
                ws.cell(row=row_idx, column=11, value=format_time(stats.get('min_mw_time'))).border = thin_border
                
                # Remarks
                ws.cell(row=row_idx, column=12, value="").border = thin_border
                
                # Center align data cells
                for c in range(4, 12): ws.cell(row=row_idx, column=c).alignment = center_align
                
                row_idx += 1
                
            # 2. ICT Separator
            row_idx += 1 # Gap
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=12)
            c = ws.cell(row=row_idx, column=1, value="ICT'S")
            c.font = bold_font; c.alignment = center_align
            # No border for separator? Or strictly follow image? Usually just text.
            row_idx += 1
            
            # 3. ICT Feeders
            start_sl = len(main_feeders) + 1
            for i, feeder in enumerate(ict_feeders, start_sl):
                f_entries = entries_by_feeder.get(feeder['id'], [])
                p_entries = [e for e in f_entries if p['start'] <= e['date'] <= p['end']]
                
                # Check for Grouping Logic (ICTs are in the group list)
                is_special, partner_names = get_feeder_group_info(feeder['name'])
                if is_special:
                    p_group_map = {}
                    p_group_map[feeder['id']] = p_entries
                    
                    # Gather partner entries
                    for pname in partner_names:
                         p_feeder = next((x for x in feeders if x['name'] == pname), None)
                         if p_feeder:
                             pf_entries = entries_by_feeder.get(p_feeder['id'], [])
                             pp_entries = [e for e in pf_entries if p['start'] <= e['date'] <= p['end']]
                             p_group_map[p_feeder['id']] = pp_entries
                    
                    leader_id = determine_leader(p_group_map)
                    if not leader_id: leader_id = feeder['id']
                    
                    leader_entries = p_group_map.get(leader_id, [])
                    stats = calculate_coincident_stats(leader_entries, feeder['id'], p_group_map, feeder['type'])
                else:
                    stats = calculate_period_stats(p_entries, feeder['type'])
                
                ws.cell(row=row_idx, column=1, value=i).border = thin_border
                ws.cell(row=row_idx, column=1).alignment = center_align
                
                ws.cell(row=row_idx, column=2, value=feeder['name']).border = thin_border
                
                ws.cell(row=row_idx, column=3, value=get_rating(feeder['name'])).border = thin_border
                ws.cell(row=row_idx, column=3).alignment = center_align
                
                # Max
                ws.cell(row=row_idx, column=4, value=stats.get('max_amps', '-')).border = thin_border
                ws.cell(row=row_idx, column=5, value=stats.get('max_mw', '-')).border = thin_border
                ws.cell(row=row_idx, column=6, value=format_date(stats.get('max_mw_date'))).border = thin_border
                ws.cell(row=row_idx, column=7, value=format_time(stats.get('max_mw_time'))).border = thin_border
                
                # Min
                ws.cell(row=row_idx, column=8, value=stats.get('min_amps', '-')).border = thin_border
                ws.cell(row=row_idx, column=9, value=stats.get('min_mw', '-')).border = thin_border
                ws.cell(row=row_idx, column=10, value=format_date(stats.get('min_mw_date'))).border = thin_border
                ws.cell(row=row_idx, column=11, value=format_time(stats.get('min_mw_time'))).border = thin_border
                
                # Remarks
                ws.cell(row=row_idx, column=12, value="").border = thin_border
                
                # Center align data cells
                for c in range(4, 12): ws.cell(row=row_idx, column=c).alignment = center_align
                
                row_idx += 1
                
            # 4. Station Load (Bottom)
            row_idx += 2
            if bus_station_feeder:
                f_entries = entries_by_feeder.get(bus_station_feeder['id'], [])
                p_entries = [e for e in f_entries if p['start'] <= e['date'] <= p['end']]
                stats = calculate_period_stats(p_entries, bus_station_feeder['type'])
                
                # Header
                ws.cell(row=row_idx, column=2, value="Station Load in MW").border = thin_border
                ws.cell(row=row_idx, column=2).font = bold_font; ws.cell(row=row_idx, column=2).alignment = center_align
                
                ws.cell(row=row_idx, column=3, value="Time").border = thin_border
                ws.cell(row=row_idx, column=3).font = bold_font; ws.cell(row=row_idx, column=3).alignment = center_align
                
                ws.cell(row=row_idx, column=4, value="Date").border = thin_border
                ws.cell(row=row_idx, column=4).font = bold_font; ws.cell(row=row_idx, column=4).alignment = center_align
                
                row_idx += 1
                
                # Value
                ws.cell(row=row_idx, column=1, value="Month").border = thin_border # Based on image 2 "Month" cell
                ws.cell(row=row_idx, column=2, value=stats.get('max_load', '-')).border = thin_border
                ws.cell(row=row_idx, column=2).alignment = center_align
                
                ws.cell(row=row_idx, column=3, value=format_time(stats.get('max_load_time'))).border = thin_border
                ws.cell(row=row_idx, column=3).alignment = center_align
                
                ws.cell(row=row_idx, column=4, value=format_date(stats.get('max_load_date'))).border = thin_border
                ws.cell(row=row_idx, column=4).alignment = center_align
                    
            # Set fixed width for Sl.No
            ws.column_dimensions['A'].width = 6
            
            # Auto-fit other columns
            for i in range(2, ws.max_column + 1):
                col_letter = get_column_letter(i)
                max_len = 0
                for cell in ws[col_letter]:
                    # Skip the first row (Title) to avoid skewed width
                    if cell.row == 1: continue
                    try:
                        if cell.value and len(str(cell.value)) > max_len:
                            max_len = len(str(cell.value))
                    except: pass
                ws.column_dimensions[col_letter].width = max_len + 2
                
        return wb
    except Exception as e:
        import traceback
        import os
        error_msg = traceback.format_exc()
        print(error_msg)
        raise e

@api_router.get("/reports/fortnight/{year}/{month}")
async def generate_fortnight_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        wb = await _generate_fortnight_report_wb(year, month)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        month_name = calendar.month_name[month]
        filename = f"Fortnight_Report_{month_name}_{year}.xlsx"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)},
            headers={
                "Access-Control-Allow-Origin": "http://localhost:3000",
                "Access-Control-Allow-Credentials": "true"
            }
        )

@api_router.get("/reports/fortnight/preview/{year}/{month}")
async def preview_fortnight_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        
        # Fetch all feeders
        feeders = await db.max_min_feeders.find({}, {"_id": 0}).to_list(100)
        
        # Sort feeders based on predefined order (INCLUDING ICTs)
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
            "220KV PARIGI-1",
            "220KV PARIGI-2",
            "220KV THANDUR",
            "220KV GACHIBOWLI-1",
            "220KV GACHIBOWLI-2",
            "220KV KETHIREDDYPALLY",
            "220KV YEDDUMAILARAM-1",
            "220KV YEDDUMAILARAM-2",
            "220KV SADASIVAPET-1",
            "220KV SADASIVAPET-2",
            "ICT-1 (315MVA)",
            "ICT-2 (315MVA)",
            "ICT-3 (315MVA)",
            "ICT-4 (500MVA)"
        ]
        
        ordered_feeders = []
        # Add Bus Station first
        bus_feeder = next((x for x in feeders if x.get('type') == "bus_station"), None)
        if bus_feeder:
            ordered_feeders.append(bus_feeder)
            
        # Add others in order
        for name in FEEDER_ORDER:
            if bus_feeder and name == bus_feeder.get('name'): continue
            f = next((x for x in feeders if x.get('name') == name), None)
            if f:
                ordered_feeders.append(f)
                
        # Helper to determine rating
        def get_rating(name):
            if "400KV" in name: return "400KV"
            if "220KV" in name: return "220KV"
            if "315MVA" in name: return "315MVA"
            if "500MVA" in name: return "500MVA"
            return "-"

        # Split feeders
        main_feeders = [f for f in ordered_feeders if f['type'] != 'ict_feeder' and f['type'] != 'bus_station']
        ict_feeders = [f for f in ordered_feeders if f['type'] == 'ict_feeder']
        bus_station_feeder = next((f for f in ordered_feeders if f['type'] == 'bus_station'), None)

        last_day = calendar.monthrange(year, month)[1]
        
        periods = [
            {"name": "1-15", "start": f"{year}-{month:02d}-01", "end": f"{year}-{month:02d}-15"},
            {"name": "16-End", "start": f"{year}-{month:02d}-16", "end": f"{year}-{month:02d}-{last_day}"},
            {"name": "Full Month", "start": f"{year}-{month:02d}-01", "end": f"{year}-{month:02d}-{last_day}"},
        ]
        
        # Fetch all entries for the month
        all_entries = await db.max_min_entries.find(
            {"date": {"$gte": f"{year}-{month:02d}-01", "$lte": f"{year}-{month:02d}-{last_day}"}},
            {"_id": 0}
        ).to_list(10000)

        entries_by_feeder = {}
        for e in all_entries:
            fid = e.get('feeder_id')
            if fid:
                if fid not in entries_by_feeder:
                    entries_by_feeder[fid] = []
                entries_by_feeder[fid].append(e)

        preview_data = {"periods": []}

        for p in periods:
            period_data = {
                "name": p['name'],
                "main_feeders": [],
                "ict_feeders": [],
                "station_load": None
            }
            
            # 1. Main Feeders
            for i, feeder in enumerate(main_feeders, 1):
                f_entries = entries_by_feeder.get(feeder['id'], [])
                p_entries = [e for e in f_entries if p['start'] <= e['date'] <= p['end']]
                
                # Check for Grouping Logic
                is_special, partner_names = get_feeder_group_info(feeder['name'])
                if is_special:
                    p_group_map = {}
                    p_group_map[feeder['id']] = p_entries
                    
                    # Gather partner entries
                    for pname in partner_names:
                         p_feeder = next((x for x in feeders if x['name'] == pname), None)
                         if p_feeder:
                             pf_entries = entries_by_feeder.get(p_feeder['id'], [])
                             pp_entries = [e for e in pf_entries if p['start'] <= e['date'] <= p['end']]
                             p_group_map[p_feeder['id']] = pp_entries
                    
                    leader_id = determine_leader(p_group_map)
                    if not leader_id: leader_id = feeder['id']
                    
                    leader_entries = p_group_map.get(leader_id, [])
                    stats = calculate_coincident_stats(leader_entries, feeder['id'], p_group_map, feeder['type'])
                else:
                    stats = calculate_period_stats(p_entries, feeder['type'])
                
                period_data["main_feeders"].append({
                    "sl_no": i,
                    "name": feeder['name'],
                    "rating": get_rating(feeder['name']),
                    "max_amps": stats.get('max_amps', '-'),
                    "max_mw": stats.get('max_mw', '-'),
                    "max_mw_date": format_date(stats.get('max_mw_date')),
                    "max_mw_time": format_time(stats.get('max_mw_time')),
                    "min_amps": stats.get('min_amps', '-'),
                    "min_mw": stats.get('min_mw', '-'),
                    "min_mw_date": format_date(stats.get('min_mw_date')),
                    "min_mw_time": format_time(stats.get('min_mw_time')),
                })
                
            # 2. ICT Feeders
            start_sl = len(main_feeders) + 1
            for i, feeder in enumerate(ict_feeders, start_sl):
                f_entries = entries_by_feeder.get(feeder['id'], [])
                p_entries = [e for e in f_entries if p['start'] <= e['date'] <= p['end']]
                
                # Check for Grouping Logic
                is_special, partner_names = get_feeder_group_info(feeder['name'])
                if is_special:
                    p_group_map = {}
                    p_group_map[feeder['id']] = p_entries
                    
                    # Gather partner entries
                    for pname in partner_names:
                         p_feeder = next((x for x in feeders if x['name'] == pname), None)
                         if p_feeder:
                             pf_entries = entries_by_feeder.get(p_feeder['id'], [])
                             pp_entries = [e for e in pf_entries if p['start'] <= e['date'] <= p['end']]
                             p_group_map[p_feeder['id']] = pp_entries
                    
                    leader_id = determine_leader(p_group_map)
                    if not leader_id: leader_id = feeder['id']
                    
                    leader_entries = p_group_map.get(leader_id, [])
                    stats = calculate_coincident_stats(leader_entries, feeder['id'], p_group_map, feeder['type'])
                else:
                    stats = calculate_period_stats(p_entries, feeder['type'])
                
                period_data["ict_feeders"].append({
                    "sl_no": i,
                    "name": feeder['name'],
                    "rating": get_rating(feeder['name']),
                    "max_amps": stats.get('max_amps', '-'),
                    "max_mw": stats.get('max_mw', '-'),
                    "max_mw_date": format_date(stats.get('max_mw_date')),
                    "max_mw_time": format_time(stats.get('max_mw_time')),
                    "min_amps": stats.get('min_amps', '-'),
                    "min_mw": stats.get('min_mw', '-'),
                    "min_mw_date": format_date(stats.get('min_mw_date')),
                    "min_mw_time": format_time(stats.get('min_mw_time')),
                })

            # 3. Station Load
            if bus_station_feeder:
                f_entries = entries_by_feeder.get(bus_station_feeder['id'], [])
                p_entries = [e for e in f_entries if p['start'] <= e['date'] <= p['end']]
                stats = calculate_period_stats(p_entries, bus_station_feeder['type'])
                
                period_data["station_load"] = {
                    "max_load": stats.get('max_load', '-'),
                    "max_load_time": format_time(stats.get('max_load_time')),
                    "max_load_date": format_date(stats.get('max_load_date'))
                }
                
            preview_data["periods"].append(period_data)

        return preview_data
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(error_msg)
        return JSONResponse(status_code=500, content={"detail": "Failed to generate Fortnight preview."})


@api_router.get("/reports/daily-max-mva/preview/{year}/{month}")
async def get_daily_max_mva_preview(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import math
        import calendar
        
        # 1. Find "Bus Voltages & Station Load" feeder
        feeder = await db.max_min_feeders.find_one({"name": "Bus Voltages & Station Load"})
        if not feeder:
            feeder = await db.max_min_feeders.find_one({"type": "bus_station"})
            
        if not feeder:
            return JSONResponse(status_code=404, content={"detail": "Bus Station feeder not found"})

        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"

        entries = await db.max_min_entries.find(
            {"feeder_id": feeder['id'], "date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).sort("date", 1).to_list(1000)
        
        report_data = []
        
        # Helper to parse float
        def parse_float(val):
            try:
                return float(val)
            except (ValueError, TypeError):
                return None

        last_day = calendar.monthrange(year, month)[1]
        entries_map = {e['date']: e for e in entries}
        
        for day in range(1, last_day + 1):
            date_str = f"{year}-{month:02d}-{day:02d}"
            entry = entries_map.get(date_str)
            
            row = {
                "date": format_date(date_str),
                "mw": "",
                "mvar": "",
                "time": "",
                "mva": ""
            }
            
            if entry:
                d = entry.get('data', {}).get('station_load', {})
                mw = d.get('max_mw')
                mvar = d.get('mvar')
                time = d.get('time')
                
                row["mw"] = mw if mw is not None else ""
                row["mvar"] = mvar if mvar is not None else ""
                row["time"] = format_time(time) if time else ""
                
                mw_val = parse_float(mw)
                mvar_val = parse_float(mvar)
                
                if mw_val is not None and mvar_val is not None:
                    mva = math.sqrt(mw_val**2 + mvar_val**2)
                    row["mva"] = f"{mva:.2f}"
            
            report_data.append(row)
            
        return report_data
        
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        with open("error_log.txt", "w") as f:
            f.write(error_msg)
        print(error_msg)
        return JSONResponse(status_code=500, content={"detail": str(e)})


async def _generate_daily_max_mva_wb(year: int, month: int):
    try:
        import math
        import calendar
        import io
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
        from openpyxl.utils import get_column_letter
        
        feeder = await db.max_min_feeders.find_one({"name": "Bus Voltages & Station Load"})
        if not feeder:
            feeder = await db.max_min_feeders.find_one({"type": "bus_station"})
            
        if not feeder:
            raise HTTPException(status_code=404, detail="Bus Station feeder not found")

        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"

        entries = await db.max_min_entries.find(
            {"feeder_id": feeder['id'], "date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).sort("date", 1).to_list(1000)
        
        def parse_float(val):
            try:
                return float(val)
            except (ValueError, TypeError):
                return None

        wb = Workbook()
        ws = wb.active
        ws.title = f"Daily Max MVA {month}-{year}"
        
        # Styles
        header_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
        header_font = Font(bold=True, color="000000")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal='center', vertical='center')
        
        headers = ["Date", "MW", "MVAR", "Time", "MVA"]
        
        for i, h in enumerate(headers):
            cell = ws.cell(row=1, column=i+1, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            
        entries_map = {e['date']: e for e in entries}
        last_day = calendar.monthrange(year, month)[1]
        
        row_idx = 2
        for day in range(1, last_day + 1):
            date_str = f"{year}-{month:02d}-{day:02d}"
            formatted_date = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d-%b-%y")
            
            entry = entries_map.get(date_str)
            
            ws.cell(row=row_idx, column=1, value=formatted_date).border = thin_border
            ws.cell(row=row_idx, column=1).alignment = center_align
            
            for c in range(2, 6):
                ws.cell(row=row_idx, column=c).border = thin_border
                ws.cell(row=row_idx, column=c).alignment = center_align
            
            if entry:
                d = entry.get('data', {}).get('station_load', {})
                mw = d.get('max_mw')
                mvar = d.get('mvar')
                time = d.get('time')
                
                if mw is not None: ws.cell(row=row_idx, column=2, value=mw)
                if mvar is not None: ws.cell(row=row_idx, column=3, value=mvar)
                if time: ws.cell(row=row_idx, column=4, value=format_time(time))
                
                mw_val = parse_float(mw)
                mvar_val = parse_float(mvar)
                
                if mw_val is not None and mvar_val is not None:
                    mva = math.sqrt(mw_val**2 + mvar_val**2)
                    ws.cell(row=row_idx, column=5, value=f"{mva:.2f}")
            
            row_idx += 1

        for i in range(1, 6):
            col_letter = get_column_letter(i)
            max_len = 0
            for cell in ws[col_letter]:
                try:
                    if cell.value and len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except: pass
            ws.column_dimensions[col_letter].width = max(max_len + 2, 12)

        return wb
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise e

@api_router.get("/reports/daily-max-mva/export/{year}/{month}")
async def export_daily_max_mva_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import io
        wb = await _generate_daily_max_mva_wb(year, month)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"Daily_Max_MVA_{month}-{year}.xlsx"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": str(e)})




async def _generate_energy_export_wb(year: int, month: int):
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
            
    return wb

@api_router.get("/energy/export-all/{year}/{month}")
async def export_all_energy_sheets(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    wb = await _generate_energy_export_wb(year, month)
    
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

async def get_boundary_meter_data(year: int, month: int):
    # 1. Find 33KV Sheet
    sheet = await db.energy_sheets.find_one({"name": "33KV"})
    if not sheet:
        raise HTTPException(status_code=404, detail="33KV Sheet not found")
    
    # 2. Find Meters (Donthapally & Kandi)
    meters_cursor = db.energy_meters.find({"sheet_id": sheet['id']})
    meters = await meters_cursor.to_list(100)
    
    target_meters = []
    # Map for easy lookup and to preserve order
    # User Order: 1. Donthapally, 2. Kandi
    donthapally = next((m for m in meters if "Donthapally" in m['name']), None)
    kandi = next((m for m in meters if "Kandi" in m['name']), None)
    
    if donthapally:
        target_meters.append({
            "meter": donthapally, 
            "display_name": "33KV Donthanpally\nservice no RRS:1445\nSerial no:14754097"
        })
    if kandi:
        target_meters.append({
            "meter": kandi, 
            "display_name": "33KV Kandi(BDL)\nservice no RRS:1445\nSerial no:04932269"
        })
        
    if not target_meters:
         raise HTTPException(status_code=404, detail="Target meters (Donthapally/Kandi) not found")

    # 3. Date Calculations
    # Current Month Start and End
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    month_start_date = f"{year}-{month:02d}-01"
    month_end_date = f"{year}-{month:02d}-{last_day}"
    
    # Previous Month Last Day (for display)
    first_date_obj = datetime(year, month, 1)
    prev_month_last_date_obj = first_date_obj - timedelta(days=1)
    prev_month_str = prev_month_last_date_obj.strftime("%d-%b-%y")
    current_month_end_str = datetime(year, month, last_day).strftime("%d-%b-%y")
    
    # 4. Fetch Readings
    # We need:
    # Initial: Initial reading of the first day of the month (or Final of prev month last day)
    # Final: Final reading of the last day of the month
    
    # Fetch first day entry
    start_entry = await db.energy_entries.find_one({
        "sheet_id": sheet['id'],
        "date": month_start_date
    })
    
    # Fetch last day entry
    end_entry = await db.energy_entries.find_one({
        "sheet_id": sheet['id'],
        "date": month_end_date
    })
    
    report_data = []
    
    for item in target_meters:
        meter = item['meter']
        meter_id = meter['id']
        mf = meter['mf']
        
        initial_val = 0.0
        final_val = 0.0
        
        # Get Initial
        if start_entry:
            reading = next((r for r in start_entry['readings'] if r['meter_id'] == meter_id), None)
            if reading:
                initial_val = reading['initial']
                
        # Get Final
        if end_entry:
            reading = next((r for r in end_entry['readings'] if r['meter_id'] == meter_id), None)
            if reading:
                final_val = reading['final']
        
        diff = final_val - initial_val
        consumption = diff * mf
        
        report_data.append({
            "name": item['display_name'],
            "initial": initial_val,
            "final": final_val,
            "diff": diff,
            "mf": mf,
            "consumption": consumption
        })
        
    return {
        "report_data": report_data,
        "prev_month_str": prev_month_str,
        "current_month_end_str": current_month_end_str,
        "month_name": calendar.month_name[month]
    }

@api_router.get("/reports/boundary-meter-33kv/data/{year}/{month}")
async def get_boundary_meter_report_json(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    return await get_boundary_meter_data(year, month)

async def _generate_boundary_meter_wb(year: int, month: int):
    data = await get_boundary_meter_data(year, month)
    report_data = data['report_data']
    prev_month_str = data['prev_month_str']
    current_month_end_str = data['current_month_end_str']
    month_name = data['month_name']

    # 5. Generate Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Boundary Meter Report"
    
    # Styling
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bold_font = Font(bold=True)
    
    # Title Row
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = f"Boundary Meter Readings of 400KV Shankarpally for the Month of {month_name[:3]}'{str(year)[-2:]}"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = center_align
    title_cell.border = thin_border
    
    # Sub-header
    ws.merge_cells('A2:H2')
    ws['A2'].value = "33KV Station Supply"
    ws['A2'].alignment = center_align
    ws['A2'].font = bold_font
    ws['A2'].border = thin_border
    
    # Header Rows (3 & 4)
    # Sl. No. (A3:A4)
    ws.merge_cells('A3:A4')
    ws['A3'] = "Sl.\nNo."
    
    # Name of Feeder (B3:B4)
    ws.merge_cells('B3:B4')
    ws['B3'] = "Name of Feeder"
    
    # Timing (C3:D3)
    ws.merge_cells('C3:D3')
    ws['C3'] = "Timing"
    ws['C4'] = "From"
    ws['D4'] = "To"
    
    # Initial Reading (E3:E4)
    ws.merge_cells('E3:E4')
    ws['E3'] = "Initial\nReading"
    
    # Final Reading (F3:F4)
    ws.merge_cells('F3:F4')
    ws['F3'] = "Final\nReading"
    
    # Diff (G3:G4)
    ws.merge_cells('G3:G4')
    ws['G3'] = "Diff."
    
    # M.F. (H3:H4)
    ws.merge_cells('H3:H4')
    ws['H3'] = "M.F."
    
    # Consumption (I3:I4) -> Wait, user image has "Consumption" then "Remarks".
    # User image columns: Sl No, Name, Timing(From, To), Initial, Final, Diff, MF, Consumption, Remarks
    # My Columns: A, B, C, D, E, F, G, H, I, J
    # Let's re-map:
    # A: Sl No
    # B: Name
    # C: Timing From
    # D: Timing To
    # E: Initial
    # F: Final
    # G: Diff
    # H: MF
    # I: Consumption
    # J: Remarks
    
    # Adjust Merges
    ws.unmerge_cells('A1:H1')
    ws.merge_cells('A1:J1')
    ws.unmerge_cells('A2:H2')
    ws.merge_cells('A2:J2')
    
    ws.merge_cells('I3:I4')
    ws['I3'] = "Consumption"
    
    ws.merge_cells('J3:J4')
    ws['J3'] = "Remarks"
    
    # Apply styles to header
    for row in ws.iter_rows(min_row=3, max_row=4, min_col=1, max_col=10):
        for cell in row:
            cell.alignment = center_align
            cell.font = bold_font
            cell.border = thin_border
            
    # Data Rows
    start_row = 5
    for idx, data in enumerate(report_data, 1):
        ws[f'A{start_row}'] = idx
        ws[f'B{start_row}'] = data['name']
        ws[f'C{start_row}'] = f"{prev_month_str}\n12:00 Hrs"
        ws[f'D{start_row}'] = f"{current_month_end_str}\n12:00 Hrs"
        ws[f'E{start_row}'] = data['initial']
        ws[f'F{start_row}'] = data['final']
        ws[f'G{start_row}'] = data['diff']
        ws[f'H{start_row}'] = data['mf']
        ws[f'I{start_row}'] = data['consumption']
        ws[f'J{start_row}'] = "-"
        
        # Style row
        for col in range(1, 11):
            cell = ws.cell(row=start_row, column=col)
            cell.alignment = center_align
            cell.border = thin_border
            # Number formatting
            if col in [5, 6, 7, 9]: # Readings, Diff, Consumption
                cell.number_format = '0.00'
                
        # Row height for multiline text
        ws.row_dimensions[start_row].height = 60
        start_row += 1
        
    # Column Widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 8
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 10

    return wb

@api_router.get("/reports/boundary-meter-33kv/{year}/{month}")
async def generate_boundary_meter_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        from io import BytesIO
        
        # We need month_name for the filename, but it's not directly returned by the wb function.
        # However, we can re-calculate it or fetch it.
        # But _generate_boundary_meter_wb calculates it internally.
        # Maybe we should return (wb, filename) from the internal function?
        # Or just recalculate it here.
        
        wb = await _generate_boundary_meter_wb(year, month)
        month_name = calendar.month_name[month]
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"Boundary_Meter_Report_{month_name}_{year}.xlsx"
        headers = {
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
        
        return StreamingResponse(
            output,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers=headers
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": str(e)})



# --- KPI Report Logic ---

KPI_FEEDER_DETAILS = {
    "400KV MAHESHWARAM-2": {"length_km": 84.13, "length_ckm": 84.13, "conductor": "ACSR Twin Moose", "capacity": 1786},
    "400KV MAHESHWARAM-1": {"length_km": 84.13, "length_ckm": 84.13, "conductor": "ACSR Twin Moose", "capacity": 1786},
    "400KV NARSAPUR-1": {"length_km": 60.92, "length_ckm": 60.92, "conductor": "ACSR Twin Moose", "capacity": 1786},
    "400KV NARSAPUR-2": {"length_km": 60.92, "length_ckm": 60.92, "conductor": "ACSR Twin Moose", "capacity": 1786},
    "400KV KETHIREDDYPALLY-1": {"length_km": 51.13, "length_ckm": 51.13, "conductor": "ACSR Quad Moose", "capacity": 3572},
    "400KV KETHIREDDYPALLY-2": {"length_km": 51.13, "length_ckm": 51.13, "conductor": "ACSR Quad Moose", "capacity": 3572},
    "400KV NIZAMABAD-1": {"length_km": 139, "length_ckm": 139, "conductor": "ACSR Twin Moose", "capacity": 1786},
    "400KV NIZAMABAD-2": {"length_km": 139, "length_ckm": 139, "conductor": "ACSR Twin Moose", "capacity": 1786},
    "220KV PARIGI-1": {"length_km": 45.05, "length_ckm": 45.05, "conductor": "ACSR Moose", "capacity": 795},
    "220KV PARIGI-2": {"length_km": 45.05, "length_ckm": 45.05, "conductor": "ACSR Moose", "capacity": 795},
    "220KV THANDUR": {"length_km": 85.057, "length_ckm": 85.057, "conductor": "ACSR Moose", "capacity": 795},
    "220KV GACHIBOWLI-1": {"length_km": 28.57, "length_ckm": 28.57, "conductor": "HTLS", "capacity": 1600},
    "220KV GACHIBOWLI-2": {"length_km": 23.3, "length_ckm": 23.3, "conductor": "HTLS", "capacity": 1600},
    "220KV KETHIREDDYPALLY": {"length_km": 29.44, "length_ckm": 29.44, "conductor": "ACSR Moose", "capacity": 795},
    "220KV YEDDUMAILARAM-1": {"length_km": 10.3, "length_ckm": 10.3, "conductor": "ACSR Moose", "capacity": 795},
    "220KV YEDDUMAILARAM-2": {"length_km": 10.2, "length_ckm": 10.2, "conductor": "ACSR Zebra", "capacity": 795},
    "220KV SADASIVAPET-1": {"length_km": 37.151, "length_ckm": 37.151, "conductor": "ACSR Zebra", "capacity": 795},
    "220KV SADASIVAPET-2": {"length_km": 37.151, "length_ckm": 37.151, "conductor": "ACSR Zebra", "capacity": 795},
}

FEEDER_ORDER_KPI = [
    "400KV MAHESHWARAM-2", "400KV MAHESHWARAM-1", 
    "400KV NARSAPUR-1", "400KV NARSAPUR-2",
    "400KV KETHIREDDYPALLY-1", "400KV KETHIREDDYPALLY-2",
    "400KV NIZAMABAD-1", "400KV NIZAMABAD-2",
    "220KV PARIGI-1", "220KV PARIGI-2",
    "220KV THANDUR",
    "220KV GACHIBOWLI-1", "220KV GACHIBOWLI-2",
    "220KV KETHIREDDYPALLY",
    "220KV YEDDUMAILARAM-1", "220KV YEDDUMAILARAM-2",
    "220KV SADASIVAPET-1", "220KV SADASIVAPET-2"
]

ICT_ORDER_KPI = ["ICT-1 (315MVA)", "ICT-2 (315MVA)", "ICT-3 (315MVA)", "ICT-4 (500MVA)"]

def calculate_kpi_stats(entries, feeder_type):
    if not entries:
        return {"avg_val": 0, "max_val": 0}
        
    total_avg = 0
    max_val = 0
    count = 0
    
    # For Lines: track Max MW to determine which Amps to pick
    max_mw_found = -1.0
    
    for e in entries:
        d = e.get('data', {})
        if not d: continue
        
        max_data = d.get('max') or {}
        avg_data = d.get('avg') or {}
        
        if feeder_type == 'ict_feeder':
            # ICT: Max Demand (MW) and Avg Load (MW)
            # Max MW comes from max.mw
            # Avg MW comes from avg.mw
            curr_max = float(max_data.get('mw', 0) or 0)
            curr_avg = float(avg_data.get('mw', 0) or 0)
            
            if curr_max > max_val:
                max_val = curr_max
            
            if curr_avg > 0:
                total_avg += curr_avg
                count += 1
        else:
            # Line: Max Line Loading (Amps) and Avg Loading (Amps)
            # Logic Update: Max Amps should be derived from the entry with Max MW
            curr_mw = float(max_data.get('mw', 0) or 0)
            curr_amps = float(max_data.get('amps', 0) or 0)
            
            if curr_mw > max_mw_found:
                max_mw_found = curr_mw
                max_val = curr_amps
            
            curr_avg = float(avg_data.get('amps', 0) or 0)
            if curr_avg > 0:
                total_avg += curr_avg
                count += 1
                
    avg_val = total_avg / count if count > 0 else 0
    return {"avg_val": avg_val, "max_val": max_val}

@api_router.get("/reports/kpi/preview/{year}/{month}")
async def get_kpi_preview(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        
        # Date Range
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"
            
        # 1. Lines Data
        lines_data = []
        feeders = await db.max_min_feeders.find({"type": {"$in": ["feeder_400kv", "feeder_220kv"]}}, {"_id": 0}).to_list(100)
        
        # Sort feeders
        feeders.sort(key=lambda x: FEEDER_ORDER_KPI.index(x['name']) if x['name'] in FEEDER_ORDER_KPI else 999)
        
        all_entries = await db.max_min_entries.find(
            {"date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).to_list(5000)
        
        entries_by_feeder = {}
        for e in all_entries:
            if e['feeder_id'] not in entries_by_feeder:
                entries_by_feeder[e['feeder_id']] = []
            entries_by_feeder[e['feeder_id']].append(e)
            
        for idx, f in enumerate(feeders):
            if f['name'] not in FEEDER_ORDER_KPI: continue # Only show defined feeders
            
            details = KPI_FEEDER_DETAILS.get(f['name'], {})
            entries = entries_by_feeder.get(f['id'], [])
            stats = calculate_kpi_stats(entries, f['type'])
            
            avg_load = stats['avg_val']
            max_load = stats['max_val']
            pct_load = (avg_load / max_load * 100) if max_load > 0 else 0
            
            lines_data.append({
                "sl_no": idx + 1,
                "zone": "Metro Zone",
                "circle": "Metro- West",
                "feeder_name": f['name'],
                "length_km": details.get('length_km', '-'),
                "length_ckm": details.get('length_ckm', '-'),
                "conductor": details.get('conductor', '-'),
                "capacity": details.get('capacity', '-'),
                "avg_loading": f"{avg_load:.2f}" if avg_load else "-",
                "max_loading": f"{max_load:.2f}" if max_load else "-",
                "pct_loading": f"{pct_load:.2f}" if pct_load else "-"
            })
            
        # 2. ICT Data
        ict_data = []
        ict_feeders = await db.max_min_feeders.find({"type": "ict_feeder"}, {"_id": 0}).to_list(100)
        ict_feeders.sort(key=lambda x: ICT_ORDER_KPI.index(x['name']) if x['name'] in ICT_ORDER_KPI else 999)
        
        for idx, f in enumerate(ict_feeders):
            entries = entries_by_feeder.get(f['id'], [])
            stats = calculate_kpi_stats(entries, f['type'])
            
            avg_load = stats['avg_val']
            max_demand = stats['max_val']
            pct_load = (avg_load / max_demand * 100) if max_demand > 0 else 0
            
            # Format Name: ICT-1 (315MVA) -> 315 MVA ICT-1
            name_parts = f['name'].split(' ')
            formatted_name = f['name']
            if len(name_parts) >= 2:
                 try:
                     ict_part = name_parts[0] # ICT-1
                     mva_part = name_parts[1].replace('(', '').replace(')', '') # 315MVA
                     formatted_name = f"{mva_part[:3]} MVA {ict_part}" # 315 MVA ICT-1
                     if "500" in mva_part: formatted_name = "500MVA ICT-4"
                 except:
                     pass
            
            ict_data.append({
                "sl_no": idx + 1,
                "zone": "Merto Zone",
                "circle": "OMC- Metro- West Circle",
                "substation": "400KV Shankarpally",
                "capacity_name": formatted_name,
                "proposed": "-",
                "max_demand": f"{max_demand:.2f}" if max_demand else "-",
                "avg_load": f"{avg_load:.2f}" if avg_load else "-",
                "pct_loading": f"{pct_load:.2f}" if pct_load else "-"
            })
            
        return {"lines": lines_data, "icts": ict_data}
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"detail": str(e)})

async def _generate_kpi_report_wb(year: int, month: int):
    import calendar
    from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
    from openpyxl.utils import get_column_letter
    
    # --- Common Styles ---
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bold_font = Font(bold=True)
    
    wb = Workbook()
    
    # --- Data Prep ---
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"
    
    month_name = calendar.month_name[month]
    
    all_entries = await db.max_min_entries.find(
        {"date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).to_list(5000)
    
    entries_by_feeder = {}
    for e in all_entries:
        if e['feeder_id'] not in entries_by_feeder:
            entries_by_feeder[e['feeder_id']] = []
        entries_by_feeder[e['feeder_id']].append(e)

    # ================= SHEET 1: Over Loading of Lines =================
    ws1 = wb.active
    ws1.title = "Over Loading of Lines"
    
    # Header Rows
    ws1.merge_cells('A1:L1')
    c = ws1.cell(row=1, column=1, value="400/220KV SHANKARPALLY SUBSTATION")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws1.merge_cells('A2:L2')
    c = ws1.cell(row=2, column=1, value=f"STATEMENT 20: FOR THE MONTH OF {month_name}-{year}")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws1.merge_cells('A3:L3')
    c = ws1.cell(row=3, column=1, value="Overloading of Lines")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    headers = [
        "Sl. No.", "Name of Zone", "Circle", "Name of feeder", 
        "Length in KM (Total Length)", "Length in CKM", "Type of Conductor",
        "Current Carrying Capacity", "Avg.Loading (Amps)", "Max.Line loading (Amps)",
        "% line loading", "Remarks"
    ]
    
    for i, h in enumerate(headers):
        cell = ws1.cell(row=4, column=i+1, value=h)
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border
        
    feeders = await db.max_min_feeders.find({"type": {"$in": ["feeder_400kv", "feeder_220kv"]}}, {"_id": 0}).to_list(100)
    feeders.sort(key=lambda x: FEEDER_ORDER_KPI.index(x['name']) if x['name'] in FEEDER_ORDER_KPI else 999)
    
    row_idx = 5
    sl_no = 1
    for f in feeders:
        if f['name'] not in FEEDER_ORDER_KPI: continue
        
        details = KPI_FEEDER_DETAILS.get(f['name'], {})
        entries = entries_by_feeder.get(f['id'], [])
        stats = calculate_kpi_stats(entries, f['type'])
        
        avg_load = stats['avg_val']
        max_load = stats['max_val']
        pct_load = (avg_load / max_load * 100) if max_load > 0 else 0
        
        row_data = [
            sl_no,
            "Metro Zone",
            "Metro- West",
            f['name'],
            details.get('length_km', '-'),
            details.get('length_ckm', '-'),
            details.get('conductor', '-'),
            details.get('capacity', '-'),
            f"{avg_load:.2f}" if avg_load else "-",
            f"{max_load:.2f}" if max_load else "-",
            f"{pct_load:.2f}" if pct_load else "-",
            "-"
        ]
        
        for i, val in enumerate(row_data):
            cell = ws1.cell(row=row_idx, column=i+1, value=val)
            cell.alignment = center_align
            cell.border = thin_border
            
        row_idx += 1
        sl_no += 1
        
    for i in range(1, ws1.max_column + 1):
        col_letter = get_column_letter(i)
        max_len = 0
        for cell in ws1[col_letter]:
            if cell.row <= 3: continue
            try:
                if cell.value and len(str(cell.value)) > max_len:
                    max_len = len(str(cell.value))
            except: pass
        ws1.column_dimensions[col_letter].width = max(max_len + 2, 10)


    # ================= SHEET 2: ICT'S =================
    ws2 = wb.create_sheet(title="ICT'S")
    
    ws2.merge_cells('A1:J1')
    c = ws2.cell(row=1, column=1, value="400/220KV SHANKARAPALLY")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws2.merge_cells('A2:J2')
    c = ws2.cell(row=2, column=1, value="ANNEXURE XVIII")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws2.merge_cells('A3:J3')
    c = ws2.cell(row=3, column=1, value="REDUCTION OF TRANSMISSION LINE FORMATS")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws2.merge_cells('A4:J4')
    c = ws2.cell(row=4, column=1, value=f"(A) Details of overloading of PTRs (70% and above) For {month_name}-{year}")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    headers = [
        "Sl.No", "Name of Zone", "TL&ss Circle", "Name of Substation", 
        "Existing ICT capacity (MVA)", "Proposed augmentation of ICT capacity",
        "Max. Demand reached during the month (MW)", "Average load in MW",
        "Average Percentage of ICT loading", "Remarks"
    ]
    
    for i, h in enumerate(headers):
        cell = ws2.cell(row=5, column=i+1, value=h)
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border
        
    ict_feeders = await db.max_min_feeders.find({"type": "ict_feeder"}, {"_id": 0}).to_list(100)
    ict_feeders.sort(key=lambda x: ICT_ORDER_KPI.index(x['name']) if x['name'] in ICT_ORDER_KPI else 999)
    
    row_idx = 6
    sl_no = 1
    for f in ict_feeders:
        entries = entries_by_feeder.get(f['id'], [])
        stats = calculate_kpi_stats(entries, f['type'])
        
        avg_load = stats['avg_val']
        max_demand = stats['max_val']
        pct_load = (avg_load / max_demand * 100) if max_demand > 0 else 0
        
        name_parts = f['name'].split(' ')
        formatted_name = f['name']
        if len(name_parts) >= 2:
                try:
                    ict_part = name_parts[0] 
                    mva_part = name_parts[1].replace('(', '').replace(')', '')
                    formatted_name = f"{mva_part[:3]} MVA {ict_part}"
                    if "500" in mva_part: formatted_name = "500MVA ICT-4"
                except: pass
        
        row_data = [
            sl_no,
            "Merto Zone", 
            "OMC- Metro- West Circle",
            "400KV Shankarpally",
            formatted_name,
            "-",
            f"{max_demand:.2f}" if max_demand else "-",
            f"{avg_load:.2f}" if avg_load else "-",
            f"{pct_load:.2f}" if pct_load else "-",
            "-"
        ]
        
        for i, val in enumerate(row_data):
            cell = ws2.cell(row=row_idx, column=i+1, value=val)
            cell.alignment = center_align
            cell.border = thin_border
            
        row_idx += 1
        sl_no += 1

    for i in range(1, ws2.max_column + 1):
        col_letter = get_column_letter(i)
        max_len = 0
        for cell in ws2[col_letter]:
            if cell.row <= 4: continue
            try:
                if cell.value and len(str(cell.value)) > max_len:
                    max_len = len(str(cell.value))
            except: pass
        ws2.column_dimensions[col_letter].width = max(max_len + 2, 12)
        
    return wb

@api_router.get("/reports/kpi/export/{year}/{month}")
async def export_kpi_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        wb = await _generate_kpi_report_wb(year, month)
        month_name = calendar.month_name[month]
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"KPI_Report_{month_name}_{year}.xlsx"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        with open("export_error_log.txt", "w") as f:
            f.write(error_msg)
        print(error_msg)
        return JSONResponse(status_code=500, content={"detail": str(e)})


# --- Line Losses Report ---

@api_router.get("/reports/line-losses/preview/{year}/{month}")
async def get_line_losses_report_preview(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        
        # 1. Date Range
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"
            
        # 2. Get Feeders (Line Losses Module)
        # We need to use 'feeders' collection
        all_feeders = await db.feeders.find({}, {"_id": 0}).to_list(100)
        
        # Mapping DB Names to Report Names
        FEEDER_MAPPING = {
            "400 KV Shanakrapally-MHRM-2": "400KV MAHESHWARAM-2",
            "400 KV Shanakrapally-MHRM-1": "400KV MAHESHWARAM-1",
            "400 KV Shanakrapally-Narsapur-1": "400KV NARSAPUR-1",
            "400 KV Shanakrapally-Narsapur-2": "400KV NARSAPUR-2",
            "400 KV KethiReddyPally-1": "400KV KETHIREDDYPALLY-1",
            "400 KV KethiReddyPally-2": "400KV KETHIREDDYPALLY-2",
            "220 KV Parigi-1": "220KV PARIGI-1",
            "220 KV Parigi-2": "220KV PARIGI-2",
            "220 KV Tandur": "220KV THANDUR",
            "220 KV Gachibowli-1": "220KV GACHIBOWLI-1",
            "220 KV Gachibowli-2": "220KV GACHIBOWLI-2",
            "220 KV KethiReddyPally": "220KV KETHIREDDYPALLY",
            "220 KV Yeddumailaram-1": "220KV YEDDUMAILARAM-1",
            "220 KV Yeddumailaram-2": "220KV YEDDUMAILARAM-2",
            "220 KV Sadasivapet-1": "220KV SADASIVAPET-1",
            "220 KV Sadasivapet-2": "220KV SADASIVAPET-2"
        }

        # Filter and Sort
        FEEDER_ORDER = [
            "400KV MAHESHWARAM-2",
            "400KV MAHESHWARAM-1",
            "400KV NARSAPUR-1",
            "400KV NARSAPUR-2",
            "400KV KETHIREDDYPALLY-1",
            "400KV KETHIREDDYPALLY-2",
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
        
        target_feeders = []
        for f in all_feeders:
            if f['name'] in FEEDER_MAPPING:
                f['display_name'] = FEEDER_MAPPING[f['name']]
                target_feeders.append(f)
        
        target_feeders.sort(key=lambda x: FEEDER_ORDER.index(x['display_name']) if x['display_name'] in FEEDER_ORDER else 999)
        
        # 3. Get Entries
        entries = await db.entries.find(
            {"date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).to_list(5000)
        
        entries_by_feeder = {}
        for e in entries:
            if e['feeder_id'] not in entries_by_feeder:
                entries_by_feeder[e['feeder_id']] = []
            entries_by_feeder[e['feeder_id']].append(e)
            
        # 4. Build Data
        report_data = []
        for idx, f in enumerate(target_feeders):
            f_entries = entries_by_feeder.get(f['id'], [])
            # Sort by date
            f_entries.sort(key=lambda x: x['date'])
            
            data = {
                "sl_no": idx + 1,
                "feeder_name": f['display_name'],
                "shankarpally": {"import": {}, "export": {}},
                "other_end": {"import": {}, "export": {}},
                "stats": {}
            }
            
            if f_entries:
                first = f_entries[0]
                last = f_entries[-1]
                
                # Shankarpally (End 1)
                # Import
                s_imp_init = first.get('end1_import_initial', 0)
                s_imp_final = last.get('end1_import_final', 0)
                s_imp_mf = f.get('end1_import_mf', 1)
                s_imp_cons = (s_imp_final - s_imp_init) * s_imp_mf
                
                data["shankarpally"]["import"] = {
                    "initial": s_imp_init,
                    "final": s_imp_final,
                    "mf": s_imp_mf,
                    "consumption": s_imp_cons
                }
                
                # Export
                s_exp_init = first.get('end1_export_initial', 0)
                s_exp_final = last.get('end1_export_final', 0)
                s_exp_mf = f.get('end1_export_mf', 1)
                s_exp_cons = (s_exp_final - s_exp_init) * s_exp_mf
                
                data["shankarpally"]["export"] = {
                    "initial": s_exp_init,
                    "final": s_exp_final,
                    "mf": s_exp_mf,
                    "consumption": s_exp_cons
                }
                
                # Other End (End 2)
                # Import
                o_imp_init = first.get('end2_import_initial', 0)
                o_imp_final = last.get('end2_import_final', 0)
                o_imp_mf = f.get('end2_import_mf', 1)
                o_imp_cons = (o_imp_final - o_imp_init) * o_imp_mf
                
                data["other_end"]["import"] = {
                    "initial": o_imp_init,
                    "final": o_imp_final,
                    "mf": o_imp_mf,
                    "consumption": o_imp_cons
                }
                
                # Export
                o_exp_init = first.get('end2_export_initial', 0)
                o_exp_final = last.get('end2_export_final', 0)
                o_exp_mf = f.get('end2_export_mf', 1)
                o_exp_cons = (o_exp_final - o_exp_init) * o_exp_mf
                
                data["other_end"]["export"] = {
                    "initial": o_exp_init,
                    "final": o_exp_final,
                    "mf": o_exp_mf,
                    "consumption": o_exp_cons
                }
                
                # Calculate Losses
                # Q = (D+L-H-P)/(D+L)
                D = s_imp_cons
                L = o_imp_cons
                H = s_exp_cons
                P = o_exp_cons
                
                numerator = (D + L) - (H + P)
                denominator = (D + L)
                
                pct_loss = (numerator / denominator * 100) if denominator != 0 else 0
                
                data["stats"]["pct_loss"] = pct_loss
                
            else:
                # No data
                for end in ["shankarpally", "other_end"]:
                    for type_ in ["import", "export"]:
                        data[end][type_] = {"initial": 0, "final": 0, "mf": 0, "consumption": 0}
                data["stats"]["pct_loss"] = 0
                
            report_data.append(data)
            
        return report_data
        
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        with open("preview_error_log.txt", "w") as f:
            f.write(error_msg)
        print(error_msg)
        return JSONResponse(status_code=500, content={"detail": str(e)})

async def _generate_line_losses_report_wb(year: int, month: int):
    import calendar
    from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
    from openpyxl.utils import get_column_letter
    
    # --- Common Styles ---
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bold_font = Font(bold=True)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Line Losses"
    
    month_name = calendar.month_name[month]
    
    # 1. Headers
    # Row 1: Title
    ws.merge_cells('A1:T1')
    c = ws.cell(row=1, column=1, value=f"400KV SHANKARAPALLY SS- ENERGY LOSSES FOR THE MONTH {month_name}-{year}")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    # Row 2: Main Headers
    ws.merge_cells('A2:A3')
    c = ws.cell(row=2, column=1, value="Sl.\nNo.")
    c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('B2:B3')
    c = ws.cell(row=2, column=2, value="Name of the Feeder")
    c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('C2:J2')
    c = ws.cell(row=2, column=3, value="Shankarapally End")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('K2:R2')
    c = ws.cell(row=2, column=11, value="Other End")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('S2:S3')
    c = ws.cell(row=2, column=19, value="% Losses/\nCumulative\nLosses")
    c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('T2:T3')
    c = ws.cell(row=2, column=20, value="Remarks")
    c.alignment = center_align; c.border = thin_border
    
    # Row 3: Sub Headers (Import/Export)
    # Shankarpally
    ws.merge_cells('C3:F3')
    c = ws.cell(row=3, column=3, value="Import")
    c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('G3:J3')
    c = ws.cell(row=3, column=7, value="Export")
    c.alignment = center_align; c.border = thin_border
    
    # Other End
    ws.merge_cells('K3:N3')
    c = ws.cell(row=3, column=11, value="Import")
    c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('O3:R3')
    c = ws.cell(row=3, column=15, value="Export")
    c.alignment = center_align; c.border = thin_border
    
    # Row 4: Detailed Headers
    sub_headers = [
        "Initial Reading\nin MWH", "Final Reading in\nMWH", "MF", "Consumption",
        "Initial Reading\nin MWH", "Final Reading in\nMWH", "MF", "Consumption",
        "Initial Reading\nin MWH", "Final Reading in\nMWH", "MF", "Consumption",
        "Initial Reading\nin MWH", "Final Reading in\nMWH", "MF", "Consumption"
    ]
    
    for i, h in enumerate(sub_headers):
        cell = ws.cell(row=4, column=i+3, value=h)
        cell.alignment = center_align; cell.border = thin_border; cell.font = bold_font
        
    # Row 5: Column Letters/Formulas
    letters = ["A", "B", "C", "D=(B-A)*C", "E", "F", "G", "H=(F-E)*G", 
               "I", "J", "K", "L=(J-I)*K", "M", "N", "O", "P=(N-M)*O"]
    
    for i, l in enumerate(letters):
        cell = ws.cell(row=5, column=i+3, value=l)
        cell.alignment = center_align; cell.border = thin_border; cell.font = bold_font
        
    ws.cell(row=5, column=19, value="Q=(D+L-H-P)/(D+L)").alignment = center_align
    ws.cell(row=5, column=19).border = thin_border
    ws.cell(row=5, column=20, value="").border = thin_border # Remarks
    
    # 2. Fetch Data
    all_feeders = await db.feeders.find({}, {"_id": 0}).to_list(100)
    
    FEEDER_MAPPING = {
        "400 KV Shanakrapally-MHRM-2": "400KV MAHESHWARAM-2",
        "400 KV Shanakrapally-MHRM-1": "400KV MAHESHWARAM-1",
        "400 KV Shanakrapally-Narsapur-1": "400KV NARSAPUR-1",
        "400 KV Shanakrapally-Narsapur-2": "400KV NARSAPUR-2",
        "400 KV KethiReddyPally-1": "400KV KETHIREDDYPALLY-1",
        "400 KV KethiReddyPally-2": "400KV KETHIREDDYPALLY-2",
        "220 KV Parigi-1": "220KV PARIGI-1",
        "220 KV Parigi-2": "220KV PARIGI-2",
        "220 KV Tandur": "220KV THANDUR",
        "220 KV Gachibowli-1": "220KV GACHIBOWLI-1",
        "220 KV Gachibowli-2": "220KV GACHIBOWLI-2",
        "220 KV KethiReddyPally": "220KV KETHIREDDYPALLY",
        "220 KV Yeddumailaram-1": "220KV YEDDUMAILARAM-1",
        "220 KV Yeddumailaram-2": "220KV YEDDUMAILARAM-2",
        "220 KV Sadasivapet-1": "220KV SADASIVAPET-1",
        "220 KV Sadasivapet-2": "220KV SADASIVAPET-2"
    }

    FEEDER_ORDER = [
        "400KV MAHESHWARAM-2", "400KV MAHESHWARAM-1", "400KV NARSAPUR-1", "400KV NARSAPUR-2",
        "400KV KETHIREDDYPALLY-1", "400KV KETHIREDDYPALLY-2", "220KV PARIGI-1", "220KV PARIGI-2",
        "220KV THANDUR", "220KV GACHIBOWLI-1", "220KV GACHIBOWLI-2", "220KV KETHIREDDYPALLY",
        "220KV YEDDUMAILARAM-1", "220KV YEDDUMAILARAM-2", "220KV SADASIVAPET-1", "220KV SADASIVAPET-2"
    ]
    
    target_feeders = []
    for f in all_feeders:
        if f['name'] in FEEDER_MAPPING:
            f['display_name'] = FEEDER_MAPPING[f['name']]
            target_feeders.append(f)

    target_feeders.sort(key=lambda x: FEEDER_ORDER.index(x['display_name']) if x['display_name'] in FEEDER_ORDER else 999)
    
    start_date = f"{year}-{month:02d}-01"
    if month == 12: end_date = f"{year + 1}-01-01"
    else: end_date = f"{year}-{month + 1:02d}-01"
        
    entries = await db.entries.find(
        {"date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).to_list(5000)
    
    entries_by_feeder = {}
    for e in entries:
        if e['feeder_id'] not in entries_by_feeder:
            entries_by_feeder[e['feeder_id']] = []
        entries_by_feeder[e['feeder_id']].append(e)
        
    row_idx = 6
    for idx, f in enumerate(target_feeders):
        f_entries = entries_by_feeder.get(f['id'], [])
        f_entries.sort(key=lambda x: x['date'])
        
        vals = [0] * 16 # 4 groups of 4
        pct_loss = 0
        
        if f_entries:
            first = f_entries[0]
            last = f_entries[-1]
            
            # S-Imp
            si_init = first.get('end1_import_initial', 0) or 0
            si_final = last.get('end1_import_final', 0) or 0
            si_mf = f.get('end1_import_mf', 1) or 1
            si_cons = (si_final - si_init) * si_mf
            
            # S-Exp
            se_init = first.get('end1_export_initial', 0) or 0
            se_final = last.get('end1_export_final', 0) or 0
            se_mf = f.get('end1_export_mf', 1) or 1
            se_cons = (se_final - se_init) * se_mf
            
            # O-Imp
            oi_init = first.get('end2_import_initial', 0) or 0
            oi_final = last.get('end2_import_final', 0) or 0
            oi_mf = f.get('end2_import_mf', 1) or 1
            oi_cons = (oi_final - oi_init) * oi_mf
            
            # O-Exp
            oe_init = first.get('end2_export_initial', 0) or 0
            oe_final = last.get('end2_export_final', 0) or 0
            oe_mf = f.get('end2_export_mf', 1) or 1
            oe_cons = (oe_final - oe_init) * oe_mf
            
            vals = [
                si_init, si_final, si_mf, si_cons,
                se_init, se_final, se_mf, se_cons,
                oi_init, oi_final, oi_mf, oi_cons,
                oe_init, oe_final, oe_mf, oe_cons
            ]
            
            D = si_cons; L = oi_cons; H = se_cons; P = oe_cons
            numerator = (D + L) - (H + P)
            denominator = (D + L)
            pct_loss = (numerator / denominator * 100) if denominator != 0 else 0

        # Write Row
        ws.cell(row=row_idx, column=1, value=idx+1).border = thin_border
        ws.cell(row=row_idx, column=2, value=f['display_name']).border = thin_border
        
        for i, v in enumerate(vals):
            cell = ws.cell(row=row_idx, column=i+3, value=v)
            cell.number_format = '0.00'
            cell.border = thin_border
            cell.alignment = center_align
            
        cell = ws.cell(row=row_idx, column=19, value=pct_loss)
        cell.number_format = '0.00'
        cell.border = thin_border
        cell.alignment = center_align
        
        ws.cell(row=row_idx, column=20, value="-").border = thin_border
        
        row_idx += 1
        
    # Auto-fit
    for i in range(1, ws.max_column + 1):
        col_letter = get_column_letter(i)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.row <= 5: continue
            try:
                if cell.value and len(str(cell.value)) > max_len:
                    max_len = len(str(cell.value))
            except: pass
        ws.column_dimensions[col_letter].width = max(max_len + 2, 10)
    
    return wb

@api_router.get("/reports/line-losses/export/{year}/{month}")
async def export_line_losses_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        import calendar
        wb = await _generate_line_losses_report_wb(year, month)
        month_name = calendar.month_name[month]
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"Line_Losses_{month_name}_{year}.xlsx"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        with open("line_losses_export_error.txt", "w") as f:
            f.write(error_msg)
        print(error_msg)
        return JSONResponse(status_code=500, content={"detail": str(e)})

# --- PTR Max-Min (Format-1) Endpoints ---

@api_router.get("/reports/ptr-max-min-format1/preview/{year}/{month}")
async def get_ptr_max_min_preview(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"

    # Fetch ICT feeders
    ict_feeders = await db.max_min_feeders.find({"type": "ict_feeder"}, {"_id": 0}).to_list(100)
    
    # Sort ICT feeders
    ICT_ORDER = ["ICT-1 (315MVA)", "ICT-2 (315MVA)", "ICT-3 (315MVA)", "ICT-4 (500MVA)"]
    ict_feeders.sort(key=lambda x: ICT_ORDER.index(x['name']) if x['name'] in ICT_ORDER else 999)
    
    data = []
    
    for feeder in ict_feeders:
        entries = await db.max_min_entries.find(
            {"feeder_id": feeder['id'], "date": {"$gte": start_date, "$lt": end_date}},
            {"_id": 0}
        ).to_list(1000)
        
        if not entries:
            # Placeholder for missing data
            rating = "315"
            if "500" in feeder['name']: rating = "500"
            data.append({
                "district": "Rangareddy",
                "substation": "400/220 KV SHANKARPALLY",
                "ptr_kv": "400/220",
                "rating": rating,
                "general": {"mw": 0, "mvar": 0},
                "max": None,
                "min": None
            })
            continue
            
        # Stats
        max_mw = -1.0
        max_entry = None
        min_mw = float('inf')
        min_entry = None
        
        total_mw = 0.0
        total_mvar_avg = 0.0 # Derived from daily max MVAR
        count = 0
        
        for e in entries:
            d = e.get('data', {})
            
            # Max MW logic (Find entry with highest Max MW)
            try:
                curr_max_mw = float(d.get('max', {}).get('mw', 0) or 0)
            except:
                curr_max_mw = 0.0
                
            if curr_max_mw > max_mw:
                max_mw = curr_max_mw
                max_entry = e
            
            # Min MW logic (Find entry with lowest Min MW)
            try:
                curr_min_mw = float(d.get('min', {}).get('mw', 0) or 0)
            except:
                curr_min_mw = 0.0
            
            # Use strict comparison, assuming data is valid.
            if curr_min_mw < min_mw:
                min_mw = curr_min_mw
                min_entry = e
                
            # Avg Calculation
            try:
                day_avg_mw = float(d.get('avg', {}).get('mw', 0) or 0)
            except:
                day_avg_mw = 0.0
                
            try:
                day_max_mvar = float(d.get('max', {}).get('mvar', 0) or 0)
                # day_min_mvar = float(d.get('min', {}).get('mvar', 0) or 0)
                day_avg_mvar = day_max_mvar # User requested average of full-month MAX MVAR values
            except:
                day_avg_mvar = 0.0
                
            total_mw += day_avg_mw
            total_mvar_avg += day_avg_mvar
            count += 1
            
        avg_mw = total_mw / count if count > 0 else 0
        avg_mvar = total_mvar_avg / count if count > 0 else 0
        
        # Extract Max Details
        max_details = {}
        if max_entry:
            md = max_entry.get('data', {}).get('max', {})
            mw_val = float(md.get('mw', 0) or 0)
            mvar_val = float(md.get('mvar', 0) or 0)
            max_details = {
                "date": max_entry['date'],
                "time": format_time(md.get('time', '')),
                "mw": mw_val,
                "mvar": mvar_val,
                "mva": (mw_val**2 + mvar_val**2)**0.5
            }
        else:
             max_details = {"date": "-", "time": "-", "mw": 0, "mvar": 0, "mva": 0}
        
        # Extract Min Details
        min_details = {}
        if min_entry and min_mw != float('inf'):
            md = min_entry.get('data', {}).get('min', {})
            mw_val = float(md.get('mw', 0) or 0)
            mvar_val = float(md.get('mvar', 0) or 0)
            min_details = {
                "date": min_entry['date'],
                "time": format_time(md.get('time', '')),
                "mw": mw_val,
                "mvar": mvar_val,
                "mva": (mw_val**2 + mvar_val**2)**0.5
            }
        else:
            min_details = {"date": "-", "time": "-", "mw": 0, "mvar": 0, "mva": 0}
            
        rating = "315"
        if "500" in feeder['name']: rating = "500"
        
        data.append({
            "district": "Rangareddy",
            "substation": "400/220 KV SHANKARPALLY",
            "ptr_kv": "400/220",
            "rating": rating,
            "general": {
                "mw": avg_mw,
                "mvar": avg_mvar
            },
            "max": max_details,
            "min": min_details
        })
        
    return data

@api_router.get("/reports/ptr-max-min-format1/export/{year}/{month}")
async def _generate_ptr_max_min_report_wb(year: int, month: int, current_user: User):
    data = await get_ptr_max_min_preview(year, month, current_user)
    
    month_name = calendar.month_name[month]
    
    wb = Workbook()
    ws = wb.active
    ws.title = "PTR Max-Min Format-1"
    
    # Styles
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Headers
    ws.merge_cells('A1:T1')
    c = ws.cell(row=1, column=1, value="FORMAT-I")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('A2:A3')
    c = ws.cell(row=2, column=1, value="ZONE: CE/METRO/HYD")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('B2:R2')
    c = ws.cell(row=2, column=2, value=f"NORMAL & MAXIMUM / MINIMUM LOADING ON PTRs FOR THE MONTH OF {month_name}-{year} in SE/OMC/Metro-West Circle")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('S2:S4')
    c = ws.cell(row=2, column=19, value=f"MD\nreached\nso far in\n{year}")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('T2:T4')
    c = ws.cell(row=2, column=20, value="MD\nreached\nso far")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    # Row 3
    ws.merge_cells('E3:F3')
    c = ws.cell(row=3, column=5, value="GENERAL\nLOADING")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('G3:L3')
    c = ws.cell(row=3, column=7, value="MAXIMUM LOAD DETAILS")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('M3:R3')
    c = ws.cell(row=3, column=13, value="MINIMUM LOAD DETAILS")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    # Row 4
    headers_r4 = [
        "DISTRICT\nNAME", "SUB-STATION\nNAME", "PTR\nHV/LV in\nKV", "PTR DETAILS/\nRATING IN\nMVA",
        "MW", "MVAR",
        "DATE", "TIME", "MW", "MVAR", "MVA", "BUS\nVOLTAGE IN\nKV @MAX",
        "DATE", "TIME", "MW", "MVAR", "MVA", "BUS VOLTAGE\nIN KV @MIN\nLOADING"
    ]
    
    for i, h in enumerate(headers_r4):
        c = ws.cell(row=4, column=1+i, value=h)
        c.font = bold_font; c.alignment = center_align; c.border = thin_border
        
    # Data
    row_idx = 5
    
    def set_style(c):
        c.border = thin_border
        c.alignment = center_align
        return c
        
    for item in data:
        set_style(ws.cell(row=row_idx, column=1, value=item['district']))
        set_style(ws.cell(row=row_idx, column=2, value=item['substation']))
        set_style(ws.cell(row=row_idx, column=3, value=item['ptr_kv']))
        set_style(ws.cell(row=row_idx, column=4, value=int(item['rating'])))
        
        set_style(ws.cell(row=row_idx, column=5, value=f"{item['general']['mw']:.2f}"))
        set_style(ws.cell(row=row_idx, column=6, value=f"{item['general']['mvar']:.2f}"))
        
        m = item['max']
        if m:
            set_style(ws.cell(row=row_idx, column=7, value=format_date(m['date'])))
            set_style(ws.cell(row=row_idx, column=8, value=m['time']))
            set_style(ws.cell(row=row_idx, column=9, value=f"{m['mw']:.2f}"))
            set_style(ws.cell(row=row_idx, column=10, value=f"{m['mvar']:.2f}"))
            set_style(ws.cell(row=row_idx, column=11, value=f"{m['mva']:.2f}"))
        else:
            for c in range(7, 12): set_style(ws.cell(row=row_idx, column=c, value="-"))

        set_style(ws.cell(row=row_idx, column=12, value=""))
        
        m = item['min']
        if m:
            set_style(ws.cell(row=row_idx, column=13, value=format_date(m['date'])))
            set_style(ws.cell(row=row_idx, column=14, value=m['time']))
            set_style(ws.cell(row=row_idx, column=15, value=f"{m['mw']:.2f}"))
            set_style(ws.cell(row=row_idx, column=16, value=f"{m['mvar']:.2f}"))
            set_style(ws.cell(row=row_idx, column=17, value=f"{m['mva']:.2f}"))
        else:
            for c in range(13, 18): set_style(ws.cell(row=row_idx, column=c, value="-"))

        set_style(ws.cell(row=row_idx, column=18, value=""))
        
        set_style(ws.cell(row=row_idx, column=19, value=""))
        set_style(ws.cell(row=row_idx, column=20, value=""))
        
        row_idx += 1
        
    # Merge District and Substation
    if row_idx > 5:
        ws.merge_cells(f'A5:A{row_idx-1}')
        ws.cell(row=5, column=1).alignment = center_align
        ws.merge_cells(f'B5:B{row_idx-1}')
        ws.cell(row=5, column=2).alignment = center_align
        
    # Auto-width
    for col in range(1, 21):
        ws.column_dimensions[get_column_letter(col)].width = 12
        
    return wb

@api_router.get("/reports/ptr-max-min-format1/export/{year}/{month}")
async def export_ptr_max_min_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        wb = await _generate_ptr_max_min_report_wb(year, month, current_user)
        month_name = calendar.month_name[month]
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"PTR_Max_Min_{month_name}_{year}.xlsx"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"Export error traceback: {error_msg}")
        logger.error(f"Export error: {error_msg}")
        return JSONResponse(status_code=500, content={"detail": error_msg})

@api_router.get("/reports/tl-max-loading-format4/preview/{year}/{month}")
async def get_tl_max_loading_preview(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year + 1}-01-01"
    else:
        end_date = f"{year}-{month + 1:02d}-01"

    # Define Feeder Order and Mapping
    TL_ORDER = [
        # 400KV
        "400KV MAHESHWARAM-2", "400KV MAHESHWARAM-1", 
        "400KV NARSAPUR-1", "400KV NARSAPUR-2",
        "400KV KETHIREDDYPALLY-1", "400KV KETHIREDDYPALLY-2",
        "400KV NIZAMABAD-1", "400KV NIZAMABAD-2",
        # 220KV
        "220KV PARIGI-1", "220KV PARIGI-2",
        "220KV THANDUR",
        "220KV GACHIBOWLI-1", "220KV GACHIBOWLI-2",
        "220KV KETHIREDDYPALLY",
        "220KV YEDDUMAILARAM-1", "220KV YEDDUMAILARAM-2",
        "220KV SADASIVAPET-1", "220KV SADASIVAPET-2"
    ]
    
    DISPLAY_NAMES = {
        "400KV MAHESHWARAM-2": "Shankraplly - Maheshwaram - II",
        "400KV MAHESHWARAM-1": "Shankraplly - Maheshwaram - I",
        "400KV NARSAPUR-1": "Shankraplly - Narsapur - I",
        "400KV NARSAPUR-2": "Shankraplly - Narsapur - II",
        "400KV KETHIREDDYPALLY-1": "Shankraplly -Kethireddypally-1",
        "400KV KETHIREDDYPALLY-2": "Shankraplly - Kethireddypally-2",
        "400KV NIZAMABAD-1": "Shankraplly - Nizamabad - I",
        "400KV NIZAMABAD-2": "Shankraplly - Nizamabad - II",
        "220KV PARIGI-1": "Shankarpally - Parigi-1",
        "220KV PARIGI-2": "Shankarpally - Parigi-2",
        "220KV THANDUR": "Shankarpally - Thandur",
        "220KV GACHIBOWLI-1": "Shankarpally - Gachibowli-1",
        "220KV GACHIBOWLI-2": "Shnakarpally - Gachibowli-2",
        "220KV KETHIREDDYPALLY": "Shankarpally -Kethireddypally",
        "220KV YEDDUMAILARAM-1": "Shankarpally - Yeddumailaram-1",
        "220KV YEDDUMAILARAM-2": "Shankarpally - Yeddumailaram-2",
        "220KV SADASIVAPET-1": "Shankarpally - Sadasivapet-1",
        "220KV SADASIVAPET-2": "Shankarpally - Sadasivapet-2"
    }

    # Fetch all feeders (needed for group logic)
    all_db_feeders = await db.max_min_feeders.find({}, {"_id": 0}).to_list(1000)
    all_feeder_map = {f['name']: f for f in all_db_feeders}
    all_feeder_id_map = {f['id']: f for f in all_db_feeders}

    # Fetch all entries for the month
    all_entries = await db.max_min_entries.find(
        {"date": {"$gte": start_date, "$lt": end_date}},
        {"_id": 0}
    ).to_list(10000)

    # Group entries by feeder_id
    entries_by_feeder = {}
    for e in all_entries:
        fid = e.get('feeder_id')
        if fid:
            if fid not in entries_by_feeder:
                entries_by_feeder[fid] = []
            entries_by_feeder[fid].append(e)
    
    data = []
    
    for idx, feeder_name in enumerate(TL_ORDER):
        feeder = all_feeder_map.get(feeder_name)
        
        mw_val = ""
        mvar_val = ""
        date_val = ""
        time_val = ""
        
        if feeder:
            feeder_entries = entries_by_feeder.get(feeder['id'], [])
            
            # Check for Grouping Logic (Pair Feeder Logic)
            is_special, partner_names = get_feeder_group_info(feeder['name'])
            
            if is_special:
                p_group_map = {}
                p_group_map[feeder['id']] = feeder_entries
                
                # Gather partner entries
                for pname in partner_names:
                    p_feeder = all_feeder_map.get(pname)
                    if p_feeder:
                         p_group_map[p_feeder['id']] = entries_by_feeder.get(p_feeder['id'], [])
                
                # Determine Leader for the Full Month
                leader_id = determine_leader(p_group_map)
                if not leader_id: leader_id = feeder['id']
                
                leader_entries = p_group_map.get(leader_id, [])
                
                # Calculate Stats
                # Note: 'type' might be missing in feeder obj if not fetched properly, default to 'feeder'
                ftype = feeder.get('type', 'feeder') 
                stats = calculate_coincident_stats(leader_entries, feeder['id'], p_group_map, ftype)
                
                # Extract values
                try:
                    max_mw = stats.get('max_mw', '-')
                    if max_mw != '-' and max_mw is not None:
                        mw_val = f"{float(max_mw):.2f}"
                        date_val = format_date(stats.get('max_mw_date', ''))
                        time_val = format_time(stats.get('max_mw_time', ''))
                except:
                    pass
            else:
                # Standard Logic (Individual Max)
                ftype = feeder.get('type', 'feeder')
                stats = calculate_standard_stats(feeder_entries, ftype)
                
                try:
                    max_mw = stats.get('max_mw', '-')
                    if max_mw != '-' and max_mw is not None:
                        mw_val = f"{float(max_mw):.2f}"
                        date_val = format_date(stats.get('max_mw_date', ''))
                        time_val = format_time(stats.get('max_mw_time', ''))
                except:
                    pass

        display_name = DISPLAY_NAMES.get(feeder_name, feeder_name)
        voltage = "400 KV" if "400" in feeder_name else "220KV"
        
        data.append({
            "sl_no": idx + 1,
            "district": "Ranga Reddy",
            "voltage": voltage,
            "substation": "Shankarpally",
            "line_name": display_name,
            "mw": mw_val,
            "mvar": mvar_val,
            "date": date_val,
            "time": time_val,
            "md_2026": "",
            "md_so_far": "",
            "remarks": ""
        })
            
    return data

@api_router.get("/reports/tl-max-loading-format4/export/{year}/{month}")
async def _generate_tl_max_loading_report_wb(year: int, month: int, current_user: User):
    print(f"Exporting TL Max Loading for {year}-{month}")
    data = await get_tl_max_loading_preview(year, month, current_user)
    print(f"Data fetched: {len(data)} rows")
    month_name = calendar.month_name[month]
    
    wb = Workbook()
    ws = wb.active
    ws.title = "TL Max Loading Format-4"
    
    # Styles
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Header Row 1
    ws.merge_cells('A1:L1')
    c = ws.cell(row=1, column=1, value=f"FORMAT-IV - SE/OMC/Metro-West Circle for the Month of {month_name}-{year}")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    # Header Row 2
    ws.merge_cells('A2:D2')
    c = ws.cell(row=2, column=1, value="ZONE CE/METRO/HYD")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    ws.merge_cells('F2:I2')
    c = ws.cell(row=2, column=6, value="MAXIMUM")
    c.font = bold_font; c.alignment = center_align; c.border = thin_border
    
    headers = [
        "SL.No", "DISTRICT\nNAME", "VOLTAGE\nLEVEL IN KV", "SUBSTATION\nNAME",
        "TRANSMISSION LINE NAME",
        "MW", "MVAR", "DATE", "TIME",
        "MD\nreached\nin 2026", "MD\nreached\nso far", "REMARKS"
    ]
    
    for i, h in enumerate(headers):
        c = ws.cell(row=3, column=1+i, value=h)
        c.font = bold_font; c.alignment = center_align; c.border = thin_border
        
    # Data Rows
    row_idx = 4
    
    def set_style(c):
        c.border = thin_border
        c.alignment = center_align
        return c
        
    for item in data:
        set_style(ws.cell(row=row_idx, column=1, value=item['sl_no']))
        set_style(ws.cell(row=row_idx, column=2, value=item['district']))
        set_style(ws.cell(row=row_idx, column=3, value=item['voltage']))
        set_style(ws.cell(row=row_idx, column=4, value=item['substation']))
        set_style(ws.cell(row=row_idx, column=5, value=item['line_name']))
        set_style(ws.cell(row=row_idx, column=6, value=item['mw']))
        set_style(ws.cell(row=row_idx, column=7, value=item['mvar']))
        set_style(ws.cell(row=row_idx, column=8, value=item['date']))
        set_style(ws.cell(row=row_idx, column=9, value=item['time']))
        set_style(ws.cell(row=row_idx, column=10, value=item['md_2026']))
        set_style(ws.cell(row=row_idx, column=11, value=item['md_so_far']))
        set_style(ws.cell(row=row_idx, column=12, value=item['remarks']))
        row_idx += 1
        
    # Merging Logic
    # District (Col 2), Substation (Col 4) -> Merge for all rows
    if row_idx > 4:
        # District
        ws.merge_cells(f'B4:B{row_idx-1}')
        
        # Substation
        ws.merge_cells(f'D4:D{row_idx-1}')
        
        # Voltage (Col 3) - Split into 400KV and 220KV blocks
        num_400 = 8
        num_220 = 10
        
        start_row = 4
        if len(data) >= num_400:
            ws.merge_cells(f'C{start_row}:C{start_row+num_400-1}')
        
        if len(data) >= num_400 + num_220:
            ws.merge_cells(f'C{start_row+num_400}:C{start_row+num_400+num_220-1}')
            
    # Column Widths
    ws.column_dimensions['E'].width = 30 # Line Name
    for col in [1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12]:
         ws.column_dimensions[get_column_letter(col)].width = 15

    return wb

@api_router.get("/reports/tl-max-loading-format4/export/{year}/{month}")
async def export_tl_max_loading_report(
    year: int,
    month: int,
    current_user: User = Depends(get_current_user)
):
    try:
        wb = await _generate_tl_max_loading_report_wb(year, month, current_user)
        month_name = calendar.month_name[month]
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"TL_Max_Loading_{month_name}_{year}.xlsx"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"Export error traceback: {error_msg}")
        logger.error(f"Export error: {error_msg}")
        # Write to file for debugging
        try:
            with open("export_error.log", "w") as f:
                f.write(error_msg)
        except:
            pass
        return JSONResponse(status_code=500, content={"detail": error_msg})

@api_router.post("/reports/send-mail")
async def send_reports_email_endpoint(
    request: EmailReportRequest,
    current_user: User = Depends(get_current_user)
):
    try:
        year = request.year
        month = request.month
        recipient_email = request.email
        month_name = calendar.month_name[month]
        report_ids = request.report_ids
        
        attachments = []
        errors = []
        
        # 1. Fortnight Report
        if not report_ids or 'fortnight' in report_ids:
            try:
                wb = await _generate_fortnight_report_wb(year, month)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"Fortnight_Report_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating Fortnight Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)
            
        # 2. Energy Consumption (Optional, not currently in frontend list)
        if report_ids and 'energy-consumption' in report_ids:
            try:
                wb = await _generate_energy_export_wb(year, month)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"Energy_Consumption_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating Energy Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)

        # 3. Boundary Meter Report
        if not report_ids or 'boundary-meter' in report_ids:
            try:
                wb = await _generate_boundary_meter_wb(year, month)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"Boundary_Meter_Report_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating Boundary Meter Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)
             
        # 4. KPI Report
        if not report_ids or 'kpi' in report_ids:
            try:
                wb = await _generate_kpi_report_wb(year, month)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"KPI_Report_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating KPI Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)
             
        # 5. Line Losses
        if not report_ids or 'line-losses' in report_ids:
            try:
                wb = await _generate_line_losses_report_wb(year, month)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"Line_Losses_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating Line Losses Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)
             
        # 6. Daily Max MVA
        if not report_ids or 'daily-max-mva' in report_ids:
            try:
                wb = await _generate_daily_max_mva_wb(year, month)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"Daily_Max_MVA_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating Daily Max MVA Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)
             
        # 7. PTR Max Min
        if not report_ids or 'ptr-max-min' in report_ids:
            try:
                wb = await _generate_ptr_max_min_report_wb(year, month, current_user)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"PTR_Max_Min_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating PTR Max Min Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)

        # 8. TL Max Loading
        if not report_ids or 'tl-max-loading' in report_ids:
            try:
                wb = await _generate_tl_max_loading_report_wb(year, month, current_user)
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                attachments.append((f"TL_Max_Loading_{month_name}_{year}.xlsx", output.read()))
            except Exception as e:
                import traceback
                error_msg = f"Error generating TL Max Loading Report: {str(e)}\n{traceback.format_exc()}\n"
                print(error_msg)
                errors.append(error_msg)
        
        if errors:
            error_content = "\n".join(errors)
            attachments.append(("generation_errors.txt", error_content.encode('utf-8')))
             
        if not attachments:
            raise HTTPException(status_code=400, detail="No reports could be generated for this period.")
            
        # Send Email in background
        import asyncio
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, send_reports_email, recipient_email, attachments)
        
        return {"message": f"Reports sent successfully to {recipient_email}"}
        
    except HTTPException as he:
        raise he
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"Send mail error: {error_msg}")
        return JSONResponse(status_code=500, content={"detail": str(e)})

app.include_router(api_router)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()
