from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from pydantic import BaseModel
from typing import List
import logic
import shutil
import os
import pandas as pd

app = FastAPI()

# ==========================================
# DATA MODELS
# ==========================================

class AuthRequest(BaseModel):
    email: str
    password: str

class SearchRequest(BaseModel):
    date: str  # Format: YYYY-MM-DD
    start_time: str  # Format: HH:MM
    end_time: str  # Format: HH:MM
    resource_type: str  # "Classroom" or "Lab"

class BookRequest(BaseModel):
    email: str
    resource_type: str
    resource_name: str
    month_name: str
    target_column: int
    slots_to_book: List[int]  # Row numbers to write to in Excel
    time_slot_labels: List[str]  # e.g., ["09:00 to 09:50", "09:50 to 10:40"]
    date: str
    purpose: str

class UtilityRequest(BaseModel):
    resource_type: str
    selected_resource: str  # e.g., "CR 1" or "All"

# ==========================================
# AUTHENTICATION ENDPOINTS
# ==========================================

@app.post("/auth/register")
def register(req: AuthRequest):
    return logic.register_user(req.email, req.password)

@app.post("/auth/login")
def login(req: AuthRequest):
    # Hardcoded Admin
    if req.email == "admin@adaniuni.ac.in" and req.password == "admin123":
        return {"role": "admin", "email": req.email}
    
    # Check regular users
    try:
        df = pd.read_excel(logic.USERS_FILE)
        user = df[(df["Email"] == req.email) & (df["Password"] == req.password)]
        if not user.empty:
            return {"role": "user", "email": req.email}
    except Exception:
        pass
    
    return {"error": "Invalid credentials"}

# ==========================================
# ADMIN ENDPOINTS
# ==========================================

@app.post("/admin/upload/classroom")
async def upload_classroom(
    file: UploadFile = File(...), 
    sheet_name: str = Form("FEST_Room Occupancy"), # Defaulting to what was in your notebook
    year: int = Form(2026)
):
    """Uploads the base classroom Excel file and generates the full year."""
    file_location = os.path.join(logic.DATA_DIR, "raw_classroom.xlsx")
    with open(file_location, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    success, msg = logic.generate_classroom_full_year(file_location, sheet_name, year)
    if not success:
        raise HTTPException(status_code=400, detail=msg)
        
    return {"message": msg}

@app.post("/admin/upload/lab")
async def upload_lab(
    file: UploadFile = File(...),
    year: int = Form(2026)
):
    """Uploads the base lab Excel file, merges it, and generates the full year."""
    file_location = os.path.join(logic.DATA_DIR, "raw_lab.xlsx")
    with open(file_location, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    # Step 1: Vertically merge
    merged_file_path = logic.generate_vertically_merged_lab(file_location)
    
    # Step 2: Generate full year
    success, msg = logic.generate_lab_full_year(merged_file_path, year)
    if not success:
        raise HTTPException(status_code=400, detail=msg)
        
    return {"message": msg}

@app.get("/admin/bookings")
def get_bookings():
    """Returns all booking logs for the admin dashboard."""
    if os.path.exists(logic.BOOKINGS_FILE) and os.path.getsize(logic.BOOKINGS_FILE) > 0:
        return pd.read_excel(logic.BOOKINGS_FILE).to_dict(orient="records")
    return []

# ==========================================
# USER ENDPOINTS (SEARCH, BOOK, UTILITY)
# ==========================================

@app.post("/search")
def search_slots(req: SearchRequest):
    """Searches for free slots in the respective timetable."""
    result = logic.search_free_slots(req.date, req.start_time, req.end_time, req.resource_type)
    if "error" in result:
        raise HTTPException(status_code=400, detail=result["error"])
    return result

@app.post("/book")
def book_slots(req: BookRequest):
    """Books the requested slots and logs the transaction."""
    # Write directly into the active timetable Excel file
    success = logic.book_slots_in_excel(
        req.slots_to_book, req.month_name, req.target_column, req.purpose, req.resource_type
    )
    
    if success:
        # Log the booking in bookings.xlsx for the admin to see
        time_slots_str = " | ".join(req.time_slot_labels)
        logic.log_booking(
            req.email, req.resource_type, req.resource_name, 
            time_slots_str, req.date, req.purpose
        )
        return {"message": "Booking successful"}
        
    raise HTTPException(status_code=500, detail="Booking failed due to an internal error.")

@app.get("/resources")
def get_resources(resource_type: str):
    """Returns a list of available classrooms or labs."""
    names = logic.get_resource_names(resource_type)
    return {"resources": names}

@app.post("/utility")
def get_utility(req: UtilityRequest):
    """Calculates occupied and free slots for charts."""
    result = logic.calculate_utility(req.resource_type, req.selected_resource)
    if "error" in result:
        raise HTTPException(status_code=400, detail=result["error"])
    return result