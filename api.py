from datetime import datetime
from email.mime.text import MIMEText
import io
import os
import re
import shutil
import smtplib
import traceback
from typing import List
import zipfile
from PyPDF2 import PdfReader
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Response, UploadFile
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from bson import ObjectId
import pymongo
import qrcode
import uvicorn
import pandas as pd
from fastapi.responses import FileResponse
from model import CostingItem, User

app = FastAPI()

# MongoDB connection
client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["kiosk"]
collection = db["costing"]
EXPORT_FOLDER = "exports"
QR_FOLDER = "exports/QR"
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(QR_FOLDER, exist_ok=True)

load_dotenv()
SMTP_SERVER = os.getenv("SMTP_SERVER")
USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")

# Function to count pages in a PDF file
def count_pdf_pages(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        return len(reader.pages)
    except Exception as e:
        print(f"Error counting pages in PDF: {e}")
        return 0
    
# Function to count pages in a PDF file
def count_pdf_pages(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        return len(reader.pages)
    except Exception as e:
        print(f"Error counting pages in PDF: {e}")
        return 0
    
@app.post("/add_item/")
def add_item(item: CostingItem):
    try:
        # Insert data into MongoDB
        collection.insert_one(item.dict())
        return {"message": "Data inserted successfully."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred: {e}")

@app.put("/edit_item/{item_id}")
def edit_item(item_id: str, item: CostingItem):
    try:
        if not ObjectId.is_valid(item_id):
            raise HTTPException(status_code=400, detail="Invalid ObjectId format.")

        # Update the item in MongoDB
        result = collection.update_one(
            {"_id": ObjectId(item_id)},
            {"$set": item.dict()}
        )
        if result.modified_count == 0:
            raise HTTPException(status_code=404, detail="Item not found or data is identical to existing.")
        return {"message": "Data updated successfully."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred: {e}")
    
@app.post("/submit/")
async def submit_user_data(
    name: str = Form(...),
    phone: str = Form(...),
    email: str = Form(...),
    description: str = Form(...),
    transaction_id: str = Form(...),
    location: str = Form(...),
    printing_type: str = Form(...),
    binding_type: str = Form(...),
    pdf_files: List[UploadFile] = None
):
    folder_path = "temp_pdfs"
    os.makedirs(folder_path, exist_ok=True)

    total_pages = 0
    saved_files = []

    for pdf_file in pdf_files:
        file_path = os.path.join(folder_path, pdf_file.filename)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(pdf_file.file, buffer)
        total_pages += count_pdf_pages(file_path)
        saved_files.append(file_path)

    printing_type_doc = collection.find_one({"field_name": "Printing type", "name": printing_type})
    binding_type_doc = collection.find_one({"field_name": "Binding and Finishing", "name": binding_type})

    if not printing_type_doc or not binding_type_doc:
        raise HTTPException(status_code=400, detail="Invalid printing or binding type")

    total_printing_cost = total_pages * printing_type_doc['cost']
    total_binding_cost = binding_type_doc['cost'] * len(saved_files)
    total_cost = total_printing_cost + total_binding_cost

    customer_id = db.customers.count_documents({}) + 1
    customer_folder = os.path.join("output", str(customer_id))
    os.makedirs(customer_folder, exist_ok=True)

    final_saved_files = []
    for idx, file_path in enumerate(saved_files, 1):
        new_file_name = f"{customer_id}.{idx} {os.path.basename(file_path)}"
        new_file_path = os.path.join(customer_folder, new_file_name)
        shutil.move(file_path, new_file_path)
        final_saved_files.append(new_file_path)

    user_data = User(
        time_stamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        name=name,
        phone=phone,
        email=email,
        description=description,
        transaction_id=transaction_id,
        total_pdfs=len(pdf_files),
        total_pages=total_pages,
        printing_type=printing_type,
        printing_cost_per_page=printing_type_doc['cost'],
        location=location,
        binding_and_finishing=binding_type,
        total_cost=total_cost,
        files=final_saved_files,
        is_printed=False
    )

    db.customers.insert_one(user_data.dict())
    return JSONResponse(content={"message": "User data saved successfully.", "total_cost": total_cost})

@app.get("/generate-excel/")
def generate_excel():
    collection = db["customers"]
    # Fetch all unprinted records
    unprinted_records = list(collection.find({"is_printed": False}))
    print("Unprinted records:", unprinted_records)
    
    if not unprinted_records:
        raise HTTPException(status_code=404, detail="No unprinted records found.")
    
    # Extract relevant data
    data = []
    for idx,record in enumerate(unprinted_records, start=1):
        data.append({
            "SL_No": idx,
            "Timestamp": record["time_stamp"],
            "Name": record["name"],
            "Phone": record["phone"],
            "Email": record["email"],
            "Description": record["description"],
            "Transaction_ID": record["transaction_id"],
            "Total_PDFs": record["total_pdfs"],
            "Total_Pages": record["total_pages"],
            "Printing_Type": record["printing_type"],
            "Printing_Cost_Per_Page": record["printing_cost_per_page"],
            "Location": record["location"],
            "Binding_and_Finishing": record["binding_and_finishing"],
            "Total_Cost": record["total_cost"],
            "Files": ", ".join(record["files"])
            
        })
    
    # Create a DataFrame and save to an Excel file
    df = pd.DataFrame(data)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_filename = f"unprinted_records_{timestamp}.xlsx"
    excel_filepath = os.path.join("exports", excel_filename)
    os.makedirs("exports", exist_ok=True)
    df.to_excel(excel_filepath, index=False)
    
    # Update the is_printed field to True
    collection.update_many({"is_printed": False}, {"$set": {"is_printed": True}})
    
    return FileResponse(excel_filepath, filename=excel_filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.put("/cleanup-printed-files/")
def cleanup_printed_files():
    collection = db["customers"]
    # Fetch all records where is_printed = True
    printed_records = list(collection.find({"is_printed": True}))
    print(f"Found {len(printed_records)} printed records for cleanup.")

    if not printed_records:
        raise HTTPException(status_code=404, detail="No printed records found for cleanup.")

    updated_count = 0

    for record in printed_records:
        file_list = record.get("files", [])

        if file_list:
            # Update MongoDB: Keep only the file names, remove file contents if stored
            collection.update_one(
                {"_id": record["_id"]}, 
                {"$set": {"files": file_list}, "$unset": {"file_contents": ""}}  # Remove file contents if present
            )
            updated_count += 1

    return {
        "message": f"Updated {updated_count} records. Only file names are retained."
    }

@app.post("/split-excel-by-location/")
async def split_excel_by_location(file: UploadFile = File(...)):
    try:
        # Read the uploaded Excel file
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))

        # Check if 'Location' column exists
        if "Location" not in df.columns:
            raise HTTPException(status_code=400, detail="'Location' column not found in the Excel file.")
        
        # Create an 'exports' folder if it doesn't exist
        export_folder = "exports"
        os.makedirs(export_folder, exist_ok=True)

        # Get unique locations
        unique_locations = df["Location"].unique()

        # Dictionary to store file paths for returning
        file_paths = []

        # Split Excel file based on Location
        for location in unique_locations:
            location_df = df[df["Location"] == location]
            
            # Save the new Excel file in the exports folder
            file_name = f"{location}_records.xlsx"
            file_path = os.path.join(export_folder, file_name)
            location_df.to_excel(file_path, index=False)
            file_paths.append(file_path)

        # Create a ZIP file of all location-based Excel files
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        zip_filename = f"split_records_{timestamp}.zip"
        zip_filepath = os.path.join(export_folder, zip_filename)

        with zipfile.ZipFile(zip_filepath, 'w') as zipf:
            for file_path in file_paths:
                zipf.write(file_path, os.path.basename(file_path))

        return FileResponse(zip_filepath, filename=zip_filename, media_type="application/zip")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred: {str(e)}")
    
# Function to send an email
def send_email(to_address, subject, body):
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, 465) as server:
            server.login(USERNAME, PASSWORD)
            msg = MIMEText(body)
            msg["Subject"] = subject
            msg["From"] = USERNAME
            msg["To"] = to_address
            server.sendmail(USERNAME, to_address, msg.as_string())
        return True
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")
        return False

@app.post("/send-emails/")
async def send_emails(
    file: UploadFile, 
    subject: str = Form(...), 
    email_message: str = Form(...)
):
    try:
        # Read Excel file
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        #df.columns = df.columns.str.lower()  # Normalize column names
        
        # Check if "email" column exists
        if "Email" not in df.columns:
            return JSONResponse(
                status_code=400, 
                content={"error": "'Email' column not found in the Excel file."}
            )

        df["Email_send"] = "Failed"  # Default status
        
        # Process each row
        for index, row in df.iterrows():
            email = row["Email"]
            customized_message = email_message
            
            # Replace placeholders with actual values
            for column_name in df.columns:
                placeholder = f"{{{column_name}}}"
                if placeholder in customized_message:
                    customized_message = customized_message.replace(placeholder, str(row[column_name]))
            
            # Send email and update status
            if send_email(email, subject, customized_message):
                df.at[index, "Email_send"] = "Sent"

        # Ensure export folder exists
        os.makedirs("exports", exist_ok=True)
        
        # Save updated Excel file
        export_path = f"exports/{file.filename}"
        df.to_excel(export_path, index=False)

        return {"message": "Emails sent successfully!", "file_path": export_path}
    
    except Exception as e:
        error_message = f"An error occurred: {str(e)}\n\n{traceback.format_exc()}"
        print(error_message)
        return JSONResponse(
            status_code=500, 
            content={"error": error_message}
        )
    
# Function to get the next serial number
def get_next_serial():
    existing_files = [f for f in os.listdir(QR_FOLDER) if f.startswith("Kiosk_QR(") and f.endswith(").png")]
    serial_numbers = [int(re.search(r"\((\d+)\)", f).group(1)) for f in existing_files if re.search(r"\((\d+)\)", f)]
    return max(serial_numbers) + 1 if serial_numbers else 1

# Function to generate a QR code
def generate_qr_code(data):
    serial_number = get_next_serial()
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill="black", back_color="white")
    
    qr_filename = os.path.join(QR_FOLDER, f"Kiosk_QR({serial_number}).png")
    img.save(qr_filename)
    return qr_filename

@app.post("/generate-qr/")
async def generate_qr(data: str = Form(...)):
    try:
        qr_filename = generate_qr_code(data)
        return {"message": "QR Code generated successfully!", "qr_file": qr_filename}
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": f"An error occurred: {str(e)}"},
        )

if __name__ == "__main__":
    uvicorn.run("api:app", host="127.0.0.1", port=8000, reload=True)
