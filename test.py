import os
import shutil
from PyPDF2 import PdfReader
import pymongo
from datetime import datetime
from model import User

# MongoDB connection
client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["kiosk"]
collection = db["costing"]

# Function to count pages in a PDF file
def count_pdf_pages(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        return len(reader.pages)
    except Exception as e:
        print(f"Error counting pages in PDF: {e}")
        return 0

def get_user_input():
    locations = [loc['name'] for loc in collection.find({"field_name": "Location"})]
    printing_types = collection.find({"field_name": "Printing type"})
    binding_options = collection.find({"field_name": "Binding and Finishing"})
    
    folder_path = r"pdfs\\"
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    total_pages = sum(count_pdf_pages(os.path.join(folder_path, pdf)) for pdf in pdf_files)
    num_pdfs = len(pdf_files)

    print("\nSelect a Location:")
    for idx, location in enumerate(locations, 1):
        print(f"{idx}. {location}")
    selected_location = locations[int(input("Enter the number for the location: ")) - 1]

    printing_type_dict = {idx: {'name': pt['name'], 'cost': pt['cost']} for idx, pt in enumerate(printing_types, 1)}
    print("\nSelect a Printing Type:")
    for idx, pt in printing_type_dict.items():
        print(f"{idx}. {pt['name']} - Cost: {pt['cost']}")
    selected_printing_type = printing_type_dict[int(input("Enter the number for Printing Type: "))]
    total_printing_cost = total_pages * selected_printing_type['cost']

    binding_dict = {idx: {'name': bd['name'], 'cost': bd['cost']} for idx, bd in enumerate(binding_options, 1)}
    print("\nSelect a Binding and Finishing Type:")
    for idx, bd in binding_dict.items():
        print(f"{idx}. {bd['name']} - Cost: {bd['cost']}")
    selected_binding = binding_dict[int(input("Enter the number for Binding and Finishing: "))]
    total_binding_cost = selected_binding['cost'] * num_pdfs

    total_cost = total_printing_cost + total_binding_cost

    name = input("\nEnter your Name: ")
    phone = input("Enter your Phone: ")
    email = input("Enter your Email: ")
    description = input("Enter a Description: ")
    transaction_id = input("Enter Transaction ID: ")

    customer_id = db.customers.count_documents({}) + 1
    customer_folder = os.path.join("output", str(customer_id))
    os.makedirs(customer_folder, exist_ok=True)

    saved_files = []
    for idx, file_name in enumerate(pdf_files, 1):
        new_file_name = f"{customer_id}.{idx} {file_name}"
        new_file_path = os.path.join(customer_folder, new_file_name)
        shutil.copy(os.path.join(folder_path, file_name), new_file_path)
        saved_files.append(new_file_path)

    user_data = User(
        time_stamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        name=name,
        phone=phone,
        email=email,
        description=description,
        transaction_id=transaction_id,
        total_pdfs=num_pdfs,
        total_pages=total_pages,
        printing_type=selected_printing_type['name'],
        printing_cost_per_page=selected_printing_type['cost'],
        location=selected_location,
        binding_and_finishing=selected_binding['name'],
        total_cost=total_cost,
        files=saved_files,
        is_printed=False
    )

    db.customers.insert_one(user_data.dict())

if __name__ == "__main__":
    get_user_input()
