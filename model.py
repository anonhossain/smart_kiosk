from pydantic import BaseModel, EmailStr, Field
from typing import List
from datetime import datetime

# Pydantic model for request validation
class CostingItem(BaseModel):
    field_name: str
    name: str
    cost: float

class User(BaseModel):
    sl_no: int
    time_stamp: str
    name: str
    phone: str
    email: EmailStr
    description: str
    transaction_id: str
    total_pdfs: int
    total_pages: int
    printing_type: str
    printing_cost_per_page: float
    location: str
    binding_and_finishing: str
    total_cost: float
    files: List[str]  # URLs from Cloudinary
    is_printed: bool = Field(default=False)
    copy_num: int