from fastapi import FastAPI
from pydantic import BaseModel, Field
from typing import List, Optional


app = FastAPI()


class Proposal(BaseModel):
    proposal_number: str
    client_name: str
    seller_id: str
    payment_method: str
    delivery_date: str
    notes: str


class Item(BaseModel):
    item_index: int
    item_name: str = Field(max_length=60)
    item_subtitle: str = Field(max_length=80)
    item_description: str = Field(max_length=300)
    item_code: str = Field(max_length=30)
    quantity: int
    unit_price: float
    item_image_url: str


class Section(BaseModel):
    section_index: int
    freight_value: Optional[float] = None
    freight_label: str
    items: List[Item] = Field(min_length=1, max_length=10)


class GenerateRequest(BaseModel):
    proposal: Proposal
    sections: List[Section] = Field(min_length=1)


@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate")
def generate_proposal(payload: GenerateRequest):
    return {
        "status": "success",
        "message": "Payload validado com sucesso",
        "sections_count": len(payload.sections),
        "items_count": sum(len(section.items) for section in payload.sections),
    }
