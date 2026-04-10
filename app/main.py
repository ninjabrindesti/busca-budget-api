from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from pptx import Presentation
import uuid

from app.services.pptx_generator import replace_text_placeholders


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
    prs = Presentation("templates/template_ninja.pptx")

    replace_text_placeholders(prs, payload.proposal.model_dump())

    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    if title:
        title.text = f"Proposta {payload.proposal.proposal_number}"

    filename = f"proposta_{payload.proposal.proposal_number}_{str(uuid.uuid4())[:4]}.pptx"
    filepath = f"/tmp/{filename}"

    prs.save(filepath)

    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
