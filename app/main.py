from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from pptx import Presentation
import uuid
import os

from app.services.pptx_generator import (
    replace_text_placeholders_on_slide,
    replace_named_images_on_slide,
)
from app.services.pptx_com import duplicate_slide_in_file

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
    template_path = "templates/template_ninja.pptx"
    working_path = f"/tmp/proposta_work_{uuid.uuid4().hex}.pptx"

    if not os.path.exists(template_path):
        raise HTTPException(status_code=500, detail="Template não encontrado.")

    section = payload.sections[0]
    item_count = len(section.items)

    # slide 9 = modelo de item
    # se já existe 1 slide base e você quer N itens,
    # precisa duplicar N-1 vezes
    copies_needed = max(0, item_count - 1)

    duplicate_slide_in_file(
        input_path=template_path,
        output_path=working_path,
        source_slide_index=9,
        copies=copies_needed,
    )

    prs = Presentation(working_path)

    ITEM_SLIDE_START_INDEX = 8  # base 0 => slide 9

    for i, item in enumerate(section.items):
        slide = prs.slides[ITEM_SLIDE_START_INDEX + i]

        item_total = item.quantity * item.unit_price
        freight_value = section.freight_value or 0

        data = {
            "proposal_number": payload.proposal.proposal_number,
            "client_name": payload.proposal.client_name,
            "payment_method": payload.proposal.payment_method,
            "delivery_date": payload.proposal.delivery_date,
            "notes": payload.proposal.notes,
            "seller_name": "João Victor Ferrigno",
            "seller_phone": "(11) 99999-9999",
            "seller_email": "joao@empresa.com",
            "seller_description": "Teste vendedor",
            "item_name": item.item_name,
            "item_subtitle": item.item_subtitle,
            "item_index": str(item.item_index),
            "item_display_index": str(item.item_index + 1),
            "item_description": item.item_description,
            "item_code": item.item_code,
            "quantity": str(item.quantity),
            "unit_price": f"{item.unit_price:.2f}",
            "item_total": f"{item_total:.2f}",
            "section_total": f"{item_total:.2f}",
            "freight": f"{freight_value:.2f}",
            "freight_label": section.freight_label,
            "seller_image_url": "https://dummyimage.com/400x400/cccccc/000000.png&text=Seller",
            "item_image_url": item.item_image_url,
        }

        replace_text_placeholders_on_slide(slide, data)
        replace_named_images_on_slide(slide, data)

    filename = f"proposta_teste_{uuid.uuid4().hex[:8]}.pptx"
    final_path = f"/tmp/{filename}"

    prs.save(final_path)

    return FileResponse(
        path=final_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
