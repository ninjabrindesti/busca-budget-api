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

    first_section = payload.sections[0]
    first_item = first_section.items[0]

    item_total = first_item.quantity * first_item.unit_price
    freight_value = first_section.freight_value or 0

    data = {
        "proposal_number": payload.proposal.proposal_number,
        "client_name": payload.proposal.client_name,
        "payment_method": payload.proposal.payment_method,
        "delivery_date": payload.proposal.delivery_date,
        "notes": payload.proposal.notes,
        "seller_name": "João Victor Ferrigno",
        "seller_phone": "(11) 99999-9999",
        "seller_email": "joao@empresa.com",
        "seller_description": (
            "Sou apaixonado por construir conexões genuínas e gerar impacto positivo na vida das pessoas. "
            "Acredito que, mais do que atender, meu papel é entender profundamente a jornada do cliente, "
            "antecipar necessidades e ser um parceiro no sucesso deles. "
            "Sou apaixonado por construir conexões genuínas e gerar impacto positivo na vida das pessoas. "
            "Acredito que, mais do que atender, meu papel é entender profundamente a jornada do cliente, "
            "antecipar necessidades e ser um parceiro no sucesso deles."
        ),
        "item_name": first_item.item_name,
        "item_subtitle": first_item.item_subtitle,
        "item_index": str(first_item.item_index),
        "item_display_index": str(first_item.item_index + 1),
        "item_description": first_item.item_description,
        "item_code": first_item.item_code,
        "quantity": str(first_item.quantity),
        "unit_price": f"{first_item.unit_price:.2f}",
        "item_total": f"{item_total:.2f}",
        "section_total": f"{item_total:.2f}",
        "freight": f"{freight_value:.2f}",
    }

    replace_text_placeholders(prs, data)

    filename = f"proposta_{payload.proposal.proposal_number}_{str(uuid.uuid4())[:4]}.pptx"
    filepath = f"/tmp/{filename}"

    prs.save(filepath)

    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
