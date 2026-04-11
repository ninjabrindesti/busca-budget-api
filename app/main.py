from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from pptx import Presentation
import uuid

from app.services.pptx_generator import (
    replace_text_placeholders_on_slide,
    replace_named_images_on_slide,
    duplicate_slide,
)


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

    # Índices internos do PowerPoint começam em 0
    # Página 9 = índice 8
    # Página 10 = índice 9
    ITEM_SLIDE_INDEX = 8
    BUDGET_SLIDE_INDEX = 9

    seller_base_data = {
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
            "antecipar necessidades e ser um parceiro no sucesso deles."
        ),
        "seller_image_url": "https://dummyimage.com/400x400/cccccc/000000.png&text=Seller",
    }

    for section in payload.sections:
        section_total_value = sum(item.quantity * item.unit_price for item in section.items)
        freight_value = section.freight_value or 0

        # 1 slide de item por item
        for item in section.items:
            item_slide = duplicate_slide(prs, ITEM_SLIDE_INDEX)
            item_total = item.quantity * item.unit_price

            item_data = {
                **seller_base_data,
                "item_name": item.item_name,
                "item_subtitle": item.item_subtitle,
                "item_index": str(item.item_index),
                "item_display_index": str(item.item_index + 1),
                "item_description": item.item_description,
                "item_code": item.item_code,
                "quantity": str(item.quantity),
                "unit_price": f"{item.unit_price:.2f}",
                "item_total": f"{item_total:.2f}",
                "section_total": f"{section_total_value:.2f}",
                "freight": f"{freight_value:.2f}",
                "item_image_url": item.item_image_url,
            }

            replace_text_placeholders_on_slide(item_slide, item_data)
            replace_named_images_on_slide(item_slide, item_data)

        # 1 slide de orçamento consolidado por orçamento
        budget_slide = duplicate_slide(prs, BUDGET_SLIDE_INDEX)

        first_item = section.items[0]
        first_item_total = first_item.quantity * first_item.unit_price

        budget_data = {
            **seller_base_data,
            "item_name": first_item.item_name,
            "item_subtitle": first_item.item_subtitle,
            "item_index": str(first_item.item_index),
            "item_display_index": str(first_item.item_index + 1),
            "item_description": first_item.item_description,
            "item_code": first_item.item_code,
            "quantity": str(first_item.quantity),
            "unit_price": f"{first_item.unit_price:.2f}",
            "item_total": f"{first_item_total:.2f}",
            "section_total": f"{section_total_value:.2f}",
            "freight": f"{freight_value:.2f}",
            "item_image_url": first_item.item_image_url,
        }

        replace_text_placeholders_on_slide(budget_slide, budget_data)
        replace_named_images_on_slide(budget_slide, budget_data)

    filename = f"proposta_{payload.proposal.proposal_number}_{str(uuid.uuid4())[:4]}.pptx"
    filepath = f"/tmp/{filename}"

    prs.save(filepath)

    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
