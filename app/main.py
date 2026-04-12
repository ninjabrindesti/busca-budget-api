"""
main.py - Gerador de Proposta Ninja Brindes
Stack: FastAPI + Railway (Linux)

Lógica de slides:
  Para cada orçamento (Section):
    - Slide 9 clonado uma vez por item:  9(A), 9(B), 9(C)
    - Slide 10 uma vez com todos itens:  10(A+B+C)
  Múltiplos orçamentos se encadeiam:
    9(A) 9(B) 9(C) 10(ABC)  →  9(D) 9(E) 10(DE)
"""

import io
import os
import uuid
from contextlib import asynccontextmanager
from copy import deepcopy as _deepcopy
from typing import List, Optional
from urllib.parse import urlparse

import httpx
import zipfile as _zipfile
from lxml import etree as _etree

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from pptx import Presentation

from app.services.pptx_generator import (
    replace_text_placeholders_on_slide,
    replace_named_images_on_slide,
)
from app.services.slide_duplicator import duplicate_slide_in_pptx


# ---------------------------------------------------------------------------
# Constantes do template
# ---------------------------------------------------------------------------

TEMPLATE_PATH = "templates/template_ninja.pptx"
TEMPLATE_URL = os.getenv("TEMPLATE_URL")

# Índices 0-based no template original
ITEM_SLIDE_INDEX = 8       # slide 9  → template de item individual
SUMMARY_SLIDE_INDEX = 9    # slide 10 → template de resumo/orçamento


# ---------------------------------------------------------------------------
# Helpers de URL / template
# ---------------------------------------------------------------------------

def _is_valid_http_url(url: Optional[str]) -> bool:
    if not url or not isinstance(url, str):
        return False

    url = url.strip()
    if " " in url:
        return False

    try:
        parsed = urlparse(url)
    except Exception:
        return False

    if parsed.scheme not in ("http", "https"):
        return False

    if not parsed.netloc:
        return False

    return True


# ---------------------------------------------------------------------------
# Startup: baixa o template do Supabase Storage se não existir localmente
# ---------------------------------------------------------------------------

@asynccontextmanager
async def lifespan(app: FastAPI):
    os.makedirs("templates", exist_ok=True)

    if not os.path.exists(TEMPLATE_PATH):
        if not _is_valid_http_url(TEMPLATE_URL):
            raise RuntimeError(
                "Template não encontrado localmente e TEMPLATE_URL está ausente ou inválida."
            )

        try:
            with httpx.Client(timeout=30, follow_redirects=True) as client:
                response = client.get(TEMPLATE_URL.strip())
                response.raise_for_status()
        except Exception as exc:
            raise RuntimeError(
                f"Falha ao baixar o template em TEMPLATE_URL: {exc}"
            ) from exc

        with open(TEMPLATE_PATH, "wb") as f:
            f.write(response.content)

    yield


app = FastAPI(title="Ninja Brindes - Gerador de Proposta", lifespan=lifespan)


# ---------------------------------------------------------------------------
# Schemas
# ---------------------------------------------------------------------------

class Proposal(BaseModel):
    proposal_number: str
    client_name: str
    seller_name: str
    seller_phone: str
    seller_email: str
    seller_description: str
    seller_image_url: Optional[str] = None
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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _format_currency(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _build_data(proposal: Proposal, item: Item, section: Section) -> dict:
    item_total = item.quantity * item.unit_price
    freight_value = section.freight_value or 0.0
    section_total = sum(i.quantity * i.unit_price for i in section.items)

    return {
        # Proposta
        "proposal_number": proposal.proposal_number,
        "client_name": proposal.client_name,
        "payment_method": proposal.payment_method,
        "delivery_date": proposal.delivery_date,
        "notes": proposal.notes,
        # Vendedor
        "seller_name": proposal.seller_name,
        "seller_phone": proposal.seller_phone,
        "seller_email": proposal.seller_email,
        "seller_description": proposal.seller_description,
        "seller_image_url": proposal.seller_image_url or "",
        # Item
        "item_name": item.item_name,
        "item_subtitle": item.item_subtitle,
        "item_index": str(item.item_index),
        "item_display_index": str(item.item_index + 1),
        "item_description": item.item_description,
        "item_code": item.item_code,
        "quantity": str(item.quantity),
        "unit_price": _format_currency(item.unit_price),
        "item_total": _format_currency(item_total),
        "item_image_url": item.item_image_url,
        # Seção
        "section_total": _format_currency(section_total),
        "freight": _format_currency(freight_value),
        "freight_label": section.freight_label,
    }


def _build_summary_data(proposal: Proposal, section: Section) -> dict:
    first_item = section.items[0]
    return _build_data(proposal, first_item, section)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate")
def generate_proposal(payload: GenerateRequest):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail="Template não encontrado.")

    with open(TEMPLATE_PATH, "rb") as f:
        pptx_bytes = f.read()

    # -----------------------------------------------------------------------
    # PASSO 1: Duplicar slides de item e resumo conforme quantidade de seções
    # -----------------------------------------------------------------------

    total_item_slides = sum(len(s.items) for s in payload.sections)
    total_summary_slides = len(payload.sections)

    item_copies = total_item_slides - 1
    summary_copies = total_summary_slides - 1

    # Duplica slides de item (clona o slide 9 original)
    for _ in range(item_copies):
        pptx_bytes = duplicate_slide_in_pptx(
            input_bytes=pptx_bytes,
            source_slide_index=ITEM_SLIDE_INDEX,
            copies=1,
            insert_after_index=ITEM_SLIDE_INDEX,
        )

    # Duplica slides de resumo (clona o slide 10 original, agora deslocado)
    summary_original_index = ITEM_SLIDE_INDEX + total_item_slides
    for _ in range(summary_copies):
        pptx_bytes = duplicate_slide_in_pptx(
            input_bytes=pptx_bytes,
            source_slide_index=summary_original_index,
            copies=1,
            insert_after_index=summary_original_index,
        )

    # -----------------------------------------------------------------------
    # PASSO 2: Reordenar slides para intercalar item/resumo por seção
    # -----------------------------------------------------------------------

    pptx_bytes = _reorder_slides(
        pptx_bytes=pptx_bytes,
        sections=payload.sections,
        total_item_slides=total_item_slides,
        total_summary_slides=total_summary_slides,
    )

    # -----------------------------------------------------------------------
    # PASSO 3: Preencher placeholders com python-pptx
    # -----------------------------------------------------------------------

    prs = Presentation(io.BytesIO(pptx_bytes))

    # Slide do vendedor (slide 4 = índice 3)
    seller_slide = prs.slides[3]
    seller_data = {
        "seller_name": payload.proposal.seller_name,
        "seller_phone": payload.proposal.seller_phone,
        "seller_email": payload.proposal.seller_email,
        "seller_description": payload.proposal.seller_description,
        "seller_image_url": payload.proposal.seller_image_url or "",
    }
    replace_text_placeholders_on_slide(seller_slide, seller_data)
    replace_named_images_on_slide(seller_slide, seller_data)

    # Slides de item e resumo
    slide_cursor = ITEM_SLIDE_INDEX

    for section in payload.sections:
        for item in section.items:
            if slide_cursor >= len(prs.slides):
                raise HTTPException(
                    status_code=500,
                    detail=f"Slide de item não encontrado. Índice: {slide_cursor}",
                )

            slide = prs.slides[slide_cursor]
            data = _build_data(payload.proposal, item, section)
            replace_text_placeholders_on_slide(slide, data)
            replace_named_images_on_slide(slide, data)
            slide_cursor += 1

        if slide_cursor >= len(prs.slides):
            raise HTTPException(
                status_code=500,
                detail=f"Slide de resumo não encontrado. Índice: {slide_cursor}",
            )

        summary_slide = prs.slides[slide_cursor]
        summary_data = _build_summary_data(payload.proposal, section)
        _expand_summary_table(summary_slide, section)
        replace_text_placeholders_on_slide(summary_slide, summary_data)
        replace_named_images_on_slide(summary_slide, summary_data)
        slide_cursor += 1

    out_buf = io.BytesIO()
    prs.save(out_buf)
    out_buf.seek(0)

    filename = f"proposta_{uuid.uuid4().hex[:8]}.pptx"
    return StreamingResponse(
        out_buf,
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation"
        ),
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ---------------------------------------------------------------------------
# Reordenação de slides
# ---------------------------------------------------------------------------

_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _reorder_slides(
    pptx_bytes: bytes,
    sections: List[Section],
    total_item_slides: int,
    total_summary_slides: int,
) -> bytes:
    """
    Reordena os slides de item e resumo para intercalar por seção.

    Antes: [0-7 fixo] [item×N] [resumo×M] [trailing...]
    Depois:[0-7 fixo] [sec1_items... sec1_resumo] [sec2_items... sec2_resumo] [trailing...]
    """
    buf: dict[str, bytes] = {}

    with _zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zf:
        prs_root = _etree.fromstring(zf.read("ppt/presentation.xml"))

        sld_id_lst = prs_root.find(f".//{{{_NS_P}}}sldIdLst")
        sld_id_els = list(sld_id_lst)

        fixed_before = sld_id_els[:ITEM_SLIDE_INDEX]
        item_slides = sld_id_els[ITEM_SLIDE_INDEX: ITEM_SLIDE_INDEX + total_item_slides]
        summary_slides = sld_id_els[
            ITEM_SLIDE_INDEX + total_item_slides:
            ITEM_SLIDE_INDEX + total_item_slides + total_summary_slides
        ]
        fixed_after = sld_id_els[
            ITEM_SLIDE_INDEX + total_item_slides + total_summary_slides:
        ]

        new_order = list(fixed_before)
        item_ptr = 0
        summary_ptr = 0

        for section in sections:
            for _ in range(len(section.items)):
                new_order.append(item_slides[item_ptr])
                item_ptr += 1

            new_order.append(summary_slides[summary_ptr])
            summary_ptr += 1

        new_order.extend(fixed_after)

        for el in list(sld_id_lst):
            sld_id_lst.remove(el)
        for el in new_order:
            sld_id_lst.append(el)

        buf["ppt/presentation.xml"] = _etree.tostring(
            prs_root,
            xml_declaration=True,
            encoding="utf-8",
            standalone=True,
        )

        out = io.BytesIO()
        with _zipfile.ZipFile(out, "w", _zipfile.ZIP_DEFLATED) as out_zf:
            written = set()
            for name, data in buf.items():
                out_zf.writestr(name, data)
                written.add(name)

            for name in zf.namelist():
                if name not in written:
                    out_zf.writestr(name, zf.read(name))

        return out.getvalue()


# ---------------------------------------------------------------------------
# Expansão da tabela de resumo para múltiplos itens
# ---------------------------------------------------------------------------

def _expand_summary_table(slide, section: Section):
    """
    Duplica as linhas de dados da tabela no slide de resumo para cada item.
    O template original tem 1 linha de item — adicionamos as demais.
    """
    first_item = section.items[0]

    for shape in slide.shapes:
        if not getattr(shape, "has_table", False):
            continue

        table = shape.table
        rows = list(table.rows)

        item_row_idx = None
        for i, row in enumerate(rows):
            row_text = " ".join(cell.text for cell in row.cells)
            if (
                first_item.item_code in row_text
                or str(first_item.item_index + 1) in row_text
                or "{{item_" in row_text
            ):
                item_row_idx = i
                break

        if item_row_idx is None:
            continue

        src_row_el = rows[item_row_idx]._tr

        for extra_item in section.items[1:]:
            new_row_el = _deepcopy(src_row_el)

            for tc in new_row_el.iter(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
            ):
                if tc.text:
                    tc.text = (
                        tc.text
                        .replace(first_item.item_code, extra_item.item_code)
                        .replace(str(first_item.item_index + 1), str(extra_item.item_index + 1))
                        .replace(first_item.item_name, extra_item.item_name)
                        .replace(first_item.item_description, extra_item.item_description)
                    )

            src_row_el.addnext(new_row_el)
            src_row_el = new_row_el

        break
