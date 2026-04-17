"""
main.py - Gerador de Proposta Ninja Brindes
"""

import io
import os
import uuid
from contextlib import asynccontextmanager
from copy import deepcopy as _deepcopy
from typing import List, Optional
from urllib.parse import urlparse
import re as _re

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
# Constantes
# ---------------------------------------------------------------------------

TEMPLATE_PATH = "templates/template_ninja.pptx"
TEMPLATE_URL = os.getenv("TEMPLATE_URL")

ITEM_SLIDE_INDEX = 8
SUMMARY_SLIDE_INDEX = 9


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

    return parsed.scheme in ("http", "https") and bool(parsed.netloc)


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


def _calculate_section_total(section: Section) -> float:
    return sum(item.quantity * item.unit_price for item in section.items)


def _calculate_grand_total(section: Section) -> float:
    return _calculate_section_total(section) + (section.freight_value or 0.0)


def _build_global_data(proposal: Proposal) -> dict:
    return {
        "proposal_number": proposal.proposal_number,
        "client_name": proposal.client_name,
        "seller_name": proposal.seller_name,
        "seller_phone": proposal.seller_phone,
        "seller_email": proposal.seller_email,
        "seller_description": proposal.seller_description,
        "seller_image_url": proposal.seller_image_url or "",
        "payment_method": proposal.payment_method,
        "delivery_date": proposal.delivery_date,
        "notes": proposal.notes,
    }


def _build_data(proposal: Proposal, item: Item, section: Section) -> dict:
    item_total = item.quantity * item.unit_price
    freight_value = section.freight_value or 0.0
    section_total = _calculate_section_total(section)
    grand_total = _calculate_grand_total(section)

    return {
        **_build_global_data(proposal),
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
        "section_total": _format_currency(section_total),
        "freight": _format_currency(freight_value),
        "grand_total": _format_currency(grand_total),
        "freight_label": section.freight_label,
    }


def _build_summary_data(proposal: Proposal, section: Section) -> dict:
    section_total = _calculate_section_total(section)
    freight_value = section.freight_value or 0.0
    grand_total = section_total + freight_value

    return {
        **_build_global_data(proposal),
        "section_total": _format_currency(section_total),
        "freight": _format_currency(freight_value),
        "grand_total": _format_currency(grand_total),
        "freight_label": section.freight_label,
        "payment_method": proposal.payment_method,
        "delivery_date": proposal.delivery_date,
    }


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate")
def generate_proposal(payload: GenerateRequest):
    print("DEBUG /generate payload:", payload.model_dump())

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail="Template não encontrado.")

    with open(TEMPLATE_PATH, "rb") as f:
        pptx_bytes = f.read()

    total_item_slides = sum(len(s.items) for s in payload.sections)
    total_summary_slides = len(payload.sections)

    if total_item_slides <= 0:
        raise HTTPException(status_code=400, detail="Nenhum item enviado.")

    item_copies = max(total_item_slides - 1, 0)
    summary_copies = max(total_summary_slides - 1, 0)

    for _ in range(item_copies):
        pptx_bytes = duplicate_slide_in_pptx(
            input_bytes=pptx_bytes,
            source_slide_index=ITEM_SLIDE_INDEX,
            copies=1,
            insert_after_index=ITEM_SLIDE_INDEX,
        )

    summary_original_index = ITEM_SLIDE_INDEX + total_item_slides
    for _ in range(summary_copies):
        pptx_bytes = duplicate_slide_in_pptx(
            input_bytes=pptx_bytes,
            source_slide_index=summary_original_index,
            copies=1,
            insert_after_index=summary_original_index,
        )

    pptx_bytes = _reorder_slides(
        pptx_bytes=pptx_bytes,
        sections=payload.sections,
        total_item_slides=total_item_slides,
        total_summary_slides=total_summary_slides,
    )

    prs = Presentation(io.BytesIO(pptx_bytes))

    global_data = _build_global_data(payload.proposal)

    # Slide fixo do vendedor
    seller_slide = prs.slides[3]
    replace_text_placeholders_on_slide(seller_slide, global_data)
    replace_named_images_on_slide(seller_slide, global_data)

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

        # força substituição dos placeholders simples fora da tabela
        for shape in summary_slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    full_text = "".join(run.text for run in paragraph.runs)

                    if not full_text:
                        continue

                    full_text = full_text.replace(
                        "{{payment_method}}",
                        summary_data["payment_method"],
                    )
                    full_text = full_text.replace(
                        "{{delivery_date}}",
                        summary_data["delivery_date"],
                    )

                    paragraph.text = full_text

        replace_named_images_on_slide(summary_slide, summary_data)
        slide_cursor += 1

    # Último slide fixo
    last_slide = prs.slides[-1]
    replace_text_placeholders_on_slide(last_slide, global_data)
    replace_named_images_on_slide(last_slide, global_data)

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


def _reorder_slides(
    pptx_bytes: bytes,
    sections: List[Section],
    total_item_slides: int,
    total_summary_slides: int,
) -> bytes:
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
# Expansão da tabela de resumo
# ---------------------------------------------------------------------------

def _expand_summary_table(slide, section):
    """
    Expande a tabela do slide de resumo criando 1 linha por item da section
    e preenche cada linha usando replace direto no text_frame das células.
    """
    for shape in slide.shapes:
        if not getattr(shape, "has_table", False):
            continue

        table = shape.table
        rows = list(table.rows)

        template_row_idx = None
        for i, row in enumerate(rows):
            row_text = " ".join(cell.text for cell in row.cells)
            if (
                "{{item_" in row_text
                or "{{ item_" in row_text
                or "{{quantity" in row_text
                or "{{ quantity" in row_text
                or "{{unit_price" in row_text
                or "{{ unit_price" in row_text
                or "{{item_total" in row_text
                or "{{ item_total" in row_text
            ):
                template_row_idx = i
                break

        if template_row_idx is None:
            continue

        template_row_el = rows[template_row_idx]._tr
        parent = template_row_el.getparent()

        current_rows = list(table.rows)
        for row in current_rows[template_row_idx + 1:]:
            row_text = " ".join(cell.text for cell in row.cells)
            if (
                "{{item_" in row_text
                or "{{ item_" in row_text
                or "{{quantity" in row_text
                or "{{ quantity" in row_text
                or "{{unit_price" in row_text
                or "{{ unit_price" in row_text
                or "{{item_total" in row_text
                or "{{ item_total" in row_text
            ):
                parent.remove(row._tr)

        for _ in section.items[1:]:
            new_row_el = _deepcopy(template_row_el)
            parent.insert(parent.index(template_row_el) + 1, new_row_el)
            template_row_el = new_row_el

        rows = list(table.rows)
        item_rows = rows[template_row_idx: template_row_idx + len(section.items)]

        for row, item in zip(item_rows, section.items):
            item_total = item.quantity * item.unit_price
            row_data = {
                "item_display_index": str(item.item_index + 1),
                "item_index": str(item.item_index),
                "item_name": item.item_name,
                "item_subtitle": item.item_subtitle,
                "item_description": item.item_description,
                "item_code": item.item_code,
                "quantity": str(item.quantity),
                "unit_price": _format_currency(item.unit_price),
                "item_total": _format_currency(item_total),
            }

            for cell in row.cells:
                if cell.text_frame is None:
                    continue

                for paragraph in cell.text_frame.paragraphs:
                    original_text = "".join(run.text for run in paragraph.runs)
                    if not original_text:
                        continue

                    new_text = original_text
                    for key, value in row_data.items():
                        new_text = _re.sub(
                            r"{{\s*" + _re.escape(key) + r"\s*}}",
                            value,
                            new_text,
                        )

                    if new_text != original_text and len(paragraph.runs) > 0:
                        paragraph.runs[0].text = new_text
                        for run in paragraph.runs[1:]:
                            run.text = ""

        break
