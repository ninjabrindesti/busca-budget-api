"""
main.py - Gerador de Proposta Ninja Brindes
Sprint 1: contrato único de placeholders + payload legado/novo
"""

import io
import os
import re as _re
import traceback
import uuid
import zipfile as _zipfile
from contextlib import asynccontextmanager
from copy import deepcopy as _deepcopy
from typing import List, Optional
from urllib.parse import urlparse

import httpx
from lxml import etree as _etree

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, ConfigDict, Field
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

COVER_SLIDE_INDEX = 0
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
                "Template não encontrado localmente e TEMPLATE_URL ausente/inválida."
            )
        try:
            with httpx.Client(timeout=30, follow_redirects=True) as client:
                response = client.get(TEMPLATE_URL.strip())
                response.raise_for_status()
        except Exception as exc:
            raise RuntimeError(f"Falha ao baixar template: {exc}") from exc

        with open(TEMPLATE_PATH, "wb") as f:
            f.write(response.content)

    yield


app = FastAPI(title="Ninja Brindes - Gerador de Proposta", lifespan=lifespan)


# ---------------------------------------------------------------------------
# Schemas — Sprint 1: tudo opcional + extra="allow" p/ payloads legado e novo
# ---------------------------------------------------------------------------

class Proposal(BaseModel):
    model_config = ConfigDict(extra="allow")

    # Únicos campos realmente obrigatórios
    proposal_number: str
    client_name: str

    # Tudo abaixo é opcional para aceitar payloads antigos/parciais
    company_name: Optional[str] = ""
    seller_name: Optional[str] = ""
    seller_phone: Optional[str] = ""
    seller_email: Optional[str] = ""
    seller_description: Optional[str] = ""
    seller_image_url: Optional[str] = None
    payment_method: Optional[str] = ""
    payment_term: Optional[str] = ""
    delivery_date: Optional[str] = ""
    notes: Optional[str] = ""
    obs_cnpj: Optional[str] = ""


class Item(BaseModel):
    model_config = ConfigDict(extra="allow")

    item_index: int
    item_name: str = Field(max_length=120)
    item_subtitle: Optional[str] = Field(default="", max_length=120)
    item_description: Optional[str] = Field(default="", max_length=500)
    item_code: Optional[str] = Field(default="", max_length=60)
    quantity: int = 1
    unit_price: float = 0.0
    item_image_url: Optional[str] = ""


class Section(BaseModel):
    model_config = ConfigDict(extra="allow")

    section_index: int
    freight_value: Optional[float] = None
    freight_label: Optional[str] = ""
    items: List[Item] = Field(min_length=1, max_length=10)


class GenerateRequest(BaseModel):
    model_config = ConfigDict(extra="allow")

    proposal: Proposal
    sections: List[Section] = Field(min_length=1)


# ---------------------------------------------------------------------------
# Helpers de formatação
# ---------------------------------------------------------------------------

def _format_currency(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _calculate_section_total(section: Section) -> float:
    return sum(item.quantity * item.unit_price for item in section.items)


def _calculate_grand_total(section: Section) -> float:
    return _calculate_section_total(section) + (section.freight_value or 0.0)


def _s(value) -> str:
    """Coerção segura: None -> ''. Preserva strings/numbers como str."""
    if value is None:
        return ""
    return str(value)


# ---------------------------------------------------------------------------
# CONTRATO ÚNICO DE PLACEHOLDERS — Sprint 1
# ---------------------------------------------------------------------------

def _build_global_data(proposal: Proposal) -> dict:
    """
    Define a saída ÚNICA usada pelo template atual.
    Aceita payload novo (com cover_*/summary_*) e legado (sem eles).
    """
    company_name   = _s(getattr(proposal, "company_name", ""))
    payment_method = _s(getattr(proposal, "payment_method", ""))
    payment_term   = _s(getattr(proposal, "payment_term", ""))
    delivery_date  = _s(getattr(proposal, "delivery_date", ""))
    notes          = _s(getattr(proposal, "notes", ""))
    obs_cnpj       = _s(getattr(proposal, "obs_cnpj", ""))

    in_summary_pm = _s(getattr(proposal, "summary_payment_method", ""))
    in_summary_dd = _s(getattr(proposal, "summary_delivery_date", ""))
    in_cover_pn   = _s(getattr(proposal, "cover_proposal_number", ""))
    in_cover_co   = _s(getattr(proposal, "cover_company", ""))
    in_cover_cl   = _s(getattr(proposal, "cover_client", ""))
    in_cover_obs  = _s(getattr(proposal, "cover_obs_cnpj", ""))

    # === REGRA ÚNICA DE FALLBACK (Sprint 1) ===
    summary_payment_method = in_summary_pm or payment_method or payment_term or ""
    summary_delivery_date  = in_summary_dd or delivery_date  or ""
    cover_obs_cnpj         = in_cover_obs  or obs_cnpj       or notes or ""
    cover_company          = in_cover_co   or company_name   or ""
    cover_client           = in_cover_cl   or _s(proposal.client_name)
    cover_proposal_number  = in_cover_pn   or _s(proposal.proposal_number)

    return {
        # Legados (mantidos só para retrocompat de slides antigos)
        "proposal_number":   _s(proposal.proposal_number),
        "client_name":       _s(proposal.client_name),
        "company_name":      company_name,
        "seller_name":       _s(getattr(proposal, "seller_name", "")),
        "seller_phone":      _s(getattr(proposal, "seller_phone", "")),
        "seller_email":      _s(getattr(proposal, "seller_email", "")),
        "seller_description":_s(getattr(proposal, "seller_description", "")),
        "seller_image_url":  _s(getattr(proposal, "seller_image_url", "")),
        "payment_method":    payment_method,
        "delivery_date":     delivery_date,
        "notes":             notes,

        # OFICIAIS do template atual (única fonte de verdade)
        "summary_payment_method": summary_payment_method,
        "summary_delivery_date":  summary_delivery_date,
        "cover_proposal_number":  cover_proposal_number,
        "cover_company":          cover_company,
        "cover_client":           cover_client,
        "cover_obs_cnpj":         cover_obs_cnpj,
    }


def _build_data(proposal: Proposal, item: Item, section: Section) -> dict:
    item_total = item.quantity * item.unit_price
    freight_value = section.freight_value or 0.0
    section_total = _calculate_section_total(section)
    grand_total = _calculate_grand_total(section)

    return {
        **_build_global_data(proposal),
        "item_name":          _s(item.item_name),
        "item_subtitle":      _s(item.item_subtitle),
        "item_index":         str(item.item_index),
        "item_display_index": str(item.item_index + 1),
        "item_description":   _s(item.item_description),
        "item_code":          _s(item.item_code),
        "quantity":           str(item.quantity),
        "unit_price":         _format_currency(item.unit_price),
        "item_total":         _format_currency(item_total),
        "item_image_url":     _s(item.item_image_url),
        "section_total":      _format_currency(section_total),
        "freight":            _format_currency(freight_value),
        "grand_total":        _format_currency(grand_total),
        "freight_label":      _s(section.freight_label),
    }


def _build_summary_data(proposal: Proposal, section: Section) -> dict:
    section_total = _calculate_section_total(section)
    freight_value = section.freight_value or 0.0
    grand_total = section_total + freight_value

    return {
        **_build_global_data(proposal),
        "section_total": _format_currency(section_total),
        "freight":       _format_currency(freight_value),
        "grand_total":   _format_currency(grand_total),
        "freight_label": _s(section.freight_label),
    }


# ---------------------------------------------------------------------------
# Fallback: substitui {{key}} preservando parágrafos (1 linha por placeholder)
# ---------------------------------------------------------------------------

def _force_replace_on_slide(slide, replacements: dict) -> None:
    """
    Substitui {{key}} em todos os text_frames do slide, parágrafo a parágrafo,
    preservando quebras de linha e a formatação do primeiro run.
    """
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            continue
        for paragraph in shape.text_frame.paragraphs:
            original_text = "".join(run.text for run in paragraph.runs)
            if not original_text:
                continue
            new_text = original_text
            for key, value in replacements.items():
                new_text = _re.sub(
                    r"{{\s*" + _re.escape(key) + r"\s*}}",
                    value or "",
                    new_text,
                )
            if new_text != original_text and paragraph.runs:
                paragraph.runs[0].text = new_text
                for run in paragraph.runs[1:]:
                    run.text = ""


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate")
def generate_proposal(payload: GenerateRequest):
    try:
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

        # === Capa: 4 placeholders cover_* (1 linha cada, preserva parágrafos)
        cover_slide = prs.slides[COVER_SLIDE_INDEX]
        replace_text_placeholders_on_slide(cover_slide, global_data)
        replace_named_images_on_slide(cover_slide, global_data)
        _force_replace_on_slide(cover_slide, {
            "cover_proposal_number": global_data["cover_proposal_number"],
            "cover_company":         global_data["cover_company"],
            "cover_client":          global_data["cover_client"],
            "cover_obs_cnpj":        global_data["cover_obs_cnpj"],
        })

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
                        detail=f"Slide de item ausente. Índice: {slide_cursor}",
                    )
                slide = prs.slides[slide_cursor]
                data = _build_data(payload.proposal, item, section)
                replace_text_placeholders_on_slide(slide, data)
                replace_named_images_on_slide(slide, data)
                slide_cursor += 1

            if slide_cursor >= len(prs.slides):
                raise HTTPException(
                    status_code=500,
                    detail=f"Slide de resumo ausente. Índice: {slide_cursor}",
                )

            summary_slide = prs.slides[slide_cursor]
            summary_data = _build_summary_data(payload.proposal, section)

            _expand_summary_table(summary_slide, section)
            replace_text_placeholders_on_slide(summary_slide, summary_data)

            # Fallback p/ summary_* + legados (preserva 1 placeholder por linha)
            _force_replace_on_slide(summary_slide, {
                "summary_payment_method": summary_data["summary_payment_method"],
                "summary_delivery_date":  summary_data["summary_delivery_date"],
                "payment_method":         summary_data["payment_method"],
                "delivery_date":          summary_data["delivery_date"],
            })

            replace_named_images_on_slide(summary_slide, summary_data)
            slide_cursor += 1

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

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "error": "render_failed",
                "type": type(e).__name__,
                "message": str(e),
            },
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
            prs_root, xml_declaration=True, encoding="utf-8", standalone=True,
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
    for shape in slide.shapes:
        if not getattr(shape, "has_table", False):
            continue

        table = shape.table
        rows = list(table.rows)

        template_row_idx = None
        for i, row in enumerate(rows):
            row_text = " ".join(cell.text for cell in row.cells)
            if any(token in row_text for token in (
                "{{item_", "{{ item_", "{{quantity", "{{ quantity",
                "{{unit_price", "{{ unit_price", "{{item_total", "{{ item_total",
            )):
                template_row_idx = i
                break

        if template_row_idx is None:
            continue

        template_row_el = rows[template_row_idx]._tr
        parent = template_row_el.getparent()

        current_rows = list(table.rows)
        for row in current_rows[template_row_idx + 1:]:
            row_text = " ".join(cell.text for cell in row.cells)
            if any(token in row_text for token in (
                "{{item_", "{{ item_", "{{quantity", "{{ quantity",
                "{{unit_price", "{{ unit_price", "{{item_total", "{{ item_total",
            )):
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
                "item_index":         str(item.item_index),
                "item_name":          _s(item.item_name),
                "item_subtitle":      _s(item.item_subtitle),
                "item_description":   _s(item.item_description),
                "item_code":          _s(item.item_code),
                "quantity":           str(item.quantity),
                "unit_price":         _format_currency(item.unit_price),
                "item_total":         _format_currency(item_total),
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
                            value, new_text,
                        )
                    if new_text != original_text and paragraph.runs:
                        paragraph.runs[0].text = new_text
                        for run in paragraph.runs[1:]:
                            run.text = ""

        break
