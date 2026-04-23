"""
main.py — Gerador de Proposta Ninja Brindes
Sprint 1: contrato único de placeholders + payload legado/novo
Sprint 2: motor central de replace (text frames, tabelas, multi-run)
Sprint 2 — ajustes de ressalvas:
  1. item_display_index determinístico (enumerate por seção)
  2. _iter_text_frames com detecção explícita de group shapes
  3. _replace_summary_table_rows com seleção robusta de linhas
  4. Removido SUMMARY_SLIDE_INDEX não utilizado
"""

import io
import os
import re as _re
import traceback
import uuid
import zipfile as _zipfile
from contextlib import asynccontextmanager
from copy import deepcopy as _deepcopy
from typing import List, Optional, Iterable
from urllib.parse import urlparse

import httpx
from lxml import etree as _etree

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, ConfigDict, Field
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from app.services.pptx_generator import replace_named_images_on_slide
from app.services.slide_duplicator import duplicate_slide_in_pptx


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

TEMPLATE_PATH = "templates/template_ninja.pptx"
TEMPLATE_URL = os.getenv("TEMPLATE_URL")

COVER_SLIDE_INDEX = 0
ITEM_SLIDE_INDEX = 8
# (SUMMARY_SLIDE_INDEX removido — não era usado em lugar nenhum.)


# ---------------------------------------------------------------------------
# Template bootstrap
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
            raise RuntimeError("Template ausente e TEMPLATE_URL inválida.")
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
# Schemas (Sprint 1)
# ---------------------------------------------------------------------------

class Proposal(BaseModel):
    model_config = ConfigDict(extra="allow")
    proposal_number: str
    client_name: str
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
    cover_cnpj: Optional[str] = ""
    cover_corporate_name: Optional[str] = ""


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
# Helpers
# ---------------------------------------------------------------------------

def _format_currency(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _calculate_section_total(section: Section) -> float:
    return sum(item.quantity * item.unit_price for item in section.items)


def _calculate_grand_total(section: Section) -> float:
    return _calculate_section_total(section) + (section.freight_value or 0.0)


def _s(value) -> str:
    return "" if value is None else str(value)


# ---------------------------------------------------------------------------
# CONTRATO ÚNICO DE PLACEHOLDERS (Sprint 1)
# ---------------------------------------------------------------------------

def _build_global_data(proposal: Proposal) -> dict:
    company_name   = _s(getattr(proposal, "company_name", ""))
    payment_method = _s(getattr(proposal, "payment_method", ""))
    payment_term   = _s(getattr(proposal, "payment_term", ""))
    delivery_date  = _s(getattr(proposal, "delivery_date", ""))
    notes          = _s(getattr(proposal, "notes", ""))
    obs_cnpj       = _s(getattr(proposal, "obs_cnpj", ""))

    in_summary_pm      = _s(getattr(proposal, "summary_payment_method", ""))
    in_summary_dd      = _s(getattr(proposal, "summary_delivery_date", ""))
    in_cover_pn        = _s(getattr(proposal, "cover_proposal_number", ""))
    in_cover_co        = _s(getattr(proposal, "cover_company", ""))
    in_cover_cl        = _s(getattr(proposal, "cover_client", ""))
    in_cover_obs       = _s(getattr(proposal, "cover_obs_cnpj", ""))
    in_cover_cnpj      = _s(getattr(proposal, "cover_cnpj", ""))
    in_cover_corp_name = _s(getattr(proposal, "cover_corporate_name", ""))

    summary_payment_method = in_summary_pm or payment_method or payment_term or ""
    summary_delivery_date  = in_summary_dd or delivery_date  or ""
    cover_obs_cnpj         = in_cover_obs  or obs_cnpj       or notes or ""
    cover_company          = in_cover_co   or company_name   or ""
    cover_client           = in_cover_cl   or _s(proposal.client_name)
    cover_proposal_number  = in_cover_pn   or _s(proposal.proposal_number)
    cover_cnpj             = in_cover_cnpj
    cover_corporate_name   = in_cover_corp_name

    return {
        # legados (retrocompat)
        "proposal_number":   _s(proposal.proposal_number),
        "client_name":       _s(proposal.client_name),
        "company_name":      company_name,
        "seller_name":       _s(getattr(proposal, "seller_name", "")),
        "seller_phone":      _s(getattr(proposal, "seller_phone", "")),
        "seller_email":      _s(getattr(proposal, "seller_email", "")),
        "seller_description":_s(getattr(proposal, "seller_description", "")),
        "seller_image_url":  _s(getattr(proposal, "seller_image_url", "")),
        "payment_method":    payment_method,
        "payment_term":      payment_term,
        "delivery_date":     delivery_date,
        "notes":             notes,
        # OFICIAIS template atual
        "summary_payment_method": summary_payment_method,
        "summary_delivery_date":  summary_delivery_date,
        "cover_proposal_number":  cover_proposal_number,
        "cover_company":          cover_company,
        "cover_client":           cover_client,
        "cover_obs_cnpj":         cover_obs_cnpj,
        # NOVOS
        "cover_cnpj":             cover_cnpj,
        "cover_corporate_name":   cover_corporate_name,
    }


def _build_data(
    proposal: Proposal,
    item: Item,
    section: Section,
    display_index: int,
) -> dict:
    """
    `display_index` é o índice visual 1-based, calculado pelo chamador
    via enumerate (NÃO depende de item.item_index do payload).
    """
    item_total = item.quantity * item.unit_price
    return {
        **_build_global_data(proposal),
        "item_name":          _s(item.item_name),
        "item_subtitle":      _s(item.item_subtitle),
        "item_index":         str(item.item_index),
        "item_display_index": str(display_index),
        "item_description":   _s(item.item_description),
        "item_code":          _s(item.item_code),
        "quantity":           str(item.quantity),
        "unit_price":         _format_currency(item.unit_price),
        "item_total":         _format_currency(item_total),
        "item_image_url":     _s(item.item_image_url),
        "section_total":      _format_currency(_calculate_section_total(section)),
        "freight":            _format_currency(section.freight_value or 0.0),
        "section_freight":    _s(section.freight_label),
        "grand_total":        _format_currency(_calculate_grand_total(section)),
        "freight_label":      _s(section.freight_label),
    }


def _build_summary_data(proposal: Proposal, section: Section) -> dict:
    section_total = _calculate_section_total(section)
    freight_value = section.freight_value or 0.0
    return {
        **_build_global_data(proposal),
        "section_total":   _format_currency(section_total),
        "freight":         _format_currency(freight_value),
        "section_freight": _s(section.freight_label),
        "grand_total":     _format_currency(section_total + freight_value),
        "freight_label":   _s(section.freight_label),
    }


# ===========================================================================
# SPRINT 2 — MOTOR CENTRAL DE REPLACE
# ===========================================================================

_PLACEHOLDER_CACHE: dict[str, _re.Pattern] = {}


def _pattern_for(key: str) -> _re.Pattern:
    pat = _PLACEHOLDER_CACHE.get(key)
    if pat is None:
        pat = _re.compile(r"\{\{\s*" + _re.escape(key) + r"\s*\}\}")
        _PLACEHOLDER_CACHE[key] = pat
    return pat


def _apply_to_paragraph(paragraph, replacements: dict) -> bool:
    """
    Substitui todos os placeholders no parágrafo, preservando o rPr do 1º run.
    Retorna True se algo mudou. Não remove runs nem altera ordem.
    """
    runs = paragraph.runs
    if not runs:
        return False

    original = "".join(r.text or "" for r in runs)
    if "{{" not in original:
        return False

    new_text = original
    for key, value in replacements.items():
        pat = _pattern_for(key)
        if pat.search(new_text):
            new_text = pat.sub(_s(value), new_text)

    if new_text == original:
        return False

    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""
    return True


def _is_group_shape(shape) -> bool:
    """Detecção explícita de group shape, sem heurística ambígua."""
    try:
        return shape.shape_type == MSO_SHAPE_TYPE.GROUP
    except (AttributeError, ValueError):
        return False


def _iter_text_frames(shapes) -> Iterable:
    """
    Itera recursivamente todos os text_frames acessíveis a partir de uma
    coleção de shapes — incluindo grupos (recursão) e tabelas (cada célula).
    Ordem preservada; não modifica nada a árvore de shapes.
    """
    for shape in shapes:
        # 1) Group shape → recursão explícita
        if _is_group_shape(shape):
            yield from _iter_text_frames(shape.shapes)
            continue

        # 2) Tabela → text_frame de cada célula
        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text_frame is not None:
                        yield cell.text_frame
            continue

        # 3) Shape de texto comum
        if getattr(shape, "has_text_frame", False) and shape.text_frame is not None:
            yield shape.text_frame


def replace_placeholders_everywhere(slide, replacements: dict) -> int:
    """
    Motor central. Aplica `replacements` em todos os text_frames do slide
    (text boxes, células, grupos). Cada parágrafo é tratado isoladamente
    (não colapsa linhas; não altera z-order).
    """
    changed = 0
    for tf in _iter_text_frames(slide.shapes):
        for paragraph in tf.paragraphs:
            if _apply_to_paragraph(paragraph, replacements):
                changed += 1
    return changed


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

        for _ in range(max(total_item_slides - 1, 0)):
            pptx_bytes = duplicate_slide_in_pptx(
                input_bytes=pptx_bytes,
                source_slide_index=ITEM_SLIDE_INDEX,
                copies=1,
                insert_after_index=ITEM_SLIDE_INDEX,
            )

        summary_original_index = ITEM_SLIDE_INDEX + total_item_slides
        for _ in range(max(total_summary_slides - 1, 0)):
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

        # Capa
        cover_slide = prs.slides[COVER_SLIDE_INDEX]
        replace_placeholders_everywhere(cover_slide, global_data)
        replace_named_images_on_slide(cover_slide, global_data)

        # Vendedor
        seller_slide = prs.slides[3]
        replace_placeholders_everywhere(seller_slide, global_data)
        replace_named_images_on_slide(seller_slide, global_data)

        # Itens + resumo por seção
        slide_cursor = ITEM_SLIDE_INDEX
        for section in payload.sections:
            # Índice visual 1-based determinístico, independente do payload
            for display_index, item in enumerate(section.items, start=1):
                if slide_cursor >= len(prs.slides):
                    raise HTTPException(status_code=500, detail=f"Slide de item ausente. Índice: {slide_cursor}")
                slide = prs.slides[slide_cursor]
                data = _build_data(payload.proposal, item, section, display_index)
                replace_placeholders_everywhere(slide, data)
                replace_named_images_on_slide(slide, data)
                slide_cursor += 1

            if slide_cursor >= len(prs.slides):
                raise HTTPException(status_code=500, detail=f"Slide de resumo ausente. Índice: {slide_cursor}")

            summary_slide = prs.slides[slide_cursor]
            summary_data = _build_summary_data(payload.proposal, section)

            _expand_summary_table_rows(summary_slide, section)
            replace_placeholders_everywhere(summary_slide, summary_data)
            _replace_summary_table_rows(summary_slide, section)
            replace_named_images_on_slide(summary_slide, summary_data)
            slide_cursor += 1

        # Slide final
        last_slide = prs.slides[-1]
        replace_placeholders_everywhere(last_slide, global_data)
        replace_named_images_on_slide(last_slide, global_data)

        out_buf = io.BytesIO()
        prs.save(out_buf)
        out_buf.seek(0)

        filename = f"proposta_{uuid.uuid4().hex[:8]}.pptx"
        return StreamingResponse(
            out_buf,
            media_type=("application/vnd.openxmlformats-officedocument"
                        ".presentationml.presentation"),
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail={"error": "render_failed", "type": type(e).__name__, "message": str(e)},
        )


# ---------------------------------------------------------------------------
# Reordenação dos slides duplicados
# ---------------------------------------------------------------------------

_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _reorder_slides(pptx_bytes, sections, total_item_slides, total_summary_slides):
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
        fixed_after = sld_id_els[ITEM_SLIDE_INDEX + total_item_slides + total_summary_slides:]

        new_order = list(fixed_before)
        item_ptr = summary_ptr = 0
        for section in sections:
            for _ in range(len(section.items)):
                new_order.append(item_slides[item_ptr]); item_ptr += 1
            new_order.append(summary_slides[summary_ptr]); summary_ptr += 1
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
                out_zf.writestr(name, data); written.add(name)
            for name in zf.namelist():
                if name not in written:
                    out_zf.writestr(name, zf.read(name))
        return out.getvalue()


# ---------------------------------------------------------------------------
# Tabela de resumo: expansão + seleção robusta
# ---------------------------------------------------------------------------

_ITEM_ROW_TOKENS = (
    "{{item_", "{{ item_", "{{quantity", "{{ quantity",
    "{{unit_price", "{{ unit_price", "{{item_total", "{{ item_total",
)


def _row_text(row) -> str:
    return " ".join(cell.text for cell in row.cells)


def _is_item_template_row(row) -> bool:
    text = _row_text(row)
    return any(tok in text for tok in _ITEM_ROW_TOKENS)


def _find_first_item_row_index(table) -> Optional[int]:
    """Retorna o índice da PRIMEIRA linha que contém placeholders de item."""
    for i, row in enumerate(table.rows):
        if _is_item_template_row(row):
            return i
    return None


def _find_summary_table(slide):
    """Retorna a primeira tabela do slide que contém placeholders de item."""
    for shape in slide.shapes:
        if not getattr(shape, "has_table", False):
            continue
        if _find_first_item_row_index(shape.table) is not None:
            return shape.table
    return None


def _expand_summary_table_rows(slide, section):
    """
    Duplica a linha-template até totalizar `len(section.items)` linhas de item.
    Não substitui placeholders — isso fica para o motor central.
    """
    table = _find_summary_table(slide)
    if table is None:
        return

    template_idx = _find_first_item_row_index(table)
    if template_idx is None:
        return

    template_row_el = list(table.rows)[template_idx]._tr
    parent = template_row_el.getparent()

    # Remove eventuais linhas-template extras abaixo da primeira
    for row in list(table.rows)[template_idx + 1:]:
        if _is_item_template_row(row):
            parent.remove(row._tr)

    # Duplica para cada item adicional (a primeira já existe)
    anchor = template_row_el
    for _ in section.items[1:]:
        new_row_el = _deepcopy(template_row_el)
        parent.insert(parent.index(anchor) + 1, new_row_el)
        anchor = new_row_el


def _replace_summary_table_rows(slide, section):
    """
    Seleção robusta e determinística:
      1. Localiza a tabela do resumo (primeira com placeholders de item).
      2. Localiza a primeira linha de item via placeholders.
      3. Pega EXATAMENTE len(section.items) linhas a partir dela.
      4. Aplica os dados de cada item via motor central, por parágrafo.

    `display_index` é 1-based via enumerate — não usa item.item_index.
    """
    table = _find_summary_table(slide)
    if table is None:
        return

    first_idx = _find_first_item_row_index(table)
    if first_idx is None:
        return  # já substituído em passada anterior

    rows = list(table.rows)
    item_rows = rows[first_idx: first_idx + len(section.items)]
    if len(item_rows) < len(section.items):
        # Estrutura inesperada — aborta sem corromper o slide
        return

    for display_index, (row, item) in enumerate(zip(item_rows, section.items), start=1):
        item_total = item.quantity * item.unit_price
        row_data = {
            "item_display_index": str(display_index),
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
                _apply_to_paragraph(paragraph, row_data)
