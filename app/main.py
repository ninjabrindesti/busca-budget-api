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
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from pptx import Presentation

from app.services.pptx_generator import (
    replace_text_placeholders_on_slide,
    replace_named_images_on_slide,
)
from app.services.slide_duplicator import duplicate_slide_in_pptx

app = FastAPI(title="Ninja Brindes - Gerador de Proposta")


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
# Constantes do template
# ---------------------------------------------------------------------------

TEMPLATE_PATH = "templates/template_ninja.pptx"

# Índices 0-based no template original
ITEM_SLIDE_INDEX    = 8   # slide 9  → template de item individual
SUMMARY_SLIDE_INDEX = 9   # slide 10 → template de resumo/orçamento


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _format_currency(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _build_data(proposal: Proposal, item: Item, section: Section) -> dict:
    item_total    = item.quantity * item.unit_price
    freight_value = section.freight_value or 0.0
    section_total = sum(i.quantity * i.unit_price for i in section.items)

    return {
        # Proposta
        "proposal_number":   proposal.proposal_number,
        "client_name":       proposal.client_name,
        "payment_method":    proposal.payment_method,
        "delivery_date":     proposal.delivery_date,
        "notes":             proposal.notes,
        # Vendedor
        "seller_name":        proposal.seller_name,
        "seller_phone":       proposal.seller_phone,
        "seller_email":       proposal.seller_email,
        "seller_description": proposal.seller_description,
        "seller_image_url":   proposal.seller_image_url or "",
        # Item
        "item_name":          item.item_name,
        "item_subtitle":      item.item_subtitle,
        "item_index":         str(item.item_index),
        "item_display_index": str(item.item_index + 1),
        "item_description":   item.item_description,
        "item_code":          item.item_code,
        "quantity":           str(item.quantity),
        "unit_price":         _format_currency(item.unit_price),
        "item_total":         _format_currency(item_total),
        "item_image_url":     item.item_image_url,
        # Seção
        "section_total":  _format_currency(section_total),
        "freight":        _format_currency(freight_value),
        "freight_label":  section.freight_label,
    }


def _build_summary_data(proposal: Proposal, section: Section) -> dict:
    """Dados para o slide de resumo (slide 10). Usa o primeiro item como base."""
    first_item = section.items[0]
    data = _build_data(proposal, first_item, section)

    # No slide de resumo, listamos todos os itens da seção
    # Os placeholders são preenchidos com dados do primeiro item;
    # linhas extras são adicionadas via _expand_summary_table().
    return data


# ---------------------------------------------------------------------------
# Geração do PPTX
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
    # PASSO 1: Calcular quantos slides precisamos e duplicar de uma vez.
    #
    # Estrutura final desejada (índices 0-based após slides 0-7 fixos):
    #   Para cada section:
    #     item_count slides de item  (clones do slide 8)
    #     1 slide de resumo          (clone do slide 9)
    #   Slides 10+ originais (depoimentos, encerramento) mantidos.
    #
    # Estratégia:
    #   1. Remove slides 8 e 9 originais da ordem
    #   2. Insere os clones no lugar certo
    #
    # Implementação mais simples e robusta:
    #   - Duplicamos slide 8 (item) e slide 9 (resumo) o número de vezes certo
    #   - Reordenamos via manipulação direta do ZIP
    # -----------------------------------------------------------------------

    total_item_slides    = sum(len(s.items) for s in payload.sections)
    total_summary_slides = len(payload.sections)
    total_new_slides     = total_item_slides + total_summary_slides

    # Duplica slide de item (índice 8, 1-based = 9)
    # Precisamos de (total_item_slides - 1) cópias além do original
    # + total_summary_slides cópias do slide de resumo além do original
    item_copies    = total_item_slides - 1
    summary_copies = total_summary_slides - 1

    # Duplica slides de item
    for _ in range(item_copies):
        pptx_bytes = duplicate_slide_in_pptx(
            input_bytes=pptx_bytes,
            source_slide_index=ITEM_SLIDE_INDEX,
            copies=1,
            insert_after_index=ITEM_SLIDE_INDEX,
        )

    # Agora o slide de resumo original está em ITEM_SLIDE_INDEX + total_item_slides
    summary_original_index = ITEM_SLIDE_INDEX + total_item_slides

    # Duplica slides de resumo
    for _ in range(summary_copies):
        pptx_bytes = duplicate_slide_in_pptx(
            input_bytes=pptx_bytes,
            source_slide_index=summary_original_index,
            copies=1,
            insert_after_index=summary_original_index,
        )

    # -----------------------------------------------------------------------
    # PASSO 2: Reordenar slides para intercalar item/resumo por seção.
    #
    # Situação atual após duplicação:
    #   [0-7 fixos] [item×N] [resumo×M] [11+ originais]
    #
    # Situação desejada:
    #   [0-7 fixos] [sec1_item1, sec1_item2, sec1_resumo, sec2_item1, ...] [11+ originais]
    #
    # Fazemos isso reordenando os índices via ZIP manipulation.
    # -----------------------------------------------------------------------

    pptx_bytes = _reorder_slides(
        pptx_bytes=pptx_bytes,
        sections=payload.sections,
        total_item_slides=total_item_slides,
        total_summary_slides=total_summary_slides,
    )

    # -----------------------------------------------------------------------
    # PASSO 3: Preencher placeholders em cada slide com python-pptx
    # -----------------------------------------------------------------------

    prs = Presentation(io.BytesIO(pptx_bytes))

    # Preenche slide do vendedor (slide 4 = índice 3)
    seller_slide = prs.slides[3]
    seller_data = {
        "seller_name":        payload.proposal.seller_name,
        "seller_phone":       payload.proposal.seller_phone,
        "seller_email":       payload.proposal.seller_email,
        "seller_description": payload.proposal.seller_description,
        "seller_image_url":   payload.proposal.seller_image_url or "",
    }
    replace_text_placeholders_on_slide(seller_slide, seller_data)
    replace_named_images_on_slide(seller_slide, seller_data)

    # Preenche slides de item e resumo
    slide_cursor = ITEM_SLIDE_INDEX  # começa no índice 8

    for section in payload.sections:
        # Slides de item desta seção
        for item in section.items:
            if slide_cursor >= len(prs.slides):
                raise HTTPException(
                    status_code=500,
                    detail=f"Slide de item não encontrado. Índice: {slide_cursor}"
                )
            slide = prs.slides[slide_cursor]
            data  = _build_data(payload.proposal, item, section)
            replace_text_placeholders_on_slide(slide, data)
            replace_named_images_on_slide(slide, data)
            slide_cursor += 1

        # Slide de resumo desta seção
        if slide_cursor >= len(prs.slides):
            raise HTTPException(
                status_code=500,
                detail=f"Slide de resumo não encontrado. Índice: {slide_cursor}"
            )
        summary_slide = prs.slides[slide_cursor]
        summary_data  = _build_summary_data(payload.proposal, section)
        _expand_summary_table(summary_slide, section, payload.proposal)
        replace_text_placeholders_on_slide(summary_slide, summary_data)
        replace_named_images_on_slide(summary_slide, summary_data)
        slide_cursor += 1

    # Salva em memória
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

import io as _io
import re as _re
import zipfile as _zipfile
from lxml import etree as _etree

_NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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

    with _zipfile.ZipFile(_io.BytesIO(pptx_bytes), "r") as zf:
        prs_bytes = zf.read("ppt/presentation.xml")
        rels_bytes = zf.read("ppt/_rels/presentation.xml.rels")

        prs_root  = _etree.fromstring(prs_bytes)
        rels_root = _etree.fromstring(rels_bytes)

        rid_to_target = {
            rel.get("Id"): rel.get("Target")
            for rel in rels_root.findall(f"{{{_NS_REL}}}Relationship")
        }

        sld_id_lst = prs_root.find(f".//{{{_NS_P}}}sldIdLst")
        sld_id_els = list(sld_id_lst)

        def path_of(el):
            rid = el.get(f"{{{_NS_R}}}id")
            tgt = rid_to_target.get(rid, "")
            return f"ppt/{tgt}" if tgt else ""

        # Fatias da lista
        fixed_before  = sld_id_els[:ITEM_SLIDE_INDEX]          # slides 0-7
        item_slides   = sld_id_els[ITEM_SLIDE_INDEX : ITEM_SLIDE_INDEX + total_item_slides]
        summary_slides = sld_id_els[ITEM_SLIDE_INDEX + total_item_slides :
                                    ITEM_SLIDE_INDEX + total_item_slides + total_summary_slides]
        fixed_after   = sld_id_els[ITEM_SLIDE_INDEX + total_item_slides + total_summary_slides:]

        # Reconstrói a ordem intercalada
        new_order = list(fixed_before)
        item_ptr    = 0
        summary_ptr = 0
        for section in sections:
            n_items = len(section.items)
            for _ in range(n_items):
                new_order.append(item_slides[item_ptr])
                item_ptr += 1
            new_order.append(summary_slides[summary_ptr])
            summary_ptr += 1
        new_order.extend(fixed_after)

        # Substitui sldIdLst
        for el in list(sld_id_lst):
            sld_id_lst.remove(el)
        for el in new_order:
            sld_id_lst.append(el)

        buf["ppt/presentation.xml"] = _etree.tostring(
            prs_root, xml_declaration=True, encoding="utf-8", standalone=True
        )

        # Monta ZIP
        out = _io.BytesIO()
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

from copy import deepcopy as _deepcopy


def _expand_summary_table(slide, section: Section, proposal: Proposal):
    """
    Duplica as linhas de dados da tabela no slide de resumo para cada item.
    O template original tem 1 linha de item — adicionamos as demais.
    """
    first_item = section.items[0]

    for shape in slide.shapes:
        if not getattr(shape, "has_table", False):
            continue

        table = shape.table
        rows  = list(table.rows)

        # Identifica a linha de dados do primeiro item
        item_row_idx = None
        for i, row in enumerate(rows):
            row_text = " ".join(
                cell.text for cell in row.cells
            )
            if (
                first_item.item_code in row_text
                or str(first_item.item_index + 1) in row_text
                or "{{item_" in row_text
            ):
                item_row_idx = i
                break

        if item_row_idx is None:
            continue  # tabela sem linha de item reconhecível

        # Para cada item adicional, clona a linha e insere após a anterior
        tbl_element = table._tbl
        src_row_el  = rows[item_row_idx]._tr

        for extra_item in section.items[1:]:
            new_row_el = _deepcopy(src_row_el)

            # Substitui textos na nova linha
            for tc in new_row_el.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                if tc.text:
                    tc.text = (
                        tc.text
                        .replace(first_item.item_code,          extra_item.item_code)
                        .replace(str(first_item.item_index + 1), str(extra_item.item_index + 1))
                        .replace(first_item.item_name,          extra_item.item_name)
                        .replace(first_item.item_description,   extra_item.item_description)
                    )

            # Insere após a última linha adicionada
            src_row_el.addnext(new_row_el)
            src_row_el = new_row_el  # atualiza referência para próxima inserção

        break  # só processa a primeira tabela do slide
