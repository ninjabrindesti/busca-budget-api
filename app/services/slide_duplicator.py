"""
slide_duplicator.py
-------------------
Duplicação de slides via manipulação direta do ZIP do PPTX (lxml puro).
Substitui o pptx_com.py que dependia de Win32COM / PowerPoint Windows.

Compatível com Railway (Linux), Docker, qualquer ambiente Unix.
"""

import copy
import io
import re
import zipfile
from lxml import etree

# Namespaces OOXML
NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT  = "http://schemas.openxmlformats.org/package/2006/content-types"

TAG_SLD_ID    = f"{{{NS_P}}}sldId"
TAG_SLD_ID_LST = f"{{{NS_P}}}sldIdLst"
ATTR_R_ID     = f"{{{NS_R}}}id"

REL_TYPE_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/slide"
)
SLIDE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "presentationml.slide+xml"
)


def _parse(data: bytes) -> etree._Element:
    return etree.fromstring(data)


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone=True)


def _slide_number(path: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", path)
    return int(m.group(1)) if m else 0


def _next_slide_number(existing_paths: set[str]) -> int:
    used = {_slide_number(p) for p in existing_paths if _slide_number(p)}
    n = 1
    while n in used:
        n += 1
    return n


def duplicate_slide_in_pptx(
    input_bytes: bytes,
    source_slide_index: int,  # 0-based
    copies: int = 1,
    insert_after_index: int | None = None,  # 0-based; None = after source
) -> bytes:
    """
    Duplica source_slide_index `copies` vezes no PPTX fornecido como bytes.
    Os clones são inseridos logo após o slide de origem (ou após insert_after_index).
    Retorna os bytes do novo PPTX.

    Preserva:
    - Backgrounds (herança do slideLayout/slideMaster)
    - Imagens embutidas (via cópia das relationships e media files)
    - Tabelas, formas, textos
    """
    buf: dict[str, bytes] = {}  # sobrescritas sobre o zip original

    with zipfile.ZipFile(io.BytesIO(input_bytes), "r") as zf:
        all_names = set(zf.namelist())

        # --- Lê presentation.xml e seus rels ---
        prs_xml   = zf.read("ppt/presentation.xml")
        prs_rels  = zf.read("ppt/_rels/presentation.xml.rels")
        prs_root  = _parse(prs_xml)
        rels_root = _parse(prs_rels)

        # Mapa rId → target (relativo a ppt/)
        rid_to_target: dict[str, str] = {
            rel.get("Id"): rel.get("Target")
            for rel in rels_root.findall(f"{{{NS_REL}}}Relationship")
        }

        # Ordem dos slides
        sld_id_lst = prs_root.find(f".//{TAG_SLD_ID_LST}")
        sld_id_els = list(sld_id_lst)

        def el_to_path(el) -> str:
            rid = el.get(ATTR_R_ID)
            target = rid_to_target.get(rid, "")
            return f"ppt/{target}" if target else ""

        ordered_paths = [el_to_path(el) for el in sld_id_els]

        # Slide de origem
        src_path = ordered_paths[source_slide_index]
        src_rels_path = f"ppt/slides/_rels/{src_path.split('/')[-1]}.rels"

        # Posição de inserção (após o índice fonte ou após insert_after_index)
        insert_after = source_slide_index if insert_after_index is None else insert_after_index

        # Contadores livres
        used_paths: set[str] = set(all_names)
        existing_ids = [int(el.get("id", 0)) for el in sld_id_els]
        next_id_num  = max(existing_ids, default=255) + 1
        existing_rids = set()
        for rel in rels_root.findall(f"{{{NS_REL}}}Relationship"):
            m = re.search(r"\d+", rel.get("Id", ""))
            if m:
                existing_rids.add(int(m.group(0)))
        next_rid_num = max(existing_rids, default=0) + 1

        new_sld_id_els = []  # (element, path) dos slides clonados

        for _ in range(copies):
            # --- Novo número e path ---
            n = _next_slide_number(used_paths)
            new_slide_path = f"ppt/slides/slide{n}.xml"
            new_rels_path  = f"ppt/slides/_rels/slide{n}.xml.rels"
            used_paths.add(new_slide_path)

            # --- Copia XML do slide ---
            src_bytes = buf.get(src_path) or zf.read(src_path)
            buf[new_slide_path] = src_bytes  # clonamos sem alterar conteúdo

            # --- Copia e remapeia .rels do slide ---
            if src_rels_path in all_names or src_rels_path in buf:
                src_rels_bytes = buf.get(src_rels_path) or zf.read(src_rels_path)
                slide_rels_root = _parse(src_rels_bytes)

                # Para cada rel de mídia, copia o arquivo de mídia
                for rel in slide_rels_root.findall(f"{{{NS_REL}}}Relationship"):
                    tgt = rel.get("Target", "")
                    if tgt.startswith("../media/"):
                        media_name = tgt.replace("../", "ppt/")
                        if media_name in all_names and media_name not in buf:
                            buf[media_name] = zf.read(media_name)

                buf[new_rels_path] = _serialize(slide_rels_root)
            else:
                # Cria rels mínimo
                min_rels = etree.Element(f"{{{NS_REL}}}Relationships")
                buf[new_rels_path] = _serialize(min_rels)

            # --- Registra em Content_Types.xml ---
            ct_bytes = buf.get("[Content_Types].xml") or zf.read("[Content_Types].xml")
            ct_root  = _parse(ct_bytes)
            part_name = f"/ppt/slides/slide{n}.xml"
            existing_parts = {el.get("PartName") for el in ct_root}
            if part_name not in existing_parts:
                override = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
                override.set("PartName", part_name)
                override.set("ContentType", SLIDE_CONTENT_TYPE)
            buf["[Content_Types].xml"] = _serialize(ct_root)

            # --- Adiciona Relationship em presentation.xml.rels ---
            new_rId = f"rId{next_rid_num}"
            next_rid_num += 1
            new_rel = etree.SubElement(rels_root, f"{{{NS_REL}}}Relationship")
            new_rel.set("Id", new_rId)
            new_rel.set("Type", REL_TYPE_SLIDE)
            new_rel.set("Target", f"slides/slide{n}.xml")
            rid_to_target[new_rId] = f"slides/slide{n}.xml"

            # --- Cria elemento sldId ---
            new_el = etree.Element(TAG_SLD_ID)
            new_el.set("id", str(next_id_num))
            new_el.set(ATTR_R_ID, new_rId)
            next_id_num += 1

            new_sld_id_els.append((new_el, new_slide_path))

        # --- Reconstrói sldIdLst com clones inseridos na posição certa ---
        for el in list(sld_id_lst):
            sld_id_lst.remove(el)

        # Reconstrói: antes do ponto de inserção, clones, depois
        for i, el in enumerate(sld_id_els):
            sld_id_lst.append(el)
            if i == insert_after:
                for new_el, _ in new_sld_id_els:
                    sld_id_lst.append(new_el)

        # Se insert_after >= len, adiciona no final
        if insert_after >= len(sld_id_els):
            for new_el, _ in new_sld_id_els:
                sld_id_lst.append(new_el)

        buf["ppt/presentation.xml"]        = _serialize(prs_root)
        buf["ppt/_rels/presentation.xml.rels"] = _serialize(rels_root)

        # --- Monta ZIP de saída ---
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as out_zf:
            written = set()
            # Escreve sobrescritas primeiro
            for name, data in buf.items():
                out_zf.writestr(name, data)
                written.add(name)
            # Copia o restante do zip original
            for name in zf.namelist():
                if name not in written:
                    out_zf.writestr(name, zf.read(name))

        return out.getvalue()


def duplicate_slide_in_file(
    input_path: str,
    output_path: str,
    source_slide_index: int,  # 1-based para manter compatibilidade com o código original
    copies: int = 1,
) -> None:
    """
    Interface compatível com a função Win32COM original.
    source_slide_index é 1-based (igual ao COM), convertemos internamente.
    """
    with open(input_path, "rb") as f:
        input_bytes = f.read()

    result = duplicate_slide_in_pptx(
        input_bytes=input_bytes,
        source_slide_index=source_slide_index - 1,  # converte para 0-based
        copies=copies,
        insert_after_index=source_slide_index - 1,  # insere após a fonte
    )

    with open(output_path, "wb") as f:
        f.write(result)
