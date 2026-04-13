def _expand_summary_table(slide, section):
    
    """
    Expande a tabela do slide de resumo criando 1 linha por item da section,
    já preenchendo cada linha com os dados do item correspondente.
    Funciona com placeholders com ou sem espaços: {{item_code}} ou {{ item_code }}.
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

        row_elements = [template_row_el]
        for _ in section.items[1:]:
            new_row_el = _deepcopy(template_row_el)
            parent.insert(parent.index(template_row_el) + len(row_elements), new_row_el)
            row_elements.append(new_row_el)

        for row_el, item in zip(row_elements, section.items):
            item_total = item.quantity * item.unit_price

            replacements = {
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

            for tc in row_el.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                if not tc.text:
                    continue

                for key, value in replacements.items():
                    tc.text = _re.sub(r"{{\s*" + _re.escape(key) + r"\s*}}", value, tc.text)

        break
