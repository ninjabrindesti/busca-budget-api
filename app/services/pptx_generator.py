import re
from pptx import Presentation

PLACEHOLDER_PATTERN = re.compile(r"{{\s*([\w\.]+)\s*}}")


def _replace_in_text_frame(text_frame, data: dict):
    for paragraph in text_frame.paragraphs:
        original_text = "".join(run.text for run in paragraph.runs)

        if not original_text:
            continue

        def repl(match):
            key = match.group(1)
            value = data.get(key)

            if value is None or value == "":
                return match.group(0)

            return str(value)

        new_text = PLACEHOLDER_PATTERN.sub(repl, original_text)

        if new_text != original_text and len(paragraph.runs) > 0:
            paragraph.runs[0].text = new_text

            for run in paragraph.runs[1:]:
                run.text = ""


def replace_text_placeholders(prs: Presentation, data: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame is not None:
                            _replace_in_text_frame(cell.text_frame, data)
                continue

            if hasattr(shape, "text_frame") and shape.text_frame is not None:
                _replace_in_text_frame(shape.text_frame, data)
