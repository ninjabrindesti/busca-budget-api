import re
from pptx import Presentation

PLACEHOLDER_PATTERN = re.compile(r"{{\s*([\w\.]+)\s*}}")


def replace_text_placeholders(prs: Presentation, data: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape.text_frame is None:
                continue

            for paragraph in shape.text_frame.paragraphs:
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
