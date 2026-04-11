import re
from pptx import Presentation

PLACEHOLDER_PATTERN = re.compile(r"{{\s*([\w\.]+)\s*}}")


def replace_text_placeholders(prs: Presentation, data: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape.text_frame is None:
                continue

            full_text = shape.text
            if not full_text:
                continue

            def repl(match):
                key = match.group(1)
                value = data.get(key)

                if value is None or value == "":
                    return match.group(0)

                return str(value)

            new_text = PLACEHOLDER_PATTERN.sub(repl, full_text)

            if new_text != full_text:
                shape.text = new_text
