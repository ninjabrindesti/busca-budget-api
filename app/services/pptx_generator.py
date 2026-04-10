import re
from pptx import Presentation

PLACEHOLDER_PATTERN = re.compile(r"{{\s*([\w\.]+)\s*}}")

def replace_text_placeholders(prs: Presentation, data: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape.text_frame is None:
                continue

            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    original_text = run.text
                    if not original_text:
                        continue

                    def repl(match):
                        key = match.group(1)
                        value = data.get(key, "")
                        return str(value) if value is not None else ""

                    run.text = PLACEHOLDER_PATTERN.sub(repl, original_text)