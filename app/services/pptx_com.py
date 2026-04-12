import os
import pythoncom
import win32com.client


def duplicate_slide_in_file(
    input_path: str,
    output_path: str,
    source_slide_index: int,
    copies: int = 1,
) -> None:
    pythoncom.CoInitialize()

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 0

    presentation = None

    try:
        abs_input = os.path.abspath(input_path)
        abs_output = os.path.abspath(output_path)

        presentation = app.Presentations.Open(
            abs_input,
            ReadOnly=False,
            Untitled=False,
            WithWindow=False,
        )

        # COM usa índice começando em 1
        for _ in range(copies):
            presentation.Slides(source_slide_index).Duplicate()

        presentation.SaveAs(abs_output)
        presentation.Close()

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass

        app.Quit()
        pythoncom.CoUninitialize()