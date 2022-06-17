from pathlib import Path
from typing import Dict
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor


def generatePPTX(name: str, model: str, options: Dict[str, str], output_dir: str) -> str:
    prs = Presentation(model)
    slide = prs.slides[0]

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        frame = shape.text_frame

        if "{{name}}" in frame.text:
            frame.text = frame.text.replace("{{name}}", name)

            for paragraph in frame.paragraphs:
                paragraph.alignment = options['align']
                paragraph.font.size = Pt(options['font_size'])
                paragraph.font.color.rgb = RGBColor.from_string(
                    options['color'])

    file_path = Path(output_dir).joinpath(f"pptx/{name}.pptx")
    prs.core_properties.title = f"Certificate for {name}"
    prs.core_properties.author = "IFPE Open Source"
    prs.core_properties.comments = "Certificate generated using IFPE Open Source Certificate Generator\nCertificado gerado usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator"
    prs.save(file_path)
    return file_path
