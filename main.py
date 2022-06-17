import os
import sys
import glob
import subprocess
from pptx import Presentation
from PyPDF2 import PdfMerger


def main():
    os.makedirs("./output", exist_ok=True)
    os.makedirs("./output/pptx", exist_ok=True)

    with open('names.txt', encoding='utf8') as file:
        lines = file.readlines()
        for name in lines:
            print(f"Generating {name.strip()} certificate")
            pptx_path = genSlide(name.strip(), './model.pptx')
            PPTXtoPDF(pptx_path)

    file_paths = glob.glob("./output/*.pdf")
    mergePDFs(file_paths)


def mergePDFs(file_paths: list) -> None:
    merger = PdfMerger()
    for file_path in file_paths:
        merger.append(file_path)

    output_path = sys.argv[1] if len(
        sys.argv) >= 2 else "./output/certificates.pdf"
    merger.write(output_path)
    merger.close()


def PPTXtoPDF(file_path: str) -> None:
    subprocess.run(["libreoffice", "--headless", "--convert-to",
                   "pdf", "--outdir", "./output", file_path])


def genSlide(name: str, model: str) -> str:
    prs = Presentation(model)
    slide = prs.slides[0]

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        for paragraph in shape.text_frame.paragraphs:
            if "{{name}}" in paragraph.text:
                paragraph.text = paragraph.text.replace("{{name}}", name)

    file_path = f"./output/pptx/{name}.pptx"
    prs.save(file_path)
    return file_path


if __name__ == '__main__':
    main()
