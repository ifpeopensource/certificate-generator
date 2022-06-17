import os
from pathlib import Path
import sys
import glob
import subprocess
from pptx import Presentation
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfMerger


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

    merger.addMetadata({
        "/Author": 'IFPE Open Source',
        "/Title": "Certificates",
        "/Subject": "Certificates generated using IFPE Open Source Certificate Generator\nCertificados gerados usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator",
        "/Creator": "IFPE Open Source Certificate Generator",
    })

    output_path = sys.argv[1] if len(
        sys.argv) >= 2 else "./output/certificates.pdf"
    merger.write(output_path)
    merger.close()


def PPTXtoPDF(origin_file_path: str) -> None:
    subprocess.run(["libreoffice", "--headless", "--convert-to",
                   "pdf", "--outdir", "./output", origin_file_path])

    generated_file_path = f"./output/{Path(origin_file_path).stem}.pdf"

    reader = PdfFileReader(generated_file_path)
    writer = PdfFileWriter()

    writer.append_pages_from_reader(reader)
    writer.addMetadata({
        "/Author": 'IFPE Open Source',
        "/Title": f"Certificate for {Path(origin_file_path).stem}",
        "/Subject": "Certificate generated using IFPE Open Source Certificate Generator\nCertificado gerado usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator",
        "/Creator": "IFPE Open Source Certificate Generator",
    })

    with open(generated_file_path, "wb") as fp:
        writer.write(fp)


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
    prs.core_properties.title = f"Certificate for {name}"
    prs.core_properties.author = "IFPE Open Source"
    prs.core_properties.comments = "Certificate generated using IFPE Open Source Certificate Generator\nCertificado gerado usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator"
    prs.save(file_path)
    return file_path


if __name__ == '__main__':
    main()
