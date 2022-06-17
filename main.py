import os
from pathlib import Path
import glob
import subprocess
from typing import Dict
import click
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
from pptx.dml.color import RGBColor
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfMerger


@click.command()
@click.argument('output_file_path', nargs=1, required=False, default="./output/certificates.pdf")
@click.option('--model', '-m', default="./model.pptx", help="Model file path", show_default=True)
@click.option('--names', '-n', default="./names.txt", help="Names file path", show_default=True)
@click.option('--output-dir', '-o', default="./output", help="Output directory", show_default=True)
@click.option('--align', '-a', default="left", help="Paragraph alignment", show_default=True, type=click.Choice(['left', 'center', 'right']))
@click.option('--font-size', '-f', default=18, help="Font size", show_default=True, type=int)
@click.option('--color', '-c', default="000000", help="Font color", show_default=True)
def main(output_file_path: str, model: str, names: str, output_dir: str, align: str, font_size: int, color: str) -> None:
    """ Generates certificates from a model PPTX file and a list of names.

        OUTPUT_FILE_PATH: Path to the output PDF file with the certificates. (default: ./output/certificates.pdf)
    """
    options = {}
    options['align'] = handleAlignOption(align)
    options['font_size'] = font_size
    options['color'] = color

    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(Path(output_dir).joinpath('pptx'), exist_ok=True)

    with open(names, encoding='utf8') as file:
        lines = file.readlines()
        for name in lines:
            print(f"Generating {name.strip()} certificate")
            pptx_path = genSlide(name.strip(), model, options, output_dir)
            PPTXtoPDF(pptx_path, output_dir)

    file_paths = glob.glob(f"{output_dir}/*.pdf")
    mergePDFs(file_paths, output_file_path)


def handleAlignOption(align: str):
    if align == "left":
        return PP_ALIGN.LEFT
    elif align == "center":
        return PP_ALIGN.CENTER
    elif align == "right":
        return PP_ALIGN.RIGHT
    else:
        return None


def mergePDFs(file_paths: list, output_path: str) -> None:
    merger = PdfMerger()
    for file_path in file_paths:
        merger.append(file_path)

    merger.addMetadata({
        "/Author": 'IFPE Open Source',
        "/Title": "Certificates",
        "/Subject": "Certificates generated using IFPE Open Source Certificate Generator\nCertificados gerados usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator",
        "/Creator": "IFPE Open Source Certificate Generator",
    })

    merger.write(output_path)
    merger.close()


def PPTXtoPDF(file_path: str, dir: str) -> None:
    subprocess.run(["libreoffice", "--headless", "--convert-to",
                   "pdf", "--outdir", dir, file_path])

    generated_file_path = Path(dir).joinpath(
        Path(file_path).stem + ".pdf")

    reader = PdfFileReader(generated_file_path)
    writer = PdfFileWriter()

    writer.append_pages_from_reader(reader)
    writer.addMetadata({
        "/Author": 'IFPE Open Source',
        "/Title": f"Certificate for {Path(file_path).stem}",
        "/Subject": "Certificate generated using IFPE Open Source Certificate Generator\nCertificado gerado usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator",
        "/Creator": "IFPE Open Source Certificate Generator",
    })

    with open(generated_file_path, "wb") as fp:
        writer.write(fp)


def genSlide(name: str, model: str, options: Dict[str, str], output_dir: str) -> str:
    prs = Presentation(model)
    slide = prs.slides[0]

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        frame = shape.text_frame

        if "{{name}}" in frame.text:
            frame.alignment = options['align']
            frame.text = frame.text.replace("{{name}}", name)

            for paragraph in frame.paragraphs:
                paragraph.font.size = Pt(options['font_size'])
                paragraph.font.color.rgb = RGBColor.from_string(
                    options['color'])

    file_path = Path(output_dir).joinpath(f"pptx/{name}.pptx")
    prs.core_properties.title = f"Certificate for {name}"
    prs.core_properties.author = "IFPE Open Source"
    prs.core_properties.comments = "Certificate generated using IFPE Open Source Certificate Generator\nCertificado gerado usando o IFPE Open Source Certificate Generator\nhttps://github.com/ifpeopensource/certificate-generator"
    prs.save(file_path)
    return file_path


if __name__ == '__main__':
    main()
