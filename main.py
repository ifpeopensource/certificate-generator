import os
from pathlib import Path
import glob
from typing import Dict
import click
from pptx.enum.text import PP_ALIGN

from src.log import Log
from src.generatePPTX import generatePPTX
from src.PPTXtoPDF import PPTXtoPDF
from src.mergePDFs import mergePDFs

logger = Log()


@click.command()
@click.argument('output_file_path', nargs=1, required=False, default="./output/certificates.pdf")
@click.option('--model', '-m', default="./model.pptx", help="Model file path", show_default=True)
@click.option('--names', '-n', default="./names.txt", help="Names file path", show_default=True)
@click.option('--output-dir', '-o', default="./output", help="Output directory", show_default=True)
@click.option('--align', '-a', default="left", help="Paragraph alignment", show_default=True, type=click.Choice(['left', 'center', 'right', 'justify']))
@click.option('--font-size', '-f', default=18, help="Font size", show_default=True, type=int)
@click.option('--color', '-c', default="000000", help="Font color", show_default=True)
@click.option('--cpf-enable', 'cpf_enabled', is_flag=True, default=False, help="Enables CPF", show_default=True, type=bool)
def main(output_file_path: str, model: str, names: str, output_dir: str, align: str, font_size: int, color: str, cpf_enabled: bool) -> None:
    """ Generates certificates from a model PPTX file and a list of names.

        OUTPUT_FILE_PATH: Path to the output PDF file with the certificates. (default: ./output/certificates.pdf)
    """
    options = {}
    options['align'] = handleAlignOption(align)
    options['font_size'] = font_size
    options['color'] = color

    if (Path(output_dir).exists() and os.listdir(output_dir) != 0):
        logger.error(f"Output directory '{output_dir}' is not empty")
        return

    if (not Path(names).exists()):
        logger.error(f"Names file '{names}' does not exist")
        return

    if (not Path(model).exists()):
        logger.error(f"Model file '{model}' does not exist")
        return

    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(Path(output_dir).joinpath('pptx'), exist_ok=True)

    with open(names, encoding='utf8') as file:
        lines = file.readlines()

    if (len(lines) == 0):
        logger.error(f"Names file '{names}' is empty")
        return
    for name in lines:
        print(f"Generating certificate for {name}")
        generateCertificate(model, name.strip(), options, output_dir)

    print("All certificates generated. Merging...")
    file_paths = glob.glob(f"{output_dir}/*.pdf")
    mergePDFs(file_paths, output_file_path)
    print("Done.")


def handleAlignOption(align: str):
    if align == "left":
        return PP_ALIGN.LEFT
    elif align == "center":
        return PP_ALIGN.CENTER
    elif align == "right":
        return PP_ALIGN.RIGHT
    elif align == "justify":
        return PP_ALIGN.JUSTIFY
    else:
        return None


def generateCertificate(model: str, name: str, options: Dict[str, str], output_dir: str):
    pptx_path = generatePPTX(name, model, options, output_dir)
    PPTXtoPDF(pptx_path, output_dir)


if __name__ == '__main__':
    main()
