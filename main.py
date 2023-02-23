import os
from pathlib import Path
import glob
import shutil
import concurrent.futures
from typing import Dict, List
import click

from src.log import Log
from src.generatePPTX import generatePPTX
from src.PPTXtoPDF import PPTXtoPDF
from src.mergePDFs import mergePDFs
from src.readConfig import readCSVConfig, readTXTConfig

logger = Log()


@click.command()
@click.argument('output_file_path', nargs=1, required=False, default="./output/certificates.pdf")
@click.option('--model', '-m', default="./model.pptx", help="Model file path", show_default=True)
@click.option('--data', '-d', help="Data file path", required=True)
@click.option(
    '--multiple-fields/--name-only',
    help="Whether to fill in multiple fields or just name fields",
    prompt=True
)
@click.option('--output-dir', '-o', default="./output", help="Output directory", show_default=True)
@click.option(
    '--align',
    '-a',
    default="left",
    help="Paragraph alignment",
    show_default=True,
    type=click.Choice(['left', 'center', 'right', 'justify'])
)
@click.option('--font-size', '-f', default=18, help="Font size", show_default=True, type=int)
@click.option('--color', '-c', default="000000", help="Font color", show_default=True)
def main(
    output_file_path: str,
    model: str,
    data: str,
    multiple_fields: bool,
    output_dir: str,
    align: str,
    font_size: int,
    color: str
) -> None:
    """ Generates certificates from a model PPTX file and a list of names.

        OUTPUT_FILE_PATH: Path to the output PDF file with the certificates. (default: ./output/certificates.pdf)
    """
    options = {}
    options['align'] = align
    options['font_size'] = font_size
    options['color'] = color

    if (Path(output_dir).exists() and os.listdir(output_dir) != 0):
        logger.error(f"Output directory '{output_dir}' is not empty")
        print(f"Should {output_dir} be emptied? (y/n): ", end="")
        if (input() == "y"):
            shutil.rmtree(output_dir)
        else:
            return

    if (not Path(data).exists()):
        logger.error(f"Fields file '{data}' does not exist")
        return

    if (not Path(model).exists()):
        logger.error(f"Model file '{model}' does not exist")
        return

    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(Path(output_dir).joinpath('pptx'), exist_ok=True)

    try:
        with concurrent.futures.ProcessPoolExecutor() as executor:
            futures = []
            if multiple_fields:
                parsed_data = readCSVConfig(data)
                for value_row in parsed_data['values']:
                    print(f"Generating certificate for {value_row[0]}")
                    futures.append(executor.submit(
                    generateCertificate, model, parsed_data['fields'],
                                        value_row, options, output_dir))
            else:
                parsed_data = readTXTConfig(data)
                for name in parsed_data:
                    print(f"Generating certificate for {name}")
                    futures.append(executor.submit(
                    generateCertificate, model, ['name'], [name], options, output_dir))
    except Exception as e:
        print(e)
        print(type(e))

    print("All certificates generated. Merging...")
    file_paths = glob.glob(f"{output_dir}/*.pdf")
    mergePDFs(file_paths, output_file_path)
    print("Cleaning tmp files...")
    shutil.rmtree("./tmp")
    print("Done.")


def generateCertificate(model: str, fields: List[str], data: List[str], options, output_dir: str):
    pptx_path = generatePPTX(model, fields, data, options, output_dir)
    PPTXtoPDF(pptx_path, output_dir)


if __name__ == '__main__':
    main()
