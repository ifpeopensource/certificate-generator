import subprocess
from pathlib import Path

from PyPDF2 import PdfFileReader, PdfFileWriter


def PPTXtoPDF(file_path: str, dir: str) -> None:
    subprocess.run(["libreoffice", "--headless", "--convert-to",
                "pdf", "--outdir", dir, file_path], stdout=subprocess.DEVNULL)

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
