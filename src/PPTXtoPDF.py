import subprocess
from pathlib import Path

from PyPDF2 import PdfFileReader, PdfFileWriter


def PPTXtoPDF(file_path: Path, dir: str) -> None:

    dir: Path = Path(dir).resolve()

    # TODO Improve tmpfolder generation
    tmpfolder = str(dir.parent) + "/tmp/" + "".join(str(file_path.stem).split()[0] + str(file_path.stem).split()[-1]) + "/0/"
    
    
    subprocess.run(["libreoffice", "--headless", "--convert-to",
                "pdf", "--outdir", str(dir), file_path,
                f"-env:UserInstallation=file://{tmpfolder}"], stdout=subprocess.DEVNULL)


    generated_file_path = dir.joinpath(
        file_path.stem + ".pdf")

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
