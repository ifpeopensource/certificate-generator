from PyPDF2 import PdfMerger


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
