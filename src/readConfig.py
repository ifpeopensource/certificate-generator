import csv
from typing import Dict, List


def readCSVConfig(file_path: str) -> Dict[List[str], List[List[str]]]:
    with open(file_path, encoding='utf8') as file:
        rows = list(csv.reader(file))
        if (len(rows) < 2):
            raise Exception(f"Invalid fields file '{file_path}'")
        return {
            "fields": rows[0],
            "values": rows[1:]
        }

def readTXTConfig(file_path: str) -> List[str]:
    with open(file_path, encoding='utf8') as file:
        rows = file.readlines()
        if (len(rows) < 1):
            raise Exception(f"Invalid fields file '{file_path}'")
        return rows
