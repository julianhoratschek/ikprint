from pathlib import Path
import xml.etree.ElementTree as xml
import re
import itertools as itt
from zipfile import ZipFile
from datetime import datetime

from typing import Optional

import subprocess
import tempfile
import argparse

# Offset before text in printed text
offset = "\n" * 34

# Where to find the admission files in .docx format
db_path = Path(__file__).parent

# XML namespaces for docx
namespaces = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}

# Pattern to find icd10 codes
icd10_pattern = re.compile(r"[A-Z]\d{2}(?:\.\d+)?")

# Which row to start reading diagnoses from
row_start = 12

# Which row to end reading diagnoses from
row_end = 17

# How many icd10-codes to fit into one column in the output
column_height = 4

# How many icd10-codes to fit into one row in the output
row_length = 7


# Using lists instead of sets to retain ordering
def refinement_loop(diagnoses: list[str]):
    """Run input loop to add or remove icd-codes separated by whitespace and
    display the result as formatted text in notepad
    :param diagnoses: list of strings representing diagnoses icd10-codes"""

    while True:
        # Transform and put output text together
        output_text = offset + "\n\n".join([
            " ".join([f"{icd:<10}" for icd in line])
            for line in itt.zip_longest(
                *list(itt.batched(diagnoses, column_height)),
                fillvalue=""
            )
        ])

        print(output_text)

        # Get user refinements
        cmd = input("\n\n[RETURN] quit;\n"
                    "[+] add; [-] remove space separated list: ")

        if not cmd:
            break

        # Read in changes
        changes = {'+': [], '-': []}
        active_list = changes['+']
        for c in cmd.split():
            if c in changes.keys():
                active_list = changes[c]
                continue
            active_list.append(c)

        # Apply changes to diagnoses
        for elem in changes['+']:
            if elem not in diagnoses:
                diagnoses.append(elem)

        for elem in changes['-']:
            while elem in diagnoses:
                diagnoses.remove(elem)

    # Create and show file to print
    with tempfile.NamedTemporaryFile(delete_on_close=False) as tmp_file:
        tmp_file.write(output_text.encode("utf-8"))
        tmp_file.close()

        subprocess.run(["notepad", tmp_file.name])


def get_patient_path(patient_name: str) -> Optional[Path]:
    """Determine correct input file from patient name"""

    patient_matches = [path for path in db_path.glob("*.docx")
                       if patient_name in path.stem.lower()]

    if not patient_matches:
        return None

    # Select correct file on multiple matches
    if (matches_count := len(patient_matches)) == 1:
        return patient_matches[0]

    patient_matches.sort()
    output_list = "\n".join([
        f"[{n:>2}]: {patient.name:.>50}"
        for n, patient in enumerate(patient_matches, start=1)
    ])

    while True:
        print(output_list)
        selection = input(f"Select correct file (1-{matches_count}): ")

        if not selection.isdecimal():
            print("!! You must select a number")
            continue

        idx = int(selection)
        if idx < 1 or idx > matches_count:
            print(f"!! Your selection must be within 1 and {matches_count}")
            continue

        return patient_matches[idx - 1]


def get_diagnoses(file_path: Path) -> list[str]:
    """Read icd-10 codes from input file"""

    # Read xml data from docx file
    with ZipFile(file_path, "r") as archive:
        with archive.open("word/document.xml", "r") as doc:
            tree = xml.fromstringlist(doc.readlines())

    # Extract icd10 codes from file
    diagnoses = []

    # Access first table through tree-indices, look for relevant rows
    for row in tree[0][0].findall("w:tr", namespaces)[row_start:row_end]:

        # Select relevant cols (only col idx 2 contains diagnoses)
        col = row.findall("w:tc", namespaces)[2]

        # Extract icd10-diagnoses via regex
        diagnoses.extend([
            icd[0]
            for icd in icd10_pattern.finditer("".join(col.itertext()))
        ])

    return diagnoses


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="ik Print",
        description="Extract ICD10-codes from docx to print for IK-Form"
    )

    parser.add_argument("-f", type=Path,
                        help="complete file name to load from database folder")

    args = parser.parse_args()

    if args.f is not None:
        file_path = db_path / args.f.name
        if not file_path.exists():
            input(f"!! Could not find {file_path}")
            exit(1)

    else:
        patient_name = input("Patient name: ").lower()
        file_path = get_patient_path(patient_name)
        if file_path is None:
            input(f"!! Could not find any files matching <{patient_name}>")
            exit(1)

    diagnoses = get_diagnoses(file_path)

    # Make sure, diagnoses will fit on paper
    if len(diagnoses) > (diagnoses_max := column_height * row_length):
        input(f"!! Too many diagnoses (won't fit on paper). Current max: {diagnoses_max}")
        exit(1)

    refinement_loop(diagnoses)


