from pathlib import Path
import xml.etree.ElementTree as xml
import re
import itertools as itt
from zipfile import ZipFile

import subprocess
import tempfile


# Offset before text in printed text
offset = "\n" * 34

# Where to find the admission files in .docx format
db_path = Path(".")

# XML namespaces for docx
namespaces = { "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" }

# Pattern to find icd10 codes
icd10_pattern = re.compile(r"[A-Z]\d+(?:\.\d+)?")

# Which row to start reading diagnoses from
diagnoses_row_start = 12

# Which row to end reading diagnoses from
diagnoses_row_end = 17

# How many icd10-codes to fit into one column in the output
column_height = 4

# How many icd10-codes to fit into one row in the output
row_length = 7


if __name__ == "__main__":
    # Find files matching patient name
    patient_name = input("Name: ").lower()
    patient_matches = [path for path in db_path.glob("*.docx") if patient_name in path.stem.lower()]

    if not patient_matches:
        print(f"!! Could not find any files matching <{patient_name}>")
        exit(0)

    # Select correct file on multiple matches
    if (matches_count := len(patient_matches)) > 1:
        patient_matches.sort(reverse=True)
        output_list = [f"[{n:>2}]: {patient.name:.>50}" for n, patient in enumerate(patient_matches, start=1)]

        while True:
            for line in output_list:
                print(line)
            selection = input(f"Select correct file (1-{matches_count}): ")

            if not selection.isdecimal():
                print("!! You must select a number")
                continue

            idx = int(selection)
            if idx < 1 or idx > matches_count:
                print(f"!! Your selection must be within 1 and {matches_count}")
                continue

            patient_matches = [patient_matches[idx - 1]]
            break

    # Read xml data from docx file
    with ZipFile(patient_matches[0], "r") as archive:
        with archive.open("word/document.xml", "r") as doc:
            tree = xml.fromstringlist(doc.readlines())

    # Extract icd10 codes from file
    diagnoses = []
    for row in tree[0][0].findall("w:tr", namespaces)[diagnoses_row_start:diagnoses_row_end]:
        col = row.findall("w:tc", namespaces)[2]
        found = [icd[0] for icd in icd10_pattern.finditer("".join(col.itertext()))]
        diagnoses.extend(found)

    # Make sure, diagnoses will fit on paper
    if len(diagnoses) > (diagnoses_max := column_height * row_length):
        print(f"!! Too many diagnoses (won't fit on paper). Current max: {diagnoses_max}")
        exit(0)

    # Transform and put output text together
    output_text = offset + "\n\n".join(
        [" ".join([f"{icd:<10}" for icd in line])
         for line in itt.zip_longest(
            *list(itt.batched(diagnoses, column_height)), fillvalue="")]
    )

    # Create and show file to print
    with tempfile.NamedTemporaryFile(delete_on_close=False) as tmp_file:
        tmp_file.write(output_text.encode("utf-8"))
        tmp_file.close()

        subprocess.run(["notepad", tmp_file.name])
        # subprocess.run(["print", '/d:""', tmp_file.name])

