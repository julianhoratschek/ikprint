from pathlib import Path
import xml.etree.ElementTree as xml
import re
import itertools as itt
from zipfile import ZipFile

import subprocess
import tempfile

# r1
    # c1 name
    # c2 birth
    # c3 room
# r2
    # c1 "Wohnort"
    # c2 Addr
# r3
    # c1 "Krankenkasse"
    # c2 insurance
# r4
    # c1 "Beruf"
    # c2 occupation
# r5
    # c1 "Wahlleistungen"
    # c2 "Chef"
    # c3 ??
    # c4 ??
    # c5 "Doppel: "
    # c6 "Regel:"
    # c7 "301:"
    # c8 ???
    # c9 "Kein PS:"
# r6
    # c1 "Teammitglieder"
    # c2 doctor
    # c3 "Psych:" <psych>
    # c4 "Physio:"
# r7
    # c1 "Aufnahme"
# r8
    # c1 "Entlassung geplant"
    # c2 date
# r9
    # c1 "Verlängerung bis"
    # c2 ???
# r10
    # c1 "Cave"
    # c2 ???
# r11
    # c1 "Allergien"
    # c2 allergies
# r12
    # c1 "Kost"
    # c2 ???
# r13
    # c1 "Diagnosen"
    # c2 "Schmerz"
    # c3 PAIN_DIAG
# r14
    # c1 ???
    # c2 "Fehlgebrauch"
    # c3 MISUSE_DIAG
# r15
    # c1
    # c2 "Komorbiditäten"
    # c3 PSYCH_DIAG
# r16
     # c1
     # c2 "Komorbidität"
     # c3 PHYS_DIAG
# r17
    # c1
    # c2 "Midas-Score"
    # c3 ???
# r18
    # c1
    # c2 "Medikation aktuell:"
    # c3 Current_acute_medis
    # c4 Current_base_medis
# r19
    # c1
    # c2
    # c3 Sonstige
# r20
    # c1
    # c2 "Medikation früher:"
    # c3 acute_previous
    # c4 base_previous
# r21 (2911)

offset = "\n" * 34

db_path = Path(".")
namespaces = { "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" }

icd10_pattern = re.compile(r"[A-Z]\d+(?:\.\d+)?")
diagnoses_row_start = 12
diagnoses_row_end = 17
column_height = 4
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
         for line in itt.zip_longest(*list(itt.batched(diagnoses, column_height)), fillvalue="")]
    )

    # Create and show file to print
    with tempfile.NamedTemporaryFile(delete_on_close=False) as tmp_file:
        tmp_file.write(output_text.encode("utf-8"))
        tmp_file.close()

        subprocess.run(["notepad", tmp_file.name])
        # subprocess.run(["print", '/d:""', tmp_file.name])

