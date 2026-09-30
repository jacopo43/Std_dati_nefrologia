PATIENT WORKBOOK CONVERTER — UPDATED VERSION
============================================

EXPECTED FOLDER STRUCTURE
-------------------------
final files/
|-- Link.xlsx
|-- PatientConverter.exe              optional, after building
|-- dict/
|   |-- Drugs_dict.xlsx
|   |-- Com_dict.xlsx
|   `-- Interact_dict.xlsx
|-- working files/
|   `-- dataset_finale_base.xlsx
|-- row_data/
|   `-- patient workbooks (.xls, .xlsx, or .ods)
|-- output/
`-- application/
    |-- app.py
    |-- converter.py
    |-- requirements.txt
    |-- install_and_run.bat
    |-- build_exe.bat
    `-- PatientConverter.spec

RUN WITH PYTHON
---------------
1. Keep the files in the structure shown above, or select them manually in the app.
2. Double-click application/install_and_run.bat.
3. Select or drag the patient workbook.
4. Review or change the LINK, template, and dictionary paths.
5. Choose an output folder and click Run conversion.

BUILD WINDOWS EXE
-----------------
1. Double-click application/build_exe.bat on Windows.
2. PatientConverter.exe is copied into the parent final files folder.
3. The EXE may use the default folder structure, but configuration files can
   also be selected manually from any location.

DICTIONARY REQUIREMENTS
-----------------------
Drugs_dict.xlsx requires:
- canonical_substance
- aliases
Optional supported columns include atc_code and match_roots.

Com_dict.xlsx requires:
- canonical_name
- aliases
- match_roots
The updated logic also supports uso_automatico, aliases_da_revisionare, and
aliases_non_automatici.

Interact_dict.xlsx requires:
- Conseguenze
- Effetti
The updated logic also supports Uso automatico, Termini da revisionare, and
Termini non automatici.
