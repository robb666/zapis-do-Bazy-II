# zapis-do-Bazy-II
Transfers data from .pdf files to .xlsx database. Recognizes PESEL and REGON by the checksum. If it finds a REGON then it downloads the entity's data from Regon API. It recognizes the names by comparing them with the file names.txt, in more difficult to parse cases it compares the vehicle brands with the brands stored in the brand.txt file.

There is a need of some sort of learning mechanism to fully automate this process as layout on the documents changes and program requires maitenance.
