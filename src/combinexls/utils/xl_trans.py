# encoding: utf-8
import os
from xls2xlsx import XLS2XLSX
import shutil
import re


def to_xlsx(parent_path):
    out_path = parent_path + "xlsx/"

    files_list = os.listdir(parent_path)
    xls_file_list = [xls_file for xls_file in files_list if xls_file.endswith(".xls")]

    os.makedirs(out_path, exist_ok=True)
    for xls_file in xls_file_list:
        in_file = parent_path + xls_file
        out_file = out_path + xls_file.split(".")[0] + ".xlsx"
        # # Read xls file
        x2x = XLS2XLSX(in_file)
        # # Write to xlsx file
        x2x.to_xlsx(out_file)

    xlsx_file_list = [
        xlsx_file
        for xlsx_file in files_list
        if xlsx_file.endswith(".xlsx") and not re.search(r"_[0-9]{10}", xlsx_file)
    ]
    if len(xlsx_file_list) > 0:
        for xlsx_file in xlsx_file_list:
            shutil.copy(parent_path + xlsx_file, out_path + xlsx_file)


if __name__ == "__main__":
    pass
    # to_xlsx("./", "./xlsx/")
