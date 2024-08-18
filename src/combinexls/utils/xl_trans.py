# encoding: utf-8
import os
from xls2xlsx import XLS2XLSX


def to_xlsx(parent_path, out_path):

    files_list = os.listdir(parent_path)
    xls_file_list = [xls_file for xls_file in files_list if xls_file.endswith(".xls")]
    print(xls_file_list)

    os.makedirs("./xlsx/", exist_ok=True)
    for xls_file in xls_file_list:
        in_file = parent_path + xls_file
        out_file = out_path + xls_file.split(".")[0] + ".xlsx"
        print(in_file, out_file)
        # # Read xls file
        x2x = XLS2XLSX(in_file)
        # # Write to xlsx file
        x2x.to_xlsx(out_file)


if __name__ == "__main__":
    pass
    # to_xlsx("./", "./xlsx/")
