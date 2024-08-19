# encoding: utf-8
from typing import Generator
from xls2xlsx import XLS2XLSX
import shutil
import re
from pathlib import Path


def to_xlsx(parent_path) -> bool | Generator[Path, None, None]:
    out_path = parent_path + "xlsx/"

    file_iter = Path(parent_path).iterdir()

    if Path(out_path).exists():
        shutil.rmtree(out_path)
    Path.mkdir(out_path, exist_ok=True)

    has_xl = False
    for file_ in file_iter:
        file = file_.name
        if file.lower().endswith(".xls"):
            in_file = parent_path + file
            out_file = out_path + file.split(".")[0] + ".xlsx"
            x2x = XLS2XLSX(in_file)
            x2x.to_xlsx(out_file)
            if not has_xl:
                has_xl = True
        elif file.lower().endswith(".xlsx") and not re.search(r"_[0-9]{10}", file):
            in_file = parent_path + file
            out_file = out_path + file.split(".")[0] + ".xlsx"
            shutil.copy(parent_path + file, out_path + file)
            if not has_xl:
                has_xl = True
    
    if not has_xl:
        return False
    else:
        return Path(out_path).iterdir()


if __name__ == "__main__":
    # pass
    to_xlsx("./")
