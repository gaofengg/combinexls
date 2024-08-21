# encoding: utf-8
import re
import shutil
from typing import Generator
from pathlib import Path
from xls2xlsx import XLS2XLSX


#
def trans_to_xlsx(parent_path_) -> bool | Generator[Path, None, None]:
    parent_path = Path(parent_path_)
    out_path = parent_path / 'xlsx'

    file_iter = Path(parent_path).iterdir()

    if Path(out_path).exists():
        shutil.rmtree(out_path)
    out_path.mkdir(parents=True, exist_ok=True)

    has_xl = False
    for file_ in file_iter:
        file = file_.name
        if file.lower().endswith(".xls"):
            in_file = parent_path / file
            xfile_name = file.split(".")[0] + ".xlsx"
            out_file = out_path / xfile_name
            x2x = XLS2XLSX(str(in_file))
            x2x.to_xlsx(str(out_file))
            if not has_xl:
                has_xl = True
        elif file.lower().endswith(".xlsx") \
                and not re.search(r"_\d{10}", file):
            shutil.copy(parent_path / file, out_path / file)
            if not has_xl:
                has_xl = True

    if not has_xl:
        return False
    else:
        return Path(out_path).iterdir()


if __name__ == "__main__":
    pass
    # to_xlsx("./")
