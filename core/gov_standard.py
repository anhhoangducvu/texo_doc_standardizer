# -*- coding: utf-8 -*-
"""Chuẩn Nghị định 30/2020/NĐ-CP – giữ tên hàm cũ, gọi bộ máy core/engine.py.
Dùng dòng lệnh:  python -m core.gov_standard <file.docx>
"""
import os
import sys

from core.engine import standardize_file


def apply_nd30_standard(input_path, output_path, opts=None):
    return standardize_file(input_path, output_path, "nd30", "thuong", opts)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        src = sys.argv[1]
        out = os.path.join(os.path.dirname(src) or ".", "ChuanHoa_ND30_" + os.path.basename(src))
        apply_nd30_standard(src, out)
        print("Đã lưu:", out)
    else:
        print("Cách dùng: python -m core.gov_standard <file.docx>")
