# -*- coding: utf-8 -*-
"""Chuẩn nội bộ TEXO – giữ tên hàm cũ, gọi bộ máy core/engine.py.
Dùng dòng lệnh:  python -m core.internal_standard <file.docx> [--letterhead]
"""
import os
import sys

from core.engine import standardize_file


def apply_texo_internal_standard(input_path, output_path, is_letterhead=False, opts=None):
    return standardize_file(input_path, output_path, "texo", "letterhead" if is_letterhead else "thuong", opts)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        src = sys.argv[1]
        out = os.path.join(os.path.dirname(src) or ".", "ChuanHoa_TEXO_" + os.path.basename(src))
        apply_texo_internal_standard(src, out, "--letterhead" in sys.argv)
        print("Đã lưu:", out)
    else:
        print("Cách dùng: python -m core.internal_standard <file.docx> [--letterhead]")
