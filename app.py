# -*- coding: utf-8 -*-
"""TEXO – Chuẩn hóa thể thức văn bản Word (.docx).
Hai bộ quy chuẩn: Quy định TEXO (giấy thường / Letterhead) và Nghị định 30/2020/NĐ-CP.
Phòng Kỹ thuật – Công ty CP TEXO Tư vấn và Đầu tư.

Chạy:  streamlit run app.py
"""
import io
import os
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from core import ui
from core.engine import PROFILES, standardize

APP_VERSION = "2.0"
TITLE = "Chuẩn hóa thể thức văn bản"
MAX_FILES = 30
ui.setup("TEXO · " + TITLE, "📄")
ui.gate("Chuẩn hóa văn bản")

ss = st.session_state
ss.setdefault("ket_qua", None)


def _rules_md(key, paper):
    P = PROFILES[key]
    t, b, l, r = P["margins"].get(paper, P["margins"]["thuong"])
    line = (f"Exactly {P['line'][1]} pt" if P["line"][0] == "exact" else f"{P['line'][1]} lần")
    tline = (f"Exactly {P['table_line'][1]} pt" if P["table_line"][0] == "exact" else f"{P['table_line'][1]} lần")
    rows = [
        ("Khổ giấy", "A4 (giữ trang ngang nếu có)"),
        ("Lề T/D/T/P", f"{t} / {b} / {l} / {r} mm"),
        ("Phông", "Times New Roman"),
        ("Cỡ chữ", f"nội dung {P['body']} · tiêu đề căn giữa in hoa {P['title']} · dòng in hoa khác {P['upper']} · bảng {P['table']}"),
        ("Giãn dòng", f"{line} · trong bảng {tline}"),
        ("Cách đoạn", f"trước {P['before']} pt, sau {P['after']} pt"),
    ]
    if P["cell_indent_mm"]:
        rows.append(("Lề trong ô bảng", f"{P['cell_indent_mm']} mm"))
    if P["first_line_cm"]:
        rows.append(("Thụt đầu dòng", f"{P['first_line_cm']} cm (đoạn nội dung)"))
    return "\n".join(f"- **{a}:** {v}" for a, v in rows)


# --------------------------------------------------------------------------- thanh bên
with st.sidebar:
    st.markdown("### ⚙️ Tùy chọn sửa chữ")
    opts = {
        "spell": st.checkbox("Sửa lỗi gõ thường gặp", True, help="ko → không, quản lí → quản lý, sai xót → sai sót…"),
        "spaces": st.checkbox("Dọn khoảng trắng thừa", True, help="Hai dấu cách liền nhau; dấu cách trước dấu câu."),
        "comma": st.checkbox("Thêm dấu phẩy sau liên từ đầu câu", True,
                             help="Tuy nhiên, Do đó, Vì vậy, Theo đó… – chỉ khi đứng đầu câu."),
        "date": st.checkbox("Ngày tháng theo NĐ 30", True,
                            help="Trong cụm 'ngày … tháng … năm …': ngày dưới 10 và tháng 1, 2 thêm số 0. "
                                 "Không đụng ngày của văn bản viện dẫn dạng 18/8/2020."),
        "kinh_gui": st.checkbox("In đậm dòng “Kính gửi”", True),
        "indent": st.checkbox("Thụt đầu dòng đoạn nội dung (NĐ 30)", True,
                              help="Chỉ áp dụng cho bộ Nghị định 30; bỏ qua đoạn đánh số, gạch đầu dòng."),
    }
    opts["text"] = any(opts[k] for k in ("spell", "spaces", "comma", "date"))
    st.caption("Mọi chỗ sửa chữ đều ghi vào nhật ký để soát lại.")
    st.divider()
    ui.logout_button()
    st.caption("⚖️ Rà soát văn bản pháp luật, tiêu chuẩn viện dẫn bằng TEXO Legal Reviewer trước khi trình ký.")
    st.caption(f"Chuẩn hóa văn bản v{APP_VERSION} · Xây dựng: Hoàng Đức Vũ")

# --------------------------------------------------------------------------- trang chính
ui.header(TITLE, "Phòng Kỹ thuật · Đưa văn bản Word về đúng khổ giấy, lề, phông, cỡ chữ, giãn dòng theo "
                 "Quy định TEXO hoặc Nghị định 30/2020/NĐ-CP")

with st.expander("📘 Hướng dẫn sử dụng"):
    st.markdown("""
1. Chọn **bộ quy chuẩn**: *Quy định TEXO* (văn bản nội bộ, chọn thêm giấy thường hay Letterhead) hoặc *Nghị định 30/2020*.
2. Tải một hoặc nhiều file **.docx** (file .doc cũ: mở bằng Word, *Lưu thành* .docx trước).
3. Bấm **Chuẩn hóa** → xem ô tổng hợp và **nhật ký** từng file (lề cũ, số đoạn căn đều, chỗ sửa chữ) → tải về từng file hoặc cả gói ZIP.

**App giữ nguyên:** căn giữa / căn phải, chữ đậm – nghiêng, số hiệu và ngày của văn bản viện dẫn, hình ảnh
(đoạn có ảnh không bị ép giãn dòng). App **không** sửa nội dung chuyên môn – người lập vẫn đọc lại trước khi trình ký.
""")

c1, c2 = st.columns([1.1, 1])
with c1:
    st.markdown("##### 📐 Bộ quy chuẩn")
    key = st.radio("Bộ quy chuẩn", ["texo", "nd30"], horizontal=True, label_visibility="collapsed",
                   format_func=lambda k: PROFILES[k]["ten"])
    paper = "thuong"
    if key == "texo":
        paper = st.radio("Loại giấy", ["thuong", "letterhead"], horizontal=True,
                         format_func=lambda k: "Giấy trắng thường" if k == "thuong" else "Giấy Letterhead (logo in sẵn)")
    st.markdown("##### 📥 Tải văn bản cần chuẩn hóa")
    files = st.file_uploader("Chọn hoặc kéo thả file .docx – có thể chọn nhiều file", type=["docx"],
                             accept_multiple_files=True)
    if files and len(files) > MAX_FILES:
        st.warning(f"Đã chọn {len(files)} file; app xử lý {MAX_FILES} file đầu tiên.")
        files = files[:MAX_FILES]
    run = st.button("📄 Chuẩn hóa", type="primary", disabled=not files, use_container_width=True)
with c2:
    st.markdown(f"##### 📋 Quy tắc áp dụng – {PROFILES[key]['ten']}")
    st.markdown(_rules_md(key, paper))
    if key == "nd30" and not opts["indent"]:
        st.caption("Đang tắt thụt đầu dòng.")

# --------------------------------------------------------------------------- xử lý
if run and files:
    prog = st.progress(0.0, text="Đang chuẩn hóa…")
    out, logs, errs = {}, [], []
    for i, f in enumerate(files, 1):
        prog.progress((i - 1) / len(files), text=f"Đang xử lý {f.name} ({i}/{len(files)})…")
        try:
            data, log = standardize(f.getvalue(), key, paper, opts)
            name = os.path.splitext(f.name)[0] + "_chuan_hoa.docx"
            out[name] = data
            logs.append((f.name, name, log))
        except Exception as e:
            errs.append(f"{f.name}: không đọc được file ({e})")
    prog.empty()
    rows = []
    for src, name, lg in logs:
        for x in lg["sua_chu"]:
            rows.append({"File": src, **x})
    tong = pd.DataFrame([{
        "File": src, "Lề cũ (T/D/T/P mm)": "; ".join("/".join(str(v) for v in m) for m in lg["le_cu"]),
        "Đoạn": lg["doan"], "Căn đều": lg["can_deu"], "Giữ căn giữa/phải": lg["giu_can"],
        "Thụt đầu dòng": lg["thut_dau_dong"], "Đoạn có ảnh": lg["anh"], "Đoạn trong bảng": lg["o_bang"],
        "Chỗ sửa chữ": len(lg["sua_chu"])} for src, name, lg in logs])
    xl = io.BytesIO()
    with pd.ExcelWriter(xl, engine="openpyxl") as w:
        tong.to_excel(w, sheet_name="Tong hop", index=False)
        pd.DataFrame(rows or [{"File": "", "Vị trí": "", "Trước": "Không có chỗ sửa chữ", "Sau": ""}]) \
            .to_excel(w, sheet_name="Sua chu", index=False)
    ss.ket_qua = {"out": out, "logs": logs, "errs": errs, "tong": tong, "sua": rows, "xlsx": xl.getvalue(),
                  "bo": PROFILES[key]["ten"] + (" – Letterhead" if paper == "letterhead" else ""),
                  "stamp": datetime.now().strftime("%Y%m%d_%H%M")}

# --------------------------------------------------------------------------- kết quả
R = ss.ket_qua
if R:
    st.markdown("---")
    h1, h2 = st.columns([3, 1.2])
    h1.markdown(f"#### 📊 Kết quả – {R['bo']}")
    if R["out"]:
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            for n, d in R["out"].items():
                z.writestr(n, d)
            z.writestr("Nhat_ky_chuan_hoa.xlsx", R["xlsx"])
        h2.download_button("📥 Tải tất cả (ZIP)", zbuf.getvalue(), f"Van_ban_chuan_hoa_{R['stamp']}.zip",
                           "application/zip", type="primary", use_container_width=True)
    t = R["tong"]
    k = st.columns(4)
    ui.kpi(k[0], len(R["out"]), "File đã chuẩn hóa", f"{len(R['errs'])} file lỗi" if R["errs"] else "không lỗi",
           "red" if R["errs"] else "green")
    ui.kpi(k[1], int(t["Đoạn"].sum() + t["Đoạn trong bảng"].sum()) if len(t) else 0, "Đoạn đã định dạng",
           "ngoài bảng + trong bảng")
    ui.kpi(k[2], len(R["sua"]), "Chỗ sửa chữ", "xem nhật ký để soát", "orange" if R["sua"] else "green")
    ui.kpi(k[3], int(t["Giữ căn giữa/phải"].sum()) if len(t) else 0, "Đoạn giữ căn giữa/phải",
           "tiêu đề, địa danh – ngày tháng")
    for e in R["errs"]:
        st.error(e)
    ui.callout("✅ Đã chuẩn hóa thể thức. Mở file bằng Word kiểm tra lại ngắt trang, bảng rộng và khối chữ ký "
               "trước khi trình ký.", "green")

    tab1, tab2 = st.tabs(["📁 Từng file", f"✏️ Chỗ sửa chữ ({len(R['sua'])})"])
    with tab1:
        st.dataframe(t, hide_index=True, use_container_width=True)
        for src, name, lg in R["logs"]:
            a, b = st.columns([4, 1])
            a.markdown(f"📄 **{name}** · {len(lg['sua_chu'])} chỗ sửa chữ")
            b.download_button("📥 Tải về", R["out"][name], name, key=f"dl_{name}", use_container_width=True,
                              mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with tab2:
        if R["sua"]:
            st.dataframe(pd.DataFrame(R["sua"]), hide_index=True, use_container_width=True, height=420)
        else:
            st.success("Không có chỗ nào phải sửa chữ.")
        st.download_button("📥 Tải nhật ký (Excel)", R["xlsx"], f"Nhat_ky_chuan_hoa_{R['stamp']}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
elif not files:
    st.info("Tải một hoặc nhiều file .docx ở trên để bắt đầu.")
