import streamlit as st
import io
import os
from core.gov_standard import apply_nd30_standard
from core.internal_standard import apply_texo_internal_standard

# --- CONFIG ---
st.set_page_config(page_title="TEXO Document Master", page_icon="📄", layout="wide")

# --- STYLE PREMIUM ---
st.markdown("""
<style>
    .stApp { background-color: #0A1931 !important; color: #ffffff !important; }
    h1, h2, h3, h4, h5, h6, p, span, div, li, label, .stMarkdown { color: #ffffff !important; }
    .main-header { color: #FFD700 !important; font-weight: 800; font-size: 32px; text-align: center; border-bottom: 2px solid #FFD700; padding-bottom: 10px; margin-bottom: 20px; }
    .stButton>button { background: #152A4A !important; color: #FFD700 !important; border: 1px solid #FFD700 !important; border-radius: 12px; font-weight: bold; height: 3.5em; width: 100%; }
    .stButton>button:hover { background: #FFD700 !important; color: #0A1931 !important; transform: scale(1.02); transition: 0.2s; }
    .stSelectbox div[data-baseweb="select"] { background-color: #152A4A !important; color: white !important; }
    .footer { text-align: center; color: #888; font-size: 12px; margin-top: 50px; }
</style>
""", unsafe_allow_html=True)

# --- AUTH ---
def check_password():
    if "authenticated" not in st.session_state: st.session_state.authenticated = False
    if st.session_state.authenticated: return True
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h2 style='text-align: center; color: #FFD700;'>🏦 TEXO DOC MASTER</h2>", unsafe_allow_html=True)
        pwd = st.text_input("Mật khẩu truy cập:", type="password")
        if st.button("XÁC THỰC"):
            if pwd == "texo2026":
                st.session_state.authenticated = True
                st.rerun()
            else: st.error("❌ Truy cập không hợp lệ.")
    return False

if not check_password(): st.stop()

# --- MAIN ---
st.markdown("<div class='main-header'>📄 CHUẨN HÓA VĂN BẢN MASTER</div>", unsafe_allow_html=True)

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("### ⚙️ Cấu hình Chuẩn hóa")
    mode = st.selectbox("Bộ quy chuẩn Elite:", ["Quy định TEXO (12 quy tắc vàng)", "Nghị định 30/2020"])
    
    is_letterhead = False
    if "TEXO" in mode:
        paper_type = st.radio("Loại phôi giấy:", ["Giấy trắng thường", "Giấy Letterhead"], horizontal=True)
        is_letterhead = True if "Letterhead" in paper_type else False

    f_up = st.file_uploader("Tải hồ sơ (.docx)", type=["docx"])

with col2:
    st.markdown("### 🚀 Thực thi & Tải về")
    if f_up:
        st.info(f"Tệp đã chọn: **{f_up.name}**")
        if st.button("🚀 BẮT ĐẦU CHUẨN HÓA"):
            with st.spinner("Đang áp dụng bộ quy chuẩn..."):
                try:
                    in_path = "temp_in.docx"
                    out_path = f"Master_{f_up.name}"
                    with open(in_path, "wb") as f:
                        f.write(f_up.getbuffer())
                    
                    if "TEXO" in mode:
                        apply_texo_internal_standard(in_path, out_path, is_letterhead)
                    else:
                        apply_nd30_standard(in_path, out_path)
                    
                    st.success("🎉 Hoàn tất chuẩn hóa.")
                    with open(out_path, "rb") as fo:
                        st.download_button("📥 TẢI VỀ BẢN CHUẨN", fo, out_path)
                    
                    # Cleanup
                    if os.path.exists(in_path): os.remove(in_path)
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}")
    else:
        st.markdown("<div style='height: 200px; border: 2px dashed #333; border-radius: 12px; display: flex; align-items: center; justify-content: center; color: #666;'>Vui lòng tải tệp .docx để bắt đầu</div>", unsafe_allow_html=True)

st.markdown("<div class='footer'>TEXO Engineering Department | Version 2.0 (Standalone) | Hoàng Đức Vũ</div>", unsafe_allow_html=True)
