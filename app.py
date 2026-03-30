import streamlit as st
import io
import os
import zipfile
from core.gov_standard import apply_nd30_standard
from core.internal_standard import apply_texo_internal_standard

# --- CONFIG ---
st.set_page_config(page_title="TEXO Document Standardizer", page_icon="📄", layout="wide")

# --- STYLE PREMIUM ---
st.markdown("""
<style>
    /* --- TỐI ƯU HÓA LÀM SẠCH CSS CHO CẢ 2 CHẾ ĐỘ --- */
    h1, h2, h3, h4, .main-header { color: #FFD700 !important; }
    
    /* Chỉ áp dụng nền tối khi ở chế độ Dark, nếu Light thì để Streamlit tự xử lý */
    [data-testid="stSidebar"] {
        border-right: 1px solid rgba(255, 215, 0, 0.2);
    }
    
    .main-header { 
        font-weight: 800; 
        font-size: 40px; 
        text-align: center; 
        border-bottom: 2px solid rgba(255, 215, 0, 0.3); 
        padding-bottom: 10px; 
        margin-bottom: 30px; 
    }
    .stButton>button { 
        background: linear-gradient(135deg, #152A4A 0%, #1e3a8a 100%) !important; 
        color: #FFD700 !important; 
        border: 1px solid #FFD700 !important; 
        border-radius: 12px; 
        font-weight: bold; 
        padding: 0.5rem 1rem;
        width: 100%; 
        box-shadow: 0 4px 15px rgba(255, 215, 0, 0.1);
    }
    .stButton>button:hover { 
        background: #FFD700 !important; 
        color: #0A1931 !important; 
        transform: scale(1.02); 
        transition: 0.2s; 
    }
    /* Dropdown color fix */
    .stSelectbox div[data-baseweb="select"] { border: 1px solid rgba(255, 215, 0, 0.2) !important; }
    
    .status-card {
        background: rgba(255, 215, 0, 0.05);
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 10px;
        border-left: 4px solid #FFD700;
    }
    .footer { text-align: center; color: #888; font-size: 12px; margin-top: 50px; border-top: 1px solid rgba(255, 215, 0, 0.2); padding-top: 20px; }
</style>
""", unsafe_allow_html=True)

# --- AUTH ---
def check_password():
    if "authenticated" not in st.session_state: st.session_state.authenticated = False
    if st.session_state.authenticated: return True
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h2 style='text-align: center; color: #FFD700; margin-top: 100px;'>🏦 TEXO STANDARDIZER AUTH</h2>", unsafe_allow_html=True)
        pwd = st.text_input("Mật khẩu truy cập hệ thống:", type="password")
        if st.button("XÁC THỰC"):
            if pwd == "texo2026":
                st.session_state.authenticated = True
                st.rerun()
            else: st.error("❌ Mật khẩu không chính xác.")
    return False

if not check_password(): st.stop()

# --- INITIALIZE STATE ---
if "standardized_files" not in st.session_state:
    st.session_state.standardized_files = {} # {filename: {"data": b"", "out_path": ""}}

# --- MAIN ---
st.markdown("<div class='main-header'>📄 CHUẨN HÓA VĂN BẢN TEXO</div>", unsafe_allow_html=True)

col1, col2 = st.columns([1, 1.2], gap="large")

with col1:
    st.markdown("### ⚙️ Cấu hình Chuẩn hóa")
    mode = st.selectbox("Bộ quy chuẩn Elite:", ["Quy định TEXO (Nâng cao)", "Nghị định 30/2020 (Cơ bản)"])
    
    # --- DYNAMIC SIDEBAR BASED ON MODE ---
    with st.sidebar:
        if "TEXO" in mode:
            st.markdown("### 🏆 QUY TẮC VÀNG TEXO")
            st.info("Hệ thống tự động áp dụng chuẩn độc quyền TEXO:")
            st.markdown("""
            - **Font:** Times New Roman
            - **Size:** 
                - Tiêu đề lớn (TỜ TRÌNH...): **14**
                - Đề mục In hoa khác: **12**
                - Nội dung thường: **13**
            - **Lề (Margins):** 
                - Thường: T20, B20, L30, R20
                - Letterhead: T40, B30, L30, R20
            - **Giãn dòng:** Exactly 17pt (Ngoài), 15pt (Trong bảng)
            - **Spacing:** Before 6pt, After 3pt
            - **Bảng biểu:** Indent lề 2mm (tránh dính cạnh)
            - **Tắt Contextual Spacing**
            """)
        else:
            st.markdown("### 📜 CHUẨN NGHỊ ĐỊNH 30")
            st.info("Tuân thủ tuyệt đối quy định của Chính phủ:")
            st.markdown("""
            - **Font:** Times New Roman
            - **Size:** Chuẩn **14** toàn văn bản
            - **Lề (Margins):** 
                - Trên/Dưới: 20-25mm
                - Trái: 30-35mm
                - Phải: 15-20mm
            - **Giãn dòng:** 1.15 - 1.5 Pt
            - **Căn lề:** Dóng đều 2 bên (Justified)
            """)
        
        st.divider()
        if st.button("♻️ LÀM MỚI DANH SÁCH"):
            st.session_state.standardized_files = {}
            st.rerun()
    
    is_letterhead = False
    if "TEXO" in mode:
        paper_type = st.radio("Loại phôi giấy:", ["Giấy trắng thường", "Giấy Letterhead"], horizontal=True)
        is_letterhead = True if "Letterhead" in paper_type else False

    uploaded_files = st.file_uploader("Tải hồ sơ (.docx) - Chọn nhiều file", type=["docx"], accept_multiple_files=True)

with col2:
    st.markdown("### 🚀 Thực thi & Tải về")
    if uploaded_files:
        if st.button("🚀 BẮT ĐẦU CHUẨN HÓA HÀNG LOẠT"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, f_up in enumerate(uploaded_files):
                status_text.markdown(f"⏳ Đang xử lý ({i+1}/{len(uploaded_files)}): **{f_up.name}**")
                try:
                    in_path = f"temp_{f_up.name}"
                    out_path = f"Standardized_{f_up.name}"
                    
                    with open(in_path, "wb") as f:
                        f.write(f_up.getbuffer())
                    
                    if "TEXO" in mode:
                        apply_texo_internal_standard(in_path, out_path, is_letterhead)
                    else:
                        apply_nd30_standard(in_path, out_path)
                    
                    with open(out_path, "rb") as fo:
                        data = fo.read()
                        st.session_state.standardized_files[f_up.name] = {"data": data, "out_path": out_path}
                    
                    # Cleanup
                    if os.path.exists(in_path): os.remove(in_path)
                    if os.path.exists(out_path): os.remove(out_path)
                except Exception as e:
                    st.error(f"Lỗi file {f_up.name}: {e}")
                
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            status_text.success(f"✅ Đã chuẩn hóa xong {len(uploaded_files)} file!")
            st.balloons()
            
    if st.session_state.standardized_files:
        # Bulk Download Zip
        if len(st.session_state.standardized_files) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for fname, fmeta in st.session_state.standardized_files.items():
                    zip_file.writestr(fmeta["out_path"], fmeta["data"])
            
            st.download_button(
                label="📥 TẢI XUỐNG TẤT CẢ (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="TEXO_Standardized_Docs.zip",
                mime="application/zip"
            )
            st.divider()

        # Individual Download
        for fname, fmeta in st.session_state.standardized_files.items():
            with st.container():
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"📄 **{fname}**")
                with c2:
                    st.download_button("📥 Tải về", fmeta["data"], fmeta["out_path"], key=f"dl_{fname}")

    else:
        if not uploaded_files:
            st.markdown("<div style='text-align: center; color: #64748b; font-weight: 500; padding: 100px 20px;'>Hệ thống đang chờ lệnh... <br>Vui lòng cấu hình và tải hồ sơ ở cột bên trái để bắt đầu.</div>", unsafe_allow_html=True)

st.markdown("<div class='footer'>TEXO Engineering Department | AI Master Standardizer | Hoàng Đức Vũ</div>", unsafe_allow_html=True)
