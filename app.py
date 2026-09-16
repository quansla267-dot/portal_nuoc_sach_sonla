import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os

# ===== DANH SÁCH TÀI KHOẢN =====
DANH_SACH_TAI_KHOAN = [
    {"TaiKhoan": "admin", "MatKhau": "", "VaiTro": "Quản trị", "DonVi": ""},
    {"TaiKhoan": "tanlap", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""},
    {"TaiKhoan": "chiengson", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""},
    {"TaiKhoan": "bacyen", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""},
    {"TaiKhoan": "yenchau", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""},
    {"TaiKhoan": "chienghac", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""},
    {"TaiKhoan": "phiengkhoai", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""},
    {"TaiKhoan": "phiengcam", "MatKhau": "", "VaiTro": "Tài khoản xã/phường", "DonVi": ""}
]

# ===== NỀN TỐI MẶC ĐỊNH =====
st.set_page_config(
    page_title="HỆ THỐNG BÁO CÁO NƯỚC SẠCH TỈNH SƠN LA",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stApp {{
        background-color: #0E1117;
        color: #FAFAFA;
    }}
    div[data-testid="stForm"] {{
        background-color: #262730;
    }}
</style>
""", unsafe_allow_html=True)

# ===== KẾT NỐI GOOGLE SHEETS =====
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource(show_spinner=False)
def ket_noi_gsheets():
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
        client = gspread.authorize(creds)
        return client.open_by_key(st.secrets["google_sheets_master_id"])
    except Exception as e:
        st.error(f"Lỗi kết nối Google: {str(e)}")
        return None

# ===== KIỂM TRA ĐĂNG NHẬP =====
def kiem_tra_dang_nhap(tk, mk):
    for u in DANH_SACH_TAI_KHOAN:
        if u["TaiKhoan"].strip() == tk.strip() and u["MatKhau"].strip() == mk.strip():
            return u
    return None

# ===== TẢI FILE MẪU =====
def tai_file_mau():
    for ten_file in os.listdir("templates"):
        if ten_file.lower().endswith((".xlsx", ".xls")):
            with open(os.path.join("templates", ten_file), "rb") as f:
                return f.read(), ten_file
    return None, None

# ===== LƯU DỮ LIỆU =====
def luu_vao_sheets(ten_donvi, df):
    sh = ket_noi_gsheets()
    if not sh:
        return False, "Không kết nối được Google Sheets"
    
    df = df.copy()
    for cot in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[cot]):
            df[cot] = df[cot].dt.strftime("%Y-%m-%d")
    
    try:
        ws = sh.worksheet(ten_donvi)
        ws.clear()
    except:
        ws = sh.add_worksheet(title=ten_donvi, rows=max(100, len(df)+5), cols=30)
    
    data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    ws.update(data)
    return True, "Đã lưu thành công"

# ===== GIAO DIỆN CHÍNH =====
st.title("🏞️ HỆ THỐNG BÁO CÁO NƯỚC SẠCH TỈNH SƠN LA")

if "nguoi_dung" not in st.session_state:
    st.session_state.nguoi_dung = None

if st.session_state.nguoi_dung is None:
    st.subheader("🔐 Đăng nhập hệ thống")
    # === NHẬP ENTER = ĐĂNG NHẬP ===
    with st.form("dangnhap_form"):
        tk_dangnhap = st.text_input("Tên tài khoản")
        mk_dangnhap = st.text_input("Mật khẩu", type="password")
        gui_nhan = st.form_submit_button("Đăng nhập")
    
    if gui_nhan:
        nd = kiem_tra_dang_nhap(tk_dangnhap, mk_dangnhap)
        if nd:
            st.session_state.nguoi_dung = nd
            st.rerun()
        else:
            st.error("❌ Sai tài khoản hoặc mật khẩu!")
else:
    nd = st.session_state.nguoi_dung
    # === CHỮ CHÀO ĐÃ SỬA ===
    st.success(f"✅ Chào mừng bạn đã đăng nhập: {nd['TaiKhoan']} — {nd['DonVi']} ({nd['VaiTro']})")
    
    tab1, tab2 = st.tabs(["📤 Nộp báo cáo", "📊 Tổng hợp"])
    
    with tab1:
        mau_data, ten_mau = tai_file_mau()
        if mau_data:
            st.download_button(
                f"📥 Tải file mẫu: {ten_mau}",
                data=mau_data,
                file_name=f"BaoCao_{nd['DonVi']}.xlsx"
            )
        else:
            st.warning("⚠️ Chưa có file mẫu!")
        
        st.markdown("---")
        file_up = st.file_uploader("Chọn file đã điền dữ liệu", type=["xlsx"])
        
        if file_up:
            try:
                df_data = pd.read_excel(file_up)
                st.dataframe(df_data, use_container_width=True)
                if st.button("✅ Nộp báo cáo"):
                    ok, msg = luu_vao_sheets(nd["DonVi"], df_data)
                    st.success(f"✅ {msg}" if ok else f"❌ {msg}")
            except Exception as e:
                st.error(f"❌ Lỗi đọc file: {str(e)}")
    
    with tab2:
        if nd["VaiTro"] in ["Quản trị", "Admin", "quản trị", "admin"]:
            link = f"https://docs.google.com/spreadsheets/d/{st.secrets['google_sheets_master_id']}/edit"
            st.link_button("📋 Mở Bảng tổng hợp Master", link)
        else:
            st.info("🔒 Chỉ quản trị viên xem được tổng hợp toàn tỉnh")
    
    if st.button("🚪 Đăng xuất"):
        st.session_state.nguoi_dung = None
        st.rerun()
