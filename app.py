import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os

# ===== DANH SÁCH TÀI KHOẢN — TỰ CẬP NHẬT TỪ EXCEL =====
DANH_SACH_TAI_KHOAN = [
    {"TaiKhoan": "admin", "MatKhau": "admin", "VaiTro": "Không", "DonVi": "ADMIN"},
    {"TaiKhoan": "tanlap", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Tân Lập"},
    {"TaiKhoan": "chiengson", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Chiềng Sơn"},
    {"TaiKhoan": "bacyen", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Bắc Yên"},
    {"TaiKhoan": "yenchau", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Yên Châu"},
    {"TaiKhoan": "chienghac", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Chiềng Hặc"},
    {"TaiKhoan": "phiengkhoai", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Phiêng Khoài"},
    {"TaiKhoan": "phiengcam", "MatKhau": "123456", "VaiTro": "Không", "DonVi": "Xã Phiêng Cằm"}
]

# ===== CẤU HÌNH KẾT NỐI =====
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
    duong_dan = "templates/Form_Mau_Nuoc_Sach_Chuan.xlsx"
    if os.path.exists(duong_dan):
        with open(duong_dan, "rb") as f:
            return f.read()
    return None

# ===== LƯU DỮ LIỆU VÀO GOOGLE SHEETS =====
def luu_vao_sheets(ten_donvi, df):
    sh = ket_noi_gsheets()
    if not sh:
        return False, "Không kết nối được Google Sheets"
    
    df = df.copy()
    for cot in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[cot]):
            df[cot] = df[cot].dt.strftime("%Y-%m-%d")
        else:
            df[cot] = df[cot].apply(
                lambda x: x.strftime("%Y-%m-%d") if hasattr(x, "strftime") else x
            )
    
    try:
        ws = sh.worksheet(ten_donvi)
        ws.clear()
    except:
        ws = sh.add_worksheet(title=ten_donvi, rows=max(100, len(df)+5), cols=30)
    
    data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    ws.update(data)
    return True, "Đã lưu thành công"

# ===== GIAO DIỆN CHÍNH =====
st.set_page_config(page_title="Báo cáo Nước sạch Sơn La", layout="wide")
st.title("🏞️ HỆ THỐNG BÁO CÁO NƯỚC SẠCH TỈNH SƠN LA")

if "nguoi_dung" not in st.session_state:
    st.session_state.nguoi_dung = None

if st.session_state.nguoi_dung is None:
    st.subheader("🔐 Đăng nhập hệ thống")
    tk_dangnhap = st.text_input("Tên tài khoản").strip()
    mk_dangnhap = st.text_input("Mật khẩu", type="password").strip()
    
    if st.button("Đăng nhập"):
        nd = kiem_tra_dang_nhap(tk_dangnhap, mk_dangnhap)
        if nd:
            st.session_state.nguoi_dung = nd
            st.rerun()
        else:
            st.error("❌ Sai tài khoản hoặc mật khẩu!")
else:
    nd = st.session_state.nguoi_dung
    st.success(f"✅ Xin chào: {nd['TaiKhoan']} — {nd['DonVi']} ({nd['VaiTro']})")
    
    tab1, tab2 = st.tabs(["📤 Nộp báo cáo", "📊 Tổng hợp"])
    
    with tab1:
        mau = tai_file_mau()
        if mau:
            st.download_button(
                "📥 Tải file mẫu chuẩn",
                data=mau,
                file_name=f"BaoCao_{nd['DonVi']}.xlsx"
            )
        
        st.markdown("---")
        file_up = st.file_uploader("Chọn file đã điền dữ liệu", type=["xlsx"])
        
        if file_up:
            try:
                df_data = pd.read_excel(file_up)
                st.dataframe(df_data, use_container_width=True)
                
                if st.button("✅ Nộp báo cáo"):
                    ok, msg = luu_vao_sheets(nd["DonVi"], df_data)
                    if ok:
                        st.success(f"✅ {msg}")
                    else:
                        st.error(f"❌ {msg}")
            except Exception as e:
                st.error(f"❌ Lỗi đọc file: {str(e)}")
    
    with tab2:
        if nd["VaiTro"] == "Quản trị":
            link_master = f"https://docs.google.com/spreadsheets/d/{st.secrets['google_sheets_master_id']}/edit"
            st.link_button("📋 Mở Bảng tổng hợp Master", link_master)
        else:
            st.info("🔒 Chỉ quản trị viên xem được tổng hợp toàn tỉnh")
    
    if st.button("🚪 Đăng xuất"):
        st.session_state.nguoi_dung = None
        st.rerun()
