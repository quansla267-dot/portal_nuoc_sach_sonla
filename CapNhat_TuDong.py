import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import subprocess
from datetime import datetime
import pandas as pd

class CapNhatTaiKhoanApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CẬP NHẬT TÀI KHOẢN & ĐẨY LÊN GITHUB — Nước sạch Sơn La")
        self.root.geometry("650x620")
        self.root.resizable(True, True)
        
        # === Thư mục gốc = thư mục chứa file CapNhat_TuDong.py ===
        self.thu_muc_goc = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.thu_muc_goc)
        
        # Thư mục lưu file mẫu
        self.thu_muc_mau = os.path.join(self.thu_muc_goc, "templates")
        os.makedirs(self.thu_muc_mau, exist_ok=True)
        
        # === Đọc tên hệ thống từ ô A1 file Excel mẫu ===
        self.ten_he_thong = "HỆ THỐNG BÁO CÁO NƯỚC SẠCH TỈNH SƠN LA"
        self.doc_ten_tu_file_mau()
        
        # Mẫu code app.py — ĐÃ CẬP NHẬT TẤT CẢ YÊU CẦU
        self.mau_code_app = '''import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os

# ===== DANH SÁCH TÀI KHOẢN — TỰ CẬP NHẬT TỪ EXCEL =====
DANH_SACH_TAI_KHOAN = [
{DANH_SACH_TAI_KHOAN}
]

# ===== CẤU HÌNH GIAO DIỆN — NỀN TỐI MẶC ĐỊNH =====
st.set_page_config(
    page_title="{TEN_HE_THONG}",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={}
)

# === BUỘC NỀN TỐI ===
st.markdown("""
<style>
    .stApp {{
        background-color: #0E1117;
        color: #FAFAFA;
    }}
    div[data-testid="stForm"] {{
        background-color: #262730;
    }}
    .stButton>button {{
        width: 100%;
    }}
</style>
""", unsafe_allow_html=True)

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
    for ten_file in os.listdir("templates"):
        if ten_file.lower().endswith((".xlsx", ".xls")):
            duong_dan = os.path.join("templates", ten_file)
            with open(duong_dan, "rb") as f:
                return f.read(), ten_file
    return None, None

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
st.title("🏞️ {TEN_HE_THONG}")

if "nguoi_dung" not in st.session_state:
    st.session_state.nguoi_dung = None

if st.session_state.nguoi_dung is None:
    st.subheader("🔐 Đăng nhập hệ thống")
    
    # === NHẬP ENTER = ĐĂNG NHẬP LUÔN ===
    with st.form("dangnhap_form"):
        tk_dangnhap = st.text_input("Tên tài khoản").strip()
        mk_dangnhap = st.text_input("Mật khẩu", type="password").strip()
        gui_nhan = st.form_submit_button("Đăng nhập")  # Enter tự động kích hoạt
    
    if gui_nhan:
        nd = kiem_tra_dang_nhap(tk_dangnhap, mk_dangnhap)
        if nd:
            st.session_state.nguoi_dung = nd
            st.rerun()
        else:
            st.error("❌ Sai tài khoản hoặc mật khẩu!")
else:
    nd = st.session_state.nguoi_dung
    
    # === SỬA CHỮ CHÀO ===
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
            st.warning("⚠️ Chưa có file mẫu — Quản trị vui lòng cập nhật!")
        
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
        if nd["VaiTro"] in ["Quản trị", "Admin", "quản trị"]:
            link_master = f"https://docs.google.com/spreadsheets/d/{st.secrets['google_sheets_master_id']}/edit"
            st.link_button("📋 Mở Bảng tổng hợp Master", link_master)
        else:
            st.info("🔒 Chỉ quản trị viên xem được tổng hợp toàn tỉnh")
    
    if st.button("🚪 Đăng xuất"):
        st.session_state.nguoi_dung = None
        st.rerun()
'''.replace("{TEN_HE_THONG}", self.ten_he_thong)
        self.tao_giao_dien()

    def doc_ten_tu_file_mau(self):
        """Đọc tên hệ thống từ ô A1 file Excel mẫu"""
        for ten_file in os.listdir(self.thu_muc_goc):
            if ten_file.lower().endswith((".xlsx", ".xls")) and "Mau" in ten_file or "mau" in ten_file:
                try:
                    duong_dan = os.path.join(self.thu_muc_goc, ten_file)
                    df = pd.read_excel(duong_dan, header=None)
                    ten_moi = str(df.iloc[0, 0]).strip()
                    if ten_moi:
                        self.ten_he_thong = ten_moi
                        self.ghi(f"📌 Đọc tên hệ thống từ ô A1: {ten_moi}")
                    return
                except:
                    pass

    def tao_giao_dien(self):
        tk.Label(self.root, text="🔄 CẬP NHẬT TÀI KHOẢN & ĐẨY LÊN GITHUB", 
                 font=("Arial", 16, "bold"), fg="#2c3e50").pack(pady=12)
        
        khung = tk.Frame(self.root)
        khung.pack(padx=30, pady=5, fill="both", expand=True)

        # === 1. File Excel tài khoản — MỞ TỪ THƯ MỤC CHƯƠNG TRÌNH ===
        khung1 = tk.LabelFrame(khung, text="📋 File Excel — Danh sách tài khoản", padx=10, pady=8)
        khung1.pack(fill="x", pady=5)
        self.duong_dan_excel_tk = tk.StringVar()
        tk.Entry(khung1, textvariable=self.duong_dan_excel_tk, width=55).pack(side="left", padx=5)
        tk.Button(khung1, text="Chọn File", command=self.chon_excel_tai_khoan,
                  bg="#27ae60", fg="white", padx=15).pack(side="right")

        # 2. File Excel mẫu dữ liệu
        khung2 = tk.LabelFrame(khung, text="📂 File Excel — Mẫu/Biểu mẫu báo cáo", padx=10, pady=8)
        khung2.pack(fill="x", pady=5)
        self.duong_dan_file_mau = tk.StringVar()
        tk.Entry(khung2, textvariable=self.duong_dan_file_mau, width=55).pack(side="left", padx=5)
        tk.Button(khung2, text="Chọn File", command=self.chon_file_mau,
                  bg="#3498db", fg="white", padx=15).pack(side="right")

        # 3. Ghi chú
        khung3 = tk.LabelFrame(khung, text="📝 Ghi chú nội dung thay đổi", padx=10, pady=8)
        khung3.pack(fill="x", pady=5)
        self.ghi_chu = tk.StringVar(value=f"Cập nhật tài khoản: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        tk.Entry(khung3, textvariable=self.ghi_chu, width=65).pack(fill="x")

        # Nút thực hiện
        tk.Button(self.root, text="🚀 TẠO app.py MỚI & ĐẨY LÊN GITHUB", 
                  command=self.thuc_hien, bg="#e74c3c", fg="white", 
                  font=("Arial", 13, "bold"), padx=40, pady=12).pack(pady=15)

        # Kết quả
        self.ket_qua = tk.Text(self.root, height=12, wrap="word", font=("Consolas", 9))
        self.ket_qua.pack(padx=20, pady=(0,15), fill="both", expand=True)
        self.ket_qua.config(state="disabled")

    def ghi(self, txt):
        self.ket_qua.config(state="normal")
        self.ket_qua.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {txt}\n")
        self.ket_qua.see("end")
        self.ket_qua.config(state="disabled")
        self.root.update()

    def chon_excel_tai_khoan(self):
        """Mở mặc định tại thư mục chứa file CapNhat_TuDong.py"""
        f = filedialog.askopenfilename(
            title="Chọn file Excel danh sách tài khoản",
            initialdir=self.thu_muc_goc,
            filetypes=[("Excel", "*.xlsx"), ("Excel 97-2003", "*.xls")]
        )
        if f: 
            self.duong_dan_excel_tk.set(f)
            self.ghi(f"✅ Đã chọn: {os.path.basename(f)}")

    def chon_file_mau(self):
        f = filedialog.askopenfilename(
            title="Chọn file Excel mẫu/biểu mẫu",
            filetypes=[("Excel", "*.xlsx"), ("Excel 97-2003", "*.xls")]
        )
        if f:
            self.duong_dan_file_mau.set(f)
            self.ghi(f"✅ Đã chọn file mẫu: {os.path.basename(f)}")
            # Đọc lại tên hệ thống từ file mẫu vừa chọn
            try:
                df = pd.read_excel(f, header=None)
                ten_moi = str(df.iloc[0, 0]).strip()
                if ten_moi:
                    self.ten_he_thong = ten_moi
                    self.ghi(f"📌 Tên hệ thống: {ten_moi}")
            except:
                pass

    def doc_danh_sach_tu_excel(self, duong_dan):
        try:
            df = pd.read_excel(duong_dan, dtype=str)
            anh_xa = {
                "TaiKhoan": "TaiKhoan", "Username": "TaiKhoan", "Tên tài khoản": "TaiKhoan",
                "MatKhau": "MatKhau", "Mat_Khau": "MatKhau", "Mật khẩu": "MatKhau",
                "VaiTro": "VaiTro", "Khoa": "VaiTro", "Vai trò": "VaiTro",
                "DonVi": "DonVi", "Ten_Xa_Phuong": "DonVi", "Đơn vị": "DonVi"
            }
            df = df.rename(columns={c: anh_xa.get(c.strip(), c.strip()) for c in df.columns})
            for cot in ["TaiKhoan", "MatKhau", "VaiTro", "DonVi"]:
                if cot not in df.columns:
                    df[cot] = ""
            df = df.fillna("")
            dong = []
            for _, row in df.iterrows():
                dong.append(f'    {{"TaiKhoan": "{row["TaiKhoan"]}", "MatKhau": "{row["MatKhau"]}", "VaiTro": "{row["VaiTro"]}", "DonVi": "{row["DonVi"]}"}}')
            return ",\n".join(dong) if dong else ""
        except Exception as e:
            messagebox.showerror("Lỗi đọc Excel", f"Không đọc được file:\n{str(e)}")
            return None

    def tao_app_py_moi(self, noi_dung_danh_sach):
        code = self.mau_code_app.replace("{DANH_SACH_TAI_KHOAN}", noi_dung_danh_sach)
        with open("app.py", "w", encoding="utf-8") as f:
            f.write(code)
        return True

    def kiem_tra_git(self):
        try:
            subprocess.run(["git", "--version"], check=True, capture_output=True, text=True)
            return True
        except:
            return False

    def thuc_hien(self):
        self.ghi("="*50)
        
        if not self.duong_dan_excel_tk.get().strip():
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file Excel tài khoản!")
            return
        
        if not self.kiem_tra_git():
            messagebox.showerror("Thiếu Git", "Chưa cài Git!\nTải tại: git-scm.com/download/win")
            return

        # Đọc Excel tài khoản
        self.ghi("📖 Đọc danh sách tài khoản...")
        ds = self.doc_danh_sach_tu_excel(self.duong_dan_excel_tk.get().strip())
        if ds is None: return
        self.ghi(f"✅ Đã đọc được {ds.count('TaiKhoan')} tài khoản")

        # Tạo app.py
        self.ghi("⚙️ Tạo file app.py...")
        self.tao_app_py_moi(ds)
        self.ghi("✅ Đã cập nhật app.py")

        # === SAO CHÉP FILE MẪU ===
        duong_dan_mau = self.duong_dan_file_mau.get().strip()
        if duong_dan_mau:
            self.ghi(f"🔍 Kiểm tra file mẫu: {os.path.basename(duong_dan_mau)}")
            if os.path.isfile(duong_dan_mau):
                try:
                    ten_file = os.path.basename(duong_dan_mau)
                    # Xóa file mẫu cũ
                    for f in os.listdir(self.thu_muc_mau):
                        if f.lower().endswith((".xlsx", ".xls")):
                            os.remove(os.path.join(self.thu_muc_mau, f))
                            self.ghi(f"🗑️  Xóa file cũ: {f}")
                    # Sao chép file mới
                    dich = os.path.join(self.thu_muc_mau, ten_file)
                    shutil.copy2(duong_dan_mau, dich)
                    self.ghi(f"✅ Đã cập nhật file mẫu: {ten_file}")
                except Exception as e:
                    self.ghi(f"❌ Lỗi sao chép: {str(e)}")
            else:
                self.ghi(f"❌ Không tìm thấy file")
        else:
            self.ghi("ℹ️ Bỏ qua file mẫu")

        # Khởi tạo Git nếu chưa có
        if not os.path.exists(".git"):
            self.ghi("🔧 Khởi tạo kết nối Git...")
            subprocess.run(["git", "init"], capture_output=True)
            subprocess.run(["git", "remote", "add", "origin",
                "https://github.com/quansla267-dot/portal_nuoc_sach_sonla.git"], capture_output=True)
            subprocess.run(["git", "branch", "-M", "main"], capture_output=True)

        # Commit & Đẩy lên GitHub
        subprocess.run(["git", "add", "."], capture_output=True)
        ghichu = self.ghi_chu.get().strip() or f"Cập nhật {datetime.now().strftime('%d/%m/%Y')}"
        kq_commit = subprocess.run(["git", "commit", "-m", ghichu], capture_output=True, text=True)
        
        if kq_commit.returncode != 0:
            if "nothing to commit" in kq_commit.stdout + kq_commit.stderr:
                self.ghi("ℹ️ Không có thay đổi mới để đẩy")
            else:
                self.ghi(f"❌ Lỗi Commit: {kq_commit.stderr}")
                return
        else:
            self.ghi("📤 Đang đẩy lên GitHub...")
            kq_push = subprocess.run(["git", "push", "-u", "origin", "main"], capture_output=True, text=True)
            
            if kq_push.returncode != 0:
                subprocess.run(["git", "pull", "origin", "main", "--no-edit"], capture_output=True)
                subprocess.run(["git", "add", "."], capture_output=True)
                subprocess.run(["git", "commit", "-m", "Đồng bộ"], capture_output=True, text=True)
                kq_push = subprocess.run(["git", "push", "--force", "-u", "origin", "main"], capture_output=True, text=True)
            
            if kq_push.returncode == 0:
                self.ghi("")
                self.ghi("🎉 THÀNH CÔNG! Đã đẩy lên GitHub ✅")
                messagebox.showinfo("Hoàn thành", "✅ Cập nhật xong!\nTải lại Streamlit để xem giao diện mới")
            else:
                self.ghi(f"❌ Lỗi đẩy: {kq_push.stderr}")
                messagebox.showerror("Lỗi", "Chạy thủ công: git push --force -u origin main")

if __name__ == "__main__":
    root = tk.Tk()
    app = CapNhatTaiKhoanApp(root)
    root.mainloop()