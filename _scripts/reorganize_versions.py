# -*- coding: utf-8 -*-
"""
reorganize_versions.py
Tổ chức lại toàn bộ cây thư mục dự án theo cấu trúc chuẩn:
- 00_Tai_Lieu_Va_Mau_Bieu_Chung/
- 01_Phien_Ban_v1_Goc/
- 02_Phien_Ban_v2/
- 03_Phien_Ban_v3/
- 04_Phien_Ban_v4/
- 05_Phien_Ban_v5/
- _scripts/
- README.md
"""
import os, sys, shutil

sys.stdout.reconfigure(encoding='utf-8')
root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

print(f"Bắt đầu sắp xếp thư mục tại: {root_dir}")

folders = [
    "00_Tai_Lieu_Va_Mau_Bieu_Chung",
    "01_Phien_Ban_v1_Goc",
    "02_Phien_Ban_v2",
    "03_Phien_Ban_v3",
    "04_Phien_Ban_v4",
    "05_Phien_Ban_v5",
    "_scripts"
]

for f in folders:
    fpath = os.path.join(root_dir, f)
    os.makedirs(fpath, exist_ok=True)

# 1. Di chuyển scripts từ workingfile/_scripts ra _scripts ở root
src_scripts = os.path.join(root_dir, "workingfile", "_scripts")
dst_scripts = os.path.join(root_dir, "_scripts")
if os.path.exists(src_scripts):
    for item in os.listdir(src_scripts):
        s_item = os.path.join(src_scripts, item)
        d_item = os.path.join(dst_scripts, item)
        if os.path.isdir(s_item):
            if not os.path.exists(d_item):
                shutil.copytree(s_item, d_item)
        else:
            shutil.copy2(s_item, d_item)
    print("Đã sao chép _scripts sang root/_scripts.")

# Các file python lẻ ở root cũng chuyển vào _scripts/
root_py_files = [
    "audit_from_raw_excel.py",
    "generate_perfect_realistic_dataset.py",
    "generate_raw_dataset_32items.py"
]
for rf in root_py_files:
    src_rf = os.path.join(root_dir, rf)
    if os.path.exists(src_rf):
        shutil.move(src_rf, os.path.join(dst_scripts, rf))
        print(f"Chuyển {rf} vào _scripts/")

# Danh mục phân bổ file vào từng folder phiên bản
file_mappings = {
    "00_Tai_Lieu_Va_Mau_Bieu_Chung": [
        "Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx",
        "Phieu_Khao_Sat_Chinh_Thuc_VietinBank.md",
        "DTL_Master_Thesis_Full_Structure_Template.docx",
        "DTL_Master_Thesis_Full_Structure_Template.md",
        "DTL_Thesis_Design_NEU_MDE_Final.docx",
        "DTL_Thesis_Design_NEU_MDE_Final.md",
        "Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.docx",
        "Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.md",
        "Bien_Ban_Hoi_Thoai_Va_Phat_Trien_De_Tai.md"
    ],
    "01_Phien_Ban_v1_Goc": [
        "DTL_Master_Thesis_Draft.docx",
        "DTL_Master_Thesis_Draft.md",
        "Du_Lieu_Khao_Sat_Tho_800_DN.xlsx",
        "Du_Lieu_Khao_Sat_Tho_800_DN.csv",
        "Codebook_Bien_Va_Quy_Trinh_Xu_Ly.xlsx",
        "figure_4_1_empirical_model.png"
    ],
    "02_Phien_Ban_v2": [
        "DTL_Master_Thesis_Draft_v2.docx",
        "DTL_Master_Thesis_Draft_v2.md",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v2.csv",
        "Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v2.xlsx",
        "figure_4_1_empirical_model_v2.png"
    ],
    "03_Phien_Ban_v3": [
        "DTL_Master_Thesis_Draft_v3.docx",
        "DTL_Master_Thesis_Draft_v3.md",
        "Du_Lieu_Khao_Sat_Goc_842_DN_v3.xlsx",
        "Du_Lieu_Khao_Sat_Goc_842_DN_v3.csv",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v3.xlsx",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v3.csv",
        "Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v3.xlsx",
        "figure_4_1_empirical_model_v3.png",
        "v3_empirical_results.json"
    ],
    "04_Phien_Ban_v4": [
        "DTL_Master_Thesis_Draft_v4.docx",
        "DTL_Master_Thesis_Draft_v4.md",
        "Du_Lieu_Khao_Sat_Goc_842_DN_v4.xlsx",
        "Du_Lieu_Khao_Sat_Goc_842_DN_v4.csv",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v4.xlsx",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v4.csv",
        "Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v4.xlsx",
        "figure_4_1_empirical_model_v4.png",
        "v4_empirical_results.json"
    ],
    "05_Phien_Ban_v5": [
        "DTL_Master_Thesis_Draft_v5.docx",
        "DTL_Master_Thesis_Draft_v5.md",
        "Du_Lieu_Khao_Sat_Goc_865_DN_v5.xlsx",
        "Du_Lieu_Khao_Sat_Goc_865_DN_v5.csv",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v5.xlsx",
        "Du_Lieu_Khao_Sat_Tho_800_DN_v5.csv",
        "Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v5.xlsx",
        "figure_4_1_empirical_model_v5.png",
        "v5_empirical_results.json"
    ]
}

# Sao chép file từ root hoặc workingfile vào thư mục chỉ định
for folder_name, file_list in file_mappings.items():
    target_folder = os.path.join(root_dir, folder_name)
    for fname in file_list:
        # Tìm file ở root hoặc ở workingfile
        src_root = os.path.join(root_dir, fname)
        src_wf = os.path.join(root_dir, "workingfile", fname)
        
        found_src = None
        if os.path.exists(src_root):
            found_src = src_root
        elif os.path.exists(src_wf):
            found_src = src_wf
            
        if found_src:
            dest_file = os.path.join(target_folder, fname)
            shutil.copy2(found_src, dest_file)
            print(f"  -> Đã đưa [{fname}] vào [{folder_name}/]")
        else:
            print(f"  [CẢNH BÁO] Không tìm thấy file: {fname}")

print("\nHoàn tất sao chép an toàn vào các thư mục phiên bản.")
