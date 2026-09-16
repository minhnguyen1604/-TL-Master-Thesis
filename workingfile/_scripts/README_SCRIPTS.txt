CAC SCRIPT SINH FILE - DOC KY TRUOC KHI CHAY
=============================================

CHAY TU THU MUC CHA (workingfile), KHONG CHAY TRONG THU MUC NAY:
    python3 _scripts/create_survey_docx.py

-------------------------------------------------------------
NHOM AN TOAN - chay lai bao nhieu lan cung duoc (ghi de file cu)
-------------------------------------------------------------
  survey_items.py ............... NGUON GOC bo cau hoi. Sua o day roi chay
                                  2 script duoi thi phieu va codebook tu khop nhau.
  create_survey_docx.py ......... sinh Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx
  create_codebook.py ............ sinh Codebook_Bien_Va_Quy_Trinh_Xu_Ly.xlsx
  generate_final_mde_thesis_design.py ....... sinh DTL_Thesis_Design_NEU_MDE_Final.docx
  generate_advisor_feedback_report_docx.py .. sinh Bao_Cao_Tong_Hop_Gop_Y...docx
  generate_full_thesis_structure_template_docx.py ... sinh file KHUNG (khong phai bai viet)

-------------------------------------------------------------
!! NHOM NGUY HIEM - TUYET DOI KHONG CHAY LAI !!
-------------------------------------------------------------
  write_chapter1.py       write_chapter2.py       write_chapter3.py
  write_frontmatter.py    write_limitations.py
  add_citations.py        update_draft_32items.py

  Cac script nay CHEN THEM noi dung vao DTL_Master_Thesis_Draft.docx,
  KHONG ghi de. Chay lai lan thu hai = toan bo Chuong 1, 2, 3 bi NHAN DOI
  trong bai. Chung chi duoc chay dung mot lan va da chay xong roi.

  Giu lai chi de tra cuu noi dung goc neu can, khong phai de chay.
