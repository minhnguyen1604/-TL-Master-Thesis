import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def create_survey_docx():
    doc = docx.Document()
    
    # Page Margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Base Normal Style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("PHIẾU KHẢO SÁT THU THẬP DỮ LIỆU NGHIÊN CỨU\n")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(0x00, 0x33, 0x66) # VietinBank Navy Blue

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Đề tài: \"Các yếu tố ảnh hưởng đến quyết định lựa chọn sản phẩm bảo lãnh ngân hàng của các doanh nghiệp tại VietinBank\"\n")
    r_sub.italic = True
    r_sub.bold = True
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(0x00, 0x55, 0x99)

    # Thư ngỏ
    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.line_spacing = 1.15
    r_intro = p_intro.add_run(
        "Kính gửi Quý Doanh nghiệp!\n"
        "Nhằm nâng cao chất lượng dịch vụ và tối ưu hóa giải pháp bảo lãnh ngân hàng đáp ứng tốt nhất nhu cầu của cộng đồng doanh nghiệp, "
        "nghiên cứu này được thực hiện để đánh giá các yếu tố tác động đến quyết định lựa chọn dịch vụ bảo lãnh tại VietinBank.\n"
        "Sự hợp tác và những đánh giá khách quan của Quý Doanh nghiệp là đóng góp vô cùng quý báu cho thành công của nghiên cứu này. "
        "Mọi thông tin Quý Doanh nghiệp cung cấp hoàn toàn được BẢO MẬT tuyệt đối và chỉ phục vụ cho mục đích nghiên cứu học thuật.\n"
        "Xin chân thành cảm ơn sự đồng hành của Quý Doanh nghiệp!"
    )
    r_intro.font.size = Pt(11)
    r_intro.italic = True

    doc.add_paragraph()

    # PHẦN I
    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("PHẦN I: THÔNG TIN CHUNG VỀ DOANH NGHIỆP (DEMOGRAPHICS)")
    r_h1.bold = True
    r_h1.font.size = Pt(13)
    r_h1.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_demo = doc.add_paragraph()
    p_demo.paragraph_format.line_spacing = 1.25
    p_demo.add_run("Quý Doanh nghiệp vui lòng đánh dấu (X) vào ô tròn tương ứng với thông tin của đơn vị mình:\n\n")
    p_demo.add_run("1. Loại hình doanh nghiệp:\n")
    p_demo.add_run("   ( ) Doanh nghiệp TNHH               ( ) Doanh nghiệp Cổ phần\n")
    p_demo.add_run("   ( ) Doanh nghiệp FDI (Vốn nước ngoài)   ( ) Doanh nghiệp Nhà nước / Khác\n\n")

    p_demo.add_run("2. Thâm niên hoạt động của doanh nghiệp:\n")
    p_demo.add_run("   ( ) Dưới 5 năm                      ( ) Từ 5 đến 10 năm                 ( ) Trên 10 năm\n\n")

    p_demo.add_run("3. Quy mô doanh thu trung bình hàng năm:\n")
    p_demo.add_run("   ( ) Dưới 50 tỷ VNĐ                  ( ) Từ 50 đến 200 tỷ VNĐ            ( ) Trên 200 tỷ VNĐ\n\n")

    p_demo.add_run("4. Loại hình bảo lãnh doanh nghiệp thường xuyên sử dụng tại VietinBank (Có thể chọn nhiều ô):\n")
    p_demo.add_run("   [ ] Bảo lãnh dự thầu                [ ] Bảo lãnh thực hiện hợp đồng     [ ] Bảo lãnh tạm ứng\n")
    p_demo.add_run("   [ ] Bảo lãnh thanh toán             [ ] Bảo lãnh bảo hành / Khác\n\n")

    p_demo.add_run("5. Vị trí/Chức danh của người đại diện thực hiện khảo sát:\n")
    p_demo.add_run("   ( ) Giám đốc / Tổng Giám đốc        ( ) Giám đốc Tài chính (CFO)       ( ) Kế toán trưởng\n")
    p_demo.add_run("   ( ) Trưởng phòng Kế toán / Thầu     ( ) Chuyên viên phụ trách bảo lãnh / Khác\n")

    doc.add_paragraph()

    # PHẦN II
    h2 = doc.add_paragraph()
    r_h2 = h2.add_run("PHẦN II: NỘI DUNG ĐÁNH GIÁ CÁC YẾU TỐ VÀ QUYẾT ĐỊNH LỰA CHỌN")
    r_h2.bold = True
    r_h2.font.size = Pt(13)
    r_h2.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_scale = doc.add_paragraph()
    p_scale.paragraph_format.line_spacing = 1.15
    r_scale_intro = p_scale.add_run("Xin Quý Doanh nghiệp vui lòng cho biết mức độ đồng ý đối với các phát biểu dưới đây theo thang đo 5 cấp độ:\n")
    r_scale_intro.italic = True
    p_scale.add_run("   1: Hoàn toàn không đồng ý   |   2: Không đồng ý   |   3: Phân vân / Trung lập   |   4: Đồng ý   |   5: Hoàn toàn đồng ý")

    doc.add_paragraph()

    # TABLE DEFINITION
    items_data = [
        # (Header/Section Name, items)
        ("A. ĐÁNH GIÁ CÁC BIẾN ĐỘC LẬP (YẾU TỐ TÁC ĐỘNG)", []),
        
        ("I. Chi phí & Biểu phí bảo lãnh (COST)", [
            ("COST1", "Mức phí phát hành bảo lãnh tại VietinBank là hợp lý so với chất lượng dịch vụ nhận được."),
            ("COST2", "Biểu phí bảo lãnh của VietinBank có tính cạnh tranh cao so với các ngân hàng thương mại khác."),
            ("COST3", "VietinBank có chính sách ưu đãi và giảm phí bảo lãnh linh hoạt cho các khách hàng thường xuyên.")
        ]),
        
        ("II. Yêu cầu Tài sản đảm bảo & Tỷ lệ Ký quỹ (COL)", [
            ("COL1", "VietinBank áp dụng tỷ lệ ký quỹ bảo lãnh linh hoạt (có chính sách giảm hoặc miễn ký quỹ)."),
            ("COL2", "VietinBank chấp nhận đa dạng các loại tài sản đảm bảo (BĐS, máy móc, dòng tiền hợp đồng...)."),
            ("COL3", "Thủ tục định giá và thế chấp tài sản đảm bảo tại VietinBank được thực hiện nhanh chóng, thuận tiện.")
        ]),
        
        ("III. Thời gian & Tốc độ xử lý hồ sơ (SPE)", [
            ("SPE1", "Thời gian thẩm định và phê duyệt hạn mức bảo lãnh tại VietinBank rất nhanh chóng."),
            ("SPE2", "Quy trình thủ tục hồ sơ đề nghị cấp bảo lãnh tại VietinBank được đơn giản hóa tối đa."),
            ("SPE3", "Thời gian từ khi nộp đủ hồ sơ đến khi nhận Thư bảo lãnh gốc đáp ứng kịp thời tiến độ hợp đồng.")
        ]),

        ("IV. Uy tín & Thương hiệu VietinBank (REP)", [
            ("REP1", "VietinBank là ngân hàng thương mại lớn, có uy tín và thương hiệu mạnh hàng đầu Việt Nam."),
            ("REP2", "Thư bảo lãnh do VietinBank phát hành luôn được các Chủ đầu tư và Bên mời thầu tin tưởng chấp nhận."),
            ("REP3", "Năng lực tài chính vững mạnh của VietinBank giúp gia tăng uy tín cho doanh nghiệp khi giao dịch.")
        ]),

        ("V. Trình độ & Thái độ phục vụ của Cán bộ (STA)", [
            ("STA1", "Cán bộ quan hệ khách hàng (RM) VietinBank có năng lực chuyên môn sâu về nghiệp vụ bảo lãnh."),
            ("STA2", "Cán bộ VietinBank am hiểu đặc thù ngành nghề kinh doanh và luôn tư vấn giải pháp phù hợp."),
            ("STA3", "Thái độ phục vụ của nhân viên VietinBank luôn tận tình, chuyên nghiệp và hỗ trợ kịp thời.")
        ]),

        ("VI. Mối quan hệ tín dụng sẵn có (REL)", [
            ("REL1", "Doanh nghiệp có lịch sử tín dụng tốt và mối quan hệ gắn kết lâu năm với VietinBank."),
            ("REL2", "Việc đang mở tài khoản thanh toán / trả lương tại VietinBank giúp thủ tục bảo lãnh thuận lợi hơn."),
            ("REL3", "Doanh nghiệp dễ dàng được cấp hạn mức bảo lãnh nhờ đã từng sử dụng các dịch vụ tài trợ khác.")
        ]),

        ("VII. Mức độ Chuyển đổi số & Bảo lãnh điện tử (DIG - Biến mới 1)", [
            ("DIG1", "Doanh nghiệp có thể dễ dàng nộp đề nghị cấp bảo lãnh online 24/7 qua VietinBank eFAST."),
            ("DIG2", "VietinBank phát hành Thư bảo lãnh điện tử ứng dụng chữ ký số CA nhanh chóng, không cần bản giấy."),
            ("DIG3", "Tính năng tra cứu mã QR xác thực Thư bảo lãnh thật/giả online trên hệ thống rất tiện ích và an toàn.")
        ]),

        ("VIII. Thiết kế Giải pháp Bảo lãnh theo yêu cầu (CUS - Biến mới 2)", [
            ("CUS1", "VietinBank có khả năng thiết kế linh hoạt mẫu thư bảo lãnh theo yêu cầu riêng của Chủ đầu tư."),
            ("CUS2", "VietinBank cung cấp cấu trúc bảo lãnh đa dạng (bảo lãnh tạm ứng, bảo lãnh thanh toán, đối ứng...)."),
            ("CUS3", "Ngân hàng sẵn sàng điều chỉnh các điều khoản bảo lãnh phù hợp với tiến độ thi công thực tế.")
        ]),

        ("IX. Năng lực Tư vấn Pháp lý & Quản trị Rủi ro (RSK - Biến mới 3)", [
            ("RSK1", "VietinBank hỗ trợ tư vấn rà soát tính pháp lý của hợp đồng gốc trước khi phát hành bảo lãnh."),
            ("RSK2", "Cán bộ VietinBank tư vấn cấu trúc điều khoản bảo lãnh giúp phòng ngừa rủi ro bị tịch thu vô lý."),
            ("RSK3", "VietinBank áp dụng tốt các tập quán bảo lãnh quốc tế (URDG 758) giúp bảo vệ quyền lợi doanh nghiệp.")
        ]),

        ("X. Chính sách Bảo lãnh Xanh & Ưu đãi ESG (ESG - Biến mới 4)", [
            ("ESG1", "VietinBank có chính sách ưu đãi riêng (giảm phí/tăng hạn mức) cho các dự án đạt tiêu chuẩn Xanh/ESG."),
            ("ESG2", "Quy trình thẩm định bảo lãnh được ưu tiên luồng xử lý nhanh cho doanh nghiệp hoạt động vì môi trường."),
            ("ESG3", "Chính sách tín dụng Xanh của VietinBank tạo động lực cho doanh nghiệp chuyển đổi phát triển bền vững.")
        ]),

        ("XI. Mạng lưới Ngân hàng Đại lý Quốc tế (NET - Biến mới 5)", [
            ("NET1", "VietinBank có mạng lưới ngân hàng đại lý rộng khắp toàn cầu, hỗ trợ tốt cho giao dịch quốc tế."),
            ("NET2", "Thư bảo lãnh đối ứng (Counter-Guarantee) của VietinBank dễ dàng được các đối tác nước ngoài chấp nhận."),
            ("NET3", "VietinBank hỗ trợ hiệu quả các thủ tục bảo lãnh bằng ngoại tệ đối với các hợp đồng xuất nhập khẩu.")
        ]),

        ("B. QUYẾT ĐỊNH LỰA CHỌN VIETINBANK (BIẾN MỤC TIÊU Y)", []),

        ("XII. Quyết định Lựa chọn & Sự hài lòng của Doanh nghiệp (DEC)", [
            ("DEC1", "Doanh nghiệp của bạn luôn ưu tiên lựa chọn VietinBank để phát hành các thư bảo lãnh."),
            ("DEC2", "Doanh nghiệp sẵn sàng tăng quy mô và tần suất sử dụng dịch vụ bảo lãnh tại VietinBank trong tương lai."),
            ("DEC3", "Doanh nghiệp sẵn sàng giới thiệu dịch vụ bảo lãnh của VietinBank cho các đối tác, khách hàng khác."),
            ("DEC4", "Nhìn chung, VietinBank là lựa chọn tối ưu nhất của doanh nghiệp khi phát sinh nhu cầu bảo lãnh.")
        ])
    ]

    # Create Table
    # Columns: STT / Code | Statement | 1 | 2 | 3 | 4 | 5
    table = doc.add_table(rows=1, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_widths = [Inches(0.9), Inches(3.8), Inches(0.4), Inches(0.4), Inches(0.4), Inches(0.4), Inches(0.4)]

    # Header Row
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Mã biến", "Các phát biểu đánh giá khảo sát", "1", "2", "3", "4", "5"]
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].width = col_widths[i]
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(hdr_cells[i], "003366") # Navy blue background

    stt_count = 1
    for group_title, items in items_data:
        if not items:
            # Section Main Header Row
            row_cells = table.add_row().cells
            # Merge all cells in row
            a = row_cells[0]
            b = row_cells[6]
            a.merge(b)
            set_cell_background(a, "E6F0FA")
            p = a.paragraphs[0]
            r = p.add_run(group_title)
            r.bold = True
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
        else:
            # Sub Group Header Row
            sub_cells = table.add_row().cells
            sub_a = sub_cells[0]
            sub_b = sub_cells[6]
            sub_a.merge(sub_b)
            set_cell_background(sub_a, "F2F4F7")
            p = sub_a.paragraphs[0]
            r = p.add_run(group_title)
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(0x00, 0x55, 0x99)

            # Add Item Rows
            for code, statement in items:
                row_cells = table.add_row().cells
                for idx, w in enumerate(col_widths):
                    row_cells[idx].width = w
                
                # Code cell
                p_code = row_cells[0].paragraphs[0]
                p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_code = p_code.add_run(code)
                r_code.bold = True
                r_code.font.size = Pt(9.5)

                # Statement cell
                p_stmt = row_cells[1].paragraphs[0]
                r_stmt = p_stmt.add_run(statement)
                r_stmt.font.size = Pt(9.5)

                # Likert 1-5 cells
                for c_idx in range(2, 7):
                    p_l = row_cells[c_idx].paragraphs[0]
                    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_l = p_l.add_run("[  ]")
                    r_l.font.size = Pt(9)
                    r_l.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                
                stt_count += 1

    # Outro
    doc.add_paragraph()
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("XIN CHÂN THÀNH CẢM ƠN SỰ HỢP TÁC VÀ HỖ TRỢ CỦA QUÝ DOANH NGHIỆP!\n"
                          "Kính chúc Quý Doanh nghiệp ngày càng Phát triển và Thành công!")
    r_end.bold = True
    r_end.italic = True
    r_end.font.size = Pt(11)
    r_end.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Save
    file_path = "c:/Users/nguyen.tuan.minh/Desktop/DTL-Master-Project/Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx"
    doc.save(file_path)
    print(f"Successfully generated DOCX file at {file_path}")

if __name__ == "__main__":
    create_survey_docx()
