import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def add_bottom_border_to_paragraph(paragraph, color_hex="888888", size="6"):
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)

def generate_full_thesis_template():
    doc = docx.Document()

    # 1. Page Margins (1.0 inch = 2.54 cm standard)
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # 2. Running Header Configuration (Different First Page)
    section.different_first_page_header_footer = True
    header = section.header
    p_head = header.paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_head = p_head.add_run("Master Thesis – MDE31 – Dang Tu Linh")
    r_head.font.name = 'Times New Roman'
    r_head.font.size = Pt(10)
    r_head.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    r_head.italic = True
    add_bottom_border_to_paragraph(p_head, color_hex="888888", size="6")

    # Configure Styles
    styles = doc.styles

    # Normal Style
    style_normal = styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    # Heading 1 Style
    style_h1 = styles['Heading 1']
    style_h1.font.name = 'Times New Roman'
    style_h1.font.size = Pt(14)
    style_h1.font.bold = True
    style_h1.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    style_h1.paragraph_format.space_before = Pt(14)
    style_h1.paragraph_format.space_after = Pt(6)

    # Heading 2 Style
    style_h2 = styles['Heading 2']
    style_h2.font.name = 'Times New Roman'
    style_h2.font.size = Pt(12.5)
    style_h2.font.bold = True
    style_h2.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    style_h2.paragraph_format.space_before = Pt(10)
    style_h2.paragraph_format.space_after = Pt(4)

    # Heading 3 Style
    style_h3 = styles['Heading 3']
    style_h3.font.name = 'Times New Roman'
    style_h3.font.size = Pt(12)
    style_h3.font.bold = True
    style_h3.font.italic = True
    style_h3.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    style_h3.paragraph_format.space_before = Pt(6)
    style_h3.paragraph_format.space_after = Pt(2)

    # Helper functions
    def add_guide(text, space_after=6):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(f"👉 [Nội dung hướng dẫn]: {text}")
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        r.italic = True
        return p

    def add_p(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.add_run(text)
        return p

    # ==================== TRANG BÌA CHÍNH (COVER PAGE) ====================
    p_cov1 = doc.add_paragraph()
    p_cov1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov1.paragraph_format.space_after = Pt(4)
    r = p_cov1.add_run("NATIONAL ECONOMICS UNIVERSITY\nVIETNAM-NETHERLANDS MASTER’S PROGRAM IN DEVELOPMENT ECONOMICS (MDE)")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_cov2 = doc.add_paragraph()
    p_cov2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov2.paragraph_format.space_before = Pt(60)
    p_cov2.paragraph_format.space_after = Pt(24)
    r = p_cov2.add_run("MASTER THESIS")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_cov3 = doc.add_paragraph()
    p_cov3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov3.paragraph_format.space_after = Pt(60)
    r = p_cov3.add_run("FACTORS AFFECTING CORPORATE CUSTOMERS’ DECISION TO CHOOSE BANK GUARANTEE SERVICES AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_cov4 = doc.add_paragraph()
    p_cov4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov4.paragraph_format.space_after = Pt(4)
    r = p_cov4.add_run("Student: DANG TU LINH")
    r.bold = True
    r.font.size = Pt(12)

    p_cov5 = doc.add_paragraph()
    p_cov5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov5.paragraph_format.space_after = Pt(4)
    r = p_cov5.add_run("Student ID / Class: MDE Class 31")
    r.font.size = Pt(12)

    p_cov6 = doc.add_paragraph()
    p_cov6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov6.paragraph_format.space_after = Pt(60)
    r = p_cov6.add_run("Supervisor: Dr. HOANG THI THUY NGA")
    r.bold = True
    r.font.size = Pt(12)

    p_cov7 = doc.add_paragraph()
    p_cov7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_cov7.add_run("Hanoi, 2026")
    r.italic = True
    r.font.size = Pt(11)

    doc.add_page_break()

    # ==================== TRANG BÌA PHỤ (TITLE PAGE) ====================
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_after = Pt(4)
    r = p_t1.add_run("NATIONAL ECONOMICS UNIVERSITY\nERASMUS UNIVERSITY ROTTERDAM - ISS")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_before = Pt(40)
    p_t2.paragraph_format.space_after = Pt(20)
    r = p_t2.add_run("MASTER’S THESIS IN DEVELOPMENT ECONOMICS")
    r.bold = True
    r.font.size = Pt(15)

    p_t3 = doc.add_paragraph()
    p_t3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t3.paragraph_format.space_after = Pt(40)
    r = p_t3.add_run("FACTORS AFFECTING CORPORATE CUSTOMERS’ DECISION TO CHOOSE BANK GUARANTEE SERVICES AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_t4 = doc.add_paragraph()
    p_t4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t4.paragraph_format.space_after = Pt(6)
    r = p_t4.add_run("Author: DANG TU LINH")
    r.bold = True
    r.font.size = Pt(12)

    p_t5 = doc.add_paragraph()
    p_t5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t5.paragraph_format.space_after = Pt(40)
    r = p_t5.add_run("Academic Supervisor: Dr. HOANG THI THUY NGA")
    r.bold = True
    r.font.size = Pt(12)

    p_t6 = doc.add_paragraph()
    p_t6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_t6.add_run("A thesis submitted in partial fulfillment of the requirements for the degree of\nMaster of Arts in Development Economics\nHanoi, June 2026")
    r.italic = True
    r.font.size = Pt(11)

    doc.add_page_break()

    # ==================== STATEMENT OF AUTHORSHIP ====================
    doc.add_heading("STATEMENT OF AUTHORSHIP (LỜI CAM ĐOAN)", level=1)
    add_guide("Trình bày lời cam đoan của học viên về tính trung thực, độc lập và đạo đức nghiên cứu khoa học. Xác nhận toàn bộ dữ liệu khảo sát 800 doanh nghiệp và kết quả ước lượng hồi quy OLS là số liệu thu thập thực tế, không sao chép bất hợp pháp từ bất kỳ công trình nào.")
    add_p("I hereby declare that this Master’s thesis entitled \"Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)\" is my own independent research work conducted under the academic supervision of Dr. Hoang Thi Thuy Nga. The survey dataset (n = 800) and econometric findings presented in this thesis are original, transparent, and have not been submitted for any other degree or qualification at any academic institution.", space_after=18)
    
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.paragraph_format.space_after = Pt(4)
    p_sig.add_run("Hanoi, June 2026\nStudent Author\n\n\n\n").italic = True
    p_sig.add_run("Dang Tu Linh").bold = True

    doc.add_page_break()

    # ==================== ACKNOWLEDGEMENTS ====================
    doc.add_heading("ACKNOWLEDGEMENTS (LỜI CẢM ƠN)", level=1)
    add_guide("Dành lời cảm ơn trân trọng tới: (1) Ban Giám hiệu NEU, Viện MDE và Đại học Erasmus Rotterdam / ISS; (2) TS. Hoàng Thị Thúy Nga - GVHD đã tận tình định hướng và phản biện; (3) Ban Lãnh đạo, cán bộ Khối KHDN Ngân hàng VietinBank trên 155 chi nhánh đã hỗ trợ phát phiếu khảo sát; (4) Đại diện 800 doanh nghiệp tham gia trả lời phiếu.")
    add_p("[Nội dung lời cảm ơn sẽ được hoàn thiện tại đây...]")

    doc.add_page_break()

    # ==================== ABSTRACT (TIẾNG ANH) ====================
    doc.add_heading("ABSTRACT", level=1)
    add_guide("Tóm tắt luận văn bằng Tiếng Anh (khoảng 300 - 500 từ) gồm 5 phần: (1) Background & Objective: Đánh giá nhân tố ảnh hưởng đến quyết định chọn bảo lãnh VietinBank; (2) Methodology: Khảo sát n = 800 DN, OLS, EFA, Cronbach's Alpha, ANOVA; (3) Key Findings: 7 nhân tố đều tác động dương (+), trong đó COST_COMP và BANK_REP mạnh nhất, eFAST có tác động ý nghĩa, phát hiện khác biệt giữa nhóm SOE/SME/FDI; (4) Policy Implications: 4 nhóm giải pháp cho VietinBank; (5) Keywords: Bank Guarantee, Corporate Customer, Selection Decision, eFAST, VietinBank, Multiple Regression.")
    add_p("[English Abstract content will be inserted here...]")
    add_p("Keywords: Bank Guarantee Services, Corporate Selection Decision, Price Competitiveness, eFAST Digital Adoption, OLS Regression, VietinBank.", space_after=18)

    doc.add_page_break()

    # ==================== TÓM TẮT LUẬN VĂN (TIẾNG VIỆT) ====================
    doc.add_heading("TÓM TẮT LUẬN VĂN", level=1)
    add_guide("Bản tóm tắt tiếng Việt tương ứng với Abstract, trình bày rõ bối cảnh, mô hình 7 biến độc lập, cỡ mẫu 800 doanh nghiệp và các kiến nghị giải pháp then chốt cho VietinBank.")
    add_p("[Nội dung Tóm tắt tiếng Việt sẽ được hoàn thiện tại đây...]")
    add_p("Từ khóa: Bảo lãnh ngân hàng, Khách hàng doanh nghiệp, Quyết định lựa chọn, Tính cạnh tranh biểu phí, Chuyển đổi số eFAST, Hồi quy OLS, VietinBank.", space_after=18)

    doc.add_page_break()

    # ==================== MỤC LỤC & DANH MỤC ====================
    doc.add_heading("TABLE OF CONTENTS (MỤC LỤC)", level=1)
    add_guide("Trang Mục lục tự động (Table of Contents) được tạo bằng tính năng References -> Table of Contents của Microsoft Word. Khi hoàn thiện bài, chỉ cần bấm Update Table là toàn bộ số trang của các Headings sẽ tự động cập nhật.")
    add_p("[Table of Contents will be generated automatically in MS Word...]")

    doc.add_page_break()

    doc.add_heading("LIST OF ABBREVIATIONS (DANH MỤC TỪ VIẾT TẮT)", level=1)
    add_guide("Bảng giải nghĩa tất cả các chữ viết tắt trong luận văn: Big4, OLS, EFA, VIF, KMO, TAM, CCF, NIM, ROE, RM, STP, SOE, SME, FDI, TG, PG, APG, BG, URDG, SBV...")
    
    t_abbr = doc.add_table(rows=1, cols=3)
    t_abbr.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = [Inches(1.2), Inches(2.3), Inches(3.0)]
    for i, w in enumerate(col_w): t_abbr.rows[0].cells[i].width = w
    hdr = t_abbr.rows[0].cells
    hdr[0].paragraphs[0].add_run("Abbreviation").bold = True
    hdr[1].paragraphs[0].add_run("Full English Term").bold = True
    hdr[2].paragraphs[0].add_run("Vietnamese Equivalent").bold = True
    for c in hdr: set_cell_background(c, "003366"); c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    sample_abbr = [
        ("APG", "Advance Payment Guarantee", "Bảo lãnh hoàn trả tiền tạm ứng"),
        ("EFA", "Exploratory Factor Analysis", "Phân tích nhân tố khám phá"),
        ("eFAST", "VietinBank Corporate Digital Banking", "Nền tảng ngân hàng số doanh nghiệp VietinBank"),
        ("FDI", "Foreign Direct Investment", "Doanh nghiệp có vốn đầu tư trực tiếp nước ngoài"),
        ("OLS", "Ordinary Least Squares", "Phương pháp bình phương bé nhất thông thường"),
        ("PG", "Performance Guarantee", "Bảo lãnh thực hiện hợp đồng"),
        ("TG", "Tender Guarantee / Bid Bond", "Bảo lãnh dự thầu"),
        ("URDG 758", "Uniform Rules for Demand Guarantees 758", "Quy tắc thống nhất về bảo lãnh theo yêu cầu (ICC)"),
        ("VIF", "Variance Inflation Factor", "Hệ số phóng đại phương sai (Đo đa cộng tuyến)")
    ]
    for a, e, v in sample_abbr:
        rc = t_abbr.add_row().cells
        for i, w in enumerate(col_w): rc[i].width = w
        rc[0].paragraphs[0].add_run(a).bold = True
        rc[1].paragraphs[0].add_run(e)
        rc[2].paragraphs[0].add_run(v)

    doc.add_page_break()

    doc.add_heading("LIST OF TABLES (DANH MỤC BẢNG BIỂU)", level=1)
    add_guide("Danh mục thống kê tự động tất cả các bảng trong luận văn kèm số trang (Table 1.1, Table 3.1, Table 4.1, Table 4.2...).")
    add_p("[List of Tables will be generated automatically in MS Word...]")

    doc.add_page_break()

    doc.add_heading("LIST OF FIGURES (DANH MỤC HÌNH VẼ & SƠ ĐỒ)", level=1)
    add_guide("Danh mục thống kê tự động tất cả các sơ đồ, hình vẽ và biểu đồ trong luận văn (Figure 2.1 Khung phân tích, Figure 3.1 Quy trình nghiên cứu, Figure 4.1 Cơ cấu mẫu...).")
    add_p("[List of Figures will be generated automatically in MS Word...]")

    doc.add_page_break()

    # ==================== CHƯƠNG 1 ====================
    doc.add_heading("CHAPTER 1: INTRODUCTION", level=1)

    doc.add_heading("1.1. Research Rationales & Background", level=2)
    add_guide("Trình bày bối cảnh kinh tế vĩ mô và tính cấp thiết: (1) Bản chất bảo lãnh ngân hàng là công cụ bảo đảm nghĩa vụ hợp đồng bắt buộc trong xây dựng, đấu thầu theo Luật Đấu thầu số 22/2023/QH15 và Nghị định số 35/2023/NĐ-CP; (2) Thông tư 61/2024/TT-NHNN tạo hành lang pháp lý mới cho bảo lãnh điện tử (e-guarantees); (3) Tầm quan trọng của thu nhập phí bảo lãnh phi tín dụng đối với ngân hàng thương mại; (4) Áp lực cạnh tranh gay gắt khiến VietinBank cần hiểu rõ các yếu tố quyết định lựa chọn của doanh nghiệp.")

    doc.add_heading("1.2. Research Problem & Industry Context at VietinBank", level=2)
    add_guide("Nhận diện vấn đề nghiên cứu thực tế tại VietinBank: Khách hàng doanh nghiệp ngày càng khắt khe về thời gian phát hành, tính cạnh tranh của biểu phí và tính linh hoạt của tài sản thế chấp. Làm rõ bối cảnh mạng lưới 155 chi nhánh của VietinBank và khoảng trống trong việc nghiên cứu hành vi lựa chọn dịch vụ bảo lãnh B2B.")

    doc.add_heading("1.3. Research Objectives", level=2)
    doc.add_heading("1.3.1. General Objective", level=3)
    add_guide("Cập nhật nguyên văn tiếng Anh của cô giáo hướng dẫn:")
    add_p("\"The general objective of this study is to identify and assess the factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, and to propose managerial recommendations for improving VietinBank’s attractiveness and competitiveness in the corporate bank guarantee market.\"")

    doc.add_heading("1.3.2. Specific Objectives", level=3)
    add_guide("Cập nhật nguyên văn 4 mục tiêu cụ thể cô giáo đã phê duyệt (bao gồm phân tích khác biệt nhóm ở Objective 3):")
    add_p("1. Identify the key factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, based on relevant theories, previous empirical studies, and the characteristics of bank guarantee services.")
    add_p("2. Assess the direction and relative importance of these factors in explaining corporate customers’ selection decisions.")
    add_p("3. Examine whether corporate customers’ selection decisions differ across major firm characteristics, such as ownership type, firm size, operating experience, and types of bank guarantees used.")
    add_p("4. Propose managerial recommendations for VietinBank to improve its bank guarantee products and services and strengthen its ability to attract and retain corporate customers.")

    doc.add_heading("1.4. Research Questions", level=2)
    add_guide("Thiết lập 4 câu hỏi nghiên cứu tương ứng trực tiếp với 4 mục tiêu cụ thể ở trên:")
    add_p("• Research Question 1: What key factors significantly influence corporate customers’ decision to choose VietinBank for bank guarantee services?")
    add_p("• Research Question 2: What is the direction and relative importance of each factor in explaining corporate selection decisions?")
    add_p("• Research Question 3: Do corporate selection decisions significantly differ across corporate ownership types, firm revenue sizes, operating tenure, and guarantee product lines?")
    add_p("• Research Question 4: What actionable managerial recommendations should VietinBank implement to enhance its competitive standing, product attractiveness, and client retention?")

    doc.add_heading("1.5. Scope and Boundaries of the Study", level=2)
    add_guide("Xác định phạm vi nghiên cứu: (1) Phạm vi đối tượng: Khách hàng doanh nghiệp có nhu cầu sử dụng bảo lãnh; (2) Phạm vi không gian: 155 chi nhánh của VietinBank trên toàn quốc; (3) Phạm vi thời gian: Khảo sát thực tế trong giai đoạn 2025 - 2026; (4) Phạm vi nội dung: Tập trung vào 7 nhân tố dịch vụ và quyết định lựa chọn ngân hàng phục vụ.")

    doc.add_heading("1.6. Significance & Contributions of the Study", level=2)
    add_guide("Nêu rõ 2 đóng góp lớn: (1) Về mặt học thuật: Bổ sung khoảng trống nghiên cứu thực nghiệm về bảo lãnh ngân hàng B2B gắn với chuyển đổi số eFAST tại thị trường mới nổi Việt Nam; (2) Về mặt thực tiễn: Cung cấp bằng chứng định lượng giúp Ban Lãnh đạo VietinBank tối ưu hóa chính sách phí, quy trình thẩm định và phân khúc khách hàng.")

    doc.add_heading("1.7. Structure of the Thesis", level=2)
    add_guide("Giới thiệu tóm tắt kết cấu 4 chương chính của luận văn (Chương 1: Mở đầu; Chương 2: Tổng quan lý thuyết; Chương 3: Phương pháp nghiên cứu; Chương 4: Kết quả nghiên cứu, Thảo luận & Hàm ý quản trị).")

    doc.add_page_break()

    # ==================== CHƯƠNG 2 ====================
    doc.add_heading("CHAPTER 2: LITERATURE REVIEW AND THEORETICAL FRAMEWORK", level=1)

    doc.add_heading("2.1. Overview of Bank Guarantee Services in Commercial Banking", level=2)
    doc.add_heading("2.1.1. Nature and Economic Functions of Bank Guarantees", level=3)
    add_guide("Phân tích bản chất kinh tế: Cam kết bảo lãnh là khoản tín dụng ngoại bảng (Off-balance sheet commitment), tạo uy tín trung gian giúp doanh nghiệp thế chấp tín nhiệm thay cho việc ký quỹ tiền mặt 100%, bảo toàn vốn lưu động cho nhà thầu.")

    doc.add_heading("2.1.2. Main Types of Corporate Bank Guarantees", level=3)
    add_guide("Trình bày đặc điểm, vai trò và tỷ trọng của 4 sản phẩm chính: (1) Bảo lãnh dự thầu (Tender Guarantee - TG); (2) Bảo lãnh thực hiện hợp đồng (Performance Guarantee - PG); (3) Bảo lãnh hoàn trả tiền tạm ứng (Advance Payment Guarantee - APG); (4) Bảo lãnh thanh toán (Payment Guarantee - BG).")

    doc.add_heading("2.1.3. Legal and Regulatory Framework", level=3)
    add_guide("Phân tích khung pháp lý: Thông tư 61/2024/TT-NHNN của NHNN (hiệu lực 01/4/2025 về bảo lãnh điện tử), Luật Đấu thầu số 22/2023/QH15, Nghị định 35/2023/NĐ-CP, Bộ luật Dân sự 2015 và Tập quán quốc tế ICC URDG 758.")

    doc.add_heading("2.2. Theoretical Foundations", level=2)
    doc.add_heading("2.2.1. Financial Intermediation & Delegated Monitoring Theory", level=3)
    add_guide("Lý thuyết Trung gian Tài chính & Giám sát Ủy thác (Diamond, 1984; Ramakrishnan & Thakor, 1984): Ngân hàng đóng vai trò sản xuất thông tin và bảo chứng tín nhiệm cho doanh nghiệp trước Chủ đầu tư.")

    doc.add_heading("2.2.2. Credit Risk Pricing & Contingent Claim Theory", level=3)
    add_guide("Lý thuyết Định giá Rủi ro Tín dụng (Merton, 1974; Stiglitz & Weiss, 1981): Mức phí bảo lãnh và tỷ lệ ký quỹ phản ánh xác suất rủi ro và giá trị tài sản đảm bảo của doanh nghiệp.")

    doc.add_heading("2.2.3. Service Quality Theory & SERVQUAL Model", level=3)
    add_guide("Mô hình Chất lượng Dịch vụ SERVQUAL (Parasuraman et al., 1988): Đo lường cảm nhận chất lượng dịch vụ của khách hàng. Lưu ý phân định rõ Uy tín thương hiệu (BANK_REP) tách biệt khỏi Độ tin cậy (Reliability) theo đúng nhắc nhở của GVHD.")

    doc.add_heading("2.2.4. Relationship Banking Theory", level=3)
    add_guide("Lý thuyết Ngân hàng Quan hệ (Boot, 2000; Berger & Udell, 1995): Khách hàng giao dịch lâu năm và sử dụng trọn gói sản phẩm sẽ giúp giảm bất cân xứng thông tin, được ngân hàng cấp hạn mức bảo lãnh lớn và ưu đãi phí.")

    doc.add_heading("2.2.5. Technology Acceptance Model (TAM) & Digital Banking", level=3)
    add_guide("Mô hình Chấp nhận Công nghệ TAM (Davis, 1989; Venkatesh et al., 2003): Cơ sở bảo chứng vững chắc cho biến Chuyển đổi số eFAST (DIGITAL_CONV) dựa trên Tính hữu ích cảm nhận và Tính dễ sử dụng.")

    doc.add_heading("2.3. Empirical Literature on Corporate Bank Selection", level=2)
    doc.add_heading("2.3.1. International Empirical Studies", level=3)
    add_guide("Tổng quan các công trình quốc tế tiêu biểu về tiêu chí lựa chọn ngân hàng của doanh nghiệp: Turnbull & Gibbs (1989), Narteh (2013), Al-Sabbagh & Al-Khathlan (2018), Kaur et al. (2021), Zelie (2023), Carletti et al. (2023).")

    doc.add_heading("2.3.2. Empirical Studies in the Vietnamese Banking Context", level=3)
    add_guide("Tổng quan các nghiên cứu tại Việt Nam về dịch vụ ngân hàng B2B: Phan Thị Hằng Nga et al. (2024), Hồ Đình Phi et al. (2023), Nguyễn et al. (2024), Lê Văn Dũng (2021), Nguyễn Thị Nhung & Nguyễn Duy Phú (2015).")

    doc.add_heading("2.4. Research Gaps", level=2)
    add_guide("Chỉ ra 4 khoảng trống nghiên cứu chính: (1) Nghiên cứu trước chủ yếu tập trung vào mảng bán lẻ hoặc cho vay chung, thiếu vắng nghiên cứu chuyên sâu về bảo lãnh ngoại bảng; (2) Chưa có mô hình tích hợp biến Chuyển đổi số eFAST và bảo lãnh điện tử; (3) Thiếu nghiên cứu so sánh khác biệt nhóm doanh nghiệp (SOE vs Private vs FDI); (4) Chưa có nghiên cứu định lượng quy mô lớn tại VietinBank (n = 800).")

    doc.add_heading("2.5. Conceptual Framework and Research Hypotheses", level=2)
    doc.add_heading("2.5.1. Conceptual Research Framework", level=3)
    add_guide("Mô tả sơ đồ khung phân tích gồm 7 biến độc lập tác động trực tiếp lên 1 biến phụ thuộc DEC (Selection Decision).")

    doc.add_heading("2.5.2. Hypothesis Development", level=3)
    add_guide("Lập luận cơ sở lý thuyết và phát biểu 7 giả thuyết nghiên cứu (tất cả đều kỳ vọng tác động dương + theo đúng ý cô):")
    add_p("• Hypothesis H1: Price Competitiveness (COST_COMP) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H2: Processing Speed (PROC_SPEED) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H3: Digital eFAST Convenience (DIGITAL_CONV) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H4: Bank Reputation (BANK_REP) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H5: Relationship Banking & Limits (RELATIONSHIP) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H6: Staff Professionalism (STAFF_QUAL) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H7: Collateral & Margin Flexibility (COLL_POLICY) has a positive impact on corporate customers' decision to choose VietinBank.")

    doc.add_page_break()

    # ==================== CHƯƠNG 3 ====================
    doc.add_heading("CHAPTER 3: RESEARCH METHODOLOGY AND EMPIRICAL DESIGN", level=1)

    doc.add_heading("3.1. Overall Research Design & Analytical Process", level=2)
    add_guide("Mô tả thiết kế nghiên cứu kết hợp định tính (hiệu chỉnh bảng hỏi) và định lượng (OLS). Trình bày sơ đồ quy trình 7 bước tuần tự: Thống kê mô tả -> Cronbach's Alpha -> EFA -> Factor Scores -> Tương quan Pearson -> Kiểm định VIF -> Hồi quy OLS và Kiểm định ANOVA/t-test theo đúng gợi ý của GVHD.")

    doc.add_heading("3.2. Questionnaire Design & Measurement Scales", level=2)
    doc.add_heading("3.2.1. Operationalization of Variables", level=3)
    add_guide("Bảng định nghĩa thang đo Likert 1-5 của 7 biến độc lập và 1 biến phụ thuộc, bám sát các câu hỏi thực tế trong phiếu khảo sát 34 câu.")

    # Table of Variables
    t_v = doc.add_table(rows=1, cols=5)
    t_v.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w_v = [Inches(1.2), Inches(1.8), Inches(2.5), Inches(0.9), Inches(0.6)]
    for i, w in enumerate(col_w_v): t_v.rows[0].cells[i].width = w
    hdr_v = t_v.rows[0].cells
    hdr_v[0].paragraphs[0].add_run("Code").bold = True
    hdr_v[1].paragraphs[0].add_run("Variable Name").bold = True
    hdr_v[2].paragraphs[0].add_run("Measurement Content (34 Items)").bold = True
    hdr_v[3].paragraphs[0].add_run("Type").bold = True
    hdr_v[4].paragraphs[0].add_run("Sign").bold = True
    for c in hdr_v: set_cell_background(c, "003366"); c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    v_data = [
        ("DEC", "Selection Decision", "Preference & priority choice of VietinBank over competitors (4 items)", "Dependent (Y)", "N/A"),
        ("COST_COMP", "Price Competitiveness", "Fee reasonableness, competitive pricing & discount incentives (4 items)", "Independent (X1)", "+"),
        ("PROC_SPEED", "Processing Speed", "Turnaround time, prompt issuance & simplified paperwork (4 items)", "Independent (X2)", "+"),
        ("DIGITAL_CONV", "Digital eFAST Convenience", "Online submission, 24/7 e-guarantees & digital status tracking (4 items)", "Independent (X3)", "+"),
        ("BANK_REP", "Bank Reputation", "Big4 brand prestige, financial strength & 100% acceptance (5 items)", "Independent (X4)", "+"),
        ("RELATIONSHIP", "Relationship & Limits", "Credit limit flexibility, relationship tenure & VIP care (5 items)", "Independent (X5)", "+"),
        ("STAFF_QUAL", "Staff Professionalism", "RM competence, legal advisory on Bidding Law & TT61 (4 items)", "Independent (X6)", "+"),
        ("COLL_POLICY", "Collateral Flexibility", "Flexible cash margin ratio & diverse pledged collateral (4 items)", "Independent (X7)", "+")
    ]
    for c, n, m, t, s in v_data:
        rc = t_v.add_row().cells
        for i, w in enumerate(col_w_v): rc[i].width = w
        rc[0].paragraphs[0].add_run(c).bold = True
        rc[1].paragraphs[0].add_run(n)
        rc[2].paragraphs[0].add_run(m)
        rc[3].paragraphs[0].add_run(t)
        rc[4].paragraphs[0].add_run(s).bold = True

    doc.add_heading("3.2.2. Mapping Scales with the Official 34-Item Survey Questionnaire", level=3)
    add_guide("Chứng minh nguyên tắc 'Có bột mới gột nên hồ': Ánh xạ chi tiết từng mã câu hỏi từ CP1-CP4, TD1-TD4, CS1-CS4, UT1-UT5, QH1-QH5, CB1-CB4, TSBĐ1-TSBĐ4 và LC1-LC4 tương ứng với 7 biến X và biến Y trong bộ dữ liệu gốc.")

    doc.add_heading("3.3. Population, Sampling Strategy and Data Collection", level=2)
    doc.add_heading("3.3.1. Target Population & Sampling Method", level=3)
    add_guide("Mô tả tổng thể doanh nghiệp có giao dịch bảo lãnh và phương pháp chọn mẫu phân tầng kết hợp thuận tiện.")

    doc.add_heading("3.3.2. Sample Size Determination", level=3)
    add_guide("Biện luận quy mô mẫu n = 800: Vượt xa tiêu chuẩn tối thiểu của EFA (n >= 5 * 34 = 170) và hồi quy OLS (n >= 50 + 8*7 = 106), đảm bảo độ tin cậy thống kê tuyệt đối.")

    doc.add_heading("3.3.3. Survey Administration across 155 VietinBank Branches", level=3)
    add_guide("Mô tả quy trình thu thập dữ liệu thông qua hệ thống 155 chi nhánh VietinBank và kênh trực tuyến Google Forms.")

    doc.add_heading("3.4. Econometric & Quantitative Analytical Methods", level=2)
    doc.add_heading("3.4.1. Descriptive Statistics", level=3)
    add_guide("Phương pháp phân tích tần số, phần trăm, giá trị trung bình (Mean) và độ lệch chuẩn (Std Dev).")

    doc.add_heading("3.4.2. Scale Reliability Testing (Cronbach’s Alpha)", level=3)
    add_guide("Tiêu chuẩn đánh giá độ tin cậy thang đo: Hệ số Alpha >= 0.60 (tốt nếu >= 0.80) và tương quan biến - tổng (Corrected Item-Total Correlation) >= 0.30.")

    doc.add_heading("3.4.3. Exploratory Factor Analysis (EFA)", level=3)
    add_guide("Tiêu chuẩn phân tích nhân tố khám phá: KMO >= 0.50, kiểm định Bartlett sig < 0.05, phương pháp trích Principal Axis Factoring / Principal Components, phép xoay Varimax, hệ số tải nhân tố Factor Loading >= 0.50, tổng phương sai trích >= 50%.")

    doc.add_heading("3.4.4. Factor Scores Extraction Method", level=3)
    add_guide("Phương pháp tính giá trị biến đại diện: Trung bình cộng các biến quan sát (Mean Score) hoặc trích xuất điểm nhân tố (Factor Score / Regression method).")

    doc.add_heading("3.4.5. Pearson Correlation Analysis & Multicollinearity Diagnostics (VIF)", level=3)
    add_guide("Kiểm tra mối quan hệ tuyến tính ban đầu giữa các biến và chẩn đoán đa cộng tuyến qua Hệ số phóng đại phương sai (VIF < 5 và Tolerance > 0.2).")

    doc.add_heading("3.4.6. Multiple Linear Regression Model Specification (OLS)", level=3)
    add_guide("Trình bày phương trình hồi quy OLS đa biến chính thức:")
    add_p("DEC = β0 + β1*COST_COMP + β2*PROC_SPEED + β3*DIGITAL_CONV + β4*BANK_REP + β5*RELATIONSHIP + β6*STAFF_QUAL + β7*COLL_POLICY + ε     (3.1)")

    doc.add_heading("3.4.7. Sub-Group Difference Testing Methods (ANOVA & t-test)", level=3)
    add_guide("Trình bày kỹ thuật kiểm định khác biệt nhóm (theo Objective 3): Sử dụng Independent Samples t-test (so sánh 2 nhóm) và One-Way ANOVA kết hợp Post-hoc Tukey test (so sánh từ 3 nhóm trở lên) để tìm kiếm sự khác biệt trong quyết định lựa chọn theo: Hình thức sở hữu, Quy mô doanh thu, Thâm niên và Loại sản phẩm bảo lãnh.")

    doc.add_page_break()

    # ==================== CHƯƠNG 4 ====================
    doc.add_heading("CHAPTER 4: EMPIRICAL RESULTS, DISCUSSION AND MANAGERIAL RECOMMENDATIONS", level=1)

    doc.add_heading("4.1. Descriptive Statistics of the Sample (n = 800)", level=2)
    doc.add_heading("4.1.1. Ownership Type Distribution", level=3)
    add_guide("Trình bày bảng và biểu đồ phân bổ mẫu theo Doanh nghiệp Nhà nước (SOEs), Doanh nghiệp Tư nhân, và Doanh nghiệp FDI.")

    doc.add_heading("4.1.2. Firm Revenue Scale Distribution", level=3)
    add_guide("Phân bổ mẫu theo quy mô doanh thu: Doanh nghiệp siêu nhỏ, Nhỏ, Vừa (SMEs) và Doanh nghiệp lớn (Corporate).")

    doc.add_heading("4.1.3. Operating Experience Distribution", level=3)
    add_guide("Phân bổ mẫu theo số năm hoạt động (< 3 năm, 3-5 năm, 5-10 năm, > 10 năm).")

    doc.add_heading("4.1.4. Usage Distribution of Bank Guarantee Products", level=3)
    add_guide("Phân bổ tần suất sử dụng các loại bảo lãnh: Dự thầu (TG), Tạm ứng (APG), Thực hiện hợp đồng (PG), Thanh toán (BG).")

    doc.add_heading("4.2. Scale Reliability Analysis Results (Cronbach’s Alpha)", level=2)
    add_guide("Bảng tổng hợp kết quả Cronbach's Alpha cho 7 biến độc lập và biến phụ thuộc DEC. Đánh giá tất cả các thang đo đều đạt độ tin cậy cao (> 0.80), không có biến quan sát nào bị loại.")

    doc.add_heading("4.3. Exploratory Factor Analysis Results (EFA)", level=2)
    doc.add_heading("4.3.1. EFA for Independent Variables", level=3)
    add_guide("Trình bày kết quả EFA của 30 biến quan sát độc lập: Hệ số KMO, Kiểm định Bartlett, Bảng Tổng phương sai trích (Cumulative Variance > 50%), và Bảng Ma trận nhân tố xoay (Rotated Component Matrix) hội tụ chuẩn vào 7 nhân tố.")

    doc.add_heading("4.3.2. EFA for Dependent Variable (DEC)", level=3)
    add_guide("Trình bày kết quả EFA cho 4 biến quan sát của DEC hội tụ vào 1 nhân tố duy nhất.")

    doc.add_heading("4.4. Correlation Analysis & Multicollinearity Diagnostics (VIF)", level=2)
    add_guide("Bảng ma trận tương quan Pearson giữa 7 biến độc lập với biến DEC. Đánh giá sơ bộ các tương quan đều có ý nghĩa thống kê (sig < 0.01). Trình bày bảng hệ số VIF đều < 2.0 (khẳng định không xảy ra hiện tượng đa cộng tuyến).")

    doc.add_heading("4.5. Multiple Linear Regression Results (OLS)", level=2)
    doc.add_heading("4.5.1. Model Summary & Goodness of Fit", level=3)
    add_guide("Bảng tóm tắt mô hình OLS: Giá trị R, R-Square, Adjusted R-Square, Thống kê kiểm định F và mức ý nghĩa sig F-change.")

    doc.add_heading("4.5.2. Estimated Coefficients and Hypothesis Testing", level=3)
    add_guide("Bảng kết quả hồi quy: Trình bày hệ số hồi quy chưa chuẩn hóa (B), Sai số chuẩn (Std. Error), Hệ số chuẩn hóa (Beta), Thống kê t, và p-value. Đánh giá chấp nhận cả 7 giả thuyết H1 đến H7 mang dấu dương (+), xếp hạng tầm quan trọng tương đối giữa các biến (theo Objective 2).")

    doc.add_heading("4.6. Sub-Group Difference Analysis Results (ANOVA & t-test)", level=2)
    doc.add_heading("4.6.1. Selection Differences across Ownership Types", level=3)
    add_guide("Trình bày kết quả ANOVA so sánh sự khác biệt trong quyết định chọn giữa DN Nhà nước, DN Tư nhân và DN FDI (Thực hiện đúng Objective 3 của cô).")

    doc.add_heading("4.6.2. Selection Differences across Firm Scales and Operating Experience", level=3)
    add_guide("Trình bày kết quả ANOVA so sánh giữa các nhóm quy mô doanh thu và thâm niên hoạt động.")

    doc.add_heading("4.6.3. Selection Differences across Guarantee Product Types", level=3)
    add_guide("Trình bày kết quả ANOVA so sánh giữa các nhóm sử dụng Bảo lãnh dự thầu, Tạm ứng, Thực hiện hợp đồng và Thanh toán.")

    doc.add_heading("4.7. Discussion of Empirical Findings", level=2)
    add_guide("Thảo luận chuyên sâu ý nghĩa học thuật và thực tiễn: Giải thích vì sao Tính cạnh tranh biểu phí (COST_COMP) và Uy tín thương hiệu (BANK_REP) là hai nhân tố tác động mạnh nhất; Phân tích vai trò ngày càng tăng của kênh Chuyển đổi số eFAST (DIGITAL_CONV).")

    doc.add_heading("4.8. Managerial Implications & Policy Recommendations for VietinBank", level=2)
    doc.add_heading("4.8.1. Enhancing Core Service Capabilities (Response to Objective 1)", level=3)
    add_guide("Giải pháp tập trung nguồn lực phát triển đồng bộ 7 giá trị cốt lõi mà doanh nghiệp quan tâm hàng đầu.")

    doc.add_heading("4.8.2. Strategies for Price Competitiveness & Processing Speed (Response to Objective 2)", level=3)
    add_guide("Giải pháp xây dựng biểu phí linh hoạt, chính sách chiết khấu phí cho khách hàng phát hành số lượng lớn; Tinh gọn quy trình thẩm định tín dụng để rút ngắn thời gian cấp bảo lãnh.")

    doc.add_heading("4.8.3. Tailored Guarantee Packages for Corporate Segments (Response to Objective 3)", level=3)
    add_guide("Giải pháp may đo sản phẩm riêng biệt dựa trên kết quả phân tích khác biệt nhóm: Gói hạn mức lớn cho SOEs; Gói phát hành nhanh online cho SMEs tư nhân; Gói bảo lãnh quốc tế theo URDG 758 cho khối FDI.")

    doc.add_heading("4.8.4. Breakthrough Digital Transformation Strategy via VietinBank eFAST (Response to Objective 4)", level=3)
    add_guide("Giải pháp chiến lược đẩy mạnh bảo lãnh điện tử (e-guarantees) 24/7 qua VietinBank eFAST, tự động hóa quy trình STP và tích hợp dịch vụ tư vấn pháp lý Luật Đấu thầu.")

    doc.add_heading("4.9. Policy Recommendations for the State Bank of Vietnam", level=2)
    add_guide("Kiến nghị với Ngân hàng Nhà nước về việc tiếp tục hoàn thiện khung pháp lý thực thi Thông tư 61/2024/TT-NHNN và thúc đẩy cơ chế xác thực chứng thư số quốc gia.")

    doc.add_heading("4.10. Research Limitations and Suggestions for Future Research", level=2)
    add_guide("Chỉ ra các giới hạn nghiên cứu (phương pháp chọn mẫu, yếu tố vĩ mô) và gợi ý hướng nghiên cứu mở rộng trong tương lai.")

    doc.add_page_break()

    # ==================== REFERENCES ====================
    doc.add_heading("REFERENCES (TÀI LIỆU THAM KHẢO)", level=1)
    add_guide("Danh mục toàn bộ 27 tài liệu tham khảo có thật 100% được xếp theo thứ tự bảng chữ cái A-Z chuẩn quốc tế Harvard.")

    refs = [
        "Al-Sabbagh, M. and Al-Khathlan, K. (2018) 'Factors influencing corporate clients’ choice of commercial banks for trade finance services', Journal of Financial Services Marketing, 23(2), pp. 71–82.",
        "Baltagi, B.H. (2008) Econometric Analysis of Panel Data. 4th edn. Chichester: Wiley.",
        "Barru, D.J. (2005) 'How to Guarantee Contractor Performance on International Construction Projects: Comparing Surety Bonds with Bank Guarantees and Standby Letters of Credit', The George Washington International Law Review, 37(1), pp. 51–94.",
        "Berger, A.N. and Udell, G.F. (1995) 'Relationship Lending and Lines of Credit in Small Firm Finance', Journal of Business, 68(3), pp. 351–381.",
        "Bertrams, R.I.V.F. (2013) Bank Guarantees in International Trade. 4th edn. The Hague: Kluwer Law International.",
        "Boot, A.W.A. (2000) 'Relationship Banking: What Do We Know?', Journal of Financial Intermediation, 9(1), pp. 7–25.",
        "Carletti, E., Leonello, A. and Marquez, R. (2023) 'Loan guarantees, bank underwriting policies and financial fragility', Journal of Financial Economics, 149(2), pp. 260–295.",
        "Davis, F.D. (1989) 'Perceived Usefulness, Perceived Ease of Use, and User Acceptance of Information Technology', MIS Quarterly, 13(3), pp. 319–340.",
        "DeYoung, R. and Roland, K.P. (2001) 'Product Mix, Revenue Mix, and Risk at Commercial Banks', Journal of Financial Intermediation, 10(2), pp. 115–144.",
        "Diamond, D.W. (1984) 'Financial Intermediation and Delegated Monitoring', Review of Economic Studies, 51(3), pp. 393–414.",
        "Hassan, A.A. et al. (2018) 'The problems and abuse of performance bond in the construction industry', IOP Conference Series: Earth and Environmental Science, 143, p. 012045.",
        "Ho Dinh Phi et al. (2023) 'Effect of Service Quality on Customer Loyalty: the Mediation of Customer Satisfaction, and Corporate Reputation in Banking Industry', Eurasian Journal of Business and Management, 11(3), pp. 145–160.",
        "International Chamber of Commerce (2010) Uniform Rules for Demand Guarantees (URDG 758). ICC Publication No. 758. Paris: ICC.",
        "Kaur, M. et al. (2021) 'The determinants of bank selection criteria of SMEs: a fuzzy analytic hierarchy approach', Journal of Science and Technology Policy Management, 12(4), pp. 580–605.",
        "Le Van Dung (2021) 'The nature of payment guarantee relationships at credit institutions', Industry and Trade Magazine, 8(April), pp. 45–52.",
        "Merton, R.C. (1974) 'On the Pricing of Corporate Debt: The Risk Structure of Interest Rates', Journal of Finance, 29(2), pp. 449–470.",
        "Narteh, B. (2013) 'SME bank selection and patronage behaviour in the Ghanaian banking industry', Management Research Review, 36(11), pp. 1061–1080.",
        "Nguyen, H. et al. (2024) 'The impact of service innovation on customer satisfaction and customer loyalty: a case in Vietnamese retail banks', Future Business Journal, 10(1), p. 14.",
        "Nguyen Thi Nhung and Nguyen Duy Phu (2015) 'Payment guarantees at Vietnamese commercial banks', Development and Integration Magazine, 25(35), pp. 62–67.",
        "Oke, A.E. (2018) 'Bonding capability of Nigerian contracting firms', Engineering, Construction and Architectural Management, 25(8), pp. 1012–1024.",
        "Parasuraman, A., Zeithaml, V.A. and Berry, L.L. (1988) 'SERVQUAL: A Multiple-Item Scale for Measuring Consumer Perceptions of Service Quality', Journal of Retailing, 64(1), pp. 12–40.",
        "Phan Thi Hang Nga et al. (2024) 'Service quality, customer satisfaction and loyalty: a case study in Vietnamese SMEs', Cogent Business & Management, 11(1), p. 2304512.",
        "Ramakrishnan, R.T.S. and Thakor, A.V. (1984) 'Information Reliability and a Theory of Financial Intermediation', Review of Economic Studies, 51(3), pp. 415–432.",
        "State Bank of Vietnam (2024) Circular No. 61/2024/TT-NHNN dated December 31, 2024, providing regulations on bank guarantees (effective April 1, 2025). Hanoi: SBV.",
        "Stiglitz, J.E. and Weiss, A. (1981) 'Credit Rationing in Markets with Imperfect Information', American Economic Review, 71(3), pp. 393–410.",
        "Turnbull, P.W. and Gibbs, M.L. (1989) 'The Selection of Banks and Banking Services among Corporate Customers in South Africa', International Journal of Bank Marketing, 7(5), pp. 36–42.",
        "Venkatesh, V. et al. (2003) 'User Acceptance of Information Technology: Toward a Unified View', MIS Quarterly, 27(3), pp. 425–478.",
        "Zeithaml, V.A. (1988) 'Consumer Perceptions of Price, Quality, and Value: A Means-End Model and Synthesis of Evidence', Journal of Marketing, 52(3), pp. 2–22.",
        "Zeithaml, V.A., Berry, L.L. and Parasuraman, A. (1996) 'The Behavioral Consequences of Service Quality', Journal of Marketing, 60(2), pp. 31–46.",
        "Zelie, E.M. (2023) 'Factors determining bank selection by micro- and small-sized enterprises: evidence from Ethiopia', International Journal of Bank Marketing, 41(5), pp. 1120–1142."
    ]

    for ref in refs:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_ref.paragraph_format.left_indent = Inches(0)
        p_ref.paragraph_format.first_line_indent = Inches(0)
        p_ref.paragraph_format.space_after = Pt(6)
        p_ref.add_run(ref).font.size = Pt(10.5)

    doc.add_page_break()

    # ==================== APPENDICES ====================
    doc.add_heading("APPENDICES (HỆ THỐNG PHỤ LỤC)", level=1)

    doc.add_heading("Appendix 1: Official 34-Item Survey Questionnaire", level=2)
    add_guide("Đính kèm toàn văn mẫu Phiếu khảo sát 34 câu hỏi chuẩn hóa Likert 1-5 đã phát cho 800 doanh nghiệp (cả bản tiếng Việt và tiếng Anh).")

    doc.add_heading("Appendix 2: Sample Demographic Characteristics Output", level=2)
    add_guide("Đính kèm bảng kết quả thống kê mô tả tần số mẫu chi tiết trích xuất từ phần mềm SPSS / Python (Frequencies & Percentages của Loại hình DN, Doanh thu, Thâm niên, Sản phẩm).")

    doc.add_heading("Appendix 3: Cronbach’s Alpha Reliability Analysis Output", level=2)
    add_guide("Đính kèm toàn bộ bảng kết quả kiểm định độ tin cậy Cronbach’s Alpha cho 7 biến độc lập và biến DEC từ phần mềm SPSS / Python.")

    doc.add_heading("Appendix 4: EFA Total Variance Explained & Rotated Component Matrix Output", level=2)
    add_guide("Đính kèm bảng KMO and Bartlett's Test, bảng Total Variance Explained và bảng Rotated Component Matrix (ma trận xoay nhân tố) đầy đủ.")

    doc.add_heading("Appendix 5: OLS Multiple Regression, VIF & Sub-group ANOVA Output", level=2)
    add_guide("Đính kèm toàn bộ bảng kết quả hồi quy OLS (Model Summary, ANOVA, Coefficients, Collinearity Statistics VIF) và các bảng kiểm định so sánh khác biệt nhóm (Independent Samples Test & One-Way ANOVA with Post-Hoc Tests).")

    # Output file path
    output_filename = "c:/Users/nguyen.tuan.minh/Desktop/DTL-Master-Project/DTL_Master_Thesis_Full_Structure_Template.docx"
    doc.save(output_filename)
    print(f"Successfully generated Full Master Thesis Structure Template at {output_filename}!")

if __name__ == "__main__":
    generate_full_thesis_template()
