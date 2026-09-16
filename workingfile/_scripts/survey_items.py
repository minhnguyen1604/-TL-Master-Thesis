# -*- coding: utf-8 -*-
"""Nguon du lieu duy nhat cho bo cau hoi - moi file khac deu doc tu day."""

SCREEN = ("S1", "Trong 12 tháng gần đây, Doanh nghiệp có phát hành thư bảo lãnh nào tại VietinBank không?",
          ["Có (đề nghị tiếp tục trả lời phiếu)", "Không (xin dừng tại đây — chân thành cảm ơn Quý Doanh nghiệp)"])

PART_I = [
 ("Q1", "OWNERSHIP", "Loại hình sở hữu của Doanh nghiệp:",
  ["Doanh nghiệp tư nhân / Công ty TNHH", "Công ty cổ phần (ngoài nhà nước)",
   "Doanh nghiệp nhà nước / có vốn nhà nước chi phối", "Doanh nghiệp có vốn đầu tư nước ngoài (FDI)",
   "Loại hình khác"]),
 ("Q2", "REVENUE", "Doanh thu năm gần nhất của Doanh nghiệp:",
  ["Dưới 20 tỷ VNĐ", "Từ 20 đến dưới 100 tỷ VNĐ", "Từ 100 đến dưới 500 tỷ VNĐ", "Từ 500 tỷ VNĐ trở lên"]),
 ("Q3", "EXPERIENCE", "Thời gian hoạt động của Doanh nghiệp:",
  ["Dưới 3 năm", "Từ 3 đến dưới 5 năm", "Từ 5 đến dưới 10 năm", "Từ 10 năm trở lên"]),
 ("Q4", "MAIN_PRODUCT", "Loại bảo lãnh Doanh nghiệp sử dụng NHIỀU NHẤT tại VietinBank (chỉ chọn 01):",
  ["Bảo lãnh dự thầu (TG)", "Bảo lãnh thực hiện hợp đồng (PG)", "Bảo lãnh tạm ứng (APG)",
   "Bảo lãnh thanh toán (BG)", "Loại khác (bảo hành, tái bảo lãnh...)"]),
 ("Q5", "NUM_BANKS", "Hiện Doanh nghiệp sử dụng dịch vụ bảo lãnh tại bao nhiêu ngân hàng?",
  ["Chỉ duy nhất VietinBank", "02 ngân hàng", "03 ngân hàng", "Từ 04 ngân hàng trở lên"]),
 ("Q6", "POSITION", "Chức danh của người trả lời phiếu:",
  ["Ban Giám đốc / CFO", "Kế toán trưởng / Trưởng phòng Tài chính",
   "Trưởng phòng Đấu thầu / Mua hàng", "Chuyên viên phụ trách bảo lãnh"]),
]

# (ma nhan to mo hinh, tieu de nhom, [(ma cau, noi dung)], co o "chua su dung" hay khong)
CONSTRUCTS = [
 ("COST_COMP", "I. Biểu phí bảo lãnh", False, [
   ("COMP1", "Mức phí phát hành bảo lãnh tại VietinBank là hợp lý so với chất lượng dịch vụ nhận được."),
   ("COMP2", "Biểu phí bảo lãnh của VietinBank có tính cạnh tranh so với các ngân hàng khác."),
   ("COMP3", "VietinBank có chính sách giảm phí cho khách hàng giao dịch thường xuyên."),
   ("COMP4", "Các khoản phí phát sinh (sửa đổi, gia hạn, tra soát) tại VietinBank ở mức chấp nhận được."),
 ]),
 ("PROC_SPEED", "II. Tốc độ xử lý hồ sơ", False, [
   ("SPEED1", "Thời gian thẩm định và phê duyệt hạn mức bảo lãnh tại VietinBank nhanh."),
   ("SPEED2", "Hồ sơ đề nghị cấp bảo lãnh tại VietinBank gọn, không yêu cầu nhiều giấy tờ."),
   ("SPEED3", "Thời gian từ khi nộp đủ hồ sơ đến khi nhận thư bảo lãnh đáp ứng kịp tiến độ hợp đồng."),
   ("SPEED4", "Việc sửa đổi hoặc gia hạn thư bảo lãnh tại VietinBank được xử lý nhanh."),
 ]),
 ("DIGITAL_CONV", "III. Bảo lãnh điện tử và kênh số eFAST", True, [
   ("DIGI1", "Doanh nghiệp nộp được đề nghị cấp bảo lãnh trực tuyến qua VietinBank eFAST."),
   ("DIGI2", "VietinBank phát hành thư bảo lãnh điện tử có chữ ký số nhanh chóng."),
   ("DIGI3", "Việc tra cứu trạng thái hồ sơ bảo lãnh trực tuyến thuận tiện."),
   ("DIGI4", "Bên thụ hưởng chấp nhận thư bảo lãnh điện tử của VietinBank mà không yêu cầu bản giấy."),
 ]),
 ("BANK_REP", "IV. Uy tín ngân hàng", False, [
   ("REPU1", "VietinBank có uy tín thuộc nhóm dẫn đầu thị trường ngân hàng Việt Nam."),
   ("REPU2", "Thư bảo lãnh của VietinBank được chủ đầu tư và bên mời thầu chấp nhận rộng rãi."),
   ("REPU3", "Năng lực tài chính của VietinBank tạo thêm uy tín cho Doanh nghiệp khi dự thầu."),
   ("REPU4", "Doanh nghiệp yên tâm về khả năng VietinBank thực hiện nghĩa vụ khi thư bảo lãnh bị yêu cầu thanh toán."),
 ]),
 ("RELATIONSHIP", "V. Quan hệ giao dịch và hạn mức", False, [
   ("RELA1", "Quan hệ tín dụng lâu năm với VietinBank giúp Doanh nghiệp thuận lợi hơn khi đề nghị cấp bảo lãnh."),
   ("RELA2", "Việc sử dụng các dịch vụ khác tại VietinBank mang lại ưu đãi cho hoạt động bảo lãnh."),
   ("RELA3", "VietinBank cấp hạn mức bảo lãnh phù hợp với nhu cầu của Doanh nghiệp."),
   ("RELA4", "VietinBank điều chỉnh hạn mức kịp thời khi nhu cầu của Doanh nghiệp thay đổi."),
 ]),
 ("STAFF_QUAL", "VI. Năng lực cán bộ", False, [
   ("STAFF1", "Cán bộ VietinBank có chuyên môn vững về nghiệp vụ bảo lãnh."),
   ("STAFF2", "Cán bộ VietinBank tư vấn được nội dung thư bảo lãnh phù hợp quy định pháp luật hiện hành."),
   ("STAFF3", "Cán bộ VietinBank cảnh báo được các rủi ro trong điều khoản bảo lãnh trước khi phát hành."),
   ("STAFF4", "Cán bộ VietinBank phản hồi kịp thời khi Doanh nghiệp có vướng mắc."),
 ]),
 ("COLL_POLICY", "VII. Ký quỹ và tài sản bảo đảm", False, [
   ("COLL1", "VietinBank áp dụng tỷ lệ ký quỹ phù hợp với năng lực của Doanh nghiệp."),
   ("COLL2", "VietinBank có chính sách giảm hoặc miễn ký quỹ cho khách hàng đủ điều kiện."),
   ("COLL3", "VietinBank chấp nhận nhiều loại tài sản bảo đảm khác nhau."),
   ("COLL4", "Thủ tục định giá và nhận tài sản bảo đảm tại VietinBank nhanh gọn."),
 ]),
]

DEC = ("DEC", "VIII. Mức độ ưu tiên lựa chọn VietinBank", [
   ("DEC1", "Khi phát sinh nhu cầu bảo lãnh, Doanh nghiệp ưu tiên lựa chọn VietinBank trước các ngân hàng khác."),
   ("DEC2", "Doanh nghiệp dành phần lớn giá trị bảo lãnh của mình cho VietinBank thay vì các ngân hàng khác."),
   ("DEC3", "Trong các gói thầu và hợp đồng sắp tới, Doanh nghiệp dự định tiếp tục chọn VietinBank."),
   ("DEC4", "Doanh nghiệp sẵn sàng giới thiệu VietinBank cho đối tác như ngân hàng phát hành bảo lãnh tốt nhất."),
])

VALIDATION = ("V1", "WALLET_SHARE",
  "Ước tính trong 12 tháng qua, VietinBank chiếm khoảng bao nhiêu phần trăm tổng giá trị bảo lãnh của Doanh nghiệp?",
  ["Dưới 25%", "Từ 25% đến 50%", "Trên 50% đến 75%", "Trên 75%"])

N_ITEMS = sum(len(c[3]) for c in CONSTRUCTS) + len(DEC[2])
if __name__ == "__main__":
    print("So bien doc lap:", len(CONSTRUCTS))
    print("So cau moi bien :", set(len(c[3]) for c in CONSTRUCTS))
    print("Tong cau Likert :", N_ITEMS)
