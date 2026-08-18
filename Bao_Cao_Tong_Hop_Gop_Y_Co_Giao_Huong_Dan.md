# CHƯƠNG TRÌNH CAO HỌC KINH TẾ PHÁT TRIỂN VIỆT NAM - HÀ LAN (MDE)
## TRƯỜNG ĐẠI HỌC KINH TẾ QUỐC DÂN

---

# BÁO CÁO TIẾP THU, GIẢI TRÌNH VÀ TỔNG HỢP NỘI DUNG CHỈNH SỬA THEO GÓP Ý CỦA GIẢNG VIÊN HƯỚNG DẪN

**Đề tài:** *Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at VietinBank*  
**Học viên:** Đặng Tú Linh – Lớp MDE Khóa 31  
**Giảng viên hướng dẫn:** TS. Hoàng Thị Thúy Nga  

---

## PHẦN I: TỔNG HỢP VÀ ĐÁNH GIÁ Ý KIẾN CỦA GIẢNG VIÊN HƯỚNG DẪN

### 1.1. Các nội dung Giảng viên hướng dẫn ĐỒNG TÌNH (Phê duyệt 100%)
* **Về Phương pháp nghiên cứu:** Đồng ý sử dụng mô hình Hồi quy tuyến tính đa biến (Multiple Linear Regression – OLS) để ước lượng tác động của các nhân tố đến quyết định lựa chọn dịch vụ bảo lãnh.
* **Về Quy trình phân tích định lượng 7 bước:** Phê duyệt 100% quy trình tuần tự: Thống kê mô tả $\rightarrow$ Đánh giá độ tin cậy Cronbach’s Alpha $\rightarrow$ Khám phá nhân tố EFA $\rightarrow$ Trích xuất biến đại diện (Factor Scores) $\rightarrow$ Phân tích tương quan Pearson $\rightarrow$ Kiểm tra đa cộng tuyến (VIF) $\rightarrow$ Hồi quy OLS và kiểm định giả thuyết nghiên cứu.
* **Về Bối cảnh và Nguồn dữ liệu:** Nhất trí với việc khai thác bộ dữ liệu khảo sát thực tế quy mô $n = 800$ khách hàng doanh nghiệp trên 155 chi nhánh toàn quốc của Ngân hàng TMCP Công thương Việt Nam (VietinBank).

### 1.2. Các nội dung Giảng viên hướng dẫn YÊU CẦU ĐIỀU CHỈNH & Hướng xử lý

| STT | Ý kiến Góp ý của Giảng viên | Phân tích Học thuật & Nguyên nhân | Phương án Tiếp thu & Xử lý |
| :---: | :--- | :--- | :--- |
| **1** | **Rút gọn số lượng biến độc lập:**<br>Nên chọn khoảng 6–7 yếu tố đặc trưng, tránh dàn trải. Lưu ý Bank Reputation không phải Reliability trong SERVQUAL. | Bản draft cũ có 10 biến làm mô hình bị phân tán, giảm bậc tự do và dễ dính đa cộng tuyến. | **Gom nhóm và rút gọn xuống đúng 7 biến độc lập trọng tâm** (nằm đúng khung 6-7 biến), giữ lại biến Chuyển đổi số eFAST và tách riêng Uy tín thương hiệu (`BANK_REP`). |
| **2** | **Điều chỉnh biến COST & Chiều tác động (+):**<br>Thang đo đo mức độ hợp lý/cạnh tranh của phí nhưng giả thuyết lại để dấu âm (-). Cần đổi tên thành Price Competitiveness và kỳ vọng tác động dương (+). | Điểm Likert 5 thể hiện 'Phí rất hợp lý và cạnh tranh'. Điểm càng cao thì càng hấp dẫn doanh nghiệp chọn VietinBank nên dấu kỳ vọng bắt buộc phải là DƯƠNG (+). | Đổi tên biến thành **`COST_COMP` (Price Competitiveness & Fee Policy)** và điều chỉnh chiều tác động trong giả thuyết H1 và mô hình sang **DƯƠNG (+)**. |
| **3** | **Rà soát biến phụ thuộc DEC (Selection Decision):**<br>Cần đo mức độ ưu tiên/lựa chọn VietinBank so với các ngân hàng khác để khớp tên đề tài. | Bản cũ bị lẫn lộn giữa hài lòng chung (Satisfaction) và quyết định lựa chọn cạnh tranh. | **Chuẩn hóa 4 câu hỏi quan sát của DEC** tập trung đo lường: Ưu tiên lựa chọn số 1, Phân bổ tỷ trọng bảo lãnh lớn tại VietinBank, Tiếp tục lựa chọn và Sẵn sàng giới thiệu đối tác. |
| **4** | **Thang đo phải bám sát Bảng hỏi gốc (n = 800):**<br>Không tự xây dựng thang đo mới rồi gán ép vào bộ dữ liệu đã thu thập. | Đảm bảo tính chân thực: Mọi biến trong đề cương phải có cột dữ liệu tương ứng trong file khảo sát 34 câu. | **Ánh xạ 100% hệ thống 7 biến độc lập và biến phụ thuộc** khớp hoàn toàn với 34 câu hỏi thực tế trong Phiếu khảo sát VietinBank. |
| **5** | **Chuẩn hóa Mục tiêu & Câu hỏi nghiên cứu:**<br>Viết lại mục tiêu tổng quát và 4 mục tiêu cụ thể tiếng Anh chuẩn NEU MDE, bổ sung phân tích khác biệt nhóm. | Nâng cao chuẩn mực học thuật quốc tế và mở rộng góc nhìn quản trị thông qua kiểm định so sánh nhóm doanh nghiệp. | **Cập nhật nguyên văn 100% bộ Mục tiêu tiếng Anh của Giảng viên** vào Mục 1.2 & 1.3; bổ sung kiểm định ANOVA & t-test ở Mục tiêu 3. |

### 1.3. Giải trình Cơ sở Học thuật & Thực tiễn về việc GIỮ LẠI biến Chuyển đổi số eFAST (DIGITAL_CONV)
Học viên xin phép được giải trình làm rõ lý do vì sao trong quá trình rút gọn mô hình từ 10 biến xuống 7 biến, **biến Chuyển đổi số eFAST (`DIGITAL_CONV`) vẫn được giữ lại làm một biến độc lập trọng tâm** trong mô hình:
1. **Tính thời sự và xu hướng chuyển đổi số tất yếu của ngành Ngân hàng:**  
   Chuyển đổi số là trọng tâm chiến lược của Chính phủ và Ngân hàng TMCP Công thương Việt Nam. Theo Thông tư số 61/2024/TT-NHNN (có hiệu lực từ ngày 01/4/2025 thay thế Thông tư 11/2022/TT-NHNN), nghiệp vụ Bảo lãnh Điện tử (e-guarantees) đã được luật hóa toàn diện. Hiện nay, tính năng nộp đề nghị online, thẩm định số hóa STP và cấp Thư bảo lãnh điện tử 24/7 qua nền tảng VietinBank eFAST đang là "vũ khí cạnh tranh sắc bén nhất" giúp VietinBank thu hút các doanh nghiệp xây dựng, nhà thầu và doanh nghiệp thương mại.
2. **Cơ sở lý thuyết vững chắc (Technology Acceptance Model - TAM):**  
   Biến `DIGITAL_CONV` được bảo chứng vững chắc bởi Mô hình Chấp nhận Công nghệ TAM (Davis, 1989; Venkatesh et al., 2003). Tính hữu ích cảm nhận (Perceived Usefulness) và Tính dễ sử dụng (Perceived Ease-of-Use) của giao dịch số hóa eFAST tác động trực tiếp, mạnh mẽ đến thái độ và quyết định lựa chọn ngân hàng phục vụ của khách hàng doanh nghiệp hiện đại.
3. **Hoàn toàn phù hợp với khuyến nghị quy mô 6–7 biến của Giảng viên:**  
   Sau khi rút gọn và tinh giản các biến phụ, việc giữ lại biến `DIGITAL_CONV` cấu thành đúng hệ thống 7 biến độc lập (nằm trọn vẹn trong khoảng 6–7 biến cô khuyên). Mô hình không bị dàn trải nhưng lại có được một điểm nhấn công nghệ rất sáng, giúp bài luận văn vừa đạt chuẩn học thuật vừa mang hơi thở thực tiễn sinh động.
4. **Dữ liệu có sẵn và đo lường độc lập trong bảng khảo sát gốc ($n = 800$):**  
   Các câu hỏi về tính năng nộp hồ sơ eFAST 24/7, tốc độ cấp mã bảo lãnh điện tử và tra cứu trực tuyến đã được thu thập đầy đủ trong tập dữ liệu 800 doanh nghiệp, đảm bảo tính khả thi tuyệt đối khi phân tích EFA và hồi quy OLS.

---

## PHẦN II: TỔNG HỢP CÁC NỘI DUNG ĐÃ ĐIỀU CHỈNH TRONG FILE THESIS DESIGN

Toàn bộ các yêu cầu chỉnh sửa trên đã được cập nhật hoàn chỉnh trong file Đề cương Luận văn chính thức: `DTL_Thesis_Design_NEU_MDE_Final.docx`. Cụ thể từng chương mục như sau:

### 2.1. Cập nhật Chương 1: Đặt vấn đề & Mục tiêu Nghiên cứu
* **Tiêu đề tiếng Anh:** `FACTORS AFFECTING CORPORATE CUSTOMERS’ DECISION TO CHOOSE BANK GUARANTEE SERVICES AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)`
* **Mục tiêu tổng quát (General Objective):** Cập nhật nguyên văn chuẩn tiếng Anh cô giao:  
  *“The general objective of this study is to identify and assess the factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, and to propose managerial recommendations for improving VietinBank’s attractiveness and competitiveness in the corporate bank guarantee market.”*
* **4 Mục tiêu cụ thể (Specific Objectives 1-4):** Cập nhật nguyên văn 4 mục tiêu tiếng Anh, trong đó Mục tiêu 3 tích hợp kiểm định khác biệt nhóm (Sub-group Analysis) theo loại hình doanh nghiệp, quy mô, thâm niên và loại bảo lãnh.

### 2.2. Cập nhật Chương 2: Khung Lý thuyết Nền tảng & Tổng quan
Bổ sung và phân định rõ ràng 5 khung lý thuyết bảo chứng cho mô hình 7 biến:
1. *Lý thuyết Trung gian Tài chính & Giám sát Ủy thác (Diamond, 1984):* Cơ sở của giá trị bảo chứng thương hiệu Big4.
2. *Lý thuyết Định giá Rủi ro Tín dụng (Merton, 1974; Stiglitz & Weiss, 1981):* Cơ sở của tính cạnh tranh biểu phí (`COST_COMP`) và ký quỹ (`COLL_POLICY`).
3. *Mô hình SERVQUAL (Parasuraman et al., 1988):* Cơ sở đo lường năng lực cán bộ (`STAFF_QUAL`), tốc độ (`PROC_SPEED`), và tách riêng Uy tín (`BANK_REP`).
4. *Lý thuyết Ngân hàng Quan hệ (Boot, 2000; Berger & Udell, 1995):* Cơ sở của biến Quan hệ tín dụng & Hạn mức (`RELATIONSHIP`).
5. *Mô hình Chấp nhận Công nghệ TAM (Davis, 1989):* Cơ sở giải thích động lực lựa chọn qua kênh số hóa eFAST (`DIGITAL_CONV`).

### 2.3. Cập nhật Chương 3: Mô hình Kinh tế lượng & Hệ thống Biến Thang đo
Phương trình Hồi quy Tuyến tính Đa biến (OLS) chính thức được thiết lập:

$$\text{DEC} = \beta_0 + \beta_1\text{COST\_COMP} + \beta_2\text{PROC\_SPEED} + \beta_3\text{DIGITAL\_CONV} + \beta_4\text{BANK\_REP} + \beta_5\text{RELATIONSHIP} + \beta_6\text{STAFF\_QUAL} + \beta_7\text{COLL\_POLICY} + \varepsilon$$

Bảng hệ thống 7 biến độc lập và 1 biến phụ thuộc đã được khớp 100% với 34 câu hỏi trong bảng khảo sát gốc:

| Mã biến | Tên biến | Nội dung thang đo (Khớp 34 câu hỏi gốc) | Loại biến | Dấu |
| :---: | :--- | :--- | :---: | :---: |
| **DEC** | **Selection Decision** | Quyết định và mức độ ưu tiên chọn VietinBank so với đối thủ | **Phụ thuộc ($Y$)** | N/A |
| **COST_COMP** | **Price Competitiveness** | Tính hợp lý, cạnh tranh của biểu phí và chiết khấu phí | **Độc lập ($X_1$)** | **+** |
| **PROC_SPEED** | **Processing Speed** | Thời gian thẩm định nhanh, thủ tục phát hành đơn giản | **Độc lập ($X_2$)** | **+** |
| **DIGITAL_CONV** | **Digital eFAST Convenience** | Nộp đề nghị online, cấp e-guarantee trực tuyến 24/7 | **Độc lập ($X_3$)** | **+** |
| **BANK_REP** | **Bank Reputation** | Uy tín thương hiệu Big4, bảo lãnh được chấp nhận 100% | **Độc lập ($X_4$)** | **+** |
| **RELATIONSHIP** | **Relationship & Limits** | Hạn mức bảo lãnh linh hoạt, chính sách chăm sóc VIP | **Độc lập ($X_5$)** | **+** |
| **STAFF_QUAL** | **Staff Professionalism** | Trình độ cán bộ RM, am hiểu Luật Đấu thầu / TT 61 | **Độc lập ($X_6$)** | **+** |
| **COLL_POLICY** | **Collateral Flexibility** | Tỷ lệ ký quỹ linh hoạt, đa dạng tài sản thế chấp | **Độc lập ($X_7$)** | **+** |

### 2.4. Cập nhật Chương 4: Kết luận & Khuyến nghị Quản trị
Thiết lập 4 nhóm giải pháp thực tiễn kết nối trực tiếp với 4 mục tiêu nghiên cứu:
* *Giải pháp 1 (Theo Objective 1):* Tập trung nguồn lực nâng cao 7 giá trị cốt lõi đã được định vị.
* *Giải pháp 2 (Theo Objective 2):* Tối ưu hóa biểu phí cạnh tranh và rút ngắn thời gian xử lý phát hành bảo lãnh.
* *Giải pháp 3 (Theo Objective 3):* Thiết kế chính sách sản phẩm may đo riêng theo phân khúc doanh nghiệp Nhà nước, Tư nhân/SME và FDI.
* *Giải pháp 4 (Theo Objective 4):* Đẩy mạnh kênh số hóa VietinBank eFAST và tư vấn pháp lý chuyên sâu.
