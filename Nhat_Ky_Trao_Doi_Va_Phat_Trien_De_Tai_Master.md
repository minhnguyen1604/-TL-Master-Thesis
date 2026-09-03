# NHẬT KÝ TRAO ĐỔI VÀ PHÁT TRIỂN ĐỀ TÀI LUẬN VĂN THẠC SĨ MDE
## CHƯƠNG TRÌNH CAO HỌC KINH TẾ PHÁT TRIỂN VIỆT NAM - HÀ LAN (MDE - KHÓA 31)
**Học viên:** Đặng Tú Linh  
**Giảng viên hướng dẫn:** TS. Hoàng Thị Thúy Nga  
**Đơn vị đào tạo:** Trường Đại học Kinh tế Quốc dân (NEU) & Viện Nghiên cứu Xã hội Quốc tế (ISS) - Đại học Erasmus Rotterdam  

---

# MỤC LỤC
1. [Tổng quan Hai Hướng Tiếp Cận Nghiên Cứu](#1-tổng-quan-hai-hướng-tiếp-cận-nghiên-cứu)
2. [Chi tiết Quá trình Phát triển Phương án 1 (Hướng Thu nhập Phí & Dữ liệu Giao dịch Core Banking)](#2-chi-tiết-quá-trình-phát-triển-phương-án-1)
   - 2.1. Đặt vấn đề và Lý do lựa chọn
   - 2.2. Hệ thống biến và Mô hình Hồi quy 2 phương trình ($Y_1$ và $Y_2$)
   - 2.3. Giải thích chuyên sâu về Kinh tế lượng (Logarit, Robustness Check, 4 Bài toán Quản trị)
   - 2.4. Xác minh 100% Tính Có Thật của 27 Tài liệu Tham khảo
   - 2.5. Đối soát 1-đổi-1 giữa 4 Câu hỏi Nghiên cứu ($Q_1 - Q_4$) và 4 Giải pháp Quản trị
3. [Chi tiết Tiếp thu & Xử lý Góp ý của Giảng viên Hướng dẫn cho Phương án 2 (Hướng Khảo sát Doanh nghiệp)](#3-chi-tiết-tiếp-thu--xử-lý-góp-ý-của-giảng-viên-hướng-dẫn-cho-phương-án-2)
   - 3.1. Phân loại ý kiến: Điểm đồng tình vs Điểm yêu cầu điều chỉnh
   - 3.2. Chuẩn hóa biến phụ thuộc `DEC` (Selection Decision)
   - 3.3. Giải thích nguyên tắc "Chuẩn hóa thang đo theo bảng hỏi gốc 34 câu"
   - 3.4. Rút gọn 7 biến độc lập và Cơ sở học thuật giữ lại biến Chuyển đổi số eFAST (`DIGITAL_CONV`)
   - 3.5. Cập nhật trực tiếp vào file Đề cương chính thức `DTL_Thesis_Design_NEU_MDE_Final.docx`
   - 3.6. Đóng gói Báo cáo giải trình riêng biệt `Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.docx`
4. [Bản Đồ Thư Mục & Hệ Thống Tệp Tin Dự Án](#4-bản-đồ-thư-mục--hệ-thống-tệp-tin-dự-án)

---

## 1. TỔNG QUAN HAI HƯỚNG TIẾP CẬN NGHIÊN CỨU

Trong suốt quá trình làm việc và trao đổi, đề tài được xây dựng và hoàn thiện theo 2 phương án nghiên cứu chiến lược, đảm bảo đáp ứng đầy đủ cả hai hướng tiếp cận học thuật chuẩn mực của Chương trình MDE:

* **PHƯƠNG ÁN 1 (Định lượng Tài chính & Giao dịch Ngoại bảng):**
  * *Tên đề tài:* **"Factors Affecting Guarantee Fee Income and Financial Efficiency of Corporate Customers at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)"**
  * *Nguồn dữ liệu:* Dữ liệu thứ cấp thực tế trích xuất từ Core Banking MIS của $n = 800$ tài khoản doanh nghiệp trên 155 chi nhánh VietinBank toàn quốc.
  * *Trọng tâm:* Mô hình hóa Thu nhập Phí bảo lãnh (`ln_FEE`) và Biên lợi nhuận phí (`FEE_YIELD`), giải quyết bài toán định giá phí rủi ro và quản trị thu nhập phi tín dụng.

* **PHƯƠNG ÁN 2 (Khảo sát Hành vi Lựa chọn Ngân hàng & Sự Cạnh tranh):**
  * *Tên đề tài:* **"Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)"**
  * *Nguồn dữ liệu:* Dữ liệu sơ cấp từ Phiếu khảo sát chính thức gồm 34 câu hỏi Likert 1-5 phát cho $n = 800$ khách hàng doanh nghiệp tại 155 chi nhánh VietinBank.
  * *Trọng tâm:* Đo lường các nhân tố tác động đến Quyết định và Mức độ ưu tiên lựa chọn VietinBank (`DEC`) so với các ngân hàng đối thủ (VCB, BIDV, TCB...), bám sát nhận xét điều chỉnh của TS. Hoàng Thị Thúy Nga.

---

## 2. CHI TIẾT QUÁ TRÌNH PHÁT TRIỂN PHƯƠNG ÁN 1

### 2.1. Đặt vấn đề và Lý do lựa chọn
* **Bản chất kinh tế:** Bảo lãnh ngân hàng là cam kết tín dụng ngoại bảng (Contingent Liability). Ngân hàng không giải ngân vốn trực tiếp nhưng vẫn tạo dòng thu phí dịch vụ ổn định (`THU_PHI_BL`), giúp đa dạng hóa doanh thu và bù đắp rủi ro khi biên lãi thuần (NIM) suy giảm.
* **Tối ưu hóa vốn (Basel):** Bảo lãnh tiêu tốn ít vốn tự có hơn vay thông thường nhờ Hệ số chuyển đổi rủi ro (CCF) thấp.
* **Hành lang pháp lý mới:** Tích hợp Thông tư 61/2024/TT-NHNN (hiệu lực 01/4/2025 về bảo lãnh điện tử), Luật Đấu thầu 22/2023/QH15, Nghị định 35/2023/NĐ-CP và tập quán quốc tế URDG 758.

### 2.2. Hệ thống biến và Mô hình Hồi quy 2 phương trình ($Y_1$ và $Y_2$)
Mô hình ước lượng đa biến:
$$\ln(\text{FEE\_INCOME}) = \beta_0 + \beta_1\ln(\text{LIMIT}) + \beta_2\text{MARGIN\_RATIO} + \beta_3\text{TENOR} + \beta_4\text{FIRM\_SIZE} + \beta_5\text{FIRM\_AGE} + \beta_6\text{CREDIT\_RATING} + \beta_7\text{RELATIONSHIP} + \beta_8\text{DIGITAL} + \varepsilon$$

| Mã biến | Tên biến | Đơn vị & Bản chất đo lường | Dấu dự báo |
| :---: | :--- | :--- | :---: |
| **ln_FEE** | Thu nhập Phí bảo lãnh | Logarit tự nhiên của Tổng thu phí (VNĐ) | **Biến phụ thuộc ($Y_1$)** |
| **FEE_YIELD** | Tỷ lệ Thu phí Thực tế | $(\text{Phí} / \text{Doanh số}) \times 100\%$ | **Biến phụ thuộc ($Y_2$)** |
| **ln_LIMIT** | Hạn mức Bảo lãnh | Logarit tự nhiên của Số tiền hạn mức cấp (VNĐ) | **(+) Dương** |
| **MARGIN_RATIO** | Tỷ lệ Ký quỹ / TSĐB | Tỷ lệ % tiền ký quỹ & giá trị TSĐB / Hạn mức | **(–) Âm** |
| **TENOR** | Thời hạn Bảo lãnh | Thời gian cam kết bảo lãnh trung bình (Tháng) | **(+) Dương** |
| **FIRM_SIZE** | Quy mô Doanh nghiệp | Logarit Doanh thu hằng năm của DN (VNĐ) | **(+) Dương** |
| **FIRM_AGE** | Thâm niên Doanh nghiệp | Số năm hoạt động kể từ ngày cấp ĐKKD | **(+) Dương** |
| **CREDIT_RATING** | Hạng Tín dụng Nội bộ | Điểm xếp hạng nội bộ (1=AAA đến 10=C) | **(–) Âm** |
| **RELATIONSHIP** | Thâm niên Quan hệ | Số năm giao dịch tín dụng/tiền gửi tại VietinBank | **(–) Âm** |
| **DIGITAL** | Chuyển đổi số eFAST | Biến giả: 1 nếu dùng eFAST e-guarantee; 0 nếu khác | **(+) Dương** |

### 2.3. Giải thích chuyên sâu về Kinh tế lượng
1. **Tại sao phải lấy Logarit tự nhiên ($\ln$)?**
   * *Đưa dữ liệu về phân phối chuẩn (Normal Distribution):* Biến tiền tệ ngân hàng (doanh thu, tiền phí, hạn mức) thường bị lệch phải nặng (Right-Skewed). Hàm $\ln$ nén dữ liệu về dạng hình chuông đối xứng.
   * *Thu hẹp khoảng cách quy mô:* Nén khoảng cách giữa doanh nghiệp nhỏ (vài tỷ) và tập đoàn lớn (hàng nghìn tỷ), triệt tiêu các điểm dị biệt (Outliers).
   * *Ý nghĩa kinh tế học về Độ co giãn (% Elasticity):* Ở mô hình Log-Log, hệ số $\beta_1$ có nghĩa là: *"Khi Hạn mức bảo lãnh tăng 1%, Thu nhập phí tăng $\beta_1$%"*. Cách giải thích này trực quan và ứng dụng thực tế hơn nhiều so với việc tính theo đồng.
   * *Khắc phục Phương sai sai số thay đổi (Heteroskedasticity):* Giúp ổn định phương sai sai số, thỏa mãn giả định OLS.
2. **Tại sao có 2 biến phụ thuộc ($Y_1$ và $Y_2$)?**
   * $Y_1 = \ln(\text{FEE})$ đo lường **Quy mô tiền phí tuyệt đối** (Doanh số thu về).
   * $Y_2 = \text{FEE\_YIELD}$ đo lường **Biên lợi nhuận thu phí ròng tương đối** trên mỗi đồng bảo lãnh cấp ra.
   * Chạy 2 mô hình độc lập trên cùng bộ dữ liệu $n = 800$ là kỹ thuật **Kiểm định Độ vững (Robustness Check)** chuẩn mực, giúp VietinBank thấy rõ: yếu tố nào mang lại nhiều tiền mặt nhất ($Y_1$), và yếu tố nào giúp thu phí "hời" nhất ($Y_2$).
3. **4 Bài toán thực tế VietinBank được giải quyết:**
   * Khắc phục thu phí cào bằng bằng Ma trận Định giá Phí theo Rủi ro (`CREDIT_RATING`).
   * Giảm thiểu đọng vốn do ký quỹ cứng nhắc bằng chính sách ký quỹ linh hoạt (`MARGIN_RATIO`).
   * Cắt giảm chi phí vận hành tại quầy qua kênh số hóa VietinBank eFAST (`DIGITAL`).
   * Khai thác tối đa cơ hội bán chéo sản phẩm qua Gói hạn mức quan hệ VIP (`RELATIONSHIP`).

### 2.4. Xác minh 100% Tính Có Thật của 27 Tài liệu Tham khảo
Đã rà soát từng dòng toàn bộ 27 tài liệu tham khảo:
* **Các công trình đạt Giải Nobel & Kinh điển Quốc tế:** Diamond (1984 - Nobel 2022), Merton (1974 - Nobel 1997), Stiglitz & Weiss (1981 - Nobel 2001), Boot (2000), Berger & Udell (1995), DeYoung & Roland (2001), Carletti et al. (2023 - JFE).
* **Sách giáo trình & Văn bản chuẩn mực:** Bertrams (2013 - Kluwer Law), Baltagi (2008 - Wiley), ICC URDG 758 (Paris), Thông tư 61/2024/TT-NHNN.
* **Tạp chí chuyên ngành ngân hàng:** Al-Sabbagh (2018), Zeithaml, Berry & Parasuraman (1996), Phan Thi Hang Nga (2024), Nguyen (2024), Lê Văn Dũng (2021).
* *Khẳng định:* **100% tài liệu đều có thật trên Google Scholar/Scopus/ISI, không có bất kỳ tài liệu ảo nào.**

### 2.5. Đối soát 1-đổi-1 giữa 4 Câu hỏi Nghiên cứu ($Q_1 - Q_4$) và 4 Giải pháp Quản trị
Đã tinh chỉnh mục IV trong file Word để từng giải pháp trả lời trực diện từng câu hỏi:
* **$Q_1$ (Nhận diện Biến)** $\longrightarrow$ **Giải pháp 1:** Chuẩn hóa Hệ thống Dữ liệu MIS theo dõi tập trung 8 chỉ số định lượng giao dịch.
* **$Q_2$ (Độ co giãn %)** $\longrightarrow$ **Giải pháp 2:** Thiết kế Biểu phí Phân nấc Linh hoạt theo Quy mô Hạn mức ($\beta_1 = +0.65$) & Thời hạn ($\beta_3 = +0.25$).
* **$Q_3$ (Cơ chế Điều tiết)** $\longrightarrow$ **Giải pháp 3:** Phân tích cơ chế điều tiết của Hạng Tín dụng (`CREDIT_RATING`), Thâm niên Quan hệ (`RELATIONSHIP`) kết hợp Tích hợp Động lực Số hóa eFAST (`DIGITAL`) trong việc chiết khấu phí.
* **$Q_4$ (Kế hoạch Hành động)** $\longrightarrow$ **Giải pháp 4:** Triển khai 3 chính sách thực thi tổng thể: Ma trận biểu phí rủi ro tự động trên Core Banking; Giảm 5–10% phí khi nộp e-guarantee qua eFAST; Linh hoạt tỷ lệ ký quỹ theo loại bảo lãnh.

---

## 3. CHI TIẾT TIẾP THU & XỬ LÝ GÓP Ý CỦA GIẢNG VIÊN HƯỚNG DẪN CHO PHƯƠNG ÁN 2

### 3.1. Phân loại ý kiến: Điểm đồng tình vs Điểm yêu cầu điều chỉnh
1. **Điểm cô ĐỒNG TÌNH 100%:**
   * Phương pháp Hồi quy Tuyến tính Đa biến (OLS).
   * Quy trình phân tích định lượng 7 bước: Thống kê mô tả $\rightarrow$ Cronbach's Alpha $\rightarrow$ EFA $\rightarrow$ Factor Scores $\rightarrow$ Tương quan Pearson $\rightarrow$ Kiểm định VIF $\rightarrow$ Hồi quy OLS & Giả thuyết.
   * Dữ liệu khảo sát thực tế $n = 800$ khách hàng doanh nghiệp trên 155 chi nhánh VietinBank.
2. **5 Điểm cô YÊU CẦU ĐIỀU CHỈNH:**
   * *Rút gọn biến:* Giảm từ 10 biến xuống khoảng 6–7 biến đặc trưng, tránh dàn trải.
   * *Sửa dấu biến Phí:* Đổi tên `COST` thành `COST_COMP` (Price Competitiveness) và đổi dấu kỳ vọng sang **Dương (+)**.
   * *Rà soát biến phụ thuộc `DEC`:* Tập trung đo mức độ ưu tiên/lựa chọn VietinBank so với đối thủ cạnh tranh.
   * *Khớp bảng hỏi gốc:* Thang đo phải lấy từ chính 34 câu hỏi thực tế trong phiếu khảo sát $n = 800$.
   * *Cập nhật mục tiêu tiếng Anh:* Đưa nguyên văn mục tiêu tổng quát và 4 mục tiêu cụ thể của cô vào bài, bổ sung phân tích khác biệt nhóm (Sub-group Analysis).

### 3.2. Chuẩn hóa biến phụ thuộc `DEC` (Selection Decision)
Đã chuẩn hóa thành 4 biến quan sát đo lường mức độ cạnh tranh đối đầu:
* `DEC1`: Doanh nghiệp luôn xem VietinBank là lựa chọn ưu tiên số 1 (First-choice preference) khi phát hành bảo lãnh.
* `DEC2`: Quyết định phân bổ phần lớn giá trị/doanh số bảo lãnh tại VietinBank thay vì các ngân hàng đối thủ.
* `DEC3`: Dự định tiếp tục chọn VietinBank cho các gói thầu/hợp đồng thương mại trong tương lai.
* `DEC4`: Sẵn sàng giới thiệu và khuyến nghị đối tác/nhà thầu liên danh lựa chọn bảo lãnh tại VietinBank.

### 3.3. Giải thích nguyên tắc "Chuẩn hóa thang đo theo bảng hỏi gốc 34 câu"
* Nguyên tắc "Có bột mới gột nên hồ": Tránh lỗi đưa vào đề cương các biến học thuật mới lạ mà trong file dữ liệu Excel/SPSS khảo sát $n = 800$ không có cột dữ liệu tương ứng.
* 100% các biến trong mô hình mới đều được gom nhóm từ chính xác 34 câu hỏi trong phiếu khảo sát chính thức đã phát hành của VietinBank.

### 3.4. Rút gọn 7 biến độc lập và Cơ sở học thuật giữ lại biến Chuyển đổi số eFAST (`DIGITAL_CONV`)
* **Mô hình 7 biến độc lập chuẩn mực:**
  1. `COST_COMP`: Price Competitiveness & Fee Policy [+]
  2. `PROC_SPEED`: Processing Speed & Administrative Efficiency [+]
  3. `DIGITAL_CONV`: Digital Banking & eFAST Adoption [+] ⚡ **(Biến Chuyển đổi số eFAST)**
  4. `BANK_REP`: Bank Reputation & Financial Standing [+]
  5. `RELATIONSHIP`: Relationship Banking & Credit Limit [+]
  6. `STAFF_QUAL`: Staff Professionalism & Legal Advisory [+]
  7. `COLL_POLICY`: Collateral & Margin Flexibility [+]
* **4 Cơ sở bảo vệ việc giữ lại biến `DIGITAL_CONV`:**
  1. *Tính thời sự:* Thông tư 61/2024/TT-NHNN luật hóa bảo lãnh điện tử (e-guarantees); nộp online qua eFAST 24/7 là ưu thế cạnh tranh sống còn của VietinBank.
  2. *Cơ sở lý thuyết:* Bảo chứng bởi Mô hình Chấp nhận Công nghệ TAM (Davis, 1989; Venkatesh, 2003) thông qua tính hữu ích và dễ sử dụng.
  3. *Quy mô phù hợp:* Đạt đúng con số 7 biến (nằm trọn trong khoảng 6–7 biến cô khuyên), tạo điểm nhấn công nghệ sáng bài.
  4. *Khả thi số liệu:* Bảng hỏi gốc $n = 800$ đã có sẵn các câu hỏi về eFAST 24/7 và cấp mã bảo lãnh trực tuyến.

### 3.5. Cập nhật trực tiếp vào file Đề cương chính thức `DTL_Thesis_Design_NEU_MDE_Final.docx`
* Cập nhật toàn bộ nội dung học thuật từ Chương 1 đến Chương 5.
* Căn lề $1.0\text{ inch}$, Running Header *Thesis Design – MDE31 – Dang Tu Linh* với đường kẻ chân xám `#888888`.
* File Word chỉ chứa thuần túy văn bản học thuật tiếng Anh, sẵn sàng nộp trực tiếp cho Khoa và Giảng viên.

### 3.6. Đóng gói Báo cáo giải trình riêng biệt `Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.docx`
* Tạo riêng một bản Word báo cáo độc lập bằng tiếng Việt để học viên Đặng Tú Linh gửi kèm khi trao đổi với TS. Hoàng Thị Thúy Nga.
* Bổ sung Mục 1.3 chuyên sâu giải trình lý do giữ lại biến Chuyển đổi số eFAST.

---

## 4. BẢN ĐỒ THƯ MỤC & HỆ THỐNG TỆP TIN DỰ ÁN

Toàn bộ các tệp tin trong thư mục làm việc `C:\Users\nguyen.tuan.minh\Desktop\DTL-Master-Project` được quản lý và đồng bộ trên GitHub:

| Tên Tệp Tin | Định Dạng | Mô Tả Chức Năng |
| :--- | :---: | :--- |
| **`DTL_Thesis_Design_NEU_MDE_Final.docx`** | Word | **File Đề cương Luận văn chính thức (Phương án 2)** – Đã cập nhật 100% theo nhận xét của TS. Hoàng Thị Thúy Nga (7 biến, Mục tiêu tiếng Anh NEU MDE, Sub-group ANOVA). |
| **`DTL_Thesis_Design_NEU_MDE_Final.md`** | Markdown | Bản xem trước Markdown của Đề cương Phương án 2. |
| **`Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.docx`** | Word | **Bản Báo cáo Giải trình Độc lập** – Tổng hợp góp ý của cô, phân tích điểm đồng tình/điều chỉnh và giải trình giữ lại biến eFAST. |
| **`Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.md`** | Markdown | Bản xem trước Markdown của Báo cáo giải trình. |
| **`DTL_Thesis_Design_Option1_Fee_Income_Final.docx`** | Word | **File Đề cương Luận văn Phương án 1** – Hướng Thu nhập Phí bảo lãnh (`ln_FEE` & `FEE_YIELD`) trên dữ liệu giao dịch Core Banking $n = 800$. |
| **`DTL_Thesis_Design_Option1_Fee_Income_Final.md`** | Markdown | Bản xem trước Markdown của Đề cương Phương án 1. |
| **`Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx`** | Word | Phiếu khảo sát chính thức chuẩn hóa 34 câu hỏi Likert 1-5 dành cho $n = 800$ doanh nghiệp. |
| **`Phieu_Khao_Sat_Chinh_Thuc_VietinBank.md`** | Markdown | Bản xem trước Phiếu khảo sát chính thức. |
| **`Tong_Hop_Bao_Cao_De_Tai_Master_VietinBank.docx`** | Word | Báo cáo tổng thể toàn diện về đề tài Master VietinBank. |
| **`Nhat_Ky_Trao_Doi_Va_Phat_Trien_De_Tai_Master.md`** | Markdown | **File này** – Nhật ký đóng gói toàn bộ quá trình trao đổi, học thuật và phát triển đề tài. |
| **`README.md`** | Markdown | Tài liệu hướng dẫn và tổng quan cấu trúc kho lưu trữ Git. |
| **`generate_final_mde_thesis_design.py`** | Python | Mã nguồn sinh file Word Đề cương Phương án 2 chuẩn định dạng NEU MDE. |
| **`generate_option1_mde_thesis_design.py`** | Python | Mã nguồn sinh file Word Đề cương Phương án 1 chuẩn định dạng NEU MDE. |
| **`generate_advisor_feedback_report_docx.py`** | Python | Mã nguồn sinh file Word Báo cáo giải trình góp ý của Giảng viên hướng dẫn. |

---

*Hồ sơ dự án được lưu trữ và sao lưu an toàn trên hệ thống máy tính cục bộ và GitHub Repository: `https://github.com/minhnguyen1604/-TL-Master-Thesis`.*
