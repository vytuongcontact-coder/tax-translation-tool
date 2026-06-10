# Báo Cáo Kỹ Thuật: Hệ Thống Trích Xuất Thuật Ngữ Thuế (Tax Glossary Extractor v2.0)

**Mục tiêu:** Xây dựng công cụ tự động trích xuất thuật ngữ chuyên ngành từ Báo cáo tài chính/Thuế (định dạng DOCX) để tạo bảng thuật ngữ (Glossary) bằng Excel, phục vụ cho quá trình dịch máy (Machine Translation).
**Đặc điểm hệ thống:** Lập trình bằng Python, xử lý ngoại tuyến (offline), sử dụng thuật toán tìm kiếm chuỗi thay cho các mô hình AI suy luận.

---

## 1. Phương Pháp và Công Cụ Sử Dụng

### 1.1. Thuật toán Aho-Corasick (Thư viện FlashText)
*   **Công cụ:** Thư viện `flashtext` trong Python.
*   **Phương pháp:** Thay vì sử dụng vòng lặp Regex để quét từng từ khóa hoặc dùng Xử lý ngôn ngữ tự nhiên (NLP) để phân tích từ loại, hệ thống dùng thuật toán Aho-Corasick. Thuật toán này tìm kiếm nhiều chuỗi cùng lúc bằng cách quét văn bản một lần duy nhất.
*   **Đặc điểm:** 
    * Khắc phục lỗi sai định dạng của NLP (các công cụ NLP trước đây thường nhận diện sai từ loại, ví dụ phân loại "Transfer" thành động từ, dẫn đến trích xuất thiếu cụm "Transfer Pricing").
    * Bật tùy chọn `span_info=True` để lấy vị trí (index) của từ khóa trong văn bản. Từ đó, công cụ trích xuất được chính xác định dạng chữ hoa/chữ thường như bản gốc trong báo cáo (ví dụ: lấy được `Transfer Pricing Policy` thay vì trả về mặc định chữ thường).

### 1.2. Tổng hợp Từ Điển (Dictionary Filtering)
*   **Công cụ:** Python (đọc file JSON, Regex, thư viện NLTK).
*   **Phương pháp:** Hệ thống sử dụng một danh sách từ khóa cố định (`tax_dictionary.json`). Danh sách này được tổng hợp từ:
    1.  **Từ điển SAPP (670 từ):** Lọc tự động từ file PDF dựa trên định dạng dòng phiên âm tiếng Anh.
    2.  **Tài liệu PWC (1.141 từ):** Lọc tự động từ file PDF. Các cụm từ bị lỗi font do quá trình trích xuất PDF được loại bỏ bằng cách đối chiếu từng từ với từ điển tiếng Anh của thư viện `nltk.corpus.words`.
    3.  **Từ điển người dùng (103 từ):** Tích hợp trực tiếp từ file Excel thuật ngữ do người dùng cung cấp.
    4.  **Lọc danh từ riêng:** Dùng thuật toán loại bỏ các tên quốc gia, vùng lãnh thổ (US, Vietnam, Hong Kong...) bằng danh sách loại trừ (blacklist).
    *   **Kết quả:** Danh sách tổng hợp gồm 1.879 thuật ngữ thuế và kế toán tiếng Anh.

### 1.3. Tiền Xử Lý Văn Bản
*   **Công cụ:** Thư viện `python-docx`.
*   **Phương pháp:** 
    * Đọc và lấy văn bản từ các đoạn văn (paragraphs) và bảng biểu (tables) trong tệp DOCX.
    * Xử lý lỗi ngắt dòng do giới hạn chiều rộng của ô trong bảng (ví dụ: tự động nối các ký tự bị ngắt xuống dòng thành một cụm từ hoàn chỉnh).
    * Đồng bộ các loại dấu nháy đơn (`’`, `‘`) về định dạng tiêu chuẩn (`'`) để đảm bảo thuật toán so sánh chuỗi hoạt động chính xác.

---

## 2. Quy Trình Hoạt Động (Workflow)

Quá trình trích xuất thuật ngữ từ một tệp DOCX bao gồm 4 bước:

### Bước 1: Đọc và Tiền Xử Lý Văn Bản
1. Giao diện Web (Flask) nhận tệp DOCX từ người dùng tải lên.
2. Thư viện `python-docx` đọc và gom nhóm toàn bộ văn bản từ tài liệu.
3. Chạy hàm xử lý ngắt dòng và đồng bộ dấu câu.

### Bước 2: Quét Từ Điển (Dictionary Scanning)
1. Danh sách 1.879 thuật ngữ được nạp vào cấu trúc dữ liệu của `FlashText`.
2. Hệ thống quét qua văn bản và ghi nhận tất cả các vị trí trùng khớp với từ khóa trong danh sách (không phân biệt hoa/thường).
3. Hệ thống sử dụng vị trí index thu được để cắt (slice) và lưu lại từ khóa theo đúng định dạng nguyên bản từ tài liệu tải lên.

### Bước 3: Định Dạng Từ Khóa
Mỗi từ khóa tìm được sẽ được định dạng thành 5 dạng khác nhau (case-sensitive) để chuẩn bị cho bước dịch thuật tự động (chạy string replacement) sau này:
*   `Original`: Giữ nguyên định dạng trong văn bản.
*   `UPPERCASE`: Chữ in hoa toàn bộ.
*   `lowercase`: Chữ viết thường toàn bộ.
*   `Sentence case`: Chỉ viết hoa chữ cái đầu tiên của cụm từ.
*   `Title Case`: Viết hoa chữ cái đầu tiên của từng từ.

### Bước 4: Xuất Tệp Excel
1. Xóa bỏ các kết quả bị trùng lặp.
2. Sắp xếp danh sách theo chiều dài từ khóa giảm dần (giúp các công cụ dịch máy ưu tiên thay thế các cụm từ dài trước, tránh việc cụm từ ngắn bị thay thế đè lên cụm từ dài).
3. Xuất kết quả ra tệp `glossary.xlsx` với 2 cột: Source (Tiếng Anh) và Target (Tiếng Việt - để trống) và trả tệp về cho người dùng tải xuống.
