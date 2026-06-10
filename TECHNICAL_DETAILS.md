# Tài liệu Kỹ thuật: Hệ thống Trích xuất Thuật ngữ Tự động (V2)

Tài liệu này cung cấp bản phân tích chi tiết về kiến trúc hệ thống, các công nghệ sử dụng, thuật toán cốt lõi và khả năng ứng dụng thực tế của nền tảng **Auto Glossary Extractor**.

## 1. Kiến trúc Hệ thống

Ứng dụng được xây dựng trên **kiến trúc phi trạng thái (stateless), thân thiện với môi trường serverless** nhằm vận hành trên các nền tảng đám mây có khả năng mở rộng (như Vercel) mà không cần phụ thuộc vào hệ thống lưu trữ tệp tin vật lý (persistent filesystem).

```
                    ┌────────────────────────┐
                    │  Frontend Browser UI   │
                    │   (HTML5, CSS3, JS)    │
                    └───────────┬────────────┘
                                │
          1. POST /upload       │      3. POST /format
          (File upload)         │      (Glossary JSON)
                                ▼
         ┌──────────────────────────────────────────┐
         │              Flask Backend               │
         │         (app.py / Vercel Serverless)     │
         └──────────────┬────────────────────┬──────┘
                        │                    │
                        ▼                    ▼
             ┌─────────────────────┐┌─────────────────┐
             │  core/extractor.py  ││core/formatter.py│
             │   (PyMuPDF/docx)    ││(Pandas/openpyxl)│
             └─────────────────────┘└─────────────────┘
```

### Các Trụ cột Kiến trúc Chính:
* **Thực thi Phi trạng thái (Stateless Execution)**: Backend không lưu tệp tin vĩnh viễn trên đĩa cứng. Các tệp tải lên được xử lý trực tiếp trong bộ nhớ hoặc thông qua các tệp tạm thời (`tempfile`) và được xóa ngay lập tức sau khi hoàn thành yêu cầu.
* **Định dạng trực tiếp trên Bộ nhớ (In-Memory Formatting)**: Tệp đầu ra Excel được biên dịch trực tiếp vào một luồng bộ nhớ (`io.BytesIO`) và truyền thẳng dưới dạng luồng dữ liệu (stream) về trình duyệt của người dùng, loại bỏ nhu cầu lưu trữ tệp trên máy chủ.
* **Quản lý Trạng thái ở Phía Trình duyệt (Client-Side State Management)**: Trình duyệt lưu giữ trạng thái hoạt động của danh sách thuật ngữ (`glossaryTerms`), cho phép người dùng kiểm tra, thêm mới, sửa đổi hoặc xóa các mục trước khi kích hoạt xuất bản tệp tin cuối cùng.

---

## 2. Các Thành phần Kỹ thuật Cốt lõi

### A. Phân tích Tài liệu & Trích xuất Văn bản (`core/extractor.py`)
Hệ thống phân tích hai định dạng tài liệu chính (.DOCX và .PDF) và hợp nhất chúng thành một chuỗi văn bản chuẩn hóa để phục vụ so khớp:
* **Tài liệu Word (`.docx`)**: Sử dụng thư viện `python-docx` để đọc tất cả các đoạn văn và ô trong bảng biểu, chuẩn hóa lỗi ngắt dòng ở giữa các từ (ví dụ: chuyển đổi `Arm's \n Length` thành `Arm's Length`).
* **Tài liệu PDF (`.pdf`)**: Sử dụng PyMuPDF (`fitz`) để trích xuất văn bản theo từng khối bố cục (layout-by-layout), xử lý các từ bị ngắt dòng giữa các trang.

### B. Thuật toán So khớp Từ khóa Tham lam Regex Tốc độ cao (`core/extractor.py`)
Đây là thuật toán cốt lõi của công cụ quét thuật ngữ:
1. **Biên dịch Từ điển (Dictionary Compiling)**: Tải tệp `core/tax_dictionary.json` (từ điển thuế chọn lọc chứa hàng trăm cặp thuật ngữ).
2. **Sắp xếp Tham lam (Greedy Sort)**: Sắp xếp tất cả các thuật ngữ theo độ dài giảm dần. Điều này đảm bảo các cụm từ ghép được đánh giá trước các từ đơn lẻ. Ví dụ: cụm từ `"transfer pricing policy"` được khớp trước để tránh việc khớp sai lệch thành các từ riêng lẻ là `"transfer pricing"` và `"policy"`.
3. **Biểu thức Chính quy Single-Pass (Single-Pass Regex)**: Escape các ký tự đặc biệt và nối tất cả các thuật ngữ bằng toán tử logic HOẶC `|`, bao bọc bởi ranh giới từ (`\b`):
   $$\text{Pattern} = \backslash\text{b}(\text{term}_1|\text{term}_2|\text{term}_3|\dots)\backslash\text{b}$$
   Mẫu Regex này được biên dịch với cờ `re.IGNORECASE` để quét toàn bộ nội dung tài liệu với tốc độ cao.
4. **Loại bỏ Trùng lặp Không Phân biệt Hoa Thường**: Các kết quả khớp được trả về với định dạng chữ hoa/thường nguyên bản trong tài liệu, sau đó được loại bỏ trùng lặp một cách không phân biệt hoa thường.

### C. Tạo & Mở rộng Biến thể Chữ Hoa/Thường (`core/formatter.py`)
Để tối đa hóa hiệu quả sử dụng của bảng thuật ngữ trong các công cụ dịch thuật hỗ trợ máy tính (CAT Tools như SDL Trados, memoQ hoặc Phrase), bộ máy xuất bản sẽ tự động mở rộng mỗi thuật ngữ thành 5 biến thể chữ hoa/thường và 1 biến thể viết tắt:
* **Chữ thường (Lowercase)**: `transfer pricing`
* **Chữ hoa (Uppercase)**: `TRANSFER PRICING`
* **Viết hoa đầu câu (Capitalized)**: `Transfer pricing`
* **Viết hoa mỗi từ (Title Case)**: `Transfer Pricing`
* **Sửa lỗi viết hoa có dấu nháy đơn**: `Arm's Length` (sửa hành vi mặc định của `.title()` trong Python vốn tạo ra lỗi `Arm'S`)
* **Từ viết tắt (Acronym)**: Tạo ra các biến thể từ viết tắt (ví dụ: `TRANSFER PRICING GLOBAL MASTER FILE` $\rightarrow$ `TPGMF`) đối với các cụm từ có $\ge 2$ từ.

### D. Đồng bộ hóa Cuộn Màn hình 2 chiều & Preview Trực tiếp (`templates/index.html`)
Màn hình frontend chứa các mã JavaScript hiệu năng cao để đồng bộ hóa danh sách thuật ngữ và giao diện preview văn bản:
* **Chuẩn hóa Dấu nháy (Quotes Normalization)**: Để ngăn chặn việc không khớp dấu nháy (dấu nháy cong `’` trong file PDF gốc so với dấu nháy thẳng `'` trong từ điển), toàn bộ văn bản được chuẩn hóa về dấu nháy thẳng (`'`) ở cả hai phía.
* **Liên kết Preview ➔ Editor**: Nhấp vào một thuật ngữ được highlight trong bảng xem trước sẽ thực hiện tìm kiếm trong mảng dữ liệu của trình duyệt, xác định chỉ mục (index), tự động cuộn dòng tương ứng trong bảng Editor vào giữa màn hình và nhấp nháy viền thẻ.
* **Liên kết Editor ➔ Preview**: Nhấp vào biểu tượng kính lúp bên cạnh ô nhập liệu của Editor sẽ kích hoạt một câu lệnh truy vấn DOM trong bảng xem trước, cuộn vị trí xuất hiện đầu tiên của thuật ngữ đó vào giữa màn hình và highlight nó bằng màu vàng.

---

## 3. Công nghệ Sử dụng (Technology Stack)

| Lớp | Công nghệ | Mục đích |
| :--- | :--- | :--- |
| **Backend Framework** | `Flask (Python 3.9+)` | Định tuyến serverless tinh gọn, tốc độ cao |
| **Trình đọc Tài liệu** | `PyMuPDF (fitz)` & `python-docx` | Trích xuất văn bản cấp độ bố cục từ tệp PDF và Word |
| **Dữ liệu & Bảng tính** | `Pandas` & `openpyxl` | Tạo tệp Excel, ánh xạ biến thể chữ hoa/thường và sắp xếp |
| **Cấu trúc Frontend** | `HTML5` & `JavaScript (ES6)` | Thao tác DOM trực tiếp, phân tích cú pháp JSON phi trạng thái và đồng bộ hóa bố cục |
| **Giao diện & Theme** | `CSS3 (Vanilla)` | Truy vấn đáp ứng (Media queries), màu sắc thương hiệu KPMG và các ô nhập liệu tối giản |
| **Cloud Hosting** | `Vercel` | Môi trường thực thi serverless của Python |

---

## 4. Khả năng Áp dụng Rộng rãi (Generalization)

Mặc dù ứng dụng này được tùy chỉnh kèm theo **từ điển thuế** (`core/tax_dictionary.json`) và Nhận diện Thương hiệu của KPMG, phần lõi kỹ thuật hoàn toàn mang tính tổng quát và có thể áp dụng cho bất kỳ hệ thống trích xuất thuật ngữ chuyên ngành nào:

### 1. Thay đổi Từ điển (Mọi Lĩnh vực)
Bằng cách thay thế tệp `core/tax_dictionary.json`, công cụ này có thể quét ngay lập tức các thuật ngữ thuộc:
* **Y học / Dược phẩm**: Quét các tài liệu thử nghiệm lâm sàng để tìm tên thuốc, tác dụng phụ và triệu chứng.
* **Pháp lý / Hợp đồng**: Quét các thỏa thuận pháp lý để tìm các điều khoản hợp đồng, tên pháp nhân và các thuật ngữ tuân thủ.
* **Kỹ thuật / Xây dựng**: Quét các hồ sơ đặc tính kỹ thuật xây dựng để tìm vật liệu, tiêu chuẩn và mã thiết bị.

### 2. Hỗ trợ Đa Ngôn ngữ
Mặc dù giao diện mặc định và các cột Excel sử dụng ngôn ngữ hiển thị là `Source` (tiếng Anh) và `Target` (tiếng Việt), các ô nhập bản dịch có thể tiếp nhận bất kỳ ký tự Unicode nào (tiếng Trung, tiếng Nhật, tiếng Hàn, tiếng Nga, tiếng Ả Rập...) và lưu thành công vào định dạng bảng tính Excel UTF-8 tiêu chuẩn.

### 3. Tích hợp với các Công cụ CAT Tool
Định dạng bảng tính được tạo ra (chứa các cột song song của các biến thể chữ hoa/thường đã được mở rộng) được thiết kế để nhập trực tiếp vào cơ sở dữ liệu dịch thuật:
* Có thể chuyển đổi sang định dạng `.tbx` (TermBase eXchange).
* Có thể nhập trực tiếp vào hệ thống Trados MultiTerm hoặc memoQ Termbase, giúp biên dịch viên tiết kiệm hàng giờ sao chép thủ công.
