import pandas as pd

def generate_excel_glossary(glossary_items: list, output_path: str):
    """
    Format Expansion: Nhân bản định dạng (Case-sensitive)
    Cho mỗi cặp thuật ngữ (EN, VI), tạo các dòng biến thể tương ứng vào Excel.
    """
    rows = []
    for item in glossary_items:
        term = item.get("english", "").strip()
        trans = item.get("vietnamese", "").strip()
        if not term:
            continue
            
        # Dùng dict để tránh trùng lặp Keyword (EN) biến thể
        mapping = {}
        
        # 1. Chữ thường (lowercase)
        mapping[term.lower()] = trans.lower() if trans else ""
        
        # 2. Chữ HOA (uppercase)
        mapping[term.upper()] = trans.upper() if trans else ""
        
        # 3. Viết hoa chữ đầu câu (capitalize)
        mapping[term.capitalize()] = trans.capitalize() if trans else ""
        
        # 4. Tiêu chuẩn Title Case của Python
        mapping[term.title()] = trans.title() if trans else ""
        
        # 5. Sửa lỗi title() của Python với dấu nháy đơn
        # VD: "arm's".title() ra "Arm'S", ta đổi thành "Arm's"
        corrected_title_en = " ".join([w.capitalize() for w in term.split()])
        corrected_title_vi = " ".join([w.capitalize() for w in trans.split()]) if trans else ""
        mapping[corrected_title_en] = corrected_title_vi
        
        # 6. Biến thể Viết tắt (Acronym) - chỉ áp dụng cho cụm từ tiếng Anh từ 2 từ trở lên
        words = term.split()
        if len(words) >= 2:
            acronym = "".join([w[0].upper() for w in words])
            mapping[acronym] = trans  # Giữ nguyên bản dịch gốc cho từ viết tắt
            
        # Đưa vào danh sách xuất
        for en_var, vi_var in mapping.items():
            rows.append({
                "Source": en_var,
                "Target": vi_var
            })
            
    # Ghi ra file Excel
    df = pd.DataFrame(rows)
    # Sắp xếp để các biến thể của cùng một từ đứng cạnh nhau
    df.sort_values(by=["Source"], key=lambda col: col.str.lower(), inplace=True)
    
    df.to_excel(output_path, index=False)
    return len(rows)

