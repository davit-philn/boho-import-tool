"""
Đọc nội dung file .docx và in ra màn hình (không cần cài thêm thư viện)
"""
import zipfile
import re
import sys
import os

def read_docx(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    lines = []
    # Tách đoạn theo <w:p>
    paras = re.split(r'</w:p>', xml)
    for para in paras:
        t = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para)
        line = "".join(t).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)

if __name__ == "__main__":
    path = r"D:\Personal\Project\insertdatasqlserver\Doc\Hop_270626.docx"
    if len(sys.argv) > 1:
        path = sys.argv[1]
    if not os.path.exists(path):
        print(f"Không tìm thấy file: {path}")
        input()
    else:
        content = read_docx(path)
        print(content)
        # Lưu ra file txt bên cạnh
        out = path.replace(".docx", "_content.txt")
        with open(out, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"\n--- Đã lưu ra: {out}")
        input("\nNhấn Enter để đóng...")
