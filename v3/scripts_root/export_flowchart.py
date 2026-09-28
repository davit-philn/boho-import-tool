"""
export_flowchart.py — Render PROCESS_FLOW.md thành HTML xem trực tiếp trên trình duyệt.
Cách dùng:
    python scripts/export_flowchart.py
"""
import re
import sys
import subprocess
from pathlib import Path

SRC  = Path(__file__).parent.parent / "docs" / "PROCESS_FLOW.md"
DEST = Path(__file__).parent.parent / "docs" / "process_flow.html"


def extract_mermaid_blocks(md: str) -> list[str]:
    return re.findall(r"```mermaid\s*\n(.*?)```", md, re.DOTALL)


def build_html(md: str, mermaid_blocks: list[str]) -> str:
    # Chuyển phần ngoài code block thành HTML đơn giản
    # (chỉ render header + bảng dạng text, Mermaid được render qua JS)
    desc_lines = []
    in_code = False
    for line in md.splitlines():
        if line.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if line.startswith("# "):
            desc_lines.append(f"<h1>{line[2:]}</h1>")
        elif line.startswith("## "):
            desc_lines.append(f"<h2>{line[3:]}</h2>")
        elif line.startswith("### "):
            desc_lines.append(f"<h3>{line[4:]}</h3>")
        elif line.startswith("> "):
            desc_lines.append(f'<blockquote>{line[2:]}</blockquote>')
        elif line.startswith("| "):
            desc_lines.append(line)          # tích lũy cho xử lý bảng bên dưới
        elif line.strip() == "---":
            desc_lines.append("<hr>")
        elif line.strip():
            desc_lines.append(f"<p>{line}</p>")
        else:
            desc_lines.append("")

    # Gộp các dòng bảng Markdown → HTML table
    html_body_parts: list[str] = []
    table_buf: list[str] = []
    for line in desc_lines:
        if line.startswith("| "):
            table_buf.append(line)
        else:
            if table_buf:
                html_body_parts.append(_md_table_to_html(table_buf))
                table_buf = []
            html_body_parts.append(line)
    if table_buf:
        html_body_parts.append(_md_table_to_html(table_buf))

    body_html = "\n".join(html_body_parts)

    # Chèn sơ đồ Mermaid vào vị trí đúng (sau h2 "Sơ đồ quy trình")
    diagram_html = ""
    for block in mermaid_blocks:
        diagram_html += (
            '<div class="diagram-wrapper">'
            f'<pre class="mermaid">{block}</pre>'
            "</div>\n"
        )

    # Thay thế placeholder (đoạn body sau h2 "Sơ đồ") bằng diagram
    body_html = body_html.replace(
        "<h2>Sơ đồ quy trình</h2>",
        f"<h2>Sơ đồ quy trình</h2>\n{diagram_html}",
    )

    return f"""<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>BOHO Import Tool — Sơ đồ quy trình</title>
  <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
  <style>
    :root {{
      --blue:   #1D4ED8; --blue-lt: #DBEAFE;
      --yellow: #CA8A04; --yellow-lt: #FEF9C3;
      --pink:   #9D174D; --pink-lt: #FCE7F3;
      --green:  #16A34A; --red: #DC2626;
      --bg: #F8FAFC; --card: #FFFFFF;
      --text: #1E293B; --muted: #64748B;
      --border: #E2E8F0;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: "Segoe UI", Arial, sans-serif;
      background: var(--bg); color: var(--text);
      padding: 32px 24px; max-width: 1400px; margin: 0 auto;
    }}
    h1 {{
      color: var(--blue); font-size: 22px;
      border-bottom: 3px solid var(--blue-lt);
      padding-bottom: 10px; margin-bottom: 20px;
    }}
    h2 {{
      color: var(--text); font-size: 16px;
      margin: 28px 0 12px; border-left: 4px solid var(--blue);
      padding-left: 10px;
    }}
    h3 {{ color: var(--muted); font-size: 14px; margin: 16px 0 8px; }}
    blockquote {{
      background: var(--blue-lt); border-left: 4px solid var(--blue);
      padding: 8px 14px; border-radius: 4px; margin: 12px 0;
      color: var(--blue); font-size: 13px;
    }}
    p {{ margin: 6px 0; font-size: 14px; line-height: 1.6; }}
    hr {{ border: none; border-top: 1px solid var(--border); margin: 20px 0; }}

    /* Diagram wrapper */
    .diagram-wrapper {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 24px;
      margin: 16px 0 24px;
      overflow: auto;
      box-shadow: 0 2px 8px rgba(0,0,0,.06);
    }}
    .mermaid svg {{ max-width: none; }}

    /* Tables */
    table {{
      border-collapse: collapse; width: 100%;
      font-size: 13px; margin: 12px 0 20px;
      background: var(--card);
      border-radius: 8px; overflow: hidden;
      box-shadow: 0 1px 4px rgba(0,0,0,.05);
    }}
    th {{
      background: var(--blue); color: #fff;
      padding: 9px 12px; text-align: left; font-weight: 600;
    }}
    td {{ padding: 8px 12px; border-bottom: 1px solid var(--border); }}
    tr:last-child td {{ border-bottom: none; }}
    tr:nth-child(even) td {{ background: #F8FAFC; }}

    /* Legend chips */
    .legend {{ display: flex; flex-wrap: wrap; gap: 8px; margin: 12px 0; }}
    .chip {{
      display: inline-flex; align-items: center; gap: 6px;
      padding: 4px 10px; border-radius: 20px;
      font-size: 12px; font-weight: 500;
    }}

    footer {{
      margin-top: 40px; padding-top: 16px;
      border-top: 1px solid var(--border);
      color: var(--muted); font-size: 12px;
    }}
  </style>
</head>
<body>
{body_html}
<footer>
  Generated by <code>scripts/export_flowchart.py</code> &mdash;
  BOHO Import BOM/THDM v2.1 &mdash; {__import__('datetime').date.today()}
</footer>
<script>
  mermaid.initialize({{
    startOnLoad: true,
    theme: "base",
    themeVariables: {{
      primaryColor:       "#DBEAFE",
      primaryBorderColor: "#1D4ED8",
      secondaryColor:     "#FEF9C3",
      tertiaryColor:      "#FCE7F3",
      edgeLabelBackground:"#F8FAFC",
      fontSize:           "13px"
    }},
    flowchart: {{ curve: "basis", useMaxWidth: false }}
  }});
</script>
</body>
</html>"""


def _md_table_to_html(rows: list[str]) -> str:
    lines = [r for r in rows if not re.match(r"\|\s*[-:]+", r)]
    if not lines:
        return ""
    header, *body = lines
    cells = lambda row: [c.strip() for c in row.strip().strip("|").split("|")]
    ths = "".join(f"<th>{c}</th>" for c in cells(header))
    trs = "".join(
        "<tr>" + "".join(f"<td>{c}</td>" for c in cells(row)) + "</tr>"
        for row in body
    )
    return f"<table><thead><tr>{ths}</tr></thead><tbody>{trs}</tbody></table>"


def main() -> int:
    if not SRC.exists():
        print(f"❌ Không tìm thấy: {SRC}")
        return 1

    md = SRC.read_text(encoding="utf-8")
    blocks = extract_mermaid_blocks(md)
    if not blocks:
        print("❌ Không có block Mermaid trong file!")
        return 1

    html = build_html(md, blocks)
    DEST.write_text(html, encoding="utf-8")
    print(f"✅ Đã tạo: {DEST}")
    print(f"   Mở: file:///{DEST.as_posix()}")

    # Tự mở trình duyệt nếu có thể
    try:
        if sys.platform == "win32":
            subprocess.Popen(["start", "", str(DEST)], shell=True)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(DEST)])
        else:
            subprocess.Popen(["xdg-open", str(DEST)])
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
