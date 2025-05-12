import fitz  # PyMuPDF
import re
import pandas as pd

pdf_path = "zuelhc1o.pdf"
output_pdf_path = "zuelhc1o_annotated.pdf"
output_excel_path = "zuelhc1o_dimensions.xlsx"
doc = fitz.open(pdf_path)

# Regex to extract full dimension values
dimension_pattern = re.compile(r'(⌀\d+(\.\d+)?|R\d+(\.\d+)?|\d+\.\d+|\b\d+\b)')

found_dimensions = []
dimension_positions = []

# Extract dimensions and positions
for page_num, page in enumerate(doc, start=1):
    blocks = page.get_text("blocks")
    for block in blocks:
        text = block[4].strip()
        matches = dimension_pattern.findall(text)
        for match in matches:
            dim = match[0]
            if dim not in found_dimensions:
                found_dimensions.append(dim)
                rect = fitz.Rect(block[0], block[1], block[2], block[3])
                dimension_positions.append((page_num, rect, dim))

# Annotate with numbers
for i, (page_num, rect, dim) in enumerate(dimension_positions, start=1):
    page = doc[page_num - 1]
    insert_point = fitz.Point(rect.x1 + 5, rect.y1)
    page.insert_text(
        insert_point,
        f"#{i}",
        fontname="helv",
        fontsize=12,
        color=(1, 0, 0),
    )

doc.save(output_pdf_path)
print(f"Annotated PDF saved: {output_pdf_path}")

# Split symbol and value for Excel
symbols = []
values = []

for dim in found_dimensions:
    if dim.startswith('R'):
        symbols.append('R')
        values.append(dim[1:])
    elif dim.startswith('⌀'):
        symbols.append('⌀')
        values.append(dim[1:])
    else:
        symbols.append('')
        values.append(dim)

# Create DataFrame
df = pd.DataFrame({
    "Serial Number": list(range(1, len(found_dimensions) + 1)),
    "Symbol": symbols,
    "Value": values
})

df.to_excel(output_excel_path, index=False)
print(f"Excel sheet saved: {output_excel_path}")
