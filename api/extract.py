from pathlib import Path
import pdfplumber
import pandas as pd

base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / "wawa policy.pdf"
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

lines = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text(x_tolerance=2, y_tolerance=0)
        for line in text.split("\n"):
            lines.append(line)
            print(line)
        # print(page.extract_tables())

df = pd.DataFrame(lines)
df.to_csv(output_dir / "test.csv")


