import io
import sys
import os
import pandas as pd

# Ensure project root on sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from app import read_excel_safely, clean_rows_by_patterns, clean_xls_specific_rules


def main(path: str):
    with open(path, "rb") as f:
        raw = f.read()
    df = read_excel_safely(raw, path)
    print("Original shape:", df.shape)

    # Save raw CSV
    raw_buf = io.StringIO()
    df.to_csv(raw_buf, sep=";", index=False, encoding="utf-8", decimal=",")
    with open(path.rsplit(".", 1)[0] + "_raw.csv", "wb") as f:
        f.write(("\ufeff" + raw_buf.getvalue()).encode("utf-8"))

    # Clean rows by patterns
    clean_df = clean_rows_by_patterns(df)
    if path.lower().endswith('.xls'):
        clean_df = clean_xls_specific_rules(clean_df)
    print("Cleaned shape:", clean_df.shape)
    print("Preview cleaned (top 10):")
    print(clean_df.head(10))

    # Save cleaned CSV
    clean_buf = io.StringIO()
    clean_df.to_csv(clean_buf, sep=";", index=False, encoding="utf-8", decimal=",")
    with open(path.rsplit(".", 1)[0] + "_limpo.csv", "wb") as f:
        f.write(("\ufeff" + clean_buf.getvalue()).encode("utf-8"))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python scripts/clean_file.py <arquivo.xlsx|csv>")
        sys.exit(1)
    main(sys.argv[1])
