import sys
import pandas as pd


def main(path: str):
    print(f"Lendo arquivo: {path}")
    # Carrega todas as planilhas para entender a estrutura
    xls = pd.ExcelFile(path)
    print("Sheets:", xls.sheet_names)
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=None)
        except Exception as e:
            print(f"[ERRO] Falha ao ler sheet '{sheet}': {e}")
            continue
        print(f"\n=== Sheet: {sheet} ===")
        print("Dimensões:", df.shape)
        print("Primeiras 8 linhas (sem header):")
        print(df.head(8))

        # Pré-processo: pular 3 linhas, remover linhas em branco
        df2 = df.iloc[3:, :].copy()
        df2 = df2.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
        print("\nApós pular 3 linhas e limpar vazios:")
        print("Dimensões:", df2.shape)
        print(df2.head(10))

        # Mostrar cabeçalhos estimados (usando a primeira linha útil como header)
        if not df2.empty:
            hdr = df2.iloc[0]
            print("\nCabeçalhos estimados (linha 1 após pulo):")
            for i, v in enumerate(hdr):
                print(f"{i}: {v}")

            # dados a partir da próxima linha
            data = df2.iloc[1:].reset_index(drop=True)
            print("\nPrévia de dados (primeiras 5 linhas):")
            print(data.head(5))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python scripts/inspect_excel.py <caminho.xlsx>")
        sys.exit(1)
    main(sys.argv[1])

