import io
import unicodedata
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st


APP_TITLE = "Conversor Uniodonto: Excel → CSV (; UTF-8)"


def remove_accents(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ASCII", "ignore")
        .decode("ASCII")
    )


def normalize(text: str) -> str:
    text = remove_accents(text or "").lower().strip()
    for ch in ["\n", "\t", ",", ".", "/", "-", "_", "(", ")", ":"]:
        text = text.replace(ch, " ")
    text = " ".join(text.split())
    return text


REQUIRED_COLUMNS: List[str] = [
    "CRO",
    "COOPERADO",
    "ATO COBERTO",
    "ATO COMPLEM.",
    "OUTROS PAGTOS",
    "TOTAL GERAL",
    "DIVERSOS",
    "TX. ADM FATMOD",
    "INSS",
    "IRRF",
    "Outras Fontes",
    "SALDO LÍQUIDO",
    "INSCRIÇÃO ISS",
]


# Known aliases (normalized) for auto-detect
ALIASES: Dict[str, List[str]] = {
    "CRO": ["cro", "cro uf", "cro/uf", "registro cro"],
    "COOPERADO": ["cooperado", "prestador", "profissional", "dentista", "nome"],
    "ATO COBERTO": ["ato coberto", "coberto", "procedimentos cobertos"],
    # Some reports call it Ato não coberto / complementar
    "ATO COMPLEM.": [
        "ato complem",
        "ato complementar",
        "ato nao coberto",
        "nao coberto",
    ],
    "OUTROS PAGTOS": ["outros pagtos", "outros pg", "outros pagamentos", "outros"],
    # Depending on report this may be Total Bruto or Total Geral
    "TOTAL GERAL": ["total geral", "total bruto", "total"],
    "DIVERSOS": ["diversos"],
    "TX. ADM FATMOD": ["tx adm fatmod", "taxa adm fatmod", "taxa administrativa fatmod", "taxa adm"],
    "INSS": ["inss"],
    "IRRF": ["irrf"],
    "Outras Fontes": ["outras fontes", "outrasfontes"],
    # Often labeled Total Líquido
    "SALDO LÍQUIDO": ["saldo liquido", "total liquido", "liquido"],
    "INSCRIÇÃO ISS": ["inscricao iss", "inscricao municipal", "inscricao-iss", "insc iss"],
}


NUMERIC_TARGETS = {
    "ATO COBERTO",
    "ATO COMPLEM.",
    "OUTROS PAGTOS",
    "TOTAL GERAL",
    "DIVERSOS",
    "TX. ADM FATMOD",
    "INSS",
    "IRRF",
    "Outras Fontes",
    "SALDO LÍQUIDO",
}

# Positional sequence for reports with merged headers where only some
# columns have data. We map remaining non-empty columns left→right to
# these canonical names provided by the user.
ORDER_SEQUENCE: List[str] = [
    "prestador",
    "cpf_cnpj",
    "inss_1",
    "iss_1",
    "orcamentos",
    "coberto",
    "nao_coberto",
    "copart",
    "outros_pagamentos",
    "total_bruto",
    "outros_desc",
    "base_desc_inss",
    "inss_2",
    "irrf",
    "iss_2",
    "csll",
    "cofins",
    "pis",
    "total_liquido",
]
ORDER_NON_NUMERIC = {"prestador", "cpf_cnpj"}


def auto_map_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    norm_cols = {normalize(c): c for c in df.columns}
    mapping: Dict[str, Optional[str]] = {k: None for k in REQUIRED_COLUMNS}
    for required in REQUIRED_COLUMNS:
        candidates = ALIASES.get(required, [])
        for cand in candidates:
            if cand in norm_cols:
                mapping[required] = norm_cols[cand]
                break
        # small heuristic: try exact normalized label of required itself
        if mapping[required] is None:
            req_norm = normalize(required)
            if req_norm in norm_cols:
                mapping[required] = norm_cols[req_norm]
    return mapping


def to_numeric(series: pd.Series) -> pd.Series:
    # Coerce to numeric, handling common BR/US decimal formats robustly
    if series.dtype.kind in {"i", "u", "f"}:
        return series.fillna(0).abs()
    def norm(x):
        if x is None:
            return None
        s = str(x).strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return None
        has_comma = "," in s
        has_dot = "." in s
        if has_comma and not has_dot:
            # 1.234,56 sometimes comes without dots; here we only saw comma
            return s.replace(" ", "").replace(".", "").replace(",", ".")
        if has_dot and not has_comma:
            # Assume dot is decimal: keep as-is
            return s.replace(" ", "")
        if has_dot and has_comma:
            # Typical BR: thousands '.' and decimal ','
            return s.replace(" ", "").replace(".", "").replace(",", ".")
        # digits only
        return s
    s = series.map(norm)
    # Remove negative sign after coercion because these fields are counted values
    return pd.to_numeric(s, errors="coerce").fillna(0.0).abs()


def build_output_df(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = pd.DataFrame()
    for col in REQUIRED_COLUMNS:
        src = mapping.get(col)
        if src and src in df.columns:
            series = df[src]
            if col in NUMERIC_TARGETS:
                series = to_numeric(series)
            out[col] = series
        else:
            # Missing: fill with sane default
            if col in NUMERIC_TARGETS:
                out[col] = 0.0
            else:
                out[col] = ""
    return out


def read_excel_safely(file_bytes: bytes, filename: str) -> pd.DataFrame:
    # Force strings and no header to preserve leading zeros (e.g. NUM_ISS)
    df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
    # If source is .xls: drop first column (artifact) and trim rows until the
    # first occurrence of 'COOPERADO' (or 'COOOPERADO') in the first column,
    # then start from the next row.
    if isinstance(filename, str) and filename.lower().endswith(".xls"):
        if df.shape[1] >= 1:
            df = df.iloc[:, 1:]
        if df.shape[0] > 0:
            first_col = df.iloc[:, 0].astype(str)
            norm = first_col.map(remove_accents).str.lower().str.strip()
            idx_matches = norm[norm.isin({"cooperado", "coooperado"})].index
            if len(idx_matches) > 0:
                start = idx_matches[0] + 1
                if start < len(df):
                    df = df.iloc[start:, :]
    return df


def convert_single_file(uploaded_file, decimal_comma: bool = True) -> Dict[str, bytes]:
    content = uploaded_file.read()
    df = read_excel_safely(content, uploaded_file.name)

    # Attempt to drop fully empty columns/rows
    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.to_series().astype(str).str.fullmatch(r"Unnamed:.*", na=False)]

    mapping = auto_map_columns(df)

    # UI-side mapping happens outside; here we just ensure order and defaults
    out_df = build_output_df(df, mapping)

    csv_buffer = io.StringIO()
    # Use decimal comma if requested
    decimal_char = "," if decimal_comma else "."
    out_df.to_csv(
        csv_buffer,
        sep=";",
        index=False,
        encoding="utf-8",
        decimal=decimal_char,
    )
    # Add UTF-8 BOM for Excel compatibility
    csv_bytes = ("\ufeff" + csv_buffer.getvalue()).encode("utf-8")
    return {
        "csv_bytes": csv_bytes,
        "mapping": mapping,
        "dataframe": out_df,
    }


def non_empty_columns_in_order(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        series = df[c]
        # Treat empty strings as NaN for the check
        is_all_na = series.replace("", pd.NA).isna().all()
        if not is_all_na:
            cols.append(c)
    return cols


def build_positional_output(df: pd.DataFrame) -> pd.DataFrame:
    cols = non_empty_columns_in_order(df)
    # Align counts
    take = min(len(cols), len(ORDER_SEQUENCE))
    cols = cols[:take]
    names = ORDER_SEQUENCE[:take]

    out = pd.DataFrame()
    for src, name in zip(cols, names):
        series = df[src]
        if name not in ORDER_NON_NUMERIC:
            series = to_numeric(series)
        out[name] = series
    return out


def preprocess_excel_df(df: pd.DataFrame, skip_rows: int = 3) -> pd.DataFrame:
    # Skip a fixed number of initial rows (used for .xlsx). For .xls we pass 0.
    df2 = df.iloc[skip_rows:, :].copy() if skip_rows else df.copy()
    # Treat blank strings as NA and drop fully empty rows
    df2 = df2.replace(r"^\s*$", pd.NA, regex=True)
    df2 = df2.dropna(how="all")
    return df2


def filter_output_rows(out_df: pd.DataFrame) -> pd.DataFrame:
    # Drop blank rows again just in case
    out = out_df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    if "NOME" in out.columns:
        s = out["NOME"].astype(str)
        s_norm = s.map(remove_accents).str.lower().str.strip()
        mask = ~(
            s_norm.str.startswith("inscricao")
            | s_norm.str.startswith("prestador")
            | s_norm.str.contains(r"\b(cooperado|total|resumo|subtotal|credenciado|radiologia)\b", regex=True)
        ) & s_norm.ne("")
        out = out[mask]
    return out.reset_index(drop=True)


def clean_rows_by_patterns(df: pd.DataFrame) -> pd.DataFrame:
    # Drop fully blank rows
    temp = df.replace(r"^\s*$", pd.NA, regex=True)
    temp = temp.dropna(how="all")
    # Join row text for pattern checks (first 12 columns are enough)
    first_col = temp.iloc[:, 0].astype(str)
    first_norm = first_col.map(remove_accents).str.lower().str.strip()
    joined = temp.iloc[:, :12].astype(str).agg(" ".join, axis=1)
    norm = joined.map(remove_accents).str.lower().str.strip()
    # Patterns seen in page headers/footers
    patterns = [
        r"resumo geral de fechamento",
        r"periodo de fechamento",
        r"fator moderador",
        r"pagina \d+ de \d+|pagina",
        r"^co+operado$|^cooperado$|^credenciado$|^radiologia$",
        r"^subtotais? do fechamento",
        r"^total fechamento|^total geral a pagar",
        r"inss retido|irrf retido|iss retido|csll retido|cofins retido|pis retido",
        r"^prestador(\s|$)",
        r"^inscricao(\s|$)",
    ]
    combined = r"(" + r"|".join(patterns) + r")"
    mask_bad = norm.str.contains(combined, regex=True)
    # Also remove when the first column is exactly a known section label
    first_exact = {"cooperado", "coooperado", "credenciado", "radiologia"}
    mask_bad = mask_bad | first_norm.isin(first_exact)
    cleaned = temp[~mask_bad].copy()
    return cleaned.reset_index(drop=True)


def clean_xls_specific_rules(df: pd.DataFrame) -> pd.DataFrame:
    # Additional filters when source is .xls, based on specific columns
    if df.empty:
        return df
    temp = df.copy()
    # Column indices are 1-based in user's description: 25 and 24
    c25 = 24 if temp.shape[1] > 24 else None
    c24 = 23 if temp.shape[1] > 23 else None
    mask25 = pd.Series(False, index=temp.index)
    mask24 = pd.Series(False, index=temp.index)
    if c25 is not None:
        t = temp.iloc[:, c25].astype(str).map(remove_accents).str.lower().str.strip()
        prefixes = (
            "inss retido:",
            "irrf retido:",
            "iss retido:",
            "csll retido:",
            "cofins retido:",
            "pis retido:",
            "total geral a pagar:",
            "total geral a pagar",
        )
        mask25 = t.str.startswith(prefixes)
    if c24 is not None:
        t = temp.iloc[:, c24].astype(str).map(remove_accents).str.lower().str.strip()
        mask24 = t.str.startswith("pagina")
    cleaned = temp[~(mask25 | mask24)].copy()
    return cleaned.reset_index(drop=True)


def find_section_indices(df: pd.DataFrame, is_xls: bool) -> tuple[int, int | None]:
    # Returns (coop_start, cred_start) based on first-column markers in raw df
    if df.empty:
        return (0, None)
    first = df.iloc[:, 0].astype(str)
    norm = first.map(remove_accents).str.lower().str.strip()
    if is_xls:
        # For .xls we already trimmed after COOPERADO, so start at 0
        coop_start = 0
        cred_idx = norm[norm == "credenciado"].index
        cred_start = int(cred_idx[0]) if len(cred_idx) else None
        return (coop_start, cred_start)
    # For .xlsx/.xlsm, locate both markers
    coop_idx = norm[norm.isin({"cooperado", "coooperado"})].index
    coop_start = int(coop_idx[0] + 1) if len(coop_idx) else 0
    cred_idx = norm[norm == "credenciado"].index
    cred_start = None
    if len(cred_idx):
        # choose first occurrence after coop_start, if any
        for ix in cred_idx:
            if ix > coop_start:
                cred_start = int(ix)
                break
    return (coop_start, cred_start)


# -------- Export mapping (final CSV) --------
EXPORT_SPEC = [
    ("CRO", "", "text", []),
    ("COOPERADO", "NOME", "text", []),
    ("ATO COBERTO", "ATOS_COBERTOS", "float", ["ATPOS_COBERTOS", "ATOS COBERTOS"]),
    ("ATO COMPLEM.", "ATOS_NAO_COBERTOS", "float", ["ATO_NAO_COBERTO", "ATO NAO COBERTO", "ATOS NAO COBERTOS"]),
    ("OUTROS PAGTOS", "OUTROS_PAGAMENTOS", "float", ["OUTROS PAGAMENTOS"]),
    ("TOTAL GERAL", "TOTAL_BRUTO", "float", ["TOTAL BRUTO"]),
    ("DIVERSOS", "OUTROS_DESCONTOS", "float", ["OUTROS DESCONTOS"]),
    ("TX. ADM FATMOD", "", "float", []),
    ("INSS", "DESC_INSS", "float", ["DESCONTO INSS", "INSS_RET"]),
    ("IRRF", "IRRF", "float", []),
    ("Outras Fontes", "", "float", []),
    ("SALDO LÍQUIDO", "TOTAL_LIQUIDO", "float", ["TOTAL LIQUIDO"]),
    ("INSCRIÇÃO ISS", "", "float", ["INSCRICAO ISS", "NUM_ISS"]),
]


def build_export_df(recognized: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    cols_lower = {c.lower(): c for c in recognized.columns}
    for dest, source, typ, aliases in EXPORT_SPEC:
        selected = None
        candidates = [source] + aliases if source else []
        for c in candidates:
            if c in recognized.columns:
                selected = c
                break
            # try case-insensitive
            key = c.lower()
            if key in cols_lower:
                selected = cols_lower[key]
                break
        if selected is None:
            if typ == "text":
                out[dest] = ""
            else:
                out[dest] = pd.NA
        else:
            series = recognized[selected]
            if typ == "float":
                series = to_numeric(series)
            else:
                series = series.astype(str).fillna("")
            out[dest] = series
    # Drop rows where COOPERADO is empty or the literal string 'nan'
    if "COOPERADO" in out.columns:
        s = out["COOPERADO"].astype(str).str.strip().str.lower()
        out = out[~(s.eq("") | s.eq("nan") | s.eq("none"))]
    return out.reset_index(drop=True)


def render_mapping_controls(df: pd.DataFrame, auto_mapping: Dict[str, Optional[str]]):
    st.subheader("Mapeamento de colunas")
    options = [None] + list(df.columns)
    mapping: Dict[str, Optional[str]] = {}
    cols = st.columns(2)
    half = (len(REQUIRED_COLUMNS) + 1) // 2
    left_fields = REQUIRED_COLUMNS[:half]
    right_fields = REQUIRED_COLUMNS[half:]

    with cols[0]:
        for field in left_fields:
            mapping[field] = st.selectbox(
                f"{field}",
                options=options,
                index=(options.index(auto_mapping[field]) if auto_mapping[field] in options else 0),
                key=f"map_left_{field}",
                help=f"Selecione a coluna correspondente a '{field}'",
            )
    with cols[1]:
        for field in right_fields:
            mapping[field] = st.selectbox(
                f"{field}",
                options=options,
                index=(options.index(auto_mapping[field]) if auto_mapping[field] in options else 0),
                key=f"map_right_{field}",
                help=f"Selecione a coluna correspondente a '{field}'",
            )
    return mapping


def apply_user_mapping(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    return build_output_df(df, mapping)


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.write(
        "Faça upload de um ou mais arquivos Excel do relatório de produção e gere um CSV com separador ';' em UTF-8 (com BOM), com as colunas padronizadas."
    )

    uploaded_files = st.file_uploader(
        "Selecione arquivo(s) Excel",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Aguardando upload dos arquivos...")
        return

    decimal_comma = st.checkbox(
        "Usar vírgula como separador decimal (recomendado)", value=True
    )

    tabs = st.tabs([f"{f.name}" for f in uploaded_files])

    for file, tab in zip(uploaded_files, tabs):
        with tab:
            # Read raw file once (need to reuse bytes)
            file_bytes = file.getvalue()
            df = read_excel_safely(file_bytes, file.name)
            # Mark original row indices for later splitting into groups
            df_raw = df.copy()
            df_raw["__orig_idx"] = range(len(df_raw))
            # Fixed preset mapping (Curitiba) 0-based indices; empty names are ignored
            preset_pairs = [
                (0, "NOME"),
                (1, ""),
                (2, "CPFCNPJ"),
                (3, ""),
                (4, "INSS"),
                (5, ""),
                (6, ""),
                (7, ""),
                (8, ""),
                (9, "NUM_ISS"),
                (10, "ORCAMENTOS_TOTAIS"),
                (11, ""),
                (12, "COBERTOS"),
                (13, ""),
                (14, ""),
                (15, "ATOS_NAO_COBERTOS"),
                (16, ""),
                (17, "COPARTICIPACAO"),
                (18, ""),
                (19, ""),
                (20, "OUTROS_PAGAMENTOS"),
                (21, ""),
                (22, ""),
                (23, "TOTAL_BRUTO"),
                (24, ""),
                (25, "OUTROS_DESCONTOS"),
                (26, ""),
                (27, ""),
                (28, ""),
                (29, "INSS_RET"),
                (30, ""),
                (31, "IRRF"),
                (32, "ISS"),
                (33, "CSLL"),
                (34, "COFINS"),
                (35, "PIS"),
                (36, "TOTAL_LIQUIDO"),
            ]

            st.caption("Prévia do Excel (completa):")
            # Build cleaned data (no UI here; UI rendered in tabs below)
            df_clean = clean_rows_by_patterns(df_raw)
            if file.name.lower().endswith('.xls'):
                df_clean = clean_xls_specific_rules(df_clean)
            
            # Prepare groups (Cooperados/Credenciados) using markers in raw data
            is_xls = file.name.lower().endswith('.xls')
            coop_start, cred_start = find_section_indices(df, is_xls=is_xls)
            import numpy as np
            n_raw = len(df)
            coop_idx = np.arange(coop_start, cred_start if cred_start is not None else n_raw)
            cred_idx = np.arange((cred_start + 1) if cred_start is not None else n_raw, n_raw)
            if "__orig_idx" in df_clean.columns:
                coop_clean = df_clean[df_clean["__orig_idx"].isin(coop_idx)].drop(columns=["__orig_idx"], errors="ignore")
                cred_clean = df_clean[df_clean["__orig_idx"].isin(cred_idx)].drop(columns=["__orig_idx"], errors="ignore")
            else:
                coop_clean = df_clean
                cred_clean = df_clean.iloc[0:0]

            # Build recognized mapped dataframes depending on file type
            xlsx_out_coop = xlsx_out_cred = None
            xls_out_coop = xls_out_cred = None
            export_coop = export_cred = None

            if file.name.lower().endswith((".xlsx", ".xlsm")):
                # Reuse the XLSX mapping defined above in this function
                xlsx_pairs = [
                    (0, "NOME"),(1, ""),(2, "CPFCNPJ"),(3, ""),(4, "INSS"),(5, ""),(6, ""),(7, "ISS"),(8, ""),(9, ""),
                    (10, "ORCAMENTO_TOTAL"),(11, ""),(12, "ATOS_COBERTOS"),(13, ""),(14, "ATOS_NAO_COBERTOS"),(15, ""),
                    (16, "COPARTICIPACAO"),(17, ""),(18, "OUTROS_PAGAMENTOS"),(19, ""),(20, ""),(21, "TOTAL_BRUTO"),(22, ""),(23, ""),
                    (24, "OUTROS_DESCONTOS"),(25, ""),(26, "BASE_INSS"),(27, ""),(28, "DESC_INSS"),(29, ""),(30, "IRRF"),(31, "ISS"),
                    (32, "CSLL"),(33, "COFINS"),(34, "PIS"),(35, ""),(36, "TOTAL_LIQUIDO"),
                ]

                def build_xlsx_view(df_in: pd.DataFrame) -> pd.DataFrame:
                    out = pd.DataFrame()
                    boundary_idx = next((idx for idx, name in xlsx_pairs if name == "ORCAMENTO_TOTAL"), None)
                    if boundary_idx is None:
                        boundary_idx = float("inf")
                    for idx, name in xlsx_pairs:
                        if not name or not (0 <= idx < df_in.shape[1]):
                            continue
                        s = df_in.iloc[:, idx]
                        if idx >= boundary_idx:
                            s = to_numeric(s)
                        out[name if name not in out.columns else f"{name}_{idx}"] = s
                    return out

                xlsx_out_coop = build_xlsx_view(coop_clean)
                xlsx_out_cred = build_xlsx_view(cred_clean)
                export_coop = build_export_df(xlsx_out_coop)
                export_cred = build_export_df(xlsx_out_cred)
            elif file.name.lower().endswith(".xls"):
                xls_pairs = [
                    (0, "NOME"),(3, "CPFCNPJ"),(5, "INSS"),(8, "ISS"),(9, "ORCAMENTO_TOTAL"),(10, "ATOS_COBERTOS"),(11, "ATOS_NAO_COBERTOS"),
                    (12, "COPARTICIPACAO"),(14, "OUTROS_PAGAMENTOS"),(15, "TOTAL_BRUTO"),(16, "OUTROS_DESCONTOS"),(17, "BASE_INSS"),(19, "DESC_INSS"),
                    (22, "IRRF"),(23, "ISS"),(24, "CSLL"),(26, "COFINS"),(27, "PIS"),(28, "TOTAL_LIQUIDO"),
                ]

                def build_xls_view(df_in: pd.DataFrame) -> pd.DataFrame:
                    out = pd.DataFrame()
                    boundary_idx = next((idx for idx, name in xls_pairs if name == "ORCAMENTO_TOTAL"), None)
                    if boundary_idx is None:
                        boundary_idx = float("inf")
                    for idx, name in xls_pairs:
                        if not (0 <= idx < df_in.shape[1]):
                            continue
                        s = df_in.iloc[:, idx]
                        if idx >= boundary_idx:
                            s = to_numeric(s)
                        out[name if name not in out.columns else f"{name}_{idx}"] = s
                    return out

                xls_out_coop = build_xls_view(coop_clean)
                xls_out_cred = build_xls_view(cred_clean)
                export_coop = build_export_df(xls_out_coop)
                export_cred = build_export_df(xls_out_cred)

            # Render inner tabs per request
            tab1, tab2 = st.tabs(["Arquivo original + Exportação", "Processados"])

            with tab1:
                st.caption("Prévia do Excel (completa):")
                st.dataframe(df, use_container_width=True)
                raw_buf = io.StringIO(); df.to_csv(raw_buf, sep=';', index=False, encoding='utf-8', decimal=',')
                st.download_button("Baixar CSV (raw - Excel original)", ("\ufeff" + raw_buf.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_raw.csv", mime='text/csv', key=f"raw_{file.name}")
                st.divider(); st.subheader("Exportação Uniodonto (padrão final)")
                if export_coop is not None and export_cred is not None:
                    st.write("Cooperados")
                    st.dataframe(export_coop, use_container_width=True)
                    st.write("Credenciados")
                    st.dataframe(export_cred, use_container_width=True)
                    ebuf = io.StringIO(); ebuf2 = io.StringIO()
                    export_coop.to_csv(ebuf, sep=';', index=False, encoding='utf-8', decimal=',', float_format='%.2f')
                    export_cred.to_csv(ebuf2, sep=';', index=False, encoding='utf-8', decimal=',', float_format='%.2f')
                    st.download_button("Baixar CSV (Exportação - Cooperados)", ("\ufeff" + ebuf.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_export_cooperados.csv", mime='text/csv', key=f"export_coop_{file.name}")
                    st.download_button("Baixar CSV (Exportação - Credenciados)", ("\ufeff" + ebuf2.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_export_credenciados.csv", mime='text/csv', key=f"export_cred_{file.name}")

            with tab2:
                st.subheader("Após limpeza de linhas (apenas remoção de cabeçalhos/rodapés)")
                st.dataframe(df_clean, use_container_width=True)
                clean2 = io.StringIO(); df_clean.to_csv(clean2, sep=';', index=False, encoding='utf-8', decimal=',')
                st.download_button("Baixar CSV (apenas limpeza)", ("\ufeff" + clean2.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_limpo.csv", mime='text/csv', key=f"clean_{file.name}")
                st.divider()
                if xlsx_out_coop is not None:
                    st.subheader("XLSX: colunas reconhecidas (mapa fixo)")
                    st.write("Cooperados (mapeado)")
                    st.dataframe(xlsx_out_coop, use_container_width=True)
                    st.write("Credenciados (mapeado)")
                    st.dataframe(xlsx_out_cred, use_container_width=True)
                    xbuf = io.StringIO(); xbuf2 = io.StringIO()
                    xlsx_out_coop.to_csv(xbuf, sep=';', index=False, encoding='utf-8', decimal=',', float_format='%.2f')
                    xlsx_out_cred.to_csv(xbuf2, sep=';', index=False, encoding='utf-8', decimal=',', float_format='%.2f')
                    st.download_button("Baixar CSV (XLSX mapeado - Cooperados)", ("\ufeff" + xbuf.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_xlsx_mapeado_cooperados.csv", mime='text/csv', key=f"xlsx_map_coop_{file.name}")
                    st.download_button("Baixar CSV (XLSX mapeado - Credenciados)", ("\ufeff" + xbuf2.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_xlsx_mapeado_credenciados.csv", mime='text/csv', key=f"xlsx_map_cred_{file.name}")
                if xls_out_coop is not None:
                    st.subheader("XLS: colunas reconhecidas (mapa fixo)")
                    st.write("Cooperados (mapeado)")
                    st.dataframe(xls_out_coop, use_container_width=True)
                    st.write("Credenciados (mapeado)")
                    st.dataframe(xls_out_cred, use_container_width=True)
                    sbuf = io.StringIO(); sbuf2 = io.StringIO()
                    xls_out_coop.to_csv(sbuf, sep=';', index=False, encoding='utf-8', decimal=',', float_format='%.2f')
                    xls_out_cred.to_csv(sbuf2, sep=';', index=False, encoding='utf-8', decimal=',', float_format='%.2f')
                    st.download_button("Baixar CSV (XLS mapeado - Cooperados)", ("\ufeff" + sbuf.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_xls_mapeado_cooperados.csv", mime='text/csv', key=f"xls_map_coop_{file.name}")
                    st.download_button("Baixar CSV (XLS mapeado - Credenciados)", ("\ufeff" + sbuf2.getvalue()).encode('utf-8'), file_name=file.name.rsplit(".",1)[0]+"_xls_mapeado_credenciados.csv", mime='text/csv', key=f"xls_map_cred_{file.name}")
            
if __name__ == "__main__":
    main()
