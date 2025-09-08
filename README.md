# Conversor Uniodonto: Excel → CSV (; UTF-8)

Aplicativo em Streamlit para converter relatórios de produção (Excel) em CSV padronizado com separador `;` e codificação UTF‑8 com BOM (compatível com Excel), contendo as colunas:

```
CRO ; COOPERADO ; ATO COBERTO ; ATO COMPLEM. ; OUTROS PAGTOS ; TOTAL GERAL ; DIVERSOS ; TX. ADM FATMOD ; INSS ; IRRF ; Outras Fontes ; SALDO LÍQUIDO ; INSCRIÇÃO ISS
```

## Como executar

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Rode o app:
   ```bash
   streamlit run app.py
   ```

## Uso

- Faça upload de um ou mais arquivos Excel.
- Revise/ajuste o mapeamento das colunas para o padrão exigido (o app tenta detectar automaticamente pelos nomes).
- Clique em “Converter este arquivo para CSV” em cada aba para gerar e baixar o CSV.
- Opcional: após converter dois ou mais arquivos, gere um CSV combinado com todos os registros.

## Notas

- O app grava os números com vírgula decimal por padrão (opção configurável na interface) e utiliza `UTF-8 com BOM` para preservar acentos no Excel.
- Se alguma coluna não existir no Excel, ela será preenchida com vazio (texto) ou `0` (número). Você pode ajustar o mapeamento manualmente antes da conversão.
- O leitor de Excel do pandas aceita `.xlsx`, `.xlsm`, `.xls` e `.xlsb` (para `.xlsb`, é necessário `pyxlsb`).
