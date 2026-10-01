# -*- coding: utf-8 -*-
"""
gerar_modelo.py — cria um MODELO LIMPO (sem dados de cliente) a partir do
arquivo real da planilha, e RECUSA se sobrar qualquer dado.

O que ele apaga:
  - CADASTRO: nomes, CNPJ, inscrições e SENHAS (linhas 3+, todas as colunas)
  - 1. FOLHA / 2. SPED / 3. FATURAMENTO / 4. CONSULTAS: todas as linhas de dados
  - COMPARATIVO: as linhas de empresa (entre "EMPRESA" e "TOTAL GERAL"),
    se ainda existirem — o bloco de faturamento foi removido em 2026-09
    (removerFaturamentoDoComparativo); em arquivo novo não há o que limpar
  - VENCIMENTOS: os prazos, se a aba ainda existir (era do calendário,
    removido em 2026-09; mantém a lista de categorias em F:G)
  - NOTAS: todas as anotações

O que ele mantém: cabeçalhos, fórmulas, listas suspensas e a estrutura.

Uso:
  py gerar_modelo.py "caminho\\da\\planilha_real.xlsx"
  (sem argumento, ele pega o .xlsx mais recente da pasta que não seja o modelo)

Saída: CONTROLE_EMPRESAS_MODELO.xlsx  (na mesma pasta)
"""
import sys, os, glob, warnings
warnings.filterwarnings("ignore")          # openpyxl avisa sobre CF avançada
import openpyxl

PASTA  = os.path.dirname(os.path.abspath(__file__))
MODELO = os.path.join(PASTA, "CONTROLE_EMPRESAS_MODELO.xlsx")
ETAPAS = ["1. FOLHA", "2. SPED", "3. FATURAMENTO", "4. CONSULTAS"]


def limpar(ws, lin_ini, col_ini, col_fim):
    for r in range(lin_ini, ws.max_row + 1):
        for c in range(col_ini, col_fim + 1):
            ws.cell(row=r, column=c).value = None


def achar(ws, texto, col=2):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=col).value
        if v and str(v).strip().upper() == texto.upper():
            return r
    return 0


def escolher_entrada():
    if len(sys.argv) >= 2:
        return sys.argv[1]
    cands = [f for f in glob.glob(os.path.join(PASTA, "*.xlsx"))
             if os.path.basename(f) != os.path.basename(MODELO)]
    if not cands:
        return None
    return max(cands, key=os.path.getmtime)      # o mais recente


def contar(ws, col, ini=3):
    return sum(1 for r in range(ini, ws.max_row + 1)
               if ws.cell(row=r, column=col).value not in (None, ""))


def main():
    entrada = escolher_entrada()
    if not entrada or not os.path.exists(entrada):
        print("Não achei o arquivo de entrada.")
        print('Uso: py gerar_modelo.py "caminho\\da\\planilha_real.xlsx"')
        return 1
    print("Entrada:", entrada)

    wb = openpyxl.load_workbook(entrada)

    if "CADASTRO" in wb.sheetnames:
        # todas as colunas: o CADASTRO pode ganhar coluna nova (a SENHA muda de lugar)
        limpar(wb["CADASTRO"], 3, 1, wb["CADASTRO"].max_column)
    for nome in ETAPAS:
        if nome in wb.sheetnames:
            limpar(wb[nome], 3, 1, wb[nome].max_column)
    if "COMPARATIVO" in wb.sheetnames:
        ws = wb["COMPARATIVO"]
        ini, fim = achar(ws, "EMPRESA"), achar(ws, "TOTAL GERAL")
        if ini and fim and fim > ini + 1:
            for r in range(ini + 1, fim):
                for c in range(2, ws.max_column + 1):
                    ws.cell(row=r, column=c).value = None
    if "VENCIMENTOS" in wb.sheetnames:
        limpar(wb["VENCIMENTOS"], 3, 1, 4)       # prazos; mantém categorias F:G
    if "NOTAS" in wb.sheetnames:
        limpar(wb["NOTAS"], 2, 1, wb["NOTAS"].max_column)

    wb.save(MODELO)

    # ---- verificação: recusa se sobrou qualquer dado de cliente ----
    v = openpyxl.load_workbook(MODELO, data_only=True)
    problemas = []
    if "CADASTRO" in v.sheetnames:
        c = v["CADASTRO"]
        # confere TODAS as colunas (nome no cabeçalho só para a mensagem)
        for col in range(1, c.max_column + 1):
            q = contar(c, col)
            if q:
                rot = c.cell(row=2, column=col).value or f"coluna {col}"
                problemas.append(f"CADASTRO/{rot}: {q}")
    for nome in ETAPAS:
        if nome in v.sheetnames and contar(v[nome], 2):
            problemas.append(f"{nome}: {contar(v[nome], 2)} empresa(s)")

    if problemas:
        os.remove(MODELO)
        print("\n[X] SOBROU DADO — modelo NAO gerado:")
        for p in problemas:
            print("   -", p)
        return 1

    print("\n[OK] Modelo limpo gerado e verificado (sem dados de cliente):")
    print("     ", MODELO)
    print("     Abas:", v.sheetnames)
    return 0


if __name__ == "__main__":
    sys.exit(main())
