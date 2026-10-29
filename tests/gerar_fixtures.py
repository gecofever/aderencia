"""Gera planilhas sintéticas com o mesmo layout esperado por acabamento_stream.py."""
import pandas as pd, numpy as np

# ---- Saída de prensas: abas 'Stábil UGB 1'/'Stábil UGB 2' com header na linha 3 ----
def planilha_prensas(path):
    linhas_p1 = [
        ["OP Acabado", "Data Stábil", "Qtd. Par Real", "Cód. Artigo", "Desc. Artigo"],
        [1001, "2026-10-01", 100, 555, "ARTIGO A"],          # lançamento no dia do limite com hora > 00:00
        [1002, None,        200, 556, "ARTIGO B"],           # sem Data Stábil
        [1003, "2026-10-05", 300, 557, "ARTIGO C"],          # sem lançamento no CD
        [1004, "2026-10-08", 400, 558, "ARTIGO D"],          # prazo cai em dias sem produção
        [1005, "2026-10-02", 500, 559, "ARTIGO E"],          # prazo cai em dia marcado '0' como TEXTO
        [1006, "2026-10-06", 600, 560, "ARTIGO F"],          # mesma OP com dois artigos
        [1006, "2026-10-06", 50,  561, "ARTIGO G"],
        [1007, None,        900, 563, "ARTIGO I"],           # sem Data Stábil, MAS com lançamento no CD
    ]
    linhas_p2 = [
        ["OP Acabado", "Data Stábil", "Qtd. Par Real", "Cód. Artigo", "Desc. Artigo"],
        [2001, "2026-10-09", 700, 562, "ARTIGO H"],
    ]
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        # header=2 => duas linhas descartáveis antes do cabeçalho
        pd.DataFrame([["relatório", "pcp"], ["gerado", "hoje"]]).to_excel(xl, sheet_name="Stábil UGB 1", header=False, index=False)
        pd.DataFrame(linhas_p1).to_excel(xl, sheet_name="Stábil UGB 1", header=False, index=False, startrow=2)
        pd.DataFrame(linhas_p2).to_excel(xl, sheet_name="Stábil UGB 2", header=False, index=False, startrow=2)
        # aba Balanceado (calendário), header=2
        cal = [["Data", "6x1"],
               ["2026-10-01", 1], ["2026-10-02", 1], ["2026-10-03", 1],
               ["2026-10-04", 0],           # sem produção (numérico)
               ["2026-10-05", "0"],         # sem produção, mas gravado como TEXTO
               ["2026-10-06", 0],           # sem produção
               ["2026-10-07", 1], ["2026-10-08", 1], ["2026-10-09", 1], ["2026-10-10", 1],
               ["2026-10-11", 0], ["2026-10-12", 0], ["2026-10-13", 1]]
        pd.DataFrame([["calendário", "outubro"]]).to_excel(xl, sheet_name="Balanceado", header=False, index=False)
        pd.DataFrame(cal).to_excel(xl, sheet_name="Balanceado", header=False, index=False, startrow=2)

# ---- YMB51 (entrada de CD) ----
def planilha_ymbt51(path):
    df = pd.DataFrame([
        ["Ordem", "Material", "Qtd.  UM registro", "Data de lançamento"],
        [1001, 555, 100, "2026-10-04 10:30:00"],   # ENTROU no dia do limite, mas com hora
        [1005, 559, 500, "2026-10-06 08:00:00"],
        [1006, 560, 600, "2026-10-09 09:00:00"],
        [2001, 562, 700, "2026-10-13 23:00:00"],
        [1007, 563, 900, "2026-10-02 07:00:00"],   # entrou no CD, mas a OP não tem Data Stábil
    ])
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="Sheet1", header=False, index=False)

if __name__ == "__main__":
    planilha_prensas("/home/hermes/.hermes/cache/scratch/aderencia_local/Saida de prensas.xlsx")
    planilha_ymbt51("/home/hermes/.hermes/cache/scratch/aderencia_local/10.outubro.xlsx")
    print("ok")
