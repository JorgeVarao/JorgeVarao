# -*- coding: utf-8 -*-
"""Extrai os dados brutos dos dois arquivos para um pickle intermediário."""
import openpyxl, pickle

SEDUC="/root/.claude/uploads/41fbd151-f5e2-530d-88f9-9b26c68b08f9/6c7f4aae-Proposta_Reordenamento_2027_SEDUC__PI.xlsx"
REORD="/root/.claude/uploads/41fbd151-f5e2-530d-88f9-9b26c68b08f9/be0ac5a9-reordenamento_2027.xlsx"

def leia(path, aba):
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws=wb[aba]
    out=[list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    # remove linhas totalmente vazias no fim
    while out and all(v in (None,"") for v in out[-1]): out.pop()
    return out

dados={}
for aba in ["Turmas","Base GDI","Panorama Municipal","Fusão Turmas","Base Tratada","Base Anexos","Validações","Reordenamento 2027"]:
    dados["SEDUC:"+aba]=leia(SEDUC,aba)
    print("SEDUC",aba,len(dados["SEDUC:"+aba]))
for aba in ["Matriculas por etapa","Cursos","Panorama Municipal","Fusão de Turmas","Base Reordenamento 2027","LISTAS"]:
    dados["REORD:"+aba]=leia(REORD,aba)
    print("REORD",aba,len(dados["REORD:"+aba]))

pickle.dump(dados, open("dados.pkl","wb"))
print("ok")
