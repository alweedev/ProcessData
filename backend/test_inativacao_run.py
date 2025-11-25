import sys
import traceback
import pandas as pd
import os

# garantir que o pacote backend está importável quando executado a partir da raiz
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)

try:
    from backend.processor import processar_inativacao_from_paths
    from backend.processor import MODEL_COLS
except Exception as e:
    print("FALHA ao importar processar_inativacao_from_paths:", e)
    traceback.print_exc()
    sys.exit(2)


def show_result(desc, out, stats):
    print('\n' + '='*40)
    print(desc)
    print('stats:', stats)
    print('out rows:', None if out is None else out.shape[0])
    if out is not None and not out.empty:
        print(out.head(10).to_dict(orient='records'))


# caso 1: CPF exato
try:
    df_base = pd.DataFrame([
        {"CPF": "111.222.333-44", "NomeCompleto":"Maria Silva", "Status":"ATIVO", "Solicitante":"S", "Terceiro":"N"},
        {"CPF": "222.333.444-55", "NomeCompleto":"Joao Pereira", "Status":"ATIVO", "Solicitante":"N", "Terceiro":"S"},
        {"CPF": "333.444.555-66", "NomeCompleto":"Mariana Costa", "Status":"INATIVO", "Solicitante":"N", "Terceiro":"N"},
    ])
    df_lista = pd.DataFrame([{"CPF":"11122233344"}])
    out, stats = processar_inativacao_from_paths(df_base, df_lista, use_fuzzy=False)
    show_result('CPF exato', out, stats)
except Exception as e:
    print('Erro no caso CPF exato:', e)
    traceback.print_exc()

# caso 2: Nome completo exato
try:
    df_lista2 = pd.DataFrame([{"NomeCompleto":"Maria Silva"}])
    out2, stats2 = processar_inativacao_from_paths(df_base, df_lista2, use_fuzzy=False)
    show_result('Nome completo exato', out2, stats2)
except Exception as e:
    print('Erro no caso Nome completo:', e)
    traceback.print_exc()

# caso 3: nome genérico que antes causava muitos matches
try:
    df_lista3 = pd.DataFrame([{"NomeCompleto":"Maria"}])
    out3, stats3 = processar_inativacao_from_paths(df_base, df_lista3, use_fuzzy=False)
    show_result('Nome genérico ("Maria")', out3, stats3)
except Exception as e:
    print('Erro no caso Nome genérico:', e)
    traceback.print_exc()

print('\nTeste concluído')
