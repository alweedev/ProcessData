import io
import sys
import os
import pandas as pd
import traceback

# garantir import do package
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)

try:
    import backend.app as backend_app
    app = backend_app.app
except Exception as e:
    print('Falha ao importar backend.app:', e)
    traceback.print_exc()
    sys.exit(2)


def make_excel_bytes(df: pd.DataFrame):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# construir base e lista
df_base = pd.DataFrame([
    {"CPF": "11122233344", "NomeCompleto": "Maria Silva", "Status": "ATIVO", "Solicitante": "S", "Terceiro": "N"},
    {"CPF": "22233344455", "NomeCompleto": "Joao Pereira", "Status": "ATIVO", "Solicitante": "N", "Terceiro": "S"},
    {"CPF": "33344455566", "NomeCompleto": "Mariana Costa", "Status": "INATIVO", "Solicitante": "N", "Terceiro": "N"}
])

# lista como arquivo Excel
df_lista = pd.DataFrame([
    {"CPF": "11122233344", "NomeCompleto": ""}
])

base_bytes = make_excel_bytes(df_base)
lista_bytes = make_excel_bytes(df_lista)

print('Iniciando teste de integração: enviando arquivos para /api/preview_inativacao')
with app.test_client() as client:
    data = {
        'base': (base_bytes, 'base.xlsx'),
        'lista': (lista_bytes, 'lista.xlsx'),
        'use_fuzzy': 'false'
    }
    resp = client.post('/api/preview_inativacao', data=data, content_type='multipart/form-data')
    print('preview status_code =', resp.status_code)
    try:
        json = resp.get_json()
    except Exception:
        json = None
    print('preview json:', json)

    print('\nEnviando para /api/process_inativacao (espera-se arquivo xlsx)')
    # recriar bytes since BytesIO may have been consumed
    base_bytes = make_excel_bytes(df_base)
    lista_bytes = make_excel_bytes(df_lista)
    data2 = {'base': (base_bytes, 'base.xlsx'), 'lista': (lista_bytes, 'lista.xlsx'), 'use_fuzzy': 'false'}
    resp2 = client.post('/api/process_inativacao', data=data2, content_type='multipart/form-data')
    print('process status_code =', resp2.status_code)
    ct = resp2.headers.get('Content-Type')
    print('Content-Type:', ct)
    if resp2.status_code == 200 and 'application' in (ct or ''):
        # salvar arquivo recebido
        out_path = os.path.join(os.path.dirname(__file__), 'saida_test_inativacao.xlsx')
        with open(out_path, 'wb') as f:
            f.write(resp2.data)
        print('Arquivo recebido salvo em', out_path)
    else:
        try:
            print('process json:', resp2.get_json())
        except Exception:
            print('process response text:', resp2.get_data(as_text=True))

print('Teste de integração finalizado')
