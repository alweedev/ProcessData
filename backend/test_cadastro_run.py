import sys
import traceback
import pandas as pd
import os

root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)

try:
    from backend.processor import processar_registros_from_files
except Exception as e:
    print('Falha ao importar:', e)
    traceback.print_exc()
    sys.exit(2)

# criar arquivo Excel fictício em memória usando DataFrame com valores variados
rows = [
    {"CPF": "11122233344", "NomeCompleto": "Ana Souza", "Solicitante": "Sim", "Terceiro": "Sim"},
    {"CPF": "22233344455", "NomeCompleto": "Bruno Lima", "Solicitante": "", "Terceiro": "Não"},
    {"CPF": "33344455566", "NomeCompleto": "Carlos Dias", "Solicitante": None, "Terceiro": "S"}
]

# o processar_registros_from_files recebe paths; vamos escrever um xlsx temporário
import tempfile
from pathlib import Path
import pandas as pd

df = pd.DataFrame(rows)
fd, path = tempfile.mkstemp(suffix='.xlsx')
os.close(fd)
df.to_excel(path, index=False)

try:
    errors, out = processar_registros_from_files([path], login_choice='CPF', fluxo='SELF')
    print('errors:', errors)
    print('out sample:', out.head().to_dict(orient='records'))
except Exception as e:
    print('Erro no teste cadastro:', e)
    traceback.print_exc()
finally:
    try:
        os.remove(path)
    except Exception:
        pass
