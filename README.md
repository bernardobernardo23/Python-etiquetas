﻿# Python-etiquetas
ARQUIVO EXECUTAVEL
def recurso_caminho(relativo):
    """Retorna o caminho absoluto do recurso, lidando com o ambiente do PyInstaller."""
    if getattr(sys, 'frozen', False):  # Executável gerado
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relativo)

credenciais_path = recurso_caminho(".json")
logo_path = recurso_caminho(".png")

CREDS = Credentials.from_service_account_file(credenciais_path, scopes=SCOPES)


ARQUIVO PARA RODAR

CREDS = Credentials.from_service_account_file(".json", scopes=SCOPES)
logo_path = ".png"

CMD 
pyinstaller --onefile --noconsole --add-data "logo.png;." --add-data "credencial.json;." gerador_de_etiquetas.py
