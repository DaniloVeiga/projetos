import os
import shutil

PASTA_PROTEGIDA = r"C:\Users\DaniloBarrosodeVeiga\OneDrive - Zeentech Brasil\Área de Trabalho\projeto\ilustracoes\base_compilada"
PASTA_BACKUP = r".\base_bkp"

def criar_backup():
    os.makedirs(PASTA_BACKUP, exist_ok=True)

    for arquivo in os.listdir(PASTA_PROTEGIDA):
        origem = os.path.join(PASTA_PROTEGIDA, arquivo)

        if os.path.isfile(origem):
            destino = os.path.join(PASTA_BACKUP, arquivo)

            try:
                shutil.copy2(origem, destino)
                print(f"Backup criado: {arquivo}")

            except PermissionError:
                print(f"Arquivo em uso, ignorado: {arquivo}")

            except Exception as e:
                print(f"Erro em {arquivo}: {e}")

criar_backup()