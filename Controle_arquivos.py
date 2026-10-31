import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime
import getpass
import time

def desbloqueio():
    lista_pastas = [r'.\base_consulta',r'.\base_compilada',r'.\base_convertida',r'.\base_sql',r'base_compilada_adq_filtrada.xlsx']
    for pasta in lista_pastas:
        subprocess.run(f'icacls {pasta} /inheritance:e')

        print("Permissões aplicadas.")

desbloqueio()
print('-')
print('Pastas Desbloqueadas com sucesso.')
print('Monitoramento de arquivos em execução...')
print('...para encerrar o processo de monitoramento basta utilizar o atalho "CTRL+C".')
PASTA = r"C:\Users\DaniloBarrosodeVeiga\OneDrive - Zeentech Brasil\Área de Trabalho\projeto\ilustracoes"
ARQUIVO_LOG = "monitoramento.txt"

class Monitor(FileSystemEventHandler):

    def on_deleted(self, event):
        registro = (
            f"{datetime.now()} | "
            f"Usuário: {getpass.getuser()} | "
            f"Arquivo: {event.src_path}\n"
        )

        print(registro.strip())

        with open(ARQUIVO_LOG, "a", encoding="utf-8") as log:
            log.write(registro)

observer = Observer()
observer.schedule(Monitor(), PASTA, recursive=True)
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()