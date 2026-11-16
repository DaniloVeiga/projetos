import pyautogui as pgui
import time
import subprocess as sbp

def painel_controle():
    #-------
    #Insumos
    #-------
    apertar_espaco = lambda tecla: pgui.press(tecla)
    executar = lambda: sbp.run('start ms-settings:notifications', shell=True)
    #acesso direto a configuração de notificações
    tela004 = lambda:(time.sleep(5), apertar_espaco('space'))
    #-------
    #Navegação
    #-------
    executar()
    tela004()
            
painel_controle()