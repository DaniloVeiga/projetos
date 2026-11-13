import pyautogui as pgui
import time
def painel_controle():
    #-------
    #Insumos
    #-------
    apertar_enter = lambda tecla: pgui.press(tecla)
    apertar_espaco = lambda tecla: pgui.press(tecla)
    executar = lambda: pgui.hotkey('win','r')
    painel = lambda: (pgui.write('control'), apertar_enter('enter'))
    atalho_voltar = lambda *atalho: pgui.hotkey(*atalho)
    tela001 = lambda: (time.sleep(3), [pgui.press('tab') for a in range(14)], apertar_enter('enter'))
    tela002 = lambda:(time.sleep(5), [atalho_voltar('shift','tab')])
    tela003 = lambda:(time.sleep(4),apertar_enter('enter'),time.sleep(3), [pgui.press('tab') for b in range(4)])
    tela004 = lambda:(time.sleep(3), [pgui.press('down') for a in range(2)], apertar_enter('enter'), time.sleep(3), apertar_espaco('space'))
    #-------
    #Navegação
    #-------
    executar()
    painel()
    tela001()
    tela002()
    tela003()
    tela004()
            
painel_controle()