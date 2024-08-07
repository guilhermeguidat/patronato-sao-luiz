import pyautogui
import time
import subprocess

def open_control_panel():
    # Abre o Painel de Controle para configuração de data e hora
    subprocess.run(['control', 'timedate.cpl'], shell=True)
    time.sleep(2)  # Aguarda a janela abrir

def configure_time_server(server):
    # Aperta TAB 5 vezes para focar na aba "Horário da Internet"
    for _ in range(4):
        pyautogui.press('tab')
        time.sleep(0.5)
    
    # Aperta seta para a direita 2 vezes para selecionar a aba "Horário da Internet"
    pyautogui.press('right', presses=2, interval=0.2)
    time.sleep(1.5)  # Aguarda a mudança de aba

    # Aperta TAB mais uma vez para focar no botão "Alterar configurações..."
    pyautogui.press('tab')
    time.sleep(0.5)

    # Aperta ESPAÇO para abrir a janela de configuração de servidor de tempo
    pyautogui.press('space')
    time.sleep(2.0)  # Aguarda a janela abrir

    # Aperta TAB mais uma vez para focar na lista de servidores
    pyautogui.press('tab')
    time.sleep(0.5)

    # Apaga o texto atual e digita o novo servidor
    pyautogui.hotkey('ctrl', 'a')  # Seleciona todo o texto
    pyautogui.write(server)
    time.sleep(1.0)

    # Aperta TAB novamente para focar no botão "Atualizar agora"
    pyautogui.press('tab')
    time.sleep(0.5)

    # Aperta ESPAÇO para aplicar as configurações
    pyautogui.press('space')
    time.sleep(2.0)  # Aguarda a aplicação das configurações

    # Aperta TAB para focar no botão "OK"
    pyautogui.press('tab')
    time.sleep(0.5)

    # Aperta ESPAÇO para confirmar e fechar a janela
    pyautogui.press('space')
    time.sleep(1.5)
    
    # Fecha a janela de Data e Hora
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1.0)  # Aguarda a janela fechar

    print(f"Servidor de tempo configurado para: {server}")
