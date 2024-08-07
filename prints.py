import os
import subprocess
import pyautogui
from PIL import ImageGrab, ImageChops, Image
import time
import pygetwindow as gw
from ctypes import windll
import sys


def capture_screenshot(filename):
    time.sleep(2)
    screenshot = ImageGrab.grab()
    screenshot.save(filename)
    print(f'Captura de tela salva como {filename}')

def open_and_capture_cmd(command, filename):
    subprocess.Popen('start cmd', shell=True)
    time.sleep(5)
    pyautogui.hotkey('win', 'up')
    time.sleep(2)
    pyautogui.typewrite(command + '\n')
    time.sleep(5)
    capture_screenshot(filename)
    pyautogui.hotkey('alt', 'f4')

def are_images_equal(img1_path, img2_path):
    img1 = Image.open(img1_path)
    img2 = Image.open(img2_path)
    diff = ImageChops.difference(img1, img2)
    return diff.getbbox() is None

def open_and_capture_programs(target_dir):
    subprocess.Popen('control appwiz.cpl', shell=True)
    time.sleep(5)
    pyautogui.hotkey('win', 'up')
    time.sleep(2)
    previous_screenshot_path = os.path.join(target_dir, 'PROGRAMAS 1.png')
    capture_screenshot(previous_screenshot_path)
    program_list_count = 1
    max_attempts = 4
    while program_list_count < max_attempts:
        pyautogui.press('pgdn')
        time.sleep(2)
        new_screenshot_path = os.path.join(target_dir, f'PROGRAMAS {program_list_count+1}.png')
        capture_screenshot(new_screenshot_path)
        if are_images_equal(previous_screenshot_path, new_screenshot_path):
            os.remove(new_screenshot_path)
            break
        previous_screenshot_path = new_screenshot_path
        program_list_count += 1
    for i in range(1, program_list_count):
        img1_path = os.path.join(target_dir, f'PROGRAMAS {i}.png')
        img2_path = os.path.join(target_dir, f'PROGRAMAS {i+1}.png')
        if os.path.exists(img2_path) and are_images_equal(img1_path, img2_path):
            os.remove(img2_path)

    pyautogui.hotkey('alt', 'f4')

def open_activation_menu():
    subprocess.run('start ms-settings:activation', shell=True)

def run_command_in_background(command):
    process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return process

def move_window(title, x_offset, y_offset):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        window = windows[0]
        window.restore()
        window.activate()
        time.sleep(1)
        hwnd = window._hWnd
        windll.user32.SetWindowPos(hwnd, 0, window.left + x_offset, window.top + y_offset, 0, 0, 0x0001 | 0x0002)
        print(f'Janela com título "{title}" movida por ({x_offset}, {y_offset}).')

def close_window(title):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        window = windows[0]
        window.close()
        print(f'Janela com título "{title}" fechada.')

def obter_diretorio_executavel():
    """Obtém o diretório onde o executável está localizado."""
    if getattr(sys, 'frozen', False):
        # Se o script estiver em um ambiente congelado (PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # Se o script estiver sendo executado como um script Python
        return os.path.dirname(os.path.abspath(__file__))       

def start_capture(folder_name, capture_mac, capture_system, capture_serial, capture_programs, capture_activation):
    script_dir = obter_diretorio_executavel()
    target_dir = os.path.join(script_dir, folder_name)
    
    try:
        os.makedirs(target_dir, exist_ok=True)
        print(f"Pasta '{target_dir}' criada com sucesso.")
    except Exception as e:
        print(f"Erro ao criar a pasta '{target_dir}': {e}")
        return
    
    if capture_mac:
        try:
            open_and_capture_cmd('getmac', os.path.join(target_dir, 'MAC.png'))
        except Exception as e:
            print(f"Erro ao capturar MAC: {e}")
    
    if capture_system:
        try:
            subprocess.Popen('msinfo32', shell=True)
            time.sleep(5)
            pyautogui.hotkey('win', 'up')
            time.sleep(2)
            capture_screenshot(os.path.join(target_dir, 'INFO.png'))
            subprocess.Popen('taskkill /IM msinfo32.exe /F', shell=True)
        except Exception as e:
            print(f"Erro ao capturar informações do sistema: {e}")

    if capture_serial:
        try:
            open_and_capture_cmd('wmic bios get serialnumber', os.path.join(target_dir, 'SERIAL.png'))
        except Exception as e:
            print(f"Erro ao capturar serial: {e}")
    
    if capture_programs:
        try:
            open_and_capture_programs(target_dir)
        except Exception as e:
            print(f"Erro ao capturar programas instalados: {e}")

    if capture_activation:
        try:
            open_activation_menu()
            time.sleep(5)
            process_dli = run_command_in_background('slmgr /dli')
            time.sleep(5)
            move_window('Windows Script Host', 300, 0)
            capture_screenshot(os.path.join(target_dir, 'WIN 1.png'))
            close_window('Windows Script Host')
            time.sleep(2)
            process_xpr = run_command_in_background('slmgr /xpr')
            time.sleep(5)
            move_window('Windows Script Host', 300, 0)
            capture_screenshot(os.path.join(target_dir, 'WIN 2.png'))
            close_window('Windows Script Host')
            close_window('Configurações')
        except Exception as e:
            print(f"Erro ao capturar ativação do Windows: {e}")

    print(f"Screenshots capturadas com sucesso e salvas na pasta {folder_name}.")

