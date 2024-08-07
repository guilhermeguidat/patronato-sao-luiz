import subprocess
import os
import time
from pdf_converter import convert_txt_to_pdf

def execute_powershell_script(script_path):
    command = ['powershell.exe', '-File', script_path]
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        print("Saída:", result.stdout)
        print("Erros:", result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar o script PowerShell: {e}")
        print("Saída:", e.stdout)
        print("Erros:", e.stderr)

def bitlocker_process(bitlocker_info_path, powershell_script_path):
    execute_powershell_script(powershell_script_path)
    time.sleep(5)  # Aguarda 5 segundos

    if not os.path.exists(bitlocker_info_path):
        print(f"Arquivo {bitlocker_info_path} não encontrado.")
        return

    nome_agr = input("Digite o nome do AGR: ")
    nome_ar = input("Digite o nome da AR: ")

    pdf_name = f"CHAVE BITLOCKER {nome_agr} AR {nome_ar}.pdf"
    pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), pdf_name)

    convert_txt_to_pdf(bitlocker_info_path, pdf_path)
    print(f"Arquivo PDF gerado: {pdf_path}")
