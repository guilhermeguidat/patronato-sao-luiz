import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import shutil
import sys
import ctypes
import winreg as reg
import time
from fpdf import FPDF
from time_configuration import configure_time_server, open_control_panel

from prints import start_capture

import pyautogui
from PIL import ImageGrab, ImageChops, Image
import pygetwindow as gw
from ctypes import windll

from tkinter import messagebox, Tk

# Caminhos
current_dir = os.path.dirname(os.path.abspath(__file__))
powershell_script_path = os.path.join(current_dir, 'script', 'EnableBitLocker.ps1')
bitlocker_info_path = r'C:\BitLockerRecoveryInfo.txt'
new_exe_name = "PRINTS.exe"  # Nome do novo arquivo .exe a ser executado

# Senhas de administrador disponíveis
admin_passwords = {
    "AR Própria": "ARcm@2050",
    "Outras": "ARcert1127"
}

def execute_powershell_script(script_path):
    activate_ps_script = 'Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force'
    command = ['powershell.exe', '-Command', f"{activate_ps_script}; . '{script_path}'"]
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        print("Script PowerShell executado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar o script PowerShell:\n{e.stderr}")

def convert_txt_to_pdf(txt_path, pdf_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    with open(txt_path, 'r') as file:
        for line in file:
            pdf.cell(200, 10, txt=line.strip(), ln=True)
    pdf.output(pdf_path)

def install_program(executable, silent_args):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        drivers_dir = os.path.join(script_dir, 'drivers')
        executable_path = os.path.join(drivers_dir, executable)
        if not os.path.isfile(executable_path):
            print(f"Erro: {executable_path} não encontrado.")
            return
        print(f"Iniciando a instalação de {executable}...")
        subprocess.run([executable_path] + silent_args, check=True)
        print(f"{executable} instalado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Falha na instalação de {executable}:\n{e}")

def criar_usuario(usuario, senha, grupo):
    try:
        subprocess.run(['net', 'user', usuario, senha, '/add'], check=True)
        subprocess.run(['net', 'localgroup', grupo, usuario, '/add'], check=True)
        print(f"Usuário {usuario} criado e adicionado ao grupo {grupo}.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao criar o usuário {usuario}:\n{e}")

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        print(f"Erro ao verificar privilégios de administrador:\n{e}")
        return False

def run_as_admin():
    script = os.path.abspath(sys.argv[0])
    params = ' '.join(sys.argv[1:])
    try:
        subprocess.run(['powershell', '-Command', f"Start-Process python -ArgumentList '{script} {params}' -Verb RunAs"], check=True)
        sys.exit()
    except subprocess.CalledProcessError as e:
        print(f"Erro ao tentar reexecutar como administrador:\n{e}")
        sys.exit(1)

def alterar_registro(chave, subchave, valor_nome, valor, tipo=reg.REG_DWORD):
    try:
        chave_reg = reg.OpenKey(chave, subchave, 0, reg.KEY_SET_VALUE)
    except FileNotFoundError:
        chave_reg = reg.CreateKey(chave, subchave)
    try:
        reg.SetValueEx(chave_reg, valor_nome, 0, tipo, valor)
        reg.CloseKey(chave_reg)
        print(f"{valor_nome} configurado para {valor} em {subchave}.")
    except WindowsError as e:
        print(f"Erro ao alterar o registro:\n{e}")

def aplicar_secedit(config):
    try:
        with open('secedit.inf', 'w') as file:
            file.write(config)
        subprocess.run(['secedit', '/configure', '/db', 'secedit.sdb', '/cfg', 'secedit.inf'], check=True)
        print("Configurações de segurança aplicadas com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao aplicar configurações de segurança:\n{e}")

def install_programs():
    programs = {
        "AWP.exe": ["/silent"],
        "ePass2003-Setup_170623.exe": ["/silent"],
        "ExtensionModule.exe": ["/S"],
        "ftrDriverSetup_win8_whql_3471.exe": ["/quiet", "/norestart"],
        "GDsetupStarsignCUTx64.exe": ["/s"],
        "sac_iti.exe": ["/silent"],
        "SafeSignIC30124-x64-win-tu-user64.exe": ["/quiet"]
    }
    for executable, args in programs.items():
        install_program(executable, args)

def executar_configuracao_data_hora():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal
    messagebox.showinfo("Aviso", "Você está prestes a configurar o servidor de data e hora. Por favor, não mexa na máquina até que o processo esteja concluído.")
    root.destroy()  # Fechar a janela oculta

def obter_diretorio_executavel():
    return os.path.dirname(os.path.abspath(__file__))

def copiar_favoritos(nome_ar):
    pasta_fav = "FAV"  # Diretório onde os arquivos HTML estão localizados
    arquivo_padrao = "PADRÃO.html"
    
    # Obtém o diretório onde o script está sendo executado
    diretorio_script = obter_diretorio_executavel()
    caminho_pasta_fav = os.path.join(diretorio_script, pasta_fav)
    
    # Verifica se a pasta FAV existe
    if not os.path.exists(caminho_pasta_fav):
        print(f"Pasta {caminho_pasta_fav} não encontrada.")
        return

    # Procura pelo arquivo correspondente ao nome da AR
    arquivo_favoritos = None
    for arquivo in os.listdir(caminho_pasta_fav):
        if arquivo.lower() == f"{nome_ar.lower()}.html":
            arquivo_favoritos = arquivo
            break

    # Define o arquivo a ser copiado (correspondente ou padrão)
    if arquivo_favoritos:
        arquivo_origem = os.path.join(caminho_pasta_fav, arquivo_favoritos)
    else:
        arquivo_origem = os.path.join(caminho_pasta_fav, arquivo_padrao)

    # Define o caminho da pasta do script
    pasta_script = diretorio_script
    arquivo_favoritos_script = os.path.join(pasta_script, "favoritos.html")

    # Verifica se o arquivo de origem existe antes de tentar copiar
    if not os.path.isfile(arquivo_origem):
        print(f"Arquivo de origem '{arquivo_origem}' não encontrado.")
        return

    try:
        # Copia o arquivo para a pasta do script
        shutil.copy(arquivo_origem, arquivo_favoritos_script)
        print(f"Arquivo '{arquivo_origem}' copiado para '{arquivo_favoritos_script}' com sucesso.")
        
        # Define o caminho da pasta de destino no disco C:
        pasta_destino = os.path.join('C:\\')
        os.makedirs(pasta_destino, exist_ok=True)
        arquivo_destino = os.path.join(pasta_destino, "favoritos.html")
        
        # Copia o arquivo da pasta do script para o disco C:
        shutil.copy(arquivo_favoritos_script, arquivo_destino)
        print(f"Arquivo '{arquivo_favoritos_script}' copiado para '{arquivo_destino}' com sucesso.")
        
    except PermissionError as e:
        print(f"Erro de permissão ao copiar o arquivo: {e}. Tente executar o programa como administrador.")
    except Exception as e:
        print(f"Erro ao copiar o arquivo: {e}")

def execute_tasks(nome_agr, nome_ar, senha_pc_admin, run_bitlocker, run_install, run_users, run_gpedit, run_time, run_capture):
    if not is_admin():
        print("O script não está sendo executado como administrador. Tentando elevar privilégios...")
        run_as_admin()

    if run_install:
        install_programs()

    if run_users:
        usuario_agr = f"AGR-{nome_agr}"
        senha_agr = "AGR12345"
        senha_admin = admin_passwords.get(senha_pc_admin, "ARcm@2050")  # Senha padrão se não encontrada
        criar_usuario(usuario_agr, senha_agr, "Usuarios")
        criar_usuario("PC_Admin", senha_admin, "Administradores")
        
    copiar_favoritos(nome_ar)

    if run_gpedit:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        backup_folder = os.path.join(script_dir, 'GroupPolicyBackup')
        source_folder = r'C:\Windows\System32\GroupPolicy'

        if os.path.exists(backup_folder):
            print("Restaurando o backup de políticas de grupo...")
            if os.path.exists(source_folder):
                shutil.rmtree(source_folder)
            shutil.copytree(backup_folder, source_folder)
            print("Restauro completo.")
            subprocess.run(['gpupdate', '/force'], shell=True)
        else:
            print("Backup não encontrado. Não é possível restaurar.")

        chave = reg.HKEY_LOCAL_MACHINE
        subchave = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        alterar_registro(chave, subchave, 'DisableCAD', 1)

        secedit_config = """
        [Unicode]
        Unicode=yes
        [System Access]
        MinimumPasswordAge = 1
        MaximumPasswordAge = 30
        MinimumPasswordLength = 8
        PasswordComplexity = 1
        PasswordHistorySize = 0
        LockoutBadCount = 3
        ResetLockoutCount = 30
        LockoutDuration = 30
        [Event Audit]
        AuditSystemEvents = 3
        AuditLogonEvents = 3
        AuditObjectAccess = 3
        AuditPrivilegeUse = 3
        AuditPolicyChange = 3
        AuditAccountManage = 3
        AuditProcessTracking = 3
        AuditDSAccess = 3
        AuditAccountLogon = 3
        [Registry Values]
        MACHINE\\Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\\EnableLUA=4,0
        [Version]
        signature="$CHICAGO$"
        Revision=1
        """
        aplicar_secedit(secedit_config)

    if run_bitlocker:
        execute_powershell_script(powershell_script_path)

        time.sleep(5)

        if not os.path.exists(bitlocker_info_path):
            print(f"Arquivo de recuperação do BitLocker não encontrado: {bitlocker_info_path}")
            return

        pdf_name = f"CHAVE_BITLOCKER_{nome_agr}_AR_{nome_ar}.pdf"
        pdf_path = os.path.join('C:', pdf_name)

        convert_txt_to_pdf(bitlocker_info_path, pdf_path)

        print(f"Arquivo PDF salvo em: {pdf_path}")

    if run_time:
        executar_configuracao_data_hora()
        time_server = "tic.syngularid.com.br"
        open_control_panel()
        configure_time_server(time_server)
    
    if run_capture:
        folder_name = f"PRINTS {nome_agr} {nome_ar}"
        messagebox.showinfo("Aviso", "O script irá tirar capturas de tela. Por favor, prepare-se.")
        start_capture(
            folder_name,
            True,  # Capturar MAC Address
            True,  # Capturar Informações do Sistema
            True,  # Capturar Serial da Máquina
            True,  # Capturar Programas Instalados
            True   # Capturar Ativação do Windows
        )
    
    # Exibir mensagem de conclusão
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Concluído", "Todas as tarefas foram concluídas com sucesso!")
    root.destroy()
    
def main():
    def on_submit():
        nome_agr = nome_agr_entry.get()
        nome_ar = nome_ar_entry.get()
        senha_pc_admin = senha_pc_admin_var.get()
        run_bitlocker = bitlocker_var.get()
        run_install = install_var.get()
        run_users = users_var.get()
        run_gpedit = gpedit_var.get()
        run_time = time_var.get()
        run_capture = capture_var.get()
        window.destroy()
        execute_tasks(nome_agr, nome_ar, senha_pc_admin, run_bitlocker, run_install, run_users, run_gpedit, run_time, run_capture)

    current_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(current_dir, 'gear_icon.ico')

    window = tk.Tk()
    window.title("Configuração")
    window.geometry("490x330")  # Define o tamanho da janela
    
    window.iconbitmap(icon_path)

    # Labels
    tk.Label(window, text="Nome do AGR:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    tk.Label(window, text="Nome da AR:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    tk.Label(window, text="Senha Administrador:").grid(row=2, column=0, padx=10, pady=10, sticky="e")

    # Entries
    nome_agr_entry = tk.Entry(window)
    nome_agr_entry.grid(row=0, column=1, padx=10, pady=10)

    nome_ar_entry = tk.Entry(window)
    nome_ar_entry.grid(row=1, column=1, padx=10, pady=10)

    # Radiobuttons
    senha_pc_admin_var = tk.StringVar(value="ARcm@2050")
    tk.Radiobutton(window, text="AR Própria", variable=senha_pc_admin_var, value="ARcm@2050").grid(row=2, column=1, padx=10, pady=5, sticky="w")
    tk.Radiobutton(window, text="Outras", variable=senha_pc_admin_var, value="ARcert1127").grid(row=2, column=1, padx=10, pady=5, sticky="e")

    # Checkbuttons
    bitlocker_var = tk.BooleanVar(value=True)
    install_var = tk.BooleanVar(value=True)
    users_var = tk.BooleanVar(value=True)
    gpedit_var = tk.BooleanVar(value=True)
    time_var = tk.BooleanVar(value=True)
    capture_var = tk.BooleanVar(value=True)

    tk.Checkbutton(window, text="BitLocker", variable=bitlocker_var).grid(row=3, column=0, padx=10, pady=5, sticky="w")
    tk.Checkbutton(window, text="Instalação de Drivers", variable=install_var).grid(row=3, column=1, padx=10, pady=5, sticky="w")
    tk.Checkbutton(window, text="Usuários", variable=users_var).grid(row=3, column=2, padx=10, pady=5, sticky="w")
    tk.Checkbutton(window, text="Gpedit", variable=gpedit_var).grid(row=4, column=0, padx=10, pady=5, sticky="w")
    tk.Checkbutton(window, text="Configurar Data e Hora", variable=time_var).grid(row=4, column=1, padx=10, pady=5, sticky="w")
    tk.Checkbutton(window, text="Capturas de Tela", variable=capture_var).grid(row=5, column=0, padx=10, pady=5, sticky="w")

    # Button
    execute_button = tk.Button(window, text="Executar", command=on_submit)
    execute_button.grid(row=6, column=0, columnspan=3, padx=10, pady=20)

    window.mainloop()

if __name__ == "__main__":
    main()
