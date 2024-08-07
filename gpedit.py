import winreg as reg
import subprocess

def alterar_registro(chave, subchave, valor_nome, valor, tipo=reg.REG_DWORD):
    try:
        chave_reg = reg.OpenKey(chave, subchave, 0, reg.KEY_SET_VALUE)
    except FileNotFoundError:
        chave_reg = reg.CreateKey(chave, subchave)
    try:
        reg.SetValueEx(chave_reg, valor_nome, 0, tipo, valor)
        reg.CloseKey(chave_reg)
        print(f'{valor_nome} alterado para {valor} em {subchave}.')
    except WindowsError as e:
        print(f'Erro ao alterar o registro: {e}')

def aplicar_secedit(config):
    try:
        with open('secedit.inf', 'w') as file:
            file.write(config)
        subprocess.run(['secedit', '/configure', '/db', 'secedit.sdb', '/cfg', 'secedit.inf'], check=True)
        print('Configurações de segurança aplicadas com sucesso.')
    except subprocess.CalledProcessError as e:
        print(f'Erro ao aplicar configurações de segurança: {e}')

def aplicar_alteracoes_gpedit():
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
