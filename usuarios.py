import subprocess

def criar_usuario(usuario, senha, grupo):
    try:
        # Criar usuário com senha que não expira
        subprocess.run(['net', 'user', usuario, senha, '/add'], check=True)
        # Adicionar o usuário ao grupo especificado
        subprocess.run(['net', 'localgroup', grupo, usuario, '/add'], check=True)
        print(f'Usuário {usuario} criado com sucesso com permissões de {grupo} e senha sem expiração.')
    except subprocess.CalledProcessError as e:
        print(f'Erro ao criar usuário {usuario}: {e}')
