# Script de inicialização para o Alfredo do Email
# Este script cria o ambiente virtual, ativa, instala dependências e executa o Streamlit

import os
import subprocess
import sys

# Passo 1: Criar ambiente virtual
if not os.path.exists('venv'):
    subprocess.run([sys.executable, '-m', 'venv', 'venv'], check=True)

# Passo 2: Ativar ambiente virtual (apenas para Windows)
activate_script = os.path.join('venv', 'Scripts', 'activate')

# Passo 3: Instalar dependências
subprocess.run([os.path.join('venv', 'Scripts', 'python.exe'), '-m', 'pip', 'install', '--upgrade', 'pip'], check=True)
subprocess.run([os.path.join('venv', 'Scripts', 'python.exe'), '-m', 'pip', 'install', 'streamlit', 'pandas', 'pywin32', 'streamlit-quill'], check=True)

# Passo 4: Rodar o Streamlit
def run_streamlit():
    subprocess.run([os.path.join('venv', 'Scripts', 'python.exe'), '-m', 'streamlit', 'run', 'gerador_email.py'])

if __name__ == '__main__':
    run_streamlit()
