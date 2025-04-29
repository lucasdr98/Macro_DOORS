import os
import sys
import gui

# Função para garantir que o PyInstaller inclua o macro.py no executável,
# mas sem carregá-lo no início para evitar problemas de UI
def include_modules():
    # Este import é apenas para o PyInstaller detectar as dependências
    # Não é usado diretamente, então não afeta o funcionamento da aplicação
    import macro
    # Outros módulos que possam ser necessários adicionar aqui
    return None

if __name__ == "__main__":
    # Configurar diretório de trabalho para garantir que os recursos sejam encontrados
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
    
    # Iniciar a interface gráfica
    gui.main() 