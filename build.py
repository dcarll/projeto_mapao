import PyInstaller.__main__
import os

def build_exe():
    print("Iniciando build...")
    
    PyInstaller.__main__.run([
        'main.py',
        '--name=ScheduleLabs',
        '--onefile',
        '--noconsole',
        '--clean',
        '--hidden-import=PIL',
        '--hidden-import=openpyxl',
        # '--add-data=README.md;.' # Exemplo se precisar adicionar arquivos
    ])
    
    print("Build concluído! Executável em dist/ScheduleLabs.exe")

if __name__ == "__main__":
    build_exe()
