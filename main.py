"""
Sistema de Agendamento de Laboratórios
Arquivo principal
"""

import os
import tkinter as tk
from gui import ScheduleGUI
"""
def setup_environment():
    #Garante que a pasta de dados do sistema exista.
    path = "C:\\SCHEDULE_LABS"
    if not os.path.exists(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
"""
def main():
    #setup_environment()
    root = tk.Tk()
    app = ScheduleGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
