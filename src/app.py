import tkinter as tk
from ui import MT5UI
# Add this to the top of app.py to ensure Excel template paths work in PyInstaller
import os, sys
# Fix for PyInstaller bundled app
if getattr(sys, 'frozen', False):
    # When running as bundled executable
    base_path = sys._MEIPASS
    os.environ['OPENPYXL_TEMPLATE_PATH'] = os.path.join(base_path, 'openpyxl', 'templates')
else:
    # When running as normal Python script
    base_path = os.path.dirname(__file__)
    
if __name__ == "__main__":
    root = tk.Tk()
    app = MT5UI(root)
    root.mainloop()
# pyinstaller --onefile --windowed --icon=assets/icon.ico --name "mt5reporter" --collect-data openpyxl --collect-data sv_ttk --hidden-import pandas --hidden-import pytz --hidden-import win32timezone --hidden-import MetaTrader5 --add-data "assets/*;assets/" src/app.py

# C:/Users/Alvin Tegutwo/AppData/Local/Programs/Python/Python313