import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from gui.main_window import AnalyticalDataProcessor

def main():
    # Initialize the ttkbootstrap theme
    root = ttk.Window(themename="darkly")  # Use a dark theme like "darkly"
    app = AnalyticalDataProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
