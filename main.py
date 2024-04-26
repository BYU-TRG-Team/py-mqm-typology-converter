from mainwindow import MainWindow
import tkinter as tk

if __name__ == "__main__":
    """
    This script creates a Tkinter application window and runs the main window of the MQM Typology Converter.
    """
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
