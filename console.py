import tkinter as tk
import tkinter.scrolledtext as tkst

class Console(tkst.ScrolledText):
    def write(self, txt):
        self.insert(tk.END, txt)
        self.see(tk.END)
        self.update_idletasks()

    def flush(self):
        self.delete('1.0', tk.END)
        self.update_idletasks()