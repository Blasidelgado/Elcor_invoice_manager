import tkinter as tk
from elcor import ElcorInvoiceManager

if __name__ == '__main__':
    root = tk.Tk()
    root.title("Elcor Invoice Manager")
    app = ElcorInvoiceManager(root)
    root.mainloop()
