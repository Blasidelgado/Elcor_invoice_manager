import dotenv
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import pdfplumber
from openpyxl import load_workbook
import veryfi

from console import Console
from Veryfi import Veryfi
from helpers import parse_afip, update_worksheet, manipulate_invoice

class ElcorInvoiceManager:
    def __init__(self, root):
        self.root = root
        self.selected_directory = None
        self.selected_xlsx = None
        self.setup_ui()
        self.verify = Veryfi(self)
        dotenv.load_dotenv() # Load .env keys
        # Call setup verify to set up keys for further api requests
        self.client = self.verify.setup_veryfi(dotenv.get_key('.env', 'CLIENT_ID'), 
            dotenv.get_key('.env', 'CLIENT_SECRET'), dotenv.get_key('.env', 'USERNAME'), 
            dotenv.get_key('.env', 'API_KEY'))
        
    def setup_ui(self):
        # Create a grid with two columns and two rows
        self.root.columnconfigure(0, weight=3)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        self.root.columnconfigure(3, weight=1)
        self.root.columnconfigure(4, weight=2)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self.root.rowconfigure(2, weight=1)
        self.root.rowconfigure(3, weight=5)

        # Add widgets to the grid
        tk.Label(self.root, text="Please select a work directory", font=(16)).grid(row=0, column=0, columnspan=4, sticky="nsew")
        tk.Button(self.root, text="Select directory", font=(16), command=self.select_directory).grid(row=0, column=4)
        tk.Label(self.root, text="Please select a xlsx file to insert info", font=(16)).grid(row=1, column=0, columnspan=4, sticky="nsew")
        tk.Button(self.root, text="Select xlsx", font=(16), command=self.select_xlsx).grid(row=1, column=4)
        tk.Button(self.root, text="START", font=(16), command=self.start_processing).grid(row=2, column=1, columnspan=2, sticky="nsew")

        self.console = Console(self.root)
        self.console.grid(row=3, column=0, columnspan=5)

    def select_directory(self):
        self.selected_directory = filedialog.askdirectory()
        self.console.write(f'Selected directory: {self.selected_directory}\n')

    def select_xlsx(self):
        self.selected_xlsx = filedialog.askopenfilename()
        self.console.write(f'Selected xlsl file: {self.selected_xlsx}\n')

    def start_processing(self):
        if self.selected_directory is None:
            self.console.write("Please select a directory\n")
            return

        if self.selected_xlsx is None:
            self.console.write("Please select an xlsx file\n")
            return
        
        # Instancing a path class to the selected directory
        p = Path(self.selected_directory)
        if not p.is_dir():
            self.console.write("Please select a valid folder\n")
            return

        # Instancing a path class to the xlsx file
        xlsx = Path(self.selected_xlsx)
        if not xlsx.suffix == '.xlsx':
            self.console.write("Please select a valid xlsx file\n")
            return

        # List all available files to scan
        extensions = ['pdf', 'jpg', 'jpeg', 'png']
        invoices = [file_path for ext in extensions for file_path in p.glob(f'*.{ext}') if file_path.is_file()]

        # Load the workbook and select worksheet
        wb = load_workbook(filename = xlsx)
        ws = wb.active

        data_list = []

        # Start scanning files
        for invoice in invoices:
            filename = invoice.name
            self.console.write(f'Processing file: {filename}\n')
            suffix = invoice.suffix

            if suffix == '.pdf':

                # Read pdf with pdfplumber
                reader = pdfplumber.open(invoice)

                # Take only the first pdf page
                page = reader.pages[0]

                # Read document metadata for further processing
                if reader.metadata.get('Creator') == 'AFIP': # Administración Federal de Ingresos Públicos   
                    data = parse_afip(page)

                    # Append data to data_list
                    data_list.append(data)

                    # Close file
                    reader.close()


                else: # PDF file is sent to our trusted API
                    info = self.verify.parse_veryfi(invoice, self.client)

                    if not self.verify.check_response(filename, info):
                        continue

                    # Append data to data_list
                    data_list.append(data)

                    # Close file
                    reader.close()

                    
            else: # File is an image that will be sent to our trusted API
                data = self.verify.parse_veryfi(invoice, self.client)
                
                if not self.verify.check_response(filename, data):
                    continue
                
                # Append data to data_list
                data_list.append(data)

                # Close file
                reader.close()


            #TODO: Take data_list and sort it by date(date format should be dd/mm/Y)

            # Plug data in xlsl selected file

            update_worksheet(ws, data)

            # Save worksheet    
            wb.save(xlsx)
            self.console.write(f'Data appended to xlsx succesfully.\n')

            # Manipulate file in respective inner directory
            self.console.write(manipulate_invoice(p, invoice, suffix, 'invoices', data))
            