from tkinter import *
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageWin
import cv2
import numpy as np
import io
import os
import win32print
import win32ui
import win32con
from google.oauth2 import service_account
from googleapiclient.discovery import build
from Mappings import (
SourceCode, FirstName, MiddleName, LastName, MmDdYyyy,
PlaceOfBirth, MotherFirstName, MotherMiddleName, MotherLastName,GenderMale,GenderFemale,
CivilStatusSingle,CivilStatusWidowed,CivilStatusMarried,CivilStatusSeperated,    
)
from Plottings import create_print_figure
from PrintControlWindow import PrintControlWindow

class LegalFormApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Legal Form Manager")
        self.paper_width = 8.5
        self.paper_height = 13.0
        self.printer_dpi = 344

        self.mappings = {
            'SourceCode': SourceCode,
            'FirstName': FirstName,
            'MiddleName': MiddleName,
            'LastName': LastName,
            'MmDdYyyy': MmDdYyyy,
            'PlaceOfBirth': PlaceOfBirth,
            'MotherFirstName': MotherFirstName,
            'MotherMiddleName': MotherMiddleName,
            'MotherLastName': MotherLastName,
            'GenderMale': GenderMale,
            'GenderFemale': GenderFemale,
            'CivilStatusSingle': CivilStatusSingle,
            'CivilStatusWidowed': CivilStatusWidowed,
            'CivilStatusMarried': CivilStatusMarried,
            'CivilStatusSeperated': CivilStatusSeperated
        }

        self.paper_dims = {
            'width': self.paper_width,
            'height': self.paper_height
        }

        self.GoogleSheetID = "1BPRsZfaz7CZk9J59FVxb6N7EWBkgYtliSsVvXSt8eQk"
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.SERVICE_ACCOUNT_FILE = 'organic.json'
        self.client_data = {}

        try:
            self.load_image()
            self.print_window = PrintControlWindow(
                root,
                fetch_callback=self.fetch_single_client,
                print_callback=self.print_selected_clients
            )
        except Exception as e:
            print(f"Initialization error: {str(e)}")
            messagebox.showerror("Error", f"Initialization failed: {str(e)}")
            raise

    def load_image(self):
        try:
            self.image_path = "front_page.png"
            self.image = cv2.imread(self.image_path)
            if self.image is None:
                raise ValueError(f"Failed to load image: {self.image_path}")
            self.image = cv2.cvtColor(self.image, cv2.COLOR_BGR2RGB)
        except Exception as e:
            print(f"Error loading image: {str(e)}")
            raise

    def fetch_single_client(self, client_no):
        """Fetch data for a single client"""
        try:
            if not os.path.exists(self.SERVICE_ACCOUNT_FILE):
                raise FileNotFoundError(f"Credentials file not found: {self.SERVICE_ACCOUNT_FILE}")
                
            creds = service_account.Credentials.from_service_account_file(
                self.SERVICE_ACCOUNT_FILE, scopes=self.SCOPES)
            
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            
            result = sheet.values().get(
                spreadsheetId=self.GoogleSheetID,
                range="'DETAILS FOR APP FORM'!A:ZZ"
            ).execute()
            
            values = result.get('values', [])
            if not values:
                return None
            
            headers = values[0]
            
            for row in values[1:]:
                if row and row[0].strip().upper() == client_no.strip().upper():
                    date_str = row[6] if len(row) > 6 else ''
                    date_parts = date_str.split('/') if date_str else []
                    formatted_date = ''.join(date_parts) if date_parts else ''
                    
                    client_data = {
                        'source_code': list(row[2].strip()) if len(row) > 2 else [],
                        'first_name': row[3] if len(row) > 3 else '',
                        'middle_name': row[4] if len(row) > 4 else '',
                        'last_name': row[5] if len(row) > 5 else '',
                        'date_of_birth': formatted_date,
                        'place_of_birth': row[7] if len(row) > 7 else '',
                        'mother_first_name': row[8] if len(row) > 8 else '',
                        'mother_middle_name': row[9] if len(row) > 9 else '',
                        'mother_last_name': row[10] if len(row) > 10 else '',
                        'gender': row[11] if len(row) > 11 else '',
                        'civil_status': row[12] if len(row) > 12 else ''
                    }
                    
                    self.client_data[client_no] = client_data
                    return client_data
            
            return None
            
        except Exception as e:
            print(f"Fetch error: {str(e)}")
            raise

    def print_selected_clients(self, client_numbers):
        """Print forms for selected clients"""
        if not client_numbers:
            messagebox.showwarning("Warning", "No clients selected")
            return
            
        try:
            printer_name = win32print.GetDefaultPrinter()
            
            for client_no in client_numbers:
                if client_no not in self.client_data:
                    if not self.fetch_single_client(client_no):
                        continue
                
                client = self.client_data[client_no]
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)
                
                dc.StartDoc(f'Legal Form - Client {client_no}')
                dc.StartPage()
                
                fig = create_print_figure(
                    self.image,
                    client,
                    self.mappings,
                    self.paper_dims
                )
                
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=344, bbox_inches='tight', pad_inches=0)
                buf.seek(0)
                
                img = Image.open(buf)
                dib = ImageWin.Dib(img)
                
                width = int(self.paper_width * self.printer_dpi)
                height = int(self.paper_height * self.printer_dpi)
                dib.draw(dc.GetHandleOutput(), (0, 0, width, height))
                
                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()
            
            messagebox.showinfo("Success", f"Printed {len(client_numbers)} forms at 344 DPI")
            
        except Exception as e:
            print(f"Print error: {str(e)}")
            messagebox.showerror("Error", str(e))