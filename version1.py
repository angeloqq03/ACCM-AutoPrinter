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
from matplotlib.figure import Figure
from google.oauth2 import service_account
from googleapiclient.discovery import build
from Mappings import SourceCode, FirstName, MiddleName, LastName

class PrintControlWindow:
    def __init__(self, master, print_callback):
        self.window = Toplevel(master)
        self.window.title("Print Control")
        self.window.geometry("200x100")
        self.window.resizable(False, False)
        print_btn = ttk.Button(self.window, text="Print (344 DPI)", 
                             command=print_callback)
        print_btn.pack(expand=True, pady=20)

class LegalFormApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Legal Form Preview")
        self.paper_width = 8.5
        self.paper_height = 13.0
        self.printer_dpi = 344

        self.SourceCode = SourceCode
        self.FirstName = FirstName
        self.MiddleName = MiddleName
        self.LastName = LastName

        self.GoogleSheetID = "1BPRsZfaz7CZk9J59FVxb6N7EWBkgYtliSsVvXSt8eQk"
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.SERVICE_ACCOUNT_FILE = 'organic.json'

        
        self.plotFirstName = "OSCAR"
        self.plotMiddleName = "COLLINS"
        self.plotLastName = "RIVERA"

        self.fetch_sheet_data(row=3)
        self.load_image()
        self.create_ui()
        self.print_window = PrintControlWindow(root, self.Output)

    def fetch_sheet_data(self, row):
        try:
            print("\n=== Starting Google Sheets Fetch ===")
            print(f"Attempting to fetch row {row}")
            
            if not os.path.exists(self.SERVICE_ACCOUNT_FILE):
                print(f"ERROR: Credentials file not found at: {self.SERVICE_ACCOUNT_FILE}")
                return
            print(f"Found credentials file: {self.SERVICE_ACCOUNT_FILE}")
            
            creds = service_account.Credentials.from_service_account_file(
                self.SERVICE_ACCOUNT_FILE, scopes=self.SCOPES)
            print("Credentials loaded successfully")
            
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            print("Service built successfully")
            
            sheet_name = 'DETAILS FOR APP FORM'
            range_name = f"'{sheet_name}'!C{row}:F{row}"
            print(f"Fetching range: {range_name}")
            
            result = sheet.values().get(
                spreadsheetId=self.GoogleSheetID,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            print(f"Raw fetch result: {values}")
            
            if values and len(values[0]) >= 4:
                self.plotFirstName = values[0][1]    # Column D
                self.plotMiddleName = values[0][2]   # Column E
                self.plotLastName = values[0][3]     # Column F
                print(f"succ - Names loaded: {self.plotFirstName} {self.plotMiddleName} {self.plotLastName}")
            else:
                print("warn: No data found in specified range")
                
        except Exception as e:
            print(f"err fetching: {str(e)}")
            messagebox.showwarning("Warning", "default values - Sheet data fetch failed")

    def load_image(self):
        self.image_path = "front_page.png"
        self.image = cv2.imread(self.image_path)
        if self.image is None:
            raise ValueError(f"failed to load image: {self.image_path}")
        self.image = cv2.cvtColor(self.image, cv2.COLOR_BGR2RGB)

    def create_ui(self):
        preview_frame = ttk.Frame(self.root)
        preview_frame.pack(fill=BOTH, expand=True)
        
        preview_width = 850
        preview_height = int(preview_width * (self.paper_height / self.paper_width))
        self.canvas = Canvas(preview_frame, width=preview_width, height=preview_height)
        
        y_scroll = ttk.Scrollbar(preview_frame, orient=VERTICAL, 
                               command=self.canvas.yview)
        x_scroll = ttk.Scrollbar(preview_frame, orient=HORIZONTAL, 
                               command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=y_scroll.set, 
                            xscrollcommand=x_scroll.set)
        
        y_scroll.pack(side=RIGHT, fill=Y)
        x_scroll.pack(side=BOTTOM, fill=X)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        self.showPrev()

        # TODO: `
        #  Will work on this later, gonna fix the sheets first 
        # `
    def showPrev(self):
        fig = Figure(figsize=(self.paper_width, self.paper_height), dpi=344)
        ax = fig.add_subplot(111)
        
        ax.imshow(self.image)
        
        coords1 = list(self.SourceCode.values())
        ax.scatter(*zip(*coords1), color="red", s=30, marker="o")
        
        font_size = 26 * (self.paper_height / 14.0)
        
        # all names
        self.plotText(ax, self.plotFirstName, self.FirstName.values(), font_size)
        self.plotText(ax, self.plotMiddleName, self.MiddleName.values(), font_size)
        self.plotText(ax, self.plotLastName, self.LastName.values(), font_size)
        
        ax.axis('off')
        
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=344, 
                   bbox_inches='tight', pad_inches=0)
        buf.seek(0)
        
        img = Image.open(buf)
        preview_width = 850
        preview_height = int(preview_width * (self.paper_height / self.paper_width))
        img = img.resize((preview_width, preview_height), Image.Resampling.LANCZOS)
        self.photo = ImageTk.PhotoImage(img)
        
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def plotText(self, ax, text, coordinates, font_size):
        for i, char in enumerate(text):
            if i < len(coordinates):
                x, y = list(coordinates)[i]
                ax.text(x, y, char,
                       fontsize=font_size, color="black",
                       ha='center', va='center',
                       fontweight='bold', fontname='Arial')

    def Output(self):
        try:
            printer_name = win32print.GetDefaultPrinter()
            dc = win32ui.CreateDC()
            dc.CreatePrinterDC(printer_name)
            
            dc.StartDoc('Legal Form')
            dc.StartPage()
            
            fig = Figure(figsize=(self.paper_width, self.paper_height), dpi=344)
            ax = fig.add_subplot(111)
            
            ax.imshow(self.image)
            coords1 = list(self.SourceCode.values())
            ax.scatter(*zip(*coords1), color="red", s=30, marker="o")
            
            font_size = 8 * (self.paper_height / 14.0)
            
            # Print all names
            self.plotText(ax, self.plotFirstName, self.FirstName.values(), font_size)
            self.plotText(ax, self.plotMiddleName, self.MiddleName.values(), font_size)
            self.plotText(ax, self.plotLastName, self.LastName.values(), font_size)
            
            ax.axis('off')
            
            buf = io.BytesIO()
            fig.savefig(buf, format='png', dpi=344, 
                       bbox_inches='tight', pad_inches=0)
            buf.seek(0)
            
            img = Image.open(buf)
            dib = ImageWin.Dib(img)
            width = int(self.paper_width * self.printer_dpi)
            height = int(self.paper_height * self.printer_dpi)
            dib.draw(dc.GetHandleOutput(), (0, 0, width, height))
            
            dc.EndPage()
            dc.EndDoc()
            dc.DeleteDC()
            
            messagebox.showinfo("Success", "Printed at 344 DPI")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # NOTE: `
    # I need to anaylze wether batch printing is possible or not
    # `

if __name__ == "__main__":
    root = Tk()
    app = LegalFormApp(root)
    root.mainloop()