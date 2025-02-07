from tkinter import *
from tkinter import ttk, messagebox
import win32print
class PrintControlWindow:
    def __init__(self, master, fetch_callback, print_callback):
        self.window = Toplevel(master)
        self.window.title("Print Control Dashboard")
        self.window.geometry("600x400")
        self.window.resizable(True, True)
        search_frame = ttk.LabelFrame(self.window, text="Search Client", padding=10)
        search_frame.pack(fill=X, padx=5, pady=5)
        
        self.client_entry = ttk.Entry(search_frame, width=20)
        self.client_entry.pack(side=LEFT, padx=5)
        
        search_btn = ttk.Button(search_frame, text="Search", 
                              command=lambda: self.search_client())
        search_btn.pack(side=LEFT, padx=5)
        
        data_frame = ttk.LabelFrame(self.window, text="Client Information", padding=10)
        data_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(data_frame, columns=("Value",), show="tree")
        self.tree.pack(fill=BOTH, expand=True)
        
        selected_frame = ttk.LabelFrame(self.window, text="Selected Clients", padding=10)
        selected_frame.pack(fill=X, padx=5, pady=5)
        
        self.selected_listbox = Listbox(selected_frame, height=3)
        self.selected_listbox.pack(side=LEFT, fill=X, expand=True)
        btn_frame = ttk.Frame(selected_frame)
        btn_frame.pack(side=RIGHT, padx=5)
        add_btn = ttk.Button(btn_frame, text="Add Current", 
                           command=self.add_to_selection)
        add_btn.pack(pady=2)
        remove_btn = ttk.Button(btn_frame, text="Remove Selected", 
                              command=self.remove_from_selection)
        remove_btn.pack(pady=2)
        print_frame = ttk.Frame(self.window)
        print_frame.pack(fill=X, padx=5, pady=5)
        printer_label = ttk.Label(print_frame, text="Printer:")
        printer_label.pack(side=LEFT, padx=5)
        self.printer_combo = ttk.Combobox(print_frame, 
                                        values=self.get_printers(),
                                        state="readonly")
        self.printer_combo.pack(side=LEFT, expand=True, fill=X, padx=5)
        self.printer_combo.set(win32print.GetDefaultPrinter())
        print_btn = ttk.Button(print_frame, text="Print Selected",
                             command=lambda: print_callback(self.get_selected_clients()))
        print_btn.pack(side=RIGHT, padx=5)
        self.fetch_callback = fetch_callback
        self.print_callback = print_callback
        self.current_client = None
        self.selected_clients = set()

    def get_printers(self):
        """Get list of available printers"""
        return [printer[2] for printer in win32print.EnumPrinters(2)]

    def search_client(self):
        """Search for client and display data"""
        client_no = self.client_entry.get().strip().upper()
        if not client_no:
            messagebox.showwarning("Warning", "Please enter a client number")
            return
            
        try:
            client_data = self.fetch_callback(client_no)
            if client_data:
                self.current_client = client_no
                self.display_client_data(client_data)
            else:
                messagebox.showinfo("Not Found", f"No data found for client {client_no}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch client data: {str(e)}")

    def display_client_data(self, client_data):
        """Display client data in treeview"""
        self.tree.delete(*self.tree.get_children())
        for key, value in client_data.items():
            if isinstance(value, list):
                value = ''.join(value)
            self.tree.insert("", END, text=key, values=(value,))

    def add_to_selection(self):
        """Add current client to selection"""
        if self.current_client:
            self.selected_clients.add(self.current_client)
            self.update_selection_display()

    def remove_from_selection(self):
        """Remove selected client from selection"""
        selection = self.selected_listbox.curselection()
        if selection:
            idx = selection[0]
            client = self.selected_listbox.get(idx)
            self.selected_clients.remove(client)
            self.update_selection_display()

    def update_selection_display(self):
        """Update listbox with selected clients"""
        self.selected_listbox.delete(0, END)
        for client in sorted(self.selected_clients):
            self.selected_listbox.insert(END, client)

    def get_selected_clients(self):
        """Return list of selected client numbers"""
        return list(self.selected_clients)