from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
import win32print
import io
from Plottings import create_print_figure

class PrintControlWindow:
    def __init__(self, master, image, mappings, paper_dims, fetch_callback, print_callback, sheet_id_callback):
        self.window = Toplevel(master)
        self.window.title("Print Control Dashboard")
        self.window.state('zoomed')

        # Store parameters
        self.image = image
        self.mappings = mappings
        self.paper_dims = paper_dims
        self.fetch_callback = fetch_callback
        self.print_callback = print_callback
        self.sheet_id_callback = sheet_id_callback
        
        # Initialize state
        self.selected_clients = set()
        self.print_history = []
        self.current_client = None
        self.preview_images = []
        self.current_page = 0

        self.setup_ui()

    def setup_ui(self):
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill=BOTH, expand=True)

        # Create frames for each tab
        self.print_frame = Frame(self.notebook)
        self.history_frame = Frame(self.notebook)
        self.preview_frame = Frame(self.notebook)
        self.settings_frame = Frame(self.notebook)
        self.batch_frame = Frame(self.notebook)

        # Add tabs to notebook
        self.notebook.add(self.print_frame, text="Print")
        self.notebook.add(self.history_frame, text="History")
        self.notebook.add(self.preview_frame, text="Preview")
        self.notebook.add(self.settings_frame, text="Settings")
        self.notebook.add(self.batch_frame, text="Batch Print")

        # Setup individual tabs
        self.setup_print_tab()
        self.setup_history_tab()
        self.setup_preview_tab()
        self.setup_settings_tab()
        self.setup_batch_tab()

    def setup_print_tab(self):
        # Search section
        search_frame = ttk.LabelFrame(self.print_frame, text="Search Client", padding=10)
        search_frame.pack(fill=X, padx=5, pady=5)

        self.client_entry = ttk.Entry(search_frame, width=20)
        self.client_entry.pack(side=LEFT, padx=5)

        search_btn = ttk.Button(search_frame, text="Search", command=self.search_client)
        search_btn.pack(side=LEFT, padx=5)

        # Client info section
        data_frame = ttk.LabelFrame(self.print_frame, text="Client Information", padding=10)
        data_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(data_frame, columns=("Value",), show="tree")
        self.tree.pack(fill=BOTH, expand=True)

        # Selection section
        selected_frame = ttk.LabelFrame(self.print_frame, text="Selected Clients", padding=10)
        selected_frame.pack(fill=X, padx=5, pady=5)

        self.selected_listbox = Listbox(selected_frame, height=3)
        self.selected_listbox.pack(side=LEFT, fill=X, expand=True)

        btn_frame = ttk.Frame(selected_frame)
        btn_frame.pack(side=RIGHT, padx=5)

        add_btn = ttk.Button(btn_frame, text="Add Current", command=self.add_to_selection)
        add_btn.pack(pady=2)

        remove_btn = ttk.Button(btn_frame, text="Remove Selected", command=self.remove_from_selection)
        remove_btn.pack(pady=2)

        # Print controls
        print_frame = ttk.Frame(self.print_frame)
        print_frame.pack(fill=X, padx=5, pady=5)

        printer_label = ttk.Label(print_frame, text="Printer:")
        printer_label.pack(side=LEFT, padx=5)

        self.printer_combo = ttk.Combobox(print_frame, values=self.get_printers(), state="readonly")
        self.printer_combo.pack(side=LEFT, expand=True, fill=X, padx=5)
        self.printer_combo.set(win32print.GetDefaultPrinter())

        print_btn = ttk.Button(print_frame, text="Print Selected", command=self.print_selected)
        print_btn.pack(side=RIGHT, padx=5)

        # Message display
        self.message_label = Label(self.print_frame, text="", fg="black")
        self.message_label.pack(pady=10)

    def setup_history_tab(self):
        self.history_listbox = Listbox(self.history_frame)
        self.history_listbox.pack(fill=BOTH, expand=True, pady=10)

    def setup_preview_tab(self):
        self.preview_label = Label(self.preview_frame, text="Print Preview")
        self.preview_label.pack(pady=10)

        # Preview canvas
        self.preview_canvas = Canvas(self.preview_frame, width=400, height=600)
        self.preview_canvas.pack(pady=10)

        # Page navigation
        self.page_control_frame = Frame(self.preview_frame)
        self.page_control_frame.pack(pady=10)

        self.prev_page_btn = Button(self.page_control_frame, text="Previous Page", command=self.show_prev_page)
        self.prev_page_btn.pack(side=LEFT, padx=5)

        self.page_label = Label(self.page_control_frame, text="Page 1")
        self.page_label.pack(side=LEFT, padx=5)

        self.next_page_btn = Button(self.page_control_frame, text="Next Page", command=self.show_next_page)
        self.next_page_btn.pack(side=LEFT, padx=5)

    def setup_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.settings_frame, text="Google Sheet Settings", padding=10)
        settings_frame.pack(fill=X, padx=5, pady=5)

        sheet_id_label = ttk.Label(settings_frame, text="Google Sheet ID:")
        sheet_id_label.pack(side=LEFT, padx=5)

        self.sheet_id_entry = ttk.Entry(settings_frame, width=40)
        self.sheet_id_entry.pack(side=LEFT, padx=5)

        save_btn = ttk.Button(settings_frame, text="Save", command=self.save_sheet_id)
        save_btn.pack(side=LEFT, padx=5)

        self.load_sheet_id()

    def setup_batch_tab(self):
        """Setup batch printing tab"""
        main_frame = ttk.LabelFrame(self.batch_frame, text="Batch Print Controls", padding=10)
        main_frame.pack(fill=X, padx=5, pady=5)

        # Start client number
        start_frame = ttk.Frame(main_frame)
        start_frame.pack(fill=X, pady=5)
        
        ttk.Label(start_frame, text="Start Client No:").pack(side=LEFT, padx=5)
        self.batch_start_entry = ttk.Entry(start_frame, width=20)
        self.batch_start_entry.pack(side=LEFT, padx=5)

        # End client number
        end_frame = ttk.Frame(main_frame)
        end_frame.pack(fill=X, pady=5)
        
        ttk.Label(end_frame, text="End Client No:").pack(side=LEFT, padx=5)
        self.batch_end_entry = ttk.Entry(end_frame, width=20)
        self.batch_end_entry.pack(side=LEFT, padx=5)

        # Progress display
        self.batch_progress_var = StringVar()
        ttk.Label(main_frame, textvariable=self.batch_progress_var).pack(fill=X, pady=5)

        # Print button
        ttk.Button(main_frame, text="Print Batch", command=self.print_batch).pack(pady=10)

        # Message display
        self.batch_message_label = Label(main_frame, text="", fg="black")
        self.batch_message_label.pack(pady=10)

    def load_sheet_id(self):
        """Load the Google Sheet ID from sheet_id.txt"""
        try:
            with open('sheet_id.txt', 'r') as f:
                sheet_id = f.read().strip()
                self.sheet_id_entry.delete(0, END)
                self.sheet_id_entry.insert(0, sheet_id)
        except FileNotFoundError:
            self.display_message("sheet_id.txt not found. Please enter a Google Sheet ID.", "warning")

    def save_sheet_id(self):
        """Save the Google Sheet ID to sheet_id.txt"""
        sheet_id = self.sheet_id_entry.get().strip()
        if sheet_id:
            try:
                self.sheet_id_callback(sheet_id)
                self.display_message("Google Sheet ID saved successfully.", "success")
            except Exception as e:
                self.display_message(f"Failed to save Sheet ID: {str(e)}", "error")
        else:
            self.display_message("Please enter a valid Google Sheet ID.", "error")

    def get_printers(self):
        """Get list of available printers"""
        return [printer[2] for printer in win32print.EnumPrinters(2)]

    def search_client(self):
        """Search for client and display data"""
        client_no = self.client_entry.get().strip().upper()
        if not client_no:
            self.display_message("Please enter a client number", "warning")
            return

        try:
            client_data = self.fetch_callback(client_no)
            if client_data:
                self.current_client = client_no
                self.display_client_data(client_data)
            else:
                self.display_message(f"No data found for client {client_no}", "info")
        except Exception as e:
            self.display_message(f"Failed to fetch client data: {str(e)}", "error")

    def display_client_data(self, client_data):
        """Display client data in treeview"""
        self.tree.delete(*self.tree.get_children())
        for key, value in client_data.items():
            if isinstance(value, list):
                value = ''.join(value)
            self.tree.insert("", END, text=key, values=(value,))

    def add_to_selection(self):
        """Add current client to selection"""
        if hasattr(self, 'current_client'):
            self.selected_clients.add(self.current_client)
            self.update_selection_display()
            self.update_preview()

    def remove_from_selection(self):
        """Remove selected client from selection"""
        selection = self.selected_listbox.curselection()
        if selection:
            idx = selection[0]
            client = self.selected_listbox.get(idx)
            self.selected_clients.remove(client)
            self.update_selection_display()
            self.update_preview()

    def update_selection_display(self):
        """Update listbox with selected clients"""
        self.selected_listbox.delete(0, END)
        for client in sorted(self.selected_clients):
            self.selected_listbox.insert(END, client)

    def update_preview(self):
        """Update the print preview with selected clients"""
        self.preview_images = []
        for client_no in self.selected_clients:
            client_data = self.fetch_callback(client_no)
            if client_data:
                fig = create_print_figure(
                    self.image,
                    client_data,
                    self.mappings,
                    self.paper_dims
                )
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=100, bbox_inches='tight', pad_inches=0)
                buf.seek(0)
                image = Image.open(buf)
                self.preview_images.append(image)

        self.current_page = 0
        if self.preview_images:
            self.show_page(self.current_page)

    def show_page(self, page_number):
        """Show the specified page in the preview"""
        if self.preview_images and 0 <= page_number < len(self.preview_images):
            image = self.preview_images[page_number]
            photo = ImageTk.PhotoImage(image)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(200, 300, image=photo)
            self.preview_canvas.image = photo
            self.page_label.config(text=f"Page {page_number + 1} of {len(self.preview_images)}")

    def show_prev_page(self):
        """Show the previous page in the preview"""
        if self.current_page > 0:
            self.current_page -= 1
            self.show_page(self.current_page)

    def show_next_page(self):
        """Show the next page in the preview"""
        if self.current_page < len(self.preview_images) - 1:
            self.current_page += 1
            self.show_page(self.current_page)

    def print_selected(self):
        """Print selected client forms"""
        client_numbers = list(self.selected_clients)
        if not client_numbers:
            self.display_message("No clients selected", "warning")
            return

        try:
            self.print_callback(client_numbers)
            self.print_history.extend(client_numbers)
            self.update_history_display()
            self.display_message("Print successful", "success")
        except Exception as e:
            self.display_message(f"Print error: {str(e)}", "error")

    def update_history_display(self):
        """Update listbox with print history"""
        self.history_listbox.delete(0, END)
        for client in self.print_history:
            self.history_listbox.insert(END, client)

    def display_message(self, message, message_type):
        """Display message in the dashboard"""
        colors = {
            "success": "green",
            "error": "red",
            "warning": "orange",
            "info": "blue"
        }
        self.message_label.config(text=message, fg=colors.get(message_type, "black"))

    def generate_client_range(self, start_no, end_no):
        """Generate range of client numbers"""
        try:
            base = ''.join(filter(str.isalpha, start_no))
            start_num = int(''.join(filter(str.isdigit, start_no)))
            end_num = int(''.join(filter(str.isdigit, end_no)))
            
            if not base or start_num > end_num:
                raise ValueError("Invalid client number range")

            return [f"{base}{i}" for i in range(start_num, end_num + 1)]
        except Exception:
            raise ValueError("Invalid client number format")

    def print_batch(self):
        """Handle batch printing"""
        start_no = self.batch_start_entry.get().strip().upper()
        end_no = self.batch_end_entry.get().strip().upper()
        
        if not (start_no and end_no):
            self.display_message("Please enter both start and end client numbers", "warning")
            return

        try:
            client_numbers = self.generate_client_range(start_no, end_no)
            total_clients = len(client_numbers)
            
            self.batch_progress_var.set(f"Starting batch print of {total_clients} clients...")
            
            # Print without generating previews
            self.print_callback(client_numbers)
            
            # Update history
            self.print_history.extend(client_numbers)
            self.update_history_display()
            
            self.batch_progress_var.set(f"Successfully printed {total_clients} clients")
            self.display_message("Batch print completed successfully", "success")

        except ValueError as e:
            self.display_message(str(e), "error")
        except Exception as e:
            self.display_message(f"Batch print error: {str(e)}", "error")
            self.batch_progress_var.set("Print failed")