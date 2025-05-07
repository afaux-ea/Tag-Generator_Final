import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from data_processor import DataProcessor
from excel_exporter import ExcelExporter
from utils import DEFAULT_WINDOW_SIZE
from datetime import datetime

class AnalyticalDataProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Tag Generator")
        self.root.geometry(DEFAULT_WINDOW_SIZE)
        self.root.minsize(900, 650)
        self.center_window()

        # Set ttkbootstrap theme
        self.style = tb.Style("flatly")  # Modern flat theme
        self.root = self.style.master
        # self.root.configure(bg=self.style.colors.background)  # Removed: invalid color attribute
        self.font_main = ("Segoe UI", 11)
        self.font_title = ("Segoe UI", 16, "bold")
        self.font_subtitle = ("Segoe UI", 12, "bold")
        self.font_button = ("Segoe UI", 11, "bold")

        # Components
        self.data_processor = DataProcessor()
        self.excel_exporter = ExcelExporter()
        self.file_path = tk.StringVar()
        self.selected_well = tk.StringVar()
        self.df = None
        self.analyte_vars = []
        self.saved_tags = []
        self.well_vars = {}  # Dictionary to store well checkbox variables
        self.show_detections_only = tk.BooleanVar(value=False)  # Variable for detection filter
        self.select_all_wells = tk.BooleanVar(value=False)  # New variable for select all wells
        
        self.create_widgets()

    def create_widgets(self):
        main_container = tb.Frame(self.root, padding=30, bootstyle="light")
        main_container.pack(fill="both", expand=True)

        # Title
        title_label = tb.Label(main_container, text="Tag Generator", font=self.font_title, bootstyle="dark")
        title_label.grid(row=0, column=0, sticky="w", padx=5, pady=(0, 25))

        # File Selection
        file_frame = tb.Labelframe(main_container, text="File Selection", padding=15, bootstyle="primary")
        file_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 18))
        file_content = tb.Frame(file_frame)
        file_content.pack(fill="x", expand=True)
        tb.Label(file_content, text="File:", font=self.font_main).pack(side="left")
        tb.Entry(file_content, textvariable=self.file_path, width=54, font=self.font_main).pack(side="left", padx=12)
        tb.Button(file_content, text="Browse", command=self.browse_file, bootstyle="info-outline", width=10).pack(side="left", padx=2)

        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_rowconfigure(0, weight=0)  # Title row
        main_container.grid_rowconfigure(1, weight=0)  # File selection row
        main_container.grid_rowconfigure(2, weight=1)  # Well selection row
        main_container.grid_rowconfigure(3, weight=1)  # Bottom container row
        
        # Well Selection
        well_frame = tb.Labelframe(main_container, text="Well Selection", padding="15", bootstyle="primary")
        well_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=(0, 15))
        
        # Add select all checkbox at the top of well selection
        select_all_frame = tb.Frame(well_frame)
        select_all_frame.pack(fill="x", pady=(0, 5))
        tb.Checkbutton(select_all_frame, text="Select All Wells", 
                       variable=self.select_all_wells, 
                       command=self.toggle_all_wells).pack(side="left")
        
        well_content = tb.Frame(well_frame)
        well_content.pack(fill="both", expand=True)
        
        # Create a canvas and scrollable frame for wells
        self.well_canvas = tk.Canvas(well_content)
        well_scrollbar = tb.Scrollbar(well_content, orient="vertical", command=self.well_canvas.yview)
        
        self.well_scrollable_frame = tb.Frame(self.well_canvas)
        self.well_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.well_canvas.configure(scrollregion=self.well_canvas.bbox("all"))
        )
        
        self.well_canvas.create_window((0, 0), window=self.well_scrollable_frame, anchor="nw")
        self.well_canvas.configure(yscrollcommand=well_scrollbar.set)
        
        # Pack the canvas and scrollbar
        self.well_canvas.pack(side="left", fill="both", expand=True, padx=(5, 0))
        well_scrollbar.pack(side="right", fill="y")
        
        # --- Mousewheel binding helpers for both canvas and frame ---
        def bind_mousewheel(widget, canvas):
            def on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            widget.bind('<Enter>', lambda e: widget.bind_all('<MouseWheel>', on_mousewheel))
            widget.bind('<Leave>', lambda e: widget.unbind_all('<MouseWheel>'))

        # Bind for well selection area
        bind_mousewheel(self.well_canvas, self.well_canvas)
        bind_mousewheel(self.well_scrollable_frame, self.well_canvas)
        
        # Create a frame to contain both analyte selection and buttons
        bottom_container = tb.Frame(main_container)
        bottom_container.grid(row=3, column=0, sticky="nsew", padx=5)
        
        # Configure grid weights for bottom_container
        bottom_container.grid_columnconfigure(0, weight=1)
        bottom_container.grid_rowconfigure(0, weight=1)
        bottom_container.grid_rowconfigure(1, weight=0)  # No weight for button row
        
        # Analyte Selection
        analyte_frame = tb.Labelframe(bottom_container, text="Analyte Selection", padding="15", bootstyle="primary")
        analyte_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        
        # Create analyte selection frame with scrollbar
        self.analyte_canvas = tk.Canvas(analyte_frame)
        analyte_scrollbar = tb.Scrollbar(analyte_frame, orient="vertical", command=self.analyte_canvas.yview)
        
        self.scrollable_frame = tb.Frame(self.analyte_canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.analyte_canvas.configure(scrollregion=self.analyte_canvas.bbox("all"))
        )
        
        self.analyte_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.analyte_canvas.configure(yscrollcommand=analyte_scrollbar.set)
        
        # Pack the canvas and scrollbar
        self.analyte_canvas.pack(side="left", fill="both", expand=True, padx=(5, 0))
        analyte_scrollbar.pack(side="right", fill="y")
        
        # Bind for analyte selection area
        bind_mousewheel(self.analyte_canvas, self.analyte_canvas)
        bind_mousewheel(self.scrollable_frame, self.analyte_canvas)
        
        # Add detection filter checkbox
        filter_frame = tb.Frame(analyte_frame)
        filter_frame.pack(fill="x", pady=(0, 10))
        tb.Label(filter_frame, text="Select Analytes:").pack(side="left", padx=(0, 10))
        tb.Checkbutton(filter_frame, text="Show Only Detections", variable=self.show_detections_only, 
                      command=lambda: self.on_well_selected(None)).pack(side="left")
        
        # Bottom buttons container
        button_container = tb.Frame(bottom_container)
        button_container.grid(row=1, column=0, sticky="ew", pady=(0, 5))
        
        # Center the buttons
        button_container.grid_columnconfigure(0, weight=1)  # Left padding
        button_container.grid_columnconfigure(3, weight=1)  # Right padding
        
        # Save Tag and Export buttons
        self.save_button = tb.Button(button_container, text="Save Tag", command=self.save_tag, 
                                    state="disabled", bootstyle="success", width=15)
        self.save_button.grid(row=0, column=1, padx=5)
        
        self.export_all_button = tb.Button(button_container, text="Export All Tags", 
                                          command=self.export_all_tags, state="disabled", 
                                          bootstyle="success", width=15)
        self.export_all_button.grid(row=0, column=2, padx=5)
        
        # Add saved tags counter with modern styling
        self.tags_label = tb.Label(bottom_container, text="Saved Tags: 0", 
                                  font=("Segoe UI", 10))
        self.tags_label.grid(row=2, column=0, pady=5)

    def bind_mousewheel(self, widget, canvas):
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        widget.bind('<Enter>', lambda e: widget.bind_all('<MouseWheel>', on_mousewheel))
        widget.bind('<Leave>', lambda e: widget.unbind_all('<MouseWheel>'))

    def bind_mousewheel_to_children(self, frame, canvas):
        for child in frame.winfo_children():
            self.bind_mousewheel(child, canvas)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[
                ("Supported files", "*.csv;*.xlsx"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.file_path.set(filename)
            self.load_data()

    def load_data(self):
        try:
            self.df = self.data_processor.load_file(self.file_path.get())
            wells = self.data_processor.get_wells(self.df)
            
            # Clear existing well checkboxes
            for widget in self.well_scrollable_frame.winfo_children():
                widget.destroy()
            self.well_vars.clear()
            
            # Create checkboxes for each well
            for well in wells:
                var = tk.BooleanVar()
                self.well_vars[well] = var
                tb.Checkbutton(
                    self.well_scrollable_frame,
                    text=well,
                    variable=var,
                    command=self.on_well_selected
                ).pack(anchor="w", padx=5, pady=2)
            
            # Bind mousewheel to children
            self.bind_mousewheel_to_children(self.well_scrollable_frame, self.well_canvas)

        except Exception as e:
            messagebox.showerror("Error", f"Error loading data: {str(e)}")

    def on_well_selected(self, event=None):
        if self.df is not None:
            # Get selected wells
            selected_wells = [well for well, var in self.well_vars.items() if var.get()]
            if not selected_wells:
                # Clear existing checkboxes if no wells selected
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                self.analyte_vars.clear()
                self.save_button.config(state="disabled")
                return
            
            # Clear existing checkboxes
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.analyte_vars.clear()
            
            # Get data for the first selected well to use as a template
            first_well = selected_wells[0]
            well_data = self.data_processor.get_well_data(self.df, first_well)
            
            # Create checkboxes for each analyte
            if well_data is not None:
                selected_wells_count = len(selected_wells)
                wells_text = f" (will apply to {selected_wells_count} selected well{'s' if selected_wells_count > 1 else ''})"
                
                # Create a label for analyte selection
                tb.Label(self.scrollable_frame, text=f"Select Analytes{wells_text}:").pack(anchor="w", pady=(0, 5))
                
                for analyte in well_data['analytes']:
                    var = tk.BooleanVar()
                    self.analyte_vars.append((var, analyte['name']))
                    tb.Checkbutton(self.scrollable_frame, text=analyte['name'], variable=var).pack(anchor="w")
                
                # Bind mousewheel to children
                self.bind_mousewheel_to_children(self.scrollable_frame, self.analyte_canvas)
                
                self.save_button.config(state="normal")
            else:
                self.save_button.config(state="disabled")

    def toggle_all_wells(self):
        """Handle select all wells checkbox"""
        select_all = self.select_all_wells.get()
        for var in self.well_vars.values():
            var.set(select_all)
        # Trigger well selection update
        self.on_well_selected()

    def save_tag(self):
        selected_wells = [well for well, var in self.well_vars.items() if var.get()]
        if not selected_wells:
            messagebox.showerror("Error", "Please select at least one well")
            return

        selected_analytes = []
        for var, name in self.analyte_vars:
            if var.get():
                selected_analytes.append(name)

        if not selected_analytes:
            messagebox.showerror("Error", "Please select at least one analyte")
            return

        # Process each selected well
        for well_id in selected_wells:
            well_data = self.data_processor.get_well_data(self.df, well_id)
            
            if well_data is None:
                messagebox.showwarning("Warning", f"Could not find data for well {well_id}")
                continue

            # Filter analytes based on selection
            analyte_dict = {a['name']: a for a in well_data['analytes']}
            filtered_analytes = []
            for name in selected_analytes:
                if name in analyte_dict:
                    analyte = analyte_dict[name]
                    # Skip ND values if show_detections_only is checked
                    if self.show_detections_only.get() and analyte['value'] == 'ND':
                        continue
                    filtered_analytes.append({
                        'name': name,
                        'value': analyte['value'],
                        'exceeds_awqs': analyte['exceeds']
                    })
            
            # Only create a tag if there are analytes to include after filtering
            if filtered_analytes:
                tag = {
                    'well_id': well_id,
                    'date': well_data['date'],
                    'analytes': filtered_analytes
                }
                self.saved_tags.append(tag)

        if self.saved_tags:
            self.tags_label.config(text=f"Saved Tags: {len(self.saved_tags)}")
            self.export_all_button.config(state="normal")
            messagebox.showinfo("Success", f"Tags saved successfully for {len(self.saved_tags)} well(s)")
        else:
            messagebox.showinfo("Info", "No tags were created as no detections were found for the selected criteria")

    def export_all_tags(self):
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Tags As"
            )
            if not filename:  # User cancelled
                return
                
            output_file = self.excel_exporter.export_tags(self.saved_tags, filename)
            messagebox.showinfo("Success", f"All tags exported to {output_file}")
            self.saved_tags = []
            self.tags_label.config(text="Saved Tags: 0")
            self.export_all_button.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting tags: {str(e)}")

    def center_window(self):
        # Get screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Calculate position coordinates
        x = (screen_width - 900) // 2
        y = (screen_height - 650) // 2
        
        # Set the position of the window to the center of the screen
        self.root.geometry(f"900x650+{x}+{y}")

def main():
    root = tk.Tk()
    app = AnalyticalDataProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()