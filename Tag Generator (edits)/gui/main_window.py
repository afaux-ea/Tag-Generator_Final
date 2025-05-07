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

        # Set a consistent app background color (matching ttkbootstrap flatly theme)
        self.app_bg = '#f8f9fa'
        self.root.configure(bg=self.app_bg)
        # Custom styles for seamless background and black border
        self.style = tb.Style("flatly")  # Modern flat theme
        self.root = self.style.master
        self.style.configure("Custom.TFrame", background=self.app_bg)
        self.style.configure("Bordered.TLabelframe", background=self.app_bg, bordercolor="black", borderwidth=2, relief="solid")
        self.style.configure("Bordered.TLabelframe.Label", background=self.app_bg)

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
        
        # New state for historical file and sampling date selections
        self.is_historical = False
        self.well_sampling_dates = {}  # {well: [(col_idx, date), ...]}
        self.selected_sampling_dates = {}  # {well: {date: tk.BooleanVar}}

        self.create_widgets()

    def create_widgets(self):
        # Main container with matching background and no border
        main_container = tb.Frame(self.root, padding=(0, 0, 0, 0), style="Custom.TFrame")
        main_container.pack(fill="both", expand=True)

        # Title (no box, just a label with matching bg)
        title_label = tk.Label(main_container, text="Tag Generator", font=self.font_title, fg="#495057", bg=self.app_bg)
        title_label.grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10), columnspan=2)

        # File Selection
        file_frame = tb.Labelframe(main_container, text="File Selection", padding=15, style="Custom.TLabelframe")
        file_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 18), columnspan=2)
        file_content = tb.Frame(file_frame, style="Custom.TFrame")
        file_content.pack(fill="x", expand=True)
        tb.Label(file_content, text="File:", font=self.font_main, background=self.app_bg).pack(side="left")
        tb.Entry(file_content, textvariable=self.file_path, width=54, font=self.font_main).pack(side="left", padx=12)
        tb.Button(file_content, text="Browse", command=self.browse_file, bootstyle="info-outline", width=10).pack(side="left", padx=2)

        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)
        main_container.grid_rowconfigure(2, weight=1, minsize=250)
        main_container.grid_rowconfigure(3, weight=0)

        # --- Horizontal container for well and analyte selection ---
        selection_container = tb.Frame(main_container, style="Custom.TFrame")
        selection_container.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=(0, 10))
        selection_container.grid_columnconfigure(0, weight=1, uniform="sel")
        selection_container.grid_columnconfigure(1, weight=1, uniform="sel")
        selection_container.grid_columnconfigure(2, weight=1, uniform="sel")
        selection_container.grid_rowconfigure(0, weight=1, minsize=250)

        # Set consistent width and height for all selection panels
        panel_width = 270
        panel_height = 300

        # Well Selection
        well_frame = tb.Labelframe(selection_container, text="Well Selection", padding="15", style="Bordered.TLabelframe", width=panel_width, height=panel_height)
        well_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 7), pady=0)
        well_frame.grid_propagate(False)
        well_frame.grid_rowconfigure(0, weight=1, minsize=200)
        well_frame.grid_columnconfigure(0, weight=1)

        select_all_frame = tb.Frame(well_frame, style="Custom.TFrame")
        select_all_frame.pack(fill="x", pady=(0, 5))
        tb.Checkbutton(select_all_frame, text="Select All Wells", variable=self.select_all_wells, command=self.toggle_all_wells, bootstyle="light").pack(side="left")

        well_content = tb.Frame(well_frame, height=200, style="Custom.TFrame")
        well_content.pack(fill="both", expand=True)
        well_content.pack_propagate(False)
        well_content.grid_rowconfigure(0, weight=1)
        well_content.grid_columnconfigure(0, weight=1)
        well_content.grid_columnconfigure(1, weight=0)

        self.well_canvas = tk.Canvas(well_content, height=200, bg=self.app_bg, highlightthickness=0, bd=0)
        well_scrollbar = tb.Scrollbar(well_content, orient="vertical", command=self.well_canvas.yview)
        self.well_scrollable_frame = tb.Frame(self.well_canvas, style="Custom.TFrame")
        self.well_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.well_canvas.configure(scrollregion=self.well_canvas.bbox("all"))
        )
        self.well_canvas.create_window((0, 0), window=self.well_scrollable_frame, anchor="nw")
        self.well_canvas.configure(yscrollcommand=well_scrollbar.set)
        self.well_canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        well_scrollbar.grid(row=0, column=1, sticky="ns")
        self.bind_mousewheel(self.well_canvas, self.well_canvas)
        self.bind_mousewheel(self.well_scrollable_frame, self.well_canvas)

        # Analyte Selection
        analyte_frame = tb.Labelframe(selection_container, text="Analyte Selection", padding="15", style="Bordered.TLabelframe", width=panel_width, height=panel_height)
        analyte_frame.grid(row=0, column=1, sticky="nsew", padx=(7, 7), pady=0)
        analyte_frame.grid_propagate(False)
        analyte_frame.grid_rowconfigure(0, weight=0)  # filter row
        analyte_frame.grid_rowconfigure(1, weight=1)  # analyte list row
        analyte_frame.grid_columnconfigure(0, weight=1)

        filter_frame = tb.Frame(analyte_frame, style="Custom.TFrame")
        filter_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        tb.Label(filter_frame, text="Select Analytes:", background=self.app_bg).pack(side="left", padx=(0, 10))
        tb.Checkbutton(filter_frame, text="Show Detections Only", variable=self.show_detections_only, command=self.on_well_selected, bootstyle="light").pack(side="left")

        self.analyte_canvas = tk.Canvas(analyte_frame, height=200, bg=self.app_bg, highlightthickness=0, bd=0)
        analyte_scrollbar = tb.Scrollbar(analyte_frame, orient="vertical", command=self.analyte_canvas.yview)
        self.scrollable_frame = tb.Frame(self.analyte_canvas, style="Custom.TFrame")
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.analyte_canvas.configure(scrollregion=self.analyte_canvas.bbox("all"))
        )
        self.analyte_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.analyte_canvas.configure(yscrollcommand=analyte_scrollbar.set)
        self.analyte_canvas.grid(row=1, column=0, sticky="nsew", padx=(5, 0))
        analyte_scrollbar.grid(row=1, column=1, sticky="ns")
        self.bind_mousewheel(self.analyte_canvas, self.analyte_canvas)
        self.bind_mousewheel(self.scrollable_frame, self.analyte_canvas)

        # Sampling Date Selection Panel
        self.sampling_date_frame = tb.Labelframe(selection_container, text="Sampling Date Selection", padding="15", style="Bordered.TLabelframe", width=panel_width, height=panel_height)
        self.sampling_date_frame.grid(row=0, column=2, sticky="nsew", padx=(7, 0), pady=0)
        self.sampling_date_frame.grid_propagate(False)
        self.sampling_date_frame.grid_rowconfigure(0, weight=1, minsize=200)
        self.sampling_date_frame.grid_columnconfigure(0, weight=1)
        self.sampling_date_canvas = tk.Canvas(self.sampling_date_frame, highlightthickness=0, bd=0, relief="flat", bg=self.app_bg, height=200)
        self.sampling_date_scrollbar = tb.Scrollbar(self.sampling_date_frame, orient="vertical", command=self.sampling_date_canvas.yview)
        self.sampling_date_canvas.grid(row=0, column=0, sticky="nsew")
        self.sampling_date_scrollbar.grid(row=0, column=1, sticky="ns")
        self.sampling_date_canvas.configure(yscrollcommand=self.sampling_date_scrollbar.set)
        self.sampling_date_scrollable_frame = tb.Frame(self.sampling_date_canvas, style="Custom.TFrame")
        self.sampling_date_window = self.sampling_date_canvas.create_window((0, 0), window=self.sampling_date_scrollable_frame, anchor="nw")
        self.sampling_date_widgets = {}  # {well: [widgets]}
        self.bind_mousewheel(self.sampling_date_canvas, self.sampling_date_canvas)
        self.bind_mousewheel_to_children(self.sampling_date_scrollable_frame, self.sampling_date_canvas)
        self.sampling_date_scrollable_frame.bind("<Configure>", lambda e: self._update_sampling_date_scrollregion_and_width())
        self.sampling_date_canvas.bind("<Configure>", lambda e: self._resize_sampling_date_frame())

        # --- Buttons at the bottom ---
        button_frame = tk.Frame(main_container, bg=self.app_bg)
        button_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 10), padx=0)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        self.save_button = tb.Button(button_frame, text="Save Tags", command=self.save_tag, bootstyle="success", width=16)
        self.save_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.save_button.config(state="disabled")

        self.export_all_button = tb.Button(button_frame, text="Export All Tags", command=self.export_all_tags, bootstyle="info", width=16)
        self.export_all_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.export_all_button.config(state="disabled")

        self.tags_label = tk.Label(button_frame, text="Saved Tags: 0", font=self.font_main, bg=self.app_bg)
        self.tags_label.grid(row=1, column=0, columnspan=2, pady=(5, 0))

    def bind_mousewheel(self, widget, canvas):
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        widget.bind("<Enter>", lambda e: widget.bind_all("<MouseWheel>", on_mousewheel))
        widget.bind("<Leave>", lambda e: widget.unbind_all("<MouseWheel>"))

    def bind_mousewheel_to_children(self, frame, canvas):
        for child in frame.winfo_children():
            self.bind_mousewheel(child, canvas)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")],
            title="Select Data File"
        )
        if file_path:
            self.file_path.set(file_path)
            self.load_data()

    def load_data(self):
        try:
            self.df = self.data_processor.load_file(self.file_path.get())
            # Detect if historical file
            self.is_historical = self.data_processor.is_historical_file(self.df)
            if self.is_historical:
                self.well_sampling_dates = self.data_processor.get_well_sampling_dates(self.df)
            else:
                self.well_sampling_dates = {}
            # Clear sampling date selections
            self.selected_sampling_dates = {}
            # Clear sampling date panel
            for widgets in getattr(self, 'sampling_date_widgets', {}).values():
                for widget in widgets:
                    widget.destroy()
            self.sampling_date_widgets = {}
            # Use unique well IDs for historical files
            if self.is_historical:
                wells = list(self.well_sampling_dates.keys())
            else:
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
            messagebox.showerror("Error", f"Error loading file: {str(e)}")

    def on_well_selected(self, event=None):
        if self.df is not None:
            selected_wells = [well for well, var in self.well_vars.items() if var.get()]
            # --- Sampling Date Panel Logic ---
            # Clear previous widgets
            for widgets in getattr(self, 'sampling_date_widgets', {}).values():
                for widget in widgets:
                    widget.destroy()
            self.sampling_date_widgets = {}
            # Only show panel if historical file and wells are selected
            if self.is_historical and selected_wells:
                for well in selected_wells:
                    widgets = []
                    frame = tb.Frame(self.sampling_date_scrollable_frame)
                    frame.pack(anchor="w", fill="x", pady=(0, 7))
                    tb.Label(frame, text=well, font=self.font_subtitle).pack(anchor="w")
                    self.selected_sampling_dates.setdefault(well, {})
                    for col_idx, date in self.well_sampling_dates.get(well, []):
                        var = self.selected_sampling_dates[well].get(date)
                        if var is None:
                            var = tk.BooleanVar()
                            self.selected_sampling_dates[well][date] = var
                        # Format date for display
                        try:
                            dt = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
                            display_date = dt.strftime("%B %Y")
                        except Exception:
                            display_date = date  # fallback if parsing fails
                        cb = tb.Checkbutton(frame, text=display_date, variable=var)
                        cb.pack(anchor="w", padx=15)
                        widgets.append(cb)
                    self.sampling_date_widgets[well] = [frame] + widgets
                # Update scrollregion after adding widgets
                self._update_sampling_date_scrollregion_and_width()
            # --- End Sampling Date Panel Logic ---
            if not selected_wells:
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                self.analyte_vars.clear()
                self.save_button.config(state="disabled")
                return
            # Clear existing checkboxes
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.analyte_vars.clear()
            # Use first selected well to populate analytes
            well_data = self.data_processor.get_well_data(self.df, selected_wells[0])
            if well_data and 'analytes' in well_data:
                for analyte in well_data['analytes']:
                    var = tk.BooleanVar()
                    self.analyte_vars.append((var, analyte['name']))
                    tb.Checkbutton(self.scrollable_frame, text=analyte['name'], variable=var).pack(anchor="w")
                self.bind_mousewheel_to_children(self.scrollable_frame, self.analyte_canvas)
                self.save_button.config(state="normal")
            else:
                self.save_button.config(state="disabled")

    def toggle_all_wells(self):
        select_all = self.select_all_wells.get()
        for var in self.well_vars.values():
            var.set(select_all)
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
        self.saved_tags = []  # Reset saved tags each time
        for well_id in selected_wells:
            if self.is_historical:
                # Gather all selected dates for this well
                selected_dates = [date for date, var in self.selected_sampling_dates.get(well_id, {}).items() if var.get()]
                if not selected_dates:
                    continue
                # Sort dates for consistency
                try:
                    sorted_dates = sorted(selected_dates, key=lambda d: datetime.strptime(d, "%Y-%m-%d %H:%M:%S"))
                except Exception:
                    sorted_dates = selected_dates
                analyte_rows = []
                for analyte_name in selected_analytes:
                    values = []
                    exceeds_list = []
                    for date in sorted_dates:
                        # Find column index for this well/date
                        col_idx = None
                        for col, d in self.well_sampling_dates.get(well_id, []):
                            if d == date:
                                col_idx = col
                                break
                        if col_idx is not None:
                            # Get value for this analyte/date
                            value = None
                            exceeds = False
                            for idx, val in enumerate(self.df.iloc[:, 0]):
                                if idx >= 6 and str(val).strip() == analyte_name:
                                    analyte_value = str(self.df.iloc[idx, col_idx]).strip()
                                    if 'U' in analyte_value:
                                        analyte_value = 'ND'
                                    value = analyte_value
                                    try:
                                        awqs = float(str(self.df.iloc[idx, 1]).strip())
                                        if value != 'ND' and float(value) > awqs:
                                            exceeds = True
                                    except Exception:
                                        pass
                                    break
                            values.append(value)
                            exceeds_list.append(exceeds)
                        else:
                            values.append(None)
                            exceeds_list.append(False)
                    analyte_rows.append({'name': analyte_name, 'values': values, 'exceeds': exceeds_list})
                tag = {
                    'well_id': well_id,
                    'dates': sorted_dates,
                    'analytes': analyte_rows
                }
                self.saved_tags.append(tag)
            else:
                well_data = self.data_processor.get_well_data(self.df, well_id)
                if well_data is None:
                    messagebox.showwarning("Warning", f"Could not find data for well {well_id}")
                    continue
                analyte_dict = {a['name']: a for a in well_data['analytes']}
                filtered_analytes = []
                for name in selected_analytes:
                    if name in analyte_dict:
                        analyte = analyte_dict[name]
                        if self.show_detections_only.get() and analyte['value'] == 'ND':
                            continue
                        filtered_analytes.append({
                            'name': name,
                            'value': analyte['value'],
                            'exceeds_awqs': analyte['exceeds']
                        })
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
            # messagebox.showinfo("Success", f"Tags saved successfully for {len(self.saved_tags)} well(s)")
            print(f"Tags saved successfully for {len(self.saved_tags)} well(s)")
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
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 900) // 2
        y = (screen_height - 650) // 2
        self.root.geometry(f"900x650+{x}+{y}")

    def _update_sampling_date_scrollregion_and_width(self):
        self.sampling_date_canvas.configure(scrollregion=self.sampling_date_canvas.bbox("all"))
        # Make the internal frame always match the canvas width
        canvas_width = self.sampling_date_canvas.winfo_width()
        self.sampling_date_canvas.itemconfig(self.sampling_date_window, width=canvas_width)

    def _resize_sampling_date_frame(self):
        # Called when the canvas is resized
        self._update_sampling_date_scrollregion_and_width()
