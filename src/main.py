import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from analyzer import ExcelAnalyzer
from excel_exporter import ExcelExporter
import json
import os
import subprocess
import platform
import pandas as pd


class ExcelAnalysisApp:
    """Main GUI application for Excel analysis."""

    def __init__(self, root):
        self.root = root
        self.root.title("DataLens - Analyze 300 Columns in 3 Clicks")
        self.root.geometry("1200x850")

        # Modern, bright color scheme
        self.bg_color = "#FFFFFF"
        self.primary_color = "#2E86DE"
        self.primary_dark = "#1E5F99"
        self.secondary_color = "#EBF5FB"
        self.accent_color = "#10AC84"
        self.text_color = "#2C3E50"
        self.text_light = "#7F8C8D"
        self.border_color = "#D5DBDB"

        self.root.configure(bg=self.bg_color)

        # Configure modern fonts
        self.title_font = ("Segoe UI", 18, "bold")
        self.heading_font = ("Segoe UI", 13, "bold")
        self.normal_font = ("Segoe UI", 11)
        self.button_font = ("Segoe UI", 10, "bold")
        self.small_font = ("Segoe UI", 9)

        self.analyzer = None
        self.current_file = None
        self.current_results = None
        self.current_group_column = None
        self.last_export_path = None
        self.available_sheets = []
        self.current_sheet = None
        self.excel_file = None  # Store the Excel file object

        self.setup_menu()
        self.setup_ui()

    def setup_menu(self):
        """Set up the menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Excel File", command=self.browse_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Open Last Export", command=self.open_last_export, state="disabled")
        self.view_menu = view_menu

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="How to Use", command=self.show_help)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

    def setup_ui(self):
        """Set up the user interface with modern design."""

        # Configure style
        style = ttk.Style()
        style.theme_use('clam')

        # Configure custom styles
        style.configure('Header.TFrame', background=self.primary_color)
        style.configure('Content.TFrame', background=self.bg_color)
        style.configure('Card.TFrame', background=self.secondary_color, relief='flat')

        style.configure('Title.TLabel',
                       background=self.primary_color,
                       foreground='white',
                       font=self.title_font,
                       padding=20)

        style.configure('Heading.TLabel',
                       background=self.bg_color,
                       foreground=self.text_color,
                       font=self.heading_font,
                       padding=(0, 10, 0, 5))

        style.configure('Normal.TLabel',
                       background=self.bg_color,
                       foreground=self.text_color,
                       font=self.normal_font)

        style.configure('Primary.TButton',
                       font=self.button_font,
                       padding=(20, 10))

        style.map('Primary.TButton',
                 background=[('active', self.primary_dark), ('!disabled', self.primary_color)],
                 foreground=[('!disabled', 'white')])

        # Header Frame with Title
        header_frame = ttk.Frame(self.root, style='Header.TFrame')
        header_frame.pack(fill="x")

        header_content = ttk.Frame(header_frame, style='Header.TFrame')
        header_content.pack(fill="x", padx=30, pady=15)

        title_label = ttk.Label(header_content,
                               text="DataLens 📊",
                               style='Title.TLabel')
        title_label.pack(side="left")

        subtitle = tk.Label(header_content,
                           text="Explore 300 columns in 3 clicks  •  by Ashraf Sinan",
                           background=self.primary_color,
                           foreground="white",
                           font=self.small_font)
        subtitle.pack(side="right", pady=5)

        # Main Content Area
        content_frame = ttk.Frame(self.root, style='Content.TFrame')
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # File Selection Card
        file_card = ttk.Frame(content_frame, style='Card.TFrame', relief='solid', borderwidth=1)
        file_card.pack(fill="x", pady=(0, 15))

        file_inner = tk.Frame(file_card, bg=self.secondary_color)
        file_inner.pack(fill="x", padx=20, pady=15)

        step1_label = tk.Label(file_inner,
                              text="STEP 1: Select Your Excel File",
                              font=self.heading_font,
                              bg=self.secondary_color,
                              fg=self.primary_color)
        step1_label.pack(anchor="w", pady=(0, 10))

        file_select_frame = tk.Frame(file_inner, bg=self.secondary_color)
        file_select_frame.pack(fill="x")

        self.file_label = tk.Label(file_select_frame,
                                   text="No file selected",
                                   font=self.normal_font,
                                   bg=self.secondary_color,
                                   fg=self.text_light,
                                   anchor="w")
        self.file_label.pack(side="left", fill="x", expand=True)

        browse_btn = tk.Button(file_select_frame,
                              text="📂 Browse Excel File",
                              command=self.browse_file,
                              font=self.button_font,
                              bg=self.primary_color,
                              fg="white",
                              activebackground=self.primary_dark,
                              activeforeground="white",
                              relief='flat',
                              padx=25,
                              pady=10,
                              cursor="hand2")
        browse_btn.pack(side="right")

        # Sheet selection (initially hidden)
        self.sheet_select_frame = tk.Frame(file_inner, bg=self.secondary_color)

        sheet_label = tk.Label(self.sheet_select_frame,
                              text="Select Sheet:",
                              font=self.normal_font,
                              bg=self.secondary_color,
                              fg=self.text_color)
        sheet_label.pack(side="left", padx=(0, 10))

        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self.sheet_select_frame,
                                        textvariable=self.sheet_var,
                                        state="readonly",
                                        width=35,
                                        font=self.normal_font)
        self.sheet_combo.pack(side="left")
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_selected)

        # Analysis Options Card
        options_card = ttk.Frame(content_frame, style='Card.TFrame', relief='solid', borderwidth=1)
        options_card.pack(fill="x", pady=(0, 15))

        options_inner = tk.Frame(options_card, bg=self.secondary_color)
        options_inner.pack(fill="x", padx=20, pady=15)

        step2_label = tk.Label(options_inner,
                              text="STEP 2: Run Analysis (Optional: Group By Column)",
                              font=self.heading_font,
                              bg=self.secondary_color,
                              fg=self.primary_color)
        step2_label.pack(anchor="w", pady=(0, 10))

        options_controls = tk.Frame(options_inner, bg=self.secondary_color)
        options_controls.pack(fill="x")

        # Group by selection
        group_label = tk.Label(options_controls,
                              text="Group By:",
                              font=self.normal_font,
                              bg=self.secondary_color,
                              fg=self.text_color)
        group_label.pack(side="left", padx=(0, 10))

        self.group_column_var = tk.StringVar(value="None")
        self.group_column_combo = ttk.Combobox(options_controls,
                                               textvariable=self.group_column_var,
                                               state="readonly",
                                               width=35,
                                               font=self.normal_font)
        self.group_column_combo.pack(side="left", padx=(0, 20))

        # Run Analysis Button
        self.run_button = tk.Button(options_controls,
                                    text="▶ Run Analysis",
                                    command=self.run_analysis,
                                    state="disabled",
                                    font=self.button_font,
                                    bg=self.accent_color,
                                    fg="white",
                                    activebackground="#0E9970",
                                    activeforeground="white",
                                    relief='flat',
                                    padx=30,
                                    pady=10,
                                    cursor="hand2")
        self.run_button.pack(side="left", padx=5)

        # Export & Open Buttons
        button_frame = tk.Frame(options_controls, bg=self.secondary_color)
        button_frame.pack(side="right")

        self.export_button = tk.Button(button_frame,
                                       text="💾 Export to Excel",
                                       command=self.export_results,
                                       state="disabled",
                                       font=self.button_font,
                                       bg=self.primary_color,
                                       fg="white",
                                       activebackground=self.primary_dark,
                                       activeforeground="white",
                                       relief='flat',
                                       padx=20,
                                       pady=10,
                                       cursor="hand2")
        self.export_button.pack(side="left", padx=5)

        self.open_excel_button = tk.Button(button_frame,
                                           text="📊 Open Excel",
                                           command=self.open_last_export,
                                           state="disabled",
                                           font=self.button_font,
                                           bg=self.text_light,
                                           fg="white",
                                           activebackground="#5F6A6A",
                                           activeforeground="white",
                                           relief='flat',
                                           padx=20,
                                           pady=10,
                                           cursor="hand2")
        self.open_excel_button.pack(side="left", padx=5)

        # Results Area with Tabs
        results_card = tk.Frame(content_frame, bg="white", relief='solid', borderwidth=1)
        results_card.pack(fill="both", expand=True)

        results_header = tk.Frame(results_card, bg=self.secondary_color)
        results_header.pack(fill="x")

        results_title = tk.Label(results_header,
                                text="STEP 3: View Results",
                                font=self.heading_font,
                                bg=self.secondary_color,
                                fg=self.primary_color)
        results_title.pack(anchor="w", padx=20, pady=15)

        # Create notebook for tabbed interface
        self.notebook = ttk.Notebook(results_card)
        self.notebook.pack(fill="both", expand=True, padx=2, pady=(0, 2))

        # Text Results Tab
        text_frame = tk.Frame(self.notebook, bg="white")
        self.notebook.add(text_frame, text="📄 Text Summary")

        self.results_text = scrolledtext.ScrolledText(text_frame,
                                                      wrap=tk.WORD,
                                                      font=("Consolas", 10),
                                                      bg="white",
                                                      relief='flat',
                                                      padx=15,
                                                      pady=15)
        self.results_text.pack(fill="both", expand=True)

        # Visual Results Tab (for future charts)
        self.visual_frame = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.visual_frame, text="📊 Visualizations")

        # Placeholder for visualizations
        self.viz_canvas = tk.Canvas(self.visual_frame, bg="white", highlightthickness=0)
        self.viz_scrollbar = ttk.Scrollbar(self.visual_frame, orient="vertical", command=self.viz_canvas.yview)
        self.viz_scrollable_frame = tk.Frame(self.viz_canvas, bg="white")

        self.viz_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.viz_canvas.configure(scrollregion=self.viz_canvas.bbox("all"))
        )

        self.viz_canvas.create_window((0, 0), window=self.viz_scrollable_frame, anchor="nw")
        self.viz_canvas.configure(yscrollcommand=self.viz_scrollbar.set)

        self.viz_canvas.pack(side="left", fill="both", expand=True)
        self.viz_scrollbar.pack(side="right", fill="y")

        # Add welcome message to visual tab
        welcome_label = tk.Label(self.viz_scrollable_frame,
                                text="📊 Run analysis to see visualizations here\n\n"
                                     "Charts and statistics will appear after you click 'Run Analysis'",
                                font=self.normal_font,
                                bg="white",
                                fg=self.text_light,
                                pady=50)
        welcome_label.pack(expand=True)

        # Status Bar
        status_frame = tk.Frame(self.root, bg=self.primary_color)
        status_frame.pack(fill="x", side="bottom")

        self.status_bar = tk.Label(status_frame,
                                   text="Ready  •  Select an Excel file to get started",
                                   font=self.normal_font,
                                   bg=self.primary_color,
                                   fg="white",
                                   anchor="w",
                                   padx=20,
                                   pady=8)
        self.status_bar.pack(fill="x")

    def browse_file(self):
        """Open file dialog to select Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            try:
                self.current_file = file_path

                # Read Excel file to get available sheets
                self.excel_file = pd.ExcelFile(file_path)
                self.available_sheets = self.excel_file.sheet_names

                # Update file label
                filename = os.path.basename(file_path)
                self.file_label.config(text=f"✓ {filename}",
                                      fg=self.accent_color,
                                      font=(self.normal_font[0], self.normal_font[1], "bold"))

                # Show sheet selection if multiple sheets
                if len(self.available_sheets) > 1:
                    self.sheet_select_frame.pack(fill="x", pady=(10, 0))
                    self.sheet_combo['values'] = self.available_sheets
                    self.sheet_combo.set(self.available_sheets[0])
                    self.current_sheet = self.available_sheets[0]
                else:
                    self.sheet_select_frame.pack_forget()
                    self.current_sheet = self.available_sheets[0] if self.available_sheets else None

                # Load the selected sheet
                self.load_sheet(self.current_sheet)

            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
                self.status_bar.config(text="Error loading file")

    def on_sheet_selected(self, event=None):
        """Handle sheet selection change."""
        selected_sheet = self.sheet_var.get()
        if selected_sheet and selected_sheet != self.current_sheet:
            self.current_sheet = selected_sheet
            self.load_sheet(selected_sheet)

    def load_sheet(self, sheet_name):
        """Load a specific sheet from the Excel file."""
        if not sheet_name:
            return

        try:
            # Create analyzer with specific sheet
            self.analyzer = ExcelAnalyzer(self.current_file, sheet_name=sheet_name)

            # Update group by dropdown
            columns = ["None"] + self.analyzer.get_columns()
            self.group_column_combo['values'] = columns
            self.group_column_combo.set("None")

            # Enable run button
            self.run_button.config(state="normal")

            sheet_info = f" (Sheet: {sheet_name})" if len(self.available_sheets) > 1 else ""
            self.status_bar.config(text=f"✓ Loaded{sheet_info}: {len(self.analyzer.df)} rows, {len(self.analyzer.df.columns)} columns")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet '{sheet_name}':\n{str(e)}")
            self.status_bar.config(text=f"Error loading sheet: {sheet_name}")

    def run_analysis(self):
        """Run the analysis based on selected options."""
        if not self.analyzer:
            messagebox.showwarning("Warning", "Please select an Excel file first")
            return

        try:
            self.status_bar.config(text="Running analysis...")
            self.results_text.delete(1.0, tk.END)

            group_column = self.group_column_var.get()

            if group_column == "None":
                # Analyze all columns without grouping
                results = self.analyzer.analyze_all_columns()
                self.current_results = results
                self.current_group_column = None
                self.display_ungrouped_results(results)
            else:
                # Analyze with grouping
                results = self.analyzer.analyze_by_group(group_column)
                self.current_results = results
                self.current_group_column = group_column
                self.display_grouped_results(results, group_column)

            self.export_button.config(state="normal")
            self.create_visual_results()
            self.status_bar.config(text="✓ Analysis completed successfully  •  View results in tabs below")

        except Exception as e:
            messagebox.showerror("Error", f"Analysis failed:\n{str(e)}")
            self.status_bar.config(text="Analysis failed")

    def display_ungrouped_results(self, results):
        """Display analysis results without grouping."""
        output = "=" * 80 + "\n"
        output += "EXCEL ANALYSIS RESULTS (Ungrouped)\n"
        output += "=" * 80 + "\n\n"

        for column_name, column_data in results.items():
            output += f"\n{'─' * 80}\n"
            output += f"Column: {column_name}\n"
            output += f"{'─' * 80}\n"

            if column_data['type'] == 'quantitative':
                output += "Type: QUANTITATIVE (Numeric)\n\n"
                data = column_data['data']
                output += f"  Count:           {data['count']}\n"
                output += f"  Minimum:         {data['min']:.2f}\n"
                output += f"  25th Percentile: {data['percentile_25']:.2f}\n"
                output += f"  Median (50th):   {data['percentile_50']:.2f}\n"
                output += f"  75th Percentile: {data['percentile_75']:.2f}\n"
                output += f"  Maximum:         {data['max']:.2f}\n"
                output += f"  Average:         {data['average']:.2f}\n"
                output += f"  Sum:             {data['sum']:.2f}\n"
                output += f"  % of Total:      {data['percent_of_total']:.2f}%\n\n"
                output += "  Frequency Distribution:\n"
                output += f"  {'Value':<15} {'Freq':<10} {'% Count':<12} {'Value Sum':<15} {'% of Total'}\n"
                output += f"  {'-' * 15} {'-' * 10} {'-' * 12} {'-' * 15} {'-' * 12}\n"
                for freq_item in data['frequency']:
                    output += f"  {freq_item['value']:<15.2f} {freq_item['frequency']:<10} {freq_item['percentage']:<12.2f} {freq_item['value_sum']:<15.2f} {freq_item['percent_of_total_column']:.2f}%\n"
            else:
                output += "Type: QUALITATIVE (Categorical)\n\n"
                output += f"  {'Label':<30} {'Frequency':<15} {'Percentage'}\n"
                output += f"  {'-' * 30} {'-' * 15} {'-' * 15}\n"
                for item in column_data['data']:
                    output += f"  {item['label']:<30} {item['frequency']:<15} {item['percentage']:.2f}%\n"

        self.results_text.insert(1.0, output)

    def display_grouped_results(self, results, group_column):
        """Display analysis results with grouping."""
        output = "=" * 80 + "\n"
        output += f"EXCEL ANALYSIS RESULTS (Grouped by: {group_column})\n"
        output += "=" * 80 + "\n\n"

        for group_name, group_data in results.items():
            output += f"\n{'═' * 80}\n"
            output += f"GROUP: {group_name} ({group_data['row_count']} rows)\n"
            output += f"{'═' * 80}\n"

            for column_name, column_data in group_data['columns'].items():
                output += f"\n  Column: {column_name}\n"
                output += f"  {'-' * 76}\n"

                if column_data['type'] == 'quantitative':
                    output += "  Type: QUANTITATIVE (Numeric)\n\n"
                    data = column_data['data']
                    output += f"    Count:           {data['count']}\n"
                    output += f"    Minimum:         {data['min']:.2f}\n"
                    output += f"    25th Percentile: {data['percentile_25']:.2f}\n"
                    output += f"    Median (50th):   {data['percentile_50']:.2f}\n"
                    output += f"    75th Percentile: {data['percentile_75']:.2f}\n"
                    output += f"    Maximum:         {data['max']:.2f}\n"
                    output += f"    Average:         {data['average']:.2f}\n"
                    output += f"    Sum:             {data['sum']:.2f}\n"
                    output += f"    % of Total:      {data['percent_of_total']:.2f}%\n\n"
                    output += "    Frequency Distribution:\n"
                    output += f"    {'Value':<13} {'Freq':<8} {'% Count':<10} {'Val Sum':<13} {'% Total'}\n"
                    output += f"    {'-' * 13} {'-' * 8} {'-' * 10} {'-' * 13} {'-' * 10}\n"
                    for freq_item in data['frequency']:
                        output += f"    {freq_item['value']:<13.2f} {freq_item['frequency']:<8} {freq_item['percentage']:<10.2f} {freq_item['value_sum']:<13.2f} {freq_item['percent_of_total_column']:.2f}%\n"
                else:
                    output += "  Type: QUALITATIVE (Categorical)\n\n"
                    output += f"    {'Label':<28} {'Frequency':<13} {'Percentage'}\n"
                    output += f"    {'-' * 28} {'-' * 13} {'-' * 13}\n"
                    for item in column_data['data']:
                        output += f"    {item['label']:<28} {item['frequency']:<13} {item['percentage']:.2f}%\n"

        self.results_text.insert(1.0, output)

    def export_results(self):
        """Export analysis results to Excel file with visualizations."""
        if not self.current_results:
            messagebox.showwarning("Warning", "No results to export")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if file_path:
            try:
                self.status_bar.config(text="Exporting to Excel...")

                exporter = ExcelExporter(self.current_results, self.current_group_column)

                if self.current_group_column:
                    exporter.export_grouped(file_path)
                else:
                    exporter.export_ungrouped(file_path)

                self.last_export_path = file_path
                self.open_excel_button.config(state="normal")
                self.view_menu.entryconfig("Open Last Export", state="normal")

                messagebox.showinfo("Success",
                    "Results exported successfully!\n\n"
                    "The Excel file contains:\n"
                    "- Index sheet with links to all columns\n"
                    "- Visualizations overview sheet with all charts\n"
                    "- Individual sheet for each column\n"
                    "- Histogram and Distribution Density charts\n\n"
                    "Click 'Open Excel File' to view the results.")
                self.status_bar.config(text=f"Results exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export results:\n{str(e)}")
                self.status_bar.config(text="Export failed")


    def create_visual_results(self):
        """Create visual representation of results in the Visualizations tab."""
        # Clear previous visualizations
        for widget in self.viz_scrollable_frame.winfo_children():
            widget.destroy()

        if not self.current_results:
            return

        # Title
        title_label = tk.Label(self.viz_scrollable_frame,
                              text="📊 Visual Analysis Summary",
                              font=self.title_font,
                              bg="white",
                              fg=self.primary_color,
                              pady=20)
        title_label.pack(anchor="w", padx=30)

        # Create statistics cards for each column
        if self.current_group_column:
            # Grouped results
            for group_name, group_data in self.current_results.items():
                # Group header
                group_frame = tk.Frame(self.viz_scrollable_frame, bg="white")
                group_frame.pack(fill="x", padx=30, pady=10)

                group_label = tk.Label(group_frame,
                                      text=f"GROUP: {group_name} ({group_data['row_count']} rows)",
                                      font=self.heading_font,
                                      bg=self.secondary_color,
                                      fg=self.primary_color,
                                      anchor="w",
                                      padx=15,
                                      pady=10)
                group_label.pack(fill="x")

                # Column cards
                for column_name, column_data in group_data['columns'].items():
                    self._create_stat_card(self.viz_scrollable_frame, column_name, column_data, group_name)
        else:
            # Ungrouped results
            for column_name, column_data in self.current_results.items():
                self._create_stat_card(self.viz_scrollable_frame, column_name, column_data)

    def _create_stat_card(self, parent, column_name, column_data, group_name=None):
        """Create a visual statistics card for a column."""
        # Card container
        card = tk.Frame(parent, bg="white", relief='solid', borderwidth=1, highlightbackground=self.border_color)
        card.pack(fill="x", padx=30, pady=10)

        # Card header
        header_bg = self.primary_color if column_data['type'] == 'quantitative' else self.accent_color
        header = tk.Frame(card, bg=header_bg)
        header.pack(fill="x")

        col_title = tk.Label(header,
                            text=f"{'📊' if column_data['type'] == 'quantitative' else '📋'} {column_name}",
                            font=self.heading_font,
                            bg=header_bg,
                            fg="white",
                            anchor="w",
                            padx=20,
                            pady=12)
        col_title.pack(side="left")

        type_label = tk.Label(header,
                             text=column_data['type'].upper(),
                             font=self.small_font,
                             bg=header_bg,
                             fg="white",
                             padx=10,
                             pady=5)
        type_label.pack(side="right", padx=20)

        # Card content
        content = tk.Frame(card, bg="white")
        content.pack(fill="both", padx=20, pady=15)

        if column_data['type'] == 'quantitative':
            self._create_quantitative_viz(content, column_data['data'])
        else:
            self._create_qualitative_viz(content, column_data['data'])

    def _create_quantitative_viz(self, parent, data):
        """Create visualization for quantitative data."""
        # Statistics grid
        stats_frame = tk.Frame(parent, bg="white")
        stats_frame.pack(fill="x", pady=(0, 15))

        stats = [
            ("Count", str(data['count']), "#3498DB"),
            ("Min", f"{data['min']:.2f}", "#E74C3C"),
            ("Median", f"{data['percentile_50']:.2f}", "#9B59B6"),
            ("Average", f"{data['average']:.2f}", "#2ECC71"),
            ("Max", f"{data['max']:.2f}", "#E67E22"),
            ("Sum", f"{data['sum']:.2f}", "#1ABC9C"),
        ]

        for idx, (label, value, color) in enumerate(stats):
            stat_box = tk.Frame(stats_frame, bg=color, relief='flat')
            stat_box.grid(row=0, column=idx, padx=5, sticky="ew")
            stats_frame.columnconfigure(idx, weight=1)

            value_label = tk.Label(stat_box,
                                  text=value,
                                  font=(self.normal_font[0], 16, "bold"),
                                  bg=color,
                                  fg="white",
                                  pady=8)
            value_label.pack()

            label_text = tk.Label(stat_box,
                                 text=label,
                                 font=self.small_font,
                                 bg=color,
                                 fg="white",
                                 pady=5)
            label_text.pack()

        # Simple bar visualization of frequency
        if len(data['frequency']) > 0 and len(data['frequency']) <= 20:
            viz_frame = tk.Frame(parent, bg="white")
            viz_frame.pack(fill="x", pady=10)

            viz_title = tk.Label(viz_frame,
                                text="Top Value Frequencies:",
                                font=self.normal_font,
                                bg="white",
                                fg=self.text_color,
                                anchor="w")
            viz_title.pack(anchor="w", pady=(0, 10))

            max_freq = max(item['frequency'] for item in data['frequency'][:10])

            for freq_item in data['frequency'][:10]:
                bar_frame = tk.Frame(viz_frame, bg="white")
                bar_frame.pack(fill="x", pady=2)

                label_width = 15
                value_label = tk.Label(bar_frame,
                                      text=f"{freq_item['value']:.2f}",
                                      font=self.small_font,
                                      bg="white",
                                      fg=self.text_color,
                                      width=label_width,
                                      anchor="w")
                value_label.pack(side="left")

                bar_container = tk.Frame(bar_frame, bg="#ECF0F1", height=20)
                bar_container.pack(side="left", fill="x", expand=True, padx=(5, 5))

                bar_width = int((freq_item['frequency'] / max_freq) * 100) if max_freq > 0 else 0
                bar = tk.Frame(bar_container, bg=self.primary_color, height=20)
                bar.place(relwidth=bar_width/100, relheight=1)

                freq_label = tk.Label(bar_frame,
                                     text=f"{freq_item['frequency']} ({freq_item['percentage']:.1f}%)",
                                     font=self.small_font,
                                     bg="white",
                                     fg=self.text_color,
                                     width=15,
                                     anchor="e")
                freq_label.pack(side="right")

    def _create_qualitative_viz(self, parent, data):
        """Create visualization for qualitative data."""
        if len(data) == 0:
            return

        # Simple bar visualization
        viz_frame = tk.Frame(parent, bg="white")
        viz_frame.pack(fill="x")

        max_freq = max(item['frequency'] for item in data[:15])

        for item in data[:15]:
            bar_frame = tk.Frame(viz_frame, bg="white")
            bar_frame.pack(fill="x", pady=3)

            label_width = 25
            value_label = tk.Label(bar_frame,
                                  text=str(item['label'])[:30],
                                  font=self.normal_font,
                                  bg="white",
                                  fg=self.text_color,
                                  width=label_width,
                                  anchor="w")
            value_label.pack(side="left")

            bar_container = tk.Frame(bar_frame, bg="#ECF0F1", height=25)
            bar_container.pack(side="left", fill="x", expand=True, padx=(10, 10))

            bar_width = int((item['frequency'] / max_freq) * 100) if max_freq > 0 else 0
            bar = tk.Frame(bar_container, bg=self.accent_color, height=25)
            bar.place(relwidth=bar_width/100, relheight=1)

            freq_label = tk.Label(bar_frame,
                                 text=f"{item['frequency']} ({item['percentage']:.1f}%)",
                                 font=self.normal_font,
                                 bg="white",
                                 fg=self.text_color,
                                 width=18,
                                 anchor="e")
            freq_label.pack(side="right")

    def open_last_export(self):
        """Open the last exported Excel file."""
        if not self.last_export_path or not os.path.exists(self.last_export_path):
            messagebox.showwarning("Warning", "No exported file found or file has been moved/deleted.")
            return

        try:
            # Open the Excel file with the default application
            if platform.system() == 'Windows':
                os.startfile(self.last_export_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', self.last_export_path])
            else:  # Linux
                subprocess.run(['xdg-open', self.last_export_path])

            self.status_bar.config(text=f"Opened {os.path.basename(self.last_export_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file:\n{str(e)}")

    def show_help(self):
        """Display help information."""
        help_window = tk.Toplevel(self.root)
        help_window.title("How to Use - DataLens")
        help_window.geometry("700x600")
        help_window.configure(bg=self.bg_color)

        # Create scrolled text for help content
        help_text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, font=("Segoe UI", 10), bg="white")
        help_text.pack(fill="both", expand=True, padx=20, pady=20)

        help_content = """
EXCEL ANALYSIS SYSTEM - USER GUIDE

═══════════════════════════════════════════════════════════════

1. LOADING DATA
   • Click "Browse Excel File" or use File > Open Excel File
   • Select your Excel file (.xlsx or .xls format)
   • The system will load and display column information

2. ANALYZING DATA
   • Choose grouping (optional):
     - Select "None" for overall analysis
     - Select a column name to group data by that column
   • Click "Run Analysis" to process the data
   • Results will appear in the analysis results area

3. EXPORTING RESULTS
   • After analysis, click "Export Results"
   • Choose a location and filename for the output
   • The system will create an Excel file containing:
     ✓ Index sheet with navigation links
     ✓ Visualizations overview with all charts
     ✓ Individual sheets for each column
     ✓ Statistical summaries for numeric columns
     ✓ Frequency distributions for all columns
     ✓ Histogram and Distribution Density charts

4. VIEWING EXPORTED FILES
   • Click "Open Excel File" to view your last export
   • Or use View > Open Last Export from the menu

5. UNDERSTANDING THE OUTPUT

   QUANTITATIVE COLUMNS (Numeric Data):
   • Statistical summary (min, max, median, average, etc.)
   • Frequency distribution table
   • Histogram chart showing value frequencies
   • Distribution Density chart showing data spread

   QUALITATIVE COLUMNS (Text/Categorical Data):
   • Category frequency table
   • Pie or bar chart showing category distribution

6. TIPS
   • Larger datasets may take longer to process
   • Group by categorical columns for meaningful comparisons
   • All charts are interactive in the Excel file
   • Use the Index sheet for easy navigation

═══════════════════════════════════════════════════════════════

For additional support or questions, contact the system owner.
"""

        help_text.insert(1.0, help_content)
        help_text.config(state=tk.DISABLED)

        # Close button
        close_btn = ttk.Button(help_window, text="Close", command=help_window.destroy)
        close_btn.pack(pady=10)

    def show_about(self):
        """Display about dialog."""
        about_window = tk.Toplevel(self.root)
        about_window.title("About - DataLens")
        about_window.geometry("500x400")
        about_window.configure(bg=self.bg_color)
        about_window.resizable(False, False)

        # Center the window
        about_window.transient(self.root)
        about_window.grab_set()

        # Main frame
        main_frame = tk.Frame(about_window, bg=self.bg_color)
        main_frame.pack(fill="both", expand=True, padx=30, pady=30)

        # Title
        title_label = tk.Label(main_frame, text="DataLens",
                              font=("Segoe UI", 20, "bold"),
                              bg=self.bg_color, fg=self.primary_color)
        title_label.pack(pady=(0, 10))

        # Version
        version_label = tk.Label(main_frame, text="Version 1.0",
                                font=("Segoe UI", 11),
                                bg=self.bg_color, fg=self.text_color)
        version_label.pack(pady=(0, 20))

        # Separator
        separator = tk.Frame(main_frame, height=2, bg=self.primary_color)
        separator.pack(fill="x", pady=10)

        # Owner/Designer information
        info_frame = tk.Frame(main_frame, bg=self.bg_color)
        info_frame.pack(pady=20)

        designer_label = tk.Label(info_frame, text="Designed and Developed by:",
                                 font=("Segoe UI", 11),
                                 bg=self.bg_color, fg=self.text_color)
        designer_label.pack()

        name_label = tk.Label(info_frame, text="Ashraf Sinan",
                             font=("Segoe UI", 16, "bold"),
                             bg=self.bg_color, fg=self.primary_color)
        name_label.pack(pady=(5, 15))

        # Description
        desc_text = """A comprehensive Excel analysis tool featuring:
• Statistical analysis and frequency distributions
• Interactive visualizations and charts
• Group-by analysis capabilities
• Professional Excel reports with navigation"""

        desc_label = tk.Label(info_frame, text=desc_text,
                             font=("Segoe UI", 9),
                             bg=self.bg_color, fg=self.text_color,
                             justify=tk.LEFT)
        desc_label.pack(pady=(0, 20))

        # Copyright
        copyright_label = tk.Label(info_frame, text="© 2024 All Rights Reserved",
                                  font=("Segoe UI", 9),
                                  bg=self.bg_color, fg=self.text_color)
        copyright_label.pack()

        # Close button
        close_btn = ttk.Button(main_frame, text="Close", command=about_window.destroy)
        close_btn.pack(pady=15)


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = ExcelAnalysisApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
