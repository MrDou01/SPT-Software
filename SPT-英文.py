import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re
import math
import os
from datetime import datetime

# Set Matplotlib font support
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False  # Fix minus sign display issue


class SPTLiquefactionSoftware:
    def __init__(self, root):
        self.root = root
        self.root.title(
            "Sand Soil Seismic Liquefaction Index Calculation and Classification Software Based on Standard Penetration Test (SPT)")
        self.root.geometry("1200x700")
        self.root.resizable(True, True)

        # Configure ttk style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=("Arial", 10))
        self.style.map("TButton", background=[("active", "#e1e1e1"), ("pressed", "#c1c1c1")])

        # ------------ Step 1: Initialize all binary choice variables (fix AttributeError core) ------------
        self.manual_acceleration_intensity_choice = tk.StringVar(value="amax")  # amax/intensity
        self.manual_particle_content_choice = tk.StringVar(value="ρc")  # ρc/Fc
        self.manual_beta_calculation_choice = tk.StringVar(value="group")  # group/formula

        # Multi-point data storage
        self.site_data = [self._create_default_site_data()]
        self.current_site_index = 0

        # Column name mapping library (only keep necessary parameters, binary choice parameters correspond to same standard column names)
        self.column_name_mapping = {
            "Site Number": ["Site Number", "Site ID", "Measurement Point Number", "ID", "Point No", "Point ID"],
            "Saturated Sand Depth ds(m)": ["Saturated Sand Depth ds(m)", "Saturation Depth(m)", "ds",
                                           "Saturated Burial Depth", "Sand Depth ds", "Burial Depth ds"],
            "Measured SPT Blow Count N": ["Measured SPT Blow Count N", "N value", "Measured Penetration Value",
                                          "Standard Penetration N", "N", "Blow Count N"],
            "Clay/Fine Particle Content(%)": ["Clay Content ρc(%)", "ρc", "Clay Content", "Clay Percentage",
                                              "Fine Particle Content Fc(%)", "Fc", "Fine Particle Percentage",
                                              "Fine Particle Content"],
            "Surface Peak Acceleration amax(g)": ["Surface Peak Acceleration amax(g)", "amax", "Peak Acceleration",
                                                  "Acceleration Peak"],
            "Seismic Intensity": ["Seismic Intensity", "Intensity", "Earthquake Grade"]
        }

        # Create menu bar
        self.create_menu_bar()

        # Create notebook interface
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Manual input tab
        self.manual_input_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.manual_input_tab, text="Manual Input")

        # Table import tab
        self.table_import_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.table_import_tab, text="Table Import")

        # Result display tab
        self.result_display_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.result_display_tab, text="Calculation Results")

        # Initialize each tab
        self.initialize_manual_input_interface()
        self.initialize_import_interface()
        self.initialize_result_interface()

        # Store calculation results (by site)
        self.calculation_results = {}
        # Store current recognized column name mapping
        self.current_column_mapping = {}

    def _create_default_site_data(self):
        """Create default site data structure (only keep necessary parameters)"""
        return {
            "acceleration_intensity_choice": "amax",
            "seismic_intensity": "7",
            "surface_peak_acceleration_amax(g)": 0.1,
            "discrimination_depth": "15",
            "groundwater_depth_dw": 2.0,
            "surface_wave_magnitude_Ms": 6.5,
            "design_earthquake_group": "1",
            "beta_calculation_method": "group",
            "particle_content_choice": "ρc",
            "soil_layer_data": [
                {"ds": 1.5, "N": 12, "content_value": 3.0},
                {"ds": 4.5, "N": 14, "content_value": 3.0},
                {"ds": 7.5, "N": 16, "content_value": 3.0},
                {"ds": 10.5, "N": 18, "content_value": 3.0},
                {"ds": 13.5, "N": 20, "content_value": 3.0}
            ]
        }

    def create_menu_bar(self):
        """Create menu bar and help menu"""
        menu_bar = tk.Menu(self.root)

        # File menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Export Calculation Results", command=self.export_calculation_results)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Help menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Software Introduction", command=self.show_software_introduction)
        help_menu.add_command(label="Parameter Description", command=self.show_parameter_description)
        help_menu.add_command(label="Liquefaction Classification Standards",
                              command=self.show_liquefaction_classification_standards)
        help_menu.add_separator()
        help_menu.add_command(label="About Software", command=self.show_about_software)

        menu_bar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menu_bar)

    def show_software_introduction(self):
        """Show software introduction (explain binary choice logic)"""
        intro_text = """
        Sand Soil Seismic Liquefaction Index Calculation and Classification Software Based on Standard Penetration Test (SPT)

        Core Optimization: All parameters are binary choice input, simplifying operation process
        1. Seismic Parameters: Surface Peak Acceleration amax / Seismic Intensity (choose one)
        2. Particle Content: Clay Content ρc / Fine Particle Content Fc (choose one)
        3. β Calculation: Design Earthquake Group / Surface Wave Magnitude Formula (choose one)

        Calculation Process Strictly Follows Document 5 Steps:
        1. Calculate Standard Penetration Critical Value Ncr (according to selected parameter type)
        2. Calculate Liquefaction Probability PL
        3. Calculate Liquefaction Critical Value Ncr,PL at Any Liquefaction Probability
        4. Calculate Liquefaction Index ILE
        5. Determine Liquefaction Grade (according to standard tables)

        Function Description:
        1. Multi-point Input: Support management and calculation of multiple site data
        2. Table Import: Automatically recognize multiple column name formats for binary choice parameters
        3. Result Display: Include liquefaction index, grade determination, liquefaction discrimination possibility and anti-liquefaction measures suggestions
        4. Result Export: Export all calculation results to Excel tables
        """
        self.show_info_dialog("Software Introduction", intro_text)

    def show_parameter_description(self):
        """Show parameter description (clarify binary choice logic)"""
        param_text = """
        Parameter Description (All parameters are binary choice, no need for duplicate input):

        1. Seismic Parameters (choose one):
        - Surface Peak Acceleration amax(g): 0.1/0.15/0.2/0.3/0.4g (direct input)
        - Seismic Intensity: 7/8/9 degrees (automatically corresponding to amax: 0.1/0.2/0.4g)

        2. Basic Parameters:
        - Discrimination Depth: 15m or 20m (only used for liquefaction grade determination)
        - Groundwater Depth dw(m): Site groundwater burial depth
        - Surface Wave Magnitude Ms: Only needed when β calculation chooses "formula"
        - Design Earthquake Group: 1/2/3 group (only needed when β calculation chooses "group")

        3. Particle Content Parameters (choose one):
        - Clay Content ρc(%): Clay content percentage, automatically takes 3 when <3%
        - Fine Particle Content Fc(%): Fine particle content percentage, automatically takes 15 when <15%

        4. Soil Layer Parameters:
        - Saturated Sand Depth ds(m): Depth of standard penetration test point in soil layer
        - Measured SPT Blow Count N: Blow count of standard penetration test

        5. β Calculation Method (choose one):
        - Design Earthquake Group: 1 group=0.80, 2 group=0.95, 3 group=1.05
        - Formula Calculation: β=0.20Ms - 0.50 (need to input surface wave magnitude Ms)

        6. Liquefaction Discrimination Possibility Description:
        - Based on liquefaction probability PL determination: PL≥0.5 (high possibility), 0.3≤PL<0.5 (medium possibility), PL<0.3 (low possibility)
        - Overall liquefaction possibility: Comprehensive judgment based on average PL of all soil layers
        """
        self.show_info_dialog("Parameter Description", param_text)

    def show_liquefaction_classification_standards(self):
        """Show liquefaction classification standards (document table content)"""
        standard_text = """
        《SPT-based Liquefaction Index Calculation Method》 Core Standards:

        I. Liquefaction Grade Discrimination Table
        +----------------+-------------------------------------+-------------------------------------+
        | Liquefaction Grade | Liquefaction Index when Discrimination Depth is 15m | Liquefaction Index when Discrimination Depth is 20m |
        +----------------+-------------------------------------+-------------------------------------+
        | Slight         | 0＜ILE≤5                           | 0＜ILE≤6                           |
        +----------------+-------------------------------------+-------------------------------------+
        | Moderate       | 5＜ILE≤15                          | 6＜ILE≤18                          |
        +----------------+-------------------------------------+-------------------------------------+
        | Severe         | ILE＞15                            | ILE＞18                            |
        +----------------+-------------------------------------+-------------------------------------+

        II. Anti-liquefaction Measures Table
        +-------------------------------+-----------------------------------------------+----------------------------------------------------+----------------------------------------------------------+
        | Building Seismic Fortification Class | Slight Liquefaction                           | Moderate Liquefaction                             | Severe Liquefaction                                      |
        +-------------------------------+-----------------------------------------------+----------------------------------------------------+----------------------------------------------------------+
        | Class B                         | Partially eliminate liquefaction settlement, or treat foundation and superstructure | Completely eliminate liquefaction settlement, or partially eliminate liquefaction settlement and treat foundation and superstructure | Completely eliminate liquefaction settlement             |
        +-------------------------------+-----------------------------------------------+----------------------------------------------------+----------------------------------------------------------+
        | Class C                         | Treat foundation and superstructure, or may not treat | Treat foundation and superstructure, or take stricter measures | Completely eliminate liquefaction settlement, or partially eliminate liquefaction settlement and treat foundation and superstructure |
        +-------------------------------+-----------------------------------------------+----------------------------------------------------+----------------------------------------------------------+
        | Class D                         | No measures needed                             | No measures needed                                 | Treat foundation and superstructure, or take other economical and reasonable measures |
        +-------------------------------+-----------------------------------------------+----------------------------------------------------+----------------------------------------------------------+

        III. Liquefaction Discrimination Possibility Determination Standards
        +----------------+-------------------+-----------------------------------+
        | Liquefaction Probability PL | Discrimination Possibility | Description                        |
        +----------------+-------------------+-----------------------------------+
        | PL ≥ 0.5       | High Possibility  | High probability of liquefaction in this soil layer under earthquake |
        +----------------+-------------------+-----------------------------------+
        | 0.3 ≤ PL ＜ 0.5 | Medium Possibility | Possible liquefaction in this soil layer under earthquake |
        +----------------+-------------------+-----------------------------------+
        | PL ＜ 0.3       | Low Possibility   | Low probability of liquefaction in this soil layer under earthquake |
        +----------------+-------------------+-----------------------------------+
        """
        self.show_info_dialog("Liquefaction Classification Standards", standard_text)

    def show_about_software(self):
        """Show about software information"""
        messagebox.showinfo("About Software",
                            "Sand Soil Seismic Liquefaction Index Calculation and Classification Software Based on Standard Penetration Test (SPT)\nOptimization Features: Binary choice parameter input, simplified operation process, added liquefaction discrimination possibility display, supports result export")

    def show_info_dialog(self, title, content):
        """Show information dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("800x600")
        dialog.resizable(True, True)

        text_widget = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=("Courier New", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=10)

    def initialize_manual_input_interface(self):
        """Initialize manual input tab (implement binary choice input logic)"""
        # Site management frame
        site_frame = ttk.LabelFrame(self.manual_input_tab, text="Site Management")
        site_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(site_frame, text="Current Site:").grid(row=0, column=0, padx=10, pady=5)
        self.site_combo = ttk.Combobox(site_frame, state="readonly", width=10)
        self.site_combo.grid(row=0, column=1, padx=5, pady=5)
        self.site_combo.bind("<<ComboboxSelected>>", self._site_switch_event)

        ttk.Button(site_frame, text="Add Site", command=self._add_site).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(site_frame, text="Delete Site", command=self._delete_site).grid(row=0, column=3, padx=5, pady=5)

        # Binary choice parameter selection frame
        choice_frame = ttk.LabelFrame(self.manual_input_tab, text="Parameter Input Selection (Binary Choice)")
        choice_frame.pack(fill=tk.X, padx=10, pady=5)

        # Seismic parameter choice (amax/intensity)
        ttk.Label(choice_frame, text="Seismic Parameters:").grid(row=0, column=0, padx=10, pady=5)
        tk.Radiobutton(choice_frame, text="Surface Peak Acceleration amax",
                       variable=self.manual_acceleration_intensity_choice,
                       value="amax", command=self._switch_seismic_parameter_input).grid(row=0, column=1, padx=5, pady=5)
        tk.Radiobutton(choice_frame, text="Seismic Intensity", variable=self.manual_acceleration_intensity_choice,
                       value="intensity", command=self._switch_seismic_parameter_input).grid(row=0, column=2, padx=5,
                                                                                             pady=5)

        # Particle content choice (ρc/Fc)
        ttk.Label(choice_frame, text="Particle Content:").grid(row=0, column=3, padx=10, pady=5)
        tk.Radiobutton(choice_frame, text="Clay Content ρc", variable=self.manual_particle_content_choice,
                       value="ρc", command=self._switch_particle_content_input).grid(row=0, column=4, padx=5, pady=5)
        tk.Radiobutton(choice_frame, text="Fine Particle Content Fc", variable=self.manual_particle_content_choice,
                       value="Fc", command=self._switch_particle_content_input).grid(row=0, column=5, padx=5, pady=5)

        # β calculation method choice
        ttk.Label(choice_frame, text="β Calculation Method:").grid(row=0, column=6, padx=10, pady=5)
        tk.Radiobutton(choice_frame, text="Design Earthquake Group", variable=self.manual_beta_calculation_choice,
                       value="group", command=self._switch_beta_calculation_method).grid(row=0, column=7, padx=5,
                                                                                         pady=5)
        tk.Radiobutton(choice_frame, text="Formula Calculation (Ms)", variable=self.manual_beta_calculation_choice,
                       value="formula", command=self._switch_beta_calculation_method).grid(row=0, column=8, padx=5,
                                                                                           pady=5)

        # Basic parameters frame
        basic_param_frame = ttk.LabelFrame(self.manual_input_tab, text="Basic Parameters")
        basic_param_frame.pack(fill=tk.X, padx=10, pady=5)
        basic_param_frame.grid_columnconfigure(5, weight=1)

        # Seismic parameter input (dynamically displayed based on choice)
        self.intensity_var = tk.StringVar(value="7")
        self.amax_var = tk.DoubleVar(value=0.1)
        ttk.Label(basic_param_frame, text="Seismic Intensity:").grid(row=0, column=0, padx=10, pady=5)
        self.intensity_entry = ttk.Combobox(basic_param_frame, textvariable=self.intensity_var, values=["7", "8", "9"],
                                            width=10)
        self.intensity_entry.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(basic_param_frame, text="Surface Peak Acceleration amax(g):").grid(row=0, column=2, padx=10, pady=5)
        self.amax_entry = ttk.Entry(basic_param_frame, textvariable=self.amax_var, width=12)
        self.amax_entry.grid(row=0, column=3, padx=10, pady=5)

        # Discrimination depth choice
        ttk.Label(basic_param_frame, text="Discrimination Depth(m):").grid(row=0, column=4, padx=10, pady=5)
        self.depth_var = tk.StringVar(value="15")
        ttk.Combobox(basic_param_frame, textvariable=self.depth_var, values=["15", "20"], width=10).grid(row=0,
                                                                                                         column=5,
                                                                                                         padx=10,
                                                                                                         pady=5)

        # Groundwater depth
        ttk.Label(basic_param_frame, text="Groundwater Depth dw(m):").grid(row=1, column=0, padx=10, pady=5)
        self.groundwater_var = tk.DoubleVar(value=2.0)
        ttk.Entry(basic_param_frame, textvariable=self.groundwater_var, width=12).grid(row=1, column=1, padx=10, pady=5)

        # β related parameters (dynamically displayed)
        self.magnitude_var = tk.DoubleVar(value=6.5)
        self.earthquake_group_var = tk.StringVar(value="1")
        ttk.Label(basic_param_frame, text="Surface Wave Magnitude Ms:").grid(row=1, column=2, padx=10, pady=5)
        self.magnitude_entry = ttk.Entry(basic_param_frame, textvariable=self.magnitude_var, width=12)
        self.magnitude_entry.grid(row=1, column=3, padx=10, pady=5)

        ttk.Label(basic_param_frame, text="Design Earthquake Group:").grid(row=1, column=4, padx=10, pady=5)
        self.group_entry = ttk.Combobox(basic_param_frame, textvariable=self.earthquake_group_var,
                                        values=["1", "2", "3"], width=10)
        self.group_entry.grid(row=1, column=5, padx=10, pady=5)

        # Soil layer parameters table (only keep necessary parameters)
        self.soil_layer_frame = ttk.LabelFrame(self.manual_input_tab, text="Soil Layer Parameters")
        self.soil_layer_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Table header (dynamically display particle content type)
        self.particle_content_header = ttk.Label(self.soil_layer_frame, text="Clay Content ρc(%)",
                                                 font=("Arial", 10, "bold"))
        self.particle_content_header.grid(row=0, column=2, padx=15, pady=5)
        ttk.Label(self.soil_layer_frame, text="Saturated Sand Depth ds(m)", font=("Arial", 10, "bold")).grid(row=0,
                                                                                                             column=0,
                                                                                                             padx=15,
                                                                                                             pady=5)
        ttk.Label(self.soil_layer_frame, text="Measured SPT Blow Count N", font=("Arial", 10, "bold")).grid(row=0,
                                                                                                            column=1,
                                                                                                            padx=15,
                                                                                                            pady=5)

        # Soil layer input rows (store Entry widgets)
        self.soil_layer_entries = []
        self._refresh_soil_layer_entries()

        # Button frame
        button_frame = ttk.Frame(self.manual_input_tab)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(button_frame, text="Add Soil Layer", command=self.add_soil_layer).pack(side=tk.LEFT, padx=8, pady=5)
        ttk.Button(button_frame, text="Delete Last Layer", command=self.delete_last_layer).pack(side=tk.LEFT, padx=8,
                                                                                                pady=5)
        ttk.Button(button_frame, text="Clear Input", command=self.clear_manual_input).pack(side=tk.RIGHT, padx=8,
                                                                                           pady=5)
        ttk.Button(button_frame, text="Calculate", command=self.manual_calculation).pack(side=tk.RIGHT, padx=8, pady=5)

        # Initialize display state
        self._switch_seismic_parameter_input()
        self._switch_particle_content_input()
        self._switch_beta_calculation_method()
        self._refresh_site_combo()

    def _switch_seismic_parameter_input(self):
        """Switch seismic parameter input (amax/intensity)"""
        if self.manual_acceleration_intensity_choice.get() == "amax":
            self.intensity_entry.config(state="disabled")
            self.amax_entry.config(state="normal")
        else:
            self.intensity_entry.config(state="normal")
            self.amax_entry.config(state="disabled")

    def _switch_particle_content_input(self):
        """Switch particle content input (ρc/Fc)"""
        choice_type = self.manual_particle_content_choice.get()
        if choice_type == "ρc":
            self.particle_content_header.config(text="Clay Content ρc(%)")
        else:
            self.particle_content_header.config(text="Fine Particle Content Fc(%)")
        # Refresh soil layer entry placeholders
        self._refresh_soil_layer_entries(update_placeholder_only=True)

    def _switch_beta_calculation_method(self):
        """Switch β calculation method (group/formula)"""
        if self.manual_beta_calculation_choice.get() == "group":
            self.group_entry.config(state="normal")
            self.magnitude_entry.config(state="disabled")
        else:
            self.group_entry.config(state="disabled")
            self.magnitude_entry.config(state="normal")

    def _refresh_site_combo(self):
        """Refresh site selection combo box"""
        site_number_list = [f"Site {i + 1}" for i in range(len(self.site_data))]
        self.site_combo['values'] = site_number_list
        if site_number_list:
            self.site_combo.current(self.current_site_index)

    def _add_site(self):
        """Add new site"""
        new_site = self._create_default_site_data()
        self.site_data.append(new_site)
        self.current_site_index = len(self.site_data) - 1
        self._refresh_site_combo()
        self._site_switch_event(None)

    def _delete_site(self):
        """Delete current site"""
        if len(self.site_data) <= 1:
            messagebox.showwarning("Warning", "At least one site must be kept")
            return
        del self.site_data[self.current_site_index]
        self.current_site_index = max(0, self.current_site_index - 1)
        self._refresh_site_combo()
        self._site_switch_event(None)

    def _site_switch_event(self, event):
        """Update interface when switching sites"""
        self.current_site_index = self.site_combo.current()
        current_site = self.site_data[self.current_site_index]
        # Update basic parameters
        self.manual_acceleration_intensity_choice.set(current_site["acceleration_intensity_choice"])
        self.manual_particle_content_choice.set(current_site["particle_content_choice"])
        self.manual_beta_calculation_choice.set(current_site["beta_calculation_method"])
        self.intensity_var.set(current_site["seismic_intensity"])
        self.amax_var.set(current_site["surface_peak_acceleration_amax(g)"])
        self.depth_var.set(current_site["discrimination_depth"])
        self.groundwater_var.set(current_site["groundwater_depth_dw"])
        self.magnitude_var.set(current_site["surface_wave_magnitude_Ms"])
        self.earthquake_group_var.set(current_site["design_earthquake_group"])
        # Refresh display state
        self._switch_seismic_parameter_input()
        self._switch_particle_content_input()
        self._switch_beta_calculation_method()
        # Update soil layer input
        self._refresh_soil_layer_entries()

    def _refresh_soil_layer_entries(self, update_placeholder_only=False):
        """Refresh soil layer input entries (only keep necessary parameters)"""
        if not update_placeholder_only:
            # Clear existing entries
            for row_entries in self.soil_layer_entries:
                for entry in row_entries:
                    entry.destroy()
            self.soil_layer_entries = []

        # Get current particle content type
        content_type = self.manual_particle_content_choice.get()
        current_site = self.site_data[self.current_site_index]

        # Create new entries
        for i, soil_layer in enumerate(current_site["soil_layer_data"]):
            if update_placeholder_only and i < len(self.soil_layer_entries):
                # Only update placeholder text
                row_entries = self.soil_layer_entries[i]
                placeholder_text = "3.0" if content_type == "ρc" else "15.0"
                row_entries[2].delete(0, tk.END)
                row_entries[2].insert(0, placeholder_text)
                continue

            row_entries = []
            # Saturated sand depth ds
            ds_entry = ttk.Entry(self.soil_layer_frame, width=15)
            ds_entry.grid(row=i + 1, column=0, padx=15, pady=3)
            ds_entry.insert(0, f"{soil_layer['ds']}")
            row_entries.append(ds_entry)
            # Measured SPT blow count N
            N_entry = ttk.Entry(self.soil_layer_frame, width=15)
            N_entry.grid(row=i + 1, column=1, padx=15, pady=3)
            N_entry.insert(0, f"{soil_layer['N']}")
            row_entries.append(N_entry)
            # Particle content (ρc/Fc)
            content_entry = ttk.Entry(self.soil_layer_frame, width=15)
            content_entry.grid(row=i + 1, column=2, padx=15, pady=3)
            placeholder_text = "3.0" if content_type == "ρc" else "15.0"
            content_entry.insert(0,
                                 f"{soil_layer['content_value']}" if soil_layer['content_value'] else placeholder_text)
            row_entries.append(content_entry)
            self.soil_layer_entries.append(row_entries)

    def add_soil_layer(self):
        """Add soil layer input row"""
        current_site = self.site_data[self.current_site_index]
        # Calculate default ds (last layer ds + 3)
        last_ds = current_site["soil_layer_data"][-1]["ds"] if current_site["soil_layer_data"] else 1.5
        new_ds = last_ds + 3.0
        new_N = 12 + len(current_site["soil_layer_data"]) * 2
        # Set default value based on current particle content type choice
        content_type = self.manual_particle_content_choice.get()
        default_content = 3.0 if content_type == "ρc" else 15.0
        new_soil_layer = {"ds": new_ds, "N": new_N, "content_value": default_content}
        current_site["soil_layer_data"].append(new_soil_layer)
        self._refresh_soil_layer_entries()

    def delete_last_layer(self):
        """Delete last soil layer"""
        current_site = self.site_data[self.current_site_index]
        if len(current_site["soil_layer_data"]) <= 1:
            messagebox.showwarning("Warning", "At least one soil layer must be kept")
            return
        current_site["soil_layer_data"].pop()
        self._refresh_soil_layer_entries()

    def clear_manual_input(self):
        """Clear current site manual input"""
        self.manual_acceleration_intensity_choice.set("amax")
        self.manual_particle_content_choice.set("ρc")
        self.manual_beta_calculation_choice.set("group")
        self.intensity_var.set("7")
        self.amax_var.set(0.1)
        self.depth_var.set("15")
        self.groundwater_var.set(2.0)
        self.magnitude_var.set(6.5)
        self.earthquake_group_var.set("1")
        # Reset to default 5 soil layers
        current_site = self.site_data[self.current_site_index]
        current_site["soil_layer_data"] = [
            {"ds": 1.5, "N": 12, "content_value": 3.0},
            {"ds": 4.5, "N": 14, "content_value": 3.0},
            {"ds": 7.5, "N": 16, "content_value": 3.0},
            {"ds": 10.5, "N": 18, "content_value": 3.0},
            {"ds": 13.5, "N": 20, "content_value": 3.0}
        ]
        # Refresh display state
        self._switch_seismic_parameter_input()
        self._switch_particle_content_input()
        self._switch_beta_calculation_method()
        self._refresh_soil_layer_entries()

    def initialize_import_interface(self):
        """Initialize table import tab (support binary choice parameter recognition)"""
        # Main container uses grid layout
        self.table_import_tab.grid_columnconfigure(0, weight=3)
        self.table_import_tab.grid_columnconfigure(1, weight=2)
        self.table_import_tab.grid_rowconfigure(0, weight=1)

        # Left area: File operations and data preview
        left_frame = ttk.Frame(self.table_import_tab)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)
        left_frame.grid_rowconfigure(3, weight=1)

        # Left title
        ttk.Label(left_frame, text="Table Data Import", font=("Arial", 12, "bold")).pack(pady=10)
        ttk.Label(left_frame, text="Support Excel format, automatically recognize binary choice parameter column names",
                  font=("Arial", 10)).pack(pady=5)

        # File selection area
        file_frame = ttk.LabelFrame(left_frame, text="File Selection")
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(file_frame, text="Browse File", command=self.browse_file).pack(side=tk.LEFT, padx=8, pady=5)
        ttk.Button(file_frame, text="Clear Import", command=self.clear_import).pack(side=tk.LEFT, padx=8, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=30).pack(side=tk.LEFT, padx=10, fill=tk.X,
                                                                              expand=True,
                                                                              pady=5)

        # Column recognition result display
        recognition_frame = ttk.LabelFrame(left_frame,
                                           text="Column Recognition Results (Auto-match Binary Choice Parameters)")
        recognition_frame.pack(fill=tk.X, padx=5, pady=5)

        self.recognition_labels = {}
        for i, standard_col_name in enumerate(self.column_name_mapping.keys()):
            ttk.Label(recognition_frame, text=f"{standard_col_name}:").grid(row=i, column=0, padx=10, pady=3,
                                                                            sticky=tk.W)
            label = ttk.Label(recognition_frame, text="Not Recognized")
            label.grid(row=i, column=1, padx=5, pady=3, sticky=tk.W)
            self.recognition_labels[standard_col_name] = label
            recognition_frame.grid_columnconfigure(1, weight=1)

        # Table preview area
        preview_frame = ttk.LabelFrame(left_frame, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.preview_canvas = tk.Canvas(preview_frame)
        preview_scrollbar_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        preview_scrollbar_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        self.preview_inner_frame = ttk.Frame(self.preview_canvas)

        self.preview_inner_frame.bind(
            "<Configure>",
            lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        )

        self.preview_canvas.create_window((0, 0), window=self.preview_inner_frame, anchor=tk.NW)
        self.preview_canvas.configure(yscrollcommand=preview_scrollbar_y.set, xscrollcommand=preview_scrollbar_x.set)

        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        preview_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Right area: Parameter settings and data statistics
        right_frame = ttk.Frame(self.table_import_tab)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)
        right_frame.grid_rowconfigure(1, weight=1)

        # Right title
        ttk.Label(right_frame, text="Calculation Parameter Settings (Binary Choice)", font=("Arial", 12, "bold")).pack(
            pady=10)

        # Import parameter settings (right upper part)
        param_frame = ttk.LabelFrame(right_frame, text="Parameter Selection")
        param_frame.pack(fill=tk.X, padx=5, pady=5)
        param_frame.grid_columnconfigure(1, weight=1)

        # Binary choice parameter selection
        row_index = 0
        # Seismic parameter choice
        ttk.Label(param_frame, text="Seismic Parameters:").grid(row=row_index, column=0, padx=5, pady=8, sticky=tk.W)
        self.import_acceleration_intensity_choice = tk.StringVar(value="amax")
        tk.Radiobutton(param_frame, text="amax (table data)", variable=self.import_acceleration_intensity_choice,
                       value="amax").grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        tk.Radiobutton(param_frame, text="Intensity (table data)", variable=self.import_acceleration_intensity_choice,
                       value="intensity").grid(
            row=row_index, column=2, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        # Particle content choice
        ttk.Label(param_frame, text="Particle Content:").grid(row=row_index, column=0, padx=5, pady=8, sticky=tk.W)
        self.import_particle_content_choice = tk.StringVar(value="ρc")
        tk.Radiobutton(param_frame, text="Clay Content ρc", variable=self.import_particle_content_choice,
                       value="ρc").grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        tk.Radiobutton(param_frame, text="Fine Particle Content Fc", variable=self.import_particle_content_choice,
                       value="Fc").grid(
            row=row_index, column=2, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        # β calculation method choice
        ttk.Label(param_frame, text="β Calculation Method:").grid(row=row_index, column=0, padx=5, pady=8, sticky=tk.W)
        self.import_beta_calculation_choice = tk.StringVar(value="group")
        tk.Radiobutton(param_frame, text="Design Earthquake Group", variable=self.import_beta_calculation_choice,
                       value="group").grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        tk.Radiobutton(param_frame, text="Formula Calculation (Ms)", variable=self.import_beta_calculation_choice,
                       value="formula").grid(
            row=row_index, column=2, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        # Supplementary parameter input
        ttk.Label(param_frame, text="Discrimination Depth(m):").grid(row=row_index, column=0, padx=5, pady=8,
                                                                     sticky=tk.W)
        self.import_depth_var = tk.StringVar(value="15")
        ttk.Combobox(param_frame, textvariable=self.import_depth_var, values=["15", "20"], width=10).grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        ttk.Label(param_frame, text="Groundwater Depth dw(m):").grid(row=row_index, column=0, padx=5, pady=8,
                                                                     sticky=tk.W)
        self.import_groundwater_var = tk.DoubleVar(value=2.0)
        ttk.Entry(param_frame, textvariable=self.import_groundwater_var, width=12).grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        ttk.Label(param_frame, text="Surface Wave Magnitude Ms:").grid(row=row_index, column=0, padx=5, pady=8,
                                                                       sticky=tk.W)
        self.import_magnitude_var = tk.DoubleVar(value=6.5)
        ttk.Entry(param_frame, textvariable=self.import_magnitude_var, width=12).grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        ttk.Label(param_frame, text="Design Earthquake Group:").grid(row=row_index, column=0, padx=5, pady=8,
                                                                     sticky=tk.W)
        self.import_earthquake_group_var = tk.StringVar(value="1")
        ttk.Combobox(param_frame, textvariable=self.import_earthquake_group_var, values=["1", "2", "3"], width=10).grid(
            row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        # Site selection (multi-point import)
        ttk.Label(param_frame, text="Select Calculation Site:").grid(row=row_index, column=0, padx=5, pady=8,
                                                                     sticky=tk.W)
        self.import_site_combo = ttk.Combobox(param_frame, state="readonly", width=10)
        self.import_site_combo.grid(row=row_index, column=1, padx=5, pady=8, sticky=tk.W)
        row_index += 1

        # Data statistics information (right middle part)
        self.statistics_frame = ttk.LabelFrame(right_frame, text="Data Statistics")
        self.statistics_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.statistics_text = scrolledtext.ScrolledText(self.statistics_frame, wrap=tk.WORD, height=10)
        self.statistics_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.statistics_text.insert(tk.END, "Please import data to display statistics...")
        self.statistics_text.config(state=tk.DISABLED)

        # Import tips (right lower part)
        tips_frame = ttk.LabelFrame(right_frame, text="Import Tips")
        tips_frame.pack(fill=tk.X, padx=5, pady=5)

        tips_text = """
1. Table must contain: Site Number, Saturated Sand Depth,
   Measured SPT Blow Count, Clay/Fine Particle Content (choose one),
   amax/Intensity (choose one)
2. Same site numbers belong to same measurement point, automatically grouped for calculation
3. Clay content <3% automatically takes 3%, fine particle content <15% automatically takes 15%
4. After import, please confirm column recognition results are correct
5. Calculation results include each layer and overall liquefaction discrimination possibility (based on PL value)
        """
        ttk.Label(tips_frame, text=tips_text, justify=tk.LEFT).pack(fill=tk.X, padx=5, pady=5)

        # Calculation buttons (fixed at bottom of import tab)
        button_frame = ttk.Frame(self.table_import_tab)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Button(button_frame, text="Batch Calculate All Sites",
                   command=self.batch_calculate_all_imported_sites).pack(
            side=tk.LEFT, padx=8, pady=8, ipady=2)
        ttk.Button(button_frame, text="Calculate Selected Site", command=self.import_calculation).pack(
            side=tk.LEFT, padx=8, pady=8, ipady=2)

        # Store imported multi-point data
        self.imported_site_data = {}

    def batch_calculate_all_imported_sites(self):
        """Batch calculate all imported sites"""
        # Check if column recognition is complete
        core_columns = ["Site Number", "Saturated Sand Depth ds(m)", "Measured SPT Blow Count N",
                        "Clay/Fine Particle Content(%)"]
        missing_columns = [col_name for col_name in core_columns if not self.current_column_mapping.get(col_name)]
        if missing_columns:
            messagebox.showwarning("Warning",
                                   f"The following necessary columns are not recognized, please check table:\n{', '.join(missing_columns)}")
            return

        # Check if seismic parameter column is recognized
        seismic_param_column = self.current_column_mapping.get(
            "Surface Peak Acceleration amax(g)") or self.current_column_mapping.get("Seismic Intensity")
        if not seismic_param_column:
            messagebox.showwarning("Warning",
                                   "Seismic parameter (amax or intensity) not recognized, please supplement data")
            return

        if not self.imported_site_data:
            messagebox.showwarning("Warning", "Please import data first")
            return

        try:
            discrimination_depth = int(self.import_depth_var.get())
            groundwater_depth = self.import_groundwater_var.get()
            magnitude = self.import_magnitude_var.get()
            earthquake_group = self.import_earthquake_group_var.get()
            acceleration_intensity_choice = self.import_acceleration_intensity_choice.get()
            particle_content_choice = self.import_particle_content_choice.get()
            beta_calculation_method = self.import_beta_calculation_choice.get()

            # Show progress window
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Calculating")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()

            ttk.Label(progress_window, text="Calculating all sites, please wait...").pack(pady=20)
            self.root.update_idletasks()

            # Calculate each site
            total_site_count = len(self.imported_site_data)
            for i, (site_number, site_data) in enumerate(self.imported_site_data.items()):
                progress_window.title(f"Calculating ({i + 1}/{total_site_count})")

                # Construct soil layer data
                soil_layer_list = []
                content_column = self.current_column_mapping["Clay/Fine Particle Content(%)"]
                seismic_param_column = self.current_column_mapping.get(
                    "Surface Peak Acceleration amax(g)") if acceleration_intensity_choice == "amax" else self.current_column_mapping.get(
                    "Seismic Intensity")

                for _, row in site_data.iterrows():
                    # Process particle content
                    content_value = row[content_column]
                    if particle_content_choice == "ρc":
                        content_value = max(3.0, float(content_value)) if pd.notna(content_value) else 3.0
                    else:
                        content_value = max(15.0, float(content_value)) if pd.notna(content_value) else 15.0

                    # Process seismic parameter
                    seismic_param_value = row[seismic_param_column] if pd.notna(row[seismic_param_column]) else 0.1

                    soil_layer_list.append({
                        "ds": row[self.current_column_mapping["Saturated Sand Depth ds(m)"]],
                        "N": row[self.current_column_mapping["Measured SPT Blow Count N"]],
                        "content_value": content_value,
                        "seismic_param_value": seismic_param_value
                    })

                # Construct calculation data
                calculation_data = {
                    "acceleration_intensity_choice": acceleration_intensity_choice,
                    "seismic_intensity": 7 if acceleration_intensity_choice != "intensity" else int(
                        seismic_param_value),
                    "surface_peak_acceleration_amax(g)": 0.1 if acceleration_intensity_choice != "amax" else float(
                        seismic_param_value),
                    "discrimination_depth": discrimination_depth,
                    "groundwater_depth_dw": groundwater_depth,
                    "surface_wave_magnitude_Ms": magnitude,
                    "design_earthquake_group": earthquake_group,
                    "beta_calculation_method": beta_calculation_method,
                    "particle_content_choice": particle_content_choice,
                    "soil_layer_data": soil_layer_list
                }

                # Calculate and store result
                result = self.calculate_liquefaction_index(calculation_data)
                self.calculation_results[f"Imported Site {site_number}"] = result
                self.root.update_idletasks()

            # Close progress window
            progress_window.destroy()

            # Update result display
            self._refresh_result_site_combo()
            if self.calculation_results:
                self.result_site_combo.current(0)
                self._result_site_switch_event(None)

            self.notebook.select(self.result_display_tab)
            messagebox.showinfo("Complete",
                                f"Successfully calculated {total_site_count} sites, results include liquefaction discrimination possibility")

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def _standardize_text(self, text):
        """Standardize text for matching: remove special characters, spaces and convert to lowercase"""
        if not text:
            return ""
        standardized_text = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', text)
        return standardized_text.lower()

    def _auto_recognize_column_names(self, column_name_list):
        """Automatically recognize table column names (support binary choice parameter matching)"""
        standardized_columns = {self._standardize_text(col_name): col_name for col_name in column_name_list}
        mapping_result = {}

        for standard_col_name, alias_list in self.column_name_mapping.items():
            best_match = None
            standardized_standard_col = self._standardize_text(standard_col_name)

            # 1. Priority exact match
            if standardized_standard_col in standardized_columns:
                best_match = standardized_columns[standardized_standard_col]
            else:
                # 2. Match aliases
                for alias in alias_list:
                    standardized_alias = self._standardize_text(alias)
                    if standardized_alias in standardized_columns:
                        best_match = standardized_columns[standardized_alias]
                        break
                # 3. Fuzzy match
                if not best_match:
                    for standardized_col, original_col_name in standardized_columns.items():
                        if standardized_standard_col in standardized_col or any(
                                self._standardize_text(alias) in standardized_col for alias in alias_list):
                            best_match = original_col_name
                            break
            mapping_result[standard_col_name] = best_match

        return mapping_result

    def _update_data_statistics(self, dataframe):
        """Update data statistics information"""
        self.statistics_text.config(state=tk.NORMAL)
        self.statistics_text.delete(1.0, tk.END)

        # Basic statistics
        self.statistics_text.insert(tk.END, f"Total Records: {len(dataframe)}\n")

        # Site count statistics
        site_column = self.current_column_mapping.get("Site Number")
        if site_column and site_column in dataframe.columns:
            site_count = dataframe[site_column].nunique()
            self.statistics_text.insert(tk.END, f"Site Count: {site_count}\n\n")

            site_record_counts = dataframe[site_column].value_counts().to_dict()
            for site, count in sorted(site_record_counts.items()):
                self.statistics_text.insert(tk.END, f"  Site {site}: {count} records\n")

        # N value statistics
        N_column = self.current_column_mapping.get("Measured SPT Blow Count N")
        if N_column and N_column in dataframe.columns:
            self.statistics_text.insert(tk.END, "\nMeasured SPT Blow Count N Statistics:\n")
            self.statistics_text.insert(tk.END, f"  Minimum: {dataframe[N_column].min()}\n")
            self.statistics_text.insert(tk.END, f"  Maximum: {dataframe[N_column].max()}\n")
            self.statistics_text.insert(tk.END, f"  Average: {dataframe[N_column].mean():.2f}\n")

        # Particle content statistics
        content_column = self.current_column_mapping.get("Clay/Fine Particle Content(%)")
        if content_column and content_column in dataframe.columns:
            self.statistics_text.insert(tk.END, "\nParticle Content Statistics:\n")
            self.statistics_text.insert(tk.END, f"  Minimum: {dataframe[content_column].min():.1f}%\n")
            self.statistics_text.insert(tk.END, f"  Maximum: {dataframe[content_column].max():.1f}%\n")
            self.statistics_text.insert(tk.END, f"  Average: {dataframe[content_column].mean():.1f}%\n")

        # Seismic parameter statistics
        amax_column = self.current_column_mapping.get("Surface Peak Acceleration amax(g)")
        intensity_column = self.current_column_mapping.get("Seismic Intensity")
        if amax_column and amax_column in dataframe.columns:
            self.statistics_text.insert(tk.END, "\namax Statistics:\n")
            self.statistics_text.insert(tk.END, f"  Minimum: {dataframe[amax_column].min():.2f}g\n")
            self.statistics_text.insert(tk.END, f"  Maximum: {dataframe[amax_column].max():.2f}g\n")
        elif intensity_column and intensity_column in dataframe.columns:
            self.statistics_text.insert(tk.END, "\nIntensity Statistics:\n")
            self.statistics_text.insert(tk.END, f"  Value Range: {sorted(dataframe[intensity_column].unique())}\n")

        self.statistics_text.config(state=tk.DISABLED)

    def browse_file(self):
        """Browse and preview Excel file (support binary choice parameters)"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.file_path_var.set(file_path)
            self.preview_file(file_path)

    def preview_file(self, file_path):
        """Preview imported file data (automatically recognize binary choice parameters)"""
        # Clear previous preview
        for widget in self.preview_inner_frame.winfo_children():
            widget.destroy()
        self.imported_site_data = {}
        self.import_site_combo['values'] = []
        self.current_column_mapping = {}

        try:
            dataframe = pd.read_excel(file_path)
            column_name_list = dataframe.columns.tolist()

            # Automatically recognize column names
            self.current_column_mapping = self._auto_recognize_column_names(column_name_list)

            # Update recognition result display
            for standard_col_name, label in self.recognition_labels.items():
                label.config(text=self.current_column_mapping[standard_col_name] if self.current_column_mapping[
                    standard_col_name] else "Not Recognized")

            # Update data statistics
            self._update_data_statistics(dataframe)

            # Check missing necessary columns
            core_columns = ["Site Number", "Saturated Sand Depth ds(m)", "Measured SPT Blow Count N",
                            "Clay/Fine Particle Content(%)"]
            missing_columns = [col_name for col_name in core_columns if not self.current_column_mapping.get(col_name)]
            if missing_columns:
                messagebox.showwarning("Recognition Prompt",
                                       f"The following necessary columns are not recognized, please check table:\n{', '.join(missing_columns)}")
                return

            # Check seismic parameter column
            seismic_param_column = self.current_column_mapping.get(
                "Surface Peak Acceleration amax(g)") or self.current_column_mapping.get("Seismic Intensity")
            if not seismic_param_column:
                messagebox.showwarning("Recognition Prompt",
                                       "Seismic parameter column (amax or intensity) not recognized, please supplement data")
                return

            # Group by site number
            site_column = self.current_column_mapping["Site Number"]
            self.imported_site_data = {
                str(site_number): group.reset_index(drop=True)
                for site_number, group in dataframe.groupby(site_column)
            }

            # Update site selection combo box
            self.import_site_combo['values'] = list(self.imported_site_data.keys())
            if self.imported_site_data:
                self.import_site_combo.current(0)

            # Display data preview (maximum 20 rows)
            display_rows = min(20, len(dataframe))
            for j, column_name in enumerate(column_name_list):
                ttk.Label(self.preview_inner_frame, text=column_name, font=("Arial", 10, "bold")).grid(row=0, column=j,
                                                                                                       padx=15, pady=5)
            for i in range(display_rows):
                row_data = dataframe.iloc[i]
                for j, value in enumerate(row_data):
                    ttk.Label(self.preview_inner_frame, text=str(value)).grid(row=i + 1, column=j, padx=15, pady=2)

        except Exception as e:
            messagebox.showerror("Error", f"File reading failed: {str(e)}")

    def clear_import(self):
        """Clear imported data"""
        self.file_path_var.set("")
        for widget in self.preview_inner_frame.winfo_children():
            widget.destroy()
        self.imported_site_data = {}
        self.import_site_combo['values'] = []
        # Reset parameters
        self.import_acceleration_intensity_choice.set("amax")
        self.import_particle_content_choice.set("ρc")
        self.import_beta_calculation_choice.set("group")
        self.import_depth_var.set("15")
        self.import_groundwater_var.set(2.0)
        self.import_magnitude_var.set(6.5)
        self.import_earthquake_group_var.set("1")
        # Reset column recognition
        for label in self.recognition_labels.values():
            label.config(text="Not Recognized")
        self.current_column_mapping = {}
        # Reset statistics
        self.statistics_text.config(state=tk.NORMAL)
        self.statistics_text.delete(1.0, tk.END)
        self.statistics_text.insert(tk.END, "Please import data to display statistics...")
        self.statistics_text.config(state=tk.DISABLED)

    def initialize_result_interface(self):
        """Initialize result interface (display binary choice parameter selection results + liquefaction discrimination possibility)"""
        # Main container with paned window for flexible resizing
        main_paned = ttk.PanedWindow(self.result_display_tab, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Top section: Results overview
        top_section = ttk.Frame(main_paned)
        main_paned.add(top_section, weight=1)

        # Bottom section: Details and chart
        bottom_section = ttk.Frame(main_paned)
        main_paned.add(bottom_section, weight=2)

        # Top section - split into left and right
        top_paned = ttk.PanedWindow(top_section, orient=tk.HORIZONTAL)
        top_paned.pack(fill=tk.BOTH, expand=True)

        # Left: Core results
        left_top_frame = ttk.LabelFrame(top_paned,
                                        text="Core Results (Including Liquefaction Discrimination Possibility)")
        top_paned.add(left_top_frame, weight=2)

        # Right: Standards and measures
        right_top_frame = ttk.LabelFrame(top_paned, text="Standards and Anti-liquefaction Measures")
        top_paned.add(right_top_frame, weight=1)

        # Bottom section - split into left and right
        bottom_paned = ttk.PanedWindow(bottom_section, orient=tk.HORIZONTAL)
        bottom_paned.pack(fill=tk.BOTH, expand=True)

        # Left: Calculation details
        left_bottom_frame = ttk.LabelFrame(bottom_paned, text="Each Soil Layer Calculation Details")
        bottom_paned.add(left_bottom_frame, weight=2)

        # Right: Chart
        right_bottom_frame = ttk.LabelFrame(bottom_paned, text="Liquefaction Probability Distribution")
        bottom_paned.add(right_bottom_frame, weight=3)

        # ========== Top Left: Core Results ==========
        # Site selection and buttons
        site_select_frame = ttk.Frame(left_top_frame)
        site_select_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(site_select_frame, text="View Site Results:").pack(side=tk.LEFT, padx=5)
        self.result_site_combo = ttk.Combobox(site_select_frame, state="readonly", width=18)
        self.result_site_combo.pack(side=tk.LEFT, padx=5)
        self.result_site_combo.bind("<<ComboboxSelected>>", self._result_site_switch_event)

        ttk.Button(site_select_frame, text="Export Results", command=self.export_calculation_results).pack(
            side=tk.RIGHT, padx=5)
        ttk.Button(site_select_frame, text="Clear Results", command=self.clear_results).pack(side=tk.RIGHT, padx=5)

        # Result display variables
        self.liquefaction_index_var = tk.DoubleVar(value=0.0)
        self.liquefaction_grade_var = tk.StringVar(value="Not Calculated")
        self.seismic_param_display_var = tk.StringVar(value="--")
        self.particle_content_display_var = tk.StringVar(value="--")
        self.beta_calculation_display_var = tk.StringVar(value="--")
        self.beta_value_display_var = tk.StringVar(value="--")
        self.depth_display_var = tk.StringVar(value="--")
        self.groundwater_display_var = tk.StringVar(value="--")
        self.overall_liquefaction_possibility_var = tk.StringVar(value="--")

        # Create scrollable frame for core results
        core_results_canvas = tk.Canvas(left_top_frame)
        core_scrollbar = ttk.Scrollbar(left_top_frame, orient=tk.VERTICAL, command=core_results_canvas.yview)
        self.core_results_frame = ttk.Frame(core_results_canvas)

        self.core_results_frame.bind(
            "<Configure>",
            lambda e: core_results_canvas.configure(scrollregion=core_results_canvas.bbox("all"))
        )

        core_results_canvas.create_window((0, 0), window=self.core_results_frame, anchor=tk.NW)
        core_results_canvas.configure(yscrollcommand=core_scrollbar.set)

        core_results_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        core_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Layout result display items
        display_items = [
            ("Seismic Parameters:", self.seismic_param_display_var),
            ("Particle Content Type:", self.particle_content_display_var),
            ("β Calculation Method:", self.beta_calculation_display_var),
            ("Design Earthquake Adjustment Coefficient β:", self.beta_value_display_var),
            ("Discrimination Depth:", self.depth_display_var),
            ("Groundwater Depth:", self.groundwater_display_var),
            ("Overall Liquefaction Possibility:", self.overall_liquefaction_possibility_var),
            ("Liquefaction Index (ILE):", self.liquefaction_index_var),
            ("Liquefaction Grade:", self.liquefaction_grade_var)
        ]

        for i, (label_text, variable) in enumerate(display_items):
            ttk.Label(self.core_results_frame, text=label_text, font=("Arial", 10)).grid(
                row=i, column=0, sticky=tk.W, pady=5, padx=10)
            if label_text in ["Liquefaction Index (ILE):", "Liquefaction Grade:", "Overall Liquefaction Possibility:"]:
                ttk.Label(self.core_results_frame, textvariable=variable, font=("Arial", 12, "bold"),
                          foreground="blue" if "Possibility" in label_text else "black").grid(
                    row=i, column=1, pady=5, padx=10, sticky=tk.W)
            else:
                ttk.Label(self.core_results_frame, textvariable=variable, font=("Arial", 10)).grid(
                    row=i, column=1, pady=5, padx=10, sticky=tk.W)

        # Building fortification class selection
        ttk.Label(self.core_results_frame, text="Building Seismic Fortification Class:").grid(
            row=len(display_items), column=0, sticky=tk.W, pady=5, padx=10)
        self.building_class_combo = ttk.Combobox(self.core_results_frame, values=["Class B", "Class C", "Class D"],
                                                 width=15)
        self.building_class_combo.current(0)
        self.building_class_combo.grid(row=len(display_items), column=1, pady=5, padx=10, sticky=tk.W)
        self.building_class_combo.bind("<<ComboboxSelected>>", self._update_anti_liquefaction_measures)

        # Liquefaction probability threshold explanation
        ttk.Label(self.core_results_frame, text="Liquefaction Probability Threshold Explanation:",
                  font=("Arial", 9, "italic")).grid(row=len(display_items) + 1, column=0, sticky=tk.W, pady=3, padx=10)
        explanation_text = "PL≥0.5: High Possibility; 0.3≤PL<0.5: Medium Possibility; PL<0.3: Low Possibility"
        ttk.Label(self.core_results_frame, text=explanation_text, font=("Arial", 9)).grid(
            row=len(display_items) + 1, column=1, sticky=tk.W, pady=3, padx=10)

        # ========== Top Right: Standards and Measures ==========
        # Create notebook for different standards
        standards_notebook = ttk.Notebook(right_top_frame)
        standards_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Tab 1: Liquefaction grade standards
        grade_tab = ttk.Frame(standards_notebook)
        standards_notebook.add(grade_tab, text="Grade Standards")

        grade_text = """Liquefaction Grade Determination Standards:

When discrimination depth is 15m:
• 0＜ILE≤5 → Slight Liquefaction
• 5＜ILE≤15 → Moderate Liquefaction
• ILE＞15 → Severe Liquefaction

When discrimination depth is 20m:
• 0＜ILE≤6 → Slight Liquefaction
• 6＜ILE≤18 → Moderate Liquefaction
• ILE＞18 → Severe Liquefaction

Liquefaction Discrimination Possibility:
• PL ≥ 0.5 → High Possibility
• 0.3 ≤ PL ＜ 0.5 → Medium Possibility
• PL ＜ 0.3 → Low Possibility"""

        grade_text_widget = scrolledtext.ScrolledText(grade_tab, wrap=tk.WORD, height=15)
        grade_text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        grade_text_widget.insert(tk.END, grade_text)
        grade_text_widget.config(state=tk.DISABLED)

        # Tab 2: Anti-liquefaction measures
        measures_tab = ttk.Frame(standards_notebook)
        standards_notebook.add(measures_tab, text="Anti-liquefaction Measures")

        self.anti_liquefaction_measures_var = tk.StringVar(value="Please calculate liquefaction grade first")
        measures_text_widget = scrolledtext.ScrolledText(measures_tab, wrap=tk.WORD, height=15)
        measures_text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        measures_text_widget.insert(tk.END,
                                    "Measures will be displayed here based on liquefaction grade and building class.")
        measures_text_widget.config(state=tk.DISABLED)
        self.measures_text_widget = measures_text_widget

        # ========== Bottom Left: Calculation Details ==========
        self.details_text = scrolledtext.ScrolledText(left_bottom_frame, wrap=tk.NONE)
        self.details_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Configure tags for better display
        self.details_text.tag_configure("header", font=("Courier New", 10, "bold"))
        self.details_text.tag_configure("data", font=("Courier New", 9))
        self.details_text.tag_configure("summary", font=("Courier New", 10, "bold"), foreground="blue")

        # ========== Bottom Right: Chart ==========
        # Create figure with better size
        self.figure, self.axes = plt.subplots(figsize=(8, 6))
        self.figure.subplots_adjust(left=0.15, right=0.95, top=0.95, bottom=0.1)
        self.canvas = FigureCanvasTkAgg(self.figure, master=right_bottom_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Add toolbar for chart interaction (optional)
        # from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        # toolbar = NavigationToolbar2Tk(self.canvas, right_bottom_frame)
        # toolbar.update()

    def export_calculation_results(self):
        """Export calculation results to Excel table"""
        if not self.calculation_results:
            messagebox.showwarning("Warning", "No calculation results to export, please calculate first")
            return

        try:
            # Show save file dialog
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"Liquefaction_Calculation_Results_{timestamp}.xlsx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialfile=default_filename
            )

            if not file_path:  # User canceled save
                return

            # Create progress window
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Exporting")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()

            ttk.Label(progress_window, text="Exporting calculation results, please wait...").pack(pady=20)
            progress_label = ttk.Label(progress_window, text="")
            progress_label.pack()
            self.root.update_idletasks()

            # Create an Excel writer
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                total_site_count = len(self.calculation_results)
                current_site = 0

                # 1. Create summary sheet
                summary_data = []
                for site_number, result in self.calculation_results.items():
                    current_site += 1
                    progress_window.title(f"Exporting ({current_site}/{total_site_count})")
                    progress_label.config(text=f"Processing site: {site_number}")
                    self.root.update_idletasks()

                    # Calculate average PL and overall liquefaction possibility
                    PL_list = [soil_layer["PL"] for soil_layer in result["soil_layer_calculation_details"]]
                    average_PL = sum(PL_list) / len(PL_list) if PL_list else 0.0
                    if average_PL >= 0.5:
                        overall_liquefaction_possibility = "High Possibility"
                    elif 0.3 <= average_PL < 0.5:
                        overall_liquefaction_possibility = "Medium Possibility"
                    else:
                        overall_liquefaction_possibility = "Low Possibility"

                    summary_data.append({
                        "Site Number": site_number,
                        "Seismic Parameter Choice": result["seismic_param_choice_result"],
                        "Particle Content Type": result["particle_content_choice_result"],
                        "β Calculation Method": result["beta_calculation_method_choice_result"],
                        "Design Earthquake Adjustment Coefficient β": round(
                            result["design_earthquake_adjustment_coefficient_β"], 3),
                        "Discrimination Depth(m)": result["discrimination_depth"],
                        "Groundwater Depth dw(m)": result["groundwater_depth_dw"],
                        "Liquefaction Index ILE": round(result["liquefaction_index_ILE"], 3),
                        "Liquefaction Grade": result["liquefaction_grade"],
                        "Average Liquefaction Probability PL": round(average_PL, 3),
                        "Overall Liquefaction Possibility": overall_liquefaction_possibility,
                        "Soil Layer Count": len(result["soil_layer_calculation_details"])
                    })

                # Write summary sheet
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name="Calculation Results Summary", index=False)

                # 2. Create detailed worksheet for each site
                for site_number, result in self.calculation_results.items():
                    # Process site name to ensure worksheet name is valid
                    worksheet_name = str(site_number)[:31]  # Excel worksheet name maximum 31 characters
                    # Remove illegal characters
                    worksheet_name = re.sub(r'[\\/*?:[\]]', '_', worksheet_name)

                    # Prepare detailed data
                    detailed_data = []
                    for i, soil_layer in enumerate(result["soil_layer_calculation_details"]):
                        # Determine each layer's liquefaction possibility
                        PL = soil_layer["PL"]
                        if PL >= 0.5:
                            liquefaction_possibility = "High Possibility"
                        elif 0.3 <= PL < 0.5:
                            liquefaction_possibility = "Medium Possibility"
                        else:
                            liquefaction_possibility = "Low Possibility"

                        detailed_data.append({
                            "Soil Layer No.": i + 1,
                            "Saturated Sand Depth ds(m)": round(soil_layer["ds"], 3),
                            "Measured SPT Blow Count N": round(soil_layer["N"], 1),
                            "Content Value(%)": round(soil_layer["content_value"], 1),
                            "Standard Penetration Critical Value Ncr": round(soil_layer["Ncr"], 3),
                            "Liquefaction Probability PL": round(soil_layer["PL"], 3),
                            "Arbitrary Probability Critical Value Ncr,PL": round(soil_layer["Ncr_PL"], 3),
                            "Liquefaction Discrimination Possibility": liquefaction_possibility,
                            "Weight wi": round(soil_layer["wi"], 1),
                            "Contribution Value": round(soil_layer["contribution_value"], 3)
                        })

                    # Create DataFrame
                    detailed_df = pd.DataFrame(detailed_data)

                    # Add summary row
                    summary_row = pd.DataFrame([{
                        "Soil Layer No.": "Summary",
                        "Saturated Sand Depth ds(m)": "",
                        "Measured SPT Blow Count N": "",
                        "Content Value(%)": "",
                        "Standard Penetration Critical Value Ncr": "",
                        "Liquefaction Probability PL": "",
                        "Arbitrary Probability Critical Value Ncr,PL": "",
                        "Liquefaction Discrimination Possibility": "",
                        "Weight wi": "",
                        "Contribution Value": round(result["liquefaction_index_ILE"], 3)
                    }])

                    detailed_df = pd.concat([detailed_df, summary_row], ignore_index=True)

                    # Write worksheet
                    detailed_df.to_excel(writer, sheet_name=worksheet_name, index=False)

                # 3. Create parameter explanation worksheet
                param_explanation_data = [
                    {"Parameter Category": "Seismic Parameters", "Option": "Surface Peak Acceleration amax",
                     "Explanation": "Directly input amax value(g)"},
                    {"Parameter Category": "Seismic Parameters", "Option": "Seismic Intensity",
                     "Explanation": "7/8/9 degrees, corresponding to amax=0.1/0.2/0.4g"},
                    {"Parameter Category": "Particle Content", "Option": "Clay Content ρc",
                     "Explanation": "Clay content percentage, automatically takes 3% when <3%"},
                    {"Parameter Category": "Particle Content", "Option": "Fine Particle Content Fc",
                     "Explanation": "Fine particle content percentage, automatically takes 15% when <15%"},
                    {"Parameter Category": "β Calculation Method", "Option": "Design Earthquake Group",
                     "Explanation": "Group 1=0.80, Group 2=0.95, Group 3=1.05"},
                    {"Parameter Category": "β Calculation Method", "Option": "Formula Calculation",
                     "Explanation": "β=0.20Ms-0.50, Ms is surface wave magnitude"},
                    {"Parameter Category": "Liquefaction Grade", "Option": "Discrimination Depth 15m",
                     "Explanation": "Slight(0<ILE≤5), Moderate(5<ILE≤15), Severe(ILE>15)"},
                    {"Parameter Category": "Liquefaction Grade", "Option": "Discrimination Depth 20m",
                     "Explanation": "Slight(0<ILE≤6), Moderate(6<ILE≤18), Severe(ILE>18)"},
                    {"Parameter Category": "Liquefaction Possibility", "Option": "High Possibility",
                     "Explanation": "PL≥0.5, high probability of liquefaction under earthquake"},
                    {"Parameter Category": "Liquefaction Possibility", "Option": "Medium Possibility",
                     "Explanation": "0.3≤PL<0.5, possible liquefaction under earthquake"},
                    {"Parameter Category": "Liquefaction Possibility", "Option": "Low Possibility",
                     "Explanation": "PL<0.3, low probability of liquefaction under earthquake"}
                ]

                param_explanation_df = pd.DataFrame(param_explanation_data)
                param_explanation_df.to_excel(writer, sheet_name="Parameter Explanation", index=False)

                # 4. Create anti-liquefaction measures worksheet
                measures_data = [
                    {"Building Class": "Class B",
                     "Slight Liquefaction": "Partially eliminate liquefaction settlement, or treat foundation and superstructure",
                     "Moderate Liquefaction": "Completely eliminate liquefaction settlement, or partially eliminate liquefaction settlement and treat foundation and superstructure",
                     "Severe Liquefaction": "Completely eliminate liquefaction settlement"},
                    {"Building Class": "Class C",
                     "Slight Liquefaction": "Treat foundation and superstructure, or may not treat",
                     "Moderate Liquefaction": "Treat foundation and superstructure, or take stricter measures",
                     "Severe Liquefaction": "Completely eliminate liquefaction settlement, or partially eliminate liquefaction settlement and treat foundation and superstructure"},
                    {"Building Class": "Class D", "Slight Liquefaction": "No measures needed",
                     "Moderate Liquefaction": "No measures needed",
                     "Severe Liquefaction": "Treat foundation and superstructure, or take other economical and reasonable measures"}
                ]

                measures_df = pd.DataFrame(measures_data)
                measures_df.to_excel(writer, sheet_name="Anti-liquefaction Measures", index=False)

            # Close progress window
            progress_window.destroy()

            # Show success message
            messagebox.showinfo("Export Successful",
                                f"Calculation results successfully exported to:\n{file_path}\n\nExported {total_site_count} sites' calculation results")

        except Exception as e:
            messagebox.showerror("Export Failed", f"Error occurred during export:\n{str(e)}")

    def clear_results(self):
        """Clear all calculation results"""
        self.calculation_results = {}
        self.result_site_combo['values'] = []
        # Reset display variables
        self.liquefaction_index_var.set(0.0)
        self.liquefaction_grade_var.set("Not Calculated")
        self.seismic_param_display_var.set("--")
        self.particle_content_display_var.set("--")
        self.beta_calculation_display_var.set("--")
        self.beta_value_display_var.set("--")
        self.depth_display_var.set("--")
        self.groundwater_display_var.set("--")
        self.overall_liquefaction_possibility_var.set("--")
        self.details_text.delete(1.0, tk.END)
        self.axes.clear()
        self.canvas.draw()
        self.measures_text_widget.config(state=tk.NORMAL)
        self.measures_text_widget.delete(1.0, tk.END)
        self.measures_text_widget.insert(tk.END,
                                         "Measures will be displayed here based on liquefaction grade and building class.")
        self.measures_text_widget.config(state=tk.DISABLED)

    def _result_site_switch_event(self, event):
        """Update display when switching result sites"""
        if not self.calculation_results:
            return
        site_number = self.result_site_combo.get()
        if site_number in self.calculation_results:
            self._update_result_display(self.calculation_results[site_number])

    def _update_anti_liquefaction_measures(self, event):
        """Update anti-liquefaction measures suggestions"""
        current_grade = self.liquefaction_grade_var.get()
        if current_grade in ["Not Calculated", "No Liquefaction"]:
            self.measures_text_widget.config(state=tk.NORMAL)
            self.measures_text_widget.delete(1.0, tk.END)
            self.measures_text_widget.insert(tk.END, "Please calculate valid liquefaction grade first")
            self.measures_text_widget.config(state=tk.DISABLED)
            return

        building_class = self.building_class_combo.get()
        # Anti-liquefaction measures mapping (document Table 2)
        measures_mapping = {
            "Class B": {
                "Slight Liquefaction": "Partially eliminate liquefaction settlement, or treat foundation and superstructure",
                "Moderate Liquefaction": "Completely eliminate liquefaction settlement, or partially eliminate liquefaction settlement and treat foundation and superstructure",
                "Severe Liquefaction": "Completely eliminate liquefaction settlement"
            },
            "Class C": {
                "Slight Liquefaction": "Treat foundation and superstructure, or may not treat",
                "Moderate Liquefaction": "Treat foundation and superstructure, or take stricter measures",
                "Severe Liquefaction": "Completely eliminate liquefaction settlement, or partially eliminate liquefaction settlement and treat foundation and superstructure"
            },
            "Class D": {
                "Slight Liquefaction": "No measures needed",
                "Moderate Liquefaction": "No measures needed",
                "Severe Liquefaction": "Treat foundation and superstructure, or take other economical and reasonable measures"
            }
        }

        measure_text = f"Building Seismic Fortification Class: {building_class}\n"
        measure_text += f"Liquefaction Grade: {current_grade}\n\n"
        measure_text += "Recommended Anti-liquefaction Measures:\n"
        measure_text += measures_mapping[building_class].get(current_grade, "No corresponding measures")
        measure_text += "\n\nNote: These measures are based on《SPT-based Liquefaction Index Calculation Method》 standards."

        self.measures_text_widget.config(state=tk.NORMAL)
        self.measures_text_widget.delete(1.0, tk.END)
        self.measures_text_widget.insert(tk.END, measure_text)
        self.measures_text_widget.config(state=tk.DISABLED)

    def get_manual_input_data(self):
        """Get current site manual input data (process binary choice logic)"""
        try:
            current_site_number = f"Site {self.current_site_index + 1}"
            acceleration_intensity_choice = self.manual_acceleration_intensity_choice.get()
            particle_content_choice = self.manual_particle_content_choice.get()
            beta_calculation_method = self.manual_beta_calculation_choice.get()

            # Process seismic parameters
            if acceleration_intensity_choice == "amax":
                seismic_param_value = self.amax_var.get()
                intensity = 7  # Placeholder, actually use amax
                amax = seismic_param_value
            else:
                seismic_param_value = int(self.intensity_var.get())
                intensity = seismic_param_value
                amax = 0.1  # Placeholder, actually use intensity

            # Process other parameters
            discrimination_depth = int(self.depth_var.get())
            groundwater_depth = self.groundwater_var.get()
            magnitude = self.magnitude_var.get()
            earthquake_group = self.earthquake_group_var.get()

            # Process soil layer data
            soil_layer_list = []
            for entry_row in self.soil_layer_entries:
                ds = float(entry_row[0].get())
                N = float(entry_row[1].get())
                content_value = float(entry_row[2].get())

                # Process limits based on selected particle content type
                if particle_content_choice == "ρc":
                    content_value = max(3.0, content_value)
                else:
                    content_value = max(15.0, content_value)

                soil_layer_list.append({
                    "ds": ds,
                    "N": N,
                    "content_value": content_value,
                    "seismic_param_value": seismic_param_value
                })

                # Save current site data
            self.site_data[self.current_site_index] = {
                "acceleration_intensity_choice": acceleration_intensity_choice,
                "seismic_intensity": intensity,
                "surface_peak_acceleration_amax(g)": amax,
                "discrimination_depth": discrimination_depth,
                "groundwater_depth_dw": groundwater_depth,
                "surface_wave_magnitude_Ms": magnitude,
                "design_earthquake_group": earthquake_group,
                "beta_calculation_method": beta_calculation_method,
                "particle_content_choice": particle_content_choice,
                "soil_layer_data": soil_layer_list
            }

            return current_site_number, {
                "acceleration_intensity_choice": acceleration_intensity_choice,
                "seismic_intensity": intensity,
                "surface_peak_acceleration_amax(g)": amax,
                "discrimination_depth": discrimination_depth,
                "groundwater_depth_dw": groundwater_depth,
                "surface_wave_magnitude_Ms": magnitude,
                "design_earthquake_group": earthquake_group,
                "beta_calculation_method": beta_calculation_method,
                "particle_content_choice": particle_content_choice,
                "soil_layer_data": soil_layer_list
            }

        except Exception as e:
            messagebox.showerror("Input Error", f"Please check input data: {str(e)}")
            return None, None

    def manual_calculation(self):
        """Manual input data calculation"""
        site_number, data = self.get_manual_input_data()
        if data:
            result = self.calculate_liquefaction_index(data)
            self.calculation_results[site_number] = result
            self._refresh_result_site_combo()
            self._result_site_switch_event(None)
            self.notebook.select(self.result_display_tab)

    def _refresh_result_site_combo(self):
        """Refresh result site selection combo box"""
        site_number_list = list(self.calculation_results.keys())
        self.result_site_combo['values'] = site_number_list
        if site_number_list:
            self.result_site_combo.current(0)

    def _update_result_display(self, result):
        """Update result display (including liquefaction discrimination possibility)"""
        # Update core results
        self.seismic_param_display_var.set(result["seismic_param_choice_result"])
        self.particle_content_display_var.set(result["particle_content_choice_result"])
        self.beta_calculation_display_var.set(result["beta_calculation_method_choice_result"])
        self.beta_value_display_var.set(f"{result['design_earthquake_adjustment_coefficient_β']:.2f}")
        self.depth_display_var.set(f"{result['discrimination_depth']}m")
        self.groundwater_display_var.set(f"{result['groundwater_depth_dw']}m")
        self.liquefaction_index_var.set(round(result["liquefaction_index_ILE"], 2))
        self.liquefaction_grade_var.set(result["liquefaction_grade"])

        # Calculate and display overall liquefaction possibility (based on average PL of all soil layers)
        PL_list = [soil_layer["PL"] for soil_layer in result["soil_layer_calculation_details"]]
        average_PL = sum(PL_list) / len(PL_list) if PL_list else 0.0
        if average_PL >= 0.5:
            overall_liquefaction_possibility = "High Possibility"
        elif 0.3 <= average_PL < 0.5:
            overall_liquefaction_possibility = "Medium Possibility"
        else:
            overall_liquefaction_possibility = "Low Possibility"
        self.overall_liquefaction_possibility_var.set(
            f"{overall_liquefaction_possibility} (Average PL={average_PL:.3f})")

        # Update detailed results
        self.details_text.delete(1.0, tk.END)
        particle_content_header = "Clay Content ρc(%)" if result[
                                                              "particle_content_choice"] == "ρc" else "Fine Particle Content Fc(%)"

        # Header
        self.details_text.insert(tk.END,
                                 f"{'Saturation Depth(m)':^15}{'Measured N':^10}{particle_content_header:^15}{'Ncr':^12}{'PL':^12}{'Ncr,PL':^12}{'Possibility':^15}{'Contribution':^10}\n",
                                 "header")
        self.details_text.insert(tk.END, "=" * 105 + "\n", "header")

        # Data rows
        for soil_layer in result["soil_layer_calculation_details"]:
            # Determine each layer's liquefaction possibility
            PL = soil_layer["PL"]
            if PL >= 0.5:
                liquefaction_possibility = "High"
                possibility_color = "red"
            elif 0.3 <= PL < 0.5:
                liquefaction_possibility = "Medium"
                possibility_color = "orange"
            else:
                liquefaction_possibility = "Low"
                possibility_color = "green"

            row_text = f"{soil_layer['ds']:^15.1f}{soil_layer['N']:^10.1f}{soil_layer['content_value']:^15.1f}{soil_layer['Ncr']:^12.2f}{soil_layer['PL']:^12.3f}{soil_layer['Ncr_PL']:^12.2f}{liquefaction_possibility:^15}{soil_layer['contribution_value']:^10.2f}\n"

            # Create tag for colored possibility
            tag_name = f"possibility_{possibility_color}"
            self.details_text.tag_configure(tag_name, foreground=possibility_color)

            start_pos = self.details_text.index("end-1c")
            self.details_text.insert(tk.END, row_text, "data")
            end_pos = self.details_text.index("end-1c")

            # Apply color to possibility column
            possibility_start = f"{start_pos}+{90}c"  # Adjust column position
            possibility_end = f"{start_pos}+{105}c"
            self.details_text.tag_add(tag_name, possibility_start, possibility_end)

        # Summary
        self.details_text.insert(tk.END, "=" * 105 + "\n", "header")
        self.details_text.insert(tk.END, f"\nSeismic Parameter Choice: {result['seismic_param_choice_result']}\n",
                                 "summary")
        self.details_text.insert(tk.END, f"Particle Content Choice: {result['particle_content_choice_result']}\n",
                                 "summary")
        self.details_text.insert(tk.END, f"β Calculation Method: {result['beta_calculation_method_choice_result']}\n",
                                 "summary")
        self.details_text.insert(tk.END,
                                 f"Design Earthquake Adjustment Coefficient β: {result['design_earthquake_adjustment_coefficient_β']:.2f}\n",
                                 "summary")
        self.details_text.insert(tk.END,
                                 f"Overall Liquefaction Possibility: {overall_liquefaction_possibility} (Average PL={average_PL:.3f})\n",
                                 "summary")
        self.details_text.insert(tk.END, f"Total Liquefaction Index (ILE): {result['liquefaction_index_ILE']:.2f}\n",
                                 "summary")
        self.details_text.insert(tk.END, f"Liquefaction Grade: {result['liquefaction_grade']}\n", "summary")

        # Draw liquefaction probability distribution chart
        self.axes.clear()
        depth_list = [soil_layer["ds"] for soil_layer in result["soil_layer_calculation_details"]]
        liquefaction_probability_list = [soil_layer["PL"] for soil_layer in result["soil_layer_calculation_details"]]

        # Categorize plotting by possibility
        high_possibility_x = []
        high_possibility_y = []
        medium_possibility_x = []
        medium_possibility_y = []
        low_possibility_x = []
        low_possibility_y = []

        for i, (pl, ds) in enumerate(zip(liquefaction_probability_list, depth_list)):
            if pl >= 0.5:
                high_possibility_x.append(pl)
                high_possibility_y.append(ds)
            elif 0.3 <= pl < 0.5:
                medium_possibility_x.append(pl)
                medium_possibility_y.append(ds)
            else:
                low_possibility_x.append(pl)
                low_possibility_y.append(ds)

        # Plot points and lines
        if high_possibility_x:
            self.axes.scatter(high_possibility_x, high_possibility_y, c='red', s=60, label='High Possibility (PL≥0.5)',
                              zorder=5)
            self.axes.plot(high_possibility_x, high_possibility_y, 'r-', linewidth=1.5, alpha=0.7)
        if medium_possibility_x:
            self.axes.scatter(medium_possibility_x, medium_possibility_y, c='orange', s=60,
                              label='Medium Possibility (0.3≤PL<0.5)', zorder=5)
            self.axes.plot(medium_possibility_x, medium_possibility_y, 'orange', linewidth=1.5, alpha=0.7)
        if low_possibility_x:
            self.axes.scatter(low_possibility_x, low_possibility_y, c='green', s=60, label='Low Possibility (PL<0.3)',
                              zorder=5)
            self.axes.plot(low_possibility_x, low_possibility_y, 'g-', linewidth=1.5, alpha=0.7)

        # Add critical lines
        self.axes.axvline(x=0.5, color='r', linestyle='--', alpha=0.5, linewidth=1)
        self.axes.axvline(x=0.3, color='orange', linestyle='--', alpha=0.5, linewidth=1)

        # Fill regions
        self.axes.axvspan(0.5, 1.0, alpha=0.1, color='red')
        self.axes.axvspan(0.3, 0.5, alpha=0.1, color='orange')
        self.axes.axvspan(0.0, 0.3, alpha=0.1, color='green')

        # Add labels and titles
        self.axes.set_xlabel('Liquefaction Probability (PL)', fontsize=12)
        self.axes.set_ylabel('Saturated Sand Depth (m)', fontsize=12)
        title_text = f'Liquefaction Probability Distribution\n{result["seismic_param_choice_result"]}'
        self.axes.set_title(title_text, fontsize=14, fontweight='bold')
        self.axes.invert_yaxis()  # Depth increases downward
        self.axes.legend(loc='lower right', fontsize=10)
        self.axes.grid(True, alpha=0.3, linestyle='--')

        # Set axis limits
        self.axes.set_xlim(0, 1.0)
        if depth_list:
            self.axes.set_ylim(max(depth_list) + 2, min(depth_list) - 2)

        # Add text annotations
        self.axes.text(0.75, 0.95, f'ILE = {result["liquefaction_index_ILE"]:.2f}',
                       transform=self.axes.transAxes, fontsize=11, verticalalignment='top',
                       bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))

        self.figure.tight_layout()
        self.canvas.draw()

        # Update anti-liquefaction measures
        self._update_anti_liquefaction_measures(None)

    def import_calculation(self):
        """Import data calculation (process binary choice logic)"""
        # Check if column recognition is complete
        core_columns = ["Site Number", "Saturated Sand Depth ds(m)", "Measured SPT Blow Count N",
                        "Clay/Fine Particle Content(%)"]
        missing_columns = [col_name for col_name in core_columns if not self.current_column_mapping.get(col_name)]
        if missing_columns:
            messagebox.showwarning("Warning",
                                   f"The following necessary columns are not recognized, please check table:\n{', '.join(missing_columns)}")
            return

        # Check seismic parameter column
        seismic_param_column = self.current_column_mapping.get(
            "Surface Peak Acceleration amax(g)") if self.import_acceleration_intensity_choice.get() == "amax" else self.current_column_mapping.get(
            "Seismic Intensity")
        if not seismic_param_column:
            messagebox.showwarning("Warning", "Selected seismic parameter column not recognized, please check data")
            return

        if not self.imported_site_data or self.import_site_combo.current() == -1:
            messagebox.showwarning("Warning", "Please import data and select calculation site first")
            return

        try:
            selected_site_number = self.import_site_combo.get()
            site_data = self.imported_site_data[selected_site_number]
            acceleration_intensity_choice = self.import_acceleration_intensity_choice.get()
            particle_content_choice = self.import_particle_content_choice.get()
            beta_calculation_method = self.import_beta_calculation_choice.get()
            discrimination_depth = int(self.import_depth_var.get())
            groundwater_depth = self.import_groundwater_var.get()
            magnitude = self.import_magnitude_var.get()
            earthquake_group = self.import_earthquake_group_var.get()

            # Construct soil layer data
            soil_layer_list = []
            content_column = self.current_column_mapping["Clay/Fine Particle Content(%)"]
            seismic_param_column = self.current_column_mapping.get(
                "Surface Peak Acceleration amax(g)") if acceleration_intensity_choice == "amax" else self.current_column_mapping.get(
                "Seismic Intensity")

            for _, row_data in site_data.iterrows():
                # Process particle content
                content_value = row_data[content_column]
                if particle_content_choice == "ρc":
                    content_value = max(3.0, float(content_value)) if pd.notna(content_value) else 3.0
                else:
                    content_value = max(15.0, float(content_value)) if pd.notna(content_value) else 15.0

                # Process seismic parameter
                seismic_param_value = row_data[seismic_param_column]
                if acceleration_intensity_choice == "amax":
                    seismic_param_value = float(seismic_param_value) if pd.notna(seismic_param_value) else 0.1
                else:
                    seismic_param_value = int(seismic_param_value) if pd.notna(seismic_param_value) else 7

                soil_layer_list.append({
                    "ds": row_data[self.current_column_mapping["Saturated Sand Depth ds(m)"]],
                    "N": row_data[self.current_column_mapping["Measured SPT Blow Count N"]],
                    "content_value": content_value,
                    "seismic_param_value": seismic_param_value
                })

            # Construct calculation data
            calculation_data = {
                "acceleration_intensity_choice": acceleration_intensity_choice,
                "seismic_intensity": 7 if acceleration_intensity_choice != "intensity" else seismic_param_value,
                "surface_peak_acceleration_amax(g)": 0.1 if acceleration_intensity_choice != "amax" else seismic_param_value,
                "discrimination_depth": discrimination_depth,
                "groundwater_depth_dw": groundwater_depth,
                "surface_wave_magnitude_Ms": magnitude,
                "design_earthquake_group": earthquake_group,
                "beta_calculation_method": beta_calculation_method,
                "particle_content_choice": particle_content_choice,
                "soil_layer_data": soil_layer_list
            }

            # Calculate and store result
            result = self.calculate_liquefaction_index(calculation_data)
            self.calculation_results[f"Imported Site {selected_site_number}"] = result
            self._refresh_result_site_combo()
            self._result_site_switch_event(None)
            self.notebook.select(self.result_display_tab)

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def calculate_liquefaction_index(self, data):
        """Core calculation (strictly follow document formulas, process binary choice parameter logic + liquefaction discrimination possibility)"""
        # Get basic parameters
        acceleration_intensity_choice = data["acceleration_intensity_choice"]
        particle_content_choice = data["particle_content_choice"]
        beta_calculation_method = data["beta_calculation_method"]
        seismic_intensity = data["seismic_intensity"]
        amax_input_value = data["surface_peak_acceleration_amax(g)"]
        discrimination_depth = data["discrimination_depth"]
        groundwater_depth_dw = data["groundwater_depth_dw"]
        surface_wave_magnitude_Ms = data["surface_wave_magnitude_Ms"]
        design_earthquake_group = data["design_earthquake_group"]
        soil_layer_data = data["soil_layer_data"]

        # Record parameter selection results (for display)
        seismic_param_choice_result = ""
        particle_content_choice_result = f"Clay Content ρc" if particle_content_choice == "ρc" else f"Fine Particle Content Fc"
        beta_calculation_method_choice_result = ""

        # Step 1: Calculate design earthquake adjustment coefficient β (binary choice, strictly follow document definition)
        design_earthquake_adjustment_coefficient_β = 0.80
        if beta_calculation_method == "group":
            beta_calculation_method_choice_result = f"Design Earthquake Group ({design_earthquake_group} group)"
            if design_earthquake_group == "1":
                design_earthquake_adjustment_coefficient_β = 0.80
            elif design_earthquake_group == "2":
                design_earthquake_adjustment_coefficient_β = 0.95
            elif design_earthquake_group == "3":
                design_earthquake_adjustment_coefficient_β = 1.05
            else:
                # When group invalid, default to group 1 calculation
                design_earthquake_adjustment_coefficient_β = 0.80
        else:
            beta_calculation_method_choice_result = f"Formula Calculation (Ms={surface_wave_magnitude_Ms})"
            design_earthquake_adjustment_coefficient_β = 0.20 * surface_wave_magnitude_Ms - 0.50
            design_earthquake_adjustment_coefficient_β = max(0.80, min(1.05,
                                                                       design_earthquake_adjustment_coefficient_β))  # Limit to reasonable range

        # Step 2: Determine seismic parameters and calculation method (binary choice, strictly follow document mapping relationship)
        amax = 0.1
        N_prime = 16
        if acceleration_intensity_choice == "amax":
            amax = amax_input_value
            seismic_param_choice_result = f"Surface Peak Acceleration amax={amax:.2f}g"
            # Verify amax validity (document specified range: 0.1-0.4g)
            if not (0.09 <= amax <= 0.41):
                raise ValueError(f"amax value {amax:.2f}g exceeds valid range (0.1-0.4g)")
        else:
            seismic_param_choice_result = f"Seismic Intensity={seismic_intensity} degrees (corresponding amax={0.1 * seismic_intensity // 7 if seismic_intensity == 7 else 0.2 if seismic_intensity == 8 else 0.4}g)"
            # When based on intensity, map amax and N' according to document table
            amax_mapping = {7: 0.1, 8: 0.2, 9: 0.4}
            N_prime_mapping = {0.10: 16, 0.15: 20, 0.20: 23, 0.30: 31, 0.40: 37}
            amax = amax_mapping.get(seismic_intensity, 0.1)
            N_prime = N_prime_mapping.get(amax, 16)

        soil_layer_calculation_details = []
        total_liquefaction_index_ILE = 0.0

        for soil_layer in soil_layer_data:
            ds = soil_layer["ds"]  # Saturated sand depth
            N = soil_layer["N"]  # Measured SPT blow count
            content_value = soil_layer[
                "content_value"]  # Clay/fine particle content (already processed limits as per document requirements: ρc≥3%, Fc≥15%)
            seismic_param_value = soil_layer["seismic_param_value"]  # Seismic parameter corresponding to soil layer

            # Step 1.1: Calculate standard penetration critical value Ncr (strictly follow document 3 formulas, binary choice logic)
            Ncr = 0.0
            if acceleration_intensity_choice == "amax":
                # Ncr calculation based on amax (document two formulas choose one)
                term1 = (69 * amax) / (amax + 0.40)
                term2 = 1 - 0.02 * groundwater_depth_dw
                term3 = 0.21 + (0.79 * ds) / (ds + 6.2)

                if particle_content_choice == "ρc":
                    # Formula 1: Contains clay content correction term (document first formula based on amax)
                    term4 = math.sqrt(3 / content_value)
                    Ncr = design_earthquake_adjustment_coefficient_β * term1 * term2 * term3 * term4
                else:
                    # Formula 2: Contains fine particle content correction term (document second formula based on amax)
                    term4 = math.sqrt(12 / (content_value - 3))
                    Ncr = term1 * term2 * term3 * term4
            else:
                # Ncr calculation based on intensity (document independent formula)
                term1 = 0.79 * N_prime
                term2 = 1 - 0.02 * groundwater_depth_dw
                term3 = 0.27 + ds / (ds + 6.2)
                Ncr = term1 * term2 * term3

            Ncr = max(1.0, Ncr)  # Avoid zero or negative values, conform to engineering practice

            # Step 1.2: Calculate liquefaction probability PL (strictly follow document formula)
            exponent_term = 0.32 * (N - Ncr)
            PL = 1.0 / (1.0 + math.exp(exponent_term))
            PL = max(0.0, min(1.0, PL))  # Limit to [0,1] range

            # Step 1.3: Calculate critical value Ncr,PL at any liquefaction probability (strictly follow document formula)
            if PL <= 0.001:
                Ncr_PL = N + 20.0  # Avoid extreme values causing calculation anomalies
            elif PL >= 0.999:
                Ncr_PL = N - 20.0  # Avoid extreme values causing calculation anomalies
            else:
                Ncr_PL = Ncr - (math.log(PL / (1.0 - PL))) / 0.32
            Ncr_PL = max(0.1, Ncr_PL)  # Avoid zero or negative values

            # Step 1.4: Calculate weight wi (reference GB50011-2010 specification, related to depth)
            if ds <= 5:
                wi = 10.0
            elif ds <= 15:
                wi = 5.0
            elif ds <= 20:
                wi = 2.0
            else:
                wi = 0.0  # Beyond 20m not included in liquefaction index calculation (conforms to document discrimination depth requirements)

            # Step 1.4: Calculate liquefaction index ILE (document weighted logic)
            # Note: Because user requested deletion of soil layer thickness di parameter, using document implied 3m standard layer thickness (engineering conventional value)
            di = 3.0  # Standard layer thickness, conforms to SPT test conventional layer thickness requirements
            contribution_value = PL * di * wi
            total_liquefaction_index_ILE += contribution_value

            soil_layer_calculation_details.append({
                "ds": ds,
                "N": N,
                "content_value": content_value,
                "Ncr": Ncr,
                "PL": PL,
                "Ncr_PL": Ncr_PL,
                "wi": wi,
                "contribution_value": contribution_value
            })

        # Step 1.5: Liquefaction grade determination (strictly follow document table)
        if total_liquefaction_index_ILE <= 0:
            liquefaction_grade = "No Liquefaction"
        elif discrimination_depth == 15:
            if total_liquefaction_index_ILE <= 5:
                liquefaction_grade = "Slight Liquefaction"
            elif total_liquefaction_index_ILE <= 15:
                liquefaction_grade = "Moderate Liquefaction"
            else:
                liquefaction_grade = "Severe Liquefaction"
        else:  # Discrimination depth 20m
            if total_liquefaction_index_ILE <= 6:
                liquefaction_grade = "Slight Liquefaction"
            elif total_liquefaction_index_ILE <= 18:
                liquefaction_grade = "Moderate Liquefaction"
            else:
                liquefaction_grade = "Severe Liquefaction"

        return {
            "seismic_intensity": seismic_intensity,
            "surface_peak_acceleration_amax(g)": amax_input_value,
            "discrimination_depth": discrimination_depth,
            "groundwater_depth_dw": groundwater_depth_dw,
            "surface_wave_magnitude_Ms": surface_wave_magnitude_Ms,
            "design_earthquake_group": design_earthquake_group,
            "design_earthquake_adjustment_coefficient_β": design_earthquake_adjustment_coefficient_β,
            "particle_content_choice": particle_content_choice,
            "acceleration_intensity_choice": acceleration_intensity_choice,
            "beta_calculation_method": beta_calculation_method,
            "seismic_param_choice_result": seismic_param_choice_result,
            "particle_content_choice_result": particle_content_choice_result,
            "beta_calculation_method_choice_result": beta_calculation_method_choice_result,
            "liquefaction_index_ILE": total_liquefaction_index_ILE,
            "liquefaction_grade": liquefaction_grade,
            "soil_layer_calculation_details": soil_layer_calculation_details
        }


if __name__ == "__main__":
    root_window = tk.Tk()
    app = SPTLiquefactionSoftware(root_window)
    root_window.mainloop()