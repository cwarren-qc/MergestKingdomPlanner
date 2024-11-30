# Version: 1.0.2
# https://github.com/cwarren-qc/MergestKingdomPlanner.git
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import messagebox
import json
import os
import re
from enum import Enum
from datetime import datetime, timedelta

class Column(Enum):
    LEVEL = 0
    ITEMS_NEEDED = 1
    EFFICIENT = 2
    NAME = 3
    TIME_TO_BUILD = 4
    ON_HAND = 5
    MAX_MERGE = 6
    NB_OVERRIDE = 7
    MERGE_COMBINATION = 8
    ITEMS_REQUIRED = 9 
    LIQUID_TO_USE = 10
    LIQUID_SAVING = 11
    ITEMS_TO_CREATE = 12
    
def calculate_small_items(items_needed, max_merge, override_merge, override_count, level):
    if items_needed <= 0:
        return 0, 0, ""
    
    merge_sizes = [3, 5, 9, 17, 33, 65]
    merge_outputs = [1, 2, 4, 8, 16, 32]
    
    allowed_sizes = [s for s in merge_sizes if s <= max_merge]
    allowed_outputs = [o for i, o in enumerate(merge_outputs) if merge_sizes[i] <= max_merge]
    
    # Add override size if applicable
    if override_merge > max_merge and override_count > 0:
        override_size_index = next((i for i, size in enumerate(merge_sizes) if size == override_merge), None)
        if override_size_index is not None:
            override_output = merge_outputs[override_size_index]
    else:
        override_count = 0
        override_output = 0
        
    remaining = items_needed
    small_items = 0
    combination = []
    merge_counts = {}
    override_used = 0
    
    while remaining > 0:
        # First try to use override if available
        if override_count > override_used and override_merge > max_merge:
            if remaining >= override_output:
                merge_counts[override_merge] = merge_counts.get(override_merge, 0) + 1
                small_items += override_merge
                remaining -= override_output
                override_used += 1
                continue
                
        # Then use regular merges
        used_merge = False
        for i in range(len(allowed_sizes)-1, -1, -1):
            if remaining >= allowed_outputs[i]:
                size = allowed_sizes[i]
                merge_counts[size] = merge_counts.get(size, 0) + 1
                small_items += size
                remaining -= allowed_outputs[i]
                used_merge = True
                break
                
        if not used_merge and remaining > 0:
            merge_counts[3] = merge_counts.get(3, 0) + 1
            small_items += 3
            remaining -= 1
    
    # Create compressed combination string
    compressed = []
    for size in sorted(merge_counts.keys()):
        count = merge_counts[size]
        if count == 1:
            compressed.append(str(size))
        else:
            compressed.append(f"{size} * {count}")
    
    direct_solution = (small_items, sum(merge_counts.values()), f"{items_needed:>5} <= " + " + ".join(compressed))
    
    # Check if using a bigger merge would be more efficient
    if override_count == 0:  # Only check for more efficient solution if no override is used
        for i in range(len(allowed_sizes)-1, -1, -1):
            if allowed_sizes[i] <= small_items and allowed_outputs[i] > items_needed:
                overflow = allowed_outputs[i] - items_needed
                return allowed_sizes[i], 1, f"{items_needed:>5} (+{overflow}) <= {allowed_sizes[i]}"
    
    if level > 1:
        return direct_solution
    return 0, 0, ""
    
def safe_get_int(value):
    try:
        return int(value)
    except:
        return 0

import re
from datetime import timedelta

def parse_time(time_str):
    """Parse time string in format D.HH:MM:SS, HH:MM:SS, or MM:SS"""
    try:
        if not time_str:
            return timedelta()
            
        pattern = r'^((((?P<day>\d+)\.)?(?P<hour>\d{1,2}):)?(?P<minute>\d{1,2}):)?(?P<second>\d{1,2})$'
        match = re.match(pattern, time_str)
        
        if not match:
            return timedelta()
            
        # Extract named groups, defaulting to 0 if None
        groups = match.groupdict()
        return timedelta(
            days=int(groups['day']) if groups['day'] else 0,
            hours=int(groups['hour']) if groups['hour'] else 0,
            minutes=int(groups['minute']) if groups['minute'] else 0,
            seconds=int(groups['second']) if groups['second'] else 0
        )
            
    except:
        return timedelta()

def format_time(td):
    """Format timedelta to human readable string"""
    total_seconds = int(td.total_seconds())
    days = total_seconds // (24 * 3600)
    remaining = total_seconds % (24 * 3600)
    hours = remaining // 3600
    remaining = remaining % 3600
    minutes = remaining // 60
    seconds = remaining % 60
    
    if days > 0:
        return f"{days}.{hours:02d}:{minutes:02d}:{seconds:02d}"
    elif hours > 0:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    elif minutes > 0:
        return f"{minutes}:{seconds:02d}"
    else:
        return f"0:{seconds:02d}"

class CalculatorGrid:
    def __init__(self, parent_frame):       
        self.frame = ttk.Frame(parent_frame, padding="5")
        self.max_merge_values = ["", "3", "5", "9", "17", "33", "65"]
            
        # Input Frame
        self.input_frame = ttk.Frame(self.frame, padding="5")
        self.input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Input Variables with trace
        self.target_level = tk.StringVar(value="10")
        self.target_level.trace_add("write", self.trigger_calculation)

        self.max_merge = tk.StringVar(value="17")       
        self.max_merge.trace_add("write", self.trigger_calculation)
        
        # Input Fields
        ttk.Label(self.input_frame, text="Target Level:").grid(row=0, column=0, padx=5)
        ttk.Entry(self.input_frame, textvariable=self.target_level, width=10).grid(row=0, column=1, padx=5)
        
        ttk.Label(self.input_frame, text="Max Merge Size:").grid(row=0, column=2, padx=5)
        ttk.Combobox(self.input_frame, textvariable=self.max_merge, values=self.max_merge_values, 
                     width=7, state="readonly").grid(row=0, column=3, padx=5)
                
        # Add completion label at the end of input fields
        self.completion_label = ttk.Label(self.input_frame, text="Completion: 0%")
        self.completion_label.grid(row=0, column=6, padx=15)
        
        # Results Frame
        self.results_frame = ttk.Frame(self.frame, padding="5")
        self.results_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.style = ttk.Style()
        self.style.theme_use('clam') # clam, alt, winnative
        self.style.configure("Error.TEntry", fieldbackground="red", foreground="white")
        self.style.configure("Normal.TEntry", fieldbackground="white", foreground="black")

        self.style.map('TCombobox', fieldbackground=[('readonly', 'white')])
        self.style.map('TCombobox', selectbackground=[('readonly', 'white')])
        self.style.map('TCombobox', selectforeground=[('readonly', 'black')])
    
        self.rows = []
        self.items_needed_vars = []
        self.efficient_vars = []
        self.name_vars = []
        self.time_vars = []
        self.on_hand_vars = []
        self.max_merge_vars = []
        self.nb_override_vars = []
        self.liquid_vars = []
        
        self.create_initial_rows()
        self.calculate()

    def create_initial_rows(self):
        headers = ["Level", "Items Needed", "Efficient", "Name", "Time to build", "On hand", 
                  "Max Merge", "Nb", "Needed <= Merge Combination", "Items required", "Liquid to use",
                  "Items / Time saved", "Items to create"]
            
        for i, header in enumerate(headers):
            ttk.Label(self.results_frame, text=header, font=('Arial', 8, 'bold')).grid(row=0, column=i, padx=2)
        
        # Create lists to store editable widgets by column
        name_entries = []
        time_entries = []
        on_hand_entries = []
        liquid_entries = []
        max_merge_entries = []
        nb_override_entries = []
        efficient_entries = []

        for i in range(15):
            row = []
            
            # Level (non-editable)
            level_label = ttk.Label(self.results_frame, text="", takefocus=0)
            level_label.grid(row=i+1, column=Column.LEVEL.value, padx=2)
            row.append(level_label)
    
            # Items Needed (editable only on the first line)
            if i == 0:
                # First row: make it editable
                items_needed_var = tk.StringVar(value="1")
                self.items_needed_vars.append(items_needed_var)
                items_needed_var.trace_add("write", self.trigger_calculation)
                items_needed = ttk.Entry(self.results_frame, textvariable=items_needed_var, width=3, justify='center')
                items_needed.grid(row=i+1, column=Column.ITEMS_NEEDED.value, padx=2)
                row.append(items_needed)
            else:
                # Other rows: keep as labels
                items_needed = ttk.Label(self.results_frame, text="", takefocus=0)
                items_needed.grid(row=i+1, column=Column.ITEMS_NEEDED.value, padx=2)
                row.append(items_needed)
            
            # Efficient checkbox
            efficient_var = tk.BooleanVar(value=False)
            self.efficient_vars.append(efficient_var)
            efficient_var.trace_add("write", self.trigger_calculation)
            efficient_check = ttk.Checkbutton(self.results_frame, variable=efficient_var)
            efficient_check.grid(row=i+1, column=Column.EFFICIENT.value, padx=2, sticky="w")
            efficient_entries.append(efficient_check)
            row.append(efficient_check)

            # Name (editable)
            name_var = tk.StringVar()
            self.name_vars.append(name_var)
            name_var.trace_add("write", self.trigger_calculation)
            name_entry = ttk.Entry(self.results_frame, textvariable=name_var, width=15)
            name_entry.grid(row=i+1, column=Column.NAME.value, padx=2)
            name_entries.append(name_entry)
            row.append(name_entry)
            
            # Time To Build (editable)
            time_var = tk.StringVar()
            self.time_vars.append(time_var)
            time_var.trace_add("write", self.trigger_calculation)
            time_entry = ttk.Entry(self.results_frame, textvariable=time_var, width=10, justify='right')
            time_entry.grid(row=i+1, column=Column.TIME_TO_BUILD.value, padx=2)
            time_entries.append(time_entry)
            row.append(time_entry)
            
            # On hand (editable)
            on_hand_var = tk.StringVar(value="0")
            self.on_hand_vars.append(on_hand_var)
            on_hand_var.trace_add("write", self.trigger_calculation)
            items_on_hand = ttk.Entry(self.results_frame, textvariable=on_hand_var, width=10)
            items_on_hand.grid(row=i+1, column=Column.ON_HAND.value, padx=2)
            on_hand_entries.append(items_on_hand)
            row.append(items_on_hand)
            
            # Max Merge (combobox)
            max_merge_var = tk.StringVar(value="")
            self.max_merge_vars.append(max_merge_var)
            max_merge_var.trace_add("write", self.trigger_calculation)
            max_merge = ttk.Combobox(self.results_frame, textvariable=max_merge_var, width=7, values=self.max_merge_values, state="readonly")
            max_merge.grid(row=i+1, column=Column.MAX_MERGE.value, padx=2)
            max_merge_entries.append(max_merge)
            row.append(max_merge)

            # Nb override (editable)
            nb_override_var = tk.StringVar(value="0")
            self.nb_override_vars.append(nb_override_var)
            nb_override_var.trace_add("write", self.trigger_calculation)
            nb_override = ttk.Entry(self.results_frame, textvariable=nb_override_var, width=5)
            nb_override.grid(row=i+1, column=Column.NB_OVERRIDE.value, padx=2)
            nb_override_entries.append(nb_override)
            row.append(nb_override)

            # Merge Combination (non-editable)
            merge_comb = ttk.Label(self.results_frame, text="", font=("Courier", "10"), takefocus=0)
            merge_comb.grid(row=i+1, column=Column.MERGE_COMBINATION.value, padx=2, sticky="w")
            row.append(merge_comb)
            
            # Items required (non-editable)
            items_required = ttk.Label(self.results_frame, text="", takefocus=0)
            items_required.grid(row=i+1, column=Column.ITEMS_REQUIRED.value, padx=2)
            row.append(items_required)
            
            # Liquid to use (editable)
            liquid_var = tk.StringVar(value="0")
            self.liquid_vars.append(liquid_var)
            liquid_var.trace_add("write", self.trigger_calculation)
            liquid = ttk.Entry(self.results_frame, textvariable=liquid_var, width=10, style="Normal.TEntry")
            liquid.grid(row=i+1, column=Column.LIQUID_TO_USE.value, padx=2)
            liquid_entries.append(liquid)
            row.append(liquid)

            # Liquid saving (non-editable)
            liquid_saving = ttk.Label(self.results_frame, text="", takefocus=0)
            liquid_saving.grid(row=i+1, column=Column.LIQUID_SAVING.value, padx=2)
            row.append(liquid_saving)
            
            # Items to create (non-editable)
            items_create = ttk.Label(self.results_frame, text="", takefocus=0)
            items_create.grid(row=i+1, column=Column.ITEMS_TO_CREATE.value, padx=2, sticky="w")
            row.append(items_create)
            
            self.rows.append(row)
        
        # Create ordered list of all entry controls for tab order
        tab_order = []
        tab_order.extend(efficient_entries)
        tab_order.extend(name_entries)
        tab_order.extend(time_entries)
        tab_order.extend(on_hand_entries)
        tab_order.extend(nb_override_entries)
        tab_order.extend(liquid_entries)
        
        # Set the tab order for all controls
        for i, entry in enumerate(tab_order):
            entry.lift()
            entry.configure(takefocus=True)
            if i < len(tab_order) - 1:
                next_widget = tab_order[i + 1]
            else:
                next_widget = tab_order[0]  # Loop back to first control
            entry.tk_focusNext = lambda next=next_widget: next


    def trigger_calculation(self, *args):
        current = self.calculate()
    
    def calculate(self, update_ui=True, min_liquid_level=0, use_items_on_hand=True):
        try:
            target_level = safe_get_int(self.target_level.get())
            global_max_merge = safe_get_int(self.max_merge.get())
            current_items_needed = safe_get_int(self.items_needed_vars[0].get())

            if update_ui:
                for row in self.rows:
                    for widget in row:
                        if isinstance(widget, ttk.Label):
                            widget.config(text="")
            
            current_level = target_level
            total_build_time = timedelta()
            items_to_create = 0
            last_current_items_needed = 0
            
            for i, row in enumerate(self.rows):
                if current_level < 1:
                    break
                
                if update_ui:
                    row[Column.LEVEL.value].config(text=str(current_level))
                    if i > 0:
                        row[Column.ITEMS_NEEDED.value].config(text=str(current_items_needed))
                    
                last_current_items_needed = current_items_needed

                try:
                    line_max_merge = safe_get_int(self.max_merge_vars[i].get())
                    max_merge_to_use = line_max_merge if line_max_merge > 0 else global_max_merge
                    nb_override = safe_get_int(self.nb_override_vars[i].get())
                except ValueError:
                    max_merge_to_use = global_max_merge
                    nb_override = 0

                if update_ui:
                    if self.efficient_vars[i].get():
                        current_items_needed = ((current_items_needed + max_merge_to_use - 1) // max_merge_to_use) * max_merge_to_use
                        row[Column.EFFICIENT.value].config(text=str(current_items_needed))
                    else:
                        row[Column.EFFICIENT.value].config(text="")
                
                items_on_hand = safe_get_int(self.on_hand_vars[i].get()) if use_items_on_hand else 0
                items_needed_after_onhand = max(0, current_items_needed - items_on_hand)

                if items_needed_after_onhand > 0:                                              
                    small_items, nb_merges, combination = calculate_small_items(
                        items_needed_after_onhand, 
                        global_max_merge if nb_override > 0 else max_merge_to_use, 
                        max_merge_to_use, nb_override if nb_override > 0 else 9999, 
                        current_level)
                    
                    liquid_to_use = safe_get_int(self.liquid_vars[i].get()) if current_level >= min_liquid_level else 0
                    liquid_to_use_effective = min(nb_merges, liquid_to_use)
                    
                    if update_ui:
                        # Show merge combination with new format
                        merge_text = f"{combination:<30}" if current_level > 1 else f"{items_needed_after_onhand:>5}"
                        row[Column.MERGE_COMBINATION.value].config(text=merge_text)
                        row[Column.ITEMS_REQUIRED.value].config(text=str(small_items))
                        
                        # Check if liquid is higher than number of merges
                        if liquid_to_use > nb_merges:
                            row[Column.LIQUID_TO_USE.value].config(style="Error.TEntry")
                        else:
                            row[Column.LIQUID_TO_USE.value].config(style="Normal.TEntry")
                           
                        # Get next item name
                        next_item_name = self.name_vars[i+1].get() if i+1 < len(self.rows) else ""
                        items_after_liquid = max(0, small_items - liquid_to_use_effective)
                        row[Column.ITEMS_TO_CREATE.value].config(
                            text=f"{items_after_liquid} {next_item_name}")
                    
                        # Calculate liquid savings
                        if liquid_to_use_effective > 0:
                            without_liquid, time_without_liquid = self.calculate(False, current_level + 1, False)
                            with_liquid, time_with_liquid = self.calculate(False, current_level, False)
                            savings = without_liquid - with_liquid
                            row[Column.LIQUID_SAVING.value].config(text=f"{savings} / {format_time(time_without_liquid - time_with_liquid)}")
                            
                    current_items_needed = max(0, small_items - liquid_to_use_effective)
                else:
                    current_items_needed = 0
                                
                current_level -= 1
            
                # Calculate build time
                if current_items_needed > 0:
                    time_str = self.time_vars[i].get()
                    if time_str:
                        build_time = parse_time(time_str)
                        total_build_time += build_time * current_items_needed

            if update_ui:
                without_items, total_time = self.calculate(False, 0, False)
                completion = (without_items - last_current_items_needed) * 100 / without_items if without_items > 0 else 0
                self.completion_label.config(
                    text=f"Completion: {completion:.1f}% - Total build time: {format_time(total_build_time)}")
            
            return last_current_items_needed, total_build_time
                
        except ValueError:
            pass
            
    def get_data(self):
        data = {
            'target_level': self.target_level.get(),
            'max_merge': self.max_merge.get(),
            'items_needed': self.items_needed_vars[0].get(),
            'rows': []
        }
        
        for i in range(len(self.rows)):
            row_data = {
                'efficient': self.efficient_vars[i].get(),
                'name': self.name_vars[i].get(),
                'time': self.time_vars[i].get(),
                'on_hand': self.on_hand_vars[i].get(),
                'liquid': self.liquid_vars[i].get(),
                'max_merge': self.max_merge_vars[i].get(),
                'nb_override': self.nb_override_vars[i].get()
             }
            data['rows'].append(row_data)
            
        return data
    
    def load_data(self, data):
        self.target_level.set(data.get('target_level', '10'))
        self.max_merge.set(data.get('max_merge', '17'))
        self.items_needed_vars[0].set(data.get('items_needed', '5'))
   
        rows_data = data.get('rows', [])
        for i, row_data in enumerate(rows_data):
            if i < len(self.rows):
                self.efficient_vars[i].set(row_data.get('efficient', False))
                self.name_vars[i].set(row_data.get('name', ''))
                self.time_vars[i].set(row_data.get('time', ''))
                self.on_hand_vars[i].set(row_data.get('on_hand', '0'))
                self.liquid_vars[i].set(row_data.get('liquid', '0'))
                self.max_merge_vars[i].set(row_data.get('max_merge', ''))
                self.nb_override_vars[i].set(row_data.get('nb_override', ''))

class BuildingCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Mergest Kingdom Planner")
       
        # Register window closing event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Create main horizontal split
        self.main_paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True)
        
        # Left side - Sheet selector
        self.left_frame = ttk.Frame(self.main_paned, padding="5")
        self.main_paned.add(self.left_frame)
        
        # Add sheet button
        self.add_button = ttk.Button(self.left_frame, text="Add Sheet", command=self.add_sheet)
        self.add_button.pack(fill=tk.X, pady=(0, 5))
        
        # Delete sheet button
        self.delete_button = ttk.Button(self.left_frame, text="Delete Sheet", command=self.delete_sheet)
        self.delete_button.pack(fill=tk.X)
        
        # Sheet treeview
        self.sheet_tree = ttk.Treeview(self.left_frame, selectmode='browse', show='tree')
        self.sheet_tree.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.sheet_tree.bind('<<TreeviewSelect>>', self.on_sheet_select)
        
        # Add right-click menu for renaming
        self.sheet_menu = tk.Menu(self.root, tearoff=0)
        self.sheet_menu.add_command(label="Rename", command=self.rename_sheet)
        self.sheet_menu.add_separator()
        self.sheet_menu.add_command(label="Move Up", command=self.move_sheet_up)
        self.sheet_menu.add_command(label="Move Down", command=self.move_sheet_down)
        self.sheet_tree.bind("<Button-3>", self.show_sheet_menu)
        
        # Save buttons
        self.save_button = ttk.Button(self.left_frame, text="Save", command=self.save_data)
        self.save_button.pack(fill=tk.X, pady=(5, 0))
        
        # Right side - Content frame
        self.right_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.right_frame)
        
        # Dictionary to store calculator grids
        self.calculators = {}
        self.sheet_counter = 1
        
        # Create initial sheet
        self.add_sheet()
        
        # Load data if exists
        self.load_data()
        
    def on_closing(self):
        """Handler for window closing event"""
        self.save_data()  # Save before closing
        self.root.destroy()
        
    def show_sheet_menu(self, event):
        selected = self.sheet_tree.selection()
        if selected:
            item = selected[0]
            # Enable/disable move options based on position
            prev = self.sheet_tree.prev(item)
            next = self.sheet_tree.next(item)
            
            self.sheet_menu.entryconfig("Move Up", state='normal' if prev else 'disabled')
            self.sheet_menu.entryconfig("Move Down", state='normal' if next else 'disabled')
            
            self.sheet_menu.post(event.x_root, event.y_root)
        
    def rename_sheet(self):
        selected = self.sheet_tree.selection()
        if not selected:
            return
        
        sheet_id = selected[0]
        old_name = self.sheet_tree.item(sheet_id)['text']
        new_name = simpledialog.askstring("Rename Sheet", "Enter new name:", initialvalue=old_name)
        
        if new_name:
            self.sheet_tree.item(sheet_id, text=new_name)
            
    def move_sheet_up(self):
        selected = self.sheet_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        prev = self.sheet_tree.prev(item)
        if prev:
            # Get the current index and calculate new index
            index = self.sheet_tree.index(item)
            self.sheet_tree.move(item, '', index - 1)

    def move_sheet_down(self):
        selected = self.sheet_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        next = self.sheet_tree.next(item)
        if next:
            # Get the current index and calculate new index
            index = self.sheet_tree.index(item)
            self.sheet_tree.move(item, '', index + 1)
        
    def save_data(self):
        data = {
            'sheets': {}
        }
        
        # Save sheets in the order they appear in sheet_tree
        for sheet_id in self.sheet_tree.get_children():
            sheet_name = self.sheet_tree.item(sheet_id)['text']
            calculator = self.calculators[sheet_id]
            data['sheets'][sheet_name] = calculator.get_data()
        
        # Save to file
        try:
            with open('MergestKingdomPlanner-Data.json', 'w') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to save data: {str(e)}")
        
    def load_data(self):
        filename = 'MergestKingdomPlanner-Data.json'
        default_filename = 'MergestKingdomPlanner-Data.default.json'
        
        # Try to load user file first, fall back to default if not found
        if not os.path.exists(filename):
            if not os.path.exists(default_filename):
                return  # No files found
            filename = default_filename
        
        try:
            with open(filename, 'r') as f:
                data = json.load(f)
            
            # Clear existing sheets
            for sheet_id in self.sheet_tree.get_children():
                self.sheet_tree.delete(sheet_id)
            self.calculators.clear()
            
            # Load each sheet
            for sheet_name, sheet_data in data.get('sheets', {}).items():
                # Create new sheet
                sheet_id = self.sheet_tree.insert('', 'end', text=sheet_name)
                calculator = CalculatorGrid(self.right_frame)
                self.calculators[sheet_id] = calculator
                calculator.load_data(sheet_data)
            
            # Select first sheet if exists
            if self.sheet_tree.get_children():
                first_sheet = self.sheet_tree.get_children()[0]
                self.sheet_tree.selection_set(first_sheet)
                self.show_sheet(first_sheet)
            else:
                # If no sheets were loaded, create a default one
                self.add_sheet()
                
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to load data from {filename}: {str(e)}")
            # Create default sheet if loading failed
            self.add_sheet() 
            
    def add_sheet(self):
        sheet_name = f"Sheet {self.sheet_counter}"
        self.sheet_counter += 1
        
        # Add to treeview
        sheet_id = self.sheet_tree.insert('', 'end', text=sheet_name)
        
        # Create new calculator grid
        calculator = CalculatorGrid(self.right_frame)
        self.calculators[sheet_id] = calculator
        
        # Select the new sheet
        self.sheet_tree.selection_set(sheet_id)
        self.show_sheet(sheet_id)

    def delete_sheet(self):
        selected = self.sheet_tree.selection()
        if not selected:
            return
        
        sheet_id = selected[0]
        sheet_name = self.sheet_tree.item(sheet_id)['text']
        
        # Show confirmation dialog
        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete '{sheet_name}'?",
            icon='warning'
        )
        
        if not confirm:
            return
        
        # Remove from dictionary and treeview
        if sheet_id in self.calculators:
            self.calculators[sheet_id].frame.grid_forget()
            del self.calculators[sheet_id]
        self.sheet_tree.delete(sheet_id)
        
        # If there are no sheets left, create a new one
        if not self.calculators:
            self.add_sheet()
        else:
            # Select the first remaining sheet
            first_sheet = self.sheet_tree.get_children()[0]
            self.sheet_tree.selection_set(first_sheet)
            self.show_sheet(first_sheet)

    def show_sheet(self, sheet_id):
        # Hide all sheets
        for calculator in self.calculators.values():
            calculator.frame.grid_forget()
        
        # Show selected sheet
        if sheet_id in self.calculators:
            self.calculators[sheet_id].frame.grid(row=0, column=0, sticky='nsew')

    def on_sheet_select(self, event):
        selected = self.sheet_tree.selection()
        if selected:
            self.show_sheet(selected[0])

def main():
    # for merge_size in [3, 5, 9, 17, 33, 65]:
    #     for items_needed in range(130):
    #         x, y = calculate_small_items(items_needed, merge_size, 5)
    #         print(f"{merge_size},{items_needed} => {x} ({y})")

    root = tk.Tk()
    app = BuildingCalculator(root)
    root.geometry("1400x600")  # Set initial window size
    root.mainloop()

if __name__ == "__main__":
    main()
