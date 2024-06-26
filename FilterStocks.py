import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pandastable import Table
import pandas as pd

class WarehouseFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Warehouse Filter Search Tool")
        
        self.file_path = ""
        
        # Add a button to load the Excel file
        self.load_button = ttk.Button(root, text="Upload Excel File", command=self.load_file)
        self.load_button.pack(pady=(10, 20))
        
        self.filters_frame = ttk.LabelFrame(root, text="Filters")
        self.filters_frame.pack(padx=10, pady=(10, 20), fill="x")
        
        self.filter_widgets = {}
        self.confirmed_selections = {}
        
        # Add buttons to perform the search and reset filters
        self.buttons_frame = ttk.Frame(root)
        self.buttons_frame.pack(pady=(10, 20))
        
        self.search_button = ttk.Button(self.buttons_frame, text="Search", command=self.search_data)
        self.search_button.grid(row=0, column=0, padx=5)
        
        self.reset_button = ttk.Button(self.buttons_frame, text="Reset Filters", command=self.reset_filters)
        self.reset_button.grid(row=0, column=1, padx=5)

        self.download_button = ttk.Button(self.buttons_frame, text="Download Results", command=self.download_results)
        self.download_button.grid(row=0, column=2, padx=5)

        self.clear_button = ttk.Button(self.buttons_frame, text="Clear Results", command=self.clear_results)
        self.clear_button.grid(row=0, column=3, padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Table to display results
        self.results_frame = ttk.LabelFrame(root, text="Results")
        self.results_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.results_table = None

        # Initialize dataframe
        self.df = pd.DataFrame()

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.df = pd.read_excel(self.file_path, engine='openpyxl')
            self.initialize_filters()

    def initialize_filters(self):
        self.clear_filters()
        
        # ValA filter
        self.add_dropdown_filter("ValA", ["8100", "8200"], filter_row=0, filter_column=0)
        
        # Material filter
        self.add_text_filter("Material", filter_row=0, filter_column=1)
        
        # Material description filter
        self.add_text_filter("Material description", filter_row=1, filter_column=0)
        
        # Long text filter
        self.add_text_filter("Long text", filter_row=1, filter_column=1)
        
        # L/O filter
        self.add_multichoice_filter("L/O", list(self.df["L/O"].dropna().unique()), filter_row=2, filter_column=0)
        
        # Manufacturer name filter
        self.add_multichoice_filter("Manufacturer name", list(self.df["Manufacturer name"].dropna().unique()), filter_row=2, filter_column=1)
        
        # MPN filter
        self.add_text_filter("MPN", filter_row=5, filter_column=0)
        
        # Mfr filter with both dropdown and text input
        self.add_combined_filter("Mfr", list(self.df["Mfr"].dropna().unique()), filter_row=5, filter_column=1)
        
        # BUn filter
        self.add_multichoice_filter("BUn", list(self.df["BUn"].dropna().unique()), filter_row=5, filter_column=0)

    def add_dropdown_filter(self, column, options, filter_row, filter_column):
        label = ttk.Label(self.filters_frame, text=column)
        label.grid(row=filter_row, column=filter_column*2, padx=5, pady=(10, 20))
        variable = tk.StringVar(self.root)
        variable.set("Select")
        dropdown = ttk.Combobox(self.filters_frame, textvariable=variable, values=options, state='readonly')
        dropdown.grid(row=filter_row, column=filter_column*2 + 1, padx=5, pady=(10, 20))
        self.filter_widgets[column] = variable

    def add_text_filter(self, column, filter_row, filter_column):
        label = ttk.Label(self.filters_frame, text=column)
        label.grid(row=filter_row, column=filter_column*2, padx=5, pady=(10, 20))
        entry = ttk.Entry(self.filters_frame)
        entry.grid(row=filter_row, column=filter_column*2 + 1, padx=5, pady=(10, 20))
        self.filter_widgets[column] = entry

    def add_multichoice_filter(self, column, options, filter_row, filter_column):
        label = ttk.Label(self.filters_frame, text=column)
        label.grid(row=filter_row, column=filter_column*2, padx=5, pady=(10, 20))
        
        frame = ttk.Frame(self.filters_frame)
        frame.grid(row=filter_row, column=filter_column*2 + 1, padx=5, pady=(10, 20))
        
        listbox = tk.Listbox(frame, selectmode="multiple", height=5)
        for option in options:
            listbox.insert(tk.END, option)
        listbox.pack(side=tk.LEFT, fill="y")
        
        scrollbar = ttk.Scrollbar(frame, orient="vertical")
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        listbox.config(yscrollcommand=scrollbar.set)

        confirm_button = ttk.Button(frame, text="Confirm", command=lambda col=column, lb=listbox: self.confirm_selection(col, lb))
        confirm_button.pack(pady=5)

        confirmed_label = ttk.Label(self.filters_frame, text="")
        confirmed_label.grid(row=filter_row + 1, column=filter_column*2, columnspan=2)
        self.confirmed_selections[column] = confirmed_label
        
        self.filter_widgets[column] = listbox

    def add_combined_filter(self, column, options, filter_row, filter_column):
        label = ttk.Label(self.filters_frame, text=column)
        label.grid(row=filter_row, column=filter_column*2, padx=5, pady=(10, 20))
        
        combined_frame = ttk.Frame(self.filters_frame)
        combined_frame.grid(row=filter_row, column=filter_column*2 + 1, padx=5, pady=(10, 20))
        
        variable = tk.StringVar(self.root)
        variable.set("Select")
        dropdown = ttk.Combobox(combined_frame, textvariable=variable, values=options, state='readonly')
        dropdown.pack(side=tk.LEFT, fill="x", expand=True)
        
        entry = ttk.Entry(combined_frame)
        entry.pack(side=tk.RIGHT, fill="x", expand=True)
        
        self.filter_widgets[column] = (variable, entry)

    def confirm_selection(self, column, listbox):
        selected_values = [listbox.get(i) for i in listbox.curselection()]
        if selected_values:
            current_selections = self.confirmed_selections[column].cget("text")
            if current_selections:
                selected_values = list(set(current_selections.split(", ") + selected_values))
            self.confirmed_selections[column].config(text=", ".join(selected_values))

    def clear_filters(self):
        for widget in self.filters_frame.winfo_children():
            widget.destroy()
        self.filter_widgets.clear()
        self.confirmed_selections.clear()

    def reset_filters(self):
        for key, widget in self.filter_widgets.items():
            if isinstance(widget, tk.StringVar):
                widget.set("Select")
            elif isinstance(widget, ttk.Entry):
                widget.delete(0, tk.END)
            elif isinstance(widget, tk.Listbox):
                widget.selection_clear(0, tk.END)
            elif isinstance(widget, tuple):
                widget[0].set("Select")
                widget[1].delete(0, tk.END)
        
        for label in self.confirmed_selections.values():
            label.config(text="")

    def search_data(self):
        if self.df.empty:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        filtered_df = self.df.copy()
        
        for column, widget in self.filter_widgets.items():
            if isinstance(widget, tk.StringVar):
                value = widget.get()
                if value != "Select":
                    filtered_df = filtered_df[filtered_df[column].astype(str) == value]
            elif isinstance(widget, ttk.Entry):
                value = widget.get().strip()
                if value:
                    filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(value, case=False, na=False)]
            elif isinstance(widget, tk.Listbox):
                confirmed_values = self.confirmed_selections[column].cget("text").split(", ")
                if confirmed_values and confirmed_values[0]:
                    filtered_df = filtered_df[filtered_df[column].astype(str).isin(confirmed_values)]
            elif isinstance(widget, tuple):
                dropdown_value = widget[0].get()
                text_value = widget[1].get().strip()
                if dropdown_value != "Select":
                    filtered_df = filtered_df[filtered_df[column].astype(str) == dropdown_value]
                if text_value:
                    filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(text_value, case=False, na=False)]

        self.progress["maximum"] = len(filtered_df)
        for i in range(len(filtered_df)):
            self.progress["value"] = i+1
            self.root.update_idletasks()

        self.display_results(filtered_df)

    def display_results(self, filtered_df):
        if self.results_table:
            self.results_table.destroy()
        
        if filtered_df.empty:
            messagebox.showinfo("No results", "No matching records found.")
        else:
            self.results_table = Table(self.results_frame, dataframe=filtered_df, showtoolbar=True, showstatusbar=True)
            self.results_table.show()

    def download_results(self):
        if self.results_table is None:
            messagebox.showerror("Error", "No results to download.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if save_path:
            filtered_df = self.results_table.model.df
            if save_path.endswith(".xlsx"):
                filtered_df.to_excel(save_path, index=False)
            else:
                filtered_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Results successfully saved to {save_path}")

    def clear_results(self):
        if self.results_table:
            self.results_table.destroy()
            self.results_table = None

if __name__ == "__main__":
    root = tk.Tk()
    app = WarehouseFilterApp(root)
    root.mainloop()
