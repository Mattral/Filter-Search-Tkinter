import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pandastable import Table
import pandas as pd
import numpy as np

class WarehouseFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Warehouse Data Filter Tool")
        
        self.file_path = ""
        
        # Button to load the Excel file
        self.load_button = ttk.Button(root, text="Upload Excel File", command=self.load_file)
        self.load_button.pack(pady=(10, 20))
        
        self.filters_frame = ttk.LabelFrame(root, text="Filters")
        self.filters_frame.pack(padx=10, pady=(10, 20), fill="x")
        
        self.filter_widgets = {}
        self.confirmed_selections = {}
        
        # Action buttons
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

        # Table for displaying results
        self.results_frame = ttk.LabelFrame(root, text="Results")
        self.results_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.results_table = None
        self.df = pd.DataFrame()

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.df = pd.read_excel(self.file_path, engine='openpyxl')
            self.initialize_filters()

    def initialize_filters(self):
        self.clear_filters()
        for column in self.df.columns:
            unique_values = self.df[column].dropna().unique()
            if len(unique_values) > 20:
                self.add_text_filter(column, 0, len(self.filter_widgets))
            else:
                self.add_multichoice_filter(column, unique_values, 0, len(self.filter_widgets))

    def add_text_filter(self, column, filter_row, filter_column):
        label = ttk.Label(self.filters_frame, text=column)
        label.grid(row=filter_row, column=filter_column, padx=5, pady=5)
        entry = ttk.Entry(self.filters_frame)
        entry.grid(row=filter_row + 1, column=filter_column, padx=5, pady=5)
        self.filter_widgets[column] = entry

    def add_multichoice_filter(self, column, options, filter_row, filter_column):
        label = ttk.Label(self.filters_frame, text=column)
        label.grid(row=filter_row, column=filter_column, padx=5, pady=5)
        
        frame = ttk.Frame(self.filters_frame)
        frame.grid(row=filter_row + 1, column=filter_column, padx=5, pady=5)
        
        listbox = tk.Listbox(frame, selectmode="multiple", height=6)
        for option in options:
            listbox.insert(tk.END, option)
        listbox.pack(side=tk.LEFT, fill="y")
        
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        listbox.config(yscrollcommand=scrollbar.set)

        confirm_button = ttk.Button(frame, text="Confirm", command=lambda col=column, lb=listbox: self.confirm_selection(col, lb))
        confirm_button.pack(pady=5)

        confirmed_label = ttk.Label(self.filters_frame, text="")
        confirmed_label.grid(row=filter_row + 2, column=filter_column)
        self.confirmed_selections[column] = confirmed_label
        
        self.filter_widgets[column] = listbox

    def confirm_selection(self, column, listbox):
        selected_values = [listbox.get(i) for i in listbox.curselection()]
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
        for label in self.confirmed_selections.values():
            label.config(text="")

    def search_data(self):
        if self.df.empty:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        filtered_df = self.df.copy()
        for column, widget in self.filter_widgets.items():
            if isinstance(widget, ttk.Entry):
                value = widget.get().strip()
                if value:
                    filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(value, case=False, na=False)]
            elif isinstance(widget, tk.Listbox):
                confirmed_values = self.confirmed_selections[column].cget("text").split(", ")
                if confirmed_values and confirmed_values[0]:
                    filtered_df = filtered_df[filtered_df[column].astype(str).isin(confirmed_values)]

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
