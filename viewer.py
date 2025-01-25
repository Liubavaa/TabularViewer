import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
# import pyreadstat


class DataViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Viewer")
        self.root.geometry("900x600")

        # Variables
        self.dataframes = {}
        self.current_dataframe = None

        # UI Elements
        self.file_label = tk.Label(root, text="No file selected")
        self.file_label.pack(pady=5)

        # In case selected file is xlsx
        self.sheet_selector_frame = tk.Frame(root)
        self.sheet_label = tk.Label(self.sheet_selector_frame, text="Select Sheet:")
        self.sheet_label.pack(side=tk.LEFT, padx=5)
        self.sheet_selector = ttk.Combobox(self.sheet_selector_frame, state="readonly")
        self.sheet_selector.pack(side=tk.LEFT, padx=5)
        self.sheet_selector.bind("<<ComboboxSelected>>", self.display_selected_file)

        self.tree = ttk.Treeview(root)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree_scroll_y = ttk.Scrollbar(self.tree, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x = ttk.Scrollbar(self.tree, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(yscroll=self.tree_scroll_y.set, xscroll=self.tree_scroll_x.set)

        self.load_button = tk.Button(root, text="Load File", command=self.load_file)
        self.load_button.pack(pady=5)

        self.summary_button = tk.Button(root, text="Show Summary", command=self.show_summary)
        self.summary_button.pack(pady=5)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Excel Files", "*.xlsx"),
            ("SAS Files", "*.xpt *.sas7bdat"),
            ("CSV Files", "*.csv"),
            ("All Files", "*.*")
        ])

        if not file_path:
            return

        try:
            # Hide the sheet selector initially
            self.sheet_selector_frame.pack_forget()

            # Load data based on file type
            if file_path.endswith(('.xlsx', '.xls')):
                self.dataframes = pd.read_excel(file_path, sheet_name=None)
                self.sheet_selector["values"] = list(self.dataframes.keys())
                if self.dataframes:
                    self.sheet_selector.current(0)
                    self.sheet_selector_frame.pack(pady=5)
            elif file_path.endswith('.xpt'):
                self.dataframes = {"File": pd.read_sas(file_path, format="xport", encoding="utf-8")}
            elif file_path.endswith('.sas7bdat'):
                # df, _ = pyreadstat.read_sas7bdat(file_path)
                # self.dataframes = {"File": df}
                self.dataframes = {"File": pd.read_sas(file_path, encoding="utf-8")}
            elif file_path.endswith('.csv'):
                with open(file_path, 'r') as csvfile:
                    delimiter = csv.Sniffer().sniff(csvfile.readline()).delimiter
                delimiter = '$' if delimiter == ' ' else delimiter
                df = pd.read_csv(file_path, sep=delimiter)
                self.dataframes = {"File": df}
            else:
                messagebox.showerror("Unsupported File", "Only Excel, SAS, CSV and XPT files are supported.")
                return

            self.file_label.config(text=f"Loaded: {file_path}")
            self.process_df()
            self.display_selected_file()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def process_df(self):
        for key, df in self.dataframes.items():
            df = df.apply(pd.to_numeric, errors='coerce').fillna(df)
            df = df.convert_dtypes()
            self.dataframes[key] = df

            # Process dates (seems to be unnecessary)
            # data = [pd.to_datetime(df[x]) if df[x].astype(str).str.match(r'\d{4}-\d{2}-\d{2}').all() else df[x] for x in
            #         df.columns]
            # df = pd.concat(data, axis=1, keys=[s.name for s in data])

    def display_selected_file(self, event=None):
        if len(self.dataframes.keys()) > 1:
            sheet_name = self.sheet_selector.get()
            if not sheet_name or sheet_name not in self.dataframes:
                return
            self.current_dataframe = self.dataframes[sheet_name]
        else:
            self.current_dataframe = self.dataframes["File"]

        # Clear tree
        for col in self.tree.get_children():
            self.tree.delete(col)

        # Add new data
        df = self.current_dataframe
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, minwidth=100, width=150)

        # Insert rows into TreeView
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def show_summary(self):
        if self.current_dataframe is None:
            messagebox.showwarning("No Data", "No data is loaded to display the summary.")
            return

        summary_window = tk.Toplevel(self.root)
        summary_window.title("Summary")
        summary_window.geometry("500x400")

        # Generate Summary
        df = self.current_dataframe
        stats = df.describe().transpose()

        # Display in Text Widget
        text_widget = tk.Text(summary_window, wrap="none")
        text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(summary_window, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)

        # Populate Text Widget
        text_widget.insert(tk.END, stats.to_string())


# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = DataViewerApp(root)
    root.mainloop()
