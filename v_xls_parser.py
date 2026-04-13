import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import *
from tkinter import ttk
import csv
import numpy as np
from tkinterdnd2 import TkinterDnD, DND_FILES
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading
import matplotlib
import io
from tkinter import messagebox
import os
from tkinter import filedialog

matplotlib.use('TkAgg')


def clean_csv_headers(csv_file_path, output_file_path):
    df = pd.read_csv(csv_file_path)
    first_header = df.columns[0]
    new_columns = [first_header] + [""] * (len(df.columns) - 1)
    df.columns = new_columns
    df.to_csv(output_file_path, index=False)


def xls_to_csv_conversion(file_path):
    # FIX: Use the correct engine based on file extension, and use it consistently.
    if file_path.endswith('.xlsx'):
        fengine = 'openpyxl'
    elif file_path.endswith('.xls'):
        fengine = 'xlrd'
    else:
        raise ValueError("Unsupported file format. Only .xls and .xlsx are supported.")

    try:
        xls = pd.ExcelFile(file_path, engine=fengine)  # FIX: was hardcoded to openpyxl
        sheet_names = xls.sheet_names
        if not sheet_names:
            raise ValueError("No sheets found in the Excel file.")
        excel_file = pd.read_excel(file_path, sheet_name=sheet_names[0], engine=fengine)
        excel_file.to_csv("tamp_output.csv", index=False)
        clean_csv_headers("tamp_output.csv", "tamp_output.csv")
        return "tamp_output.csv"
    except Exception as e:
        raise ValueError(f"Excel to CSV conversion failed: {str(e)}")


def func2(csvfilename):
    df = pd.read_csv(csvfilename)
    columns = df.columns.tolist()

    num_plots = len(columns) - 1
    if num_plots < 1:
        return None

    fig, axes = plt.subplots(1, num_plots, figsize=(20, 4))

    # FIX: When there's only one plot, plt.subplots returns a single Axes, not a list.
    if num_plots == 1:
        axes = [axes]

    for i in range(num_plots):
        ylabel = columns[i + 1].split(" ")
        axes[i].bar(df[columns[0]], df[columns[i + 1]])
        axes[i].set_title(columns[i + 1])
        axes[i].set_ylabel(ylabel[0])
        axes[i].tick_params(labelrotation=45, axis='x')
        # FIX: Check column names as a list membership check using `any`, and apply per-axis.
        if any(kw in col for col in columns for kw in ("Variance", "Goal")):
            axes[i].set_ylim(0, 100)

    plt.tight_layout()
    plt.close()
    return fig


def get_row_count(csv_file_path):
    with open(csv_file_path, 'r') as file:
        count = sum(1 for _ in file)
    return count


def create_data2(csv_file_path):
    # FIX: Rewritten entirely. The original deduplication logic was broken —
    # it compared raw unsplit string rows, meaning the blank-row check almost
    # never fired correctly. This version strips newlines and properly skips
    # runs of fully-blank rows while keeping all data rows.
    data2 = []
    prev_blank = False

    with open(csv_file_path, 'r') as file:
        for raw_line in file:
            stripped = raw_line.rstrip('\n').rstrip('\r')
            cells = [cell.strip() for cell in stripped.split(',')]
            is_blank = all(c == '' for c in cells)

            if is_blank:
                if not prev_blank:
                    # Keep one blank row as a table separator
                    data2.append(cells)
                prev_blank = True
            else:
                data2.append(cells)
                prev_blank = False

    return data2


def mark_df_as_seen(marked_cells, starting_i, starting_j, height, width):
    for row in range(starting_i, starting_i + height):
        for column in range(starting_j, starting_j + width):
            marked_cells[(row, column)] = True


def try_convert_to_numeric(cell):
    """
    Try to convert a cell value to numeric.
    Handles percentages, decimals, integers, and invalid values.
    Returns the numeric value or the original string if conversion fails.
    """
    if isinstance(cell, (int, float)):
        return cell
    
    cell_str = str(cell).strip()
    
    # Handle percentage signs
    if cell_str.endswith("%"):
        try:
            return float(cell_str[:-1])
        except ValueError:
            return cell_str
    
    # Try direct numeric conversion
    try:
        if '.' in cell_str:
            return float(cell_str)
        else:
            return int(cell_str)
    except ValueError:
        return cell_str


def create_df_no_title(darray, i, j, marked_cells):
    """Parse a table where the first row IS the header (no separate title row above)."""
    title = darray[i][j]
    starting_i = i
    starting_j = j
    width = 0
    height = 0

    # Measure width along the header row
    while j < len(darray[starting_i]) and darray[starting_i][j] != "":
        width += 1
        j += 1

    # Measure height downward
    while i < len(darray) and starting_j < len(darray[i]) and darray[i][starting_j] != "":
        height += 1
        i += 1

    if height < 2 or width < 1:
        return None

    matrix = []
    for g in range(starting_i, starting_i + height):
        temp_list = []
        for f in range(starting_j, starting_j + width):
            # FIX: guard against rows shorter than expected
            cell = darray[g][f] if f < len(darray[g]) else ""
            # FIX: Convert ALL numeric-looking values to numbers, not just percentages
            temp_list.append(try_convert_to_numeric(cell))
        matrix.append(temp_list)

    df = pd.DataFrame(matrix)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    
    # FIX: Convert string columns to numeric where possible
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
    
    mark_df_as_seen(marked_cells, starting_i, starting_j, height, width)
    return {title: df}


def create_df_yes_title(darray, i, j, marked_cells):
    """Parse a table that has a title cell above the header row."""
    title = darray[i][j]
    marked_cells[(i, j)] = True
    i += 1

    if i >= len(darray) or j >= len(darray[i]) or darray[i][j] == "":
        return None

    starting_i = i
    starting_j = j
    width = 0
    height = 0

    while j < len(darray[starting_i]) and darray[starting_i][j] != "":
        width += 1
        j += 1

    while i < len(darray) and starting_j < len(darray[i]) and darray[i][starting_j] != "":
        height += 1
        i += 1

    if height < 2 or width < 1:
        return None

    matrix = []
    for g in range(starting_i, starting_i + height):
        temp_list = []
        for f in range(starting_j, starting_j + width):
            cell = darray[g][f] if f < len(darray[g]) else ""
            # FIX: Convert ALL numeric-looking values to numbers, not just percentages
            temp_list.append(try_convert_to_numeric(cell))
        matrix.append(temp_list)

    df = pd.DataFrame(matrix)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    
    # FIX: Convert string columns to numeric where possible
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
    
    mark_df_as_seen(marked_cells, starting_i - 1, starting_j, height + 1, width)
    return {title: df}


def parse_csv_file(list_of_dfs, csv_file_path, marked_cells):
    data2 = create_data2(csv_file_path)
    return_value = list(list_of_dfs)

    for row in range(len(data2)):
        for column in range(len(data2[row])):
            if (row, column) in marked_cells:
                continue

            cell = data2[row][column]
            # FIX: strip newlines from cells before comparing
            cell = cell.strip()

            if cell == "":
                continue

            # Try to parse table WITH a title (current cell is title, next cell starts headers)
            result = create_df_yes_title(data2, row, column, marked_cells)
            if result:
                return_value.extend([result])
                continue

            # Try to parse table WITHOUT a title (current cell is first header)
            result = create_df_no_title(data2, row, column, marked_cells)
            if result:
                return_value.extend([result])

    return return_value


def QuickMake(file_path):
    if file_path.endswith('.csv'):
        temp_csv = file_path
    else:
        temp_csv = xls_to_csv_conversion(file_path)

    a = parse_csv_file([], temp_csv, {})
    all_figures = []

    # FIX: was `range(len(a) - 1)` which silently dropped the last table.
    for i in range(len(a)):
        try:
            for key in a[i].keys():
                df = a[i][key]

                # Skip tables that have no numeric data to plot
                numeric_cols = df.select_dtypes(include='number').columns.tolist()
                if not numeric_cols:
                    print(f"Skipping table '{key}' (row {i}): No numeric columns found")
                    continue

                new_data = [df.columns.tolist()]
                for j in range(len(df)):
                    new_data.append(df.iloc[j].tolist())

                with open("tamp_output.csv", 'w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerows(new_data)

                fig = func2("tamp_output.csv")
                if fig is not None:
                    all_figures.append(fig)
                else:
                    print(f"Warning: func2() returned None for table '{key}'")

        except Exception as e:
            # FIX: replaced silent `except: x = 2` with an informative print
            print(f"Skipping table #{i} due to error: {e}")

    if not all_figures:
        print("WARNING: No plottable tables found in the file!")
    
    return all_figures


class FileProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Plot Viewer")
        self.root.geometry("1000x700")
        self.file_queue = []
        self.figures = []   # FIX: initialize figures list here, not only in process_queue
        self.last_save_dir = None
        self.create_widgets()

    def create_widgets(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        control_panel = tk.Frame(self.main_frame, height=100, bg="#f0f0f0")
        control_panel.pack(fill=tk.X, pady=5)

        self.drop_label = tk.Label(control_panel, text="Drop files here",
                                   bg="lightgray", width=50, height=2,
                                   font=('Arial', 10))
        self.drop_label.pack(pady=5, padx=10, fill=tk.X)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.add_to_queue)

        self.process_btn = tk.Button(control_panel, text="Generate Plots",
                                     command=self.process_queue,
                                     bg="#4CAF50", fg="white")
        self.process_btn.pack(pady=5, padx=10, ipadx=20)

        self.save_all_btn = tk.Button(control_panel, text="Save All Plots as PDF",
                                      command=self.save_all_plots,
                                      bg="#2196F3", fg="white")
        self.save_all_btn.pack(pady=5, padx=10, ipadx=20, side=tk.RIGHT)

        self.queue_status = tk.Label(control_panel, text="Files ready: 0")
        self.queue_status.pack()

        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def add_to_queue(self, event):
        files = self._parse_dropped_files(event.data)
        for file in files:
            if file not in self.file_queue:
                self.file_queue.append(file)
        self.queue_status.config(text=f"Files ready: {len(self.file_queue)}")
        messagebox.showinfo("Files Added", f"Added {len(files)} file(s) to queue")

    def _parse_dropped_files(self, data):
        try:
            if isinstance(data, (list, tuple)):
                return [f.strip('{}') for f in data]
            if data.startswith('{') and data.endswith('}'):
                return [data.strip('{}')]
            if os.path.exists(data):
                return [data]
            parts = data.split()
            possible_files = []
            current_file = parts[0]
            for part in parts[1:]:
                test_path = f"{current_file} {part}"
                if os.path.exists(test_path):
                    current_file = test_path
                else:
                    possible_files.append(current_file)
                    current_file = part
            possible_files.append(current_file)
            return [f for f in possible_files if os.path.exists(f)]
        except Exception as e:
            print(f"Error parsing dropped files: {e}")
            return []

    def process_queue(self):
        if not self.file_queue:
            messagebox.showwarning("Empty Queue", "No files to process")
            return

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.figures = []

        for file_path in self.file_queue:
            try:
                figures = self.generate_plots(file_path)
                if figures:
                    self.figures.extend(figures)
                    self._display_figures(figures, file_path)
                else:
                    # FIX: Show error message if no plots were generated
                    messagebox.showwarning("No Plots Generated", 
                        f"No plottable tables found in {os.path.basename(file_path)}.\n"
                        "Make sure your data has:\n"
                        "- Headers in the first row\n"
                        "- At least one column with numeric values")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process {file_path}:\n{str(e)}")

        self.file_queue.clear()
        self.queue_status.config(text=f"Files ready: {len(self.file_queue)}")
        self.canvas.yview_moveto(0)

    def save_all_plots(self):
        # FIX: Save as PDF (original goal) instead of a stitched PNG.
        # Also fixed: dpi=(300,300) tuple was invalid for PIL save().
        if not self.figures:
            messagebox.showwarning("No Plots", "No plots to save")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save All Plots as PDF",
            defaultextension=".pdf",
            filetypes=[('PDF Document', '*.pdf'), ('All Files', '*.*')]
        )

        if not save_path:
            return

        try:
            from matplotlib.backends.backend_pdf import PdfPages

            with PdfPages(save_path) as pdf:
                for fig in self.figures:
                    pdf.savefig(fig, bbox_inches='tight')

            messagebox.showinfo("Success", f"All plots saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save PDF:\n{str(e)}")

    def generate_plots(self, file_path):
        try:
            figures = QuickMake(file_path)
            return figures
        except Exception as e:
            print(f"Error in QuickMake: {e}")
            return []

    def _display_figures(self, figures, file_path):
        file_header = tk.Label(self.scrollable_frame,
                               text=f"Plots from: {os.path.basename(file_path)}",
                               font=('Arial', 12, 'bold'),
                               bg="#e0e0e0")
        file_header.pack(fill=tk.X, pady=(10, 5), padx=5)

        for fig in figures:
            frame = tk.Frame(self.scrollable_frame, bd=2, relief=tk.GROOVE)
            frame.pack(fill=tk.X, padx=5, pady=5)

            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            save_btn = tk.Button(frame, text="Save Plot",
                                 command=lambda f=fig: self.save_plot(f, file_path))
            save_btn.pack(side=tk.BOTTOM, pady=5)

    def save_plot(self, fig, source_file_path):
        try:
            base_name = os.path.splitext(os.path.basename(source_file_path))[0]
            file_types = [
                ('PNG Image', '*.png'),
                ('PDF Document', '*.pdf'),
                ('SVG Vector', '*.svg'),
                ('All Files', '*.*')
            ]
            initialdir = self.last_save_dir or os.path.dirname(source_file_path)
            save_path = filedialog.asksaveasfilename(
                title="Save Plot As",
                initialdir=initialdir,
                initialfile=f"{base_name}_plot",
                defaultextension=".png",
                filetypes=file_types
            )
            if save_path:
                fig.savefig(save_path, bbox_inches='tight', dpi=300)
                self.last_save_dir = os.path.dirname(save_path)
                messagebox.showinfo("Success", f"Plot saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save plot:\n{str(e)}")


# Run the application
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = FileProcessorApp(root)
    root.mainloop()