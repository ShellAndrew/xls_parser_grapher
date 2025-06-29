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

#csv_file_path = "Bellevue Nissan Reports.csv"



def clean_csv_headers(csv_file_path, output_file_path):
    # Read the CSV file
    df = pd.read_csv(csv_file_path)
    
    # Get the first header ("Bellevue Nissan")
    first_header = df.columns[0]
    
    # Replace all other headers with empty strings
    new_columns = [first_header] + [""] * (len(df.columns) - 1)
    df.columns = new_columns
    
    # Save the cleaned CSV
    df.to_csv(output_file_path, index=False)


def xls_to_csv_conversion(file_path):
    if file_path.endswith('.xlsx'):
        fengine = 'openpyxl'
    elif file_path.endswith('.xls'):
        fengine = 'xlrd'
    else:
        raise ValueError("Unsupported file format")
    '''
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    excel_file = pd.read_excel("BellevueNissanReports sandbox.xlsx", sheet_name=sheet_names[0], engine=fengine)
    excel_file.to_csv("tamp_output.csv", index=False)
    return "tamp_output.csv"
    '''
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')  # Explicit engine
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
    
    fig, axes = plt.subplots(1, len(columns) - 1, figsize=(20,4))
    data = []
    for i in range(len(columns) - 1):
        ylabel = columns[i].split(" ")
        axes[i].bar(df[columns[0]], df[columns[i + 1]])
        axes[i].set_title(columns[i + 1])
        axes[i].set_ylabel(ylabel[0])
        axes[i].tick_params(labelrotation=45, axis='x') 
        if ("Variance" in columns or "Goal" in columns):
            plt.ylim(0, 100)
    plt.tight_layout()
    plt.close()
    return fig

def get_row_count(csv_file_path):  
    with open(csv_file_path, 'r') as file:
        count = 0
        for i in file:
            count += 1
    file.close()
    return count

def create_data2(csv_file_path):
    data = []
    data2 = []
    if (csv_file_path.endswith("xlsx")):
        csv_file_path = xls_to_csv_conversion(csv_file_path)

    with open(csv_file_path) as file:
        row_count = get_row_count(csv_file_path)
        for row in file:
            data.append(row)
        data2.append(data[0].split(","))
        for i in range(1, len(data) - 1, 1):

            for a in data[i]:
                line_with_value = False
                if a != ",":
                    if a != "\n":
                        line_with_value = True
                if (data[i - 1] != data[i + 1]) or line_with_value:
                    data2.append(data[i].split(","))
                    break
        file.close()
    return data2

        
def mark_df_as_seen(marked_cells, starting_i, starting_j, height, width):
    for row in range(starting_i, starting_i + height):
        for column in range(starting_j, starting_j + width):
            marked_cells[(row, column)] = True
    #return None

def create_df_no_title(darray, i, j, marked_cells):
    title = darray[i][j]
    starting_i = i
    starting_j = j
    height = 0
    width = 0
    while j < len(darray[starting_i]) and darray[starting_i][j] != "":
        width += 1
        j += 1
    while i < len(darray) and darray[i][starting_j] != "":
        height += 1
        i += 1
    matrix = []
    
    for g in range(starting_i, starting_i + height):
        temp_list = []
        for f in range(starting_j, starting_j + width):
            if "%" in darray[g][f]:
                temp_list.append(float(darray[g][f][:-1]))
                
            else:
                temp_list.append(darray[g][f])
        matrix.append(temp_list)
    df = pd.DataFrame(matrix)
    df.columns = df.iloc[0]
    df = df[1:]
    mark_df_as_seen(marked_cells, starting_i, starting_j, height, width)
    return {title: df}

def create_df_yes_title(darray, i, j, marked_cells):
    title = darray[i][j]
    marked_cells[(i,j)] = True
    i += 1
    if darray[i][j] == "":
        return None
    starting_i = i
    starting_j = j
    height = 0
    width = 0
    
    while j < len(darray[starting_i]) and darray[starting_i][j] != "":
        width += 1
        j += 1
    while i < len(darray) and darray[i][starting_j] != "":
        height += 1
        i += 1
    matrix = []
    
    for g in range(starting_i,  starting_i + height):
        temp_list = []
        for f in range(starting_j, starting_j + width):
            if "%" in darray[g][f]:
                temp_list.append(float(darray[g][f][:-1]))
                
            else:
                temp_list.append(darray[g][f])
        matrix.append(temp_list)
    df = pd.DataFrame(matrix)
    df.columns = df.iloc[0]
    df = df[1:]
    mark_df_as_seen(marked_cells, starting_i - 1, starting_j, height + 1, width)
    return {title: df}

def parse_csv_file(list_of_dfs, csv_file_path, marked_cells):
    data2 = []
    data2 = create_data2(csv_file_path)
    return_value = list_of_dfs
    
    for row in range(len(data2)):
        for column in range(len(data2[row]) - 1):
            #print((row,column))
            if (row, column) in marked_cells:
                continue
            else:
                
                #marked_cells[(row,column)] = True
                if (data2[row][column] == "") or (data2[row][column] == "\n"):
                    continue
                else:
                    if (data2[row][column + 1] != ""):
                        return_value.append(create_df_no_title(data2, row, column, marked_cells))
                    else:
                        return_value.append(create_df_yes_title(data2, row, column, marked_cells))
    return return_value

def QuickMake(csv_file_path):  
    # Convert Excel to CSV if needed
    if csv_file_path.endswith(('.xlsx', '.xls')):
        csv_file_path = xls_to_csv_conversion(csv_file_path)
    #print(f"Processing file: {csv_file_path}")
    # Check if the CSV exists and is non-empty
    if not os.path.exists(csv_file_path):
        raise FileNotFoundError(f"CSV file not found: {csv_file_path}")
    
    with open(csv_file_path, 'r') as f:
        if not f.read(1):  # Check if file is empty
            raise ValueError("CSV file is empty after conversion.")


    all_figures = [] 
    marked_cells = {}
    data = []
    data2 = []
    list_of_dfs = []   
    data2 = create_data2(csv_file_path)

    a = parse_csv_file(list_of_dfs, csv_file_path, marked_cells)
    #print(f'length of a is: {len(a)}')
    
    for i in range(len(a) - 1):
        try:
            for key in a[i].keys():
                new_data = []
                new_data.append(a[i][key].columns.tolist())
                for j in range(len(a[i][key])):
                    new_data.append(a[i][key].iloc[j].tolist())
                    
                with open("tamp_output.csv", 'w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerows(new_data)
                    file.close()
                all_figures.append(func2("tamp_output.csv"))
                

        except:
            x = 2
    #print(f"Plotting graph #{len(all_figures) + 1}")
    return all_figures  

class FileProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Plot Viewer")
        self.root.geometry("1000x700")
        self.file_queue = []
        # Create UI elements
        self.create_widgets()
        self.last_save_dir = None

    
    def create_widgets(self):
        # Main container using grid
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        # Top control panel
        control_panel = tk.Frame(self.main_frame, height=100, bg="#f0f0f0")
        control_panel.pack(fill=tk.X, pady=5)
        
        # Drag and drop area
        self.drop_label = tk.Label(control_panel, text="Drop files here", 
                                 bg="lightgray", width=50, height=2,
                                 font=('Arial', 10))
        self.drop_label.pack(pady=5, padx=10, fill=tk.X)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.add_to_queue)
        
        # Process button
        self.process_btn = tk.Button(control_panel, text="Generate Plots", 
                                   command=self.process_queue,
                                   bg="#4CAF50", fg="white")
        self.process_btn.pack(pady=5, padx=10, ipadx=20)
        
        self.save_all_btn = tk.Button(control_panel, text="Save All Plots", 
                                    command=self.save_all_plots,
                                    bg="#2196F3", fg="white")
        self.save_all_btn.pack(pady=5, padx=10, ipadx=20, side=tk.RIGHT)

        # Queue status
        self.queue_status = tk.Label(control_panel, text="Files ready: 0")
        self.queue_status.pack()
        
        # Scrollable plot area
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
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
            # Handle case where data is already a list
            if isinstance(data, (list, tuple)):
                return [f.strip('{}') for f in data]
            
            # Handle Windows paths (enclosed in curly braces)
            if data.startswith('{') and data.endswith('}'):
                return [data.strip('{}')]
            
            # Handle single file with spaces (most common case)
            if os.path.exists(data):
                return [data]
            
            # Fallback - split on spaces but try to reconstruct paths
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
        
        # Clear previous plots
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.figures = []  
        
        # Process files
        for file_path in self.file_queue:
            try:
                figures = self.generate_plots(file_path)  # Changed from QuickMake
                if figures:
                    self.figures.extend(figures)
                    self._display_figures(figures, file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process {file_path}:\n{str(e)}")
        
        self.file_queue.clear()
        self.queue_status.config(text=f"Files ready: {len(self.file_queue)}")
        self.canvas.yview_moveto(0)

    def save_all_plots(self):
        if not self.figures:
            messagebox.showwarning("No Plots", "No plots to save")
            return
            
        try:
            from matplotlib.backends.backend_agg import FigureCanvasAgg
            from PIL import Image
            
            # Render each figure to an image
            images = []
            for fig in self.figures:
                canvas = FigureCanvasAgg(fig)
                canvas.draw()
                img = np.array(canvas.renderer.buffer_rgba())
                images.append(Image.fromarray(img))
            
            if not images:
                messagebox.showwarning("No Plots", "No valid plots to save")
                return
                
            # Calculate total dimensions
            widths, heights = zip(*(i.size for i in images))
            total_height = sum(heights)
            max_width = max(widths)
            
            # Create new image
            combined_img = Image.new('RGB', (max_width, total_height), (255, 255, 255))
            
            # Paste all images vertically
            y_offset = 0
            for img in images:
                combined_img.paste(img, (0, y_offset))
                y_offset += img.size[1]
            
            # Save the combined image
            save_path = filedialog.asksaveasfilename(
                title="Save Combined Plots As",
                defaultextension=".png",
                filetypes=[('PNG Image', '*.png'), ('All Files', '*.*')]
            )
            
            if save_path:
                combined_img.save(save_path, dpi=(300, 300))
                messagebox.showinfo("Success", f"Combined plots saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save combined plots:\n{str(e)}")

    def generate_plots(self, file_path):
        try:
            figures = QuickMake(file_path)  # Gets ALL figures
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
        count = 0
        for fig in figures:
            frame = tk.Frame(self.scrollable_frame, bd=2, relief=tk.GROOVE)
            frame.pack(fill=tk.X, padx=5, pady=5)
            
            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Optional: Add save button per figure
            save_btn = tk.Button(frame, text="Save Plot",
                               command=lambda f=fig: self.save_plot(f, file_path))
            save_btn.pack(side=tk.BOTTOM, pady=5)

    def save_plot(self, fig, source_file_path):
        """Save a matplotlib figure to a file"""
        try:
            # Get default filename based on source file and plot number
            base_name = os.path.splitext(os.path.basename(source_file_path))[0]
            
            # Ask user for save location
            file_types = [
                ('PNG Image', '*.png'),
                ('PDF Document', '*.pdf'),
                ('SVG Vector', '*.svg'),
                ('All Files', '*.*')
            ]
            
            save_path = filedialog.asksaveasfilename(
                title="Save Plot As",
                initialdir=os.path.dirname(source_file_path),
                initialfile=f"{base_name}_plot",
                defaultextension=".png",
                filetypes=file_types
            )
            
            if save_path:  # If user didn't cancel
                fig.savefig(save_path, bbox_inches='tight', dpi=300)
                messagebox.showinfo("Success", f"Plot saved to:\n{save_path}")
            
            initialdir = self.last_save_dir or os.path.dirname(source_file_path)
            if save_path:
                self.last_save_dir = os.path.dirname(save_path)

                
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save plot:\n{str(e)}")

# Run the application
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = FileProcessorApp(root)
    root.mainloop()

