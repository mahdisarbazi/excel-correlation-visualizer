import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import itertools
import seaborn as sns
import os
from matplotlib.figure import Figure
import io
from PIL import Image

class ExcelVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Correlation Visualizer")
        self.root.geometry("900x700")
        
        # Font setup for support
        self.default_font = ('Tahoma', 10)
        
        # Class variables
        self.excel_file = None
        self.excel_data = None
        self.sheet_name = None
        self.dataframe = None
        self.preview_window = None
        self.preview_tree = None
        self.correlation_figure = None
        self.scatter_figures = []
        
        # Create select file button
        self.open_button = tk.Button(root, text="Select Excel File", command=self.open_file, font=self.default_font)
        self.open_button.pack(pady=20)
        
        # Status
        self.status_label = tk.Label(root, text="Please select an Excel file\n @ Mahdi Sarbazi", font=self.default_font)
        self.status_label.pack(pady=10)
        
    def open_file(self):
        """Select Excel file and display sheet selection window"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.excel_file = file_path
            self.status_label.config(text=f"Selected file: {file_path}")
            
            try:
                # Read Excel file sheet information
                self.excel_data = pd.ExcelFile(file_path)
                # Show sheet selector window
                self.show_sheet_selector()
            except Exception as e:
                self.status_label.config(text=f"Error opening file: {str(e)}")
    
    def show_sheet_selector(self):
        """Display sheet selection window"""
        sheet_window = tk.Toplevel(self.root)
        sheet_window.title("Select Sheet")
        sheet_window.geometry("300x400")
        
        # Sheet list
        tk.Label(sheet_window, text="Please select a sheet:", font=self.default_font).pack(pady=10)
        
        sheet_list = tk.Listbox(sheet_window, font=self.default_font, height=15, width=30)
        sheet_list.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Add sheet names to list
        for sheet in self.excel_data.sheet_names:
            sheet_list.insert(tk.END, sheet)
        
        # Select button
        select_button = tk.Button(
            sheet_window, 
            text="Select", 
            font=self.default_font,
            command=lambda: self.preview_sheet(sheet_list.get(tk.ACTIVE), sheet_window)
        )
        select_button.pack(pady=10)

    def preview_sheet(self, sheet_name, sheet_window):
        """Preview sheet and determine data start row"""
        if not sheet_name:
            return
            
        try:
            self.sheet_name = sheet_name
            
            # Close sheet selection window
            sheet_window.destroy()
            
            # Read first 20 rows for preview
            preview_df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None, nrows=20)
            
            # Create preview window
            self.preview_window = tk.Toplevel(self.root)
            self.preview_window.title(f"Sheet Preview: {sheet_name}")
            self.preview_window.geometry("800x600")
            
            # Title
            tk.Label(
                self.preview_window, 
                text="Please specify the header row and data start row:", 
                font=self.default_font
            ).pack(pady=10)
            
            # Frame for controls
            control_frame = tk.Frame(self.preview_window)
            control_frame.pack(pady=10, fill=tk.X)
            
            # Header row selection
            tk.Label(control_frame, text="Header Row:", font=self.default_font).pack(side=tk.LEFT, padx=5)
            header_var = tk.IntVar(value=0)
            header_spinbox = tk.Spinbox(control_frame, from_=0, to=19, textvariable=header_var, width=5)
            header_spinbox.pack(side=tk.LEFT, padx=5)
            
            # Data start row selection
            tk.Label(control_frame, text="Data Start Row:", font=self.default_font).pack(side=tk.LEFT, padx=5)
            data_start_var = tk.IntVar(value=1)
            data_start_spinbox = tk.Spinbox(control_frame, from_=1, to=20, textvariable=data_start_var, width=5)
            data_start_spinbox.pack(side=tk.LEFT, padx=5)
            
            # Auto detect button
            auto_detect_button = tk.Button(
                control_frame,
                text="Auto Detect",
                font=self.default_font,
                command=lambda: self.auto_detect_headers_and_data(preview_df, header_var, data_start_var)
            )
            auto_detect_button.pack(side=tk.LEFT, padx=20)
            
            # Checkbox for using first row as column names
            use_first_row_var = tk.BooleanVar(value=False)
            use_first_row_check = tk.Checkbutton(
                control_frame, 
                text="Use first data row as column names", 
                variable=use_first_row_var,
                font=self.default_font
            )
            use_first_row_check.pack(side=tk.LEFT, padx=5)
            
            # Treeview for displaying data
            tree_frame = tk.Frame(self.preview_window)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Create treeview
            self.preview_tree = ttk.Treeview(tree_frame)
            
            # Vertical scrollbar
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.preview_tree.yview)
            self.preview_tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Horizontal scrollbar
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.preview_tree.xview)
            self.preview_tree.configure(xscrollcommand=hsb.set)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            
            self.preview_tree.pack(fill=tk.BOTH, expand=True)
            
            # Update treeview with preview data
            self.update_preview_tree(preview_df)
            
            # Confirm button
            confirm_button = tk.Button(
                self.preview_window,
                text="Confirm and Continue",
                font=self.default_font,
                command=lambda: self.load_sheet_with_options(
                    header_var.get() if not use_first_row_var.get() else None,
                    data_start_var.get()
                )
            )
            confirm_button.pack(pady=10)
            
            # Auto detect at start
            self.auto_detect_headers_and_data(preview_df, header_var, data_start_var)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading sheet preview: {str(e)}")

    def auto_detect_headers_and_data(self, preview_df, header_var, data_start_var):
        """Auto-detect header row and data start row"""
        try:
            # Different strategies for detecting header row and data start
            
            # 1. Check for empty rows at the beginning
            first_non_empty_row = None
            for i in range(len(preview_df)):
                if not preview_df.iloc[i].isna().all():
                    first_non_empty_row = i
                    break
            
            if first_non_empty_row is None:
                first_non_empty_row = 0
            
            # 2. Check data type differences between rows
            header_row = first_non_empty_row
            data_start_row = first_non_empty_row + 1
            
            # Check if first row contains numbers, likely not a header
            potential_header_row = preview_df.iloc[header_row]
            numeric_count = sum(pd.to_numeric(potential_header_row, errors='coerce').notna())
            
            # If more than half of columns are numeric, likely not a header
            if numeric_count > len(potential_header_row) / 2:
                header_row = None
                data_start_row = first_non_empty_row
            
            # Set values
            header_var.set(header_row if header_row is not None else 0)
            data_start_var.set(data_start_row)
            
            # Update preview
            self.update_preview_tree(preview_df, header_row)
            
        except Exception as e:
            messagebox.showwarning("Warning", f"Error in auto detection: {str(e)}")

    def update_preview_tree(self, preview_df, header_row=None):
        """Update treeview display with preview data"""
        # Clear previous data
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # Clear previous columns
        self.preview_tree['columns'] = ()
        
        # Number of columns
        num_cols = len(preview_df.columns)
        column_ids = [f"#{i}" for i in range(num_cols)]
        
        # Set columns
        self.preview_tree['columns'] = column_ids
        
        # Set row number column width
        self.preview_tree.column("#0", width=60, stretch=tk.NO)
        self.preview_tree.heading("#0", text="Row")
        
        # Set column headings and widths
        for i, col_id in enumerate(column_ids):
            self.preview_tree.column(col_id, width=100, stretch=tk.YES)
            if header_row is not None and 0 <= header_row < len(preview_df):
                header_text = str(preview_df.iloc[header_row, i])
                self.preview_tree.heading(col_id, text=header_text)
            else:
                self.preview_tree.heading(col_id, text=f"Column {i+1}")
        
        # Add data
        for i in range(len(preview_df)):
            row_values = preview_df.iloc[i].tolist()
            # Different color for header row
            if i == header_row:
                self.preview_tree.insert("", tk.END, text=f"{i}", values=row_values, tags=('header',))
            else:
                self.preview_tree.insert("", tk.END, text=f"{i}", values=row_values)
        
        # Set header row color
        self.preview_tree.tag_configure('header', background='light blue')

    def load_sheet_with_options(self, header_row, data_start_row):
        """Load sheet data with specified settings"""
        try:
            # Close preview window
            if self.preview_window:
                self.preview_window.destroy()
            
            # Load data with specified settings
            self.dataframe = pd.read_excel(
                self.excel_file, 
                sheet_name=self.sheet_name,
                header=header_row,
                skiprows=range(1, data_start_row) if data_start_row > 0 and header_row is None else None
            )
            
            # Display status
            self.status_label.config(text=f"Selected sheet: {self.sheet_name} - Data loaded from row {data_start_row}")
            
            # Show analysis
            self.show_analysis()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading data: {str(e)}")
    
    def show_analysis(self):
        """Show correlation analysis and scatter plots"""
        # Remove previous widgets
        for widget in self.root.winfo_children():
            if widget not in [self.open_button, self.status_label]:
                widget.destroy()
        
        # Clear plot list
        self.scatter_figures = []
        
        # Create notebook for organizing tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # Correlation tab
        correlation_frame = tk.Frame(notebook)
        correlation_frame.pack(fill=tk.BOTH, expand=True)
        notebook.add(correlation_frame, text="Correlation Matrix")
        
        # Scatter plots tab
        scatter_frame = tk.Frame(notebook)
        scatter_frame.pack(fill=tk.BOTH, expand=True)
        notebook.add(scatter_frame, text="Scatter Plots")
        
        # Convert non-numeric columns to numeric if possible
        for col in self.dataframe.columns:
            try:
                # Try to convert column to numeric if possible
                if self.dataframe[col].dtype == object:
                    self.dataframe[col] = pd.to_numeric(self.dataframe[col], errors='coerce')
            except:
                pass
                
        # Calculate correlation matrix
        numeric_df = self.dataframe.select_dtypes(include=[np.number])
        if numeric_df.empty:
            tk.Label(correlation_frame, text="No numeric data found", font=self.default_font).pack(pady=20)
            return
            
        correlation_matrix = numeric_df.corr()
        
        # Create frame for correlation plot and save button
        corr_content_frame = tk.Frame(correlation_frame)
        corr_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Display heatmap in correlation tab
        self.correlation_figure = Figure(figsize=(8, 6))
        ax_corr = self.correlation_figure.add_subplot(111)
        sns.heatmap(correlation_matrix, annot=True, cmap="coolwarm", ax=ax_corr)
        ax_corr.set_title("Correlation Matrix")
        
        # Add heatmap to correlation frame
        canvas_corr = FigureCanvasTkAgg(self.correlation_figure, corr_content_frame)
        canvas_corr.draw()
        canvas_corr.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Save correlation plot button
        corr_save_button = tk.Button(
            correlation_frame,
            text="Save Correlation Plot",
            font=self.default_font,
            command=lambda: self.save_figure(self.correlation_figure, "heatmap")
        )
        corr_save_button.pack(pady=10)
        
        # Create scatter plots for each column pair
        cols = numeric_df.columns
        column_pairs = list(itertools.combinations(cols, 2))
        
        # Create scroll frame for scatter plots
        scatter_container = tk.Frame(scatter_frame)
        scatter_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Vertical scrollbar
        vsb = ttk.Scrollbar(scatter_container, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Content frame
        scatter_content = tk.Canvas(scatter_container, yscrollcommand=vsb.set)
        vsb.config(command=scatter_content.yview)
        scatter_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Inner frame for plots
        scatter_inner_frame = tk.Frame(scatter_content)
        scatter_content.create_window((0, 0), window=scatter_inner_frame, anchor=tk.NW)
        
        # Determine number of columns for displaying plots
        n_cols = 2  # Number of columns in plot grid
        
        # Create and display each scatter plot in a separate frame
        for i, (col1, col2) in enumerate(column_pairs):
            # Create frame for each plot
            pair_frame = tk.Frame(scatter_inner_frame, borderwidth=1, relief=tk.RAISED)
            pair_frame.grid(row=i // n_cols, column=i % n_cols, padx=10, pady=10, sticky="nsew")
            
            # Create scatter plot
            fig = Figure(figsize=(5, 4))
            ax = fig.add_subplot(111)
            
            # Draw scatter plot with smaller points
            ax.scatter(numeric_df[col1], numeric_df[col2], alpha=0.7, s=1)  # Smaller point size (s=1)
            ax.set_xlabel(col1)
            ax.set_ylabel(col2)
            ax.set_title(f"{col1} vs {col2}")
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # Add regression line
            if len(numeric_df) > 1:  # At least two points needed for regression
                try:
                    # Remove NaN values before calculating regression
                    valid_data = numeric_df[[col1, col2]].dropna()
                    if len(valid_data) > 1:
                        z = np.polyfit(valid_data[col1], valid_data[col2], 1)
                        p = np.poly1d(z)
                        x_range = np.linspace(valid_data[col1].min(), valid_data[col1].max(), 100)
                        ax.plot(x_range, p(x_range), "r--", alpha=0.7)
                        
                        # Display regression equation
                        equation = f"y = {z[0]:.4f}x + {z[1]:.4f}"
                        ax.text(0.05, 0.95, equation, transform=ax.transAxes, 
                                verticalalignment='top', fontsize=10, 
                                bbox=dict(boxstyle='round', facecolor='white', alpha=0.7))
                except Exception as e:
                    print(f"Error calculating regression for {col1} and {col2}: {str(e)}")
            
            # Add correlation coefficient
            try:
                corr_val = correlation_matrix.loc[col1, col2]
                ax.text(0.05, 0.85, f"Correlation: {corr_val:.4f}", transform=ax.transAxes,
                        verticalalignment='top', fontsize=10, 
                        bbox=dict(boxstyle='round', facecolor='white', alpha=0.7))
            except:
                pass
                
            fig.tight_layout()
            
            # Add plot to list
            self.scatter_figures.append((fig, f"{col1}_vs_{col2}"))
            
            # Display plot
            canvas = FigureCanvasTkAgg(fig, pair_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Save button for each plot
            save_button = tk.Button(
                pair_frame,
                text="Save",
                font=self.default_font,
                command=lambda fig=fig, name=f"{col1}_vs_{col2}": self.save_figure(fig, name)
            )
            save_button.pack(pady=5)
        
        # Update scroll
        scatter_inner_frame.update_idletasks()
        scatter_content.config(scrollregion=scatter_content.bbox("all"))
        
        # Bottom page buttons
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10, fill=tk.X)
        
        # Back to file selection button
        back_button = tk.Button(
            bottom_frame, 
            text="Select Another File", 
            command=self.open_file,
            font=self.default_font
        )
        back_button.pack(side=tk.LEFT, padx=10)
        
        # Save all plots button
        save_all_button = tk.Button(
            bottom_frame, 
            text="Save All Plots", 
            command=self.save_all_figures,
            font=self.default_font
        )
        save_all_button.pack(side=tk.RIGHT, padx=10)

    def save_figure(self, figure, name_prefix):
        """Save a plot as a JPG file"""
        try:
            # Select save path
            file_path = filedialog.asksaveasfilename(
                title="Save Plot",
                defaultextension=".jpg",
                filetypes=[("JPEG files", "*.jpg"), ("PNG files", "*.png"), ("All files", "*.*")],
                initialfile=f"{name_prefix}.jpg"
            )
            
            if file_path:
                # Save plot
                figure.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Plot successfully saved to:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving plot: {str(e)}")

    def save_all_figures(self):
        """Save all plots in a folder"""
        try:
            # Select folder for saving plots
            folder_path = filedialog.askdirectory(title="Select folder to save plots")
            
            if not folder_path:
                return
                
            # Save correlation plot
            if self.correlation_figure:
                corr_path = os.path.join(folder_path, "correlation_heatmap.jpg")
                self.correlation_figure.savefig(corr_path, dpi=300, bbox_inches='tight')
            
            # Save scatter plots
            for fig, name in self.scatter_figures:
                scatter_path = os.path.join(folder_path, f"{name}.jpg")
                fig.savefig(scatter_path, dpi=300, bbox_inches='tight')
            
            messagebox.showinfo("Success", f"All plots successfully saved to:\n{folder_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving plots: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelVisualizer(root)
    root.mainloop()
