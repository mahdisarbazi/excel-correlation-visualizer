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
        self.root.title("نمایشگر همبستگی فایل اکسل")
        self.root.geometry("900x700")
        
        # تنظیم فونت برای پشتیبانی از فارسی
        self.default_font = ('Tahoma', 10)
        
        # متغیرهای کلاس
        self.excel_file = None
        self.excel_data = None
        self.sheet_name = None
        self.dataframe = None
        self.preview_window = None
        self.preview_tree = None
        self.correlation_figure = None
        self.scatter_figures = []
        
        # ایجاد دکمه انتخاب فایل
        self.open_button = tk.Button(root, text="انتخاب فایل اکسل", command=self.open_file, font=self.default_font)
        self.open_button.pack(pady=20)
        
        # وضعیت
        self.status_label = tk.Label(root, text="لطفاً یک فایل اکسل انتخاب کنید\n @ Mahdi Sarbazi", font=self.default_font)
        self.status_label.pack(pady=10)
        
    def open_file(self):
        """انتخاب فایل اکسل و نمایش پنجره انتخاب شیت"""
        file_path = filedialog.askopenfilename(
            title="انتخاب فایل اکسل",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.excel_file = file_path
            self.status_label.config(text=f"فایل انتخاب شده: {file_path}")
            
            try:
                # خواندن اطلاعات شیت‌های فایل اکسل
                self.excel_data = pd.ExcelFile(file_path)
                # نمایش پنجره انتخاب شیت
                self.show_sheet_selector()
            except Exception as e:
                self.status_label.config(text=f"خطا در باز کردن فایل: {str(e)}")
    
    def show_sheet_selector(self):
        """نمایش پنجره انتخاب شیت"""
        sheet_window = tk.Toplevel(self.root)
        sheet_window.title("انتخاب شیت")
        sheet_window.geometry("300x400")
        
        # لیست شیت‌ها
        tk.Label(sheet_window, text="لطفاً یک شیت را انتخاب کنید:", font=self.default_font).pack(pady=10)
        
        sheet_list = tk.Listbox(sheet_window, font=self.default_font, height=15, width=30)
        sheet_list.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # اضافه کردن نام شیت‌ها به لیست
        for sheet in self.excel_data.sheet_names:
            sheet_list.insert(tk.END, sheet)
        
        # دکمه انتخاب
        select_button = tk.Button(
            sheet_window, 
            text="انتخاب", 
            font=self.default_font,
            command=lambda: self.preview_sheet(sheet_list.get(tk.ACTIVE), sheet_window)
        )
        select_button.pack(pady=10)

    def preview_sheet(self, sheet_name, sheet_window):
        """نمایش پیش‌نمایش شیت و تعیین سطر شروع داده‌ها"""
        if not sheet_name:
            return
            
        try:
            self.sheet_name = sheet_name
            
            # بستن پنجره انتخاب شیت
            sheet_window.destroy()
            
            # خواندن اولین 20 سطر برای پیش‌نمایش
            preview_df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None, nrows=20)
            
            # ایجاد پنجره پیش‌نمایش
            self.preview_window = tk.Toplevel(self.root)
            self.preview_window.title(f"پیش‌نمایش شیت: {sheet_name}")
            self.preview_window.geometry("800x600")
            
            # عنوان
            tk.Label(
                self.preview_window, 
                text="لطفاً سطر شروع داده‌ها و سطر هدر (نام ستون‌ها) را مشخص کنید:", 
                font=self.default_font
            ).pack(pady=10)
            
            # فریم برای کنترل‌ها
            control_frame = tk.Frame(self.preview_window)
            control_frame.pack(pady=10, fill=tk.X)
            
            # انتخاب سطر هدر
            tk.Label(control_frame, text="سطر هدر:", font=self.default_font).pack(side=tk.LEFT, padx=5)
            header_var = tk.IntVar(value=0)
            header_spinbox = tk.Spinbox(control_frame, from_=0, to=19, textvariable=header_var, width=5)
            header_spinbox.pack(side=tk.LEFT, padx=5)
            
            # انتخاب سطر شروع داده‌ها
            tk.Label(control_frame, text="سطر شروع داده‌ها:", font=self.default_font).pack(side=tk.LEFT, padx=5)
            data_start_var = tk.IntVar(value=1)
            data_start_spinbox = tk.Spinbox(control_frame, from_=1, to=20, textvariable=data_start_var, width=5)
            data_start_spinbox.pack(side=tk.LEFT, padx=5)
            
            # دکمه تشخیص خودکار
            auto_detect_button = tk.Button(
                control_frame,
                text="تشخیص خودکار",
                font=self.default_font,
                command=lambda: self.auto_detect_headers_and_data(preview_df, header_var, data_start_var)
            )
            auto_detect_button.pack(side=tk.LEFT, padx=20)
            
            # چک‌باکس نام ستون‌ها از اولین سطر داده
            use_first_row_var = tk.BooleanVar(value=False)
            use_first_row_check = tk.Checkbutton(
                control_frame, 
                text="استفاده از اولین سطر داده به عنوان نام ستون‌ها", 
                variable=use_first_row_var,
                font=self.default_font
            )
            use_first_row_check.pack(side=tk.LEFT, padx=5)
            
            # تری‌ویو برای نمایش داده‌ها
            tree_frame = tk.Frame(self.preview_window)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # ایجاد تری‌ویو
            self.preview_tree = ttk.Treeview(tree_frame)
            
            # اسکرول‌بار عمودی
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.preview_tree.yview)
            self.preview_tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            
            # اسکرول‌بار افقی
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.preview_tree.xview)
            self.preview_tree.configure(xscrollcommand=hsb.set)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            
            self.preview_tree.pack(fill=tk.BOTH, expand=True)
            
            # بروزرسانی تری‌ویو با داده‌های پیش‌نمایش
            self.update_preview_tree(preview_df)
            
            # دکمه تأیید
            confirm_button = tk.Button(
                self.preview_window,
                text="تأیید و ادامه",
                font=self.default_font,
                command=lambda: self.load_sheet_with_options(
                    header_var.get() if not use_first_row_var.get() else None,
                    data_start_var.get()
                )
            )
            confirm_button.pack(pady=10)
            
            # تشخیص خودکار در ابتدا
            self.auto_detect_headers_and_data(preview_df, header_var, data_start_var)
            
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در بارگذاری پیش‌نمایش شیت: {str(e)}")

    def auto_detect_headers_and_data(self, preview_df, header_var, data_start_var):
        """تشخیص خودکار سطر هدر و سطر شروع داده‌ها"""
        try:
            # استراتژی‌های مختلف برای تشخیص سطر هدر و شروع داده‌ها
            
            # 1. بررسی سطرهای خالی در ابتدا
            first_non_empty_row = None
            for i in range(len(preview_df)):
                if not preview_df.iloc[i].isna().all():
                    first_non_empty_row = i
                    break
            
            if first_non_empty_row is None:
                first_non_empty_row = 0
            
            # 2. بررسی تفاوت نوع داده‌ها بین سطرها
            header_row = first_non_empty_row
            data_start_row = first_non_empty_row + 1
            
            # بررسی اگر سطر اول حاوی اعداد است، احتمالاً هدر نیست
            potential_header_row = preview_df.iloc[header_row]
            numeric_count = sum(pd.to_numeric(potential_header_row, errors='coerce').notna())
            
            # اگر بیش از نیمی از ستون‌ها عددی هستند، احتمالاً هدر نیست
            if numeric_count > len(potential_header_row) / 2:
                header_row = None
                data_start_row = first_non_empty_row
            
            # تنظیم مقادیر
            header_var.set(header_row if header_row is not None else 0)
            data_start_var.set(data_start_row)
            
            # بروزرسانی پیش‌نمایش
            self.update_preview_tree(preview_df, header_row)
            
        except Exception as e:
            messagebox.showwarning("هشدار", f"خطا در تشخیص خودکار: {str(e)}")

    def update_preview_tree(self, preview_df, header_row=None):
        """بروزرسانی نمایش تری‌ویو با داده‌های پیش‌نمایش"""
        # پاک کردن داده‌های قبلی
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # پاک کردن ستون‌های قبلی
        self.preview_tree['columns'] = ()
        
        # تعداد ستون‌ها
        num_cols = len(preview_df.columns)
        column_ids = [f"#{i}" for i in range(num_cols)]
        
        # تنظیم ستون‌ها
        self.preview_tree['columns'] = column_ids
        
        # تنظیم عرض ستون شماره سطر
        self.preview_tree.column("#0", width=60, stretch=tk.NO)
        self.preview_tree.heading("#0", text="سطر")
        
        # تنظیم عناوین و عرض ستون‌ها
        for i, col_id in enumerate(column_ids):
            self.preview_tree.column(col_id, width=100, stretch=tk.YES)
            if header_row is not None and 0 <= header_row < len(preview_df):
                header_text = str(preview_df.iloc[header_row, i])
                self.preview_tree.heading(col_id, text=header_text)
            else:
                self.preview_tree.heading(col_id, text=f"ستون {i+1}")
        
        # اضافه کردن داده‌ها
        for i in range(len(preview_df)):
            row_values = preview_df.iloc[i].tolist()
            # رنگ متفاوت برای سطر هدر
            if i == header_row:
                self.preview_tree.insert("", tk.END, text=f"{i}", values=row_values, tags=('header',))
            else:
                self.preview_tree.insert("", tk.END, text=f"{i}", values=row_values)
        
        # تنظیم رنگ سطر هدر
        self.preview_tree.tag_configure('header', background='light blue')

    def load_sheet_with_options(self, header_row, data_start_row):
        """بارگذاری داده‌های شیت با تنظیمات مشخص شده"""
        try:
            # بستن پنجره پیش‌نمایش
            if self.preview_window:
                self.preview_window.destroy()
            
            # بارگذاری داده‌ها با تنظیمات مشخص شده
            self.dataframe = pd.read_excel(
                self.excel_file, 
                sheet_name=self.sheet_name,
                header=header_row,
                skiprows=range(1, data_start_row) if data_start_row > 0 and header_row is None else None
            )
            
            # نمایش وضعیت
            self.status_label.config(text=f"شیت انتخاب شده: {self.sheet_name} - داده‌ها از سطر {data_start_row} بارگذاری شدند")
            
            # نمایش تحلیل
            self.show_analysis()
            
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در بارگذاری داده‌ها: {str(e)}")
    
    def show_analysis(self):
        """نمایش تحلیل همبستگی و نمودارهای نقطه‌ای"""
        # حذف ویجت‌های قبلی
        for widget in self.root.winfo_children():
            if widget not in [self.open_button, self.status_label]:
                widget.destroy()
        
        # پاک‌سازی لیست نمودارها
        self.scatter_figures = []
        
        # ایجاد نوت‌بوک برای سازماندهی تب‌ها
        notebook = ttk.Notebook(self.root)
        notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # تب همبستگی
        correlation_frame = tk.Frame(notebook)
        correlation_frame.pack(fill=tk.BOTH, expand=True)
        notebook.add(correlation_frame, text="نمودار همبستگی")
        
        # تب نمودارهای نقطه‌ای
        scatter_frame = tk.Frame(notebook)
        scatter_frame.pack(fill=tk.BOTH, expand=True)
        notebook.add(scatter_frame, text="نمودارهای نقطه‌ای")
        
        # تبدیل ستون‌های غیر عددی به عددی در صورت امکان
        for col in self.dataframe.columns:
            try:
                # سعی در تبدیل ستون به عددی اگر ممکن باشد
                if self.dataframe[col].dtype == object:
                    self.dataframe[col] = pd.to_numeric(self.dataframe[col], errors='coerce')
            except:
                pass
                
        # محاسبه ماتریس همبستگی
        numeric_df = self.dataframe.select_dtypes(include=[np.number])
        if numeric_df.empty:
            tk.Label(correlation_frame, text="داده‌های عددی یافت نشد", font=self.default_font).pack(pady=20)
            return
            
        correlation_matrix = numeric_df.corr()
        
        # ایجاد فریم برای نمودار همبستگی و دکمه ذخیره
        corr_content_frame = tk.Frame(correlation_frame)
        corr_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # نمایش هیت‌مپ در تب همبستگی
        self.correlation_figure = Figure(figsize=(8, 6))
        ax_corr = self.correlation_figure.add_subplot(111)
        sns.heatmap(correlation_matrix, annot=True, cmap="coolwarm", ax=ax_corr)
        ax_corr.set_title("Correlation Matrix")  # تغییر به انگلیسی
        
        # اضافه کردن هیت‌مپ به فریم همبستگی
        canvas_corr = FigureCanvasTkAgg(self.correlation_figure, corr_content_frame)
        canvas_corr.draw()
        canvas_corr.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # دکمه ذخیره نمودار همبستگی
        corr_save_button = tk.Button(
            correlation_frame,
            text="ذخیره نمودار همبستگی",
            font=self.default_font,
            command=lambda: self.save_figure(self.correlation_figure, "heatmap")
        )
        corr_save_button.pack(pady=10)
        
        # ایجاد نمودارهای نقطه‌ای برای هر جفت ستون
        cols = numeric_df.columns
        column_pairs = list(itertools.combinations(cols, 2))
        
        # ایجاد فریم اسکرول برای نمودارهای نقطه‌ای
        scatter_container = tk.Frame(scatter_frame)
        scatter_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # اسکرول‌بار عمودی
        vsb = ttk.Scrollbar(scatter_container, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # فریم محتوا
        scatter_content = tk.Canvas(scatter_container, yscrollcommand=vsb.set)
        vsb.config(command=scatter_content.yview)
        scatter_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # فریم داخلی برای نمودارها
        scatter_inner_frame = tk.Frame(scatter_content)
        scatter_content.create_window((0, 0), window=scatter_inner_frame, anchor=tk.NW)
        
        # تعیین تعداد ستون‌ها برای نمایش نمودارها
        n_cols = 2  # تعداد ستون در گرید نمودارها
        
        # ایجاد و نمایش هر نمودار نقطه‌ای در یک فریم جداگانه
        for i, (col1, col2) in enumerate(column_pairs):
            # ایجاد فریم برای هر نمودار
            pair_frame = tk.Frame(scatter_inner_frame, borderwidth=1, relief=tk.RAISED)
            pair_frame.grid(row=i // n_cols, column=i % n_cols, padx=10, pady=10, sticky="nsew")
            
            # ایجاد نمودار نقطه‌ای
            fig = Figure(figsize=(5, 4))
            ax = fig.add_subplot(111)
            
            # رسم نمودار نقطه‌ای با نقاط کوچکتر
            ax.scatter(numeric_df[col1], numeric_df[col2], alpha=0.7, s=1)  # اندازه نقاط کوچکتر (s=10)
            ax.set_xlabel(col1)
            ax.set_ylabel(col2)
            ax.set_title(f"{col1} vs {col2}")
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # اضافه کردن خط رگرسیون
            if len(numeric_df) > 1:  # حداقل دو نقطه برای رگرسیون لازم است
                try:
                    # حذف مقادیر NaN قبل از محاسبه رگرسیون
                    valid_data = numeric_df[[col1, col2]].dropna()
                    if len(valid_data) > 1:
                        z = np.polyfit(valid_data[col1], valid_data[col2], 1)
                        p = np.poly1d(z)
                        x_range = np.linspace(valid_data[col1].min(), valid_data[col1].max(), 100)
                        ax.plot(x_range, p(x_range), "r--", alpha=0.7)
                        
                        # نمایش معادله رگرسیون
                        equation = f"y = {z[0]:.4f}x + {z[1]:.4f}"
                        ax.text(0.05, 0.95, equation, transform=ax.transAxes, 
                                verticalalignment='top', fontsize=10, 
                                bbox=dict(boxstyle='round', facecolor='white', alpha=0.7))
                except Exception as e:
                    print(f"خطا در محاسبه رگرسیون برای {col1} و {col2}: {str(e)}")
            
            # اضافه کردن ضریب همبستگی
            try:
                corr_val = correlation_matrix.loc[col1, col2]
                ax.text(0.05, 0.85, f"Correlation: {corr_val:.4f}", transform=ax.transAxes,  # تغییر به انگلیسی
                        verticalalignment='top', fontsize=10, 
                        bbox=dict(boxstyle='round', facecolor='white', alpha=0.7))
            except:
                pass
                
            fig.tight_layout()
            
            # اضافه کردن نمودار به لیست
            self.scatter_figures.append((fig, f"{col1}_vs_{col2}"))
            
            # نمایش نمودار
            canvas = FigureCanvasTkAgg(fig, pair_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # دکمه ذخیره برای هر نمودار
            save_button = tk.Button(
                pair_frame,
                text="ذخیره",
                font=self.default_font,
                command=lambda fig=fig, name=f"{col1}_vs_{col2}": self.save_figure(fig, name)
            )
            save_button.pack(pady=5)
        
        # بروزرسانی اسکرول
        scatter_inner_frame.update_idletasks()
        scatter_content.config(scrollregion=scatter_content.bbox("all"))
        
        # دکمه‌های پایین صفحه
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10, fill=tk.X)
        
        # دکمه بازگشت به انتخاب فایل
        back_button = tk.Button(
            bottom_frame, 
            text="انتخاب فایل دیگر", 
            command=self.open_file,
            font=self.default_font
        )
        back_button.pack(side=tk.LEFT, padx=10)
        
        # دکمه ذخیره همه نمودارها
        save_all_button = tk.Button(
            bottom_frame, 
            text="ذخیره تمام نمودارها", 
            command=self.save_all_figures,
            font=self.default_font
        )
        save_all_button.pack(side=tk.RIGHT, padx=10)

    def save_figure(self, figure, name_prefix):
        """ذخیره یک نمودار به صورت فایل JPG"""
        try:
            # انتخاب مسیر ذخیره فایل
            file_path = filedialog.asksaveasfilename(
                title="ذخیره نمودار",
                defaultextension=".jpg",
                filetypes=[("JPEG files", "*.jpg"), ("PNG files", "*.png"), ("All files", "*.*")],
                initialfile=f"{name_prefix}.jpg"
            )
            
            if file_path:
                # ذخیره نمودار
                figure.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("موفقیت", f"نمودار با موفقیت در مسیر زیر ذخیره شد:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره نمودار: {str(e)}")

    def save_all_figures(self):
        """ذخیره تمام نمودارها در یک پوشه"""
        try:
            # انتخاب پوشه برای ذخیره نمودارها
            folder_path = filedialog.askdirectory(title="انتخاب پوشه برای ذخیره نمودارها")
            
            if not folder_path:
                return
                
            # ذخیره نمودار همبستگی
            if self.correlation_figure:
                corr_path = os.path.join(folder_path, "correlation_heatmap.jpg")
                self.correlation_figure.savefig(corr_path, dpi=300, bbox_inches='tight')
            
            # ذخیره نمودارهای نقطه‌ای
            for fig, name in self.scatter_figures:
                scatter_path = os.path.join(folder_path, f"{name}.jpg")
                fig.savefig(scatter_path, dpi=300, bbox_inches='tight')
            
            messagebox.showinfo("موفقیت", f"تمام نمودارها با موفقیت در پوشه زیر ذخیره شدند:\n{folder_path}")
            
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره نمودارها: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelVisualizer(root)
    root.mainloop()
