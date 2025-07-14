import os
os.environ['TF_ENABLE_ONEDNN_OPTS'] = '0'
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tensorflow as tf
import cv2
import numpy as np
import fitz  # PyMuPDF
import uuid
import pandas as pd
from datetime import datetime
import time
import threading
from openpyxl import Workbook

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Blank Page Remover - PDF Processor")
        self.root.geometry("800x600")
        
        self.model_path = ""
        self.input_folder = ""
        self.output_folder = ""
        self.first_page_folder = ""
        self.excel_output_folder = ""
        self.processing = False
        self.progress = 0
        self.model = None
        
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        
        self.create_widgets()
        self.load_default_paths()

    def create_widgets(self):
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.title_label = ttk.Label(self.main_frame, text="Blank Page Remover - PDF Processor", style='Title.TLabel')
        self.title_label.pack(pady=(0, 20))
        
        model_frame = ttk.LabelFrame(self.main_frame, text="Model Configuration")
        model_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(model_frame, text="Model Path:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.model_path_entry = ttk.Entry(model_frame, width=50)
        self.model_path_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(model_frame, text="Browse", command=self.browse_model).grid(row=0, column=2, padx=5, pady=5)
        
        folder_frame = ttk.LabelFrame(self.main_frame, text="Folder Configuration")
        folder_frame.pack(fill=tk.X, pady=5)
        
        labels = ["Input PDF Folder:", "Output PDF Folder:", "First Page PDF Folder:", "Excel Records Folder:"]
        self.input_folder_entry = ttk.Entry(folder_frame, width=50)
        self.output_folder_entry = ttk.Entry(folder_frame, width=50)
        self.first_page_folder_entry = ttk.Entry(folder_frame, width=50)
        self.excel_output_folder_entry = ttk.Entry(folder_frame, width=50)
        entries = [self.input_folder_entry, self.output_folder_entry, self.first_page_folder_entry, self.excel_output_folder_entry]

        for i, (label, entry) in enumerate(zip(labels, entries)):
            ttk.Label(folder_frame, text=label).grid(row=i, column=0, sticky=tk.W, padx=5, pady=5)
            entry.grid(row=i, column=1, padx=5, pady=5)
            ttk.Button(folder_frame, text="Browse", command=lambda e=entry: self.browse_folder(e)).grid(row=i, column=2, padx=5, pady=5)
        
        progress_frame = ttk.LabelFrame(self.main_frame, text="Processing Status")
        progress_frame.pack(fill=tk.X, pady=10)
        self.progress_label = ttk.Label(progress_frame, text="Ready to process")
        self.progress_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress_bar.pack(pady=5)
        
        log_frame = ttk.LabelFrame(self.main_frame, text="Processing Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, font=('Courier', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        self.process_button = ttk.Button(button_frame, text="Start Processing", command=self.start_processing_thread)
        self.process_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

    def load_default_paths(self):
        default_paths = {
            'model': r"E:\FYMCA\SEM 2\Major Project\blank_detector_model.h5",
            'input': r"E:\FYMCA\SEM 2\Major Project\Prototype\Input_PDFs",
            'output': r"E:\FYMCA\SEM 2\Major Project\Prototype\Output_PDFs",
            'first_page': r"E:\FYMCA\SEM 2\Major Project\Prototype\First_Page_PDFs",
            'excel': r"E:\FYMCA\SEM 2\Major Project\Prototype\Records"
        }

        for path_type, path in default_paths.items():
            if os.path.exists(path):
                if path_type == "model":
                    entry_widget = self.model_path_entry
                else:
                    entry_widget = getattr(self, f"{path_type}_folder_entry")
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, path)

    def browse_model(self):
        file_path = filedialog.askopenfilename(title="Select Model File", filetypes=[("H5 Files", "*.h5"), ("All Files", "*.*")])
        if file_path:
            self.model_path_entry.delete(0, tk.END)
            self.model_path_entry.insert(0, file_path)

    def browse_folder(self, entry_widget):
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder_path)

    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def update_progress(self, value, message=None):
        self.progress_bar['value'] = value
        if message:
            self.progress_label.config(text=message)
        self.root.update_idletasks()

    def validate_paths(self):
        self.model_path = self.model_path_entry.get()
        self.input_folder = self.input_folder_entry.get()
        self.output_folder = self.output_folder_entry.get()
        self.first_page_folder = self.first_page_folder_entry.get()
        self.excel_output_folder = self.excel_output_folder_entry.get()

        if not self.model_path or not os.path.exists(self.model_path):
            messagebox.showerror("Error", f"Model file not found at {self.model_path}")
            return False
        if not self.input_folder or not os.path.exists(self.input_folder):
            messagebox.showerror("Error", f"Input folder not found at {self.input_folder}")
            return False
        try:
            os.makedirs(self.output_folder, exist_ok=True)
            os.makedirs(self.first_page_folder, exist_ok=True)
            os.makedirs(self.excel_output_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not create folders: {e}")
            return False
        return True

    def start_processing_thread(self):
        if self.processing or not self.validate_paths():
            return
        self.processing = True
        self.process_button.config(state=tk.DISABLED)
        self.clear_log()
        threading.Thread(target=self.process_pdfs, daemon=True).start()

    def process_pdfs(self):
        try:
            self.log_message("Loading model...")
            self.update_progress(5, "Loading model...")
            self.model = tf.keras.models.load_model(self.model_path)
            self.model.summary(print_fn=lambda x: self.log_message(x))
            self.log_message("Model loaded successfully.")

            excel_data = []
            current_date = datetime.now().strftime("%Y-%m-%d")
            current_time = datetime.now().strftime("%H%M%S")
            excel_file_name = f"BlankOutRecords_{current_date}_{current_time}.xlsx"

            files = [f for f in os.listdir(self.input_folder) if f.lower().endswith('.pdf')]
            if not files:
                self.log_message(f"No PDF files found in {self.input_folder}.")
                return

            for i, file in enumerate(files, 1):
                unique_code = str(uuid.uuid4())[:10]
                input_path = os.path.join(self.input_folder, file)
                output_path = os.path.join(self.output_folder, f"{os.path.splitext(file)[0]}_{unique_code}.pdf")
                first_page_path = os.path.join(self.first_page_folder, f"{unique_code}.pdf")
                execution_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                start = time.time()
                self.log_message(f"\nProcessing {file}...")

                try:
                    with fitz.open(input_path) as doc:
                        if doc.page_count == 0:
                            continue
                        fitz.open().insert_pdf(doc, from_page=0, to_page=0).save(first_page_path)

                        output_doc = fitz.open()
                        for p in range(1, doc.page_count):
                            pix = doc[p].get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
                            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                            if not self.is_blank_page(img):
                                output_doc.insert_pdf(doc, from_page=p, to_page=p)

                        if len(output_doc) > 0:
                            output_doc.save(output_path)
                            excel_data.append({
                                "Serial Number": len(excel_data) + 1,
                                "Folder": self.input_folder,
                                "Input file": file,
                                "Proof file": os.path.basename(first_page_path),
                                "Output file": os.path.basename(output_path),
                                "Date and Time": execution_time,
                                "Time": f"{time.time() - start:.2f} seconds",
                                "Location Output file": self.output_folder,
                                "Location Proof file": self.first_page_folder
                            })
                            self.log_message(f"Saved: {output_path}")

                except Exception as e:
                    self.log_message(f"Error: {e}")

                self.update_progress(int(i / len(files) * 100), f"Processed {i}/{len(files)}")

            if excel_data:
                try:
                    df = pd.DataFrame(excel_data)
                    with pd.ExcelWriter(os.path.join(self.excel_output_folder, excel_file_name), engine="openpyxl") as writer:
                        df.to_excel(writer, index=False)
                    self.log_message("Excel file saved.")
                except PermissionError:
                    self.log_message("Close the Excel file before writing.")
                except Exception as e:
                    self.log_message(f"Excel write error: {e}")

            self.update_progress(100, "Processing completed")
            messagebox.showinfo("Success", "Processing completed successfully!")

        except Exception as e:
            self.log_message(f"Fatal error: {e}")
        finally:
            self.processing = False
            self.root.after(0, lambda: self.process_button.config(state=tk.NORMAL))

    def is_blank_page(self, img):
        try:
            preprocessed = self.preprocess_image(img)
            return self.model.predict(preprocessed, verbose=0)[0][0] < 0.5
        except Exception as e:
            self.log_message(f"Prediction error: {e}")
            return False

    def preprocess_image(self, img, size=(224, 224)):
        try:
            if img.shape[-1] == 4:
                img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            img = cv2.resize(img, size)
            img = img / 255.0
            return np.expand_dims(img, axis=0)
        except Exception as e:
            self.log_message(f"Preprocess error: {e}")
            return np.zeros((1, *size, 3))

def main():
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
