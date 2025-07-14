import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
os.environ['TF_ENABLE_ONEDNN_OPTS'] = '0'  # Add this before importing tensorflow
import tensorflow as tf
import cv2
import numpy as np
import fitz  # PyMuPDF
import uuid
import pandas as pd
from datetime import datetime
import time
import threading

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Blank Page Remover - PDF Processor")
        self.root.geometry("800x600")
        
        # Initialize variables
        self.model_path = ""
        self.input_folder = ""
        self.output_folder = ""
        self.first_page_folder = ""
        self.excel_output_folder = ""
        self.processing = False
        self.progress = 0
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        
        # Create main container
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        self.title_label = ttk.Label(
            self.main_frame, 
            text="Blank Page Remover - PDF Processor", 
            style='Title.TLabel'
        )
        self.title_label.pack(pady=(0, 20))
        
        # Model Selection Frame
        model_frame = ttk.LabelFrame(self.main_frame, text="Model Configuration")
        model_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(model_frame, text="Model Path:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.model_path_entry = ttk.Entry(model_frame, width=50)
        self.model_path_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(
            model_frame, 
            text="Browse", 
            command=self.browse_model
        ).grid(row=0, column=2, padx=5, pady=5)
        
        # Folder Selection Frame
        folder_frame = ttk.LabelFrame(self.main_frame, text="Folder Configuration")
        folder_frame.pack(fill=tk.X, pady=5)
        
        # Input Folder
        ttk.Label(folder_frame, text="Input PDF Folder:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.input_folder_entry = ttk.Entry(folder_frame, width=50)
        self.input_folder_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(
            folder_frame, 
            text="Browse", 
            command=lambda: self.browse_folder(self.input_folder_entry)
        ).grid(row=0, column=2, padx=5, pady=5)
        
        # Output Folder
        ttk.Label(folder_frame, text="Output PDF Folder:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.output_folder_entry = ttk.Entry(folder_frame, width=50)
        self.output_folder_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(
            folder_frame, 
            text="Browse", 
            command=lambda: self.browse_folder(self.output_folder_entry)
        ).grid(row=1, column=2, padx=5, pady=5)
        
        # First Page Folder
        ttk.Label(folder_frame, text="First Page PDF Folder:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.first_page_folder_entry = ttk.Entry(folder_frame, width=50)
        self.first_page_folder_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(
            folder_frame, 
            text="Browse", 
            command=lambda: self.browse_folder(self.first_page_folder_entry)
        ).grid(row=2, column=2, padx=5, pady=5)
        
        # Excel Output Folder
        ttk.Label(folder_frame, text="Excel Records Folder:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.excel_output_folder_entry = ttk.Entry(folder_frame, width=50)
        self.excel_output_folder_entry.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(
            folder_frame, 
            text="Browse", 
            command=lambda: self.browse_folder(self.excel_output_folder_entry)
        ).grid(row=3, column=2, padx=5, pady=5)
        
        # Progress Frame
        progress_frame = ttk.LabelFrame(self.main_frame, text="Processing Status")
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to process")
        self.progress_label.pack(pady=5)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient=tk.HORIZONTAL, 
            length=400, 
            mode='determinate'
        )
        self.progress_bar.pack(pady=5)
        
        # Log Frame
        log_frame = ttk.LabelFrame(self.main_frame, text="Processing Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(
            log_frame, 
            height=10, 
            wrap=tk.WORD, 
            state=tk.DISABLED,
            font=('Courier', 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Button Frame
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.process_button = ttk.Button(
            button_frame, 
            text="Start Processing", 
            command=self.start_processing_thread,
            state=tk.NORMAL
        )
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Clear Log", 
            command=self.clear_log
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Exit", 
            command=self.root.quit
        ).pack(side=tk.RIGHT, padx=5)
        
        # Initialize model variable
        self.model = None
        
        # Load default paths if they exist in the original code
        self.load_default_paths()
        
    def load_default_paths(self):
        """Load default paths from the original code if they exist"""
        default_model = r"D:\erp\blank_detector_model.h5"
        default_input = r"E:\FYMCA\SEM 2\Major Project\Prototype\Input_PDFs"
        default_output = r"E:\FYMCA\SEM 2\Major Project\Prototype\Output_PDFs"
        default_first_page = r"E:\FYMCA\SEM 2\Major Project\Prototype\First_Page_PDFs"
        default_excel = r"E:\FYMCA\SEM 2\Major Project\Prototype\Records"
        
        if os.path.exists(default_model):
            self.model_path_entry.insert(0, default_model)
        if os.path.exists(default_input):
            self.input_folder_entry.insert(0, default_input)
        if os.path.exists(default_output):
            self.output_folder_entry.insert(0, default_output)
        if os.path.exists(default_first_page):
            self.first_page_folder_entry.insert(0, default_first_page)
        if os.path.exists(default_excel):
            self.excel_output_folder_entry.insert(0, default_excel)
    
    def browse_model(self):
        """Browse for model file"""
        file_path = filedialog.askopenfilename(
            title="Select Model File",
            filetypes=[("H5 Files", ".h5"), ("All Files", ".*")]
        )
        if file_path:
            self.model_path_entry.delete(0, tk.END)
            self.model_path_entry.insert(0, file_path)
    
    def browse_folder(self, entry_widget):
        """Browse for folder"""
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder_path)
    
    def log_message(self, message):
        """Add message to log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        """Clear the log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def update_progress(self, value, message=None):
        """Update progress bar and label"""
        self.progress_bar['value'] = value
        if message:
            self.progress_label.config(text=message)
        self.root.update_idletasks()
    
    def validate_paths(self):
        """Validate all required paths"""
        self.model_path = self.model_path_entry.get()
        self.input_folder = self.input_folder_entry.get()
        self.output_folder = self.output_folder_entry.get()
        self.first_page_folder = self.first_page_folder_entry.get()
        self.excel_output_folder = self.excel_output_folder_entry.get()
        
        if not self.model_path:
            messagebox.showerror("Error", "Please select a model file")
            return False
        
        if not os.path.exists(self.model_path):
            messagebox.showerror("Error", f"Model file not found at {self.model_path}")
            return False
        
        if not self.input_folder:
            messagebox.showerror("Error", "Please select an input folder")
            return False
        
        if not os.path.exists(self.input_folder):
            messagebox.showerror("Error", f"Input folder not found at {self.input_folder}")
            return False
        
        # Create output folders if they don't exist
        try:
            if not os.path.exists(self.output_folder):
                os.makedirs(self.output_folder)
                self.log_message(f"Created output folder: {self.output_folder}")

            if not os.path.exists(self.first_page_folder):
                os.makedirs(self.first_page_folder)
                self.log_message(f"Created first page folder: {self.first_page_folder}")

            if not os.path.exists(self.excel_output_folder):
                os.makedirs(self.excel_output_folder)
                self.log_message(f"Created Excel output folder: {self.excel_output_folder}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not create folders: {e}")
            return False
        
        return True
    
    def start_processing_thread(self):
        """Start processing in a separate thread"""
        if self.processing:
            return
            
        if not self.validate_paths():
            return
            
        self.processing = True
        self.process_button.config(state=tk.DISABLED)
        self.clear_log()
        
        # Start processing in a separate thread
        processing_thread = threading.Thread(target=self.process_pdfs, daemon=True)
        processing_thread.start()
    
    def process_pdfs(self):
        """Main processing function"""
        try:
            # Step 1: Suppress oneDNN warnings (optional, based on your logs)
            os.environ['TF_ENABLE_ONEDNN_OPTS'] = '0'
            
            # Step 2: Load the model
            self.log_message("Loading model...")
            self.update_progress(5, "Loading model...")
            
            try:
                self.model = tf.keras.models.load_model(self.model_path)
                self.log_message("Model loaded successfully.")
            except Exception as e:
                self.log_message(f"Could not load model from {self.model_path}. Error: {e}")
                messagebox.showerror("Error", f"Could not load model from {self.model_path}. Error: {e}")
                self.processing = False
                self.process_button.config(state=tk.NORMAL)
                return
            
            # Step 3: Initialize Excel data
            excel_data = []
            serial_number = 1
            current_date = datetime.now().strftime("%Y-%m-%d")
            current_time = datetime.now().strftime("%H%M%S")  # Format time as HHMMSS
            excel_file_name = f"BlankOutRecords_{current_date}_{current_time}.xlsx"
            excel_file_path = os.path.join(self.excel_output_folder, excel_file_name)
            
            # Step 4: Process all PDFs in the input folder
            files = os.listdir(self.input_folder)
            if not files:
                self.log_message(f"No files found in {self.input_folder}.")
                messagebox.showinfo("Info", f"No files found in {self.input_folder}.")
                self.processing = False
                self.process_button.config(state=tk.NORMAL)
                return
            
            total_files = len([f for f in files if f.lower().endswith('.pdf')])
            if total_files == 0:
                self.log_message(f"No PDF files found in {self.input_folder}.")
                messagebox.showinfo("Info", f"No PDF files found in {self.input_folder}.")
                self.processing = False
                self.process_button.config(state=tk.NORMAL)
                return
            
            processed_files = 0
            
            for file in files:
                if not self.processing:  # Check if processing was cancelled
                    break
                    
                # Skip non-PDF files and print a comment
                if not file.lower().endswith('.pdf'):
                    self.log_message(f"Skipping {file}: Not a PDF file.")
                    continue  # Proceed to the next file

                # Generate unique 10-character code for the PDF
                unique_code = str(uuid.uuid4())[:10]
                input_pdf_path = os.path.join(self.input_folder, file)
                output_pdf_path = os.path.join(self.output_folder, f"{os.path.splitext(file)[0]}_{unique_code}.pdf")
                first_page_path = os.path.join(self.first_page_folder, f"{unique_code}.pdf")
                self.log_message(f"\nProcessing PDF: {file} (Unique code: {unique_code})")

                # Record start time
                start_time = time.time()
                execution_date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Open the input PDF
                try:
                    input_pdf = fitz.open(input_pdf_path)
                    if input_pdf.page_count == 0:
                        self.log_message(f"Skipping {file}: Empty PDF.")
                        input_pdf.close()
                        continue

                    # Save the first page as a separate PDF
                    if input_pdf.page_count > 0:
                        first_page_pdf = fitz.open()
                        first_page_pdf.insert_pdf(input_pdf, from_page=0, to_page=0)
                        if os.path.exists(first_page_path):
                            os.remove(first_page_path)
                        first_page_pdf.save(first_page_path)
                        first_page_pdf.close()
                        self.log_message(f"First page saved as {first_page_path}")

                    # Create a new PDF for non-blank pages (excluding the first page)
                    output_pdf = fitz.open()

                    # Process each page, starting from the second page
                    for page_num in range(1, input_pdf.page_count):
                        if not self.processing:  # Check if processing was cancelled
                            break
                            
                        self.log_message(f"Processing page {page_num + 1} of {input_pdf.page_count} in {file}...")
                        page = input_pdf[page_num]
                        # Convert page to image in memory
                        pix = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72))  # 150 DPI for speed
                        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

                        # Classify the page
                        if not self.is_blank_page(img):
                            output_pdf.insert_pdf(input_pdf, from_page=page_num, to_page=page_num)
                            self.log_message(f"Page {page_num + 1} included (not blank).")
                        else:
                            self.log_message(f"Page {page_num + 1} skipped (blank).")

                    # Check if output PDF has pages
                    page_count = len(output_pdf)
                    if page_count == 0:
                        self.log_message(f"No non-blank pages found in {file} (excluding first page). Skipping output.")
                        output_pdf.close()
                        input_pdf.close()
                        continue

                    # Remove output file if it already exists
                    if os.path.exists(output_pdf_path):
                        os.remove(output_pdf_path)

                    # Save the output PDF
                    output_pdf.save(output_pdf_path)
                    output_pdf.close()
                    input_pdf.close()
                    self.log_message(f"Output PDF saved as {output_pdf_path} with {page_count} pages.")

                    # Calculate execution time
                    execution_time = time.time() - start_time

                    # Append record to Excel data
                    excel_data.append({
                        "Serial Number": serial_number,
                        "Folder": self.input_folder,
                        "Input file": file,
                        "Proof file": f"{unique_code}.pdf",
                        "Output file": f"{os.path.splitext(file)[0]}_{unique_code}.pdf",
                        "Date and Time": execution_date_time,
                        "Time": f"{execution_time:.2f} seconds",
                        "Location Output file": self.output_folder,
                        "Location Proof file": self.first_page_folder
                    })
                    serial_number += 1
                    processed_files += 1
                    
                    # Update progress
                    progress = int((processed_files / total_files) * 100)
                    self.update_progress(progress, f"Processed {processed_files} of {total_files} files")

                except Exception as e:
                    self.log_message(f"Error processing {file}: {e}")
                    if 'input_pdf' in locals():
                        input_pdf.close()
                    if 'output_pdf' in locals():
                        output_pdf.close()
                    if 'first_page_pdf' in locals():
                        first_page_pdf.close()

            # Save Excel file
            if excel_data and self.processing:
                df = pd.DataFrame(excel_data)
                df.to_excel(excel_file_path, index=False)
                self.log_message(f"Excel record saved as {excel_file_path}")
            
            if self.processing:
                self.update_progress(100, "Processing completed successfully!")
                messagebox.showinfo("Success", "Processing completed successfully!")
            else:
                self.update_progress(0, "Processing cancelled")
                
        except Exception as e:
            self.log_message(f"An error occurred: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
            
        finally:
            self.processing = False
            self.process_button.config(state=tk.NORMAL)
    
    def is_blank_page(self, image):
        """Classify a page as blank or not"""
        processed_img = self.preprocess_image(image)
        prediction = self.model.predict(processed_img, verbose=0)[0][0]
        return prediction < 0.5  # Threshold of 0.5 for blank (same as training)
    
    def preprocess_image(self, img, target_size=(224, 224)):
        """Preprocess image for model prediction"""
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)  # Convert to RGB
        img = cv2.resize(img, target_size)  # Resize to match model input
        img = img / 255.0  # Normalize pixel values to [0, 1]
        img = np.expand_dims(img, axis=0)  # Add batch dimension
        return img

def main():
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
