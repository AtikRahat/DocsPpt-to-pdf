import customtkinter as ctk
import os
import sys
import threading
from tkinter import filedialog, messagebox
from converter import convert_to_pdf

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class FileRow(ctk.CTkFrame):
    def __init__(self, master, file_path, remove_callback, **kwargs):
        super().__init__(master, **kwargs)
        self.file_path = file_path
        self.remove_callback = remove_callback
        
        self.filename = os.path.basename(file_path)
        
        # Grid layout
        self.grid_columnconfigure(1, weight=1) # Spacer
        
        self.label_name = ctk.CTkLabel(self, text=self.filename, width=200, anchor="w")
        self.label_name.grid(row=0, column=0, padx=10, pady=5)
        
        self.entry_start = ctk.CTkEntry(self, placeholder_text="Start Pg", width=70)
        self.entry_start.grid(row=0, column=1, padx=5, pady=5)
        
        self.entry_end = ctk.CTkEntry(self, placeholder_text="End Pg", width=70)
        self.entry_end.grid(row=0, column=2, padx=5, pady=5)
        
        self.entry_name = ctk.CTkEntry(self, placeholder_text="Output Name (Optional)", width=150)
        self.entry_name.grid(row=0, column=3, padx=5, pady=5)
        self.entry_name.insert(0, os.path.splitext(self.filename)[0] + ".pdf")
        
        self.btn_remove = ctk.CTkButton(self, text="X", width=30, fg_color="red", command=self.remove)
        self.btn_remove.grid(row=0, column=4, padx=5, pady=5)
        
        self.btn_open = ctk.CTkButton(self, text="Open", width=50, state="disabled", command=self.open_pdf)
        self.btn_open.grid(row=0, column=5, padx=5, pady=5)
        
        self.output_path = None

    def remove(self):
        self.remove_callback(self)
        self.destroy()

    def open_pdf(self):
        if self.output_path and os.path.exists(self.output_path):
            os.startfile(self.output_path)

    def set_converted(self, path):
        self.output_path = path
        self.btn_open.configure(state="normal")

    def get_options(self):
        return {
            "path": self.file_path,
            "start": self.entry_start.get(),
            "end": self.entry_end.get(),
            "output_name": self.entry_name.get()
        }

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Docx & PPT to PDF Converter")
        self.geometry("800x600")
        
        # Set window icon
        try:
            icon_path = resource_path("iconfinder-file-4341289_120551.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass
        
        self.files = []
        
        # Header
        self.header = ctk.CTkLabel(self, text="Docx & PPT to PDF Converter", font=("Arial", 20, "bold"))
        self.header.pack(pady=20)
        
        # Add Files Button - Large and Prominent
        self.btn_add = ctk.CTkButton(
            self, 
            text="📁 Add Files to Convert", 
            command=self.add_files,
            height=50,
            font=("Arial", 16, "bold"),
            fg_color="#1f6aa5",
            hover_color="#144870"
        )
        self.btn_add.pack(pady=20)
        
        # Scrollable Area for File Rows
        self.scroll_frame = ctk.CTkScrollableFrame(self, width=700, height=300)
        self.scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Convert Button
        self.btn_convert = ctk.CTkButton(self, text="Convert All", command=self.start_conversion, height=40, font=("Arial", 16))
        self.btn_convert.pack(pady=20)
        
        # Progress
        self.progress_bar = ctk.CTkProgressBar(self, width=600)
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(self, text="Ready")
        self.status_label.pack(pady=5)

    def add_files(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[
                ("Supported Files", "*.docx;*.doc;*.pptx;*.ppt"),
                ("Word Documents", "*.docx;*.doc"),
                ("PowerPoint Presentations", "*.pptx;*.ppt")
            ]
        )
        for path in file_paths:
            # Check if already added
            if any(f.file_path == path for f in self.files):
                continue
            
            row = FileRow(self.scroll_frame, path, self.remove_file)
            row.pack(fill="x", pady=2)
            self.files.append(row)

    def remove_file(self, row_obj):
        if row_obj in self.files:
            self.files.remove(row_obj)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("No Files", "Please add files to convert.")
            return
            
        self.btn_convert.configure(state="disabled")
        self.btn_add.configure(state="disabled")
        self.progress_bar.set(0)
        
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def run_conversion(self):
        total = len(self.files)
        success_count = 0
        errors = []
        
        for i, row in enumerate(self.files):
            options = row.get_options()
            input_path = options["path"]
            output_folder = os.path.dirname(input_path)
            
            # Determine output filename
            out_name = options["output_name"]
            if not out_name.lower().endswith(".pdf"):
                out_name += ".pdf"
            
            output_path = os.path.join(output_folder, out_name)
            
            # Parse pages
            start_page = options["start"].strip()
            end_page = options["end"].strip()
            
            start_page = int(start_page) if start_page.isdigit() else None
            end_page = int(end_page) if end_page.isdigit() else None
            
            self.status_label.configure(text=f"Converting: {os.path.basename(input_path)}...")
            
            success, msg = convert_to_pdf(input_path, output_path, start_page, end_page)
            
            if success:
                success_count += 1
                # Update row UI to enable Open button
                row.set_converted(output_path)
            else:
                errors.append(f"{os.path.basename(input_path)}: {msg}")
                
            self.progress_bar.set((i + 1) / total)
        
        self.status_label.configure(text=f"Completed: {success_count}/{total}")
        self.btn_convert.configure(state="normal")
        self.btn_add.configure(state="normal")
        
        if errors:
            messagebox.showerror("Errors Occurred", "\n".join(errors))
        else:
            messagebox.showinfo("Success", "All files converted successfully!")

if __name__ == "__main__":
    app = App()
    app.mainloop()
