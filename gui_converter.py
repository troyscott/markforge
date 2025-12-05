"""
==============================================================================
SCRIPT NAME:   AI Doc Converter PRO (GUI Version)
DESCRIPTION:   A modern graphical interface for the GPU-accelerated 
               Document to Markdown converter.
               
               Features:
               - Drag-and-drop style folder selection
               - Real-time log monitoring
               - Background processing (Anti-Freeze)
               - Dark Mode UI
               
DEPENDENCIES:  pip install customtkinter marker-pdf markitdown pymupdf pillow
==============================================================================
"""

import os
import sys
import threading
import tkinter as tk
import customtkinter as ctk
import shutil
import gc
import pymupdf
from pathlib import Path
from PIL import Image

# --- IMPORT AI LIBRARIES (Wrapped to prevent GUI freeze on import) ---
# We import these later in the thread or globally if fast enough. 
# For this script, global is fine as long as we don't load models yet.
from markitdown import MarkItDown
from marker.models import load_all_models
from marker.convert import convert_single_pdf

# --- CONFIGURATION & CONSTANTS ---
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
MAX_PAGES_PER_CHUNK = 25

# --- SMART LOGGING REDIRECTOR ---
class TextRedirector(object):
    """Redirects print() to GUI but filters out messy progress bars."""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string):
        # 1. Filter out known "noisy" progress bar patterns
        if "it/s]" in string or "%|" in string or "it/s" in string:
            return
        
        # 2. Filter out raw carriage returns (the \r characters used by progress bars)
        if string == "\r" or string.startswith("\r"):
            return

        # 3. Only show clean, non-empty messages
        if string.strip():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", string + "\n") # Force new line for clarity
            self.text_widget.see("end") 
            self.text_widget.configure(state="disabled")

    def flush(self):
        pass

# --- CORE LOGIC (Copied from your stable script) ---
def split_pdf(input_path, chunk_size):
    temp_files = []
    try:
        with pymupdf.open(input_path) as doc:
            total_pages = len(doc)
            if total_pages <= chunk_size:
                return [input_path]

            print(f"   ✂️  Splitting {input_path.name} ({total_pages} pages)...")
            for i in range(0, total_pages, chunk_size):
                end_page = min(i + chunk_size, total_pages)
                with pymupdf.open() as new_doc:
                    new_doc.insert_pdf(doc, from_page=i, to_page=end_page - 1)
                    chunk_name = f"{input_path.stem}_temp_chunk_{i // chunk_size}.pdf"
                    chunk_path = input_path.parent / chunk_name
                    new_doc.save(chunk_path)
                    temp_files.append(chunk_path)
    except Exception as e:
        print(f"   [!] Error splitting PDF: {e}")
        return [input_path]
    return temp_files

def convert_pdf_safely(input_path, current_output_dir, model_list):
    files_to_process = split_pdf(input_path, MAX_PAGES_PER_CHUNK)
    full_markdown = ""
    doc_folder_name = input_path.stem.replace(" ", "_")
    images_dir = current_output_dir / "images" / doc_folder_name
    os.makedirs(images_dir, exist_ok=True)
    
    for i, current_file in enumerate(files_to_process):
        try:
            print(f"   Processing part {i+1}/{len(files_to_process)}...")
            text, images, meta = convert_single_pdf(str(current_file), model_list, batch_multiplier=1)
            
            if len(images) > 0:
                print(f"      📸 Found {len(images)} images in this chunk.")
            
            for original_filename, image_obj in images.items():
                unique_name = f"chunk_{i}_{original_filename}"
                save_path = images_dir / unique_name
                image_obj.save(save_path)
                relative_link = f"images/{doc_folder_name}/{unique_name}"
                text = text.replace(original_filename, relative_link)
            
            full_markdown += text + "\n\n"
        except Exception as e:
            print(f"   [!] Failed on chunk {current_file.name}: {e}")
        finally:
            if current_file != input_path:
                try:
                    os.remove(current_file)
                except:
                    pass
            gc.collect()
    return full_markdown

def convert_office(input_path, md_converter):
    try:
        result = md_converter.convert(str(input_path))
        return result.text_content
    except Exception as e:
        print(f"   [!] Office error: {e}")
        return None

def process_logic(source_root, dest_root, model_list, md_converter):
    source_root = Path(source_root)
    dest_root = Path(dest_root)
    
    if not source_root.exists():
        print(f"❌ Input '{source_root}' not found!")
        return

    for root, dirs, files in os.walk(source_root):
        rel_path = os.path.relpath(root, source_root)
        current_dest_dir = dest_root / rel_path
        os.makedirs(current_dest_dir, exist_ok=True)

        for file in files:
            input_file_path = Path(root) / file
            output_file_name = f"{input_file_path.stem}.md"
            output_file_path = current_dest_dir / output_file_name

            if output_file_path.exists():
                print(f"⏭️  Skipping {file} (Done)")
                continue

            print(f"Processing: {file}...")
            markdown_content = None
            ext = input_file_path.suffix.lower()

            if ext == ".pdf":
                markdown_content = convert_pdf_safely(input_file_path, current_dest_dir, model_list)
            elif ext in [".docx", ".pptx", ".xlsx"]:
                markdown_content = convert_office(input_file_path, md_converter)
            
            if markdown_content:
                with open(output_file_path, "w", encoding="utf-8") as f:
                    f.write(markdown_content)
                print(f"   ✅ Saved {output_file_path.name}")

# --- GUI APPLICATION CLASS ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Document Converter Ultra")
        self.geometry("800x600")

        # Layout Config
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. Header
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.label_title = ctk.CTkLabel(self.header_frame, text="AI Document Converter", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=10)

        # 2. Controls Frame
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        # Input Button & Label
        self.btn_input = ctk.CTkButton(self.controls_frame, text="Select Input Folder", command=self.select_input)
        self.btn_input.grid(row=0, column=0, padx=10, pady=10)
        self.lbl_input = ctk.CTkLabel(self.controls_frame, text="No folder selected", text_color="gray")
        self.lbl_input.grid(row=0, column=1, padx=10, sticky="w")

        # Output Button & Label
        self.btn_output = ctk.CTkButton(self.controls_frame, text="Select Output Folder", command=self.select_output)
        self.btn_output.grid(row=1, column=0, padx=10, pady=10)
        self.lbl_output = ctk.CTkLabel(self.controls_frame, text="No folder selected", text_color="gray")
        self.lbl_output.grid(row=1, column=1, padx=10, sticky="w")

        # Start Button
        self.btn_start = ctk.CTkButton(self.controls_frame, text="START CONVERSION", fg_color="green", hover_color="darkgreen", command=self.start_thread)
        self.btn_start.grid(row=2, column=0, columnspan=2, pady=20, sticky="ew")

        # 3. Log Window
        self.textbox = ctk.CTkTextbox(self, width=760, height=300)
        self.textbox.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")
        self.textbox.configure(state="disabled")

        # Variables
        self.input_dir = ""
        self.output_dir = ""

        # Redirect Stdout
        sys.stdout = TextRedirector(self.textbox)
        sys.stderr = TextRedirector(self.textbox)

    def select_input(self):
        path = ctk.filedialog.askdirectory()
        if path:
            self.input_dir = path
            self.lbl_input.configure(text=f".../{os.path.basename(path)}")

    def select_output(self):
        path = ctk.filedialog.askdirectory()
        if path:
            self.output_dir = path
            self.lbl_output.configure(text=f".../{os.path.basename(path)}")

    def start_thread(self):
        if not self.input_dir or not self.output_dir:
            print("⚠️ Please select both Input and Output folders first.")
            return
        
        self.btn_start.configure(state="disabled", text="Processing...")
        
        # Run in separate thread to keep GUI alive
        thread = threading.Thread(target=self.run_conversion)
        thread.start()

    def run_conversion(self):
        try:
            print("-" * 30)
            print("🚀 Initializing AI Models (This may take a moment)...")
            
            # Load Models INSIDE the thread (safe because we are already in __main__)
            marker_model_list = load_all_models()
            md_converter = MarkItDown()
            
            print("✅ Models Loaded. Starting Batch...")
            process_logic(self.input_dir, self.output_dir, marker_model_list, md_converter)
            
            print("\n✨ ALL TASKS COMPLETED SUCCESSFULLY!")
        except Exception as e:
            print(f"\n❌ CRITICAL ERROR: {e}")
        finally:
            # Re-enable button (must be done via main thread in strict environments, but usually safe in CTK)
            self.btn_start.configure(state="normal", text="START CONVERSION")

if __name__ == "__main__":
    app = App()
    app.mainloop()