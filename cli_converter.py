"""==============================================================================
SCRIPT NAME:   AI Universal Document to Markdown Converter (GPU-Accelerated)
DESCRIPTION:   Recursively scans an input folder, converts PDFs, Word, Excel, 
               and PowerPoint files into clean Markdown format. 
               
               Features:
               1. PDF OCR & Layout: Uses 'surya-ocr' and 'marker' for high-fidelity 
                  PDF conversion (tables, equations, layout preservation).
               2. Image Extraction: Automatically extracts images from PDFs, saves 
                  them locally, and links them in the Markdown.
               3. Smart Chunking: Splits large PDFs into smaller 25-page chunks 
                  to prevent GPU Memory (VRAM) overflows.
               4. Windows Safe: Implements safe multiprocessing guards to prevent
                  infinite loading loops on Windows machines.

DEPENDENCIES:  pip install marker-pdf markitdown pymupdf pillow
SYSTEM REQ:    NVIDIA GPU (Recommended for speed), Python 3.9+
==============================================================================
"""

import os
import shutil
import pymupdf 
import gc
import uuid 
from pathlib import Path
from PIL import Image

# --- LIBRARY IMPORTS ---
print("🚀 Importing libraries...")
from markitdown import MarkItDown
from marker.models import load_all_models
from marker.convert import convert_single_pdf

# --- CONFIGURATION ---
INPUT_DIR = Path("./in")
OUTPUT_DIR = Path("./out")
MAX_PAGES_PER_CHUNK = 25  

def split_pdf(input_path, chunk_size):
    """Splits PDF into safe, bite-sized temporary chunks."""
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

# FIX: Added model_list argument
def convert_pdf_safely(input_path, current_output_dir, model_list):
    files_to_process = split_pdf(input_path, MAX_PAGES_PER_CHUNK)
    full_markdown = ""
    
    doc_folder_name = input_path.stem.replace(" ", "_")
    images_dir = current_output_dir / "images" / doc_folder_name
    os.makedirs(images_dir, exist_ok=True)
    
    for i, current_file in enumerate(files_to_process):
        try:
            print(f"   Processing part {i+1}/{len(files_to_process)}...")
            
            # FIX: Use the passed model_list
            text, images, meta = convert_single_pdf(
                str(current_file), 
                model_list,
                batch_multiplier=1
            )
            
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

# FIX: Added md_converter argument
def convert_office(input_path, md_converter):
    try:
        result = md_converter.convert(str(input_path))
        return result.text_content
    except Exception as e:
        print(f"   [!] Office error: {e}")
        return None

# FIX: Added models argument to pass down
def process_directory(source_root, dest_root, model_list, md_converter):
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
            elif ext == ".txt":
                try:
                    with open(input_file_path, "r", encoding="utf-8") as f:
                         markdown_content = f.read()
                except: 
                     pass

            if markdown_content:
                with open(output_file_path, "w", encoding="utf-8") as f:
                    f.write(markdown_content)
                print(f"   ✅ Saved {output_file_path.name}")

# --- MAIN EXECUTION BLOCK ---
if __name__ == "__main__":
    # 1. Models are loaded ONLY here, once.
    print("⚙️  Loading AI models (Main Process)...")
    marker_model_list = load_all_models() 
    md_converter = MarkItDown()

    # 2. Pass the models into the processing function
    process_directory(INPUT_DIR, OUTPUT_DIR, marker_model_list, md_converter)
    
    print(f"\n✨ Done! Check folder: {OUTPUT_DIR.absolute()}")