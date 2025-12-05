# AI Document Converter (GPU-Accelerated)

A robust, GPU-accelerated pipeline to convert PDF, Word, Excel, and PowerPoint documents into clean Markdown. 

Features **Surya-OCR** and **Marker** for high-fidelity PDF extraction (tables, equations, layout) and **MarkItDown** for Office documents. Includes both a **Modern GUI** and a headless **CLI** for automation.

## 📂 Project Structure

* `gui_converter.py`: **(Recommended)** The modern graphical interface. Features dark mode, folder selection, and non-freezing background processing.
* `cli_converter.py`: The headless script. Best for servers or automated batch jobs.
* `cleanup.py`: A maintenance tool to wipe the output folder and remove stray temporary files.

## 📋 Prerequisites

* **OS:** Windows 10/11
* **GPU:** NVIDIA GPU (Required for reasonable performance).
* **Manager:** [Micromamba](https://mamba.readthedocs.io/en/latest/installation/micromamba-installation.html).

## 🛠️ Installation

### 1. Create Environment

Open Command Prompt (`cmd`) and run:

```cmd
:: Create environment
micromamba create -n doc-convert python=3.10 -c conda-forge -y

:: Activate
micromamba activate doc-convert
````

### 2\. Install PyTorch (CUDA)

**Critical:** You must install the CUDA version of PyTorch to use your GPU.

```cmd
pip install torch torchvision torchaudio --index-url [https://download.pytorch.org/whl/cu121](https://download.pytorch.org/whl/cu121)
```

### 3\. Install Dependencies

```cmd
pip install customtkinter marker-pdf markitdown pymupdf pillow
```

-----

## 🚀 How to Run

### Option A: The GUI (Easiest)

Launch the graphical interface to select folders and watch the log in real-time.

```cmd
python gui_converter.py
```

1. Click **Select Input Folder** (Put your PDFs/Docs here).
2. Click **Select Output Folder**.
3. Click **START CONVERSION**.

### Option B: The CLI (Automated)

Run the headless script. *Ensure you edit the `INPUT_DIR` and `OUTPUT_DIR` paths inside the script first.*

```cmd
python cli_converter.py
```

### 🧹 Maintenance

To clean up the `/out` directory and remove any temporary chunks left behind after a crash:

```cmd
python cleanup.py
```

## ⚠️ Troubleshooting

* **"OSError: Page file too small":** If you crash on huge files, ensure `MAX_PAGES_PER_CHUNK` is set to 25 or lower in the script.
  * **Stuck Logs:** If the GUI logs look messy with progress bars, ensure you are using the latest version of `gui_converter.py` with the "Smart Filter" enabled.
  