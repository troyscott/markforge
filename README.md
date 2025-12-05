# AI Document Converter (GPU-Accelerated)

This project provides a robust, GPU-accelerated pipeline to convert PDF, Word, Excel, and PowerPoint documents into clean Markdown. It utilizes **Surya-OCR** and **Marker** for high-fidelity PDF extraction and **MarkItDown** for Office documents.

## 📋 Prerequisites

* **OS:** Windows 10/11
* **Interface:** Command Prompt (cmd.exe)
* **GPU:** NVIDIA GPU with updated drivers.
* **Package Manager:** [Micromamba](https://mamba.readthedocs.io/en/latest/installation/micromamba-installation.html) (Assumed installed and in your PATH).

## 🛠️ Environment Setup

We use `micromamba` to ensure clean handling of Python dependencies and GPU libraries without administrative privileges.

### 1. Create the Environment

Open your Command Prompt (`cmd`) and run the following commands one by one:

```cmd
:: Create a fresh environment with Python 3.10
micromamba create -n doc-convert python=3.10 -c conda-forge -y

:: Activate the environment
micromamba activate doc-convert
````

### 2\. Install PyTorch with CUDA (Critical)

To ensure the AI models use your GPU instead of your CPU, we must install the CUDA-enabled version of PyTorch first.

```cmd
:: Install PyTorch for CUDA 12.1 (Standard for modern GPUs)
pip install torch torchvision torchaudio --index-url [https://download.pytorch.org/whl/cu121](https://download.pytorch.org/whl/cu121)
```

### 3\. Install Application Dependencies

Now install the converter libraries.

```cmd
pip install marker-pdf markitdown pymupdf pillow
```

-----

## 🚀 How to Run

### 1\. Prepare your folders

Ensure your directory looks like this:

```text
\ProjectRoot
  ├── convert.py           # The main script
  ├── cleanup.py           # The cleanup script
  ├── \in                  # PUT YOUR FILES HERE (PDF, DOCX, etc.)
  └── \out                 # Results will appear here
```

### 2\. Execute the Converter

Run the main script. It will automatically detect your GPU and start processing.

```cmd
python convert.py
```

*Note: The first time you run this, it will download several GBs of AI models. This is normal.*

### 3\. Cleaning Up

If you want to clear the output folder and reset for a fresh batch:

```cmd
python cleanup.py
```

## ⚠️ Troubleshooting

**"OSError: Page file too small" or Out of Memory:**
If processing large batches, ensure `convert.py` is configured with `MAX_PAGES_PER_CHUNK = 25` (or lower).

**GPU not being used:**
Run this command to check if CUDA is active:

```cmd
python -c "import torch; print(torch.cuda.is_available())"
```

If it says `False`, reinstall the PyTorch step above.
