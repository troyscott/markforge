from __future__ import annotations

import json
import queue
import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from .engine import ConversionCancelled, ConversionEngine
from .extractors import detect_device
from .models import ConversionOptions


class MarkForgeApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MarkForge")
        self.geometry("940x720")
        self.minsize(820, 620)
        ctk.set_appearance_mode("dark")
        self.events: queue.Queue[dict[str, object]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar()
        self.split_var = ctk.StringVar(value="chapters")
        self.device_var = ctk.StringVar(value="auto")
        self.processing_var = ctk.StringVar(value="20")
        self.output_pages_var = ctk.StringVar(value="25")
        self.combined_var = ctk.BooleanVar(value=False)
        self.status_var = ctk.StringVar(value=f"Ready — detected device: {detect_device()}")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        self._build_ui()
        self._bind_shortcuts()
        self.after(100, self._drain_events)

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self)
        header.grid(row=0, column=0, padx=18, pady=(18, 8), sticky="ew")
        ctk.CTkLabel(header, text="MarkForge", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(12, 2))
        ctk.CTkLabel(header, text="Local, chapter-aware document conversion").pack(pady=(0, 12))

        paths = ctk.CTkFrame(self)
        paths.grid(row=1, column=0, padx=18, pady=8, sticky="ew")
        paths.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(paths, text="Input file", command=self._select_input_file).grid(row=0, column=0, padx=10, pady=8)
        ctk.CTkButton(paths, text="Input folder", command=self._select_input_folder).grid(row=1, column=0, padx=10, pady=8)
        ctk.CTkEntry(paths, textvariable=self.input_var).grid(row=0, column=1, rowspan=2, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(paths, text="Output folder", command=self._select_output).grid(row=2, column=0, padx=10, pady=8)
        ctk.CTkEntry(paths, textvariable=self.output_var).grid(row=2, column=1, padx=10, pady=8, sticky="ew")

        options = ctk.CTkFrame(self)
        options.grid(row=2, column=0, padx=18, pady=8, sticky="ew")
        for column in range(5):
            options.grid_columnconfigure(column, weight=1)
        values = [
            ("Output split", ctk.CTkOptionMenu(options, variable=self.split_var, values=["chapters", "pages", "single"])),
            ("Device", ctk.CTkOptionMenu(options, variable=self.device_var, values=["auto", "mps", "cuda", "cpu"])),
            ("Processing pages", ctk.CTkEntry(options, textvariable=self.processing_var, width=90)),
            ("Fallback pages", ctk.CTkEntry(options, textvariable=self.output_pages_var, width=90)),
        ]
        for column, (label, widget) in enumerate(values):
            ctk.CTkLabel(options, text=label).grid(row=0, column=column, padx=8, pady=(10, 2))
            widget.grid(row=1, column=column, padx=8, pady=(2, 10))
        ctk.CTkCheckBox(options, text="Combined copy", variable=self.combined_var).grid(row=1, column=4, padx=8, pady=(2, 10))

        actions = ctk.CTkFrame(self)
        actions.grid(row=3, column=0, padx=18, pady=8, sticky="ew")
        actions.grid_columnconfigure(3, weight=1)
        self.preview_button = ctk.CTkButton(actions, text="Preview structure", command=self._preview)
        self.preview_button.grid(row=0, column=0, padx=8, pady=10)
        self.start_button = ctk.CTkButton(actions, text="Start conversion", fg_color="green", command=self._start)
        self.start_button.grid(row=0, column=1, padx=8, pady=10)
        self.cancel_button = ctk.CTkButton(actions, text="Cancel", state="disabled", command=self._cancel)
        self.cancel_button.grid(row=0, column=2, padx=8, pady=10)
        ctk.CTkLabel(actions, textvariable=self.status_var, anchor="w").grid(row=0, column=3, padx=12, pady=10, sticky="ew")
        self.log = ctk.CTkTextbox(self)
        self.log.grid(row=4, column=0, padx=18, pady=(8, 18), sticky="nsew")
        self.log.configure(state="disabled")

    def _bind_shortcuts(self) -> None:
        for sequence in ("<Command-p>", "<Control-p>"):
            self.bind(sequence, lambda _event: self._preview())
        for sequence in ("<Command-r>", "<Control-r>"):
            self.bind(sequence, lambda _event: self._start())

    def _select_input_file(self) -> None:
        value = filedialog.askopenfilename(filetypes=[("Documents", "*.pdf *.docx *.pptx *.xlsx *.txt"), ("All files", "*")])
        if value:
            self.input_var.set(value)

    def _select_input_folder(self) -> None:
        value = filedialog.askdirectory()
        if value:
            self.input_var.set(value)

    def _select_output(self) -> None:
        value = filedialog.askdirectory()
        if value:
            self.output_var.set(value)

    def _conversion_options(self) -> ConversionOptions:
        return ConversionOptions(
            output_dir=Path(self.output_var.get() or "out"), split=self.split_var.get(),
            processing_pages=int(self.processing_var.get()), output_pages=int(self.output_pages_var.get()),
            device=self.device_var.get(), combined=self.combined_var.get(),
        )

    def _preview(self) -> None:
        source = Path(self.input_var.get())
        if not source.is_file() or source.suffix.lower() != ".pdf":
            self._append("Preview requires a PDF file.")
            return
        try:
            self._append(json.dumps(ConversionEngine().inspect(source, self._conversion_options()).to_dict(), indent=2, ensure_ascii=False))
        except Exception as exc:
            self._append(f"Preview failed: {exc}")

    def _start(self) -> None:
        source = Path(self.input_var.get())
        if not source.exists() or not self.output_var.get():
            self._append("Select a valid input and output folder.")
            return
        try:
            options = self._conversion_options()
            options.validate()
        except Exception as exc:
            self._append(f"Invalid options: {exc}")
            return
        self.cancel_event.clear()
        self.start_button.configure(state="disabled")
        self.preview_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.status_var.set("Converting…")
        threading.Thread(target=self._run, args=(source, options), daemon=True).start()

    def _run(self, source: Path, options: ConversionOptions) -> None:
        try:
            manifests = ConversionEngine(progress=self.events.put, cancel_event=self.cancel_event).convert(source, options)
            message = f"Completed {len(manifests)} document(s)."
        except ConversionCancelled:
            message = "Conversion cancelled."
        except Exception as exc:
            message = f"Conversion failed: {exc}"
        self.events.put({"type": "finished", "message": message})

    def _cancel(self) -> None:
        self.cancel_event.set()
        self.status_var.set("Cancelling after the current chunk…")

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                event_type = event.get("type")
                if event_type == "segment_started":
                    self.status_var.set(f"{event.get('index')}/{event.get('total')}: {event.get('title')}")
                    self._append(f"Pages {event.get('start_page')}-{event.get('end_page')}: {event.get('title')}")
                elif event_type == "segment_completed":
                    self._append(f"{event.get('status')}: {event.get('title')}")
                elif event_type == "document_completed":
                    self._append(f"Manifest: {event.get('manifest')}")
                elif event_type == "finished":
                    self._append(str(event.get("message")))
                    self.status_var.set(str(event.get("message")))
                    self.start_button.configure(state="normal")
                    self.preview_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _append(self, value: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", value + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    MarkForgeApp().mainloop()


if __name__ == "__main__":
    main()
