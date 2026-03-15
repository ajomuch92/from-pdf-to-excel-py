from tkinter import BOTH, DISABLED, NORMAL, W, X
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
import threading
import os

from ai import parse_pdf_orders_ai
from utils import export_orders_to_excel

available_models = [
    "nvidia/nemotron-nano-12b-v2-vl:free",
    "google/gemma-3-4b-it:free",
    "google/gemma-3-12b-it:free",
    "google/gemma-3-27b-it:free"
]


class App(ttk.Window):

    def __init__(self):
        super().__init__(themename="flatly")

        self.title("PDF Exporter")
        self.geometry("520x520")

        self.resizable(False, False)

        self.selected_file = None
        self.selected_excel = None
        self.export_running = False
        self.cancel_export = False

        self.create_ui()

    def create_ui(self):

        container = ttk.Frame(self, padding=20)
        container.pack(fill=BOTH, expand=True)

        ttk.Label(container, text="Document Type").pack(anchor=W)

        self.doc_type = ttk.StringVar()

        self.combo = ttk.Combobox(
            container,
            textvariable=self.doc_type,
            values=["Invoices", "Refunds"],
            state="readonly"
        )

        self.combo.pack(fill=X, pady=10)
        self.combo.bind("<<ComboboxSelected>>", self.update_ui_state)

        ttk.Label(container, text="Enter your Open AI API key:").pack(anchor=W, pady=5)
        self.api_key = ttk.StringVar()

        self.api_key_entry = ttk.Entry(
            container,
            textvariable=self.api_key,
            style="info.TEntry",
            show="*"
        )

        self.api_key_entry.pack(fill=X, pady=5)

        self.api_key.trace_add("write", self.update_ui_state)

        ttk.Label(container, text="Choose your model:").pack(anchor=W)
        ttk.Label(container, text="- Some models are slower than others.").pack(anchor=W, padx=5)
        ttk.Label(container, text="- Some models require more resources.").pack(anchor=W, padx=5)
        ttk.Label(container, text="- Some models have rate limits.").pack(anchor=W, padx=5)

        self.model = ttk.StringVar()

        self.combo_model = ttk.Combobox(
            container,
            textvariable=self.model,
            values=available_models,
            state="readonly"
        )

        self.combo_model.pack(fill=X, pady=10)
        self.combo_model.bind("<<ComboboxSelected>>", self.update_ui_state)

        # ── Labels de archivos ──────────────────────────────────────────────
        files_labels_frame = ttk.Frame(container)
        files_labels_frame.pack(fill=X, pady=(5, 0))
        files_labels_frame.columnconfigure(0, weight=1)
        files_labels_frame.columnconfigure(1, weight=1)

        self.file_label = ttk.Label(
            files_labels_frame,
            text="No PDF selected.\nMax 10 pages.",
            justify="center"
        )
        self.file_label.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.excel_label = ttk.Label(
            files_labels_frame,
            text="No Excel selected.",
            justify="center"
        )
        self.excel_label.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        # ── Botones Select PDF / Select Excel ───────────────────────────────
        select_frame = ttk.Frame(container)
        select_frame.pack(fill=X, pady=5)
        select_frame.columnconfigure(0, weight=1)
        select_frame.columnconfigure(1, weight=1)

        self.select_btn = ttk.Button(
            select_frame,
            text="Select PDF",
            command=self.select_file,
            bootstyle="primary",
            state=DISABLED
        )
        self.select_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.select_excel_btn = ttk.Button(
            select_frame,
            text="Select Excel",
            command=self.select_excel,
            bootstyle="primary-outline",
            state=DISABLED
        )
        self.select_excel_btn.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        # ── Botones Export / Reset ──────────────────────────────────────────
        action_frame = ttk.Frame(container)
        action_frame.pack(fill=X, pady=5)
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        self.export_btn = ttk.Button(
            action_frame,
            text="Export",
            command=self.start_export,
            bootstyle="success",
            state=DISABLED
        )
        self.export_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.reset_btn = ttk.Button(
            action_frame,
            text="Reset",
            command=self.reset,
            bootstyle="warning",
            state=DISABLED
        )
        self.reset_btn.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        # ── Mensajes de estado ──────────────────────────────────────────────
        self.warning_label = ttk.Label(
            container,
            text="",
            bootstyle="danger",
            justify="center"
        )
        self.warning_label.pack(pady=10, fill=X)

        self.progress = ttk.Progressbar(
            container,
            mode="indeterminate",
            bootstyle="info"
        )

        self.status_label = ttk.Label(
            container,
            text="",
            bootstyle="info"
        )
        self.status_label.pack(pady=1, fill=X)

    def update_ui_state(self, *args):
        has_doc_type = bool(self.doc_type.get())
        has_api_key = bool(self.api_key.get().strip())
        has_model = bool(self.model.get())
        has_file = self.selected_file is not None
        has_excel = self.selected_excel is not None  # ahora requerido

        base_ready = has_doc_type and has_api_key and has_model

        self.select_btn.config(state=NORMAL if base_ready else DISABLED)
        self.select_excel_btn.config(state=NORMAL if base_ready else DISABLED)

        # Export: requiere PDF y Excel
        self.export_btn.config(
            state=NORMAL if (base_ready and has_file and has_excel) else DISABLED
        )

        self.api_key_entry.configure(
            bootstyle="success" if has_api_key else "danger"
        )

    def select_file(self):
        file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.selected_file = file
            self.file_label.config(text=os.path.basename(file))
        self.update_ui_state()

    def select_excel(self):
        file = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.selected_excel = file
            self.excel_label.config(text=os.path.basename(file))
        self.update_ui_state()

    def start_export(self):
        if not self.selected_file:
            return

        self.export_running = True
        self.cancel_export = False

        self.export_btn.config(state=DISABLED)
        self.select_btn.config(state=DISABLED)
        self.select_excel_btn.config(state=DISABLED)
        self.combo.config(state=DISABLED)
        self.warning_label.config(
            text="The export process may take some time depending on the size of the PDF"
        )

        self.progress.pack(fill=X, pady=10)
        self.progress.start()

        thread = threading.Thread(target=self.export)
        thread.start()

    def export(self):
        try:
            orders = parse_pdf_orders_ai(
                self.selected_file,
                self.api_key.get().strip(),
                self.model.get().strip(),
                self.status_label
            )
            export_orders_to_excel(orders, self.combo.get())
            self.after(0, self.export_finished)
        except Exception as e:
            self.after(0, lambda: ttk.dialogs.Messagebox.show_error(
                f"An error occurred during export: {str(e)}",
                title="Export failed",
            ))
            self.reset_btn.config(state=NORMAL)
            self.warning_label.config(text="")

    def export_finished(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.export_running = False
        self.reset_btn.config(state=NORMAL)
        self.warning_label.config(text="")

        ttk.dialogs.Messagebox.show_info(
            "Export completed",
            "Export finished successfully. The Excel file has been saved to your Downloads folder."
        )
        self.status_label.config(text="Export completed successfully.")

    def reset(self):
        if self.export_running:
            self.cancel_export = True

        self.export_running = False
        self.selected_file = None
        self.selected_excel = None

        self.file_label.config(text="No PDF selected.\nMax 10 pages.")
        self.excel_label.config(text="No Excel selected.\n(Optional)")

        self.doc_type.set("")
        self.combo.config(state="readonly")

        self.select_btn.config(state=DISABLED)
        self.select_excel_btn.config(state=DISABLED)
        self.export_btn.config(state=DISABLED)

        self.progress.stop()
        self.progress.pack_forget()
        self.reset_btn.config(state=DISABLED)
        self.status_label.config(text="")
        self.update_ui_state()


if __name__ == "__main__":
    app = App()
    app.mainloop()