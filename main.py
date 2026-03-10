from tkinter import BOTH, DISABLED, NORMAL, W, X
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
import threading
import os

from utils import export_orders_to_excel, parse_pdf_orders


class App(ttk.Window):

    def __init__(self):
        super().__init__(themename="flatly")

        self.title("PDF Exporter")
        self.geometry("500x320")

        # No maximizar
        self.resizable(False, False)

        self.selected_file = None
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
        self.combo.bind("<<ComboboxSelected>>", self.enable_file_selector)

        self.file_label = ttk.Label(container, text="No file selected")
        self.file_label.pack(pady=10)

        self.select_btn = ttk.Button(
            container,
            text="Select PDF",
            command=self.select_file,
            bootstyle="primary",
            state=DISABLED
        )

        self.select_btn.pack(fill=X)

        self.export_btn = ttk.Button(
            container,
            text="Export",
            command=self.start_export,
            bootstyle="success",
            state=DISABLED
        )

        self.export_btn.pack(fill=X, pady=10)

        # RESET SIEMPRE DISPONIBLE
        self.reset_btn = ttk.Button(
            container,
            text="Reset",
            command=self.reset,
            bootstyle="warning",
            state=DISABLED
        )

        self.reset_btn.pack(fill=X)

        self.progress = ttk.Progressbar(
            container,
            mode="indeterminate",
            bootstyle="info"
        )

    def enable_file_selector(self, event=None):
        self.select_btn.config(state=NORMAL)

    def select_file(self):

        file = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )

        if file:
            self.selected_file = file
            self.file_label.config(text=os.path.basename(file))
            self.export_btn.config(state=NORMAL)

    def start_export(self):

        if not self.selected_file:
            return

        self.export_running = True
        self.cancel_export = False

        self.export_btn.config(state=DISABLED)
        self.select_btn.config(state=DISABLED)
        self.combo.config(state=DISABLED)

        self.progress.pack(fill=X, pady=10)
        self.progress.start()

        thread = threading.Thread(target=self.export)
        thread.start()

    def export(self):
        try:
            orders = parse_pdf_orders(self.selected_file)

            export_orders_to_excel(orders, self.combo.get())          

            self.after(0, self.export_finished)
        except Exception as e:
            self.after(0, lambda: ttk.dialogs.Messagebox.show_error(
                f"An error occurred during export: {str(e)}"
                "Export failed",
            ))
            self.after(0, self.reset)

    def export_finished(self):

        self.progress.stop()
        self.progress.pack_forget()

        self.export_running = False
        self.reset_btn.config(state=NORMAL)

        ttk.dialogs.Messagebox.show_info(
            "Export completed",
            "Export finished successfully. The Excel file has been saved to your Downloads folder."
        )

    def reset(self):

        # Cancelar export si está corriendo
        if self.export_running:
            self.cancel_export = True

        self.export_running = False
        self.selected_file = None

        self.file_label.config(text="No file selected")

        self.doc_type.set("")
        self.combo.config(state="readonly")

        self.select_btn.config(state=DISABLED)
        self.export_btn.config(state=DISABLED)

        self.progress.stop()
        self.progress.pack_forget()
        self.reset_btn.config(state=DISABLED)


if __name__ == "__main__":
    app = App()
    app.mainloop()