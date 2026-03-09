import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import BOTH, DISABLED, NORMAL, W, X, filedialog
import os
import threading
import time


class App(ttk.Window):

    def __init__(self):
        super().__init__(themename="flatly")

        self.title("PDF Exporter")
        self.geometry("500x300")

        # deshabilitar maximizar
        self.resizable(False, False)

        self.selected_file = None

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

        self.export_btn.pack(fill=X, pady=20)

        # loader
        self.progress = ttk.Progressbar(
            container,
            mode="indeterminate",
            bootstyle="success"
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

        self.export_btn.config(state=DISABLED)
        self.select_btn.config(state=DISABLED)

        self.progress.pack(fill=X, pady=10)
        self.progress.start()

        thread = threading.Thread(target=self.export)
        thread.start()

    def export(self):

        # simulación del proceso
        time.sleep(3)

        self.after(0, self.export_finished)

    def export_finished(self):

        self.progress.stop()
        self.progress.pack_forget()

        self.export_btn.config(state=NORMAL)
        self.select_btn.config(state=NORMAL)

        ttk.dialogs.Messagebox.show_info(
            "Export completed",
            "The PDF was exported successfully"
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()