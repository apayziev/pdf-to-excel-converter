import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import threading
import sys
import os
import io
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# Import the extraction function directly
try:
    from extract_to_excel import main as extract_main
except ImportError:
    extract_main = None


class PDFToExcelGUI:
    def __init__(self, root):
        self.root = root
        root.title("PDF → Excel Converter")
        root.geometry("700x600")
        root.resizable(False, False)

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill="both", expand=True)

        # ========== INPUT SECTION ==========
        ttk.Label(frame, text="Select PDF File", font=("Arial", 14, "bold")).pack(pady=5)

        self.input_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.input_path_var, width=70,
                  bootstyle="info").pack(pady=3)

        ttk.Button(frame, text="Browse PDF", bootstyle="primary-outline",
                   width=20, command=self.browse_file).pack(pady=5)

        ttk.Button(frame, text="Start Conversion", bootstyle="success",
                   padding=8, width=22, command=self.start_conversion).pack(pady=10)

        self.progress = ttk.Progressbar(frame, mode="indeterminate", bootstyle="info")
        self.progress.pack(fill="x", pady=5)

        # ========== LOGS ==========
        ttk.Label(frame, text="Logs", font=("Arial", 13, "bold")).pack(pady=4)

        # Shrink log height to keep output path visible
        self.log_area = ScrolledText(frame, width=82, height=10, font=("Consolas", 10))
        self.log_area.pack(pady=3)

        # ========== OUTPUT PATH SECTION ==========
        card = ttk.Labelframe(frame, text="Generated Excel Path",
                              padding=12, bootstyle="info")
        card.pack(fill="x", pady=10)

        self.output_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.output_var, width=85,
                  bootstyle="success").pack(pady=3)


    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.input_path_var.set(file_path)

    def log(self, text):
        self.log_area.insert("end", text + "\n")
        self.log_area.see("end")

    def start_conversion(self):
        pdf_path = self.input_path_var.get()
        if not pdf_path:
            self.log("❌ Please select a PDF file.")
            return

        self.progress.start(10)
        self.log("▶ Starting conversion...\n")

        threading.Thread(target=self.run_script, args=(pdf_path,), daemon=True).start()

    def run_script(self, pdf_path):
        try:
            # Redirect stdout to capture print statements
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            sys.stdout = io.StringIO()
            sys.stderr = io.StringIO()

            # Run the extraction function directly
            try:
                extract_main(pdf_path)
                success = True
                error_msg = None
            except Exception as e:
                success = False
                error_msg = str(e)

            # Capture output
            output = sys.stdout.getvalue()
            errors = sys.stderr.getvalue()
            
            # Restore stdout/stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr

            # Display output in log area
            for line in output.split('\n'):
                if line.strip():
                    self.log(line.strip())
                    if "Excel file created:" in line:
                        out_file = line.split("Excel file created:")[1].strip()
                        self.output_var.set(out_file)

            if errors:
                for line in errors.split('\n'):
                    if line.strip():
                        self.log(line.strip())

            self.progress.stop()

            if success:
                self.log("\n✅ Finished successfully!")
                self.log("Closing in 3 seconds...")
                self.root.after(3000, self.root.destroy)
            else:
                self.log(f"❌ An error occurred: {error_msg}")

        except Exception as e:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.progress.stop()
            self.log(f"❌ Exception: {e}")


# ----- RUN -----
if __name__ == "__main__":
    app = ttk.Window(themename="flatly")
    GUI = PDFToExcelGUI(app)
    app.mainloop()
