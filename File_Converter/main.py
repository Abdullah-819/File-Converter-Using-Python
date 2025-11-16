import os
import threading
import logging
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk, ImageDraw, ImageFont

from Converters import *
from Progress import CircularProgressBar

# Setup logging
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    filename=os.path.join(LOG_DIR, "app.log"),
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal File Converter")
        self.root.geometry("950x550")
        self.root.resizable(False, False)
        self.file_path = None

        # ===== LEFT FRAME (Neon Illustration) =====
        self.left_frame = tk.Frame(root, width=500, height=550)
        self.left_frame.pack(side="left", fill="both")
        self.left_frame.pack_propagate(False)
        bg_path = os.path.join(os.path.dirname(__file__), "background.png")
        try:
            img = Image.open(bg_path).resize((500,550), Image.LANCZOS)
        except:
            img = Image.new('RGB', (500,550), '#4A1A7F')
            d = ImageDraw.Draw(img)
            font = ImageFont.load_default()
            d.text((50,250),"Universal\nConverter", fill=(255,255,255), font=font)
        self.bg_image = ImageTk.PhotoImage(img)
        self.bg_label = tk.Label(self.left_frame, image=self.bg_image)
        self.bg_label.pack(fill="both", expand=True)

        # ===== RIGHT FRAME (Controls) =====
        self.right_frame = tk.Frame(root, width=450, height=550, bg="#1A1A2E", padx=20, pady=20)
        self.right_frame.pack(side="right", fill="both", expand=True)
        self.right_frame.pack_propagate(False)

        # Title
        self.title_label = tk.Label(self.right_frame, text="Universal File Converter",
                                    font=("Segoe UI", 24, "bold"), bg="#1A1A2E", fg="white")
        self.title_label.pack(pady=(20,15))

        # Drag & Drop
        self.drop_frame = tk.Frame(self.right_frame, bg="#2E1A47", highlightbackground="#9D4EDD",
                                   highlightthickness=2)
        self.drop_frame.pack(pady=10, fill="x", ipady=15)
        self.drop_label_text = tk.Label(self.drop_frame, text="Drag & Drop Your File Here\nor Click 'Upload File'",
                                        bg="#2E1A47", fg="#EEE", font=("Segoe UI", 12, "italic"))
        self.drop_label_text.pack(expand=True)
        self.drop_frame.bind("<Button-1>", lambda e: self.upload_file())
        self.drop_label_text.bind("<Button-1>", lambda e: self.upload_file())

        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<Drop>>', self.drop_file)
        self.drop_frame.dnd_bind('<<Drop>>', self.drop_file)
        self.drop_label_text.dnd_bind('<<Drop>>', self.drop_file)

        # File label
        self.file_label = tk.Label(self.right_frame, text="No file selected", font=("Segoe UI",10),
                                   bg="#1A1A2E", fg="#CCC")
        self.file_label.pack(pady=(5,15))

        # Conversion dropdown
        self.conversion_var = tk.StringVar()
        self.dropdown = ttk.Combobox(self.right_frame, textvariable=self.conversion_var, state="readonly",
                                     width=30, font=("Segoe UI",11))
        self.dropdown['values'] = [
            "DOCX → PDF", "PDF → DOCX", "PDF → PPTX",
            "TXT → PDF", "PDF → TXT", "DOCX → PPTX", "PPTX → DOCX"
        ]
        self.dropdown.pack(pady=10)
        self.dropdown.set("Select a conversion...")

        # Convert button
        self.convert_btn = tk.Button(self.right_frame, text="🔄 Convert File", bg="#9D4EDD",
                                     fg="white", font=("Segoe UI", 12, "bold"),
                                     activebackground="#7B2CBF", command=self.start_conversion_thread)
        self.convert_btn.pack(pady=20, ipadx=10, ipady=5)

        # Circular Progress
        self.circular_progress = CircularProgressBar(self.right_frame, size=120, fg="#9D4EDD", bg="#1A1A2E")
        self.circular_progress.pack(pady=(0,10))

    # ===== File Selection =====
    def upload_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.file_path = path
            self.file_label.config(text=f"Selected File: {os.path.basename(path)}")
            self.drop_label_text.config(text=f"File Loaded: {os.path.basename(path)}")
            logging.info(f"File loaded: {self.file_path}")

    def drop_file(self, event):
        self.file_path = event.data.strip("{}")
        if self.file_path:
            self.file_label.config(text=f"Selected File: {os.path.basename(self.file_path)}")
            self.drop_label_text.config(text=f"File Loaded: {os.path.basename(self.file_path)}")
            logging.info(f"File dropped: {self.file_path}")

    # ===== Threaded Conversion =====
    def start_conversion_thread(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return
        if not self.conversion_var.get() or self.conversion_var.get()=="Select a conversion...":
            messagebox.showerror("Error", "Please choose a conversion option.")
            return
        self.convert_btn.config(state="disabled")
        self.dropdown.config(state="disabled")
        threading.Thread(target=self._convert_file, daemon=True).start()

    def _convert_file(self):
        try:
            conv = self.conversion_var.get()
            ext_map = {
                "DOCX → PDF":".pdf","PDF → DOCX":".docx","PDF → PPTX":".pptx",
                "TXT → PDF":".pdf","PDF → TXT":".txt","DOCX → PPTX":".pptx","PPTX → DOCX":".docx"
            }
            save_path = filedialog.asksaveasfilename(
                defaultextension=ext_map.get(conv,""),
                initialfile=os.path.splitext(os.path.basename(self.file_path))[0]+ext_map.get(conv,""),
                parent=self.root
            )
            if not save_path:
                self._reset_gui()
                return

            # Map conversions
            funcs = {
                "DOCX → PDF": docx_to_pdf_conv,
                "PDF → DOCX": pdf_to_docx_conv,
                "PDF → PPTX": pdf_to_pptx_conv,
                "TXT → PDF": txt_to_pdf_conv,
                "PDF → TXT": pdf_to_txt_conv,
                "DOCX → PPTX": docx_to_pptx_conv,
                "PPTX → DOCX": pptx_to_docx_conv
            }
            func = funcs.get(conv)

            logging.info(f"Starting conversion: {conv} - {self.file_path} -> {save_path}")
            func(self.file_path, save_path,
                 callback=lambda p: self.root.after(0, lambda: self.circular_progress.update_progress(p)))

            messagebox.showinfo("Success", f"Conversion complete!\nSaved as:\n{save_path}")
            logging.info(f"Conversion successful: {save_path}")
        except Exception as e:
            logging.exception("Conversion failed")
            messagebox.showerror("Error", str(e))
        finally:
            self._reset_gui()

    def _reset_gui(self):
        self.convert_btn.config(state="normal")
        self.dropdown.config(state="readonly")
        self.circular_progress.update_progress(0)


if __name__=="__main__":
    root = TkinterDnD.Tk()
    app = FileConverterApp(root)
    root.mainloop()
