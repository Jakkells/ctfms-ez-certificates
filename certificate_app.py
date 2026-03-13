"""
CTFMS EZ Certificates - GUI application.
Uses SERVICE CERTIFICATE.docx template. User selects output and PEP folders at startup.
"""
from docx import Document
from datetime import datetime
import os
import shutil
from docx2pdf import convert
import zipfile
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json

# Template and settings are in the same folder as this script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "SERVICE CERTIFICATE.docx")
SETTINGS_PATH = os.path.join(SCRIPT_DIR, "certificate_settings.json")


def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%d.%m.%Y").strftime("%d.%m.%Y")
    except ValueError:
        return None


def calculate_precise_font_size(name: str) -> int:
    length = len(name.strip())
    size_pt = 28 - ((length - 30) * 0.3)
    size_pt = max(8, round(size_pt))
    return size_pt * 2


def simple_xml_replace(file_path, placeholder, replacement, name_for_font_size=None):
    with zipfile.ZipFile(file_path, 'r') as docx_zip:
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_zip.extractall(temp_dir)
            doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            if os.path.exists(doc_xml_path):
                with open(doc_xml_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                if placeholder in content:
                    if placeholder == '{{NAME}}' and name_for_font_size:
                        font_size = calculate_precise_font_size(name_for_font_size)
                        replacement_xml = f"""
<w:r>
  <w:rPr>
    <w:rFonts w:ascii=\"Daytona\" w:hAnsi=\"Daytona\" w:eastAsia=\"Daytona\" w:cs=\"Daytona\"/>
    <w:sz w:val=\"{font_size}\"/>
    <w:b/>
  </w:rPr>
  <w:t>{replacement}</w:t>
</w:r>
"""
                        content = content.replace(f"<w:t>{placeholder}</w:t>", replacement_xml)
                    else:
                        content = content.replace(placeholder, replacement)
                    with open(doc_xml_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    new_zip_path = file_path.replace('.docx', '_modified.docx')
                    with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                file_path_full = os.path.join(root, file)
                                arcname = os.path.relpath(file_path_full, temp_dir)
                                new_zip.write(file_path_full, arcname)
                    return new_zip_path
    return None


def process_entries(entries, output_folder, pep_folder, template_path, log_callback):
    if not os.path.exists(template_path):
        log_callback(f"❌ Template not found: {template_path}\n")
        return
    if not output_folder:
        log_callback("❌ Please set Output folder.\n")
        return
    if not pep_folder:
        log_callback("❌ Please set PEP folder.\n")
        return
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(pep_folder, exist_ok=True)

    for name, date, next_date in entries:
        log_callback(f"🔄 Processing certificate for {name}...\n")
        temp_file = os.path.join(output_folder, f"temp_{name.replace(' ', '_')}.docx")
        shutil.copy2(template_path, temp_file)

        modified_file = simple_xml_replace(temp_file, '{{NAME}}', name, name_for_font_size=name)
        if modified_file:
            os.replace(modified_file, temp_file)
        modified_file = simple_xml_replace(temp_file, '{{DATE}}', date)
        if modified_file:
            os.replace(modified_file, temp_file)
        modified_file = simple_xml_replace(temp_file, '{{NEXTDATE}}', next_date)
        if modified_file:
            os.replace(modified_file, temp_file)

        parts = name.split(" - ")
        first_part = parts[0].strip().lower()
        pep_variants = ["pep", "pep home", "pep cell"]
        is_pep = first_part in pep_variants

        if is_pep and len(parts) >= 3:
            filename = f"{parts[1].strip()} - {parts[2].strip()}"
        elif not is_pep and len(parts) >= 3:
            filename = f"{parts[0].strip()} - {parts[1].strip()} - {parts[2].strip()}"
        else:
            filename = name

        out_folder = pep_folder if is_pep else output_folder
        docx_path = os.path.join(out_folder, f"{filename}.docx")
        pdf_path = os.path.join(out_folder, f"{filename}.pdf")

        os.rename(temp_file, docx_path)
        log_callback(f"📄 DOCX saved: {docx_path}\n")
        try:
            convert(docx_path, pdf_path)
            os.remove(docx_path)
            log_callback(f"✅ PDF saved: {pdf_path}\n")
        except Exception as e:
            log_callback(f"❌ PDF conversion failed: {e}\n")
            log_callback(f"📄 DOCX file kept: {docx_path}\n")


class FolderSetupDialog:
    """Initial dialog to choose Output and PEP folders."""
    def __init__(self, parent, default_output: str = "", default_pep: str = ""):
        self.result = None
        self.win = tk.Toplevel(parent)
        self.win.title("Select folders")
        # Slightly wider so buttons and paths are not cut off
        self.win.geometry("650x200")
        self.win.transient(parent)
        self.win.grab_set()
        # Ensure the dialog appears on top (helps if it opens behind other windows)
        self.win.lift()
        self.win.attributes("-topmost", True)
        self.win.after(300, lambda: self.win.attributes("-topmost", False))

        ttk.Label(self.win, text="Output folder (certificates):").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        self.output_var = tk.StringVar(value=default_output)
        ttk.Entry(self.win, textvariable=self.output_var, width=55).grid(row=0, column=1, padx=4, pady=8)
        ttk.Button(self.win, text="Browse...", command=self._browse_output).grid(row=0, column=2, padx=4, pady=8)

        ttk.Label(self.win, text="PEP folder (PEP certificates):").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        self.pep_var = tk.StringVar(value=default_pep)
        ttk.Entry(self.win, textvariable=self.pep_var, width=55).grid(row=1, column=1, padx=4, pady=8)
        ttk.Button(self.win, text="Browse...", command=self._browse_pep).grid(row=1, column=2, padx=4, pady=8)

        ttk.Button(self.win, text="Continue", command=self._ok).grid(row=2, column=1, pady=16)
        self.win.protocol("WM_DELETE_WINDOW", self._cancel)
        parent.wait_window(self.win)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select Output folder")
        if path:
            self.output_var.set(path)

    def _browse_pep(self):
        path = filedialog.askdirectory(title="Select PEP folder")
        if path:
            self.pep_var.set(path)

    def _ok(self):
        out = self.output_var.get().strip()
        pep = self.pep_var.get().strip()
        if not out:
            messagebox.showwarning("Missing folder", "Please select the Output folder.")
            return
        if not pep:
            messagebox.showwarning("Missing folder", "Please select the PEP folder.")
            return
        self.result = (out, pep)
        self.win.destroy()

    def _cancel(self):
        self.result = None
        self.win.destroy()


class MainApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CTFMS EZ Certificates")
        self.root.geometry("640x520")
        self.root.minsize(500, 400)

        self.output_folder = ""
        self.pep_folder = ""
        self.entries = []  # list of (name, date, next_date)

        # Load previously saved folders if they exist
        saved_out, saved_pep = self._load_settings()

        if saved_out and saved_pep:
            # Paths already known: use them directly and open app immediately
            self.output_folder, self.pep_folder = saved_out, saved_pep
        else:
            # First run (or incomplete settings): show folder selection dialog once
            dlg = FolderSetupDialog(self.root, default_output=saved_out, default_pep=saved_pep)
            if dlg.result is None:
                # User cancelled without picking folders; exit cleanly
                self.root.destroy()
                return
            self.output_folder, self.pep_folder = dlg.result
            # Persist the chosen folders for next run
            self._save_settings(self.output_folder, self.pep_folder)

        self._build_ui()
        if self.output_folder and self.pep_folder:
            self.root.mainloop()

    def _log(self, msg):
        self.log_area.insert(tk.END, msg)
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def _build_ui(self):
        # Notebook with main app and settings
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True)

        cert_frame = ttk.Frame(notebook)
        settings_frame = ttk.Frame(notebook)
        notebook.add(cert_frame, text="Certificates")
        notebook.add(settings_frame, text="Settings")

        # --- Certificates tab ---
        # Folder display
        frm = ttk.LabelFrame(cert_frame, text="Folders (current)", padding=8)
        frm.pack(fill=tk.X, padx=8, pady=6)
        self.output_label = ttk.Label(frm, text=f"Output: {self.output_folder}", foreground="gray")
        self.output_label.pack(anchor="w")
        self.pep_label = ttk.Label(frm, text=f"PEP: {self.pep_folder}", foreground="gray")
        self.pep_label.pack(anchor="w")

        # Add entry
        add_frm = ttk.LabelFrame(cert_frame, text="Add entry", padding=8)
        add_frm.pack(fill=tk.X, padx=8, pady=6)
        row = ttk.Frame(add_frm)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Name:").pack(side=tk.LEFT, padx=(0, 4))
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(row, textvariable=self.name_var, width=40)
        name_entry.pack(side=tk.LEFT, padx=4)
        ttk.Label(row, text="Date (DD.MM.YYYY):").pack(side=tk.LEFT, padx=(12, 4))
        self.date_var = tk.StringVar()
        date_entry = ttk.Entry(row, textvariable=self.date_var, width=12)
        date_entry.pack(side=tk.LEFT, padx=4)
        # Pressing Enter in the date field triggers Add
        date_entry.bind("<Return>", lambda event: self._add_entry())
        ttk.Button(row, text="Add", command=self._add_entry).pack(side=tk.LEFT, padx=8)

        # List
        list_frm = ttk.LabelFrame(cert_frame, text="Entries", padding=8)
        list_frm.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.listbox = tk.Listbox(list_frm, height=8, selectmode=tk.SINGLE)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(list_frm, orient=tk.VERTICAL, command=self.listbox.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scroll.set)
        btn_row = ttk.Frame(list_frm)
        btn_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btn_row, text="Edit selected", command=self._edit_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Remove selected", command=self._remove_selected).pack(side=tk.LEFT)

        # Generate
        gen_frm = ttk.Frame(cert_frame)
        gen_frm.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(gen_frm, text="Generate PDFs", command=self._generate).pack(side=tk.LEFT)

        # Log
        log_frm = ttk.LabelFrame(cert_frame, text="Log", padding=4)
        log_frm.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.log_area = scrolledtext.ScrolledText(log_frm, height=8, state=tk.NORMAL, wrap=tk.WORD)
        self.log_area.pack(fill=tk.BOTH, expand=True)

        # --- Settings tab ---
        sfrm = ttk.LabelFrame(settings_frame, text="Folder locations", padding=8)
        sfrm.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(sfrm, text="Output folder (certificates):").grid(row=0, column=0, padx=4, pady=6, sticky="w")
        self.settings_output_var = tk.StringVar(value=self.output_folder)
        ttk.Entry(sfrm, textvariable=self.settings_output_var, width=55).grid(row=0, column=1, padx=4, pady=6)
        ttk.Button(sfrm, text="Browse...", command=self._settings_browse_output).grid(row=0, column=2, padx=4, pady=6)

        ttk.Label(sfrm, text="PEP folder (PEP certificates):").grid(row=1, column=0, padx=4, pady=6, sticky="w")
        self.settings_pep_var = tk.StringVar(value=self.pep_folder)
        ttk.Entry(sfrm, textvariable=self.settings_pep_var, width=55).grid(row=1, column=1, padx=4, pady=6)
        ttk.Button(sfrm, text="Browse...", command=self._settings_browse_pep).grid(row=1, column=2, padx=4, pady=6)

        ttk.Button(settings_frame, text="Save folders", command=self._settings_save).pack(padx=8, pady=10, anchor="w")

    def _load_settings(self):
        """Read last used folders from disk, if available."""
        if not os.path.exists(SETTINGS_PATH):
            return "", ""
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("output_folder", ""), data.get("pep_folder", "")
        except Exception:
            return "", ""

    def _save_settings(self, output_folder: str, pep_folder: str):
        """Persist folders to disk for the next run."""
        data = {"output_folder": output_folder, "pep_folder": pep_folder}
        try:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            # Non-fatal if settings cannot be saved
            pass

    def _settings_browse_output(self):
        path = filedialog.askdirectory(title="Select Output folder")
        if path:
            self.settings_output_var.set(path)

    def _settings_browse_pep(self):
        path = filedialog.askdirectory(title="Select PEP folder")
        if path:
            self.settings_pep_var.set(path)

    def _settings_save(self):
        out = self.settings_output_var.get().strip()
        pep = self.settings_pep_var.get().strip()
        if not out or not pep:
            messagebox.showwarning("Missing folder", "Both Output and PEP folders are required.")
            return
        self.output_folder = out
        self.pep_folder = pep
        # Update display labels
        self.output_label.config(text=f"Output: {self.output_folder}")
        self.pep_label.config(text=f"PEP: {self.pep_folder}")
        # Save to disk
        self._save_settings(self.output_folder, self.pep_folder)
        messagebox.showinfo("Saved", "Folder locations have been updated and saved.")

    def _add_entry(self):
        name = self.name_var.get().strip().upper()
        date_str = self.date_var.get().strip()
        if not name:
            messagebox.showwarning("Missing name", "Enter a name.")
            return
        formatted = format_date(date_str)
        if not formatted:
            messagebox.showwarning("Invalid date", "Use DD.MM.YYYY format for the date.")
            return
        next_date = datetime.strptime(formatted, "%d.%m.%Y").replace(
            year=datetime.strptime(formatted, "%d.%m.%Y").year + 1
        ).strftime("%d.%m.%Y")
        self.entries.append((name, formatted, next_date))
        self.listbox.insert(tk.END, f"{name} | {formatted}")
        self.name_var.set("")
        self.date_var.set("")

    def _edit_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo("No selection", "Select an entry to edit.")
            return
        idx = sel[0]
        name, date, _ = self.entries[idx]
        # Simple edit dialog
        d = tk.Toplevel(self.root)
        d.title("Edit entry")
        d.geometry("360x120")
        d.transient(self.root)
        ttk.Label(d, text="Name:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        name_ent = ttk.Entry(d, width=40)
        name_ent.grid(row=0, column=1, padx=8, pady=6)
        name_ent.insert(0, name)
        ttk.Label(d, text="Date (DD.MM.YYYY):").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        date_ent = ttk.Entry(d, width=14)
        date_ent.grid(row=1, column=1, padx=8, pady=6)
        date_ent.insert(0, date)
        def save():
            new_name = name_ent.get().strip().upper()
            new_date_str = date_ent.get().strip()
            if not new_name:
                messagebox.showwarning("Missing name", "Enter a name.", parent=d)
                return
            new_date = format_date(new_date_str) if new_date_str else date
            if not new_date:
                messagebox.showwarning("Invalid date", "Use DD.MM.YYYY format.", parent=d)
                return
            next_date = datetime.strptime(new_date, "%d.%m.%Y").replace(
                year=datetime.strptime(new_date, "%d.%m.%Y").year + 1
            ).strftime("%d.%m.%Y")
            self.entries[idx] = (new_name, new_date, next_date)
            self.listbox.delete(idx)
            self.listbox.insert(idx, f"{new_name} | {new_date}")
            d.destroy()
        ttk.Button(d, text="Save", command=save).grid(row=2, column=1, padx=8, pady=8, sticky="w")

    def _remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo("No selection", "Select an entry to remove.")
            return
        idx = sel[0]
        del self.entries[idx]
        self.listbox.delete(idx)

    def _generate(self):
        if not self.entries:
            messagebox.showwarning("No entries", "Add at least one name/date entry.")
            return
        self.log_area.delete(1.0, tk.END)
        process_entries(
            self.entries,
            self.output_folder,
            self.pep_folder,
            TEMPLATE_PATH,
            self._log,
        )
        messagebox.showinfo("Done", "Certificate generation finished. Check the log.")


if __name__ == "__main__":
    app = MainApp()
