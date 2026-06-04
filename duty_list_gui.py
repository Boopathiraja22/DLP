"""
Duty List PDF Parser — Portable GUI
=====================================
Runs on Windows, macOS, and Linux.
Uses only tkinter (built into Python) + auto-installs pdfplumber & reportlab.

Run:
    python duty_list_gui.py
"""

import subprocess, sys, os, re, zipfile, threading
from collections import defaultdict

# ── Auto-install dependencies ─────────────────────────────────────────────────
def install(pkg):
    subprocess.run([sys.executable, "-m", "pip", "install", pkg, "-q"], check=True)

try:
    import pdfplumber
except ImportError:
    install("pdfplumber"); import pdfplumber

try:
    from reportlab.lib.pagesizes import A4
except ImportError:
    install("reportlab")

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# ── PDF logic ─────────────────────────────────────────────────────────────────
DATE_PATTERN    = re.compile(r"Exam Date[:\s]+(\d{2}\.\d{2}\.\d{4})")
TIMING_PATTERN  = re.compile(r"Timings?[:\s]+([\d:APM\s\-]+)", re.IGNORECASE)
FACULTY_PATTERN = re.compile(r"^\s*(\d+)\s+(?:(\S+)\s+)?([A-Z][a-zA-Z.\s]+)")
SKIP_KEYWORDS   = [
    "Faculty of", "Ramaiah", "Room Superintendent",
    "Exam Date", "Sl. No", "Dy. Supt", "Chief Supt",
    "Dy. CoE", "CReport", "Alternative", "Reporting"
]

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def extract_duty_data(path):
    faculty_duties = defaultdict(list)
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            date_m = DATE_PATTERN.search(text)
            exam_date = date_m.group(1) if date_m else "Unknown Date"
            time_m = TIMING_PATTERN.search(text)
            timing = time_m.group(1).strip() if time_m else "Unknown Time"
            for line in text.splitlines():
                if any(kw in line for kw in SKIP_KEYWORDS):
                    continue
                m = FACULTY_PATTERN.match(line)
                if m:
                    name = re.sub(r"\s{2,}.*$", "", m.group(3).strip()).strip()
                    if len(name) >= 3:
                        faculty_duties[name].append({"date": exam_date, "timing": timing})
    return faculty_duties

HEADER_BG = colors.HexColor("#003366")
ROW_ALT   = colors.HexColor("#EAF1FB")
_styles   = getSampleStyleSheet()

def make_faculty_pdf(faculty_name, duties, filepath):
    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )
    story = []
    story.append(Paragraph("Ramaiah University of Applied Sciences", ParagraphStyle(
        "Uni", parent=_styles["Normal"],
        fontSize=12, textColor=HEADER_BG, spaceAfter=2,
        alignment=1, fontName="Helvetica-Bold"
    )))
    story.append(Paragraph("Ramaiah Technology Campus (RTC), Peenya", ParagraphStyle(
        "Campus", parent=_styles["Normal"],
        fontSize=9, textColor=colors.HexColor("#555555"), spaceAfter=10, alignment=1
    )))
    story.append(Paragraph(f"<b>Faculty Name:</b> {faculty_name}", ParagraphStyle(
        "FName", parent=_styles["Normal"],
        fontSize=12, textColor=colors.HexColor("#111111"), spaceAfter=6, alignment=1
    )))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Exam Duty Schedule", ParagraphStyle(
        "SchHead", parent=_styles["Heading2"],
        fontSize=11, textColor=HEADER_BG, spaceBefore=8, spaceAfter=6
    )))

    table_data = [["S.No", "Exam Date", "Timings"]]
    for i, d in enumerate(duties, 1):
        table_data.append([str(i), d["date"], d["timing"]])

    t = Table(table_data, colWidths=[2*cm, 5*cm, 9*cm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  HEADER_BG),
        ("TEXTCOLOR",     (0,0), (-1,0),  colors.white),
        ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,0),  10),
        ("ALIGN",         (0,0), (-1,0),  "CENTER"),
        ("TOPPADDING",    (0,0), (-1,0),  8),
        ("BOTTOMPADDING", (0,0), (-1,0),  8),
        ("FONTNAME",      (0,1), (-1,-1), "Helvetica"),
        ("FONTSIZE",      (0,1), (-1,-1), 9),
        ("ALIGN",         (0,1), (1,-1),  "CENTER"),
        ("TOPPADDING",    (0,1), (-1,-1), 6),
        ("BOTTOMPADDING", (0,1), (-1,-1), 6),
        ("GRID",          (0,0), (-1,-1), 0.5, HEADER_BG),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, ROW_ALT]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        f"<b>Total exam duties assigned: {len(duties)}</b>",
        ParagraphStyle("sm", parent=_styles["Normal"], fontSize=9)
    ))
    doc.build(story)

# ── GUI ───────────────────────────────────────────────────────────────────────
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Duty List PDF Parser")
        self.resizable(False, False)
        self.configure(bg="#f0f4f8")
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 520, 420
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#003366", pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Duty List PDF Parser",
                 bg="#003366", fg="white",
                 font=("Helvetica", 16, "bold")).pack()
        tk.Label(hdr, text="Ramaiah University of Applied Sciences",
                 bg="#003366", fg="#aac8ff",
                 font=("Helvetica", 9)).pack()

        body = tk.Frame(self, bg="#f0f4f8", padx=24, pady=18)
        body.pack(fill="both", expand=True)

        # ── Input PDF ─────────────────────────────────────────────────────────
        tk.Label(body, text="Input PDF", bg="#f0f4f8",
                 font=("Helvetica", 10, "bold"), anchor="w").pack(fill="x")

        row1 = tk.Frame(body, bg="#f0f4f8")
        row1.pack(fill="x", pady=(4, 12))
        self.pdf_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.pdf_var, width=44,
                 font=("Helvetica", 9)).pack(side="left", padx=(0,8))
        tk.Button(row1, text="Browse…", command=self._browse_pdf,
                  bg="#003366", fg="white", font=("Helvetica", 9),
                  relief="flat", padx=10, cursor="hand2").pack(side="left")

        # ── Output Folder ─────────────────────────────────────────────────────
        tk.Label(body, text="Output Folder", bg="#f0f4f8",
                 font=("Helvetica", 10, "bold"), anchor="w").pack(fill="x")

        row2 = tk.Frame(body, bg="#f0f4f8")
        row2.pack(fill="x", pady=(4, 18))
        self.out_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.out_var, width=44,
                 font=("Helvetica", 9)).pack(side="left", padx=(0,8))
        tk.Button(row2, text="Browse…", command=self._browse_out,
                  bg="#003366", fg="white", font=("Helvetica", 9),
                  relief="flat", padx=10, cursor="hand2").pack(side="left")

        # ── Progress bar ──────────────────────────────────────────────────────
        self.progress = ttk.Progressbar(body, mode="determinate", length=460)
        self.progress.pack(fill="x", pady=(0, 6))

        # ── Log box ───────────────────────────────────────────────────────────
        self.log = tk.Text(body, height=7, font=("Courier", 8),
                           bg="#1e1e2e", fg="#cdd6f4",
                           relief="flat", state="disabled")
        self.log.pack(fill="x", pady=(0, 14))

        # ── Generate button ───────────────────────────────────────────────────
        self.btn = tk.Button(body, text="⚙  Generate PDFs",
                             command=self._start,
                             bg="#0055aa", fg="white",
                             font=("Helvetica", 11, "bold"),
                             relief="flat", pady=8, cursor="hand2")
        self.btn.pack(fill="x")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Select Duty List PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_var.set(path)
            if not self.out_var.get():
                self.out_var.set(os.path.join(os.path.dirname(path), "faculty_duties"))

    def _browse_out(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.out_var.set(path)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _start(self):
        pdf  = self.pdf_var.get().strip()
        outd = self.out_var.get().strip()

        if not pdf or not os.path.isfile(pdf):
            messagebox.showerror("Error", "Please select a valid PDF file."); return
        if not outd:
            messagebox.showerror("Error", "Please select an output folder."); return

        self.btn.configure(state="disabled", text="Processing…")
        self.progress["value"] = 0
        self.log.configure(state="normal"); self.log.delete("1.0", "end"); self.log.configure(state="disabled")

        threading.Thread(target=self._run, args=(pdf, outd), daemon=True).start()

    def _run(self, pdf_path, output_dir):
        try:
            self._log(f"📄 Parsing: {os.path.basename(pdf_path)}")
            faculty_duties = extract_duty_data(pdf_path)
            total = len(faculty_duties)
            self._log(f"👥 Found {total} faculty member(s)\n")

            os.makedirs(output_dir, exist_ok=True)
            created = []

            for idx, (faculty_name, duties) in enumerate(sorted(faculty_duties.items()), 1):
                filename = sanitize_filename(faculty_name) + ".pdf"
                filepath = os.path.join(output_dir, filename)
                make_faculty_pdf(faculty_name, duties, filepath)
                created.append(filepath)
                self._log(f"  ✓  {filename}  ({len(duties)} duty slot(s))")
                pct = int(idx / total * 100)
                self.progress["value"] = pct
                self.update_idletasks()

            # Zip
            zip_path = os.path.join(output_dir, "faculty_duties.zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                for fp in created:
                    zf.write(fp, os.path.basename(fp))

            self._log(f"\n📦 ZIP saved: {zip_path}")
            self._log(f"✅ Done! {total} PDF(s) generated.")
            self.progress["value"] = 100

            messagebox.showinfo(
                "Done ✅",
                f"{total} PDF(s) saved to:\n{output_dir}\n\nZIP: faculty_duties.zip"
            )

            # Open output folder
            if sys.platform == "win32":
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", output_dir])
            else:
                subprocess.Popen(["xdg-open", output_dir])

        except Exception as e:
            self._log(f"\n❌ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn.configure(state="normal", text="⚙  Generate PDFs")


if __name__ == "__main__":
    app = App()
    app.mainloop()
