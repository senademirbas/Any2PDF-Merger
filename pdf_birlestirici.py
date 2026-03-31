"""
PDF Birleştirici v2 - Dosyaları PDF'e çevirip tek bir PDF'te birleştirir.

Desteklenen formatlar:
  - Görseller: JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP, ICO, SVG
  - Office: DOCX, DOC, XLSX, XLS, PPTX, PPT (Microsoft Office gerektirir)
  - Metin: TXT, CSV, LOG, MD, JSON, XML, HTML, HTM, PY, JS, CSS, JAVA, C, CPP, H
  - PDF: doğrudan birleştirilir

Gerekli kütüphaneler:
  pip install Pillow pypdf reportlab nbformat nbconvert
  (HTML ve IPYNB için Microsoft Edge yüklü olmalıdır)
"""

import os
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# --- Kütüphane kontrolleri ---
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Pillow kütüphanesi gerekli: pip install Pillow")
    sys.exit(1)

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("pypdf kütüphanesi gerekli: pip install pypdf")
    sys.exit(1)

# --- Dosya türleri ---
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp", ".ico"}
OFFICE_EXTENSIONS = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"}
TEXT_EXTENSIONS = {
    ".txt", ".csv", ".log", ".md", ".json", ".xml",
    ".py", ".js", ".css", ".java", ".c", ".cpp", ".h", ".ini", ".cfg", ".yaml", ".yml",
}
WEB_EXTENSIONS = {".html", ".htm"}
IPYNB_EXTENSIONS = {".ipynb"}
PDF_EXTENSIONS = {".pdf"}
ALL_EXTENSIONS = IMAGE_EXTENSIONS | OFFICE_EXTENSIONS | TEXT_EXTENSIONS | WEB_EXTENSIONS | IPYNB_EXTENSIONS | PDF_EXTENSIONS


# ==================== DÖNÜŞTÜRME FONKSİYONLARI ====================

def _get_font(size=11):
    """Sistemde kullanılabilir bir font bul."""
    font_paths = [
        "C:/Windows/Fonts/consola.ttf",
        "C:/Windows/Fonts/cour.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


def image_to_pdf(image_path: str, output_pdf: str):
    """Görsel dosyasını orijinal kalitesini koruyarak PDF'e çevirir.
    
    Görsel hiçbir şekilde yeniden boyutlandırılmaz veya sıkıştırılmaz.
    PDF sayfası görselin tam boyutuna göre oluşturulur.
    """
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    # Orijinal görseli oku (sadece boyut bilgisi için Pillow kullanıyoruz)
    img = Image.open(image_path)
    if hasattr(img, "n_frames") and img.n_frames > 1:
        img.seek(0)  # GIF ise ilk kare

    # Şeffaflık varsa beyaz arka plan ekle (PNG alpha desteği)
    needs_flatten = img.mode in ("RGBA", "LA", "P")
    if needs_flatten:
        flat = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "P":
            img = img.convert("RGBA")
        flat.paste(img, mask=img.split()[-1] if "A" in img.mode else None)
        # Geçici dosyaya kaydet (reportlab için)
        flat_path = output_pdf + "_flat.png"
        flat.save(flat_path, "PNG")
        img_source = flat_path
        img_w, img_h = flat.size
    else:
        img_source = image_path
        img_w, img_h = img.size
    
    img.close()

    # PDF sayfasını görselin piksel boyutunda oluştur (72 DPI referans)
    # Böylece görsel 1:1 piksel eşlemesiyle PDF'e gömülür
    page_w = float(img_w)
    page_h = float(img_h)

    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))
    c.drawImage(img_source, 0, 0, width=page_w, height=page_h,
                preserveAspectRatio=True, anchor='c')
    c.save()

    # Geçici flatten dosyasını temizle
    if needs_flatten and os.path.exists(flat_path):
        os.remove(flat_path)


def text_to_pdf(text_path: str, output_pdf: str):
    """Metin/kod dosyasını PDF'e çevirir."""
    content = None
    for enc in ["utf-8", "utf-8-sig", "cp1254", "latin-1"]:
        try:
            with open(text_path, "r", encoding=enc) as f:
                content = f.read()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    if content is None:
        with open(text_path, "rb") as f:
            content = f.read().decode("latin-1")

    font = _get_font(10)
    A4_W, A4_H = 595, 842
    MARGIN = 40
    line_h = 14
    usable_w = A4_W - 2 * MARGIN
    usable_h = A4_H - 2 * MARGIN
    lines_per_page = usable_h // line_h

    # Satırları hazırla ve word-wrap uygula
    raw_lines = content.replace("\t", "    ").split("\n")
    wrapped = []
    for line in raw_lines:
        if not line:
            wrapped.append("")
            continue
        while len(line) > 95:
            bp = line.rfind(" ", 0, 95)
            if bp == -1:
                bp = 95
            wrapped.append(line[:bp])
            line = line[bp:].lstrip()
        wrapped.append(line)

    pages = []
    for i in range(0, max(len(wrapped), 1), lines_per_page):
        chunk = wrapped[i : i + lines_per_page]
        page = Image.new("RGB", (A4_W, A4_H), (255, 255, 255))
        draw = ImageDraw.Draw(page)
        y = MARGIN
        for txt in chunk:
            draw.text((MARGIN, y), txt, fill=(30, 30, 30), font=font)
            y += line_h
        pages.append(page)

    if not pages:
        pages.append(Image.new("RGB", (A4_W, A4_H), (255, 255, 255)))

    pages[0].save(output_pdf, "PDF", resolution=72.0, save_all=True, append_images=pages[1:])


def office_to_pdf(file_path: str, output_pdf: str):
    """Office dosyasını (Word/Excel/PowerPoint) PDF'e çevirir."""
    ext = Path(file_path).suffix.lower()
    abs_in = os.path.abspath(file_path)
    abs_out = os.path.abspath(output_pdf)

    try:
        import comtypes.client
    except ImportError:
        # comtypes yoksa docx2pdf'i dene (sadece Word için)
        if ext in (".docx", ".doc"):
            try:
                from docx2pdf import convert as docx_convert
                docx_convert(abs_in, abs_out)
                return
            except ImportError:
                pass
        raise RuntimeError(
            f"Office dönüşümü için 'comtypes' gerekli:\n"
            f"  pip install comtypes\n"
            f"Ayrıca Microsoft Office yüklü olmalıdır."
        )

    if ext in (".docx", ".doc"):
        app = comtypes.client.CreateObject("Word.Application")
        app.Visible = False
        try:
            doc = app.Documents.Open(abs_in)
            doc.SaveAs(abs_out, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close()
        finally:
            app.Quit()

    elif ext in (".xlsx", ".xls"):
        app = comtypes.client.CreateObject("Excel.Application")
        app.Visible = False
        try:
            wb = app.Workbooks.Open(abs_in)
            wb.ExportAsFixedFormat(0, abs_out)  # 0 = xlTypePDF
            wb.Close(SaveChanges=False)
        finally:
            app.Quit()

    elif ext in (".pptx", ".ppt"):
        app = comtypes.client.CreateObject("PowerPoint.Application")
        try:
            prs = app.Presentations.Open(abs_in, ReadOnly=True, WithWindow=False)
            prs.SaveAs(abs_out, FileFormat=32)  # 32 = ppSaveAsPDF
            prs.Close()
        finally:
            app.Quit()


def get_edge_path():
    """Sistemde Microsoft Edge'in kurulu olduğu yolu bulur."""
    paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def web_to_pdf(html_path: str, output_pdf: str):
    """HTML dosyalarını Microsoft Edge (Headless) kullanarak PDF'e çevirir."""
    edge_path = get_edge_path()
    if not edge_path:
        raise RuntimeError("HTML/IPYNB dönüşümü için Microsoft Edge (msedge.exe) bulunamadı.")
    
    abs_in = os.path.abspath(html_path)
    abs_out = os.path.abspath(output_pdf)
    
    import subprocess
    cmd = [
        edge_path,
        "--headless",
        "--disable-gpu",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={abs_out}",
        abs_in
    ]
    try:
        subprocess.run(cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Edge PDF dönüşümü başarısız oldu: {e}")


def ipynb_to_pdf(ipynb_path: str, temp_dir: str, output_pdf: str):
    """Jupyter dosyasını HTML'ye, oradan da Edge ile görselleriyle birlikte tam teşekküllü PDF'e çevirir."""
    try:
        import nbformat
        from nbconvert import HTMLExporter
        import warnings
        warnings.filterwarnings("ignore")
    except ImportError:
        raise RuntimeError("Jupyter Notebook desteği için 'nbformat' ve 'nbconvert' gerekli:\n  pip install nbformat nbconvert")

    # 1. IPYNB'yi okuyup HTML'ye çevir (Tüm plot ve outputlar dahil)
    html_out = os.path.join(temp_dir, os.path.basename(ipynb_path) + "_temp.html")
    with open(ipynb_path, "r", encoding="utf-8") as f:
        nb = nbformat.read(f, as_version=4)
        
    html_exporter = HTMLExporter()
    (body, resources) = html_exporter.from_notebook_node(nb)
    with open(html_out, "w", encoding="utf-8") as f:
        f.write(body)
    
    # 2. HTML dosyasını MS Edge ile PDF'e çevir
    web_to_pdf(html_out, output_pdf)


def convert_file_to_pdf(file_path: str, temp_dir: str) -> str:
    """Herhangi bir dosyayı PDF'e çevirir."""
    ext = Path(file_path).suffix.lower()

    if ext in PDF_EXTENSIONS:
        return file_path

    uid = abs(hash(file_path)) % 10**8
    output = os.path.join(temp_dir, f"{Path(file_path).stem}_{uid}.pdf")

    if ext in IMAGE_EXTENSIONS:
        image_to_pdf(file_path, output)
    elif ext in WEB_EXTENSIONS:
        web_to_pdf(file_path, output)
    elif ext in IPYNB_EXTENSIONS:
        ipynb_to_pdf(file_path, temp_dir, output)
    elif ext in TEXT_EXTENSIONS:
        text_to_pdf(file_path, output)
    elif ext in OFFICE_EXTENSIONS:
        office_to_pdf(file_path, output)
    else:
        raise ValueError(f"Desteklenmeyen format: {ext}")

    return output


def merge_pdfs(pdf_paths: list, output_path: str):
    """PDF'leri birleştirir."""
    writer = PdfWriter()
    for p in pdf_paths:
        for page in PdfReader(p).pages:
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)


# ==================== MODERN GUI ====================

# Renk paleti
C = {
    "bg": "#1a1b2e",
    "surface": "#252742",
    "surface2": "#2d2f4e",
    "accent": "#7c5cfc",
    "accent_hover": "#9278ff",
    "danger": "#ff4d6a",
    "danger_hover": "#ff6b83",
    "success": "#34d399",
    "text": "#e8e8f0",
    "text2": "#9a9ab8",
    "border": "#3a3c5c",
    "orange": "#ff9f43",
}


class PDFBirlestiriciApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Birleştirici")
        self.root.geometry("720x620")
        self.root.minsize(600, 500)
        self.root.configure(bg=C["bg"])

        # Windows dark title bar
        try:
            self.root.update()
            from ctypes import windll
            windll.dwmapi.DwmSetWindowAttribute(
                windll.user32.GetParent(self.root.winfo_id()),
                20, byref := __import__("ctypes").byref(__import__("ctypes").c_int(2)), 4
            )
        except Exception:
            pass

        self.files = []  # (path, mtime) tuples
        self._build_ui()

    def _build_ui(self):
        # --- Stil ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=C["surface2"], background=C["accent"],
                        bordercolor=C["border"], lightcolor=C["accent"], darkcolor=C["accent"])

        # --- Başlık ---
        header = tk.Frame(self.root, bg=C["bg"])
        header.pack(fill="x", padx=24, pady=(20, 0))

        tk.Label(header, text="PDF Birleştirici", font=("Segoe UI", 20, "bold"),
                 bg=C["bg"], fg=C["text"]).pack(side="left")

        tk.Label(header, text="v2", font=("Segoe UI", 10),
                 bg=C["bg"], fg=C["accent"]).pack(side="left", padx=(6, 0), pady=(8, 0))

        # --- Alt başlık ---
        tk.Label(self.root, text="Dosyaları seç → otomatik sırala → tek PDF olarak birleştir",
                 font=("Segoe UI", 10), bg=C["bg"], fg=C["text2"]).pack(anchor="w", padx=24, pady=(4, 12))

        # --- Üst butonlar ---
        top_bar = tk.Frame(self.root, bg=C["bg"])
        top_bar.pack(fill="x", padx=24)

        self._btn(top_bar, "＋  Dosya Ekle", C["accent"], C["accent_hover"], self._add_files).pack(side="left")
        self._btn(top_bar, "✕  Kaldır", C["danger"], C["danger_hover"], self._remove_selected).pack(side="left", padx=(8, 0))
        self._btn(top_bar, "↑", C["surface2"], C["border"], self._move_up, w=3).pack(side="left", padx=(8, 0))
        self._btn(top_bar, "↓", C["surface2"], C["border"], self._move_down, w=3).pack(side="left", padx=(4, 0))
        self._btn(top_bar, "🗑  Temizle", C["surface2"], C["border"], self._clear_all).pack(side="right")

        # --- Sayaç ---
        self.counter_var = tk.StringVar(value="0 dosya")
        tk.Label(self.root, textvariable=self.counter_var, font=("Segoe UI", 9),
                 bg=C["bg"], fg=C["text2"]).pack(anchor="e", padx=24, pady=(6, 2))

        # --- Dosya listesi ---
        list_frame = tk.Frame(self.root, bg=C["border"], bd=0, highlightthickness=0)
        list_frame.pack(fill="both", expand=True, padx=24, pady=(0, 12))

        inner = tk.Frame(list_frame, bg=C["surface"], bd=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        scrollbar = tk.Scrollbar(inner, bg=C["surface2"], troughcolor=C["surface"],
                                 activebackground=C["accent"], bd=0, highlightthickness=0)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            inner, font=("Consolas", 10), selectmode="single",
            yscrollcommand=scrollbar.set,
            bg=C["surface"], fg=C["text"],
            selectbackground=C["accent"], selectforeground="white",
            relief="flat", bd=0, highlightthickness=0,
            activestyle="none",
        )
        self.listbox.pack(fill="both", expand=True, padx=(8, 0), pady=4)
        scrollbar.config(command=self.listbox.yview)

        # --- Bilgi ---
        info = tk.Frame(self.root, bg=C["bg"])
        info.pack(fill="x", padx=24)
        exts = "jpg png bmp pdf docx xlsx pptx txt py html ipynb"
        tk.Label(info, text=f"Desteklenen: {exts}",
                 font=("Segoe UI", 8), bg=C["bg"], fg=C["text2"], wraplength=680, justify="left"
                 ).pack(side="left")

        # --- Birleştir butonu ---
        merge_frame = tk.Frame(self.root, bg=C["bg"])
        merge_frame.pack(fill="x", padx=24, pady=(8, 20))

        self.merge_btn = tk.Button(
            merge_frame, text="⚡  PDF Olarak Birleştir", font=("Segoe UI", 13, "bold"),
            bg=C["orange"], fg="white", activebackground="#ffb366", activeforeground="white",
            relief="flat", bd=0, cursor="hand2", pady=10,
            command=self._merge,
        )
        self.merge_btn.pack(fill="x")
        self.merge_btn.bind("<Enter>", lambda e: self.merge_btn.config(bg="#ffb366"))
        self.merge_btn.bind("<Leave>", lambda e: self.merge_btn.config(bg=C["orange"]))

    def _btn(self, parent, text, bg, hover_bg, cmd, w=None):
        """Küçük modern buton oluştur."""
        kw = dict(text=text, font=("Segoe UI", 9), bg=bg, fg="white",
                  activebackground=hover_bg, activeforeground="white",
                  relief="flat", bd=0, cursor="hand2", padx=12, pady=5, command=cmd)
        if w:
            kw["width"] = w
        btn = tk.Button(parent, **kw)
        btn.bind("<Enter>", lambda e, b=btn, c=hover_bg: b.config(bg=c))
        btn.bind("<Leave>", lambda e, b=btn, c=bg: b.config(bg=c))
        return btn

    def _update_counter(self):
        self.counter_var.set(f"{len(self.files)} dosya")

    def _refresh_listbox(self):
        """Listbox'ı dosya listesinden yenile."""
        self.listbox.delete(0, tk.END)
        for path, mtime in self.files:
            name = os.path.basename(path)
            ext = Path(path).suffix.lower()
            from datetime import datetime
            date_str = datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M")
            self.listbox.insert(tk.END, f"  {name}   ·   {date_str}   ({ext})")
        self._update_counter()

    def _add_files(self):
        exts = " ".join(f"*{e}" for e in sorted(ALL_EXTENSIONS))
        paths = filedialog.askopenfilenames(
            title="Dosya Seç",
            filetypes=[
                ("Tüm desteklenen dosyalar", exts),
                ("Görseller", " ".join(f"*{e}" for e in sorted(IMAGE_EXTENSIONS))),
                ("Office dosyaları", " ".join(f"*{e}" for e in sorted(OFFICE_EXTENSIONS))),
                ("Jupyter dosyaları", "*.ipynb"),
                ("Web sayfaları", "*.html *.htm"),
                ("PDF", "*.pdf"),
                ("Metin dosyaları", " ".join(f"*{e}" for e in sorted(TEXT_EXTENSIONS))),
                ("Tüm dosyalar", "*.*"),
            ],
        )
        if not paths:
            return

        existing = {p for p, _ in self.files}
        for p in paths:
            if p not in existing:
                mtime = os.path.getmtime(p)
                self.files.append((p, mtime))

        # Değişiklik tarihine göre sırala (en eski önce)
        self.files.sort(key=lambda x: x[1])
        self._refresh_listbox()

    def _remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.files.pop(idx)
        self._refresh_listbox()
        if self.files:
            self.listbox.select_set(min(idx, len(self.files) - 1))

    def _move_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        self.files[i], self.files[i - 1] = self.files[i - 1], self.files[i]
        self._refresh_listbox()
        self.listbox.select_set(i - 1)

    def _move_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.files) - 1:
            return
        i = sel[0]
        self.files[i], self.files[i + 1] = self.files[i + 1], self.files[i]
        self._refresh_listbox()
        self.listbox.select_set(i + 1)

    def _clear_all(self):
        self.files.clear()
        self._refresh_listbox()

    def _merge(self):
        if not self.files:
            messagebox.showwarning("Uyarı", "Lütfen en az bir dosya ekleyin.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Çıktı PDF'i Kaydet",
            defaultextension=".pdf",
            filetypes=[("PDF dosyası", "*.pdf")],
            initialfile="birlestirilmis.pdf",
        )
        if not output_path:
            return

        # İlerleme overlay'i
        overlay = tk.Frame(self.root, bg=C["bg"])
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)

        center = tk.Frame(overlay, bg=C["surface"], bd=0, highlightthickness=1, highlightbackground=C["border"])
        center.place(relx=0.5, rely=0.5, anchor="center", width=420, height=160)

        tk.Label(center, text="⏳  İşleniyor...", font=("Segoe UI", 14, "bold"),
                 bg=C["surface"], fg=C["text"]).pack(pady=(20, 8))

        progress = ttk.Progressbar(center, length=360, mode="determinate",
                                   style="Custom.Horizontal.TProgressbar")
        progress.pack(pady=(0, 8))

        status_var = tk.StringVar(value="Hazırlanıyor...")
        tk.Label(center, textvariable=status_var, font=("Segoe UI", 9),
                 bg=C["surface"], fg=C["text2"]).pack()

        progress["maximum"] = len(self.files)
        self.root.update()

        def run_merge():
            try:
                with tempfile.TemporaryDirectory() as temp_dir:
                    pdf_paths = []
                    for i, (file_path, _) in enumerate(self.files):
                        name = os.path.basename(file_path)
                        status_var.set(f"Dönüştürülüyor: {name}")
                        self.root.update_idletasks()

                        pdf = convert_file_to_pdf(file_path, temp_dir)
                        pdf_paths.append(pdf)

                        progress["value"] = i + 1
                        self.root.update_idletasks()

                    status_var.set("Birleştiriliyor...")
                    self.root.update_idletasks()
                    merge_pdfs(pdf_paths, output_path)

                overlay.destroy()
                messagebox.showinfo(
                    "Başarılı ✅",
                    f"{len(self.files)} dosya birleştirildi.\n\n📄 {output_path}"
                )
            except Exception as e:
                overlay.destroy()
                messagebox.showerror("Hata ❌", f"Bir hata oluştu:\n\n{str(e)}")

        # UI donmasın diye update_idletasks kullanıyoruz
        self.root.after(100, run_merge)


def main():
    root = tk.Tk()
    PDFBirlestiriciApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
