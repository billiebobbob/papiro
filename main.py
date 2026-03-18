"""
EtiquetaTron - Separa etiquetas de PDFs, preview y manda a imprimir
Desarrollado para Mawida Dispensario
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io
import os
import re
import sys
import subprocess
import threading

DPI = 300

def mm_to_px(mm):
    return int(mm * DPI / 25.4)

CANVAS_WIDTH_MM = 62

LABEL_FORMATS = {
    "Envio": {
        "description": "Envio (15.2cm x 3cm)",
        "tape_width_mm": CANVAS_WIDTH_MM,  # 62mm ancho de cinta = ancho de imagen
        "cut_length_mm": 152,               # largo de corte = alto de imagen
        "tape_mm": CANVAS_WIDTH_MM,
    },
    "Producto": {
        "description": "Producto (6.2cm x 3.3cm)",
        "tape_width_mm": CANVAS_WIDTH_MM,  # 62mm ancho de cinta
        "cut_length_mm": CANVAS_WIDTH_MM,  # 62mm corte (canvas cuadrado, etiqueta centrada)
        "tape_mm": CANVAS_WIDTH_MM,
    },
}

PDF_LABEL_SPACING_PTS = 130
PDF_LABEL_HEIGHT_PTS = 120
PDF_MARGIN_TOP_PTS = 5
PDF_MARGIN_SIDES_PTS = 10
PDF_MAX_LABELS_PER_PAGE = 6

PREVIEW_MAX_WIDTH = 520
PREVIEW_MAX_HEIGHT = 160

# Color palette
BG_DARK = "#1a1a2e"
BG_CARD = "#16213e"
BG_CARD_LIGHT = "#1c2a4a"
ACCENT_BLUE = "#0f7dff"
ACCENT_GREEN = "#00c853"
ACCENT_ORANGE = "#ff6d00"
ACCENT_CYAN = "#00e5ff"
TEXT_PRIMARY = "#e8eaf6"
TEXT_SECONDARY = "#7986cb"
TEXT_MUTED = "#455a80"
BORDER_COLOR = "#263159"


def get_printers():
    printers = []
    try:
        if sys.platform == 'darwin' or sys.platform.startswith('linux'):
            result = subprocess.run(['lpstat', '-p'], capture_output=True, text=True, timeout=5)
            for line in result.stdout.strip().split('\n'):
                if line.startswith('printer '):
                    printers.append(line.split()[1])
        elif sys.platform == 'win32':
            result = subprocess.run(['wmic', 'printer', 'get', 'name'], capture_output=True, text=True, timeout=5)
            for line in result.stdout.strip().split('\n')[1:]:
                name = line.strip()
                if name:
                    printers.append(name)
    except Exception:
        pass
    return printers


def get_default_printer():
    try:
        if sys.platform == 'darwin' or sys.platform.startswith('linux'):
            result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True, timeout=5)
            for line in result.stdout.strip().split('\n'):
                if 'default' in line.lower() and ':' in line:
                    return line.split(':')[-1].strip()
    except Exception:
        pass
    return None


def print_image(filepath, printer_name=None, tape_mm=62):
    """Imprime una imagen en la impresora seleccionada.
    tape_mm: ancho de cinta (62mm para Brother QL-800), usa rollo continuo.
    """
    if sys.platform in ('darwin',) or sys.platform.startswith('linux'):
        cmd = ['lp']
        if printer_name:
            cmd.extend(['-d', printer_name])
        # Usar rollo continuo (ej: "62mm") - la impresora corta segun largo de imagen
        cmd.extend(['-o', f'media={tape_mm}mm'])
        cmd.extend(['-o', 'fit-to-page'])
        cmd.append(filepath)
        return subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    elif sys.platform == 'win32':
        return _print_windows(filepath, printer_name, tape_mm)


def _print_windows(filepath, printer_name, tape_mm=62):
    """Imprime en Windows via PowerShell/.NET con tamaño de papel exacto.
    Calcula las dimensiones desde el DPI de la imagen para que la Brother QL-800
    corte al largo correcto (152mm para envio, 62mm para producto).
    """
    fp = filepath.replace('\\', '\\\\')
    pn = (printer_name or "").replace("'", "''")

    ps_script = f'''
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$bitmap = [System.Drawing.Image]::FromFile("{fp}")

$pd = New-Object System.Drawing.Printing.PrintDocument
$pd.PrinterSettings.PrinterName = "{pn}"

# Calcular tamaño exacto desde la imagen (en centesimas de pulgada)
# Imagen de 62mm x 152mm a 300dpi → el papel debe ser exactamente ese tamaño
$wHundredths = [int]($bitmap.Width / $bitmap.HorizontalResolution * 100)
$hHundredths = [int]($bitmap.Height / $bitmap.VerticalResolution * 100)
$customSize = New-Object System.Drawing.Printing.PaperSize("EtiquetaTron", $wHundredths, $hHundredths)
$pd.DefaultPageSettings.PaperSize = $customSize
$pd.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(0, 0, 0, 0)

$pd.add_PrintPage({{
    param($sender, $e)
    $destRect = New-Object System.Drawing.RectangleF(0, 0, $e.PageBounds.Width, $e.PageBounds.Height)
    $e.Graphics.DrawImage($bitmap, $destRect)
}})

$pd.Print()
$bitmap.Dispose()
$pd.Dispose()
Write-Output "OK"
'''
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
            capture_output=True, text=True, timeout=30
        )
        return result
    except Exception as e:
        # Fallback: abrir con visor de imagenes
        try:
            os.startfile(filepath, 'print')
        except Exception:
            pass
        return None


class EtiquetaSeparador(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("EtiquetaTron")
        self.geometry("640x600")
        self.resizable(True, True)
        self.minsize(600, 540)
        self.configure(fg_color=BG_DARK)

        ctk.set_appearance_mode("dark")

        self.pdf_path = None
        self.processing = False
        self.selected_format = ctk.StringVar(value="Envio")
        self.selected_printer = ctk.StringVar(value="")

        self.preview_labels = []
        self.preview_index = 0
        self.preview_tk_image = None
        self.last_output_dir = None

        self.printers = get_printers()
        default = get_default_printer()
        if default and default in self.printers:
            self.selected_printer.set(default)
        elif self.printers:
            self.selected_printer.set(self.printers[0])

        self.create_widgets()
        self.center_window()

        self.bind('<Left>', self._preview_prev)
        self.bind('<Right>', self._preview_next)

    def center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f'{w}x{h}+{x}+{y}')

    def _load_logo(self):
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(base_path, 'logo.png')
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                max_width = 38
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), Image.LANCZOS)
                return ctk.CTkImage(light_image=img, dark_image=img, size=(max_width, new_height))
        except Exception:
            pass
        return None

    def _get_canvas_size(self):
        """Returns (width_px, height_px) for the image.
        width = largo de etiqueta (152mm envio), height = ancho cinta (62mm)
        Imagen LANDSCAPE: el PowerShell crea PaperSize custom desde el DPI."""
        fmt = LABEL_FORMATS[self.selected_format.get()]
        return mm_to_px(fmt["cut_length_mm"]), mm_to_px(fmt["tape_width_mm"])

    def _on_format_change(self, *args):
        fmt = LABEL_FORMATS[self.selected_format.get()]
        self.format_detail.configure(text=f"{fmt['tape_width_mm']}x{fmt['cut_length_mm']}mm")

    def _refresh_printers(self):
        self.printers = get_printers()
        if self.printers:
            self.printer_menu.configure(values=self.printers)
            if self.selected_printer.get() not in self.printers:
                default = get_default_printer()
                self.selected_printer.set(default if default and default in self.printers else self.printers[0])
        else:
            self.printer_menu.configure(values=["Sin impresora"])
            self.selected_printer.set("Sin impresora")

    def create_widgets(self):
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=BG_DARK)
        self.main_frame.pack(fill="both", expand=True, padx=16, pady=12)

        # === HEADER ===
        header = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        self.logo_image = self._load_logo()
        if self.logo_image:
            ctk.CTkLabel(header, image=self.logo_image, text="").pack(side="left", padx=(0, 10))

        ctk.CTkLabel(header, text="EtiquetaTron",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")

        ctk.CTkLabel(header, text="Mawida Dispensario",
                     font=ctk.CTkFont(size=10),
                     text_color=TEXT_MUTED).pack(side="right", padx=(0, 4))

        # === CONFIG CARD ===
        config_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=BORDER_COLOR)
        config_card.pack(fill="x", pady=(0, 8))

        # Row 1: Format
        row1 = ctk.CTkFrame(config_card, fg_color="transparent")
        row1.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(row1, text="FORMATO", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=TEXT_MUTED).pack(side="left", padx=(0, 12))

        for fmt_name in LABEL_FORMATS:
            ctk.CTkRadioButton(
                row1, text=fmt_name, variable=self.selected_format,
                value=fmt_name, font=ctk.CTkFont(size=11),
                text_color=TEXT_PRIMARY,
                command=self._on_format_change,
                radiobutton_width=16, radiobutton_height=16,
                fg_color=ACCENT_BLUE, hover_color=ACCENT_BLUE,
                border_color=TEXT_MUTED
            ).pack(side="left", padx=8)

        self.format_detail = ctk.CTkLabel(row1, text="140x62mm",
                                          font=ctk.CTkFont(size=10),
                                          text_color=ACCENT_CYAN)
        self.format_detail.pack(side="right")
        self._on_format_change()

        # Divider
        ctk.CTkFrame(config_card, height=1, fg_color=BORDER_COLOR).pack(fill="x", padx=14, pady=4)

        # Row 2: Printer
        row2 = ctk.CTkFrame(config_card, fg_color="transparent")
        row2.pack(fill="x", padx=14, pady=(4, 10))

        ctk.CTkLabel(row2, text="IMPRESORA", font=ctk.CTkFont(size=9, weight="bold"),
                     text_color=TEXT_MUTED).pack(side="left", padx=(0, 8))

        printer_values = self.printers if self.printers else ["Sin impresora"]
        self.printer_menu = ctk.CTkOptionMenu(
            row2, variable=self.selected_printer,
            values=printer_values, font=ctk.CTkFont(size=10),
            width=280, height=28, corner_radius=6,
            fg_color=BG_CARD_LIGHT, button_color=ACCENT_BLUE,
            button_hover_color="#0d6efd",
            text_color=TEXT_PRIMARY
        )
        self.printer_menu.pack(side="left", padx=4)

        ctk.CTkButton(
            row2, text="Refresh", width=60, height=26, corner_radius=6,
            font=ctk.CTkFont(size=9), command=self._refresh_printers,
            fg_color="transparent", border_width=1, border_color=TEXT_MUTED,
            text_color=TEXT_SECONDARY, hover_color=BG_CARD_LIGHT
        ).pack(side="right")

        # === LOAD BUTTON ===
        load_row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        load_row.pack(fill="x", pady=(0, 6))

        self.select_button = ctk.CTkButton(
            load_row, text="Cargar PDF",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=36, width=160, corner_radius=8,
            fg_color=ACCENT_BLUE, hover_color="#0d6efd",
            text_color="white",
            command=self.load_and_preview
        )
        self.select_button.pack(side="left")

        self.file_label = ctk.CTkLabel(
            load_row, text="Ningun archivo seleccionado",
            font=ctk.CTkFont(size=10), text_color=TEXT_MUTED
        )
        self.file_label.pack(side="left", padx=(10, 0))

        # === PREVIEW CARD ===
        preview_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12,
                                    border_width=1, border_color=BORDER_COLOR)
        preview_card.pack(fill="x", pady=(0, 8))

        # Nav bar
        nav = ctk.CTkFrame(preview_card, fg_color="transparent")
        nav.pack(fill="x", padx=10, pady=(8, 0))

        self.prev_btn = ctk.CTkButton(
            nav, text="<", width=30, height=26, corner_radius=6,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=BG_CARD_LIGHT, hover_color=BORDER_COLOR,
            text_color=TEXT_SECONDARY,
            command=lambda: self._preview_prev(None), state="disabled"
        )
        self.prev_btn.pack(side="left")

        self.preview_counter_label = ctk.CTkLabel(
            nav, text="Sin etiquetas",
            font=ctk.CTkFont(size=11), text_color=TEXT_SECONDARY
        )
        self.preview_counter_label.pack(side="left", expand=True)

        self.preview_name_label = ctk.CTkLabel(
            nav, text="", font=ctk.CTkFont(size=11, weight="bold"),
            text_color=ACCENT_CYAN
        )
        self.preview_name_label.pack(side="left", expand=True)

        self.next_btn = ctk.CTkButton(
            nav, text=">", width=30, height=26, corner_radius=6,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=BG_CARD_LIGHT, hover_color=BORDER_COLOR,
            text_color=TEXT_SECONDARY,
            command=lambda: self._preview_next(None), state="disabled"
        )
        self.next_btn.pack(side="right")

        # Preview image
        self.preview_image_label = ctk.CTkLabel(
            preview_card, text="Carga un PDF para previsualizar",
            font=ctk.CTkFont(size=10), text_color=TEXT_MUTED,
            height=PREVIEW_MAX_HEIGHT
        )
        self.preview_image_label.pack(padx=10, pady=(4, 10))

        # === ACTION BUTTONS ===
        action_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        action_frame.pack(fill="x", pady=(0, 6))

        # Center buttons using inner frame
        btn_container = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_container.pack(anchor="center")

        self.save_button = ctk.CTkButton(
            btn_container, text="Guardar",
            font=ctk.CTkFont(size=11, weight="bold"),
            height=34, width=130, corner_radius=8,
            fg_color="transparent", border_width=1,
            border_color=ACCENT_BLUE, text_color=ACCENT_BLUE,
            hover_color=BG_CARD,
            command=lambda: self.save_and_print(send_to_printer=False),
            state="disabled"
        )
        self.save_button.pack(side="left", padx=4)

        self.print_button = ctk.CTkButton(
            btn_container, text="Guardar + Imprimir",
            font=ctk.CTkFont(size=11, weight="bold"),
            height=34, width=170, corner_radius=8,
            fg_color=ACCENT_GREEN, hover_color="#00a844",
            text_color="#0a0a0a",
            command=lambda: self.save_and_print(send_to_printer=True),
            state="disabled"
        )
        self.print_button.pack(side="left", padx=4)

        self.print_all_button = ctk.CTkButton(
            btn_container, text="Imprimir Todas",
            font=ctk.CTkFont(size=11, weight="bold"),
            height=34, width=140, corner_radius=8,
            fg_color=ACCENT_ORANGE, hover_color="#e65100",
            text_color="#0a0a0a",
            command=lambda: self.save_and_print(send_to_printer=True),
            state="disabled"
        )
        self.print_all_button.pack(side="left", padx=4)

        # === PROGRESS ===
        self.progress_bar = ctk.CTkProgressBar(
            self.main_frame, height=3, corner_radius=2,
            progress_color=ACCENT_BLUE, fg_color=BORDER_COLOR
        )
        self.progress_bar.pack(fill="x", padx=4, pady=(0, 6))
        self.progress_bar.set(0)

        # === LOG ===
        log_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=10,
                                border_width=1, border_color=BORDER_COLOR)
        log_card.pack(fill="both", expand=True)

        mono = "Menlo" if sys.platform == 'darwin' else ("Consolas" if sys.platform == 'win32' else "DejaVu Sans Mono")
        self.result_text = ctk.CTkTextbox(
            log_card, font=ctk.CTkFont(size=9, family=mono),
            fg_color="transparent", text_color=TEXT_SECONDARY, height=50
        )
        self.result_text.pack(fill="both", expand=True, padx=8, pady=8)
        self.result_text.insert("1.0", "Listo para trabajar...")
        self.result_text.configure(state="disabled")

    # --- Preview ---
    def _preview_prev(self, event):
        if self.preview_labels:
            self.preview_index = (self.preview_index - 1) % len(self.preview_labels)
            self._show_preview()

    def _preview_next(self, event):
        if self.preview_labels:
            self.preview_index = (self.preview_index + 1) % len(self.preview_labels)
            self._show_preview()

    def _show_preview(self):
        if not self.preview_labels:
            self.preview_image_label.configure(image=None, text="Carga un PDF para previsualizar")
            self.preview_counter_label.configure(text="Sin etiquetas")
            self.preview_name_label.configure(text="")
            self.prev_btn.configure(state="disabled")
            self.next_btn.configure(state="disabled")
            return

        data = self.preview_labels[self.preview_index]
        img = data['image']
        img_w, img_h = img.size
        scale = min(PREVIEW_MAX_WIDTH / img_w, PREVIEW_MAX_HEIGHT / img_h, 1.0)
        dw, dh = max(int(img_w * scale), 1), max(int(img_h * scale), 1)

        img_resized = img.resize((dw, dh), Image.LANCZOS)
        self.preview_tk_image = ctk.CTkImage(light_image=img_resized, dark_image=img_resized, size=(dw, dh))
        self.preview_image_label.configure(image=self.preview_tk_image, text="", height=PREVIEW_MAX_HEIGHT)

        total = len(self.preview_labels)
        self.preview_counter_label.configure(text=f"{self.preview_index + 1} / {total}")
        self.preview_name_label.configure(text=data['venta'])

        state = "normal" if total > 1 else "disabled"
        self.prev_btn.configure(state=state)
        self.next_btn.configure(state=state)

    # --- Load ---
    def load_and_preview(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar PDF de etiquetas",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        if not file_path:
            return
        self.pdf_path = file_path
        self.file_label.configure(text=os.path.basename(file_path), text_color=TEXT_PRIMARY)
        self.log_message(f"Cargando: {os.path.basename(file_path)}")
        self.progress_bar.set(0)
        self.select_button.configure(state="disabled")
        thread = threading.Thread(target=self._load_preview_thread)
        thread.daemon = True
        thread.start()

    def _load_preview_thread(self):
        doc = None
        try:
            canvas_w, canvas_h = self._get_canvas_size()
            doc = fitz.open(self.pdf_path)
            total_pages = len(doc)
            if total_pages == 0:
                raise ValueError("PDF vacio")

            self.after(0, lambda t=total_pages: self.append_log(f"PDF: {t} pag"))
            labels = []
            render_scale = DPI / 72

            for page_num in range(total_pages):
                self.after(0, lambda p=(page_num+1)/total_pages: self.progress_bar.set(p))
                page = doc[page_num]
                page_rect = page.rect
                page_text = page.get_text()

                ventas = re.findall(r'Venta:\s*(S\d+)', page_text)
                if not ventas:
                    ventas = [f"etiqueta_p{page_num+1}"]

                for i, venta in enumerate(ventas):
                    if i >= PDF_MAX_LABELS_PER_PAGE:
                        break
                    y_top = PDF_MARGIN_TOP_PTS + (i * PDF_LABEL_SPACING_PTS)
                    y_bottom = y_top + PDF_LABEL_HEIGHT_PTS
                    label_rect = fitz.Rect(PDF_MARGIN_SIDES_PTS, y_top, page_rect.width - PDF_MARGIN_SIDES_PTS, y_bottom) & page_rect
                    pix = page.get_pixmap(matrix=fitz.Matrix(render_scale, render_scale), clip=label_rect)
                    label_img = Image.open(io.BytesIO(pix.tobytes("png")))
                    labels.append({'image': self._fit_to_canvas(label_img, canvas_w, canvas_h), 'venta': venta})

                self.after(0, lambda p=page_num+1, t=total_pages, v=len(ventas):
                    self.append_log(f"Pag {p}/{t}: {v} etiquetas"))

            doc.close()
            doc = None
            self.after(0, lambda: self._set_preview_data(labels))
        except Exception as e:
            self.after(0, lambda err=str(e): self._handle_error(err))
        finally:
            if doc:
                try: doc.close()
                except: pass
            self.after(0, lambda: self.select_button.configure(state="normal"))

    def _set_preview_data(self, labels):
        self.preview_labels = labels
        self.preview_index = 0
        if labels:
            self.append_log(f"Total: {len(labels)} etiquetas")
            self.save_button.configure(state="normal")
            self.print_button.configure(state="normal")
            self.print_all_button.configure(state="normal")
        else:
            self.save_button.configure(state="disabled")
            self.print_button.configure(state="disabled")
            self.print_all_button.configure(state="disabled")
        self._show_preview()
        self.progress_bar.set(1.0)

    # --- Save/Print ---
    def save_and_print(self, send_to_printer=False):
        if self.processing or not self.preview_labels:
            return
        if send_to_printer:
            printer = self.selected_printer.get()
            if not printer or printer == "Sin impresora":
                messagebox.showwarning("Sin impresora", "No hay impresora seleccionada.")
                return
        self.processing = True
        self.save_button.configure(state="disabled")
        self.print_button.configure(state="disabled")
        self.select_button.configure(state="disabled")
        self.progress_bar.set(0)
        thread = threading.Thread(target=self._save_print_thread, args=(send_to_printer,))
        thread.daemon = True
        thread.start()

    def _save_print_thread(self, send_to_printer):
        try:
            fmt_name = self.selected_format.get()
            fmt = LABEL_FORMATS[fmt_name]
            printer = self.selected_printer.get() if send_to_printer else None
            tape_mm = fmt.get("tape_mm", CANVAS_WIDTH_MM)

            doc = fitz.open(self.pdf_path)
            first_text = doc[0].get_text()
            doc.close()

            date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', first_text)
            if date_match:
                parts = date_match.group(1).split('/')
                folder_date = f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
            else:
                folder_date = "sin_fecha"

            exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(exe_dir, "etiquetas", fmt_name.lower(), folder_date)
            os.makedirs(output_dir, exist_ok=True)

            self.after(0, lambda: self.log_message("Guardando..."))
            saved = printed = 0
            used_names = {}
            total = len(self.preview_labels)

            for i, label in enumerate(self.preview_labels):
                self.after(0, lambda p=(i+1)/total: self.progress_bar.set(p))
                venta = label['venta']
                if venta in used_names:
                    used_names[venta] += 1
                    filename = f"{venta}_{used_names[venta]}.png"
                else:
                    used_names[venta] = 1
                    filename = f"{venta}.png"

                filepath = os.path.join(output_dir, filename)
                label['image'].save(filepath, 'PNG', dpi=(DPI, DPI))
                saved += 1

                if send_to_printer and printer:
                    try:
                        result = print_image(filepath, printer, tape_mm)
                        if result and result.returncode == 0:
                            printed += 1
                    except Exception:
                        pass

            self.after(0, lambda: self.progress_bar.set(1.0))
            self.after(0, lambda: self._finish_processing(saved, printed, output_dir))
        except Exception as e:
            self.after(0, lambda err=str(e): self._handle_error(err))

    def print_all_from_folder(self):
        if not self.last_output_dir or not os.path.isdir(self.last_output_dir):
            messagebox.showwarning("Sin carpeta", "Primero debes guardar las etiquetas.")
            return
        printer = self.selected_printer.get()
        if not printer or printer == "Sin impresora":
            messagebox.showwarning("Sin impresora", "No hay impresora seleccionada.")
            return

        files = sorted([f for f in os.listdir(self.last_output_dir) if f.lower().endswith('.png')])
        if not files:
            messagebox.showwarning("Sin archivos", "No hay etiquetas en la carpeta.")
            return

        self.print_all_button.configure(state="disabled")
        self.log_message(f"Imprimiendo {len(files)} etiquetas...")

        fmt = LABEL_FORMATS[self.selected_format.get()]
        tape_mm = fmt.get("tape_mm", CANVAS_WIDTH_MM)

        def _print_thread():
            printed = 0
            for i, fname in enumerate(files):
                filepath = os.path.join(self.last_output_dir, fname)
                try:
                    result = print_image(filepath, printer, tape_mm)
                    if result and result.returncode == 0:
                        printed += 1
                        self.after(0, lambda f=fname: self.append_log(f"  OK: {f}"))
                    elif result:
                        self.after(0, lambda f=fname, e=result.stderr.strip():
                            self.append_log(f"  Error: {f} - {e}"))
                except Exception as e:
                    self.after(0, lambda f=fname, err=str(e):
                        self.append_log(f"  Error: {f} - {err}"))
                self.after(0, lambda p=(i+1)/len(files): self.progress_bar.set(p))

            self.after(0, lambda: self.print_all_button.configure(state="normal"))
            self.after(0, lambda: self.append_log(f"Listo: {printed}/{len(files)} impresas"))
            self.after(0, lambda: messagebox.showinfo("Impresion", f"{printed} etiquetas enviadas"))

        thread = threading.Thread(target=_print_thread)
        thread.daemon = True
        thread.start()

    def _fit_to_canvas(self, img, canvas_w, canvas_h):
        src_w, src_h = img.size
        margin = mm_to_px(1)
        avail_w, avail_h = canvas_w - 2*margin, canvas_h - 2*margin
        scale = min(avail_w/src_w, avail_h/src_h)
        new_w, new_h = int(src_w*scale), int(src_h*scale)
        img_scaled = img.resize((new_w, new_h), Image.LANCZOS)
        canvas = Image.new('RGB', (canvas_w, canvas_h), (255, 255, 255))
        canvas.paste(img_scaled, ((canvas_w-new_w)//2, (canvas_h-new_h)//2))
        return canvas

    def _finish_processing(self, saved, printed, output_dir):
        self.processing = False
        self.save_button.configure(state="normal")
        self.print_button.configure(state="normal")
        self.select_button.configure(state="normal")
        self.last_output_dir = output_dir
        self.print_all_button.configure(state="normal")
        self.append_log(f"Listo: {saved} guardadas" + (f", {printed} impresas" if printed else ""))
        try:
            if os.name == 'nt': os.startfile(output_dir)
            elif sys.platform == 'darwin': subprocess.run(['open', output_dir], check=False)
            else: subprocess.run(['xdg-open', output_dir], check=False)
        except: pass
        msg = f"{saved} etiquetas guardadas"
        if printed: msg += f"\n{printed} enviadas a impresora"
        messagebox.showinfo("Listo", msg + f"\n\n{output_dir}")

    def _handle_error(self, error_message):
        self.processing = False
        self.save_button.configure(state="normal")
        self.print_button.configure(state="normal")
        self.select_button.configure(state="normal")
        self.progress_bar.set(0)
        self.append_log(f"ERROR: {error_message}")
        messagebox.showerror("Error", error_message)

    def log_message(self, msg):
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", msg)
        self.result_text.configure(state="disabled")

    def append_log(self, msg):
        self.result_text.configure(state="normal")
        self.result_text.insert("end", f"\n{msg}")
        self.result_text.see("end")
        self.result_text.configure(state="disabled")


def main():
    app = EtiquetaSeparador()
    app.mainloop()

if __name__ == "__main__":
    main()
