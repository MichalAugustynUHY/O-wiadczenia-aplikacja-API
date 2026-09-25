import os
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import win32com.client as win32
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
from datetime import datetime
import threading
import pythoncom
import numpy as np
import ctypes
from ctypes import wintypes, windll
import io
from reportlab.pdfgen import canvas
from PIL import Image, ImageFilter, ImageEnhance
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
import tempfile
import random
from reportlab.lib.pagesizes import A4

# ------------------------- Get Desktop Path -------------------------
CSIDL_DESKTOP = 0
_SHGetFolderPath = windll.shell32.SHGetFolderPathW
_SHGetFolderPath.argtypes = [wintypes.HWND,
                             ctypes.c_int,
                             wintypes.HANDLE,
                             wintypes.DWORD, wintypes.LPCWSTR]
path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
result = _SHGetFolderPath(0, CSIDL_DESKTOP, 0, 0, path_buf)
desktop_path = path_buf.value
print("Pulpit:", desktop_path)

APP_FOLDER_NAME = "Oświadczenia aplikacja"

def find_app_folder():
    candidates = []

    for env in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        base = os.environ.get(env)
        if base:
            candidates.append(os.path.join(base, APP_FOLDER_NAME))

    # 2) Pulpit (obejmuje też Pulpit przekierowany do OneDrive)
    candidates.append(os.path.join(desktop_path, APP_FOLDER_NAME))

    # 3) Katalog profilu użytkownika
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        candidates.append(os.path.join(userprofile, APP_FOLDER_NAME))

    # Zwróć pierwszą lokalizację, która realnie istnieje
    for path in candidates:
        if os.path.isdir(path):
            print("Znaleziono folder aplikacji:", path)
            return path

    # Nic nie znaleziono – zwróć najlepsze przypuszczenie (pierwsze z listy)
    fallback = candidates[0] if candidates else os.path.join(desktop_path, APP_FOLDER_NAME)
    print("UWAGA: nie znaleziono folderu aplikacji. Używam:", fallback)
    return fallback

base_dir = find_app_folder()

# Excel data file and signatures folder
input_excel_path = os.path.join(base_dir, "Dane sharepoint.xlsx")
signature_dir = os.path.join(base_dir, "podpisy")

# ------------------------- Excel Refresh Function -------------------------
def refresh_excel_file(file_path, progress):
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        # Jasny komunikat, gdy pliku nie ma pod obliczoną ścieżką
        if not os.path.exists(file_path):
            messagebox.showerror(
                "Error",
                "Nie znaleziono pliku danych:\n" + file_path + "\n\n"
                "Sprawdź, czy folder \"" + APP_FOLDER_NAME + "\" (z plikiem "
                "\"Dane sharepoint.xlsx\") znajduje się w OneDrive lub na Pulpicie."
            )
            return

        progress['value'] = 10
        progress.update_idletasks()

        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        progress['value'] = 30
        progress.update_idletasks()

        wb.RefreshAll()  # Refresh all data connections
        excel.CalculateUntilAsyncQueriesDone()
        progress['value'] = 70
        progress.update_idletasks()

        wb.Save()
        progress['value'] = 100
        progress.update_idletasks()
    except Exception as e:
        messagebox.showerror("Error", f"Nie udało się odświeżyć danych: {e}")
    finally:
        if wb is not None:
            wb.Close(False)
        excel.Application.Quit()
        pythoncom.CoUninitialize()

def refresh_data():
    progress_window = tk.Toplevel(root)
    progress_window.title("Odświeżanie danych źródłowych")
    progress_window.geometry("300x100")
    progress_window.resizable(False, False)

    progress_label = ttk.Label(progress_window, text="Odświeżanie danych listy źródłowej, proszę czekać...", anchor="center")
    progress_label.pack(pady=10)

    progress = ttk.Progressbar(progress_window, orient=tk.HORIZONTAL, length=250, mode='determinate')
    progress.pack(pady=10)

    def run_refresh():
        refresh_excel_file(input_excel_path, progress)
        progress_window.destroy()
        load_data()

    threading.Thread(target=run_refresh, daemon=True).start()

def load_data():
    global data, client_names
    try:
        # Now try to load the file
        data = pd.read_excel(input_excel_path)
        client_names = data['Nazwa firmy'].dropna().astype(str).unique().tolist()
        client_name_combobox['values'] = client_names
    except PermissionError:
        messagebox.showerror("Error",
            "Nie można otworzyć pliku - jest on obecnie używany przez inny proces.\n\n" +
            "Proszę:\n" +
            "1. Zamknąć wszystkie okna Excel\n" +
            "2. Sprawdzić czy plik nie jest otwarty w innym programie\n" +
            "3. Spróbować ponownie")
    except Exception as e:
        messagebox.showerror("Error", f"Błąd wczytywania danych: {e}")

# ------------------------- Helper Functions -------------------------
def add_signature(sheet, image_path, cell):
    img = OpenpyxlImage(image_path)
    sheet.add_image(img, cell)

# ------------------------- Scanned Effect Function -------------------------
def add_scanned_effect(img):
    # Slight random rotation
    angle = np.random.uniform(-0.5, 0.4)
    img = img.rotate(angle, expand=1, fillcolor=(255,255,255))
    # Add slight blur
    img = img.filter(ImageFilter.GaussianBlur(radius=0.7))
    # Add noise
    arr = np.array(img)
    noise = np.random.normal(0, 8, arr.shape).astype(np.int16)
    arr = np.clip(arr + noise, 0, 255).astype(np.uint8)
    img = Image.fromarray(arr)
    # Adjust contrast and brightness
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(1.15)
    enhancer = ImageEnhance.Brightness(img)
    img = enhancer.enhance(1.05)
    return img

# ------------------------- Client Record Helpers -------------------------
def normalize_task_type(value):
    return str(value).strip().lower() if pd.notnull(value) else ''

def get_client_records(client_name, audit_type=None, dzien_otw_bil=None):
    records = []
    for _, row in data[data['Nazwa firmy'] == client_name].iterrows():
        if audit_type and "przegląd" in audit_type.lower():
            if row['Data rozpoczęcia'].year != int(dzien_otw_bil[-4:]):
                continue
        records.append({
            'signer_name': row['Osoba odpowiedzialna'],
            'task_type': row['Typ zadania'],
            'appearance_date': row['Data rozpoczęcia'],
            'type_audit': row['Rodzaj sprawozdania'],
        })
    return records

def compute_earliest_date(signer_data_list, client_name):
    dates = [entry['appearance_date'] for entry in signer_data_list
             if normalize_task_type(entry['task_type']) != "oświadczenie"]
    if not dates:
        messagebox.showerror("Error", f"Nie znaleziono odpowiednich danych dla: {client_name} (zadania poza 'Oświadczenie')")
        return None
    return min(dates)

# ------------------------- Signer Selection Dialog -------------------------
def open_signer_selection_dialog(client_name, dzien_otw_bil, dzien_bil, audit_type, data_podpisu_umowy, data_podpisu_badania):
    client_records = data[data['Nazwa firmy'] == client_name].copy()
    if client_records.empty:
        messagebox.showerror("Error", f"Nie znaleziono danych dla klienta: {client_name}")
        return

    signer_data_list = get_client_records(client_name)

    computed_earliest_date = compute_earliest_date(signer_data_list, client_name)
    if computed_earliest_date is None:
        return

    signer_tasks = defaultdict(list)
    for entry in signer_data_list:
        signer_tasks[entry['signer_name']].append({
            'task_type': entry['task_type'],
            'appearance_date': entry['appearance_date']
        })

    valid_signers = []
    for signer, tasks in signer_tasks.items():
        non_oswiadczenie_tasks = [task for task in tasks if normalize_task_type(task['task_type']) != "oświadczenie"]
        if non_oswiadczenie_tasks:
            valid_signers.append(signer)

    sel_dialog = tk.Toplevel(root)
    sel_dialog.title("Wybór daty i podpisujących")
    sel_dialog.geometry("800x800")

    # Table frame
    table_frame = ttk.LabelFrame(sel_dialog, text="Rekordy klienta", padding="10")
    table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Create Treeview with scrollbars
    tree_frame = ttk.Frame(table_frame)
    tree_frame.pack(fill=tk.BOTH, expand=True)
    vsb = ttk.Scrollbar(tree_frame, orient="vertical")
    vsb.pack(side="right", fill="y")
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
    hsb.pack(side="bottom", fill="x")
    columns = ("Data rozpoczęcia", "Osoba odpowiedzialna", "Typ zadania", "Rodzaj sprawozdania")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                        yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)
    tree.pack(fill=tk.BOTH, expand=True)
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    client_records['Data rozpoczęcia'] = pd.to_datetime(client_records['Data rozpoczęcia'], errors='coerce')
    client_records = client_records.sort_values(by='Data rozpoczęcia')

    # Insert records
    for idx, row in client_records.iterrows():
        date_str = row['Data rozpoczęcia'].strftime("%Y-%m-%d") if pd.notnull(row['Data rozpoczęcia']) else ''
        tree.insert("", "end", values=(date_str, row['Osoba odpowiedzialna'], row['Typ zadania'], row['Rodzaj sprawozdania']))

    # Options frame
    options_frame = ttk.LabelFrame(sel_dialog, text="Wybór osób podpisujących i daty", padding="10")
    options_frame.pack(fill=tk.X, padx=10, pady=10)
    ttk.Label(options_frame, text="Osoby podpisujące:").grid(row=0, column=0, sticky=tk.W)
    checkbox_frame = ttk.Frame(options_frame)
    checkbox_frame.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    signer_vars = {}
    for i, signer in enumerate(valid_signers):
        var = tk.IntVar(value=1)
        signer_vars[signer] = var
        cb = ttk.Checkbutton(checkbox_frame, text=signer, variable=var)
        cb.grid(row=i // 3, column=i % 3, sticky=tk.W, padx=5, pady=2)

    ttk.Label(options_frame, text="Data kontaktu (pierwszego zlecenia):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    custom_date_var = tk.StringVar(value=computed_earliest_date.strftime("%d.%m.%Y"))
    if na_dzien_podpisu_var.get() == 1:
        custom_date_entry = ttk.Entry(options_frame, textvariable=custom_date_var, width=15, state='disabled')
    else:
        custom_date_entry = ttk.Entry(options_frame, textvariable=custom_date_var, width=15)
    custom_date_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

    # New frame for client name override
    override_frame = ttk.Frame(options_frame)
    override_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
    ttk.Label(override_frame, text="Zmień nazwę klienta (opcjonalnie):").grid(row=0, column=0, sticky=tk.W)
    client_name_override_var = tk.StringVar(value=client_name)
    ttk.Entry(override_frame, textvariable=client_name_override_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=5)

    def on_tree_select(event):
        item = tree.focus()
        if not item:
            return
        values = tree.item(item, "values")
        date_val = values[0]
        try:
            selected_date = pd.to_datetime(date_val)
            formatted_date = selected_date.strftime("%d.%m.%Y")
        except Exception:
            formatted_date = str(date_val)
            selected_date = None
        custom_date_var.set(formatted_date)
        # Automatically uncheck signers not having any record on or after the selected date
        if selected_date is not None:
            for signer, tasks in signer_tasks.items():
                # If none of the tasks' appearance_date is greater or equal to selected_date, uncheck:
                valid = any(pd.to_datetime(task['appearance_date']) >= selected_date for task in tasks)
                if not valid:
                    signer_vars[signer].set(0)
    tree.bind("<<TreeviewSelect>>", on_tree_select)

    # OK and Cancel buttons
    def on_ok():
        selected_signers = [signer for signer, var in signer_vars.items() if var.get() == 1]
        if not selected_signers:
            messagebox.showerror("Error", "Musisz wybrać co najmniej jedną osobę podpisującą.")
            return
        sel_dialog.destroy()
        # Read the override client name for display only.
        display_client_name = client_name_override_var.get().strip()
        # Pass the original client name (for filtering) and the override (for display)
        process_form(client_name, display_client_name, dzien_otw_bil, dzien_bil, audit_type,
                     data_podpisu_umowy, data_podpisu_badania, selected_signers, custom_date_var.get().strip())
    def on_cancel():
        sel_dialog.destroy()
    btn_frame = ttk.Frame(sel_dialog)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="OK", command=on_ok).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="Anuluj", command=on_cancel).grid(row=0, column=1, padx=10)

# ------------------------- Process Form -------------------------
def process_form(selected_client, display_client, dzien_otw_bil, dzien_bil, audit_type, 
                 data_podpisu_umowy, data_podpisu_badania, selected_signers, custom_date_cell):
    try:
        data_podpisu_umowy = datetime.strptime(data_podpisu_umowy, '%d.%m.%Y').date()
    except ValueError:
        messagebox.showerror("Error", "Nieprawidłowy format daty podpisu umowy. Użyj formatu DD.MM.YYYY.")
        return

    signer_data_list = get_client_records(selected_client, audit_type, dzien_otw_bil)

    earliest_date = compute_earliest_date(signer_data_list, selected_client)
    if earliest_date is None:
        return

    skip_second_signature = na_dzien_podpisu_var.get() == 1
    if skip_second_signature:
        template_file = "Szablon dzien podpisu.xlsx"
        name_cell = 'M6'
        date_cell = 'N9'
        dzien_otw_bil_cell = 'N7'
        dzien_bil_cell = 'N8'
        audit_type_cell = 'N10'
        data_podpisu_umowy_cell = 'N12'
        data_podpisu_badania_cell = 'N13'
    else:
        template_file = "Szablon.xlsx"
        name_cell = 'L2'
        date_cell = 'M5'
        dzien_otw_bil_cell = 'M3'
        dzien_bil_cell = 'M4'
        audit_type_cell = 'M6'
        data_podpisu_umowy_cell = 'M8'

    signer_name_col = 'D'
    signature1_col, signature1_row = 'F', 22
    signature2_col, signature2_row = 'F', 35

    template_path = os.path.join(base_dir, template_file)
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    if custom_date_cell:
        try:
            custom_date = datetime.strptime(custom_date_cell, '%d.%m.%Y').date()
            date_to_use = custom_date
        except ValueError:
            messagebox.showerror("Error", "Nieprawidłowy format daty wpisanej. Użyj formatu DD.MM.YYYY.")
            return
    else:
        date_to_use = earliest_date.date()

    # Write the client name to Excel – use display_client if provided, otherwise the original
    final_client_name = display_client if display_client else selected_client
    if grupa_kapitalowa_var.get():
        ws[name_cell] = "Grupa kapitałowa " + final_client_name
    else:
        ws[name_cell] = final_client_name

    ws[date_cell] = date_to_use
    ws[dzien_otw_bil_cell] = dzien_otw_bil
    ws[dzien_bil_cell] = dzien_bil
    ws[audit_type_cell] = audit_type
    ws[data_podpisu_umowy_cell] = data_podpisu_umowy
    if skip_second_signature:
        ws[data_podpisu_badania_cell] = data_podpisu_badania

    # Get the selected signature color (global variable declared in the GUI setup)
    signature_color = signature_color_var.get()
    
    for i, signer in enumerate(selected_signers):
        available_signatures = []
        for j in range(1, 4):
            candidate_path = os.path.join(signature_dir, f"{signer}_{signature_color}_{j}.png")
            if os.path.exists(candidate_path):
                available_signatures.append(candidate_path)
        if not available_signatures:
            print(f"Podpis dla {signer} nie został znaleziony.")
            continue

        signature_path = random.choice(available_signatures)
        ws[f'{signer_name_col}{signature1_row + i}'] = signer
        add_signature(ws, signature_path, f'{signature1_col}{signature1_row + i}')
        if not skip_second_signature:
            ws[f'{signer_name_col}{signature2_row + i}'] = signer
            add_signature(ws, signature_path, f'{signature2_col}{signature2_row + i}')

    for row in ws.iter_rows(min_row=22, max_row=42, min_col=4, max_col=4):
        if row[0].row == 32:
            continue
        if row[0].value is None:
            ws.row_dimensions[row[0].row].hidden = True

    # Set the custom print area based on the template type.
    if skip_second_signature:
        # For "Szablon dzien podpisu.xlsx"
        ws.print_area = 'A1:J32'
    else:
        # For "Szablon.xlsx"
        ws.print_area = 'A1:I31'

    filled_excel_path = 'oświadczenie.xlsx'
    wb.save(filled_excel_path)

    # ------------------------- PDF Export and Flattening -------------------------
    def excel_to_pdf(excel_path, pdf_path, folder_path):
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb_pdf = None
        try:
            wb_pdf = excel.Workbooks.Open(os.path.abspath(excel_path))
            ws_pdf = wb_pdf.Worksheets(1)
            # Set print area using Excel's COM interface
            if skip_second_signature:
                ws_pdf.PageSetup.PrintArea = "A1:J32"
            else:
                ws_pdf.PageSetup.PrintArea = "A1:I31"

            ws_pdf.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        finally:
            if wb_pdf is not None:
                wb_pdf.Close(False)
            excel.Application.Quit()
            pythoncom.CoUninitialize()

        # Flatten the PDF by converting each page into an image and reassembling.
        temp_flattened_pdf = "temp_flattened.pdf"
        flatten_pdf(pdf_path, temp_flattened_pdf, dpi=300)
        os.replace(temp_flattened_pdf, pdf_path)

    def flatten_pdf(input_pdf_path, output_pdf_path, dpi=300):
        # Convert PDF pages to images.
        images = convert_from_path(input_pdf_path, dpi=dpi)
        a4_width, a4_height = A4  # A4 page dimensions in points (approx. 595x842)
        c = canvas.Canvas(output_pdf_path, pagesize=A4)
        for img in images:
            img = add_scanned_effect(img)
            width, height = img.size  # in pixels
            
            # Convert pixel dimensions to points (72 points per inch).
            img_width_pt = width * 72 / dpi
            img_height_pt = height * 72 / dpi
            
            # Calculate scale to fit image into A4 while preserving aspect ratio.
            scale = min(a4_width / img_width_pt, a4_height / img_height_pt)
            scaled_width = img_width_pt * scale
            scaled_height = img_height_pt * scale
            
            # Center the image on the A4 page.
            x = (a4_width - scaled_width) / 2
            y = (a4_height - scaled_height) / 2
            
            # Save the image temporarily.
            temp_image_path = os.path.join(tempfile.gettempdir(), "temp_page.png")
            img.save(temp_image_path, 'PNG')
            
            # Draw the image on an A4 page.
            c.drawImage(temp_image_path, x, y, width=scaled_width, height=scaled_height)
            c.showPage()
            os.remove(temp_image_path)
        c.save()

    if grupa_kapitalowa_var.get():
        output_client_name = "Grupa kapitałowa " + selected_client
    else:
        output_client_name = selected_client

    if skip_second_signature:
        initial_name = f'Oświadczenie_{output_client_name}_na dzień SzB.pdf'
    else:
        initial_name = f'Oświadczenie_{output_client_name}.pdf'

    # Create the folder only once if it does not exist.
    folder_path = os.path.join(desktop_path, f"Oświadczenia_{selected_client}")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Set the default directory for the save dialog to folder_path.
    output_pdf_path = asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialdir=folder_path,
        initialfile=initial_name
    )
    
    if output_pdf_path:
        # Call excel_to_pdf without creating additional folders.
        try:
            excel_to_pdf(filled_excel_path, output_pdf_path, folder_path)
        except Exception as e:
            messagebox.showerror("Error", f"Nie udało się wyeksportować PDF: {e}")
            return
        finally:
            os.remove(filled_excel_path)
        messagebox.showinfo("Sukces", f"Oświadczenie dla {selected_client} zostało wypełnione i zapisane.")
    else:
        messagebox.showinfo("Anulowano", "Zapis PDF został anulowany.")
        os.remove(filled_excel_path)

# ------------------------- Main Form Functions -------------------------
def submit_form():
    client_name = client_name_var.get()
    year = year_var.get()
    audit_type = audit_type_var.get()
    data_podpisu_umowy = data_podpisu_umowy_var.get()
    data_podpisu_badania = data_podpisu_badania_var.get()
    if not client_name or not year or not audit_type or not data_podpisu_umowy:
        messagebox.showerror("Error", "Proszę wypełnić wszystkie pola.")
        return
    dzien_otw_bil = f"01.01.{year}"
    dzien_bil = f"31.12.{year}"
    if custom_dates_var.get():
        dzien_otw_bil = dzien_otw_bil_var.get()
        dzien_bil = dzien_bil_var.get()
    open_signer_selection_dialog(client_name, dzien_otw_bil, dzien_bil, audit_type, data_podpisu_umowy, data_podpisu_badania)

def toggle_custom_dates():
    state = tk.NORMAL if custom_dates_var.get() else tk.DISABLED
    dzien_otw_bil_box.config(state=state)
    dzien_bil_box.config(state=state)

def toggle_data_podpisu():
    state = tk.NORMAL if na_dzien_podpisu_var.get() else tk.DISABLED
    data_podpisu_badania_box.config(state=state)

def on_client_name_entry(event):
    value = client_name_var.get()
    if len(value) < 3:
        client_name_combobox['values'] = []
    else:
        data_list = [item for item in client_names if value.lower() in item.lower()]
        client_name_combobox['values'] = data_list
        client_name_combobox.event_generate('<Down>')

# ------------------------- Main GUI Setup -------------------------
root = tk.Tk()
root.title("Oświadczenia DAA")
root.minsize(500, 500)

# Apply ttk styles for a modern look
style = ttk.Style(root)
style.theme_use('clam')
style.configure('TLabel', font=('Calibri', 11), padding=5)
style.configure('TButton', font=('Calibri', 11), padding=5)
style.configure('TEntry', font=('Calibri', 11))
style.configure('TCombobox', font=('Calibri', 11))
style.configure('TLabelframe', font=('Calibri', 11, 'bold'), padding=10)
style.configure('TLabelframe.Label', font=('Calibri', 12, 'bold'))

main_frame = ttk.Frame(root, padding="15")
main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=2)

title_label = ttk.Label(main_frame, text="Generator Oświadczeń DAA", font=("Calibri", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))

grupa_kapitalowa_var = tk.IntVar()
grupa_kapitalowa_checkbutton = ttk.Checkbutton(main_frame, text="Grupa kapitałowa", variable=grupa_kapitalowa_var)
grupa_kapitalowa_checkbutton.grid(row=1, column=0, sticky=tk.W)

separator = ttk.Separator(main_frame, orient='horizontal')
separator.grid(row=2, column=0, columnspan=2, sticky='ew', pady=10)

client_frame = ttk.LabelFrame(main_frame, text="Dane oświadczenia")
client_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

ttk.Label(client_frame, text="Wybierz klienta:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
client_name_var = tk.StringVar()
client_name_combobox = ttk.Combobox(client_frame, textvariable=client_name_var)
client_name_combobox['values'] = []
client_name_combobox.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
client_name_combobox.bind('<KeyRelease>', on_client_name_entry)
client_frame.columnconfigure(1, weight=1)

ttk.Label(client_frame, text="Rok badania:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
year_var = tk.StringVar()
year_combobox = ttk.Combobox(client_frame, textvariable=year_var)
year_combobox['values'] = [str(y) for y in range(2021, datetime.now().year + 1)]
year_combobox.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

custom_dates_var = tk.IntVar()
custom_dates_checkbutton = ttk.Checkbutton(
    client_frame, 
    text="Przesunięty rok obrotowy", 
    variable=custom_dates_var, 
    command=toggle_custom_dates
)
custom_dates_checkbutton.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

ttk.Label(client_frame, text="Dzień otwarcia bilansu:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
dzien_otw_bil_var = tk.StringVar(value="01.01.2021")
dzien_otw_bil_box = ttk.Entry(client_frame, textvariable=dzien_otw_bil_var, state=tk.DISABLED)
dzien_otw_bil_box.grid(row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(client_frame, text="Dzień bilansowy:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
dzien_bil_var = tk.StringVar(value="31.12.2021")
dzien_bil_box = ttk.Entry(client_frame, textvariable=dzien_bil_var, state=tk.DISABLED)
dzien_bil_box.grid(row=4, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(client_frame, text="Data podpisu umowy:").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
data_podpisu_umowy_var = tk.StringVar()
data_podpisu_umowy_box = ttk.Entry(client_frame, textvariable=data_podpisu_umowy_var)
data_podpisu_umowy_box.grid(row=5, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(client_frame, text="Badanie/Przegląd:").grid(row=6, column=0, padx=5, pady=5, sticky=tk.W)
audit_type_var = tk.StringVar()
audit_type_combobox = ttk.Combobox(client_frame, textvariable=audit_type_var)
audit_type_combobox['values'] = [
    "badaniem ",
    "badaniem jednostkowego ",
    "badaniem skonsolidowanego ",
    "badaniem jednostkowego i skonsolidowanego ",
    "badaniem śródrocznego skróconego jednostkowego ",
    "przeglądem śródrocznego jednostkowego i skonsolidowanego ",
    "przeglądem śródrocznego ",
    "przeglądem śródrocznego jednostkowego ",
    "przeglądem śródrocznego skonsolidowanego ",
    "przeglądem śródrocznego skróconego ",
    "przeglądem śródrocznego skróconego jednostkowego ",
    "przeglądem śródrocznego skróconego skonsolidowanego "
]
audit_type_combobox.grid(row=6, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(client_frame, text="Kolor podpisu:").grid(row=7, column=0, padx=5, pady=5, sticky=tk.W)
signature_color_var = tk.StringVar(value="czarny")
signature_color_combobox = ttk.Combobox(client_frame, textvariable=signature_color_var)
signature_color_combobox['values'] = ["czarny", "niebieski"]
signature_color_combobox.grid(row=7, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

na_dzien_podpisu_var = tk.IntVar()
na_dzien_podpisu_checkbutton = ttk.Checkbutton(
    client_frame, 
    text="Na dzień podpisu SzB", 
    variable=na_dzien_podpisu_var, 
    command=toggle_data_podpisu
)
na_dzien_podpisu_checkbutton.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

ttk.Label(client_frame, text="Data podpisu SzB:").grid(row=9, column=0, padx=5, pady=5, sticky=tk.W)
data_podpisu_badania_var = tk.StringVar(value=datetime.today().strftime("%d.%m.%Y"))
data_podpisu_badania_box = ttk.Entry(client_frame, textvariable=data_podpisu_badania_var, state=tk.DISABLED)
data_podpisu_badania_box.grid(row=9, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

separator2 = ttk.Separator(main_frame, orient='horizontal')
separator2.grid(row=4, column=0, columnspan=2, sticky='ew', pady=10)

button_frame = ttk.Frame(main_frame)
button_frame.grid(row=5, column=0, columnspan=2, pady=10)
submit_button = ttk.Button(button_frame, text="Dalej", command=submit_form)
submit_button.grid(row=0, column=0, padx=10)
refresh_button = ttk.Button(button_frame, text="Odśwież dane", command=refresh_data)
refresh_button.grid(row=0, column=1, padx=10)
exit_button = ttk.Button(button_frame, text="Wyjdź", command=root.quit)
exit_button.grid(row=0, column=2, padx=10)
button_frame.columnconfigure((0, 1, 2), weight=1)

refresh_data()
root.mainloop()
