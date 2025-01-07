import pikepdf
import pdfplumber
import re
import os
import shutil
import tkinter as tk
from tkinter import ttk

def empty_folder(folder_path):
    """Empty the folder by deleting all its contents."""
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            shutil.rmtree(item_path)  # Remove subdirectories and their contents
        else:
            os.remove(item_path)  # Remove files

def decrypt_pdf(input_path, decrypted_folder):
    """Decrypts a PDF and saves it to the decrypted_folder."""
    try:
        with pikepdf.open(input_path) as pdf:
            decrypted_path = os.path.join(decrypted_folder, os.path.basename(input_path))
            pdf.save(decrypted_path)  # Save decrypted PDF to the decrypted folder
            # print(f"Decrypted and saved: {decrypted_path}")
            return decrypted_path
    except pikepdf.PasswordError:
        print(f"Error: Could not open {input_path} - incorrect password or strong encryption.")
        return None
    except Exception as e:
        print(f"Error decrypting {input_path}: {e}")
        return None

def extract_information(decrypted_path, output_folder, error_folder):
    """Extracts information from a decrypted PDF and saves it with a new filename in output_folder."""
    codigo, punto_venta, nro_factura, cuit = None, None, None, None
    translate = {"6": "3",
                 "11": "5",
                 "15": "6"}
    try:
        with pdfplumber.open(decrypted_path) as pdf:
            text = "".join(page.extract_text() or "" for page in pdf.pages)
            
            # Regex to find Codigo
            codigo_match = re.search(r'COD\.?\s*0*(\d+)', text, re.IGNORECASE)
            if codigo_match:
                codigo = codigo_match.group(1)
            else:
                codigo_match = re.search(r'Codigo\s*nº\s*(\d+)', text, re.IGNORECASE)
                codigo = codigo_match.group(1) if codigo_match else None
            # translate codigo 
            codigo = translate[codigo]


            cuit_match = re.search(r'CUIT\:?\s*(\d{11})', text)
            if cuit_match:
                cuit = cuit_match.group(1)
                # cuit1 = cuit_match1.group(1) if cuit_match1 and len(cuit_match1.group(1))==11 else None
            else:
                cuit_match = re.search(r':\s*(\d{2})-(\d{8})-(\d)', text)
                cuit = cuit_match.group(1) + cuit_match.group(2) + cuit_match.group(3) if cuit_match else None

            # if len(cuit) != 11:
            #     cuit = None

            punto_venta_match = re.search(r'Punto de Venta:\s*(\d+)', text)
            if punto_venta_match:
                nro_factura_match = re.search(r'Comp\. Nro:\s*(\d+)', text)
                punto_venta = punto_venta_match.group(1).lstrip('0') if punto_venta_match else None
                nro_factura = nro_factura_match.group(1).lstrip('0') if nro_factura_match else None
            else:
                punto_venta_factura_match = re.search(r'\D*0*(\d{5})\s*-\s*0*(\d{8})', text)
                if punto_venta_factura_match:
                    punto_venta = punto_venta_factura_match.group(1).lstrip('0')
                    nro_factura = punto_venta_factura_match.group(2).lstrip('0')
                else:
                    punto_venta = None
                    nro_factura = None

        # Create a filename if all information is available
        if codigo and cuit and punto_venta and nro_factura:
            output_filename = f"{cuit}_{codigo}_{punto_venta}_{nro_factura}.pdf"
            output_path = os.path.join(output_folder, output_filename)
            os.rename(decrypted_path, output_path)
        else:
            output_filename = os.path.basename(decrypted_path)
            output_path = os.path.join(error_folder, output_filename)
            os.rename(decrypted_path, output_path)
            print(f"Error: {decrypted_path} CUIT: {cuit} Cod: {codigo} Punto de Venta: {punto_venta} Nro Factura: {nro_factura}")
            return None

    except Exception as e:
        print(f"Error extracting information: {e}")
        return None

    return output_path

def process_all_pdfs(input_folder, decrypted_folder, output_folder, error_folder, progress_bar, total_files, processed_files):

    # Loop through each file in the input directory
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            input_path = os.path.join(input_folder, filename)
            # print(f"Processing: {input_path}")
            
            # Step 1: Decrypt the PDF and save to decrypted folder
            decrypted_path = decrypt_pdf(input_path, decrypted_folder)
            if decrypted_path:
                # Step 2: Extract information and save to outpcd ut folder
                extract_information(decrypted_path, output_folder, error_folder)
            processed_files += 1
            progress_bar['value'] = (processed_files / total_files) * 100
            progress_bar.update_idletasks()
    return processed_files

# GUI con tkinter
def start_processing():
    #  * * * * * * * * * * * * * * * *
    #
    #
    #root_path = "G:\Mi unidad\Capacitacion\InSoft\AOMAOSAM\SignReady" # NEED  TO BE CHANGED DEPENDING ON ENVIRONMENT
    root_path = "s:\Contaduria\InSoft\SignReady"

    input_folder = os.path.join(root_path, "para procesar")
    input_subfolders = [item for item in os.listdir(input_folder) if os.path.isdir(os.path.join(input_folder, item))] # fill input_subfolders list
    decrypted_folder = os.path.join(root_path, "desencriptados")
    output_folder = os.path.join(root_path, "procesados")
    error_folder = os.path.join(root_path, "errores")

    # Ensure all folders exist
    os.makedirs(decrypted_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(error_folder, exist_ok=True)
    os.makedirs(input_folder, exist_ok=True)

    # Empty the folders
    empty_folder(decrypted_folder)
    empty_folder(output_folder)
    empty_folder(error_folder)

    total_files = sum(1 for f in os.listdir(input_folder) if f.endswith(".pdf"))
    processed_files = 0
    for item in input_subfolders:
        os.makedirs(os.path.join(output_folder, item), exist_ok=True)
        total_files += sum(1 for f in os.listdir(os.path.join(input_folder, item)) if f.endswith(".pdf"))

    processed_files = process_all_pdfs(input_folder, decrypted_folder, output_folder, error_folder, progress_bar, total_files, processed_files)

    for item in input_subfolders:
        processed_files = process_all_pdfs(os.path.join(input_folder, item), decrypted_folder, os.path.join(output_folder, item), error_folder, progress_bar, total_files, processed_files)

    # Cambiar el texto del botón al finalizar
    start_button.config(text="Finalizado", command=lambda: app.quit())


# Crear ventana principal
app = tk.Tk()
app.title("SignReady")
app.geometry("400x200")

# Elementos de la interfaz
tk.Label(app, text="SignReady", font=("Helvetica", 16)).pack(pady=10)
tk.Label(app, text="Desarrollado por InSoft").pack()

progress_bar = ttk.Progressbar(app, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=20)

start_button = tk.Button(app, text="Comenzar", command=start_processing)
start_button.pack(pady=10)

app.mainloop()