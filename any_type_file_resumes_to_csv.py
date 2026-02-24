import os
import re
import csv
import logging
import subprocess
import pdfplumber
from datetime import datetime
from tkinter import Tk, filedialog, messagebox
from tqdm import tqdm
import aspose.words as aw

# --- Extraction Utilities ---
def find_email(text):
    if not text: return "N/A"
    clean_text = " ".join(text.split())
    pattern = r'[a-zA-Z0-9._%+-]+\s*@\s*[a-zA-Z0-9.-]+\s*\.\s*[a-zA-Z]{2,}'
    match = re.search(pattern, clean_text)
    return match.group().strip() if match else "N/A"

def find_phone(text):
    if not text: return "Not Found"
    pattern = r'\b\+?[\d\s\-\(\).]{10,20}\b'
    matches = re.finditer(pattern, text)
    for match in matches:
        candidate = match.group().strip()
        clean = ''.join(filter(str.isdigit, candidate))
        if 10 <= len(clean) <= 15: return clean
    return "Not Found"

def get_pdf_text_robust(path):
    try:
        with pdfplumber.open(path) as pdf:
            full_text = "".join([page.extract_text(layout=True) or "" for page in pdf.pages])
            return {"status": "Success", "text": full_text}
    except Exception:
        return {"status": "Error", "text": ""}

# --- Main Engine ---
def main():
    logging.basicConfig(filename='tpc_master_log.txt', level=logging.INFO, format='%(asctime)s - %(message)s')
    root = Tk()
    root.withdraw()

    input_folder = filedialog.askdirectory(title="TPC ATS - Select Resume Folder")
    if not input_folder: return

    output_dir = os.path.join(input_folder, "Converted_PDFs")
    if not os.path.exists(output_dir): os.makedirs(output_dir)

    valid_exts = (".docx", ".doc", ".rtf", ".odt", ".txt")
    files_to_process = [f for f in os.listdir(input_folder) if f.lower().endswith(valid_exts)]

    all_results = []
    missing_data_details = []
    stats = {"total": len(files_to_process), "emails": 0, "phones": 0, "errors": 0}

    for filename in tqdm(files_to_process, desc="Processing"):
        source_path = os.path.join(input_folder, filename)
        pdf_path = os.path.join(output_dir, os.path.splitext(filename)[0] + ".pdf")
        try:
            doc = aw.Document(source_path)
            doc.save(pdf_path)
            result = get_pdf_text_robust(pdf_path)
            email, phone = find_email(result["text"]), find_phone(result["text"])
            if email != "N/A": stats["emails"] += 1
            if phone != "Not Found": stats["phones"] += 1
            all_results.append({"File": filename, "Email": email, "Phone": phone})
            if email == "N/A" or phone == "Not Found":
                missing_data_details.append(f"{filename}: {'Email Missing ' if email == 'N/A' else ''}{'Phone Missing' if phone == 'Not Found' else ''}")
        except Exception as e:
            stats["errors"] += 1
            logging.error(f"Error on {filename}: {str(e)}")

    current_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    csv_path = os.path.join(output_dir, "TPC_Results.csv")
    
    with open(csv_path, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=["File", "Email", "Phone", "Processed_At"])
        writer.writeheader()
        for row in all_results:
            row["Processed_At"] = current_now
            writer.writerow(row)

    with open("missing_data_report.txt", "w") as txtfile:
        txtfile.write(f"TPC QUALITY REPORT - {current_now}\n" + "-"*40 + "\n")
        for entry in missing_data_details: txtfile.write(f"- {entry}\n")

    logging.info(f"SUCCESS: Batch Complete. CSV saved to {csv_path}")
    messagebox.showinfo("TPC Complete", f"Success! 📧 {stats['emails']} | 📱 {stats['phones']}\nCSV: {csv_path}")
    subprocess.Popen(f'explorer "{os.path.normpath(output_dir)}"')

if __name__ == "__main__":
    main()