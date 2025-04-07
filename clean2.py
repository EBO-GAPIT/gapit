import os
import re
import fitz
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

class Config:
    def __init__(self):
        self.pdf_folder = ""
        self.bms_file = ""
        self.output_folder = ""
        self.cancel = False

config = Config()

def parse_pdfs(pdf_folder, update_status, update_progress, search_subfolders=False):
    results = []
    skipped_files = []
    pdf_files = []

    # Collect PDF files
    if search_subfolders:
        for root, _, files in os.walk(pdf_folder):
            pdf_files.extend([os.path.join(root, f) for f in files if f.lower().endswith('.pdf')])
    else:
        pdf_files = [os.path.join(pdf_folder, f) for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]

    total = len(pdf_files)

    def extract_tfm_code(filename):
        base = os.path.basename(filename).replace(".pdf", "").strip()
        cleaned = re.split(r"\s-\s", base)[0]
        return cleaned.strip().upper()

    for i, file_path in enumerate(pdf_files, 1):
        if config.cancel:
            update_status("Operation canceled.")
            return []
        update_status(f"Parsing: {os.path.basename(file_path)}")
        try:
            with fitz.open(file_path) as doc:
                text = "\n".join([p.get_text() for p in doc])
        except Exception as e:
            skipped_files.append((file_path, str(e)))
            continue

        pod = re.search(r"Pod serial number:\s*(\d+)", text)
        tested_by = re.search(r"Tested by:\s*(.*)", text)
        test_date = re.search(r"Test completed:\s*(.*)", text)
        modbus = re.search(r"Successful Modbus queries:\s*(.*)", text)
        traceroute_ips = re.findall(r"IP address:\s*(10(?:\.\d+){3})", text)
        ip = traceroute_ips[-1] if traceroute_ips else None

        tfm_code = extract_tfm_code(file_path)

        results.append({
            "Filename": os.path.basename(file_path),
            "Folder": os.path.dirname(file_path),
            "PodSerialNumber": pod.group(1) if pod else "",
            "TestedBy": tested_by.group(1).strip() if tested_by else "",
            "TestCompleted": test_date.group(1).strip() if test_date else "",
            "TestedIP": ip,
            "DeviceTag": tfm_code,
            "ModbusQueries": modbus.group(1) if modbus else ""
        })

        if total > 0:
            progress = int((i / total) * 100)
            update_progress(progress)

    return results

def compare_with_bms(parsed_data, bms_path, update_status):
    update_status("Reading BMS Integration List...")
    bms_excel = pd.ExcelFile(bms_path)
    comparison = []

    site_mode = "Integration List Customer" in bms_excel.sheet_names

    if site_mode:
        df_site = pd.read_excel(bms_path, sheet_name="Integration List Customer", header=3).fillna("")
    else:
        sheets = {sheet: pd.read_excel(bms_path, sheet_name=sheet, header=3).fillna("") for sheet in bms_excel.sheet_names}

    for entry in parsed_data:
        tag = entry['DeviceTag'].strip().upper()
        tested_ip = entry['TestedIP']
        pod = entry['PodSerialNumber']
        filename = entry['Filename']

        expected_ip = ""
        site = ""
        building = ""
        description = ""
        match_status = "❓ Tag not found"

        if site_mode:
            norm_tag = re.split(r"[\\s\\-]", tag)[0].strip().upper()
            for _, row in df_site.iterrows():
                dev_tag = str(row.get("TFM Code", "")).strip().upper()
                ip = str(row.get("IP-Address", "")).strip()
                if norm_tag == dev_tag:
                    expected_ip = ip
                    description = row.get("Description", "")
                    site = row.get("Site", "")
                    building = row.get("Building", "")
                    match_status = "✅ Match" if ip == tested_ip else "❌ Mismatch"
                    break
        else:
            if pod in sheets:
                df = sheets[pod]
                for _, row in df.iterrows():
                    dev_tag = str(row.get("Device Tag", "")).strip().upper().replace("-", "")
                    ip = str(row.get("IP-Address", "")).strip()
                    if tag.replace("-", "") == dev_tag:
                        expected_ip = ip
                        description = row.get("Description", "")
                        site = row.get("Site", "")
                        building = row.get("Building", "")
                        match_status = "✅ Match" if ip == tested_ip else "❌ Mismatch"
                        break

        comparison.append({
            **entry,
            "ExpectedIP": expected_ip,
            "Site": site,
            "Building": building,
            "Description": description,
            "Match": match_status
        })

    return comparison

def run_all():
    if not config.pdf_folder or not config.bms_file or not config.output_folder:
        messagebox.showwarning("Missing", "Please select PDF folder, BMS file, and output folder.")
        return

    progress_var.set(0)
    update_status("Parsing Ginspector Reports...")

    parsed_data = parse_pdfs(config.pdf_folder, update_status, progress_var.set, search_subfolders=search_subfolders_var.get())
    if not parsed_data:
        messagebox.showwarning("No PDFs", "No reports were parsed.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    parsed_output_path = os.path.join(config.output_folder, f"ginspector_parsed_output_{timestamp}.xlsx")
    pd.DataFrame(parsed_data).to_excel(parsed_output_path, index=False)

    update_status("Comparing with BMS list...")
    compared = compare_with_bms(parsed_data, config.bms_file, update_status)

    final_output_path = os.path.join(config.output_folder, f"ginspector_bms_verified_result_{timestamp}.xlsx")
    pd.DataFrame(compared).to_excel(final_output_path, index=False)

    update_status("✅ Complete!")
    match_stats = pd.DataFrame(compared)["Match"].value_counts().to_dict()
    stats_text = f"✅ Matches: {match_stats.get('✅ Match', 0)}\n❌ Mismatches: {match_stats.get('❌ Mismatch', 0)}\n❓ Not found: {match_stats.get('❓ Tag not found', 0)}"
    messagebox.showinfo("Done", f"Parsed:\n{parsed_output_path}\n\nCompared:\n{final_output_path}\n\n" + stats_text)

def select_pdf_folder():
    folder = filedialog.askdirectory(title="Select Folder with Ginspector PDF Reports")
    if folder:
        config.pdf_folder = folder
        gin_var.set(folder)

def select_bms():
    file = filedialog.askopenfilename(title="Select BMS Integration List Excel", filetypes=[("Excel Files", "*.xlsx")])
    if file:
        config.bms_file = file
        bms_var.set(file)

def select_output_folder():
    folder = filedialog.askdirectory(title="Select Folder to Save Results")
    if folder:
        config.output_folder = folder
        out_var.set(folder)

def create_labeled_button(frame, text, command, var, row):
    tk.Button(frame, text=text, command=command).grid(row=row, column=0, padx=10, pady=5)
    tk.Label(frame, textvariable=var, width=60, anchor="w").grid(row=row, column=1)

def download_pdfs_from_sharepoint(site_url, folder_url, local_folder, username, password):
    """
    Downloads all PDF files from a SharePoint folder to a local folder.
    """
    try:
        # Authenticate with SharePoint
        ctx_auth = AuthenticationContext(site_url)
        if not ctx_auth.acquire_token_for_user(username, password):
            raise Exception("Authentication failed. Check your credentials.")

        ctx = ClientContext(site_url, ctx_auth)
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()

        # Download each PDF file
        for file in files:
            if file.name.lower().endswith(".pdf"):
                local_path = os.path.join(local_folder, file.name)
                with open(local_path, "wb") as local_file:
                    file_data = File.open_binary(ctx, file.serverRelativeUrl)
                    local_file.write(file_data.content)
                print(f"Downloaded: {file.name}")

        messagebox.showinfo("Success", f"PDFs downloaded to {local_folder}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to download PDFs: {str(e)}")

def select_sharepoint_folder():
    site_url = "https://yourcompany.sharepoint.com/sites/YourSite"
    folder_url = "/sites/YourSite/Shared Documents/YourFolder"
    local_folder = filedialog.askdirectory(title="Select Local Folder to Save PDFs")
    if local_folder:
        username = "your_email@yourcompany.com"
        password = "your_password"
        download_pdfs_from_sharepoint(site_url, folder_url, local_folder, username, password)
        config.pdf_folder = local_folder
        gin_var.set(local_folder)

# GUI
root = tk.Tk()
root.title("Ginspector IP Parser & BMS Comparator")
root.geometry("700x420")

tk.Label(root, text="Automatically Parse Ginspector PDFs and Compare to BMS", font=("Helvetica", 13, "bold")).pack(pady=10)

frame = tk.Frame(root)
frame.pack()

gin_var = tk.StringVar(value="No folder selected")
bms_var = tk.StringVar(value="No file selected")
out_var = tk.StringVar(value="No folder selected")

# Add buttons for selecting folders and files
create_labeled_button(frame, "Select PDF Folder", select_pdf_folder, gin_var, 0)

# Add a checkbox for searching subfolders in the same row as "Select PDF Folder"
search_subfolders_var = tk.BooleanVar(value=False)
tk.Checkbutton(frame, text="Search Subfolders", variable=search_subfolders_var).grid(row=0, column=2, padx=10, pady=5, sticky="w")

create_labeled_button(frame, "Select BMS Excel", select_bms, bms_var, 1)
create_labeled_button(frame, "Select Output Folder", select_output_folder, out_var, 2)

# Add the Parse & Compare button
tk.Button(root, text="🚀 Parse & Compare", command=run_all, width=30, height=2).pack(pady=10)

# Add a progress bar
progress_var = tk.IntVar()
ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate", variable=progress_var).pack(pady=5)

# Add a status label
status_label = tk.Label(root, text="Waiting...")
status_label.pack()

def update_status(msg):
    status_label.config(text=msg)
    root.update_idletasks()

# Add a footer label
tk.Label(root, text="© Gapit Nordics", font=("Arial", 9, "italic")).pack(side="bottom", pady=5)

root.mainloop()
