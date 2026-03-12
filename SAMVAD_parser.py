import os
import re 
import json
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
import logging
from datetime import datetime
import SAMVAD_mapping
#---------------Logging-----------------------

def setup_logger(log_folder):

    if not os.path.exists(log_folder):
        os.makedirs(log_folder)

    log_path = os.path.join(log_folder, "LOGS.log")

    # Prevent duplicate logs if script runs multiple times
    if logging.getLogger().hasHandlers():
        logging.getLogger().handlers.clear()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, mode="a", encoding="utf-8"),  # Append mode
            logging.StreamHandler()
        ]
    )

    logging.info("========== NEW SAMVAD EXECUTION STARTED ==========")

    return log_path



pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ------------------------required fields that is important to exist----------------------

REQUIRED_FIELDS = [
        "Agency Name",
        "Agency Address",
        "Client Name",
        "Client Address",
        "RO Number",
        "RO FROM",
        "GSTIN",
        "Remark",
        "Newspaper Name",
        "Booking Centre",
        "Package",
        "Publish Date",
        "Not Later Than",
        "Category",
        "Color",
        "Height",
        "Width",
        "AD_CAT",
        "Product",
        "BRAND",
        "Executive",
        "RO Rate",
        "RO Amount"
]

# Extraction using PdfReader

def extract_text_from_pdf(pdf_path):
    text = ""

    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    except:
        pass

# If PdfReader data extraction fails then 'OCR (Optical Character Recognition) using pytesseract' will execute

    if not text.strip():
        print("No text found. Applying OCR...")

        images = convert_from_path(
            pdf_path,
            poppler_path = r"C:\Users\admin\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin"
        )

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            text += ocr_text + "\n"

    return text


def extract_agency_name(text):

    match = re.search(r'for .*?\((.*?)\)', text, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    heading_match = re.search(r'\n([A-Z\s]{15,})\nAND VALUE ADDED', text)
    if heading_match:
        return heading_match.group(1).strip()

    return None


def extract_invoice_data(text):
    data = {}

    agency_match = re.search(r'([A-Z\s\n]{30,})\s+ADVERTISEMENT RELEASE ORDER',text)
    if agency_match:
        agency_block = agency_match.group(1)
        agency_block = agency_block.replace("\n", " ")
        agency_block = re.sub(r"\s+", " ", agency_block).strip()
        data["Agency Name"] = agency_block


    Agency_ad_match = re.search(r'(SAMVAD Office:.*?Ph:\s*[0-9\-\s]+)',text)
    if Agency_ad_match:
        Agency_ad = Agency_ad_match.group(1)
        data["Agency Address"] = Agency_ad
    

    ro_match = re.search(r'From\s*:\s*(.*?)\s*GSTIN',text,re.DOTALL)
    if ro_match:
        RO_FROM = ro_match.group(1)
        RO_FROM = RO_FROM.replace('\n', ' ')
        RO_FROM = re.sub(r',\s*', ', ', RO_FROM)
        RO_FROM = re.sub(r'\s+', ' ', RO_FROM).strip()
        data["RO FROM"] = RO_FROM
    
    
    normalized = re.sub(r'\s+', '', text)
    gstin_match = re.search(r'\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z]\d',normalized)
    if gstin_match:
        gstin = gstin_match.group(0)
        data["GSTIN"] = gstin
    
    
    client_match = re.search(r'Dept\.?\s*to\s*which\s*advt\.?\s*relates\s*(.*?)\s*(?:\n|Office|Managing|Director|Under)',text,re.IGNORECASE | re.DOTALL)
    if client_match:
        data["Client Name"] = client_match.group(1).replace("\n", " ").strip()
            
        
    Client_ad_match = re.search(r'2\.\s*Office/.*?relates\s*(.*?)\s*3\.\s*Ref\.',text,re.DOTALL)
    Client_ad = Client_ad_match.group(1)
    Client_ad = re.sub(r'-\s*\n\s*', '-', Client_ad)
    Client_ad = Client_ad.replace('\n', ' ')
    Client_ad = re.sub(r',\s*', ', ', Client_ad)
    Client_ad = re.sub(r'\s+', ' ', Client_ad).strip()
    data["Client Address"] = Client_ad 
    

    ro_match = re.search(
    r'RO\s*No\.?\s*[:-]?\s*(?:-\s*)?[\r\n\s]*([A-Z0-9][A-Z0-9/\-\n]*?)(?=\s*Dated)',
    text,
    re.IGNORECASE
    )

    if ro_match:
        ro_number = ro_match.group(1).replace("\n", "").strip()
        data["RO Number"] = ro_number

    # Extracting Edition
    edition_matches = re.findall(r'Amar Ujala,\s*([A-Za-z ]+)', text, re.IGNORECASE)

    if edition_matches:

        editions_clean = [e.strip().upper() for e in edition_matches]

    # Package
        data["Package"] = ", ".join(editions_clean)

    # Newspaper Name
        data["Newspaper Name"] = ", ".join([f"Amar Ujala, {e}" for e in editions_clean])

    # Booking Centre
        chandigarh_editions = {"ROHTAK", "KARNAL", "HISAR"}

        if any(e in chandigarh_editions for e in editions_clean):
            data["Booking Centre"] = "CHANDIGARH"
        else:
            data["Booking Centre"] = "NATIONAL"

    else:
        data["Package"] = ""
        data["Newspaper Name"] = ""
        data["Booking Centre"] = ""



    #--------------------------------Agency_code_subcode logic-------------------------------
    
    ro_number = data.get("RO Number", "")
    client_name = data.get("Client Name", "")
    booking_centre = data.get("Booking Centre", "")

    has_C_in_ro = "C" in ro_number.upper()
    is_multi_department = "MULTI DEPARTMENT" in client_name.upper()

    if has_C_in_ro or is_multi_department:

        if booking_centre == "CHANDIGARH":
            data["Agency_code_subcode"] = "SA1254SAM190"
        else:
            data["Agency_code_subcode"] = "SA1463SAM214"

    else:

        if booking_centre == "CHANDIGARH":
            data["Agency_code_subcode"] = "SA329SAM81"
        else:
            data["Agency_code_subcode"] = "SA1462SAM213"
    
    
    #----------Mapping of Package---------------------

    
    package_matches = re.findall(r'Amar Ujala,\s*([A-Za-z ]+)', text, re.IGNORECASE)

    if package_matches:

        packages = []

        for edition in package_matches:
            edition = edition.strip().upper()
            mapped_package = SAMVAD_mapping.PACKAGE_NAME_MAP.get(edition, edition)
            packages.append(mapped_package)

        data["Package"] = ", ".join(packages)

    else:
        data["Package"] = ""
      
 
    # Agency name + Package are combined
    if "Agency Name" in data and package_matches:
        data["Agency Name"] = data["Agency Name"] + ", " + package_matches[0]
    
    
    remark_match = re.search(r'Remarks?\s+([A-Z ]+)', text)
    if remark_match:
        data["Remark"] = remark_match.group(1).strip().upper()


    date_match = re.findall(r'\d{2}-\d{2}-\d{4}', text)
    if len(date_match) >= 2:
        data["Publish Date"] = date_match[0]
        data["Not Later Than"] = date_match[1]

    #-----------------------Default Data------------------------------

    data["Category"] = "DISPLAY"
    data["AD_CAT"] = "G20"
    data["AD_SUBCAT"] = 0
    data["Product"] = "DISPLAY-MISC"
    data["BRAND"] = "None"
    data["Executive"] = "None"
    
    size_match = re.search(r'/\s*([\d.]+)\s*\(Sq\.?\s*cm', text, re.IGNORECASE)
    if size_match:
        data["Height"] = 1
        data["Width"] = float(size_match.group(1))

    
    color_match = re.search(r'/Any Page\s*\n\s*(B&W|Colored)',text,re.IGNORECASE)
    if color_match:
        color = color_match.group(1).strip().upper()

        if "B" in color:
            data["Color"] = "B"
        else:
            data["Color"] = "C"
    else:
        data["Color"] = ""
    

#--------------- Extract ALL RO Rates correctly ---------------------
# ---------------- Extract ALL RO Rates and SUM them ----------------

    rate_matches = re.findall(r'Rs\.?\s*([\d]+\.\d+)\s*\(Per\s*Sq\.?\s*cm\)',text.replace('\n', ' '),re.IGNORECASE)
    rates = [float(rate) for rate in rate_matches]
    data["RO Rate"] = round(sum(rates), 2) if rates else 0.0



    amounts = re.findall(r'Rs\.?\s*([\d,]+\.\d+)', text)
    if amounts:
        final_amount = amounts[-1].replace(",", "")
        data["RO Amount"] = float(final_amount)
    return data

#  Save JSON
def save_json(data, output_path):
    with open(output_path, "w") as f:
        json.dump(data, f, indent=4)

# here data is converted to csv
def append_to_csv(data, csv_path):

    columns = [
        "Agency Name",
        "Agency Address",
        "Agency_code_subcode",
        "Client Name",
        "Client Address",
        "RO Number",
        "RO FROM",
        "GSTIN",
        "Remark",
        "Newspaper Name",
        "Booking Centre",
        "Package",
        "Publish Date",
        "Not Later Than",
        "Category",
        "Color",
        "Height",
        "Width",
        "AD_CAT",
        "AD_SUBCAT",
        "Product",
        "BRAND",
        "Executive",
        "RO Rate",
        "RO Amount"
    ]

    row = {col: data.get(col, "") for col in columns}
    df = pd.DataFrame([row], columns=columns)

    write_header = not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0

    df.to_csv(
        csv_path,
        mode='a',
        header=write_header,
        index=False
    )

def process_pdf(pdf_path, csv_path, json_folder):

    print(f"Processing: {pdf_path}")

    text = extract_text_from_pdf(pdf_path)
    data = extract_invoice_data(text)

    # Create JSON file name
    base_name = os.path.basename(pdf_path).replace(".pdf", ".json")
    json_path = os.path.join(json_folder, base_name)

    # Step 1 — Save extracted data to JSON
    save_json(data, json_path)

    # Step 2 — Read back from JSON (CSV will use this)
    with open(json_path, "r") as f:
        json_data = json.load(f)

    # Step 3 — Append JSON data to CSV
    append_to_csv(json_data, csv_path)
    print("Done.\n")
    
    
    #---------------------------------------------------------------------------
    
    
def generate_csv_from_json(json_path, csv_path):
    with open(json_path, "r", encoding="utf-8") as f:
        data_list = json.load(f)

    columns = [
        "Agency Name",
        "Agency Address",
        "Agency_code_subcode",
        "Client Name",
        "Client Address",
        "RO Number",
        "RO FROM",
        "GSTIN",
        "Remark",
        "Newspaper Name",
        "Booking Centre",
        "Package",
        "Publish Date",
        "Not Later Than",
        "Category",
        "Color",
        "Height",
        "Width",
        "AD_CAT",
        "AD_SUBCAT",
        "Product",
        "BRAND",
        "Executive",
        "RO Rate",
        "RO Amount"
    ]

    df = pd.DataFrame(data_list)
    df = df.reindex(columns=columns)
    df.to_csv(csv_path, index=False)


def process_folder(folder_path, csv_path, json_output_path, error_csv_path):

    all_data = []
    valid_data = []
    error_records = []

    total = 0

    for file in os.listdir(folder_path):
        if file.lower().endswith(".pdf"):

            total += 1
            pdf_path = os.path.join(folder_path, file)

            logging.info(f"Processing file: {file}")

            try:
                text = extract_text_from_pdf(pdf_path)
                data = extract_invoice_data(text)
                data["Source File"] = file

                all_data.append(data)

                # Validate required fields
                missing_fields = [
                    field for field in REQUIRED_FIELDS
                    if not data.get(field) or str(data.get(field)).strip() == ""
                ]

                if missing_fields:
                    logging.warning(
                        f"Missing fields in {file}: {missing_fields}"
                    )

                    error_records.append({
                        "PDF File": file,
                        "Missing Fields": ", ".join(missing_fields)
                    })
                else:
                    valid_data.append(data)
                    logging.info(f"Valid record: {file}")

            except Exception as e:
                logging.error(f"Critical error in {file}: {str(e)}")

                error_records.append({
                    "PDF File": file,
                    "Missing Fields": f"EXTRACTION_ERROR: {str(e)}"
                })

    # Save ALL records to single JSON
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, indent=4)

    logging.info(f"All records saved to JSON: {json_output_path}")

    # Save valid records to CSV
    if valid_data:
        pd.DataFrame(valid_data).to_csv(csv_path, index=False)
        logging.info(f"Valid records saved to CSV: {csv_path}")
    else:
        logging.warning("No valid records found.")

    # Save error CSV
    if error_records:
        pd.DataFrame(error_records).to_csv(error_csv_path, index=False)
        logging.info(f"Errors saved to: {error_csv_path}")
    else:
        logging.info("No errors found.")

    # Summary
    logging.info("========== PROCESS SUMMARY ==========")
    logging.info(f"Total PDFs Processed: {total}")
    logging.info(f"Valid Records: {len(valid_data)}")
    logging.info(f"Invalid Records: {len(error_records)}")
    logging.info("========== SAMVAD Extraction Finished ==========")
    
#-------------------------------Folder Locations----------------------------------------------
if __name__ == "__main__":
    folder_path = r"C:\Users\admin\OneDrive\Desktop\OCR\input\SAMVAD" #-------Location of folder that has SAMVAD RO's-------
    csv_output = r"C:\Users\admin\OneDrive\Desktop\OCR\output.csv"   #-------Location of file that stores output-------
    json_output = r"C:\Users\admin\OneDrive\Desktop\OCR\Workflow\parser\SAMVAD\SAMVAD.json"  #-------Location of json file-------
    error_csv = r"C:\Users\admin\OneDrive\Desktop\OCR\error\SAMVAD_error.csv" #-------Location of file that stores error-------
    log_folder = r"C:\Users\admin\OneDrive\Desktop\OCR\logs"   #-------Location of Logs folder-------

    setup_logger(log_folder)

    process_folder(folder_path, csv_output, json_output, error_csv)


    print("processed successfully.")
