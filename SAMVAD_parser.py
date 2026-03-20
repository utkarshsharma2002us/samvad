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



#-----------------------------------------------Logging-------------------------------------------------------
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

#--------------------------------------------------------------------------------------------------------

# Use of pytesseract for OCR (Optical Character Recognition)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"



# Loading of client code mapping file to extract CLIENT_CODE based on RO_CLIENT_NAME
mapping_df = pd.read_csv(r"C:\Users\admin\OneDrive\Desktop\OCR\client_code_mapped.csv")

# Create dictionary: RO_CLIENT_NAME → MASTER_CLIENT_CODE
client_code_map = dict(
    zip(
        mapping_df["Ro Client Name"].str.strip().str.upper(),
        mapping_df["MASTER_CLIENT_CODE"]
    )
)

client_name_map = dict(
    zip(
        mapping_df["Ro Client Name"].str.strip().str.upper(),
        mapping_df["MASTER_CLIENT_NAME"]
    )
)

# ------------------------required fields that is important to exist----------------------

REQUIRED_FIELDS = [
        "FILE_NAME",
        "AGENCY_NAME",
        "AGENCY_CODE",
        "Agency_code_subcode",
        "CLIENT_CODE",
        "RO_CLIENT_NAME",
        "RO_CLIENT_CODE",
        "RO_NUMBER",
        "RO_DATE",
        #"GSTIN",
        "KEY_NUMBER",
        "CATEGORY",
        "COLOUR",
        "AD_CAT",
        "AD_SUBCAT",
        "PRODUCT",
        "BRAND",
        "PACKAGE_NAME",
        "INSERT_DATE",
        "RO_REMARKS",
        #"Newspaper Name",
        "AD_HEIGHT",
        "AD_WIDTH",
        "AD_SIZE",
        "Executive",
        "PAGE_PREMIUM",
        "POSITIONING",
        "RO_RATE",
        "RO_AMOUNT",
]

CATEGORY_TO_SUBCAT = {
    "Display": "GOVT. DISPLAY",
    "Tender": "GOVT.TENDER",
    "Public Notices": "GOVT. DISPLAY",
    "Auction": "GOVT.TENDER",
    "Recruitment": "GOVT. DISPLAY",
    "Others": "GOVT. DISPLAY",
    "Admission Notice": "GOVT. DISPLAY",
    "Announcements": "GOVT. DISPLAY",
}


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

#-----------------------------Extracyion of Agency Name using multiple regex patterns-------------------------------
def extract_agency_name(text):

    match = re.search(r'for .*?\((.*?)\)', text, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    heading_match = re.search(r'\n([A-Z\s]{15,})\nAND VALUE ADDED', text)
    if heading_match:
        return heading_match.group(1).strip()

    return None

# Extraction of all required fields using regex patterns and data cleaning
def extract_invoice_data(text):
    data = {}

    agency_match = re.search(r'([A-Z\s\n]{30,})\s+ADVERTISEMENT RELEASE ORDER',text)
    if agency_match:
        agency_block = agency_match.group(1)
        agency_block = agency_block.replace("\n", " ")
        agency_block = re.sub(r"\s+", " ", agency_block).strip()
        data["AGENCY_NAME"] = agency_block



#---------------Extracted data not need to be in CSV file is commented-----------------------------------

    # Agency_ad_match = re.search(r'(SAMVAD Office:.*?Ph:\s*[0-9\-\s]+)',text)
    # if Agency_ad_match:
    #     Agency_ad = Agency_ad_match.group(1)
    #     data["Agency Address"] = Agency_ad
    

    # ro_match = re.search(r'From\s*:\s*(.*?)\s*GSTIN',text,re.DOTALL)    #RO FROM field is extracted using regex pattern.
    # if ro_match:
    #     RO_FROM = ro_match.group(1)
    #     RO_FROM = RO_FROM.replace('\n', ' ')
    #     RO_FROM = re.sub(r',\s*', ', ', RO_FROM)
    #     RO_FROM = re.sub(r'\s+', ' ', RO_FROM).strip()
    #     data["RO FROM"] = RO_FROM
    
    
    #normalized = re.sub(r'\s+', '', text)
    # gstin_match = re.search(r'\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z]\d',normalized)
    # if gstin_match:
    #     gstin = gstin_match.group(0)
    #     data["GSTIN"] = gstin
    
            
        
    # Client_ad_match = re.search(r'2\.\s*Office/.*?relates\s*(.*?)\s*3\.\s*Ref\.',text,re.DOTALL)
    # Client_ad = Client_ad_match.group(1)
    # Client_ad = re.sub(r'-\s*\n\s*', '-', Client_ad)
    # Client_ad = Client_ad.replace('\n', ' ')
    # Client_ad = re.sub(r',\s*', ', ', Client_ad)
    # Client_ad = re.sub(r'\s+', ' ', Client_ad).strip()
    # data["Client Address"] = Client_ad 
    
    
    client_match = re.search(r'Dept\.?\s*to\s*which\s*advt\.?\s*relates\s*(.*?)\s*(?:\n|Office|Managing|Director|Under)',
    text,re.IGNORECASE | re.DOTALL)

    if client_match:

        Ro_client_name = client_match.group(1).replace("\n", " ").strip()
        data["RO_CLIENT_NAME"] = Ro_client_name

    # Normalize for matching
        lookup_name = re.sub(r'\s+', ' ', Ro_client_name).strip().upper()

    # Fetch from mapping
        client_code = client_code_map.get(lookup_name, "")
        client_name = client_name_map.get(lookup_name, "")

    # Assign values
        data["CLIENT_CODE"] = client_code
        data["CLIENT_NAME"] = client_name

    # Logging (optional but very useful)
        if not client_code:
            logging.warning(f"CLIENT_CODE not found for: {Ro_client_name}")
        if not client_name:
            logging.warning(f"CLIENT_NAME not found for: {Ro_client_name}")            # Hard value
    

    ro_match = re.search(
    r'RO\s*No\.?\s*[:-]?\s*(?:-\s*)?[\r\n\s]*([A-Z0-9][A-Z0-9/\-\n]*?)(?=\s*Dated)',
    text,
    re.IGNORECASE
    )

    if ro_match:
        ro_number = ro_match.group(1).replace("\n", "").strip()
        data["RO_NUMBER"] = ro_number

    # Extracting Edition
    edition_matches = re.findall(r'Amar Ujala,\s*([A-Za-z ]+)', text, re.IGNORECASE)

    if edition_matches:

        editions_clean = [e.strip().upper() for e in edition_matches]

    # Package
        data["PACKAGE_NAME"] = ", ".join(editions_clean)

    # Newspaper Name
       # data["Newspaper Name"] = ", ".join([f"Amar Ujala, {e}" for e in editions_clean])

    # Booking Centre
        chandigarh_editions = {"ROHTAK", "KARNAL", "HISAR"}

        if any(e in chandigarh_editions for e in editions_clean):
            data["BOOKING_CENTER"] = "CH0"
        else:
            data["BOOKING_CENTER"] = "NA1"

    else:
        data["PACKAGE_NAME"] = ""
        #data["Newspaper Name"] = ""
        data["BOOKING_CENTER"] = ""
    
    subject_match = re.search(r'Subject matter of the advertisement\s*([^\n\r]+)',text,re.IGNORECASE)
    if subject_match:
        data["CATEGORY"] = subject_match.group(1).strip()
        
        
    code_match = re.search(r'Advertisement Code\s*SAMVAD:-\s*([A-Z0-9/\n]+?)(?=\s*RO\b|$)',text,re.IGNORECASE)
    if code_match:
         data["KEY_NUMBER"] = re.sub(r'\s+', '', code_match.group(1))
         
         
         
    position_match = re.search(r'\(Sq\.?\s*cm\)\s*/\s*([A-Za-z ]*?)(?=\s*(?:B&W|Colored))',text,re.IGNORECASE | re.DOTALL)
    if position_match:
        data["POSITIONING"] = position_match.group(1).strip() if position_match.group(1).strip() else ""
    else:
        data["POSITIONING"] = ""
    

    position = data.get("POSITIONING", "").strip().upper()
    if position in ["ANY PAGE", "FRONT PAGE"]:
        data["PAGE_PREMIUM"] = "YES"
    else:
        data["PAGE_PREMIUM"] = "NO"


    #-----------------------------------------Agency_code_subcode logic---------------------------------------------
    
    ro_number = data.get("RO_NUMBER", "")
    Ro_client_name = data.get("RO_CLIENT_NAME", "")
    booking_centre = data.get("BOOKING_CENTER", "")

    has_C_in_ro = "C" in ro_number.upper()
    is_multi_department = "MULTI DEPARTMENT" in Ro_client_name.upper()

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
    
    
    #----------------------------------------Mapping of Package-----------------------------------------------------

    
    package_matches = re.findall(r'Amar Ujala,\s*([A-Za-z ]+)', text, re.IGNORECASE)

    if package_matches:

        packages = []
        for edition in package_matches:
            edition = edition.strip().upper()
            mapped_package = SAMVAD_mapping.PACKAGE_NAME_MAP.get(edition, edition)
            packages.append(mapped_package)

        data["PACKAGE_NAME"] = ", ".join(packages)

    else:
        data["PACKAGE_NAME"] = ""
      
 
    # Agency name + Package are combined
    if "AGENCY_NAME" in data and package_matches:
        data["AGENCY_NAME"] = data["AGENCY_NAME"] + ", " + package_matches[0]
    
    
    remark_match = re.search( r'Remarks?\s*(.*?)\s*B\.\s*Advertisement',text,re.IGNORECASE | re.DOTALL)

    if remark_match:
        remark = remark_match.group(1)
        remark = remark.replace('\n', ' ')
        remark = re.sub(r'\s+', ' ', remark).strip()
        data["RO_REMARKS"] = remark.upper()
    else:
        data["RO_REMARKS"] = ""
        
    # date_match = re.findall(r'\d{2}-\d{2}-\d{4}', text)
    # if len(date_match) >= 2:
    #     data["INSERT_DATE"] = date_match[0]
        # data["Not Later Than"] = date_match[1]
        
    pub_date_match = re.search(r"Publication\s*Date.*?(\d{2}-\d{2}-\d{4})",text, re.IGNORECASE | re.DOTALL)
    if pub_date_match:
        data["INSERT_DATE"] = pub_date_match.group(1)
    else:
        data["INSERT_DATE"] = ""
        
        
        
    RO_DATE_match = re.search(r"Dated(.*?)From\s*:", text, re.DOTALL)

    if RO_DATE_match:
        section = RO_DATE_match.group(1)

    # Extract only the date
        date_match = re.search(r"\d{2}/\d{2}/\d{4}", section)

        if date_match:
            data["RO_DATE"] = date_match.group(0)
        else:
            data["RO_DATE"] = ""
    else:
        data["RO_DATE"] = ""


#--------------------------------------------Hardcoded values------------------------------------------------------

    data["AD_CAT"] = "GO2"
    data["PRODUCT"] = "DISPLAY-MISC"
    data["BRAND"] = "None"
    data["Executive"] = "None"
    data["AGENCY_CODE"] = "None"
    data["RO_CLIENT_CODE"] = "None"
#-------------------------------------------------------------------------------------------------- ----------------

    category = data.get("CATEGORY", "")
    mapped_subcat = ""

    for key in CATEGORY_TO_SUBCAT:
        if key in category:
            mapped_subcat = CATEGORY_TO_SUBCAT[key]
            break
    data["AD_SUBCAT"] = mapped_subcat




    size_match = re.search(r'/\s*([\d.]+)\s*\(Sq\.?\s*cm', text, re.IGNORECASE)
    if size_match:
        data["AD_HEIGHT"] = 1
        data["AD_WIDTH"] = float(size_match.group(1))
        data["AD_SIZE"] = data["AD_HEIGHT"] * data["AD_WIDTH"]

    
    color_match = re.search( r'(?:/(Any Page|Front Page))?\s*\n\s*(B&W|Colored)',text,re.IGNORECASE | re.DOTALL)
    if color_match:
        color = color_match.group(2).strip().upper()

        if "B" in color:
            data["COLOUR"] = "B"
        else:
            data["COLOUR"] = "C"
    else:
        data["COLOUR"] = ""
    

#-----if multiple RO rates are found, they are all extracted and summed up to get the total RO Rate-----


    rate_matches = re.findall(r'Rs\.?\s*([\d]+\.\d+)\s*\(Per\s*Sq\.?\s*cm\)',text.replace('\n', ' '),re.IGNORECASE)
    rates = [float(rate) for rate in rate_matches]
    data["RO_RATE"] = round(sum(rates), 2) if rates else 0.0



    amounts = re.findall(r'Rs\.?\s*([\d,]+\.\d+)', text)
    if amounts:
        final_amount = amounts[-1].replace(",", "")
        data["RO_AMOUNT"] = float(final_amount)
    return data

#  Save JSON
def save_json(data, output_path):
    with open(output_path, "w") as f:
        json.dump(data, f, indent=4)

# here data is converted to csv
def append_to_csv(data, csv_path):
    
    
#------Name of the columns in CSV file is defined in the same sequence as required in the output CSV-----------
    
    
    #     DB_COLUMNS = [
#     "FILE_NAME", 
#     "AGENCY_CODE",
#     "AGENCY_NAME",
#     "CLIENT_CODE",
#     "CLIENT_NAME",
#     "RO_CLIENT_CODE",
#     "RO_CLIENT_NAME", 
#     "RO_NUMBER",
#     "RO_DATE",
#     "KEY_NUMBER",
#     "COLOUR",
#     "AD_CAT",
#     "AD_SUBCAT",
#     "PRODUCT",
#     "BRAND",
#     "PACKAGE_NAME",
#     "INSERT_DATE",
#     "AD_HEIGHT",
#     "AD_WIDTH",
#     "AD_SIZE",
#     "PAGE_PREMIUM",
#     "RO_AMOUNT",
#     "RO_RATE",
#     "RO_REMARKS",
#     "EXTRACTED_TEXT",
#     "POSITIONING"
# ]
    
    columns = [
        "FILE_NAME",
        "AGENCY_NAME",
        "AGENCY_CODE",
        "Agency_code_subcode",
        "CLIENT_CODE",
        "CLIENT_NAME",
        "RO_CLIENT_CODE",
        "RO_NUMBER",
        "RO_DATE",
        #"GSTIN",
        "KEY_NUMBER",
        "CATEGORY",
        "COLOUR",
        "AD_CAT",
        "AD_SUBCAT",
        "PRODUCT",
        "BRAND",
        "PACKAGE_NAME",
        "INSERT_DATE",
        "RO_REMARKS",
        #"Newspaper Name",
        "AD_HEIGHT",
        "AD_WIDTH",
        "AD_SIZE",
        "Executive",
        "PAGE_PREMIUM",
        "POSITIONING",
        "RO_RATE",
        "RO_AMOUNT",
    ]
    if "CLIENT_CODE" not in data:
        data["CLIENT_CODE"] = ""
        
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

    # Create JSON FILE_NAME
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
    
    
#--------------------------------------------Generate CSV from JSON file--------------------------------------------
def generate_csv_from_json(json_path, csv_path):
    with open(json_path, "r", encoding="utf-8") as f:
        data_list = json.load(f)

    
    columns = [
        "FILE_NAME",
        "AGENCY_NAME",
        "AGENCY_CODE",
        "Agency_code_subcode",
        "CLIENT_CODE",
        "CLIENT_NAME",
        "RO_CLIENT_CODE",
        "RO_NUMBER",
        "RO_DATE"
        #"GSTIN",
        "KEY_NUMBER",
        "CATEGORY",
        "COLOUR",
        "AD_CAT",
        "AD_SUBCAT",
        "PRODUCT",
        "BRAND",
        "PACKAGE_NAME",
        "INSERT_DATE",
        "RO_REMARKS",
        #"Newspaper Name",
        "AD_HEIGHT",
        "AD_WIDTH",
        "AD_SIZE",
        "Executive",
        "PAGE_PREMIUM",
        "POSITIONING"
        "RO_RATE",
        "RO_AMOUNT",
    ]
    
    df = pd.DataFrame(data_list)
    df = df.reindex(columns=columns)
    df.to_csv(csv_path, index=False)


#------------------------------------Main processing function for folder of PDFs------------------------------------
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
                data["FILE_NAME"] = file

                all_data.append(data)

                # Validate required fields
                OPTIONAL_FIELDS = ["POSITIONING","CLIENT_CODE"]
                missing_fields = [
                    field for field in REQUIRED_FIELDS
                   if field not in OPTIONAL_FIELDS and (not data.get(field) or str(data.get(field)).strip() == "")
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
    
#------------------------------------------------------------------------------------------------------------------
#--Path of the folder containing SAMVAD PDFs, output CSV file path, output JSON file path and error CSV file path--

if __name__ == "__main__":
    folder_path = r"C:\Users\admin\OneDrive\Desktop\OCR\input\SAMVAD" #folder path containing SAMVAD's PDFs
    csv_output = r"C:\Users\admin\OneDrive\Desktop\OCR\output.csv" #output CSV file path
    json_output = r"C:\Users\admin\OneDrive\Desktop\OCR\Workflow\parser\SAMVAD\SAMVAD.json" #output JSON file path
    error_csv = r"C:\Users\admin\OneDrive\Desktop\OCR\error\SAMVAD_error.csv" #error CSV file path
    log_folder = r"C:\Users\admin\OneDrive\Desktop\OCR\logs"#log folder path
    #folder_path = r"C:\Users\admin\OneDrive\Desktop\test"
    
    setup_logger(log_folder)
    process_folder(folder_path, csv_output, json_output, error_csv)
    print("processed successfully.")