#----------------------------------Approach on the basis of PDF name-------------------------------------------


# import os
# import shutil

# SOURCE_FOLDER = r"C:\Users\admin\OneDrive\Desktop\OCR\RO input"
# DEST_FOLDER = r"C:\Users\admin\OneDrive\Desktop\OCR\input"

# samvad_folder = os.path.join(DEST_FOLDER, "SAMVAD")
# davp_folder = os.path.join(DEST_FOLDER, "DAVP")
# others_folder = os.path.join(DEST_FOLDER, "Others")

# def classify_by_filename(filename):
#     name = filename.upper()

#     if "CENTRAL BUREAU OF COMMUNICATION" in name or "CBC" in name:
#         return "DAVP"

#     elif "-" in name and "SIZE" in name:
#         return "SAMVAD"

#     else:
#         return "Others"


# #------PROCESS FILES----------

# for filename in os.listdir(SOURCE_FOLDER):
#     if filename.lower().endswith(".pdf"):
#         file_path = os.path.join(SOURCE_FOLDER, filename)
#         category = classify_by_filename(filename)

#         if category == "SAMVAD":
#             shutil.move(file_path, os.path.join(samvad_folder, filename))
#         elif category == "DAVP":
#             shutil.move(file_path, os.path.join(davp_folder, filename))
#         else:
#             shutil.move(file_path, os.path.join(others_folder, filename))

#         print(f"{filename} → {category}")

# print("Classification Completed!")





#------------------------------------Approach basis of the data inside the PDFs--------------------------------




import os
import shutil
import pdfplumber

SOURCE_FOLDER = r"C:\Users\admin\OneDrive\Desktop\OCR\RO input"
DEST_FOLDER = r"C:\Users\admin\OneDrive\Desktop\OCR\Sorted_PDFs"


samvad_folder = os.path.join(DEST_FOLDER, "SAMVAD")
davp_folder = os.path.join(DEST_FOLDER, "DAVP")
others_folder = os.path.join(DEST_FOLDER, "Others")


SAMVAD_KEYWORDS = [
    "SOCIETY FOR ADVANCED MANAGEMENT OF COMMUNICATION",
    "SAMVAD",
    "MD -Cum- CEO"
]

CBC_KEYWORDS = [
    "CENTRAL BUREAU OF COMMUNICATION",
    "Government of India",
    "RO Code",
    "CBC"
]

def classify_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""

        text = text.upper()

        if any(keyword.upper() in text for keyword in SAMVAD_KEYWORDS):
            return "SAMVAD"

        elif any(keyword.upper() in text for keyword in CBC_KEYWORDS):
            return "DAVP"

        else:
            return "Others"

    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return "Others"


# ===== PROCESS FILES =====
for filename in os.listdir(SOURCE_FOLDER):
    if filename.lower().endswith(".pdf"):
        file_path = os.path.join(SOURCE_FOLDER, filename)
        category = classify_pdf(file_path)

        if category == "SAMVAD":
            shutil.move(file_path, os.path.join(samvad_folder, filename)) # path for SAMVAD folder
        elif category == "DAVP":
            shutil.move(file_path, os.path.join(davp_folder, filename)) # path for DAVP folder
        else:
            shutil.move(file_path, os.path.join(others_folder, filename))

        print(f"{filename} → {category}")

print("Classification Completed!")
