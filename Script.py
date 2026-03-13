from docx import Document
from datetime import datetime
import os
import shutil
from docx2pdf import convert
import zipfile
import tempfile

TEMPLATE_PATH = 
OUTPUT_FOLDER = 
PEP_FOLDER = 

def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%d.%m.%Y").strftime("%d.%m.%Y")
    except ValueError:
        print("❌ Use DD.MM.YYYY format for the date.")
        return None

def calculate_precise_font_size(name: str) -> int:
    """
    Starts at 32pt for <=30 chars.
    Decreases ~0.41pt per character up to 69 chars.
    Floors at 8pt.
    Returns size in Word half-points.
    """
    length = len(name.strip())
    
    # Linear decrease
    size_pt = 28 - ((length - 30) * 0.3)

    size_pt = max(8, round(size_pt))  # Floor and round
    return size_pt * 2  # Convert to Word half-points

def simple_xml_replace(file_path, placeholder, replacement, name_for_font_size=None):
    with zipfile.ZipFile(file_path, 'r') as docx_zip:
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_zip.extractall(temp_dir)
            doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            if os.path.exists(doc_xml_path):
                with open(doc_xml_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                if placeholder in content:
                    if placeholder == '{{NAME}}' and name_for_font_size:
                        font_size = calculate_precise_font_size(name_for_font_size)
                        replacement_xml = f"""
<w:r>
  <w:rPr>
    <w:rFonts w:ascii=\"Daytona\" w:hAnsi=\"Daytona\" w:eastAsia=\"Daytona\" w:cs=\"Daytona\"/>
    <w:sz w:val=\"{font_size}\"/>
    <w:b/>
  </w:rPr>
  <w:t>{replacement}</w:t>
</w:r>
"""
                        content = content.replace(f"<w:t>{placeholder}</w:t>", replacement_xml)
                    else:
                        content = content.replace(placeholder, replacement)
                    with open(doc_xml_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    new_zip_path = file_path.replace('.docx', '_modified.docx')
                    with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                file_path_full = os.path.join(root, file)
                                arcname = os.path.relpath(file_path_full, temp_dir)
                                new_zip.write(file_path_full, arcname)
                    return new_zip_path
    return None

def review_and_edit_entries(entries):
    while True:
        print("\n📋 Current Entries:")
        for i, (name, date, next_date) in enumerate(entries):
            print(f"{i + 1}. {name} | {date}")
        choice = input("\n🔧 Type a number to edit, or 'print' to generate PDFs: ").strip().lower()
        if choice == 'print':
            return entries
        elif choice.isdigit() and 1 <= int(choice) <= len(entries):
            idx = int(choice) - 1
            name, date, _ = entries[idx]
            new_name = input(f"Edit Name [{name}]: ").strip() or name
            new_date_input = input(f"Edit Date [{date}] (DD.MM.YYYY): ").strip()
            new_date = format_date(new_date_input) if new_date_input else date
            if new_date:
                next_date = datetime.strptime(new_date, "%d.%m.%Y").replace(year=datetime.strptime(new_date, "%d.%m.%Y").year + 1).strftime("%d.%m.%Y")
                entries[idx] = (new_name, new_date, next_date)
        else:
            print("❌ Invalid input. Try again.")

def main():
    name_date_pairs = []
    print("📥 Enter names and dates (type 'done' to finish):")
    while True:
        name = input("Name: ").strip()
        if name.lower() == 'done':
            break
        date_str = input("Date (DD.MM.YYYY): ").strip()
        formatted_date = format_date(date_str)
        if formatted_date:
            next_date = datetime.strptime(formatted_date, "%d.%m.%Y").replace(year=datetime.strptime(formatted_date, "%d.%m.%Y").year + 1).strftime("%d.%m.%Y")
            name_date_pairs.append((name, formatted_date, next_date))

    if not name_date_pairs:
        print("No entries provided. Exiting.")
        return

    name_date_pairs = review_and_edit_entries(name_date_pairs)

    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
    if not os.path.exists(PEP_FOLDER):
        os.makedirs(PEP_FOLDER)

    for name, date, next_date in name_date_pairs:
        print(f"\n🔄 Processing certificate for {name}...")
        temp_file = os.path.join(OUTPUT_FOLDER, f"temp_{name.replace(' ', '_')}.docx")
        shutil.copy2(TEMPLATE_PATH, temp_file)

        modified_file = simple_xml_replace(temp_file, '{{NAME}}', name, name_for_font_size=name)
        if modified_file:
            os.replace(modified_file, temp_file)
        modified_file = simple_xml_replace(temp_file, '{{DATE}}', date)
        if modified_file:
            os.replace(modified_file, temp_file)
        modified_file = simple_xml_replace(temp_file, '{{NEXTDATE}}', next_date)
        if modified_file:
            os.replace(modified_file, temp_file)

        parts = name.split(" - ")
        first_part = parts[0].strip().lower()
        pep_variants = ["pep", "pep home", "pep cell"]
        is_pep = first_part in pep_variants

        if is_pep and len(parts) >= 3:
            filename = f"{parts[1].strip()} - {parts[2].strip()}"
        elif not is_pep and len(parts) >= 3:
            filename = f"{parts[0].strip()} - {parts[1].strip()} - {parts[2].strip()}"
        else:
            filename = name

        output_folder = PEP_FOLDER if is_pep else OUTPUT_FOLDER
        docx_path = os.path.join(output_folder, f"{filename}.docx")
        pdf_path = os.path.join(output_folder, f"{filename}.pdf")

        os.rename(temp_file, docx_path)
        print(f"📄 DOCX saved: {docx_path}")

        try:
            convert(docx_path, pdf_path)
            os.remove(docx_path)
            print(f"✅ PDF saved: {pdf_path}")
        except Exception as e:
            print(f"❌ PDF conversion failed: {e}")
            print(f"📄 DOCX file kept: {docx_path}")

if __name__ == "__main__":
    main()