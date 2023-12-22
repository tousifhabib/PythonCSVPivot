import tabula

def convert_pdf_to_excel(pdf_file_path, excel_file_path):
    try:
        tabula.convert_into(pdf_file_path, excel_file_path, output_format='xlsx', pages='all')
        print("Excel generation from PDF completed.")
    except Exception as e:
        print(f"Error converting PDF to Excel: {e}")
