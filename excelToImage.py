import os
import win32com.client as win32
import fitz  # PyMuPDF library
from PIL import Image

# Define the directories
excel_dir = "C:\\Users\\tthompson\\excelToImage\\input"
output_dir = "C:\\Users\\tthompson\\excelToImage\\output"

# Create the output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

print("Starting conversion process...")

# Create a COM object for Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

try:
    for filename in os.listdir(excel_dir):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            excel_path = os.path.join(excel_dir, filename)
            base_name = os.path.splitext(filename)[0]

            print(f"Processing '{filename}'...")
            
            temp_pdf_path = os.path.join(output_dir, f"temp_{base_name}.pdf")
            
            try:
                # Step 1: Open the Excel workbook and apply formatting
                workbook = excel.Workbooks.Open(excel_path)
                
                # Iterate through each worksheet and apply formatting
                for worksheet in workbook.Worksheets:
                    # Disable Zoom to allow FitToPages to work
                    worksheet.PageSetup.Zoom = False
                    
                    # Set page orientation to landscape (xlLandscape = 2)
                    worksheet.PageSetup.Orientation = 2
                    
                    # Set all columns to fit on one page
                    worksheet.PageSetup.FitToPagesWide = 1
                    worksheet.PageSetup.FitToPagesTall = False
                
                # Export the workbook as a standard PDF
                workbook.ExportAsFixedFormat(0, temp_pdf_path)
                workbook.Close(False)
                print(f"  -> Saved intermediate PDF: {temp_pdf_path}")
            except Exception as e:
                # Handle Excel's "nothing to print" error
                if "We didn't find anything to print" in str(e):
                    print(f"  -> Error: Excel file '{filename}' has no printable content. Skipping.")
                    continue
                else:
                    print(f"  -> Error processing {filename}: {e}")
                    continue

            # Step 2: Convert the temporary PDF pages to a single multi-page PDF
            try:
                doc = fitz.open(temp_pdf_path)
                images = []

                for page in doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(600 / 72, 600 / 72))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images.append(img)
                
                if images:
                    combined_pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
                    images[0].save(combined_pdf_path, save_all=True, append_images=images[1:])
                    print(f"  -> Saved flattened multi-page PDF: {combined_pdf_path}")
                
                doc.close()
                os.remove(temp_pdf_path)
                print(f"  -> Removed temporary PDF.")

            except Exception as e:
                print(f"  -> Error during image processing for {filename}: {e}")

finally:
    excel.Quit()
    print("\nConversion complete! ✨")