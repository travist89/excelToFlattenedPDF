import os
import win32com.client as win32
import fitz
from PIL import Image
import io

# Define the root directory to start searching
input_dir = "C:\\Users\\tthompson\\excelToImage\\input"
# Define a temporary directory for intermediate files
temp_dir = "C:\\Users\\tthompson\\excelToImage\\temp"

# Create the temporary directory if it doesn't exist
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)

print("Starting conversion process...")

# Create a COM object for Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

try:
    # Use os.walk to search through all subfolders
    for root, dirs, files in os.walk(input_dir):
        for filename in files:
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                excel_path = os.path.join(root, filename)
                base_name = os.path.splitext(filename)[0]

                print(f"Processing '{excel_path}'...")
                
                # Define the paths for the temporary and final PDFs
                temp_pdf_path = os.path.join(temp_dir, f"temp_{base_name}.pdf")
                final_pdf_path = os.path.join(root, f"{base_name}.pdf")

                try:
                    # Step 1: Open the Excel workbook and apply formatting
                    workbook = excel.Workbooks.Open(excel_path)

                    for worksheet in workbook.Worksheets:
                        worksheet.PageSetup.Zoom = False
                        worksheet.PageSetup.Orientation = 2
                        worksheet.PageSetup.FitToPagesWide = 1
                        worksheet.PageSetup.FitToPagesTall = False

                    workbook.ExportAsFixedFormat(0, temp_pdf_path)
                    workbook.Close(False)
                    print(f"  -> Saved intermediate PDF to temp folder: {temp_pdf_path}")
                except Exception as e:
                    if "We didn't find anything to print" in str(e) or "Open method of Workbooks class failed" in str(e):
                        print(f"  -> Error: Excel file has no printable content, is corrupted, or could not be opened. Skipping.")
                        continue
                    elif "The RPC server is unavailable" in str(e):
                        print(f"  -> Error: The COM server connection was lost. Retrying...")
                        continue
                    else:
                        print(f"  -> Error processing {filename}: {e}")
                        continue

                # Step 2: Convert the temporary PDF pages to a single multi-page PDF
                try:
                    doc = fitz.open(temp_pdf_path)
                    output_doc = fitz.open()

                    for page_num in range(len(doc)):
                        # Render the page to a pixmap at a lower DPI
                        pixmap = doc.load_page(page_num).get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))

                        # Use a buffer to handle compressed image data
                        img_buffer = io.BytesIO()
                        # Convert to a Pillow Image and save to the buffer with compression
                        img = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                        img.save(img_buffer, format='jpeg', quality=75)
                        
                        # Create a new page and insert the compressed image from the buffer
                        page = output_doc.new_page(width=pixmap.width, height=pixmap.height)
                        page.insert_image(page.rect, stream=img_buffer)

                    # Save the final multi-page PDF
                    output_doc.save(final_pdf_path, garbage=3, deflate=True)

                    print(f"  -> Saved flattened multi-page PDF to original directory: {final_pdf_path}")

                    doc.close()
                    output_doc.close()
                    os.remove(temp_pdf_path)
                    print(f"  -> Removed temporary PDF.")
                
                except Exception as e:
                    print(f"  -> Error during image processing for {filename}: {e}")

finally:
    excel.Quit()
    print("\nConversion complete! ✨")