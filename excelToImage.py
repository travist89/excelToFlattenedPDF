import os
import win32com.client as win32
import fitz

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
            final_pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

            try:
                # Step 1: Open the Excel workbook and apply formatting
                workbook = excel.Workbooks.Open(excel_path)

                # Iterate through each worksheet and apply formatting
                for worksheet in workbook.Worksheets:
                    worksheet.PageSetup.Zoom = False
                    worksheet.PageSetup.Orientation = 2
                    worksheet.PageSetup.FitToPagesWide = 1
                    worksheet.PageSetup.FitToPagesTall = False

                # Export the workbook as a standard PDF
                workbook.ExportAsFixedFormat(0, temp_pdf_path)
                workbook.Close(False)
                print(f"  -> Saved intermediate PDF: {temp_pdf_path}")
            except Exception as e:
                if "We didn't find anything to print" in str(e):
                    print(f"  -> Error: Excel file '{filename}' has no printable content. Skipping.")
                    continue
                else:
                    print(f"  -> Error processing {filename}: {e}")
                    continue

            # Step 2: Convert the temporary PDF pages to a single multi-page PDF
            try:
                doc = fitz.open(temp_pdf_path)
                output_doc = fitz.open()

                for page_num in range(len(doc)):
                    # Render the page to a pixmap at 600 DPI
                    pixmap = doc.load_page(page_num).get_pixmap(matrix=fitz.Matrix(600 / 72, 600 / 72))

                    # Create a new page with the pixmap dimensions
                    page = output_doc.new_page(width=pixmap.width, height=pixmap.height)
                    
                    # Insert the compressed image onto the new page
                    page.insert_image(page.rect, pixmap=pixmap)

                # Save the final multi-page PDF with aggressive compression
                output_doc.save(final_pdf_path, garbage=3, deflate=True, compress_images=fitz.PDF_IMAGE_COMPRESS_JPEG, compress_image_quality=75)

                print(f"  -> Saved flattened multi-page PDF: {final_pdf_path}")

                doc.close()
                output_doc.close()
                os.remove(temp_pdf_path)
                print(f"  -> Removed temporary PDF.")

            except Exception as e:
                print(f"  -> Error during image processing for {filename}: {e}")

finally:
    excel.Quit()
    print("\nConversion complete! ✨")