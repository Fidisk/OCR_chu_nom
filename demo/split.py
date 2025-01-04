from pdf2image import convert_from_path
import os

# Path to your PDF file
pdf_path = "aa.pdf"

# Output folder for images
output_folder = "BVTT"

# Convert PDF to images (one image per page)
pages = convert_from_path(pdf_path, dpi=300)  # Adjust dpi if needed

# Save each page as an image
for i, page in enumerate(pages):
    output_path = os.path.join(output_folder, f"page_{i+1}.png")
    page.save(output_path, 'PNG')

print(f"PDF split into {len(pages)} images.")
