import os
from PIL import Image
from openpyxl import Workbook
from paddleocr import PaddleOCR
from PIL import Image, ImageOps
import sys
import getopt

from engine import *

folder_path = "image"
excel_file = "output_coordinates.xlsx"

# Create a new workbook and add a worksheet
wb = Workbook()
ws = wb.active
ws.title = "Image Coordinates"

book_name = "Buom_hoa_tan_truyen."

# Add headers to the Excel file
ws.append(["ID","Image Name", "BBox", "SinoNom"])

for filename in os.listdir(folder_path):
    # Check if the filename matches the pattern "aaad.xx.jpeg"
    if filename.startswith("page_") and filename.endswith(".png"):
        # Construct the new filename
        new_filename = filename.replace("page_", book_name, 1)  # Replace only the first occurrence of "aaad"
        
        # Construct full paths
        old_path = os.path.join(folder_path, filename)
        new_path = os.path.join(folder_path, new_filename)
        
        # Rename the file
        os.rename(old_path, new_path)
        print(f"Renamed: {old_path} -> {new_path}")

exit()

ocr = PaddleOCR(
    use_gpu=False,
    use_angle_cls=False,
    cls=False,
    det=False,
    rec=True,
    rec_model_dir="model/rec",
    rec_image_inverse=False,
    max_text_length=1,
    rec_char_dict_path="model/rec/vocab.txt",
    use_space_char=False,
    lang="ch",  # Specify language (e.g., 'ch' for Chinese)
)


OCNE = OCR_Chu_Nom_Engine()
    # Iterate through all files in the folder
for i in range(1,5):
    # Construct the full path to the file
    filename = f"{book_name}.{i}.jpeg"
    file_path = os.path.join(folder_path, filename)
    
    #command = f"python predict.py --input={file_path}"
    #os.system(command)  

    OCNE.pipeline(image_path=file_path, print_img=False)

    file_path2 = os.path.splitext(file_path)[0]
    file_path3 = os.path.splitext(filename)[0]
    coordinates_file = f"result\\char\\{book_name}.txt"
    output_folder = "result\\output\\crop"

    image = Image.open(f"{file_path}")


    os.makedirs(output_folder, exist_ok=True)


    with open(coordinates_file, "r", encoding="utf-8") as file:
        coordinates_data = []
        cnt = 0

        # Read coordinates from the file
        for line in file:
            parts = line.strip().split()
            
            if len(parts) == 5:  # Ensure correct format
                label, y1, x1, y2, x2 = parts
                x1, y1, x2, y2 = map(float, [x1, y1, x2, y2])  # Convert to float

                if x2 - x1 <= 40 or y2 - y1 <= 40:
                    continue
                
                # Collect coordinates and their respective y-values for grouping horizontally
                coordinates_data.append((label, x1, y1, x2, y2))

        # Sort coordinates by y1 to group horizontally
        coordinates_data.sort(key=lambda x: x[2])  # Sort by y1 (vertical position)

        # Group coordinates with similar y-values (horizontal line group)
        grouped_coordinates = []
        current_group = []
        previous_y1 = None

        for data in coordinates_data:
            label, x1, y1, x2, y2 = data
            if previous_y1 is None or abs(y1 - previous_y1) <= 20:  # Allow small variation for grouping
                current_group.append((x1, y1, x2, y2))
            else:
                grouped_coordinates.append(current_group)
                current_group = [(x1, y1, x2, y2)]  # Start new group
            previous_y1 = y1

        if current_group:
            grouped_coordinates.append(current_group)  # Add last group

    for group in grouped_coordinates:
        group.sort(key = lambda x:x[0])
        cnt += 1
        cnt2 = 0
        for data in group:
            try:
                x1, y1, x2, y2 = data
                cnt2 += 1

                # Crop the image for each character
                cropped_image = image.crop((x1, y1, x2, y2))

                # Save the cropped image
                output_path = os.path.join(output_folder, f"{file_path3}.{cnt}.{cnt2}.jpeg")
                cropped_image.save(output_path)
                print(f"Cropped and saved: {output_path}")

                img = Image.open(output_path).convert("RGB")
    
                # Resize the image to 40x40 while maintaining aspect ratio
                img = ImageOps.fit(img, (40, 40), method=Image.Resampling.LANCZOS)
                
                # Create a blank white image of size 48x320
                field = Image.new("RGB", (320, 48), (255, 255, 255))
                
                # Calculate coordinates to center the 40x40 image in the 48x320 field
                x_offset = (320 - 40) // 2
                y_offset = (48 - 40) // 2
                
                # Paste the resized character into the center of the field
                field.paste(img, (x_offset, y_offset))
                
                # Save the result
                field.save(output_path)

                # Add the data to the Excel sheet
                coordinates = f"[({x1},{y1}),({x2},{y1}),({x1},{y2}),({x2},{y2})]"
                #ws.append([f"{file_path3}.{cnt}.{cnt2}", filename, coordinates])

                result = ocr.ocr(output_path, cls=False, det=False, rec=True)

                # Extract the recognized text and confidence
                if result:
                    recognized_text = result[0][0][0]
                    confidence = result[0][0][1]
                else:
                    recognized_text = ""
                    confidence = 0

                # Add the data to the Excel sheet
                ws.append([f"{file_path3}.{cnt}.{cnt2}", filename, coordinates, recognized_text, confidence])

            except Exception as e:
                print(f"Error processing file {file_path3}.{cnt}.{cnt2}: {e}")

# Save the Excel file
wb.save(excel_file)
print(f"Excel file saved as {excel_file}")