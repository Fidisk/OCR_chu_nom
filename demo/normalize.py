import pandas as pd
from PIL import Image
import ast
import json

# Load the Excel file
file_path = 'output_coordinates.xlsx'  # Replace with the path to your Excel file
df = pd.read_excel(file_path)

image_dimensions = {}

# Calculate width and height for each unique image name
for image_name in df['Image Name'].unique():
    image_path = f'image/{image_name}'  # Replace with your image folder path
    with Image.open(image_path) as img:
        width, height = img.size
        image_dimensions[image_name] = (width, height)

# Load the JSON file
with open("label.json", "r", encoding="utf-8") as file:
    data = json.load(file)

# Normalize points
def normalize_points(points, image_name):
    width, height = image_dimensions[image_name]
    print(points)
    print(width, height)
    if (width+height<1200):
        return points
    factor = (width + height) / 1200.0
    normalized = [[((width/factor-x) * factor), (y * factor)] for x, y in points]
    return normalized

for item in data:
    item["image_name"] = item["image_name"].replace(" ", "_")

with open("output.json", "w", encoding="utf-8") as file:
    json.dump(data, file, indent=4, ensure_ascii=False)

with open("output.json", "r", encoding="utf-8") as file:
    data = json.load(file)

# Process the JSON data
for item in data:
    image_name = item["image_name"]
    for ocr_detail in item["ocr_details"]:
        points = ocr_detail["points"]
        ocr_detail["normalized_points"] = normalize_points(points, image_name)

# Save the updated JSON file
with open("output.json", "w", encoding="utf-8") as file:
    json.dump(data, file, indent=4, ensure_ascii=False)

print("Normalization completed and saved to output.json.")
