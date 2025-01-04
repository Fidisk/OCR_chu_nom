import pandas as pd
from PIL import Image
import ast

# Load the Excel file
file_path = 'output_coordinates.xlsx'  # Replace with the path to your Excel file
df = pd.read_excel(file_path)

# Dictionary to store image dimensions
image_dimensions = {}

# Calculate width and height for each unique image name
for image_name in df['Image Name'].unique():
    image_path = f'image/{image_name}'  # Replace with your image folder path
    with Image.open(image_path) as img:
        width, height = img.size
        image_dimensions[image_name] = (width, height)

# Normalize the BBox coordinates
def normalize_bbox(bbox_str, image_name, maxsize = 1200):
    # Convert the BBox string to a list of tuples
    bbox = ast.literal_eval(bbox_str)
    width, height = image_dimensions[image_name]
    if (width+height<maxsize):
        return bbox
    factor = (width+height)/1200
    # Normalize coordinates
    normalized_bbox = [(int(x / factor), int(y / factor)) for x, y in bbox]
    return normalized_bbox

# Apply normalization
df['Normalized BBox'] = df.apply(
    lambda row: normalize_bbox(row['BBox'], row['Image Name']), axis=1
)

# Save the updated DataFrame to a new Excel file
output_file = 'normalized_bboxes.xlsx'
df.to_excel(output_file, index=False)
print(f'Normalized bounding boxes saved to {output_file}')
