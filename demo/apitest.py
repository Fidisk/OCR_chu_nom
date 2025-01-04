import requests
import json
import os
import re

def get_tor_session():
    session = requests.session()
    # Tor uses the 9050 port as the default socks port
    session.proxies = {'http':  'socks5://127.0.0.1:9050',
                       'https': 'socks5://127.0.0.1:9050'}
    return session

from stem import Signal
from stem.control import Controller

# signal TOR for a new connection 
def renew_connection():
    with Controller.from_port(port = 9051) as controller:
        controller.authenticate(password="password")
        controller.signal(Signal.NEWNYM)

def callAPI(image_path, session):
    url_upload = "https://tools.clc.hcmus.edu.vn/api/web/clc-sinonom/image-upload"
    url_ocr = "https://tools.clc.hcmus.edu.vn/api/web/clc-sinonom/image-ocr"
    url_download = "https://tools.clc.hcmus.edu.vn/api/web/clc-sinonom/image-download"

    headers = {
        "User-Agent": "MyWebApp v1.0"
    }


    #image_path = "D:\\NLP_Midterm\\image\\Bach_van_thi_tap.9.jpeg"

    with open(image_path, "rb") as f:
        files = {
            "image_file": f
        }
        response_upload = session.post(url_upload, headers=headers, files=files)
    if response_upload.status_code == 200:
        response_json_upload = response_upload.json()
        if response_json_upload.get("code") == "000000":
            file_name = response_json_upload["data"]["file_name"]
            print(f"Image uploaded successfully. File name: {file_name}")
        else:
            print(f"Error uploading image: {response_json_upload.get('message')}")
    else:
        print(f"HTTP Error while uploading: {response_upload.status_code}")
        exit(1)


    data_ocr = {
        "ocr_id": 4,  # Use 1 for administrative OCR
        "file_name": file_name  
    }

    response_ocr = session.post(url_ocr, headers=headers, json=data_ocr)

    if response_ocr.status_code == 200:
        response_json_ocr = response_ocr.json()
        if response_json_ocr.get("code") == "000000":
            # OCR result
            result_text = response_json_ocr["data"]["result_ocr_text"]
        
            #print(f"{response_json_ocr}")
            print(response_json_ocr["data"]["details"]["details"])
            return response_json_ocr["data"]["details"]["details"]
        else:
            print(f"Error in OCR: {response_json_ocr.get('message')}")
    else:
        print(f"HTTP Error while processing OCR: {response_ocr.status_code}")
def extract_number(file_name):
    match = re.search(r'(\d+)', file_name)
    return int(match.group(1)) if match else 0

def process_and_merge(folder_path, output_file):
    # Get all files in the folder
    files = os.listdir(folder_path)
    
    # Filter to include only common image file types
    image_extensions = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
    image_files = [file for file in files if os.path.splitext(file)[1].lower() in image_extensions]
    image_files.sort(key=extract_number)
    # Initialize a list to hold all OCR results
    merged_results = []
    cnt = 0
    # Process the first 20 images directly
    for i, image_file in enumerate(image_files):
        if cnt == 0:
            cnt = 10
            renew_connection()
            session = get_tor_session()
        cnt = cnt - 1 
        image_path = os.path.join(folder_path, image_file)
        print(f"Processing image {i + 1}: {image_file}")
        details = callAPI(image_path, session)
        
        # Add the details to the merged results with a reference to the image
        merged_results.append({
            "image_name": image_file,
            "ocr_details": details
        })
        print(f"Processed {image_file}")
    
    # Save all results into a single output file
    with open(output_file, "w", encoding="utf-8") as file:
        json.dump(merged_results, file, ensure_ascii=False, indent=4)
    print(f"All OCR results saved to {output_file}")

# Specify the folder containing images
image_folder = "image"  # Image folder path
output_file = os.path.join(image_folder, "label.json")  # Output file path

process_and_merge(image_folder, output_file)