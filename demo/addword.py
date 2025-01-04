import pandas as pd
import re

def normalize_text(text):
    text = text.lower()
    text = re.sub(r'[^\w\s]', '', text)  
    text = text.strip()  
    return text


# Step 1: Read the text file and extract words
with open("Bach van thi tap _Dich.txt", "r", encoding="utf-8") as file:
    words = normalize_text(file.read()).split()  # Split by whitespace to get individual words

# Step 2: Read the existing Excel file



df = pd.read_excel("output_coordinates.xlsx")

# Step 3: Add the words as a new column
# Ensure the number of words matches the number of rows in the DataFrame
if len(words) < len(df):
    words += [""] * (len(df) - len(words))  # Pad with empty strings if not enough words
elif len(words) > len(df):
    words = words[:len(df)]  # Truncate the list if too many words

df['NewColumn'] = words  # Add the words as a new column

# Step 4: Save the updated DataFrame back to an Excel file
df.to_excel("output_coordinates.xlsx", index=False)

print("Column added with words from the text file.")
