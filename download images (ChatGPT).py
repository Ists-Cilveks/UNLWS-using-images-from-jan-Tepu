from odf.opendocument import load
import os
import requests
from odf.table import Table, TableRow, TableCell
from odf.text import P

# Load the ODS file
ods_file_path = "scraped data.ods"  # Replace with the path to your ODS file
doc = load(ods_file_path)

# Create a directory to store downloaded images
output_dir = "images"
os.makedirs(output_dir, exist_ok=True)

first_row = True

# Iterate through the first column (assumed to be in the first table)
table = doc.getElementsByType(Table)[0]  # Assuming the table is the first in the document
for row in table.getElementsByType(TableRow):
    if first_row:
        first_row = False
        continue
    # Get the first cell in each row (assuming it contains URLs)
    cell = row.getElementsByType(TableCell)[0]
    
    # Get the text (URL) from the cell
    url = cell.getElementsByType(P)[0].__str__()
    
    # Define the file name by extracting the last part of the URL
    file_name = os.path.join(output_dir, os.path.basename(url))
    
    # Send an HTTP GET request to download the image
    response = requests.get(url)
    
    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Save the image to the specified file path
        with open(file_name, "wb") as f:
            f.write(response.content)
        print(f"Downloaded: {url} -> {file_name}")
    else:
        print(f"Failed to download: {url}")

print("Downloaded all images.")
