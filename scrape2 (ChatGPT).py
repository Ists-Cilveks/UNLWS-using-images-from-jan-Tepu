import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TextProperties
from odf.table import Table, TableRow, TableCell
from odf.text import P, Span
import re
from odf.draw import Frame, Image
import xml.etree.ElementTree as ET
import shutil
import os


def get_svg_dimensions(image_path):
    """
    Get the width and height of an SVG image from its metadata.

    Args:
        image_path (str): Path to the SVG image file.

    Returns:
        tuple: A tuple containing the width and height (in pixels) as strings.
    """
    try:
        tree = ET.parse(image_path)
        root = tree.getroot()
        width = root.get("width") + "px"
        height = root.get("height") + "px"
        return width, height
    except Exception as e:
        print(f"Failed to parse SVG metadata: {str(e)}")
        return None, None

def create_image_cell(image_path, ods_doc):
    """
    Create a TableCell containing an image from the given image_path and add it to the ods_doc.

    Args:
        image_path (str): Path to the image file (e.g., 'images/people.svg').
        ods_doc (OpenDocumentSpreadsheet): The ODS document to which the image cell will be added.

    Returns:
        TableCell: A TableCell containing the image.
    """
    # Get the width and height from SVG metadata
    width, height = get_svg_dimensions(image_path)
    
    # Create a frame for the image with width and height attributes
    # image_frame = Frame(width=width, height=height)
    image_frame = Frame(width="50px", height="50px")
    
    # Add the image to the frame
    href = ods_doc.addPicture(image_path)
    image = Image(href=href)
    image_frame.addElement(image)

    # Create a table cell and add the image frame to it
    image_cell = TableCell(valuetype="string")
    image_cell.addElement(image_frame)

    return image_cell


# Define the URL of the webpage
base_url = "https://davidar.github.io/tp/kama-sona"
url = base_url

# Send an HTTP GET request to the URL
response = requests.get(url)

# Check if the request was successful
if response.status_code == 200:
    # Parse the HTML content of the page using BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find the table on the webpage
    table = soup.find('table')

    # Initialize an empty list to store the scraped rows
    data = []

    # Iterate through the rows in the table
    for row in table.find_all('tr'):
        columns = row.find_all('td')
        if len(columns) >= 2:
            # Extract the image relative path (assuming the SVG is in the first column)
            image_path = columns[0].find('img')['src']
            # Combine the base URL with the relative path to get the full URL
            image_url = urljoin(base_url, image_path)

            # Extract the text (assuming it's in the second column)
            text = columns[1].encode_contents().decode()  # Keep only the content of the second column

            # Remove all HTML tags except <em> tags
            text = re.sub(r'<(?!/?(em)\b)[^>]*>', '', text)

            # Append the extracted data as a tuple to the data list
            data.append(("Row "+str(len(data))+", Cell ", image_url, text))

    # # Print the scraped rows
    # for row in data:
    #     print("Image URL:", row[0])
    #     print("Text:", row[1])
    #     print()

    # Create a new ODS document
    ods_doc = OpenDocumentSpreadsheet()

    # Create a table in the ODS document
    ods_table = Table(name="kama sona data")

    # Add headers to the table
    header_row = TableRow()
    header_row.addElement(TableCell(valuetype="string", value="Column 1"))
    header_row.addElement(TableCell(valuetype="string", value="Column 2"))
    ods_table.addElement(header_row)


    
    # # Sample data with emphasized text
    # data = [
    #     ("Row 1, Cell 1", "This is a <em>sample</em> text with <em>emphasis</em> in color."),
    #     ("Row 2, Cell 1", "Another example without emphasis."),
    #     ("Row 3, Cell 1", "Here, we have <em>more</em> <em>emphasized</em> text."),
    # ]

    # data = []
    # for i in range(len(data)):
    #     row = data[i]
    #     data.append(("Row "+str(i+1)+", Cell ", row[0], row[1]))

    # Style definition for emphasized text color
    style = Style(
        name="EmphasizedTextColor",
        family="text",
    )
    text_properties = TextProperties(
        color="#FFA500",  # Orange color for emphasis
    )
    style.addElement(text_properties)
    ods_doc.automaticstyles.addElement(style)

    editable_image_folder = "editable_images"
    os.makedirs(editable_image_folder, exist_ok=True)
    
    img_number = 0
    # Iterate through the sample data and add rows to the table
    for item in data:
        row = TableRow()
        # cell1 = TableCell(valuetype="string", value=item[0]+"1")
        
        original_image_path = "images/"+item[1].split("/")[-1] # The image from the website (downloaded)
        safe_editable_image_name = f"{str(img_number)} {item[2]}".replace(":", ",").replace("<em>", "").replace("</em>", "").replace("zzzz", "")
        editable_image_path = f"{editable_image_folder}/{safe_editable_image_name}.svg" # The image that will be drawn on top of
        
        print("images/"+item[1].split("/")[-1])
        cell1 = create_image_cell(original_image_path, ods_doc)
        cell2 = TableCell(valuetype="string")
        
        # Duplicate images with numbers and toki pona names
        shutil.copyfile(original_image_path, editable_image_path)
        img_number+=1
        
        # Create paragraphs
        # svg_paragraph = P()
        # svg_paragraph.addText(item[1])
        paragraph = P()
        
        # Split the text by "<em>" tags to handle emphasis
        parts = re.split(r'(<em>.*?</em>)', item[2])
        for part in parts:
            if part.startswith("<em>") and part.endswith("</em>"):
                # Handle emphasized text
                em_text = part[4:-5]  # Remove "<em>" and "</em>" tags
                em_span = Span(stylename="EmphasizedTextColor")
                em_span.addText(em_text)  # Use addText to add text to the span
                paragraph.addElement(em_span)
            else:
                # Regular text
                paragraph.addText(part)

        # cell1.addElement(svg_paragraph)
        cell2.addElement(paragraph)
        row.addElement(cell1)
        row.addElement(TableCell())
        row.addElement(cell2)
        ods_table.addElement(row)

    # Add the table to the ODS document
    ods_doc.spreadsheet.addElement(ods_table)

    # Save the ODS document
    ods_doc.save("scraped data.ods")

    print("ODS file with emphasized text created: scraped data.ods")

else:
    print("Failed to retrieve the webpage")
