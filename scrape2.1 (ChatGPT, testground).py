from odf.opendocument import OpenDocumentSpreadsheet
from odf.draw import Frame, Image
from odf.table import Table, TableRow, TableCell
import xml.etree.ElementTree as ET

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
    image_frame = Frame(width=width, height=height)
    
    # Add the image to the frame
    href = ods_doc.addPicture(image_path)
    image = Image(href=href)
    image_frame.addElement(image)

    # Create a table cell and add the image frame to it
    image_cell = TableCell(valuetype="string")
    image_cell.addElement(image_frame)

    return image_cell

# Create a new ODS document
ods_doc = OpenDocumentSpreadsheet()

# Create a table and add the cell to it
table = Table(name="Image Table")
table.addElement(TableRow())

image_cell = create_image_cell('images/people.svg', ods_doc)

table._get_lastChild().addElement(image_cell)

# Add the table to the ODS document
ods_doc.spreadsheet.addElement(table)

# Save the ODS document with the image
ods_doc.save("image_in_ods.ods")

print("Image inserted into the ODS file: image_in_ods.ods")
