
# """
# Code for inserting images into spreadsheet:
# Author: Gareth Watkins, Collections Data Manager @ Te Papa Tongarewa, Museum of New Zealand
# """

# Import modules
# import os
import xlsxwriter
# from io import BytesIO
# from PIL import Image

# Define our folder that contains the source images and the name of the output file (note each time the script is run
# it will overwrite the output file).
# The folder where the images are located:
working_folder = "spreadsheet output/"

# The name of the output file. It will be created in the same folder as the source images:
out_file = working_folder + 'result.xlsx'

# # Define the size of the image cell in the spreadsheet and the slightly smaller max size of the thumbnail
# cell_size = 95
# max_image_size = (cell_size-5, cell_size-5)

# # Generate a list of jpg images from the working folder (and any child folders)
# file_list = []
# for root, dirs, files in os.walk(working_folder):
#   for file in files:
#     if file.endswith(".jpg"):
#       file_list.append(os.path.join(root, file))

# Create a workbook. Col A will contain filenames, Col B will contain images
workbook = xlsxwriter.Workbook(out_file)
worksheet = workbook.add_worksheet()

# # Size the image column (B) in the worksheet
# worksheet.set_column_pixels('B:B', cell_size)

# Add column headers to the worksheet
worksheet.write('A1', 'Original image')
worksheet.write('B1', 'toki pona')
worksheet.write('C1', 'UNLWS image')

# Define which row we start at (row 1 being the column headers)
row = 2


from bs4 import BeautifulSoup as BS
# import urllib2
import urllib.request

html = urllib.request.urlopen("https://davidar.github.io/tp/kama-sona")
soup = BS(html)
# elem = soup.findAll('tr', {'title': 'title here'})
raw_rows = soup.findAll('tr')
td_rows = []
for r in raw_rows:
  r = r.findAll('td')
  if not r:
    continue
  td_rows.append(r)
  speech = r[0].find(['.speech-left', '.speech-right'])
  if speech:
    pass
  r[0] = r[0].find()
  
  image_link = r[0].find()
  
  this_filename_cell = 'A' + str(row)
  worksheet.write(this_filename_cell, image_link)
# print(td_rows)
# elem[0].text



# # Main loop to iterate through the file_list and add each image to the worksheet
# for full_filename in file_list:

  # # Get the name of the current image
  # this_filename = os.path.basename(full_filename)

#   # As a courtesy, output the filename of the current image to the console window so that we can track progress
#   print("Processing " + this_filename)

#   # Define the worksheet cells we are about populate
  # this_filename_cell = 'A' + str(row)
#   this_image_cell = 'B' + str(row)

#   # Set the size of the current row in the worksheet.
#   # Note the row has an offset of -1, e.g. 0 actually equals the first row in the spreadsheet
#   worksheet.set_row_pixels(row - 1, cell_size)

  

  # # Load the current image into memory
  # im = Image.open(full_filename)

  # # Reduce the image using the thumbnail method
  # im.thumbnail(max_image_size)

  # # Sometimes the files may be encoded in the RGBA colour space so we re-encode as RGB
  # if im.mode in ("RGBA", "P"):
  #   im = im.convert("RGB")

  # # Save the modified image into memory
  # this_modified_image = BytesIO()
  # im.save(this_modified_image, format="JPEG", quality=50)

  # # Write the modified image into Col B of the worksheet.
  # # We offset the image slightly (2px) so that it falls within the cell
  # # Note: setting object_position to 1 sets the image property in the worksheet so that it 
  # # will move and resize (allowing for user filtering)
  # worksheet.insert_image(this_image_cell, this_filename, {'image_data': this_modified_image, 'object_position': 1,
  # 'x_offset': 2, 'y_offset': 2})

  # # Advance the row pointer by 1
  # row = row + 1

# Close the workbook once the loop has finished
workbook.close()