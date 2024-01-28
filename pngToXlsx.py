from PIL import Image
from openpyxl import Workbook
import argparse

# use https://jamestomasino.github.io/stitchy/ to downsample you image and put a grid over it


def parse_args():
    parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter,
    description='''\
    This is a script which takes in an image and converts it into an xlsx document where the cell values are the RGB value of the image. 
    This is to crochet a pattern, and use google sheets to track it as you go.

    1: Find an image you would like to turn into a crochet pattern
    2: Run the image through this tool https://jamestomasino.github.io/stitchy/ which will downsample the image into a grid. Save the output image
    3: Use the output of step 2 as the input to this script, and run this script, providing the necessary information
    4: Upload the output xlsx document to google sheets
    5: Run the google apt script colorCellsBasedOnRGB.gs on xlsx document to color the cells in the sheet (instructions on how to run the apt script are commented in that file)
    ''')
    parser.add_argument('input_image_path')
    parser.add_argument('output_excel_path')
    parser.add_argument('rows', type=int)
    parser.add_argument('columns', type=int)
    return parser.parse_args()

def jpeg_to_excel(input_image_path, output_excel_path, rows, columns):
    # Open the JPEG image
    img = Image.open(input_image_path)

    # Get image dimensions
    width, height = img.size
    print(width, height)

    # Create a new Excel workbook and get the active sheet
    workbook = Workbook()
    sheet = workbook.active

    excel_row = 1
    excel_col = 1

    mega_pixel_width = width / columns
    mega_pixel_height = height / rows

    # Loop through each pixel in the image and write RGB values to the Excel sheet
    for row in range(rows-1):
        for col in range(columns-1):
            # get the RBG value of a pixel roughly in the center of the megapixel
            rgb = img.getpixel((col*mega_pixel_width + (mega_pixel_width/2), row*mega_pixel_height + (mega_pixel_height/2)))[0:3]
            rgb_string = ','.join(map(str, rgb))

            # Write RGB values to the Excel sheet
            sheet.cell(row=excel_row, column=excel_col).value = rgb_string
            excel_col = excel_col + 1
        excel_row = excel_row + 1
        excel_col = 1

    # Save the Excel file
    workbook.save(output_excel_path)
    print(f"Conversion complete. Excel file saved at {output_excel_path}")

if __name__ == "__main__":
    # Specify the input JPEG image and output Excel file paths
    input_image_path = "stitchy-pattern.png"
    output_excel_path = "output_excel.xlsx"

    args = parse_args()


    # Convert JPEG to Excel
    jpeg_to_excel(args.input_image_path, args.output_excel_path, args.rows, args.columns)

