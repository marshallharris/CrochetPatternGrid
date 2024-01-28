    This is a script which takes in a png image and converts it into an xlsx document where the cell values are the RGB value of the image. 
    This is to crochet a pattern, and use google sheets to track the progress in the pattern.

    1: Find an image you would like to turn into a crochet pattern
    2: Run the image through this tool: https://jamestomasino.github.io/stitchy/ which will downsample the image into a grid with a specified number of colors. Save the output image.
    3: Use the output of step 2 as the input to pngToXlsx.py, providing the necessary information (grid columns, grid rows and output_path)
    4: Upload the output xlsx document to google sheets
    5: Run the google apt script colorCellsBasedOnRGB.gs on spreadsheet to color the cells in the sheet (instructions on how to run the apt script are commented in that file)
