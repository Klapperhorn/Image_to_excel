# Image_to_excel
This notebook can be used to rename &amp; resize images in differnt folders and put them into an excel sheet. This is useful to transparently annotate the corpus creation in qualitative studies that employ images (e.g. cartoons or tweets).
It uses openpyxl to put images into an excel sheet.

- Specify the Excel output filename
excel_filename="2025-04-10 Political_Cartoons.xlsx"

- Specify the two folders you want to search
folders_to_search = ['political cartoons saved\\']

- specift image endings to be included from these folders:
file_endings='*.*g'

- Specify TEMP image folder (for resized and renamed images)
temp_folder='TEMP\\'
