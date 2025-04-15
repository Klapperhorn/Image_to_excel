# Image_to_excel
This notebook can be used to rename &amp; resize images in differnt folders and put them into an excel sheet. This is useful to transparently annotate the corpus creation in qualitative studies that employ images (e.g. cartoons or tweets).
It uses openpyxl to put images into an excel sheet.

Users specify
- the Excel output filename
- the folders (subfolers) you want to search for image files 
- image endings to be included from these folders
- the TEMP image folder for resized and renamed images (if needed).
