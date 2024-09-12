# Libraries
import pytesseract                      # Optical Character Recognition
import cv2                              # Image Modification
import os                               # Directory handling
import pandas as pd                     # Excel sheets
import xlsxwriter                       # Modifying excel sheets
import matplotlib.pyplot as plt         # Displaying images and plots during debugging
import numpy as np                      # Handling arrays

# Tesseract Setup
pytesseract.pytesseract.tesseract_cmd = r"_internal\Tesseract-OCR\tesseract.exe"

# Function:    processImage
#    Input:    Path to image to process.
#   Output:    Preprocessed image ready for OCR.

def processImage(imagePath):

    # Get Image
    image = cv2.imread(imagePath)                           # Import image
    imageGray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)     # Turn it into grayscale

    # Apply Contrast
    alpha = 1.5  
    beta = 10
    contrasted = cv2.convertScaleAbs(imageGray, alpha=alpha, beta=beta) 

    # Bilateral Filter
    filtered = cv2.bilateralFilter(contrasted,9,75,75)

    # Segment the image by appling thresholds
    thresholded = cv2.adaptiveThreshold(filtered,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,11,2)

    # Remove background
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (20,20))   # Make kernel
    morph = cv2.morphologyEx(thresholded, cv2.MORPH_CLOSE, kernel)   # Apply morphology
    
    result = cv2.bitwise_and(contrasted, contrasted, mask=morph)     # Apply mask to image

    return result

# Function:    setupColumns
#    Input:    Name of dataframe to modify.
#   Output:    Modifies input dataframe when called.

def setupColumns(dataframe):
    dataframe.insert(1, "Processed Image", None, False)
    dataframe.insert(2, "Raw OCR", None, False)
    dataframe.insert(3, "OCR Serial", None, False)
    dataframe.insert(4, "OCR Brand", None, False)
    dataframe.insert(5, "OCR Model", None, False)
    dataframe.insert(6, "OCR Name", None, False)
    dataframe.insert(7, "Engineer Serial", None, False)
    dataframe.insert(8, "Engineer Brand", None, False)
    dataframe.insert(9, "Engineer Model", None, False)
    dataframe.insert(10, "Final Name", None, False)

# Function:    setupDataframe
#    Input:    Path to site photo directory.
#   Output:    Returns dataframe ready to be converted into Excel sheet.

def setupDataframe(siteDir):
    # Target folder
    dir = siteDir

    # Create dataframe file to serve as UI when converted into excel file
    df = pd.DataFrame(data=[])

    # Getting file paths
    filenames = []
    for item in os.listdir(dir):
        if item.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
            filenames.append(item)
        else:
            print("Not image")
            continue

    filepaths = []
    for item in os.listdir(dir):
        if item.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
            filepaths.append(dir + "\\" + item)
        else:
            print("Not image")
            continue
        
    # Debug
    print(filenames)
    print(filepaths)

    # Save file names into UI
    df["File Name"] = filenames

    # Columns
    setupColumns(df)

    return df, filenames, filepaths


# Function:    setupExcel
#    Input:    Filename for excel, Name of dataframe to convert.
#   Output:    Returns workbook and worksheet object

def setupExcel(filename, dataframe):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    dataframe.to_excel(writer, sheet_name='Sheet1')

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.freeze_panes(1, 0)                 # Freeze the first row.

    # Setting column widths
    cell_format1 = workbook.add_format()
    cell_format1.set_align('center')             # Text formats
    cell_format1.set_text_wrap()
    cell_format1.set_align('vcenter')
    cell_format1.set_border()                    # Border format
    cell_format1.set_bg_color("#f1e2dc")

    # Setting column widths
    cell_format2 = workbook.add_format()
    cell_format2.set_align('center')             # Text formats
    cell_format2.set_text_wrap()
    cell_format2.set_align('vcenter')
    cell_format2.set_border()                    # Border format
    cell_format2.set_bg_color("#C4DCF4")

    # Setting column widths
    cell_format3 = workbook.add_format()
    cell_format3.set_align('center')             # Text formats
    cell_format3.set_text_wrap()
    cell_format3.set_align('vcenter')
    cell_format3.set_border()                    # Border format
    cell_format3.set_bg_color('#FFAEBC')
    
    worksheet.set_column(first_col=0, last_col=0, width=15.6, cell_format=cell_format1)   # Index
    worksheet.set_column(first_col=1, last_col=1, width=9, cell_format=cell_format1)      # Filename
    worksheet.set_column(first_col=2, last_col=2, width=73, cell_format=cell_format1)     # Image 
    worksheet.set_column(first_col=3, last_col=3, width=30, cell_format=cell_format1)     # Raw
    worksheet.set_column(first_col=4, last_col=4, width=10, cell_format=cell_format1)     # Serial
    worksheet.set_column(first_col=5, last_col=6, width=15.6, cell_format=cell_format1)   # Brand
    worksheet.set_column(first_col=7, last_col=7, width=32, cell_format=cell_format1)     # Suggested Name
    worksheet.set_column(first_col=8, last_col=10, width=20, cell_format=cell_format2)    # Engineer Suggestions
    worksheet.set_column(first_col=11, last_col=11, width=20, cell_format=cell_format3)   # Engineer Name

    # Insert Button
    worksheet.insert_button('M2', {'macro':'rename_files','caption':'Rename Files', 'width': 80, 'height': 30})
    worksheet.set_column(first_col=12, last_col=12, width=15)    # Button width
    
    # Hide raw OCR results by default
    worksheet.set_column(0,0,options={'hidden': 1})
    worksheet.set_column(3,3,options={'hidden': 1})

    # Insert formula that will suggest new file name
    for row_num in range(2, 1000):
        worksheet.write_dynamic_array_formula("H%s:H%s" % (row_num, row_num), formula='=_xlfn.CONCAT(F%s," ",G%s," ",E%s," ","Nameplate")' % (row_num, row_num, row_num))

    # Insert formula that will suggest new file name
    for row_num in range(2, 1000):
        worksheet.write_dynamic_array_formula("L%s:L%s" % (row_num, row_num), formula='_xlfn.CONCAT(IF(J%s="",F%s,J%s)," ",IF(K%s="",G%s,K%s)," ",IF(I%s="",E%s,I%s)," ", "Nameplate.jpg")' % (row_num, row_num, row_num, row_num, row_num, row_num, row_num, row_num, row_num))

    return writer, workbook, worksheet

def iterateImages(sitedir, filenames, filepaths, excelfile, dataframe):
    # OCR Results
    results  = []

    # Create temp folder to store pre-processing of images
    try:
        os.mkdir("temp")
    except FileExistsError:
        print("File already exists")

    # Clean temp folder
    for files in os.listdir("temp"):
        os.remove("temp\\"+files)

    # Setup excel file
    writer, workbook, worksheet = setupExcel(excelfile, dataframe)

    for paths in filepaths:
  
        processed = processImage(sitedir + "\\" + filenames[filepaths.index(paths)])

        # Save processed image
        cv2.imwrite("temp\\"+ "processed_" +filenames[filepaths.index(paths)], processed)

        # Insert the processed image
        worksheet.set_row(row=filepaths.index(paths)+1, height=250)                                             
        worksheet.embed_image(filepaths.index(paths)+1, 2, "temp\\"+ "processed_" +filenames[filepaths.index(paths)])

        # OCR
        ocr_result = pytesseract.image_to_string("temp\\"+ "processed_" +filenames[filepaths.index(paths)])
        ocr_result = ocr_result.split("\n")

        if ocr_result == " " or ocr_result == "":
            worksheet.write(filepaths.index(paths)+1, 3, "unreadable")
            continue
        else:
            results.append(ocr_result)
            worksheet.write(filepaths.index(paths)+1, 3, str(ocr_result))

            print(filenames[filepaths.index(paths)])
            #print(ocr_result)

        # Filter OCR results for strings of interest
        for word in ocr_result:

            # Hussmann Serial Numbers
            if "SERIAL" in word:
                result = word[word.index("SERIAL")+10:17]
                worksheet.write(filepaths.index(paths)+1, 4, result)

            # RDC Brands
            if "ARN" in word:
                worksheet.write(filepaths.index(paths)+1, 5, "Arneg")

            if "HU" in word:
                worksheet.write(filepaths.index(paths)+1, 5, "Hussmann")
            
            # Model Number Test
            if "KG" in word:
                result = word[word.index("KG"):]
                worksheet.write(filepaths.index(paths)+1, 6, result)

            if "LIS" in word:
                result = word[word.index("LIS"):]
                worksheet.write(filepaths.index(paths)+1, 6, result)
            
            if "Lis" in word:
                result = word[word.index("Lis"):]
                worksheet.write(filepaths.index(paths)+1, 6, result)

            if "KM" in word:
                result = word[word.index("KM"):]
                worksheet.write(filepaths.index(paths)+1, 6, result)

    workbook.filename = sitedir + '/output.xlsm'
    workbook.add_vba_project('_internal/vbaProject.bin')

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()

    # Remove xlsx file
    os.remove(excelfile)

# Get target directory
dir_path = input("Paste the directory with the RDC photos: ")
dir_path = dir_path.replace('"','')

# Execute program
df, names, paths = setupDataframe(dir_path)
iterateImages(dir_path, names, paths, "output.xlsx", df)

# Let user know it is done
print("Program has finished running.")

# Let user read
input("Press Enter to Exit Program")