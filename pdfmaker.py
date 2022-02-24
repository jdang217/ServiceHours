import csv
import json
from fpdf import FPDF

#call: python3 pdfmaker.py


# Function to convert a CSV to JSON
# Takes the file paths as arguments
def make_json(csvFilePath, jsonFilePath):
     
    # create a dictionary
    data = {}
     
    # Open a csv reader called DictReader
    with open(csvFilePath, encoding='utf-8') as csvf:
        csvReader = csv.DictReader(csvf)
         
        # Convert each row into a dictionary
        # and add it to data
        for rows in csvReader:
             
            # Assuming a column named 'No' to
            # be the primary key
            key = rows['Discord']
            data[key] = rows
 
    # Open a json writer, and use the json.dumps()
    # function to dump data
    with open(jsonFilePath, 'w', encoding='utf-8') as jsonf:
        jsonf.write(json.dumps(data, indent=4))

def make_pdf(input):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(40, 10, input)
    pdf.output('tuto1.pdf', 'F')



# Driver Code
 
# Decide the two file paths according to your
# computer system
csvFilePath = r'ServiceHoursTestInput.csv'
jsonFilePath = r'ServiceHoursTestInput.json'
 
# Call the make_json function
make_json(csvFilePath, jsonFilePath)
make_pdf("hello world")