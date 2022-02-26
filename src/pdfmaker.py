import csv
import json
from fpdf import FPDF

#call: python3 pdfmaker.py


# Function to convert a CSV to JSON
# Takes the file paths as arguments
# May become obsolete if we read live data from google sheet
def make_json(csvFilePath, jsonFilePath):
     
    # will hold json data
    data = {}
     
    # Open a csv reader called DictReader
    with open(csvFilePath, encoding='utf-8') as csvf:
        csvReader = csv.DictReader(csvf)
         
        # Convert each row into a dictionary
        # and add it to data
        for rows in csvReader:
             
            # Assuming a column named 'Discord' to
            # be the primary key
            key = rows['Discord']
            data[key] = rows
 
    # Open a json writer, and use the json.dumps()
    # function to dump data
    with open(jsonFilePath, 'w', encoding='utf-8') as jsonf:
        jsonf.write(json.dumps(data, indent=4))

    # pass json data to pdf maker
    make_pdf(data)


# makes a pdf to fit on standard sized page
def make_pdf(json_data):

    #REQUEST TEMPLATE THAT WE ARE TRYING TO CREATE
    #https://docs.google.com/document/d/1elTQRFnLgp2ktDsm021bli2rcBwSgzgwS4W-XwMqR44/edit?usp=sharing
    for entry in json_data:

        #all json members that will be on the service hours sheet
        #DOES NOT INCLUDE EVERY MEMBER OF THE JSON OBJECT
        info_categories = ["Timestamp", "Name", "Email", "Address", "Phone", "Service Hours", 
         "Responsibilities", "School Name", "School Address", "School Phone Number"]

        #makes a dictionary with empty values for all categories
        data = dict.fromkeys(info_categories)

        #sets values in dictionary for current entry
        for category in info_categories:
            data[category] = json_data[entry][category]
        
        #Format values onto pdf to recreate CS template
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(40, 10, data['Name'])
        pdf.output('../output/' + data['Name'].replace(" ", "") + '.pdf', 'F')



# Driver Code
 
# Decide the two file paths 
csvFilePath = '../input/ServiceHoursTestInput.csv'
jsonFilePath = '../output/ServiceHoursTestInput.json'
 
# Call the make_json function, which calls the make_pdf function
make_json(csvFilePath, jsonFilePath)

