import datetime
import logging
import os
#import csv
#import json
from fpdf import FPDF
import phonenumbers
import tempfile

#gsheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#email
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, From, Attachment, FileContent, FileName, FileType, Disposition, ContentId
import base64

import azure.functions as func

TEMPLATE_ID = os.environ.get("TEMPLATE_ID")
FROM_EMAIL = From('servicehours@schoolsimplified.org', "School Simplified Service Hours")

def create_keyfile_dict():
    keys = {
        "type": os.environ.get('GOOG_TYPE'),
        "project_id": os.environ.get('GOOG_PROJECT_ID'),
        "private_key_id": os.environ.get('GOOG_PRIVATE_KEY_ID'),
        "private_key": os.environ.get('GOOG_PRIVATE_KEY'),
        "client_email": os.environ.get('GOOG_CLIENT_EMAIL'),
        "client_id": os.environ.get('GOOG_CLIENT_ID'),
        "auth_uri": os.environ.get('GOOG_AUTH_URI'),
        "token_uri": os.environ.get('GOOG_TOKEN_URI'),
        "auth_provider_x509_cert_url": os.environ.get('GOOG_AUTH_PROVIDER_X509_CERT_URL'),
        "client_x509_cert_url": os.environ.get('GOOG_CLIENT_X509_CERT_URL')
    }
    return keys

# makes a pdf to fit on standard sized page
def make_pdf():
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(create_keyfile_dict(), scope)
    client = gspread.authorize(creds)   #NEED TO CHANGE TO ENV VARIABLE
    sheet = client.open("Community Service Hours Request (Responses)").sheet1  # Open the spreadsheet
    sheet_data = sheet.get_all_records() #Get all records


    #REQUEST TEMPLATE THAT WE ARE TRYING TO CREATE
    #https://docs.google.com/document/d/1elTQRFnLgp2ktDsm021bli2rcBwSgzgwS4W-XwMqR44/edit?usp=sharing
    #only process 100 (50, 2 times a day) requests to keep sendgrid free
    emails_sent = 0
    for entry in sheet_data:

        #only process 100 (50, 2 times a day) requests to keep sendgrid free
        if emails_sent == 50:
            break

        #check if entry is validated by manager
        worksheet = client.open("Community Service Hours Request (Responses)").worksheet(entry["Team"])
        timestamp_cell = worksheet.find(entry["Timestamp"])
        validation_cell = worksheet.cell(timestamp_cell.row, 1)

        #if not validated, skip
        if validation_cell.value == "FALSE":
            continue
    
        #all json members that will be on the service hours sheet
        #DOES NOT INCLUDE EVERY MEMBER OF THE JSON OBJECT
        info_categories = ["Timestamp", "Name", "Email", "Address", "Phone", "Service Hours", 
        "Responsibilities", "School Name", "School Address", "School Phone Number"]

        #makes a dictionary with empty values for all categories
        data = dict.fromkeys(info_categories)

        #sets values in dictionary for current entry
        for category in info_categories:
            data[category] = entry[category]
        
        #check if valid row
        if not data["Name"] or not data["Email"]:
            continue
            
        #Format values onto pdf to recreate CS template
        file_name = "" + data['Name'].replace(" ", "") + 'ServiceHours.pdf'
        
        temp_file_path = tempfile.gettempdir()

        format_pdf(data, file_name, temp_file_path)

        #send pdf back to user
        message = Mail(
            from_email= FROM_EMAIL,
            to_emails=data['Email'])

        first_name = data['Name'].split()[0].capitalize()
        message.dynamic_template_data = {
            'name': first_name
        }

        message.template_id = TEMPLATE_ID

        with open(temp_file_path + "/" + file_name, 'rb') as f:
            pdf_data = f.read()
        encoded = base64.b64encode(pdf_data).decode()
        os.remove(temp_file_path + "/" + file_name)
        #encoded = base64.b64encode(pdf.encode("latin-1")).decode("latin-1")
        
        attachment = Attachment()
        attachment.file_content = FileContent(encoded)
        attachment.file_type = FileType("application/pdf")
        attachment.file_name = FileName(file_name)
        attachment.disposition = Disposition("attachment")
        attachment.content_id = ContentId("PDF Document file")

        message.add_attachment(attachment)

        try:
            sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
            response = sg.send(message)
            print(response.status_code)
            print(response.body)
            print(response.headers)
        except Exception as e:
            print(e)
        
        emails_sent += 1

        #delete record
        main_timestamp_cell = sheet.find(entry["Timestamp"])
        sheet.delete_rows(main_timestamp_cell.row)
        
    
    #Delete 100 records after processing
    #num_records = getattr(sheet, "row_count")

    #if there is only one record, cant delete that row so just clear it
    #if num_records == 2:
    #    sheet.batch_clear(['A2:M2'])
    #more than one record
    #else:
    #    start_index = 2
    #    end_index = min(100, num_records) #min of either 100 or number of records
    #    sheet.add_rows(1) #prevents deleting frozen row error
    #    sheet.delete_rows(start_index, end_index)
    #    #sheet.batch_clear(['A2:M100'])

def format_pdf(data, file_name, temp_file_path):
    pdf = FPDF()
    pdf.add_page()

    #page constants
    PDF_CENTER_WIDTH = pdf.w * 0.5
    PDF_CENTER_HEIGHT = pdf.h * 0.5
    SIDE_MARGIN = 25

    pdf.add_font('Montserrat-Bold', '', './PDFMaker/Fonts/Montserrat-Bold.ttf', uni=True)
    pdf.add_font('Montserrat-Regular', '', './PDFMaker/Fonts/Montserrat-Regular.ttf', uni=True)
    #puts background image in center of page; gets half of width of page and subtracts half of width of image
    pdf.image('./PDFMaker/Images/SchoolSimpBG.png', PDF_CENTER_WIDTH - 100, PDF_CENTER_HEIGHT - 100, 200, 200)

    #header
    pdf.image('./PDFMaker/Images/BlueHeader.png', 0, 0, pdf.w, 15)
    pdf.set_font('Montserrat-Regular', '', 14)
    pdf.set_text_color(108,124,251)
    pdf.cell(0, 15, "", 0, 1)
    pdf.image('./PDFMaker/Images/SchoolSimpLogoBlue.png', pdf.w - 38, 19, 15, 15)
    pdf.cell(163, 10, 'School Simplified, Inc.', 0, 1, 'R')
    pdf.set_text_color(0,0,0)
    pdf.cell(0, 3, "", 0, 1)

    #title
    pdf.set_font('Montserrat-Bold', '', 14)
    pdf.cell(0, 10, 'Community Service Verification Certificate', 0, 1, 'C')
    
    #student name and email
    pdf.set_left_margin(SIDE_MARGIN)
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(0, 15, 'This form is to certify that the following student:', 0, 1)
    pdf.cell(PDF_CENTER_WIDTH - SIDE_MARGIN, 7, 'Student Name:', 0, 0)  #name
    pdf.cell(0, 7, 'Student Email:', 0, 1)                     #email
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(PDF_CENTER_WIDTH - SIDE_MARGIN, 5, str(data['Name']), 0, 0)
    pdf.cell(0, 5, str(data["Email"]), 0, 1)
    pdf.line(SIDE_MARGIN, 75, PDF_CENTER_WIDTH - 10, 75)
    pdf.line(PDF_CENTER_WIDTH, 75, pdf.w - SIDE_MARGIN, 75)
    pdf.cell(0, 2, "", 0, 1)

    #student address and phone
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(PDF_CENTER_WIDTH - SIDE_MARGIN, 7, 'Student Address:', 0, 0) #address
    pdf.cell(0, 7, 'Student/Parent Phone Number:', 0, 1)         #phone
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(PDF_CENTER_WIDTH - SIDE_MARGIN, 5, str(data['Address']), 0, 0)
    phone = phonenumbers.format_number(phonenumbers.parse(str(data["Phone"]), 'US'), phonenumbers.PhoneNumberFormat.NATIONAL)
    pdf.cell(0, 5, phone, 0, 1)
    pdf.line(SIDE_MARGIN, 89, PDF_CENTER_WIDTH - 10, 89)
    pdf.line(PDF_CENTER_WIDTH, 89, pdf.w - SIDE_MARGIN, 89)
    pdf.cell(0, 3, "", 0, 1)

    #student service hours
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(21, 10, 'Completed ', 0, 0)                    #service hours
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, str(data["Service Hours"]), 0, 0, 'C') 
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(100, 10, ' unpaid hours of service on the following date(s):', 0, 0)
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, str(data["Timestamp"]).split(' ', 1)[0], 0, 1, 'C')
    pdf.line(46, 99, 56, 99)
    pdf.line(140, 99, pdf.w - SIDE_MARGIN, 99)
    pdf.cell(0, 2, "", 0, 1)

    #student responsibilities
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(0, 10, 'The volunteer\'s service consisted of the following responsibilities:', 0, 1) 
    x = pdf.get_x() #gets position before multicell
    y = pdf.get_y()
    pdf.set_font('Montserrat-Regular', '', 9) 
    pdf.multi_cell(pdf.w - 50, 7, str(data["Responsibilities"]), 0, 'L')
    pdf.set_x(x)
    pdf.set_y(y)
    pdf.cell(pdf.w - 50, 42, "", 1, 1)
    pdf.cell(0, 4, "", 0, 1)

    #school name, address, and phone
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(30, 10, 'School Name:', 0, 0)                    
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, str(data["School Name"]), 0, 1) 

    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(30, 10, 'School Address:', 0, 0)                    
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, str(data["School Address"]), 0, 1) 

    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(30, 10, 'School Phone:', 0, 0)                    
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, str(data["School Phone Number"]), 0, 1) 

    pdf.line(55, 167, pdf.w - SIDE_MARGIN, 167)
    pdf.line(55, 177, pdf.w - SIDE_MARGIN, 177)
    pdf.line(55, 187, pdf.w - SIDE_MARGIN, 187)
    pdf.cell(0, 1, "", 0, 1)

    #admin verification
    #Hardcoded for now
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(0, 15, 'This is verified by the following administrator, a current board member of the organization:', 0, 1) 
    pdf.cell(45, 10, 'Administrator:', 0, 0)                    
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, "Param Patil", 0, 1) 

    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(45, 10, 'Administrator Signature:', 0, 1)         
    pdf.image('./PDFMaker/Images/ParamPatilSign.png', 70, 218, 35, 5)           

    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(45, 10, 'Verification Phone:', 0, 0)                    
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, "(518) 886-2954", 0, 1) 

    current_date = datetime.date.today()
    formatted_date = datetime.date.strftime(current_date, "%m/%d/%Y")
    pdf.set_font('Montserrat-Bold', '', 9)
    pdf.cell(45, 10, 'Date:', 0, 0)                    
    pdf.set_font('Montserrat-Regular', '', 9)
    pdf.cell(10, 10, formatted_date, 0, 1)

    pdf.line(70, 213, pdf.w - SIDE_MARGIN, 213)
    pdf.line(70, 223, pdf.w - SIDE_MARGIN, 223)
    pdf.line(70, 233, pdf.w - SIDE_MARGIN, 233)
    pdf.line(70, 243, pdf.w - SIDE_MARGIN, 243)

    #footer
    pdf.cell(0, 20, "", 0, 1)
    pdf.set_font('Montserrat-Regular', '', 8)
    pdf.cell(pdf.w - (SIDE_MARGIN * 2) + 1, 3, "School Simplified, Inc.", 0, 1, 'R')
    pdf.cell(pdf.w - (SIDE_MARGIN * 2) + 1, 3, "8 The Green, Dover DE 19901", 0, 1, 'R')
    pdf.cell(pdf.w - (SIDE_MARGIN * 2) + 1, 3, "support@schoolsimplified.org", 0, 1, 'R')

    #complete_pdf = pdf.output('S')
    #return complete_pdf
    pdf.output(temp_file_path + "/" + file_name, 'F')


def main(makePDF: func.TimerRequest) -> None:
    utc_timestamp = datetime.datetime.utcnow().replace(
        tzinfo=datetime.timezone.utc).isoformat()

    if makePDF.past_due:
        logging.info('The timer is past due!')

    make_pdf()
    logging.info('Python timer trigger function ran at %s', utc_timestamp)