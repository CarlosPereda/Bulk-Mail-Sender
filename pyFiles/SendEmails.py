# ======================================================================================================================
# Author:           CarlosPereda
# Repository:       https://github.com/CarlosPereda/Bulk-Mail-Sender
# Execution-Notes:  Run this program whit GUI_SendChunkEmail (prefered) or from CMD
# Tested with:      3.9.11
# Tested on:        Win 10 Home (x64)
# License:          Feel free to modify this software and in case of publishing mention the author.
# Credits:          This program was inspired by the tutorials of Izzy Analytics: 
                    # https://www.youtube.com/watch?v=J3SiyMingRk&list=PLHnSLOMOPT11njaNmENJN6p2ro9MTc7t_
# Notes:            To successfully run this program you must:
                    # Have the outlook application installed and closed
                    # Have a txt email template
                    # Have a Database of all recipients in an excel
# ======================================================================================================================
# This software is provided 'as-is', without any express or implied warranty.
# In no event will the author be held liable for any damages arising from the use of this software.
# Please consider the law of your region related to spam and avoid missuse of the following program. 
# ======================================================================================================================

from datetime import datetime
from dataclasses import dataclass, field
from time import sleep
from openpyxl import load_workbook
import win32com.client as client # Install pypiwin32 if win32com does not work
import sys
import re

class EmailAddressNotFoundError(Exception):
    def __init__(self, message):
        self.message = message

@dataclass
class SendBulkMail:
    mail_from: str
    mail_cc: str
    mail_subject: str

    def __set_sender(self, message, send_on_behalf=True):
        """Set sendOfBehalfOfName bahaviour"""
        if send_on_behalf:
            message.SentOnBehalfOfName = self.mail_from
            return
        
        SENDER = None
        for email_address in self.outlook.Session.Accounts:
            if self.mail_from in str(email_address):
                SENDER = email_address
                break
        
        if SENDER is None:
            raise EmailAddressNotFoundError(f"Sender address \"{self.mail_from}\" could not be found \
                                            in outlook.Session.Accounts")

        message._oleobj_.Invoke(*(64209, 0, 8, 0, SENDER))

    
    def __set_cc(self, message, cc):
        """Sets and formats the cc. """
        if cc:
            cc = re.sub(' ', '', cc)
            cc = re.sub(',', ';', cc)
            message.cc = cc
        else:
            message.cc = ""


    def load_files(self, html_path, excel_path, sheet_name):
        """Loads and sets the html and excel files."""
        self.wb = load_workbook(excel_path)
        self.sheet = self.wb[sheet_name]
        self.HTML_TXT = open(html_path, 'r', encoding="utf8")
        self.HTML_TXT = self.HTML_TXT.read()


    def get_student_rows(self, headers_row, first_data_row, last_data_row, columns=None):
        """Make a dictionary for each student in the excel file"""
        headers_row = int(headers_row)
        first_data_row = int(first_data_row)
        last_data_row = int(last_data_row)
        
        if columns is None:
            columns = [chr(x) for x in range(65, 91)]

        self.students_list = []

        for row in range(first_data_row, last_data_row+1):
            student_data = {}
            for col in columns:
                header_cell = col + str(headers_row)
                header_value = self.sheet[header_cell].value
                if header_value == None:
                    break

                data_cell = col + str(row)
                data_value = self.sheet[data_cell].value
                student_data[header_value] = data_value
            
            if 'CC' not in student_data.keys():
                student_data['CC'] = self.mail_cc

            self.students_list.append(student_data)
            

        for dynamic_key in {'Email Address', 'Attachments'}:
            if dynamic_key not in self.students_list[0].keys():
                raise Exception(f"The header '{dynamic_key}' was not found in the spreadsheet headers")


    def set_message_parameters(self, message, DynamicContent):
        self.__set_sender(message)
        self.__set_cc(message, DynamicContent['CC'])
        message.To = DynamicContent['Email Address']
        message.Subject = self.mail_subject 
        message.HTMLBody = eval(self.HTML_TXT)

        if DynamicContent['Attachments'] != None:
            for attachment in DynamicContent['Attachments'].split(";"): # TODO: erase spaces between attachments with regex
                message.Attachments.Add(attachment)


    def send_mails(self):
        try:
            outlook = client.GetActiveObject('Outlook.Application')
        except:
            outlook = client.Dispatch('Outlook.Application')
        
        print("Sending mails... Please wait")
        for DynamicContent in self.students_list:
            if DynamicContent["Email Address"] == None:
                print("A row with no data was attempted to be read")
                break
            message = outlook.CreateItem(0)
            self.set_message_parameters(message, DynamicContent)
            message.Send()
            print(f"The email was sent to {DynamicContent['Email Address']} succesfully!")
            sleep(2)


if __name__ == "__main__":
    new_message = SendBulkMail(mail_from=sys.argv[1],
                        mail_cc=sys.argv[2],
                        mail_subject=sys.argv[3])

    new_message.load_files(excel_path= sys.argv[4],
                                html_path = sys.argv[5],
                                sheet_name= sys.argv[6])
    
    new_message.get_student_rows(headers_row=sys.argv[7], 
                                    first_data_row=sys.argv[8],
                                    last_data_row=sys.argv[9])

    new_message.send_mails()
    print("Program ended")