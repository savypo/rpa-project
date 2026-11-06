import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE

while True:    
    directory = single_input("Please input the directory where the ProductData.xlsx, MasterInvoice.pdf, and Order files are located")
    directory = directory.strip(' " ').strip(" ' ")
    if directory[-1] != ("\\"):
        directory = directory + "\\"
    if (os.path.exists(directory + "ProductData.xlsx") and os.path.exists(directory + "MasterInvoice.pdf") and os.path.exists(directory + "Order1.xlsx") and os.path.exists(directory + "Order2.xlsx") and os.path.exists(directory +"Order3.xlsx") and os.path.exists(directory +"Order4.xlsx")):
        break
    else:
        show_message("Error", "Please select a different directory.")
print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
descriptions = pd.read_excel(directory+"ProductData.xlsx", sheet_name= "Descriptions")
for index, row in descriptions.iterrows():
    description = row.Description
    unitnumber = row['Unit number']
    product_descriptions[unitnumber] = description

prices = pd.read_excel(directory+"ProductData.xlsx",sheet_name= "Prices")
for index, row in prices.iterrows():
    price_per_unit = row['Price per unit']
    unitnumber = row['Unit number']
    product_prices[unitnumber] = price_per_unit


product_descriptions
prices
product_prices

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num,filename,directory):
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    CustomerEmail = str(selected_sheet['C13'].value)
    update_dict = {}
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = selected_sheet['C9'].value.strftime("%m/%d/%Y")
    total = 0 
    if selected_sheet['B24'].value is not None:
        amount1 = (selected_sheet['C24'].value *product_prices[selected_sheet['B24'].value])
        total = total + amount1
        update_dict['Price0'] = f"${product_prices[selected_sheet['B24'].value]:,.2f}"
        update_dict['Quantity0'] = str(selected_sheet['C24'].value)
        update_dict['Amount0'] = f'${amount1:.2f}'
        update_dict['Description0'] = str(product_descriptions[selected_sheet['B24'].value])
    if selected_sheet['B25'].value is not None:
        amount2 = (selected_sheet['C25'].value *product_prices[selected_sheet['B25'].value])
        total = total + amount2
        update_dict['Price1'] = f"${product_prices[selected_sheet['B25'].value]:,.2f}"
        update_dict['Quantity1'] = str(selected_sheet['C25'].value)
        update_dict['Amount1'] = f'${amount2:.2f}'
        update_dict['Description1'] = str(product_descriptions[selected_sheet['B25'].value])
    if selected_sheet['B26'].value is not None:
        amount3 = (selected_sheet['C26'].value *product_prices[selected_sheet['B26'].value])
        total = total + amount3
        update_dict['Price2'] = f"${product_prices[selected_sheet['B26'].value]:,.2f}"
        update_dict['Quantity2'] = str(selected_sheet['C26'].value)
        update_dict['Amount2'] = f'${amount3:.2f}'
        update_dict['Description2'] = str(product_descriptions[selected_sheet['B26'].value])
    if selected_sheet['B27'].value is not None:
        amount4 = (selected_sheet['C27'].value *product_prices[selected_sheet['B27'].value])
        total = total + amount4
        update_dict['Price3'] = f"${product_prices[selected_sheet['B27'].value]:,.2f}"
        update_dict['Quantity3'] = str(selected_sheet['C27'].value)
        update_dict['Amount3'] = f'${amount4:.2f}'
        update_dict['Description3'] = str(product_descriptions[selected_sheet['B27'].value])
    if selected_sheet['B28'].value is not None:
        amount5 = (selected_sheet['C28'].value *product_prices[selected_sheet['B28'].value])
        update_dict['Price4'] = f"${product_prices[selected_sheet['B28'].value]:,.2f}"
        update_dict['Quantity4'] = str(selected_sheet['C28'].value)
        update_dict['Amount4'] = f'${amount5:.2f}'
        update_dict['Description4'] = str(product_descriptions[selected_sheet['B28'].value])
        total = total + amount5
    update_dict['Total'] = f'${total:.2f}'
    update_dict['CustomerAddress1'] = str(selected_sheet['C16'].value)
    update_dict['CustomerName'] = str(selected_sheet['C14'].value)
    update_dict['CustomerCompany'] = str(selected_sheet['C15'].value)
    update_dict['CustomerAddress2'] = str(selected_sheet['C17'].value) + ", " + str(selected_sheet['C18'].value) + " " + str(selected_sheet['C19'].value)



    return update_dict, update_dict['CustomerName'], CustomerEmail, update_dict['Date']

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx',directory)

order_data

email

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)
create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE

gmail_password_path= select_file("Plase select the gmail password file")
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    gmail_user = "savyjbrouwer@gmail.com"

    recipients = email

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    path_to_file = directory+ f"Invoice {invoice_num}.pdf"
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, "Thomas Holt", "savyjbrouwer@gmail.com", "11/13/2023", directory )

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames
invoice_num = 8515
for filename in filenames:
    if "Order" in filename and "xlsx" in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1



