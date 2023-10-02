from docxtpl import DocxTemplate #import docxtpl from external library
from datetime import datetime
import pandas as pd

doc = DocxTemplate("en-template-manager-info.docx") #1. reading template

my_name = "Kyle Tran"
my_phone = "(702)-281-7901"
my_email = "kylet@gmail.com"
my_address = "123 Main Street"
today_date = datetime.today().strftime("%d %b, %Y") #gives format to date

my_context = {'my_name': my_name, 'my_phone': my_phone, 'my_email': my_email, 'my_address': my_address,
           'today_date': today_date} #indicate which elements to replace
#here the keys from the template need to match the variables above

df = pd.read_csv('en_fake_data.csv') #extracting datafram from external excel file

for index, row in df.iterrows():
    context = {
        'hiring_manager_name': row['name'],
        'address': row['address'],
        'phone_number': row['phone_number'],
        'email': row['email'],
        'job_position': row['job'],
        'company_name': row['company']
    }
    context.update(my_context)
    doc.render(context) #render/save in another file
    doc.save(f"generated_doc_{index}.docx")
