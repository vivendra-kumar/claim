# pip install --upgrade pip  
# pip install openpyxl     
# pip install pandas PyPDF2 reportlab
# pip install pdfrw PyMuPDF

import fitz
import pandas as pd

# Read the Excel file
excel_data = pd.read_excel('claim_from.xlsx')
value_dict = dict(zip(excel_data.Field, excel_data.Value))


def list_form_fields(pdf_path):
    doc = fitz.open(pdf_path)
    field_mapping = {}
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        fields = page.widgets()
        if fields:
            for field in fields:
                field_name = field.field_name
                field_type = field.field_type
                field_value = field.field_value
                rect = field.rect
                # print(f"Field Name: {field_name}, Type: {field_type}, Value: {field_value}, Rect: {rect}")
                # data[field_name] = field_name
                field_mapping[field_name] = rect
    return field_mapping

def fill_form(pdf_path, output_path, data):
    doc = fitz.open(pdf_path)
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        fields = page.widgets()
        if fields:
            for field in fields:
                field_name = field.field_name
                if field_name in data:
                    field.field_value = data[field_name]
                    field.update()
    doc.save(output_path)

pdf_path = "cashless-claim-form.pdf"
output_path = "filled-cashless-claim-form.pdf"

# Create a mapping of the generic field names to the actual data you want to input
data = {'Text9': value_dict['Name of the hospital'],
 'Text26': value_dict['Hospital location'],
 'Text28': 'Hospital ID',
 'Text27': 'Hospital email ID',
 'Text30': 'ROHINI ID',
 'Text31': 'Name of the patient',
 'Text32': 'Contact no',
 'Text33': 'Alternate contact no.',
 'Text34': 'DOB',
 'Text35': 'Insurer ID card no',
 'Text36': 'Policy number/Name of corporate',
 'Text37': 'Employee ID',
 'Text38': 'Give details:',
 'Text39': 'Do you have a family physician, if yes: Name',
 'Text40': 'Contact no.:',
 'Text41': 'Occupation of insured patient',
 'Text42': 'Address of insured patient:',
 'Text43': 'Name of Illness/disease with presenting complaints:',
 'Text44': 'Relevant clinical findings:',
 'Text45': 'Duration of the present ailment:',
 'Text46': 'Date of first consultation:',
 'Text47': 'Past history of present ailment if any:',
 'Text49': 'Provisional diagnosis:',
 'Text50': 'Route of drug administration:',
 'Text51': 'If investigation and/or medical management, provide details:',
 'Text52': 'If Surgical, name of surgery:',
 'Text54': 'ICD 10 PCS code:',
 'Text55': 'How did injury occur:',
 'Text56': 'If other treatments provide details:',
 'Text57': 'Room type:',
 'Text58': 'In case of maternity: g',
 'Text59': 'In case of maternity:  p',
 'Text60': 'In case of maternity:  L',
 'Text61': 'In case of maternity:  A',
 'Text62': 'Expected no. of days stay in hospital:',
 'Text66': 'Days in ICU:',
 'Text99': 'ICD 10 code:',
 'Text100': 'contact no. doctor',
 'Text101': 'DOB',
 'Text102': 'Insurer name:',
 'Text103': 'Text103',
 'Text104': 'Text104',
 'Text105': 'Date of injury:',
 'Text106': 'FIR no.:',
 'Text107': 'Expected date of delivery:',
 'Text108': 'Date of admission:',
 'Text109': 'Time of admission:',
 'Button110': 'Button110',
 'Button111': 'Button111',
 'Button112': 'Button112',
 'Button113': 'Button113',
 'Button114': 'Button114',
 'Button115': 'Button115',
 'Button116': 'Button116',
 'Button118': 'Button118',
 'Button119': 'Button119',
 'Button120': 'Button120',
 'Button121': 'Button121',
 'Button122': 'Button122',
 'Button123': 'Button123',
 'Button124': 'Button124',
 'Button125': 'Button125',
 'Button126': 'Button126',
 'Button127': 'Button127',
 'Button128': 'Button128',
 'Button129': 'Button129',
 'Button130': 'Button130',
 'Button131': 'Button131',
 'Button132': 'Button132',
 'Button133': 'Button133',
 'Text67': 'Per Day Room Rent + Nursing & Service charges + Patient’s D',
 'Text68': 'Expected cost for investigation + diagnostics:',
 'Text69': 'ICU Charges:',
 'Text70': 'OT Charges',
 'Text71': 'Text71',
 'Text72': 'Text72',
 'Text73': 'Text73',
 'Text74': 'Text74',
 'Text75': 'Text75',
 'Text76': 'Text76',
 'Text77': 'Text77',
 'Text78': 'Text78',
 'Text79': 'Text79',
 'Text80': 'Text80',
 'Text81': 'Text81',
 'Text82': 'Text82',
 'Text83': 'Text83',
 'Text84': 'Text84',
 'Text85': 'Text85',
 'Text86': 'Text86',
 'Text87': 'Text87',
 'Text88': 'Text88',
 'Text89': 'Text89',
 'Text90': 'Text90',
 'Text91': 'Text91',
 'Text92': 'Text92',
 'Text93': 'Text93',
 'Text94': 'Text94',
 'Text95': 'Text95',
 'Text96': 'Text96',
 'Text97': 'Text97',
 'Text98': 'Text98',
 'Button134': 'Button134',
 'Button135': 'Button135',
 'Button137': 'Button137',
 'Button138': 'Button138',
 'Button139': 'Button139',
 'Button140': 'Button140',
 'Button141': 'Button141',
 'Button142': 'Button142',
 'Button143': 'Button143'}

# List the form fields to help create the mapping
field_mapping = list_form_fields(pdf_path)

# Fill the form using the mapping
fill_form(pdf_path, output_path, data)
