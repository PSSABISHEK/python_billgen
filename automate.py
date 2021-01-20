from mailmerge import MailMerge
from datetime import date
import os
import pandas as pd


read_csv = pd.read_csv('./temp_V2.csv').T.to_dict()
fName = str(date.today())
fName = fName.replace('-', '')
if not os.path.exists(fName):
    os.makedirs(fName)

for i in read_csv:
    template = 'ReceiptTemplate.docx'
    document = MailMerge(template)
    docs_headers = document.get_merge_fields()
    data = read_csv[i]
    document.merge(
        amountPaid=str(data['amountPaid']),
        cAddress=data['cAddress'],
        cName=data['cName'],
        details=data['details'],
        duration=data['duration'],
        eMailID=data['eMailID'],
        noteAmount=str(data['amountPaid']),
        phNumber=str(data['phNumber']),
        subEndDate=str(data['subEndDate']),
        subStartDate=str(data['subStartDate'])
    )
    document.write('./'+fName+'/'+data['PDF']+'.docx')