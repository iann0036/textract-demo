import boto3
import requests
import io
import time
import json
import pickle
import pprint
import xlwt

BUCKET_NAME = "<REPLACE ME>"
KEY_NAME = "asicreport.pdf"
OUTPUT_XLS = "output.xls"

##### STEP 1. Submit document to Textract

textract_client = boto3.client('textract', region_name='us-west-2')

try:
    analysis_result = pickle.load(open("save.p","rb"))
except FileNotFoundError:
    print("About to send document to Textract")
    time.sleep(10)

    s3_client = boto3.client('s3', region_name='us-west-2')
    with open(KEY_NAME, "rb") as f:
        s3_client.put_object(Body=f.read(), Bucket=BUCKET_NAME, Key=KEY_NAME, ACL='public-read')

    analysis_job = textract_client.start_document_analysis(
        DocumentLocation={
            'S3Object': {
                'Bucket': BUCKET_NAME,
                'Name': KEY_NAME
            }
        },
        FeatureTypes=[
            'TABLES'
        ]
    )

    analysis_result = textract_client.get_document_analysis(
        JobId=analysis_job['JobId']
    )

    while analysis_result['JobStatus'] == "IN_PROGRESS":
        time.sleep(10)
        analysis_result = textract_client.get_document_analysis(
            JobId=analysis_job['JobId']
        )

    while 'NextToken' in analysis_result:
        tmp_result = textract_client.get_document_analysis(
            JobId=analysis_job['JobId'],
            NextToken=analysis_result['NextToken']
        )
        tmp_result['Blocks'].extend(analysis_result['Blocks'])
        analysis_result = tmp_result

    pickle.dump(analysis_result, open("save.p", "wb"))

##### STEP 2. Extract table information into Excel workbook

book = xlwt.Workbook()

table_number = 1

for block in analysis_result['Blocks']: # document tables
    if block['BlockType'] == 'TABLE':

        sheet = book.add_sheet("Table {}".format(table_number))
        table_number+=1
        if table_number > 20:
            break
        table = {}

        for relationship in block['Relationships']:
            if relationship['Type'] == 'CHILD':
                for rid in relationship['Ids']:
                    for block2 in analysis_result['Blocks']: # table cells
                        if block2['Id'] == rid:
                            if 'Relationships' in block2:
                                for relationship2 in block2['Relationships']:
                                    if relationship2['Type'] == 'CHILD':
                                        for rid2 in relationship2['Ids']:
                                            for block3 in analysis_result['Blocks']: # cell words
                                                if block3['Id'] == rid2:
                                                    row = block2['RowIndex'] - 1
                                                    column = block2['ColumnIndex'] - 1
                                                    text = block3['Text']
                                                    
                                                    if row not in table:
                                                        table[row] = {}
                                                    
                                                    if column in table[row]:
                                                        table[row][column] += " " + text
                                                    else:
                                                        table[row][column] = text

        for rowindex, row in table.items():
            for columnindex, text in row.items():
                sheet.write(rowindex, columnindex, text)

book.save(OUTPUT_XLS)
