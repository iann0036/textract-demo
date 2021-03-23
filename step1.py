import boto3
import io
import time
import json

BUCKET_NAME = "my-textract-demo-bucket"
KEY_NAME = "contactlist.pdf"

##### STEP 1. Submit document to Textract

s3_client = boto3.client('s3', region_name='us-east-1')

with open(KEY_NAME, "rb") as f:
    s3_client.put_object(Body=f.read(), Bucket=BUCKET_NAME, Key=KEY_NAME)

textract_client = boto3.client('textract', region_name='us-east-1')

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
    time.sleep(5)
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

with open("analysisresult.json", "w") as f:
    f.write(json.dumps(analysis_result))
