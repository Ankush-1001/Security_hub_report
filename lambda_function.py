import boto3
import datetime
import os
import xlsxwriter
from ses_sender import *


def lambda_handler(event, context):
    creating_xlsx()


def creating_xlsx():

    securityhub_client = boto3.client('securityhub' )

    #### List to store all the Security Findings
    all_findings = []

    # ### Retrieve all the Security Findings that match the specified filters using pagination
    paginator = securityhub_client.get_paginator('get_findings')
    response_iterator = paginator.paginate(Filters={
        'RecordState': [{'Value': 'ACTIVE', 'Comparison': 'EQUALS'}],
        'AwsAccountId': [{'Value': '573989815911', 'Comparison': 'EQUALS'}],
        'WorkflowStatus': [
            {'Value': 'NEW', 'Comparison': 'EQUALS'},
            {'Value': 'NOTIFIED', 'Comparison': 'EQUALS'}
        ],
        'FindingProviderFieldsSeverityLabel': [
            {'Value': 'CRITICAL','Comparison': 'EQUALS'},
            {'Value': 'HIGH','Comparison': 'EQUALS'},
            {'Value': 'MEDIUM','Comparison': 'EQUALS'}

        ],
    })

    for response in response_iterator:
        all_findings.extend(response['Findings'])
        sorted_findings = sorted(all_findings, key=lambda x: x['Severity']['Normalized'], reverse=True)
        # print(sorted_findings[8])
        ### Create an XLSX file with the sorted Security Findings in the /tmp directory
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
        filename = f'security_findings_{timestamp}.xlsx'
        xlsx_file = f'/tmp/{filename}'

        ### Create a workbook and add a worksheet
        workbook = xlsxwriter.Workbook(xlsx_file)
        worksheet = workbook.add_worksheet('Security_hub')

        ### Adding formats
        worksheet.set_default_row(30) # Sets the default row height to 30 pixels
        worksheet.set_column(0, 0, 15) # 1 column width. Column used for Severity.
        worksheet.set_column(1, 1, 75) # 2 column width. Column used for Title.
        worksheet.set_column(2, 2, 15) # 2 column width. Column used for WorkFlow.
        worksheet.set_column(3, 3, 15) # 3 column width. Column used for Account ID.
        worksheet.set_column(4, 4, 140) # 4 column width. Column used for Description.
        worksheet.set_column(5, 5, 25) # 5 column width. Column used for Updated at.
        worksheet.set_column(6, 6, 25) # 5 column width. Column used for Updated at.
        worksheet.set_column(7, 7, 20) # 6 column width. Column used for Resource Type.
        worksheet.set_column(8, 8, 125) # 7 column width. Column used for Resource ID.
        worksheet.set_column(9, 9, 100) # 8 column width. Column used for Remediation.

        title_format = workbook.add_format({'bold': True, 'border': 1})
        raws_format = workbook.add_format({'text_wrap': True, 'border': 1})
        critical_format = workbook.add_format({'bg_color': '#7d2105', 'border': 1})
        high_format = workbook.add_format({'bg_color': '#ba2e0f', 'border': 1})
        medium_format = workbook.add_format({'bg_color': '#cc6021', 'border': 1})
        low_format = workbook.add_format({'bg_color': '#b49216', 'border': 1}) # Low and Informal

        worksheet.write_row(0, 0, ['Severity', 'Title', 'Workflow', 'AWS Account ID', 'Description', 'Updated at', 'ProductName', 'Resource Type', 'Resource ID', 'Remediation'], title_format)
        current_line = 1

        # Write each Security Finding row
        for i, finding in enumerate(sorted_findings, start=1):
            severity = finding['Severity']['Label']
            if severity == 'CRITICAL':
                # Apply the critical_format to the "Severity" column cell
                worksheet.write(current_line, 0, severity, critical_format)
            elif severity == 'HIGH':
                # Apply the high_format to the "Severity" column cell
                worksheet.write(current_line, 0, severity, high_format)
            elif severity == 'MEDIUM':
                # Apply the medium_format to the "Severity" column cell
                worksheet.write(current_line, 0, severity, medium_format)
            else:
                worksheet.write(current_line, 0, severity, low_format)

            worksheet.write_row(current_line, 1, [finding['Title'], finding['Workflow']['Status'], finding['AwsAccountId'], finding['Description'], finding['UpdatedAt'], finding['ProductName'], finding['Resources'][0]['Type'], finding['Resources'][0]['Id']])
            current_line+=1

            try:
                worksheet.write_url(current_line, 9, finding['Remediation']['Recommendation']['Url'], string=finding['Remediation']['Recommendation']['Text'])
            except:
                continue
        # Closing workbook
        workbook.close()
    Email_ID = <add receiver's email address>
    
    send_email_with_attachment(Email_ID, filename, xlsx_file)
    
