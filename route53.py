import boto3
import pandas as pd
from openpyxl import Workbook

def get_hosted_zones_and_records():
    # Initialize boto3 Route 53 client
    client = boto3.client('route53')

    # Create an Excel writer
    excel_writer = pd.ExcelWriter('route53_hosted_zones.xlsx', engine='openpyxl')

    # Get all hosted zones
    response = client.list_hosted_zones()
    hosted_zones = response['HostedZones']

    print(f"Number of Hosted Zones found: {len(hosted_zones)}")

    # Loop through each hosted zone
    for zone in hosted_zones:
        hosted_zone_id = zone['Id'].split('/')[-1]
        zone_name = zone['Name']

        print(f"Fetching records for Hosted Zone: {zone_name} (ID: {hosted_zone_id})")

        # Fetch records for this hosted zone
        records = []
        paginator = client.get_paginator('list_resource_record_sets')
        for page in paginator.paginate(HostedZoneId=hosted_zone_id):
            records.extend(page['ResourceRecordSets'])

        # Prepare data for this hosted zone
        zone_data = []
        for record in records:
            record_data = {
                'Name': record.get('Name'),
                'Type': record.get('Type'),
                'TTL': record.get('TTL', 'N/A'),
                'Value': ', '.join([r['Value'] for r in record.get('ResourceRecords', [])]) if 'ResourceRecords' in record else 'Alias'
            }
            zone_data.append(record_data)

        # Convert to DataFrame
        df = pd.DataFrame(zone_data)

        # Add to Excel (1 sheet per Hosted Zone)
        sheet_name = zone_name.strip('.').replace('.', '_')[:31]  # Ensure Excel sheet name is valid
        df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
        print(f"Added Hosted Zone {zone_name} to Excel sheet.")

    # Save the Excel file
    excel_writer.close()
    print("Excel file 'route53_hosted_zones.xlsx' has been created successfully.")

if __name__ == "__main__":
    get_hosted_zones_and_records()
