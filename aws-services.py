import boto3
import pandas as pd
from botocore.exceptions import NoCredentialsError, PartialCredentialsError

def list_services():
    """
    Returns a list of AWS service names.
    """
    return boto3.Session().get_available_services()

def describe_service_resources(service_name):
    """
    Lists resources for a specific AWS service using Boto3.
    Returns a list of resource dictionaries.
    """
    try:
        session = boto3.Session()
        client = session.client(service_name)
        data = []
        
        # Add specific logic for each supported AWS service
        if service_name == "ec2":
            response = client.describe_instances()
            for reservation in response.get("Reservations", []):
                for instance in reservation.get("Instances", []):
                    data.append(instance)
                    
        elif service_name == "s3":
            response = client.list_buckets()
            for bucket in response.get("Buckets", []):
                data.append(bucket)
                
        elif service_name == "rds":
            response = client.describe_db_instances()
            for db_instance in response.get("DBInstances", []):
                data.append(db_instance)
        
        elif service_name == "lambda":
            response = client.list_functions()
            for function in response.get("Functions", []):
                data.append(function)
                
        # Add more service-specific logic here as needed
        
        else:
            data.append({"Service": f"No custom logic implemented for {service_name}"})
        
        return data
    
    except Exception as e:
        return [{"Error": str(e)}]

def main():
    try:
        print("Fetching AWS services...")
        services = list_services()
        
        # Create an Excel writer
        with pd.ExcelWriter("aws_services_report.xlsx", engine="openpyxl") as writer:
            for service in services:
                print(f"Fetching resources for {service}...")
                resources = describe_service_resources(service)
                
                # Convert resources to DataFrame
                if resources:
                    df = pd.DataFrame(resources)
                else:
                    df = pd.DataFrame([{"Message": "No resources found"}])
                
                # Write to a new sheet in the Excel file
                sheet_name = service[:31]  # Sheet names max 31 chars
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("Report saved to aws_services_report.xlsx")
        
    except NoCredentialsError:
        print("Error: No AWS credentials found.")
    except PartialCredentialsError:
        print("Error: Incomplete AWS credentials configuration.")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")

if __name__ == "__main__":
    main()