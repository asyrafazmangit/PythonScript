import boto3
import pandas as pd
import re

def is_valid_instance_id(target_id):
    """
    Check if the given target ID is a valid EC2 instance ID.
    """
    return re.match(r'^i-[0-9a-f]{17}$', target_id)

def list_load_balancers_and_resources():
    try:
        elbv2_client = boto3.client('elbv2')
        ec2_client = boto3.client('ec2')

        # Data storage for Excel
        load_balancer_data = []
        target_instance_data = []

        # Fetch all load balancers
        load_balancers = elbv2_client.describe_load_balancers().get('LoadBalancers', [])
        
        for lb in load_balancers:
            print(f"Processing Load Balancer: {lb['LoadBalancerName']}")
            
            # Fetch Target Groups
            target_groups = elbv2_client.describe_target_groups(LoadBalancerArn=lb['LoadBalancerArn'])['TargetGroups']
            for tg in target_groups:
                targets = elbv2_client.describe_target_health(TargetGroupArn=tg['TargetGroupArn'])['TargetHealthDescriptions']
                
                for target in targets:
                    target_id = target['Target']['Id']
                    
                    # Check if the target is an EC2 Instance ID or IP
                    if is_valid_instance_id(target_id):
                        # Describe EC2 Instance
                        ec2_details = ec2_client.describe_instances(InstanceIds=[target_id])
                        for reservation in ec2_details['Reservations']:
                            for instance in reservation['Instances']:
                                target_instance_data.append({
                                    "LoadBalancerName": lb['LoadBalancerName'],
                                    "TargetGroupName": tg['TargetGroupName'],
                                    "InstanceID": target_id,
                                    "State": instance['State']['Name'],
                                    "PrivateIP": instance['PrivateIpAddress']
                                })
                    else:
                        # It's an IP target, add it directly
                        target_instance_data.append({
                            "LoadBalancerName": lb['LoadBalancerName'],
                            "TargetGroupName": tg['TargetGroupName'],
                            "InstanceID": "N/A",
                            "State": "N/A",
                            "PrivateIP": target_id
                        })

        # Save results to Excel
        with pd.ExcelWriter("aws_load_balancers_fixed.xlsx", engine="openpyxl") as writer:
            pd.DataFrame(target_instance_data).to_excel(writer, sheet_name="Target Instances", index=False)
        
        print("Data has been saved to 'aws_load_balancers_fixed.xlsx'")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    list_load_balancers_and_resources()
