import boto3
import pandas as pd
from openpyxl import Workbook

def fetch_ec2_instances():
    """Retrieve EC2 instances and their details."""
    ec2_client = boto3.client('ec2')
    instances = []
    response = ec2_client.describe_instances()
    for reservation in response['Reservations']:
        for instance in reservation['Instances']:
            launch_time = instance.get('LaunchTime')
            if launch_time:
                launch_time = launch_time.replace(tzinfo=None)  # Remove timezone

            instances.append({
                "Instance ID": instance.get('InstanceId'),
                "State": instance.get('State', {}).get('Name'),
                "Type": instance.get('InstanceType'),
                "AZ": instance.get('Placement', {}).get('AvailabilityZone'),
                "Public IP": instance.get('PublicIpAddress'),
                "Private IP": instance.get('PrivateIpAddress'),
                "Launch Time": launch_time
            })
    return pd.DataFrame(instances)

def fetch_security_groups():
    """Retrieve Security Groups."""
    ec2_client = boto3.client('ec2')
    security_groups = []
    response = ec2_client.describe_security_groups()
    for sg in response['SecurityGroups']:
        security_groups.append({
            "Group ID": sg.get('GroupId'),
            "Group Name": sg.get('GroupName'),
            "Description": sg.get('Description'),
            "VPC ID": sg.get('VpcId')
        })
    return pd.DataFrame(security_groups)

def fetch_alb():
    """Retrieve ALBs."""
    elbv2_client = boto3.client('elbv2')
    albs = []
    response = elbv2_client.describe_load_balancers()
    for lb in response['LoadBalancers']:
        created_time = lb.get('CreatedTime')
        if created_time:
            created_time = created_time.replace(tzinfo=None)

        albs.append({
            "Load Balancer Name": lb.get('LoadBalancerName'),
            "DNS Name": lb.get('DNSName'),
            "Type": lb.get('Type'),
            "State": lb.get('State', {}).get('Code'),
            "Scheme": lb.get('Scheme'),
            "Created Time": created_time
        })
    return pd.DataFrame(albs)

def fetch_rds_instances():
    """Retrieve RDS instances."""
    rds_client = boto3.client('rds')
    rds_instances = []
    response = rds_client.describe_db_instances()
    for db in response['DBInstances']:
        instance_time = db.get('InstanceCreateTime')
        if instance_time:
            instance_time = instance_time.replace(tzinfo=None)

        rds_instances.append({
            "DB Instance ID": db.get('DBInstanceIdentifier'),
            "Engine": db.get('Engine'),
            "Status": db.get('DBInstanceStatus'),
            "Instance Class": db.get('DBInstanceClass'),
            "Endpoint": db.get('Endpoint', {}).get('Address'),
            "AZ": db.get('AvailabilityZone'),
            "Created Time": instance_time
        })
    return pd.DataFrame(rds_instances)

def fetch_iam_users():
    """Retrieve IAM users."""
    iam_client = boto3.client('iam')
    users = []
    response = iam_client.list_users()
    for user in response['Users']:
        created_on = user.get('CreateDate')
        if created_on:
            created_on = created_on.replace(tzinfo=None)

        users.append({
            "User Name": user.get('UserName'),
            "User ID": user.get('UserId'),
            "ARN": user.get('Arn'),
            "Created On": created_on
        })
    return pd.DataFrame(users)

def fetch_ecs_clusters():
    """Retrieve ECS clusters."""
    ecs_client = boto3.client('ecs')
    clusters = []
    response = ecs_client.list_clusters()
    for arn in response['clusterArns']:
        clusters.append({"Cluster ARN": arn})
    return pd.DataFrame(clusters)

def fetch_s3_buckets():
    """Retrieve S3 buckets."""
    s3_client = boto3.client('s3')
    buckets = []
    response = s3_client.list_buckets()
    for bucket in response['Buckets']:
        creation_date = bucket.get('CreationDate')
        if creation_date:
            creation_date = creation_date.replace(tzinfo=None)

        buckets.append({
            "Bucket Name": bucket.get('Name'),
            "Creation Date": creation_date
        })
    return pd.DataFrame(buckets)

def fetch_cloudfront_distributions():
    """Retrieve CloudFront distributions."""
    cf_client = boto3.client('cloudfront')
    distributions = []
    response = cf_client.list_distributions()
    for dist in response['DistributionList'].get('Items', []):
        distributions.append({
            "ID": dist.get('Id'),
            "Domain Name": dist.get('DomainName'),
            "Status": dist.get('Status'),
            "ARN": dist.get('ARN'),
            "Comment": dist.get('Comment')
        })
    return pd.DataFrame(distributions)

def fetch_acm_certificates():
    """Retrieve ACM certificates."""
    acm_client = boto3.client('acm')
    certificates = []
    response = acm_client.list_certificates()
    for cert in response['CertificateSummaryList']:
        certificates.append({
            "Domain Name": cert.get('DomainName'),
            "Certificate ARN": cert.get('CertificateArn'),
            "Status": cert.get('Status'),
            "Type": cert.get('Type')
        })
    return pd.DataFrame(certificates)

def main():
    print("Fetching AWS service data...")

    # Fetch data for all services
    data_collectors = {
        "EC2": fetch_ec2_instances,
        "SECURITY GROUP": fetch_security_groups,
        "ALB": fetch_alb,
        "RDS": fetch_rds_instances,
        "IAM USER": fetch_iam_users,
        "ECS": fetch_ecs_clusters,
        "S3": fetch_s3_buckets,
        "CLOUDFRONT": fetch_cloudfront_distributions,
        "ACM": fetch_acm_certificates
    }

    # Save to Excel
    output_file = "aws_services_report.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, fetch_function in data_collectors.items():
            print(f"Fetching {sheet_name} data...")
            df = fetch_function()
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Data successfully written to {output_file}")

if __name__ == "__main__":
    main()
