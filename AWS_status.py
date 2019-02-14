import datetime
import xlwt
from xlwt import Workbook
import os
import boto3
from botocore.exceptions import ClientError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

region_list = ['us-east-1', 'us-east-2', 'us-west-1', 'us-west-2', 'ca-central-1']

xlFilename = "awsStatus.xls"

wb = Workbook()
now = datetime.datetime.now()
START = "{}-{}-{}".format(now.year, now.month, now.day)
END = "{}-{}-{}".format(now.year, now.month, now.day)

def awsCost(region):
    client = boto3.client('ce', region)
    a = client.get_cost_and_usage(
    TimePeriod={"Start": START, "End": END},
    Granularity="DAILY",
    Metrics=["UnblendedCost"],)
    return("Total spend from {} to {} is {} USD".format(
        START, END, a['ResultsByTime'][0]['Total']['UnblendedCost']['Amount']))

for region in region_list:
    print 'REGION:', region
    sheet = wb.add_sheet(region)
    style = xlwt.easyxf('font: bold 1, color red')
    awsDict =["name", "instance_type", "instance.id", "instance.state", "instance.launch_time", "instance.private_ip_address", "instance.public_ip_address", "volume.id/Volume.Size"]
    RowStarting = 0
    for k,l in enumerate(awsDict):
        sheet.write(RowStarting, k, l,style)
    ec2 = boto3.resource('ec2', region)
    for instance in ec2.instances.all():
        RowStarting +=1
        print instance
        for tag in instance.tags:
            if 'Name'in tag['Key']:
                name = tag['Value']

        sheet.write(RowStarting, 0, name)
        sheet.write(RowStarting, 1, instance.instance_type)
        sheet.write(RowStarting, 2, instance.id)
        sheet.write(RowStarting, 3, instance.state["Name"])
        sheet.write(RowStarting, 4, str(instance.launch_time))
        sheet.write(RowStarting, 5, instance.private_ip_address)
        sheet.write(RowStarting, 6, instance.public_ip_address)
        volData = ""
        for volume in instance.volumes.all():
            volData = volData +"//" + str(volume.size)+"/"+volume.id
        sheet.write(RowStarting,7, volData)
    try:
        cost= awsCost(region)
    except:
        cost ="Not Authorised"
    sheet.write(RowStarting+5,0, "AverageRegionCost:",style)
    sheet.write(RowStarting+5,2, cost,style)

wb.save(xlFilename)

SENDER = "Sender-Address"
RECIPIENT = "Receiver's Address"
AWS_REGION = "Region"
SUBJECT = "Subject"
ATTACHMENT = xlFilename
BODY_TEXT = "Hello,\r\nPlease see the attached file of aws status."
BODY_HTML = """<html>
<head></head>
<body>
<h1>Hello!</h1>
<p>Please see the attached file of aws status.</p>
</body>
</html>
            """

CHARSET = "utf-8"

client = boto3.client('ses',region_name=AWS_REGION)
msg = MIMEMultipart('mixed')
msg['Subject'] = SUBJECT
msg['From'] = SENDER
msg['To'] = RECIPIENT
msg_body = MIMEMultipart('alternative')
textpart = MIMEText(BODY_TEXT.encode(CHARSET), 'plain', CHARSET)
htmlpart = MIMEText(BODY_HTML.encode(CHARSET), 'html', CHARSET)
msg_body.attach(textpart)
msg_body.attach(htmlpart)
att = MIMEApplication(open(ATTACHMENT, 'rb').read())
att.add_header('Content-Disposition','attachment',filename=os.path.basename(ATTACHMENT))
msg.attach(msg_body)
msg.attach(att)
try:
    response = client.send_raw_email(
        Source=SENDER,
        Destinations=[
            RECIPIENT
        ],
        RawMessage={
            'Data':msg.as_string(),
        },
    )
except ClientError as e:
    print(e.response['Error']['Message'])
else:
    print("Email sent! Message ID:"),
    print(response['MessageId'])
