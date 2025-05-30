import msal
import requests
import webbrowser
from flask import Flask, request
from io import BytesIO
import pandas as pd
from datetime import datetime
import ibm_boto3
from ibm_botocore.client import Config, ClientError
import pytz
from bs4 import BeautifulSoup
import time
import json  
import base64  # <-- Added for decoding base64

app = Flask(__name__)

client_id = 'c0654350-450c-4ba2-a030-4c82b6fa25bd'
client_secret = 'zKL8Q~R4UDIKi9s9yMC2bnzRUjrDF5C.qcChzcKm'
redirect_uri = 'http://localhost:8000/getAToken'
authority = f'https://login.microsoftonline.com/common'
app2_api_key = "6VS7iOH0uxS1Lh3HRYhsu1s30--Rqr7kFzpZOXm2kGLu"
COS_ENDPOINT = "https://s3.eu-gb.cloud-object-storage.appdomain.cloud"
COS_API_KEY_ID = "6VS7iOH0uxS1Lh3HRYhsu1s30--Rqr7kFzpZOXm2kGLu"
COS_INSTANCE_CRN = "crn:v1:bluemix:public:cloud-object-storage:global:a/c5c9c7d1f51e4fa1864855f56508e356:d7cc671f-1238-4111-930b-a159132a0e5f:bucket:mas-email-extraction-bucket"
BUCKET_NAME = "mas-email-extraction-bucket"
EXCEL_FILE_NAME = "mainExcel.xlsx"

scopes = [
    'User.Read',
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Files.Read',
    'Files.ReadWrite'
]

app_ms = msal.PublicClientApplication(client_id, authority=authority)

cos = ibm_boto3.resource("s3",
                         ibm_api_key_id=COS_API_KEY_ID,
                         ibm_service_instance_id=COS_INSTANCE_CRN,
                         config=Config(signature_version="oauth"),
                         endpoint_url=COS_ENDPOINT,
                         verify=False)

# Uploads a file to the specified IBM Cloud Object Storage bucket from memory.
def upload_file_to_bucket(bucket_name, file_bytes, key):
    try:
        cos.Bucket(bucket_name).put_object(Key=key, Body=file_bytes)
        # print(f"File uploaded to {bucket_name} with key {key}")
    except ClientError as e:
        # print(f"Error uploading file: {e}")
        pass

# Downloads a file from the specified IBM Cloud Object Storage bucket to memory.
def download_file_from_bucket(bucket_name, key):
    try:
        file_obj = cos.Bucket(bucket_name).Object(key).get()
        return file_obj['Body'].read()
    except ClientError as e:
        if e.response['Error']['Code'] == '404':
            # print(f"The file {key} does not exist in the bucket {bucket_name}.")
            return None
        else:
            # print(f"Error downloading file: {e}")
            return None

# Generates the authorization URL for Microsoft login.
def generate_auth_url():
    auth_url = app_ms.get_authorization_request_url(scopes, redirect_uri=redirect_uri)
    # print("Generated authorization URL")
    return auth_url

# Redirects the user to the Microsoft login page.
@app.route('/')
def index():
    auth_url = generate_auth_url()
    webbrowser.open(auth_url)
    # print("Opened web browser for Microsoft login")
    return "Redirecting to Microsoft login..."

# Exchanges the authorization code for an access token.
def exchange_code_for_token(auth_code):
    token_url = f"https://login.microsoftonline.com/common/oauth2/v2.0/token"
    token_data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'code': auth_code,
        'redirect_uri': redirect_uri,
        'grant_type': 'authorization_code',
        'scope': ' '.join(scopes),
    }
    response = requests.post(token_url, data=token_data, verify=False)
    if response.status_code == 200:
        # print("Exchanged authorization code for access token")
        return response.json()['access_token']
    else:
        return None

# Cleans the HTML content and returns plain text.
def clean_html(raw_html):
    soup = BeautifulSoup(raw_html, "html.parser")
    # print("Cleaned HTML content")
    return soup.get_text()

# Fetches PDF attachments from an email message.
def get_pdf_attachments(message_id, access_token):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments'
    response = requests.get(url, headers=headers, verify=False)
    
    if response.status_code == 200:
        attachments = response.json().get('value', [])
        pdf_attachments = [attachment for attachment in attachments if attachment.get('name', '').lower().endswith('.pdf')]
        # print(f"Fetched PDF attachments for message ID {message_id}")
        return pdf_attachments
    else:
        return []

# Marks an email message as read.
def mark_message_as_read(message_id, access_token):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'
    data = {
        'isRead': True
    }
    response = requests.patch(url, headers=headers, json=data, verify=False)
    if response.status_code == 200:
        # print(f"Marked message ID {message_id} as read")
        pass
    return response.status_code == 200

# Updates the Excel file in memory and uploads it to the IBM Cloud Object Storage.
def update_excel_file(all_messages):
    excel_data = BytesIO()  # Create an in-memory file for Excel
    try:
        # Try to download the existing Excel file from IBM Cloud Object Storage
        existing_excel_data = download_file_from_bucket(BUCKET_NAME, EXCEL_FILE_NAME)
        if existing_excel_data:
            existing_df = pd.read_excel(BytesIO(existing_excel_data))
            df = pd.DataFrame(all_messages)
            combined_df = pd.concat([df, existing_df], ignore_index=True)
        else:
            combined_df = pd.DataFrame(all_messages)
    except Exception as e:
        combined_df = pd.DataFrame(all_messages)

    combined_df.to_excel(excel_data, index=False)
    excel_data.seek(0)  # Move to the beginning of the BytesIO stream
    upload_file_to_bucket(BUCKET_NAME, excel_data, EXCEL_FILE_NAME)
    # print(f"Updated Excel file '{EXCEL_FILE_NAME}'")

# Fetches data from the endpoint using the provided API key and PDF filename.
def fetch_data_from_endpoint(app2_api_key,attachment_name):
    API_KEY = app2_api_key
    try:
        print("Requesting access token...")
        token_response = requests.post(
            'https://iam.cloud.ibm.com/identity/token',
            data={"apikey": API_KEY, "grant_type": 'urn:ibm:params:oauth:grant-type:apikey'},
            verify=False
        )
        token_response.raise_for_status()
        mltoken = token_response.json()["access_token"]
        print(mltoken)
        print("Access token obtained successfully.")
    except requests.exceptions.RequestException as e:
        print(f"Error obtaining access token: {e}")
        return None

    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer ' + mltoken}
    
    payload_scoring = {
        "input_data": [
            {
                "fields": ["input_bucket_name", "input_file_name"],
                "values": [["mas-email-extraction-bucket", attachment_name]]
            }
        ]
    }
    params = {
        "version": "2021-05-01"
    }

    try:
        # Debugging: Print the payload as a JSON string
        print("Sending scoring request with payload:", json.dumps(payload_scoring, indent=4))
        
        response_scoring = requests.post(
            'https://eu-de.ml.cloud.ibm.com/ml/v4/deployments/7325dd32-512b-4713-be37-b7d4e50640b8/predictions?version=2021-05-01',
            json=payload_scoring,  # Use `json` to send the payload as JSON
            headers=headers,
            params=params,  # Pass the `params` dictionary correctly
            verify=False
        )
        print("Waiting for response...")
        time.sleep(10)
        response_scoring.raise_for_status()
        print("Scoring response received successfully.")
        print("Scoring response:", response_scoring.json())
        return response_scoring.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data from endpoint: {e}")
        if e.response is not None:
            print("Response content:", e.response.content)
        return None

# Fetches unread emails and processes them.
def fetch_unread_emails(access_token):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    url = 'https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false&$top=50'
    response = requests.get(url, headers=headers, verify=False)
    
    if response.status_code == 200:
        messages = response.json()
        if messages['value']:
            all_messages = []
            all_pdfs = []
            for message in messages['value']:
                # print(f"Processing message ID: {message['id']}")
                sender_info = message.get('from', {})
                sender_name = sender_info.get('emailAddress', {}).get('name', 'Unknown Sender')
                sender_email = sender_info.get('emailAddress', {}).get('address', 'Unknown Email')
                # print(f"Sender Name: {sender_name}, Sender Email: {sender_email}")
                message_body_text = clean_html(message.get('bodyPreview', ''))
                received_time = message.get('receivedDateTime', 'Unknown Time')
                received_time_utc = datetime.strptime(received_time, '%Y-%m-%dT%H:%M:%SZ')
                received_time_ist = received_time_utc.astimezone(pytz.timezone('Asia/Kolkata')).strftime('%Y-%m-%d %H:%M:%S')
                pdf_attachments = get_pdf_attachments(message['id'], access_token)
                all_pdfs.extend(pdf_attachments)

                for attachment in pdf_attachments:
                    attachment_name = attachment['name']
                    attachment_content = attachment['contentBytes']
                    # Decode base64 to get original PDF bytes
                    attachment_bytes = BytesIO(base64.b64decode(attachment_content))
                    upload_file_to_bucket(BUCKET_NAME, attachment_bytes, attachment_name)

                    all_messages.append({
                        'Sender Name': "ABC Conmpany Pvt. Ltd.",
                        'Sender Email': "ABCcomapny@outlook.com",
                        'Received Time': received_time_ist,
                        'Message Body': message_body_text,
                        'PDF Attachment': attachment_name
                    })

                    endpoint_data = fetch_data_from_endpoint(app2_api_key, attachment_name)
                    if endpoint_data:
                        endpoint_values = endpoint_data.get('predictions', [{}])[0].get('values', {})
                        all_messages[-1].update({
                            'Asset Numbers': ', '.join(endpoint_values.get('asset numbers', [])),
                            'Assignment End Date': endpoint_values.get('assignment end date', ''),
                            'Assignment Start Date': endpoint_values.get('assignment start date', ''),
                            'Comments': endpoint_values.get('comments', ''),
                            'Location': endpoint_values.get('location', ''),
                            'PO Number': endpoint_values.get('po number', ''),
                            'Supplier': endpoint_values.get('supplier', ''),
                            'Workorder': endpoint_values.get('workorder', '')
                        })

                mark_message_as_read(message['id'], access_token)

            update_excel_file(all_messages)
            # print("Fetched unread emails and saved to Excel file")

            return f"Unread messages have been saved to '{EXCEL_FILE_NAME}' and marked as read. PDF attachments: {all_pdfs}"
        else:
            return "No unread messages found."
    else:
        return f"Error fetching messages: {response.status_code} - {response.text}"

# Handles the /getAToken route to exchange the authorization code for an access token and fetch unread emails.
@app.route('/getAToken', methods=['GET'])
def get_token():
    auth_code = request.args.get('code')
    if auth_code:
        access_token = exchange_code_for_token(auth_code)
        if access_token:
            result = fetch_unread_emails(access_token)
            return result
        else:
            return "Failed to obtain access token", 400
    else:
        return "Authorization code not found in the request", 400

if __name__ == "__main__":
    app.run(port=8000)