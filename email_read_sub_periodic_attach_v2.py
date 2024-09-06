import os
import uuid
import requests
import json
from datetime import datetime, timedelta
import sys
import base64

# Function to get OAuth2 token
def get_oauth2_token(client_id, client_secret, tenant_id):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    # Data for requesting the token
    token_data = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',  # Scope to access Microsoft Graph API
        'client_secret': client_secret,
        'grant_type': 'client_credentials',
    }
    
    # Request token
    response = requests.post(token_url, data=token_data)
    if response.status_code == 200:
        token_json = response.json()
        return token_json.get('access_token')  # Return the access token
    else:
        raise Exception(f"Failed to get access token: {response.text}")

# Function to get the latest messages from the mailbox
def get_latest_messages(token, user_email, within_last_hour=False):
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
    headers = {
        'Authorization': f'Bearer {token}',  # Bearer token for authorization
        'Content-Type': 'application/json'
    }
    
    # If we want to filter emails from the past hour
    if within_last_hour:
        current_time = datetime.utcnow()
        one_hour_ago = current_time - timedelta(hours=1)
        one_hour_ago_iso = one_hour_ago.strftime('%Y-%m-%dT%H:%M:%SZ')  # ISO format
        url += f"?$filter=receivedDateTime ge {one_hour_ago_iso}"

    # Send GET request to Microsoft Graph API to retrieve messages
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()  # Return messages in JSON format
    else:
        raise Exception(f"Failed to get messages: {response.text}")

# Function to check for attachments in the email and download them
def download_attachments(token, user_email, message_id, attachment_dir):
    # URL to fetch attachments using the user_email and message_id
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{message_id}/attachments"
    headers = {
        'Authorization': f'Bearer {token}',  # Bearer token for authorization
        'Content-Type': 'application/json'
    }
    
    # Request the attachments
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        attachments = response.json().get('value', [])
        
        for attachment in attachments:
            # Only process file attachments (ignore item attachments like inline images)
            if attachment['@odata.type'] == '#microsoft.graph.fileAttachment':
                file_name = attachment['name']
                file_content = attachment['contentBytes']
                
                # Decode the Base64 content
                decoded_content = base64.b64decode(file_content)
                
                # Generate a unique filename by appending a UUID
                unique_id = uuid.uuid4()  # Generate a UUID
                file_extension = os.path.splitext(file_name)[1]  # Get the file extension
                new_file_name = f"{os.path.splitext(file_name)[0]}_{unique_id}{file_extension}"
                
                # Save the decoded attachment to the specified directory
                file_path = os.path.join(attachment_dir, new_file_name)
                with open(file_path, 'wb') as f:
                    f.write(decoded_content)
                
                print(f"Downloaded attachment: {new_file_name}")
                print(f"##gbStart##attachment_file_path##splitKeyValue##{file_path}##gbEnd##")
    else:
        raise Exception(f"Failed to get attachments for message {message_id}: {response.text}")

# Function to load the last processed email details from a file
def load_last_processed_email(last_email_file):
    if os.path.exists(last_email_file):
        with open(last_email_file, 'r') as f:
            return json.load(f)
    return None

# Function to save the last processed email details to a file
def save_last_processed_email(last_email_file, last_email):
    with open(last_email_file, 'w') as f:
        json.dump(last_email, f)

# Function to save matched emails to a file
def save_matched_email(matched_emails_file, email_details):
    if not os.path.exists(matched_emails_file):
        with open(matched_emails_file, 'w') as f:
            json.dump([], f)

    with open(matched_emails_file, 'r+') as f:
        emails = json.load(f)
        emails.append(email_details)
        f.seek(0)
        json.dump(emails, f, indent=4)

# Function to check and process new emails after the last processed email
def process_new_emails(token, user_email, target_subject, attachment_dir, last_email_file, matched_emails_file):
    # Load the last processed email
    last_processed_email = load_last_processed_email(last_email_file)

    # Check if the last processed email is not available
    within_last_hour = False
    if not last_processed_email:
        print("No last processed email found. Fetching emails from the last hour...")
        within_last_hour = True

    # Get the latest messages from the mailbox, filter by past hour if needed
    messages = get_latest_messages(token, user_email, within_last_hour)

    for message in messages['value']:
        # If last processed email exists, skip emails that were received before it
        if last_processed_email and message['receivedDateTime'] <= last_processed_email['receivedDateTime']:
            continue

        # Check if the subject matches the target subject
        if target_subject.lower() in message['subject'].lower():
            # Prepare the email details to store
            email_details = {
                'subject': message['subject'],
                'from': message['from']['emailAddress']['address'],
                'receivedDateTime': message['receivedDateTime'],
                'bodyPreview': message['bodyPreview']
            }

            # Save the matching email to the matched emails file
            save_matched_email(matched_emails_file, email_details)

            # Check for attachments and download them
            download_attachments(token, user_email, message['id'], attachment_dir)

            # Update the last processed email
            save_last_processed_email(last_email_file, {
                'id': message['id'],
                'receivedDateTime': message['receivedDateTime']
            })

            print(f"Matching email processed: {email_details}")

# Main execution to monitor the mailbox
if __name__ == "__main__":
    # Check for required arguments using sys.argv
    if len(sys.argv) != 6:
        print("Usage: python script.py <attachment_dir> <client_id> <client_secret> <tenant_id> <mailbox_email>")
        sys.exit(1)

    # Get arguments from sys.argv
    ATTACHMENT_DIR = sys.argv[1]
    client_id = sys.argv[2]
    client_secret = sys.argv[3]
    tenant_id = sys.argv[4]
    mailbox_email = sys.argv[5]

    # Set up paths based on the provided attachment directory
    LAST_EMAIL_FILE = os.path.join(ATTACHMENT_DIR, 'last_processed_email.json')
    MATCHED_EMAILS_FILE = os.path.join(ATTACHMENT_DIR, 'matched_emails.json')

    # Ensure the attachment directory exists
    if not os.path.exists(ATTACHMENT_DIR):
        os.makedirs(ATTACHMENT_DIR)

    # Define the target subject to look for
    target_subject = "User Creation/Modification Request"

    try:
        # Step 1: Get OAuth2 token
        token = get_oauth2_token(client_id, client_secret, tenant_id)
        print("Access token acquired successfully.")

        # Step 2: Process new emails after the last processed email
        process_new_emails(token, mailbox_email, target_subject, ATTACHMENT_DIR, LAST_EMAIL_FILE, MATCHED_EMAILS_FILE)
        print("Email processing completed.")

    except Exception as e:
        print(f"An error occurred: {e}")
