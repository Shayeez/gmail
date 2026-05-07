import os
import io
import sys
import time
import subprocess
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file drive_token.json.
SCOPES = ['https://www.googleapis.com/auth/drive']
DRIVE_TOKEN_FILE = 'drive_token.json'
CREDENTIALS_FILE = 'credentials.json'

# The ID of the shared folder "@Gdrive Shamim_shayeez"
PARENT_FOLDER_ID = '1Qo_94sphOgbE8x9G42zAk0w6ksSsbY98'

def get_drive_service():
    """Authenticates and returns the Google Drive API service."""
    creds = None
    if os.path.exists(DRIVE_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(DRIVE_TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"Error: {CREDENTIALS_FILE} not found. Please place it in the same directory.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            # This will open a browser to authenticate.
            creds = flow.run_local_server(port=0)
        with open(DRIVE_TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def find_folder(service, name, parent_id):
    """Finds a folder by name inside a specific parent folder."""
    query = f"name='{name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(
        q=query, 
        spaces='drive',
        corpora='allDrives',
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        fields="files(id, name)"
    ).execute()
    items = results.get('files', [])
    if not items:
        return None
    return items[0]['id']

def download_file(service, file_id, file_name, dest_folder="."):
    """Downloads a file from Google Drive."""
    os.makedirs(dest_folder, exist_ok=True)
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    file_path = os.path.join(dest_folder, file_name)
    fh = io.FileIO(file_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    print(f"Downloading {file_name} to {dest_folder}/...")
    while done is False:
        status, done = downloader.next_chunk()

def upload_file(service, file_path, parent_id):
    """Uploads or updates a file on Google Drive."""
    file_name = os.path.basename(file_path)
    print(f"Uploading {file_name}...")
    
    # Check if file exists to update it, else create a new one
    query = f"name='{file_name}' and '{parent_id}' in parents and trashed=false"
    results = service.files().list(
        q=query, 
        corpora='allDrives',
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        fields="files(id, name)"
    ).execute()
    items = results.get('files', [])
    
    # Determine mime type based on extension
    mime_type = 'application/octet-stream'
    if file_name.endswith('.xlsx'):
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    elif file_name.endswith('.csv'):
        mime_type = 'text/csv'
    elif file_name.endswith('.json'):
        mime_type = 'application/json'
        
    media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
    
    if items:
        # Update existing file
        file_id = items[0]['id']
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
    else:
        # Create new file
        file_metadata = {'name': file_name, 'parents': [parent_id]}
        service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()

def sync_down(service, gmail_folder_id):
    """Downloads all code and config files from Drive to local."""
    print("--- DOWNLOADING LATEST CODE AND SETTINGS ---")
    query = f"'{gmail_folder_id}' in parents and trashed=false"
    results = service.files().list(
        q=query, 
        corpora='allDrives',
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        fields="files(id, name, mimeType)"
    ).execute()
    
    tokens_folder_id = None
    for item in results.get('files', []):
        if item['mimeType'] == 'application/vnd.google-apps.folder':
            if item['name'] == 'tokens':
                tokens_folder_id = item['id']
            continue
            
        # Download strictly necessary code and config
        if item['name'] in ['mail.py', 'priorities.json', 'credentials.json']:
            download_file(service, item['id'], item['name'])
                
    # Download tokens
    if tokens_folder_id:
        query = f"'{tokens_folder_id}' in parents and trashed=false"
        token_results = service.files().list(
            q=query, 
            corpora='allDrives',
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            fields="files(id, name)"
        ).execute()
        for item in token_results.get('files', []):
            if item['name'].endswith('.json'):
                download_file(service, item['id'], item['name'], 'tokens')

def sync_up(service, gmail_folder_id, run_start_time):
    """Uploads the generated outputs back to Drive."""
    print("--- UPLOADING RESULTS ---")
    


    # 2. Upload output excel/csv files (ignoring temporary ~$ files)
    output_folder_id = find_folder(service, 'output', gmail_folder_id)
    if not output_folder_id:
        metadata = {'name': 'output', 'parents': [gmail_folder_id], 'mimeType': 'application/vnd.google-apps.folder'}
        folder = service.files().create(body=metadata, supportsAllDrives=True).execute()
        output_folder_id = folder['id']
        
    if os.path.exists('output'):
        for f in os.listdir('output'):
            if not f.startswith('~$') and (f.endswith('.xlsx') or f.endswith('.csv')):
                file_path = os.path.join('output', f)
                # Only upload if the file was modified by mail.py in this run
                if os.path.getmtime(file_path) > run_start_time:
                    upload_file(service, file_path, output_folder_id)

    # 3. Upload logs
    logs_folder_id = find_folder(service, 'logs', gmail_folder_id)
    if not logs_folder_id:
        metadata = {'name': 'logs', 'parents': [gmail_folder_id], 'mimeType': 'application/vnd.google-apps.folder'}
        folder = service.files().create(body=metadata, supportsAllDrives=True).execute()
        logs_folder_id = folder['id']
        
    if os.path.exists('logs'):
        for f in os.listdir('logs'):
            if f.endswith('.log'):
                file_path = os.path.join('logs', f)
                # Only upload if the file was modified by mail.py in this run
                if os.path.getmtime(file_path) > run_start_time:
                    upload_file(service, file_path, logs_folder_id)

def main():
    print("Starting Server Runner...")
    service = get_drive_service()
    
    # 1. Find the Gmail folder inside the shared folder
    print("\nLocating Gmail folder in Google Drive...")
    gmail_folder_id = find_folder(service, 'Gmail', PARENT_FOLDER_ID)
    if not gmail_folder_id:
        print("Error: Could not find 'Gmail' folder inside the shared Drive folder.")
        sys.exit(1)
        
    # 2. Download code and settings
    sync_down(service, gmail_folder_id)
    
    # 3. Run the mail script
    print("\n--- RUNNING MAIL SCRIPT ---")
    
    # Record exactly when mail.py starts
    run_start_time = time.time()
    
    # Pass along any arguments given to this script down to mail.py
    cmd = [sys.executable, 'mail.py'] + sys.argv[1:]
    if len(sys.argv) == 1:
        # Default behavior if no args passed: run all for today
        cmd.extend(['--run-all', '--today'])
        
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"mail.py execution failed with error: {e}")
    
    # 4. Upload results
    print("\n")
    sync_up(service, gmail_folder_id, run_start_time)
    
    print("\nServer Runner complete!")

if __name__ == '__main__':
    main()
