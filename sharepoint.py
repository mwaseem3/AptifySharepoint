import requests
import json
import os

def get_access_token(tenant_id, client_id, client_secret):
    """
    Obtain an OAuth2 access token from Azure Active Directory.

    :param tenant_id: Azure AD tenant ID.
    :param client_id: Azure AD application (client) ID.
    :param client_secret: Azure AD application secret (client secret).
    :return: Access token as a string or None if the request fails.
    """
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    token_data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default',
        'grant_type': 'client_credentials',
    }

    try:
        response = requests.post(token_url, data=token_data)
        response.raise_for_status()
        return response.json().get('access_token')
    except Exception as e:
        print(f"Failed to obtain access token: {e}")
        return None

def get_site_id(access_token, site_url):
    """
    Retrieve the unique ID of a SharePoint site using the Microsoft Graph API.

    :param access_token: OAuth2 access token for authentication.
    :param site_url: The URL to the Microsoft Graph endpoint for retrieving site details.
    :return: Site ID as a string or None if the request fails.
    """
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(site_url, headers=headers)
        response.raise_for_status()
        return response.json().get('id')
    except Exception as e:
        print(f"Failed to retrieve site ID: {e}")
        return None

def get_shared_documents_drive_id(access_token, drive_url):
    """
    Fetch the drive ID associated with the Shared Documents library in SharePoint.

    :param access_token: OAuth2 access token for authentication.
    :param drive_url: The URL to the Microsoft Graph endpoint for retrieving drive details.
    :return: Drive ID as a string or None if the request fails.
    """
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(drive_url, headers=headers)
        x=response.json()
        print(x)
        response.raise_for_status()
        drives = x['value']
        return drives[0]['id']  
    except Exception as e:
        print(f"Failed to retrieve Shared Documents drive ID: {e}")
        return None

def list_folders_in_drive(access_token, folders_url):
    """
    List all folders in a specified drive and save the data to a JSON file.

    :param access_token: OAuth2 access token for authentication.
    :param folders_url: The URL to the Microsoft Graph endpoint for listing folders in the drive.
    """
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(folders_url, headers=headers)
        response.raise_for_status()
        data = response.json()['value']
        with open('folders_data_value.json', 'w') as file:
            json.dump(data, file, indent=4)
        print("Folder data saved to folders_data_value.json")
    except Exception as e:
        print(f"Failed to list folders: {e}")


def create_folder(access_token, drive_id, folder_name):
    """
    Create a new folder in a specified SharePoint drive or return the link if it already exists.

    :param access_token: OAuth2 access token for authentication.
    :param drive_id: The ID of the SharePoint drive where the folder will be created or found.
    :param folder_name: The name of the folder to be created or found.
    :return: The link to the existing or newly created folder, or None if the operation fails.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Check if the folder already exists
    # search_folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children/{folder_name}"
    # try:
    #     search_response = requests.get(search_folder_url, headers=headers)
    #     if search_response.status_code == 200:
    #         # Folder exists, return the folder's link
    #         folder_data = search_response.json()
    #         print('already exists')
    #         return folder_data.get('webUrl')  # Adjust this key if necessary
        
    # except Exception as e:
    #     print(f"Failed to search for the existing folder: {e}")

    # Folder doesn't exist, create a new folder
    create_folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
    folder_data = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }

    try:
        create_response = requests.post(create_folder_url, headers=headers, json=folder_data)
        create_response.raise_for_status()
        created_folder = create_response.json()
        return created_folder.get('webUrl')  # Adjust this key if necessary to get the correct link
    except Exception as e:
        print(f"Failed to create folder: {e}")
        return None

    

def upload_pdf(access_token, drive_id, folder_name, pdf_file_name, pdf_file_path):
    """
    Upload a PDF file to a specified folder within a SharePoint drive. If a file with the same name
    exists, the file is uploaded with a modified name by appending a numerical suffix.

    :param access_token: OAuth2 access token for authentication.
    :param drive_id: The ID of the SharePoint drive.
    :param folder_name: The name of the folder where the PDF will be uploaded.
    :param pdf_file_name: The name of the PDF file to be uploaded.
    :param pdf_file_path: The path to the PDF on the local machine.
    :return: JSON response from the server or None if the upload fails.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/pdf'
    }

    base_filename, file_extension = pdf_file_name.rsplit('.', 1)
    file_exists = True
    count = 1

    while file_exists:
        check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_name}/{pdf_file_name}"
        response = requests.get(check_url, headers={'Authorization': f'Bearer {access_token}', 'Content-Type' :'application/json' })

        if response.status_code == 200:
            file_url = response.json().get('webUrl')
            print("File already exists: ", file_url)
            return file_url, False
        elif response.status_code==401:
            return None,True
        else:
            # File does not exist, ready to upload
            file_exists = False

    upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_name}/{pdf_file_name}:/content"

    try:
        with open(pdf_file_path, "rb") as pdf_file:
            file_content = pdf_file.read()
        upload_response = requests.put(upload_url, headers=headers, data=file_content)
        upload_response.raise_for_status()
        created_file= upload_response.json()
        print(created_file.get('webUrl'))
        return created_file.get('webUrl'),False
    except Exception as e:
        if upload_response.status_code==401:
            return None,True
        print(f"Failed to upload PDF file: {e}")
        return None,False

def get_network_link(sharepoint_link):
    # sharepoint_link = "https://nbcrna.sharepoint.com/sites/HistoricalTranscripts/Shared%20Documents/Test2/Aalamjot%20Bhuller1_Transcript_2.pdf"

# Step 1: Remove the base URL and decode URL-encoded spaces (%20)
    relative_path = sharepoint_link.replace("https://nbcrna.sharepoint.com/sites/HistoricalTranscripts/Shared%20Documents/", "").replace("%20", " ")

# Step 2: Replace forward slashes with backslashes
    file_path = relative_path.replace("/", "\\")

# Step 3: Prepend the network drive path
    network_path = f"\\\\nbcrna-file1\\transcripts$\\{file_path}"

    return network_path




client_id = os.getenv("client_id")
tenant_id = os.getenv("tenant_id")
client_secret = os.getenv("client_secret")

site_path = "/sites/AptifyAttachment"
base_url = "https://graph.microsoft.com/v1.0"
site_url = f"{base_url}/sites/nbcrna.sharepoint.com:{site_path}"

at=get_access_token(tenant_id, client_id, client_secret)

did=os.getenv("drive_id")
folders_url = f"{base_url}/drives/{did}/root/children"
# list_folders_in_drive(at, folders_url)
create_folder(at,did,'TestingAptify2')