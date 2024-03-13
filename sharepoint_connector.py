import json
import os
import requests
from urllib.request import urlretrieve
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv()

class SharepointDownloader():
    """
    This class provides a reusable tool for connecting to SharePoint and manipulating dataframes sourced from it.

    The behavior of the class revolves around six main points:
    1. The initialization method gathers essential information, except for two pieces of information required during the download method call:
        - CLIENT_ID, CLIENT_SECRET, and TENANT_ID: These are from the Azure APP created, necessary for the application to run successfully. They should be set as environment variables with secret encapsulation.
        - SITE_NAME: This variable is set for each call, representing the SharePoint site to be accessed. Without it, the URL requesting (get() method) cannot proceed.
        - COMPANY_TENANT_ID: This variable sets the company tenant ID, similar to the initial part of the SharePoint URL (e.g., your_company.sharepoint.com).

    The class provides five main functions:
    1. get_token: Retrieves the header information required for each get() method.
    2. get_response_id: Searches for the ID for every folder/file level, crucial for constructing the URL to be accessed.
    3. get_drive_id: Retrieves the drive ID information from the first level of folders in the SharePoint library.
    4. find_file: Searches for the desired file within each folder on the pipeline. It avoids looping the same search after finding each level of subfolder.
    5. download_file: Initiates the entire pipeline. This method requires setting two variables:
        - target_file_name: The name of the file, along with its type (e.g., 'file.xlsx' for Excel, 'file.csv' for CSV).
        - folder_match: The folder name to be matched, to find the root folder ID.

    """
    
    def __init__(self, company_tenant_id=None, client_id=None, client_secret=None, tenant_id=None, site_name=None) -> None:
        """
        Initializes the SharepointDownloader object.

        params: 
            - company_tenant_id (str): you may pass it directly to the constructor as a string. 
            - client_id (str): you may pass it directly to the constructor as a string. 
            - client_secret (str): you may pass it directly to the constructor as a string. 
            - tenant_id (str): you may pass it directly to the constructor as a string. 
            - site_name (str): you may pass it directly to the constructor as a string. 
            
        example:
            sharepointdownloader = SharePointDownloader(
                                                        company_tenant_id = 'your_company', 
                                                        client_id = 'you client ID from Azure',
                                                        client_secret = 'you client secret from Azure', 
                                                        tenant_id = 'you tenant ID from Azure',
                                                        site_name = 'The sharepoint site name as shown at the sharepoint's URL',
                                                        ). 
            Alternatively, keep company_tenant_id = None to try to collect it from the environment (or any/all other variables).
        
        """
        # Environment variables loading
        self.COMPANY_TENANT_ID = company_tenant_id or os.environ.get('COMPANY_TENANT_ID')   
        self.CLIENT_ID = client_id or os.environ.get('CLIENT_ID')
        self.CLIENT_SECRET = client_secret or os.environ.get('CLIENT_SECRET')
        self.TENANT_ID = tenant_id or os.environ.get('TENANT_ID')
        self.SITE_NAME = site_name or os.environ.get('SITE_NAME')
            
        # Authentication Config
        self.authority = 'https://login.microsoftonline.com/' + self.TENANT_ID
        self.scope = ['https://graph.microsoft.com/.default']

        # APP Initialization
        self.app = ConfidentialClientApplication(
            self.CLIENT_ID, 
            authority=self.authority,
            client_credential=self.CLIENT_SECRET
        )
        
        # This variable is retrieved after the Env variableS "site name" and "company tenant (like your_company.sharepoint.com)" is set, dynamically searching the sharepoint site ID.
        self.SHAREPOINT_SITE_ID = requests.get(f'https://graph.microsoft.com/v1.0/sites/{self.COMPANY_TENANT_ID}:/sites/{self.SITE_NAME}', headers=self.get_token()).json()['id']
       
        # This url is used at each request, as it is the main request url
        self.url = f'https://graph.microsoft.com/v1.0/sites/{self.SHAREPOINT_SITE_ID}/drives/'
    
    def get_token(self) -> None:
        """
        Retrieves the authorization token required for API requests.

        Returns:
        - dict: A dictionary containing the authorization header.

        """
        try:
            result = self.app.acquire_token_for_client(scopes=self.scope)
            access_token = result['access_token']
            headers = {'Authorization': 'Bearer ' + access_token}
            return headers
        
        except Exception as e:
            raise RuntimeError(f'Erro ao baixar arquivo: {e}')
         
    def get_response_id(self, result_json, folder_match: str):
        """
        Retrieves the ID for the specified folder or file.

        Args:
        - result_json (str): The JSON response containing folder/file information.
        - folder_match (str): The name of the folder/file to be matched.

        Returns:
        - str: The ID of the matching folder/file.

        """
        try:
            # Converting the response into JSON format
            result_json = json.loads(result_json.content)
            
            # Verifying the id for the object
            for item in result_json['value']:
                if item['name'] == folder_match:
                    return item['id']
            return None
        
        except Exception as e:
            raise RuntimeError(f'Erro ao baixar arquivo: {e}')
    
    def get_drive_id(self, folder_match: str):
        """
        Retrieves the root drive and folder IDs.

        Args:
        - folder_match (str): The folder name to be matched.

        Returns:
        - tuple: A tuple containing the drive ID and the root folder ID.

        """
        try:
            headers = self.get_token()
            response = requests.get(self.url, headers=headers)
            drive_id = self.get_response_id(result_json=response, folder_match='Documents')
            drive_response = requests.get(self.url + f'{drive_id}/root/children', headers=headers)
            root_folder_id = self.get_response_id(result_json=drive_response, folder_match=folder_match)
            
            return drive_id, root_folder_id
        
        except Exception as e:
            raise RuntimeError(f'Erro ao baixar arquivo: {e}')

    def find_file(self, drive_id: str, folder_id: str, target_file_name: str, headers):
        """
        Searches for the specified file within the SharePoint folders.

        Args:
        - drive_id (str): The ID of the root folder on the API.
        - folder_id (str): The ID of the folder to search within.
        - target_file_name (str): The name of the file to be searched.
        - headers (dict): The authorization headers for the request.

        Returns:
        - str: The ID of the matching file.

        """
        try:
            # Requests to search the file ID.
            folder_url = f"{self.url}/{drive_id}/items/{folder_id}/children"
            folder_response = requests.get(folder_url, headers=headers)
            folder_result = folder_response.json()
            
            # This loops looks for the Download URL. If do not finds the file in the loop, it looks inside each folder.
            for item in folder_result['value']:
                if '@microsoft.graph.downloadUrl' in item:  # Verifica se é um arquivo
                    if item['name'] == target_file_name:
                        return item['id']
                else:
                    if 'folder' in item:
                        sub_folder_id = item['id']
                        file_id = self.find_file(drive_id, sub_folder_id, target_file_name, headers)
                        if file_id:
                            return file_id
            return None
        
        except Exception as e:
            raise RuntimeError(f'Erro ao baixar arquivo: {e}')

    def download_file(self, target_file_name: str, folder_match: str):
        """
        Downloads the specified file from SharePoint.

        Args:
        - target_file_name (str): The name of the file to download.
        - folder_match (str): The name of the folder to match.

        """
        try:
            # Getting each important ID to download the file
            headers = self.get_token()
            drive_id, root_folder_id = self.get_drive_id(folder_match)
            file_id = self.find_file(drive_id, root_folder_id, target_file_name, headers)
            
            # Requests the Download URL from the API and downloads it within the system set up.
            if file_id:
                file_url = f"{self.url}/{drive_id}/items/{file_id}"
                file_result = requests.get(file_url, headers=headers).json()
                file_download_url = file_result["@microsoft.graph.downloadUrl"]
                urlretrieve(file_download_url, file_result['name'])
                print("Arquivo baixado com sucesso!")
            else:
                print("Arquivo não encontrado.")
                
        except Exception as e:
            raise RuntimeError(f'Erro ao baixar arquivo: {e}')
