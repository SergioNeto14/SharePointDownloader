import json
import os
import requests
from urllib.request import urlretrieve
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv()

class SharepointDownloader(company_tenant_id = None):
    """
    This class is a reusable tool for connection in order to manipulate dataframes which uses Sharepoint as the main source.
    
    The behavior of the class has 6 main points:
    - The init method gather all important information needed apart 2 information that must be placed at the download method call:
        - CLIENT_ID, CLIENT_SECRET and TENANT_ID: from the Azure APP that was created, this information is required for the application to run sucessfully.
        As they are relevant information, must be set as environment variables, with secrets encapsulation.
        - SITE_NAME: this variable is set to every call, being replaced in every different sharepoint site you might call.
        Without the variable above, the URL requesting (get() method) is not able to run any call from the API.
        - COMPANY_TENANT_ID: this variable is to set the company tenant id, which is similar to the initial part of the sharepoint url.
        You can identify it in your url address as: your_company.sharepoint.com
    
    We have 5 main functions:
    - get_token: brings the header information to be run at every get() method
    
    - get_response_id: searchs the ID for every folder/file level, as an important part of the url to be run sucessfully. 
    The arguments, explained at each method level, are self set by any call.
    
    - get_drive_id: gets the drive id information from the first level of folders. It is important to search on the files library of each sharepoint site.
    
    - find_file: This function is able to search for the desired file in each folder on the pipeline. 
    There's a searcher that avoids looping the same seek after finding each level of subfolder.
    
    - download_file: This method is the only that must be called to run the entire pipeline. 
    Also, here's the one place you must set 2 variables, completing the mentioned 6 main parts (see above):
        - target_file_name: the name of the file, and also its type (e.g.: if excel, must contains 'file.xlsx', if csv, 'file.csv')
        - folder_match: folder name to be matched, in order to find the root folder ID.
    """
    def __init__(self, company_tenant_id) -> None:
        
        if company_tenant_id: # run if it is not None
            self.COMPANY_TENANT_ID = company_tenant_id
        
        else:
            self.COMPANY_TENANT_ID = os.environ.get('COMPANY_TENANT_ID')        
        
        # Environment variables loading
        self.CLIENT_ID = os.environ.get('CLIENT_ID')
        self.CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
        self.TENANT_ID = os.environ.get('TENANT_ID')
        self.SITE_NAME = os.environ.get('SITE_NAME')
        # self.COMPANY_TENANT_ID = os.environ.get('COMPANY_TENANT_ID')
            
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
        This function aims to create a 'headers' object to be used at each get request on the Graph API.
        The return is a json/dictionary object to be read by the headers argument, as it can be seen in the functions below.
        """
        try:
            result = self.app.acquire_token_for_client(scopes=self.scope)
            access_token = result['access_token']
            headers = {'Authorization': 'Bearer ' + access_token}
            return headers
        
        except Exception as e:
            print(f'Error: {e}')
         
    def get_response_id(self, result_json, folder_match: str):
        """
            This function is responsible to read all API response and return the folder/file IDs.
            It converts the response from get() method into a JSON format and gets the intended ID by matching the subsequent name on arg.
            
            The arguments of this function are:
            - result_json: inputs the obtject as type requests.models.Response, which is converted to json and read by the function
            - folder_match: string of folder/file to be found by the function. 
            
            If the name is not match as required, the return is none.
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
            print(f'Error: {e}')
    
    def get_drive_id(self, folder_match: str):
        """
        This functions intends to create the root connection to the sharepoint library.
        It allows the user to search the desired document by finding the root folder over the sharepoint files and folders available.
        
        The arguments of this function are:
        - match: the folder that will indicate where the file must be located, within any levels that may have.
        """
        try:
            headers = self.get_token()
            response = requests.get(self.url, headers=headers)
            drive_id = self.get_response_id(result_json=response, folder_match='Documents')
            drive_response = requests.get(self.url + f'{drive_id}/root/children', headers=headers)
            root_folder_id = self.get_response_id(result_json=drive_response, folder_match=folder_match)
            
            return drive_id, root_folder_id
        
        except Exception as e:
            print(f'Error: {e}')

    def find_file(self, drive_id: str, folder_id: str, target_file_name: str, headers):
        """
            find_file is a API scraping function to find the ID of the target file on sharepoint.
            
            The arguments of this function are:
            - drive_id: string, the ID of the root folder on the API.
            - folder_id: string, uses the IDs of the folders from the root folder to search the target file.
            - target_file_name: string, the name of file to be searched. If found, stops the search.
            - headers: JSON, the authorization to run the get() request method.
            
            If the name is not match as required, the return is none.
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
            print(f'Error: {e}')

    def download_file(self, target_file_name: str, folder_match: str):
        """
            The function is able to read and download the sharepoint hosted file from the result given the file ID obtained by the method find_file()
            
            The arguments of this function are:
            - target_file_name: string, the name of file to be searched. If found, stops the search.
            - match: match the root folder of the application, after the documents folder is already accessed.
            
            If the name is not match as required, the return is none.
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
            print(f'Error: {e}')
    
sharepoint_downloader = SharepointDownloader(company_tenant_id=os.environ.get('COMPANY_TENANT_ID')   )
sharepoint_downloader.download_file(target_file_name="DataLAB Line Inspector.xlsx", folder_match='SPC Barranquilla')