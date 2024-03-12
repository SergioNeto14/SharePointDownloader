# Sharepoint Downloader Documentation

This documentation provides an overview and usage guide for the `SharepointDownloader` class, which is designed for interacting with Sharepoint data through Microsoft Graph API.

## Overview

The `SharepointDownloader` class serves as a tool for connecting and manipulating dataframes utilizing Sharepoint as the primary data source. It facilitates authentication, retrieval of Sharepoint resources, and downloading files.

## Usage

### Initialization

The class must be initialized with environment variables set for authentication and site information. The following variables are required:

- `CLIENT_ID`: Client ID from the Azure APP.
- `CLIENT_SECRET`: Client Secret from the Azure APP.
- `TENANT_ID`: Tenant ID from the Azure APP.
- `SITE_NAME`: Name of the Sharepoint site. Example: `MySharepointName`
- `COMPANY_TENANT_ID`: Tenant ID for the company, typically the initial part of the Sharepoint URL. Example: `companygroup.sharepoint.com`

### Functions

#### `get_token()`

- **Description**: Generates a token for accessing Microsoft Graph API.
- **Returns**: Dictionary containing the authorization header.

#### `get_response_id(result_json, folder_match)`

- **Description**: Retrieves the ID of a folder or file from the API response.
- **Arguments**:
  - `result_json`: JSON response from the API.
  - `folder_match`: Name of the folder or file to be found.
- **Returns**: ID of the matched folder or file.

#### `get_drive_id(folder_match)`

- **Description**: Retrieves the root drive ID of the Sharepoint library.
- **Arguments**:
  - `folder_match`: Name of the root folder.
- **Returns**: Tuple containing drive ID and root folder ID.

#### `find_file(drive_id, folder_id, target_file_name, headers)`

- **Description**: Finds the ID of a target file within the Sharepoint library.
- **Arguments**:
  - `drive_id`: ID of the root folder.
  - `folder_id`: ID of the current folder.
  - `target_file_name`: Name of the file to be searched.
  - `headers`: Authorization headers for API access.
- **Returns**: ID of the target file.

#### `download_file(target_file_name, folder_match)`

- **Description**: Downloads a file from Sharepoint.
- **Arguments**:
  - `target_file_name`: Name of the file to be downloaded.
  - `folder_match`: Name of the root folder.
- **Returns**: None.

## Example Usage

```python
from sharepoint_downloader import SharepointDownloader

# Initialize SharepointDownloader
downloader = SharepointDownloader()

# Download a file
downloader.download_file(target_file_name="example.xlsx", folder_match="Documents")
