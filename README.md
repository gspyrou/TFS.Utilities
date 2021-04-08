# TFS.Utilities
## Introduction
The utility uses an input the Test cases excel file as exported from VSTS and it appends the Description column stripped from HTML tags
## Configuration
In order to configure communication with VSTS the following settings need to be modified at the app.config file.
### Personal Access Token
Gerenate a Personal Access Token (navigate to User progile -> Security -> Personal Access Token -> Add) and use this value for the appSetting with key "PersonalAccessToken" 
### REST API base Url
For the appSetting with key "RestApiBaseUri" replace the value with https://{server:port}/tfs/{collection}
## Usage
- Generate an excel file from VSTS for a test plan and donwload it locally
- Build and run the TestCasesExport project
- Pass the full path of the Test Cases exported Excel file 
- The tool will strip HTML tags and append the "System.Description" column to excel file
