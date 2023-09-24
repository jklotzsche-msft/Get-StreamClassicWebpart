# Get-StreamClassicWebpart

Identify all sharepoint sites containing a stream classic webpart

## Description

This script will identify all sharepoint sites containing a stream classic webpart. It will also identify the page, the embeded url to the video and the owner of the site.

## Prerequisites

1. Create a Azure Function App in your Azure Subscription.

2. Enable the "system-assigned managed identity" of the Azure Function App.

3. Create a Azure Storage Account in your Azure Subscription.

4. Create a Azure Container in the Azure Storage Account.

5. Assign the "Storage Account Contributor" role to the "system-assigned managed identity" of the Azure Function App on the Azure Container.

6. Assign the "Sites.Read.All" Application Permission to the "system-assigned managed identity" of the Azure Function App on the Microsoft Graph API.

7. Upload all files from the [function Folder](./function) to a Azure Function App in your Azure Subscription or replace the existing files with the content of the [function Folder](./function). If you replace the existing files, create a new timer trigger function and copy the content of the [function\GetStreamClassicWebpart](./function/GetStreamClassicWebpart) folder to the run.ps1 and function.json file of the new timer trigger function.

## Usage

1. Run the function to get all sites containing a stream classic webpart. Please check the comment-based help section of the "run.ps1" file for more details!
2. Download the csv file(s) from the Azure Container.
3. Merge the csv file(s) to one csv file using the [Merge-CsvFiles.ps1](./Merge-CsvFiles.ps1) script.
4. Import the csv file to Excel and filter the data.
