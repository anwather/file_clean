# Word Document Processing Toolkit

This toolkit provides PowerShell scripts for managing and processing Microsoft Word documents in bulk.

## Overview

The toolkit consists of three PowerShell scripts:

1. **Clean-WordDocuments.ps1** - Opens, saves, and closes Word documents to refresh them
2. **Create-TestWordDocuments.ps1** - Creates test Word documents for validation testing
3. **Test-WordDocumentCleanup.ps1** - An all-in-one script that creates and processes documents

These scripts are particularly useful for:

- Refreshing Word documents to apply template updates
- Applying sensitivity labels to documents in bulk
- Testing document processing workflows
- Creating test document sets for validation

## Prerequisites

- Windows operating system
- Microsoft Word installed
- PowerShell 5.1 or higher

## Script Descriptions

### Clean-WordDocuments.ps1

This script processes existing Word documents by opening each document, saving it, and then closing it. It's particularly useful for refreshing documents to apply template changes or for applying sensitivity labels to documents in bulk.

#### Usage

```powershell
.\Clean-WordDocuments.ps1 -FolderPath "C:\Path\To\Documents"
```

#### Parameters

- `-FolderPath` (Required): The path to the folder containing Word documents to process

#### Features

- Recursively processes all .doc and .docx files in the specified folder
- Applies sensitivity labels to documents (configured in script)
- Provides detailed logging with timestamps
- Handles errors gracefully and continues processing
- Properly cleans up COM objects to prevent memory leaks

### Create-TestWordDocuments.ps1

This script creates blank Word documents for testing purposes. It can create documents in both the main folder and in subfolders.

#### Usage

```powershell
.\Create-TestWordDocuments.ps1 -FolderPath "C:\Path\To\Output" -NumberOfDocuments 5 -CreateSubfolders
```

#### Parameters

- `-FolderPath` (Required): The path where test documents will be created
- `-NumberOfDocuments` (Optional): Number of documents to create (default: 5)
- `-CreateSubfolders` (Optional): Switch to create documents in subfolders as well

#### Features

- Creates a specified number of Word documents with minimal content
- Option to distribute documents between main folder and a subfolder
- Creates the target folder if it doesn't exist
- Provides detailed logging with timestamps

### Test-WordDocumentCleanup.ps1

This is an all-in-one script that combines the functionality of both creating test documents and processing them in one operation. Useful for end-to-end testing of document processing workflows.

#### Usage

```powershell
.\Test-WordDocumentCleanup.ps1 -FolderPath "C:\Path\To\TestArea" -NumberOfDocuments 3 -CreateSubfolders
```

#### Parameters

- `-FolderPath` (Required): The path where test documents will be created and processed
- `-NumberOfDocuments` (Optional): Number of documents to create (default: 3)
- `-CreateSubfolders` (Optional): Switch to create documents in subfolders as well

#### Features

- Creates test documents with minimal content
- Option to distribute documents between main folder and a subfolder
- Processes all created documents in a single operation
- Includes pausing between operations to ensure proper cleanup
- Uses colored console output for better visibility
- Provides comprehensive logging throughout the process

## Example Workflows

### Refreshing Existing Documents

To refresh all Word documents in a specific folder:

```powershell
.\Clean-WordDocuments.ps1 -FolderPath "C:\Company\Documents"
```

### Creating and Testing with Sample Documents

To create a set of test documents and then process them:

```powershell
# First create test documents
.\Create-TestWordDocuments.ps1 -FolderPath "C:\Temp\TestDocs" -NumberOfDocuments 10 -CreateSubfolders

# Then process them
.\Clean-WordDocuments.ps1 -FolderPath "C:\Temp\TestDocs"
```

### End-to-End Testing

To create and process documents in one command:

```powershell
.\Test-WordDocumentCleanup.ps1 -FolderPath "C:\Temp\TestArea" -NumberOfDocuments 5 -CreateSubfolders
```

### Adding Additional Processing

You can extend the document processing by adding additional operations after the document is opened and before it is saved. For example, to update document properties:

```powershell
$doc = $word.Documents.Open($docPath)

# Add custom processing here
$doc.BuiltInDocumentProperties("Title").Value = "New Title"
$doc.BuiltInDocumentProperties("Company").Value = "My Company"

$doc.Save()
```

## Troubleshooting

### Common Issues

1. **"Cannot convert value of type 'psobject' to type 'Object'"**
   - This is related to how PowerShell handles COM object parameters
   - The scripts use `[string]$fileName` to properly cast the string parameter

2. **Word process remains in memory**
   - The scripts include COM object cleanup and garbage collection
   - If Word processes remain, you may need to manually end them in Task Manager

3. **Permission denied errors**
   - Ensure you have write permissions to the document folders
   - Check if documents are currently open in Word or locked by another process

## License

This project is available under the MIT License.

## Last Updated

May 30, 2025
