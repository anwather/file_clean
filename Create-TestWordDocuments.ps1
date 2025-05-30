# Create-TestWordDocuments.ps1
# This script creates blank Word documents for testing purposes

param (
    [Parameter(Mandatory = $true)]
    [string]$FolderPath,
    
    [Parameter(Mandatory = $false)]
    [int]$NumberOfDocuments = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$CreateSubfolders = $false
)

# Function to log messages with timestamp
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

# Validate input path exists or create it
if (-not (Test-Path -Path $FolderPath -PathType Container)) {
    try {
        New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
        Write-Log "Created folder: $FolderPath"
    }
    catch {
        Write-Log "Error: Could not create folder: $FolderPath"
        Write-Log "Error details: $_"
        exit 1
    }
}

Write-Log "Starting creation of $NumberOfDocuments test Word document(s) in folder: $FolderPath"

try {
    # Create Word application object
    $word = New-Object -ComObject Word.Application
    
    # Set Word to be invisible 
    $word.Visible = $false
    
    # Create documents in main folder
    $mainFolderCount = if ($CreateSubfolders) { [Math]::Ceiling($NumberOfDocuments / 2) } else { $NumberOfDocuments }
    
    for ($i = 1; $i -le $mainFolderCount; $i++) {
        $fileName = Join-Path -Path $FolderPath -ChildPath "TestDoc_$i.docx"
        Write-Log "Creating document: $fileName"
        # Create a new document
        $doc = $word.Documents.Add()
        
        # Add some minimal content
        $selection = $word.Selection
        $selection.TypeText("This is test document $i created on $(Get-Date)")
        $selection.TypeParagraph()
        $selection.TypeText("Use this file to test the Clean-WordDocuments.ps1 script.")
        
        # Save the document
        $doc.SaveAs([string]$fileName)
        
        # Close the document
        $doc.Close()
        
        Write-Log "Document created: $fileName"
    }
    
    # Create documents in subfolders if requested
    if ($CreateSubfolders) {
        $subfolder = Join-Path -Path $FolderPath -ChildPath "Subfolder"
        
        if (-not (Test-Path -Path $subfolder -PathType Container)) {
            New-Item -Path $subfolder -ItemType Directory -Force | Out-Null
            Write-Log "Created subfolder: $subfolder"
        }
        
        $subFolderCount = $NumberOfDocuments - $mainFolderCount
        
        for ($i = 1; $i -le $subFolderCount; $i++) {
            $fileName = Join-Path -Path $subfolder -ChildPath "SubfolderDoc_$i.docx"
            Write-Log "Creating document: $fileName"
            
            # Create a new document
            $doc = $word.Documents.Add()
            # Add some minimal content
            $selection = $word.Selection
            $selection.TypeText("This is a subfolder test document $i created on $(Get-Date)")
            $selection.TypeParagraph()
            $selection.TypeText("Use this file to test the Clean-WordDocuments.ps1 script.")
            
            # Save the document
            $doc.SaveAs([string]$fileName)
            
            # Close the document
            $doc.Close()
            
            Write-Log "Document created: $fileName"
        }
    }
    
    # Quit Word application
    $word.Quit()
    
    # Release COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    
    Write-Log "Creation complete. Created $NumberOfDocuments test document(s)."
    
}
catch {
    Write-Log "An error occurred: $_"
}
finally {
    # Make sure Word is closed and COM objects are released
    if ($null -ne $word) {
        try {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        catch {
            # Ignore errors in cleanup
        }
    }
    
    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Print instructions for next steps
Write-Log ""
Write-Log "Test documents created successfully. To test the Clean-WordDocuments.ps1 script, run:"
Write-Log ".\Clean-WordDocuments.ps1 -FolderPath `"$FolderPath`""
