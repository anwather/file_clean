# Test-WordDocumentCleanup.ps1
# This script creates test Word documents and then cleans them in one operation

param (
    [Parameter(Mandatory = $true)]
    [string]$FolderPath,
    
    [Parameter(Mandatory = $false)]
    [int]$NumberOfDocuments = 3,
    
    [Parameter(Mandatory = $false)]
    [switch]$CreateSubfolders = $false
)

# Function to log messages with timestamp
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message" -ForegroundColor Cyan
}

Write-Log "===== Starting Word Document Test Suite ====="

# Step 1: Create a temp folder if the specified path doesn't exist
if (-not (Test-Path -Path $FolderPath -PathType Container)) {
    try {
        New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
        Write-Log "Created test folder: $FolderPath"
    }
    catch {
        Write-Log "Error: Could not create folder: $FolderPath"
        Write-Log "Error details: $_"
        exit 1
    }
}

# Step 2: Create test documents
Write-Log "Creating $NumberOfDocuments test Word document(s) in folder: $FolderPath"

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
        $selection.TypeText("Use this file to test the document cleanup process.")
        
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
            $selection.TypeText("Use this file to test the document cleanup process.")
            
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
    
    Write-Log "Document creation complete. Created $NumberOfDocuments test document(s)."
    
} 
catch {
    Write-Log "An error occurred during document creation: $_"
    exit 1
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

# Step 3: Add a small pause to ensure Word is fully released
Start-Sleep -Seconds 2
Write-Log "Waiting for Word processes to clean up..."

# Step 4: Process/clean the documents
Write-Log "===== Starting to process the test documents ====="
Write-Log "Processing Word documents in folder: $FolderPath"

try {
    # Get all Word documents in the specified path and subfolders
    $wordFiles = Get-ChildItem -Path $FolderPath -Filter "*.doc*" -Recurse -File
    
    if ($wordFiles.Count -eq 0) {
        Write-Log "No Word documents found in the specified folder."
        exit 0
    }
    
    Write-Log "Found $($wordFiles.Count) Word document(s) to process."
    
    # Create Word application object
    $word = New-Object -ComObject Word.Application
    
    # Set Word to be invisible
    $word.Visible = $false
    
    # Process each Word document
    foreach ($file in $wordFiles) {
        try {
            $docPath = $file.FullName
            Write-Log "Processing: $docPath"
            
            # Open the document
            $doc = $word.Documents.Open($docPath)
            
            # Save the document
            $doc.Save()
            
            # Close the document
            $doc.Close()
            
            Write-Log "Document cleaned: $docPath"
            
        } 
        catch {
            Write-Log "Error processing file $($file.FullName): $_"
        }
    }
    
    # Quit Word application
    $word.Quit()
    
    # Release COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    
    Write-Log "Processing complete. Processed $($wordFiles.Count) document(s)."
    
} 
catch {
    Write-Log "An error occurred during document processing: $_"
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

Write-Log "===== Word Document Test Suite Completed ====="
Write-Log "All operations completed successfully on the folder: $FolderPath"
