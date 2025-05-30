# Clean-WordDocuments.ps1
# This script opens all Word documents in a specified path, saves them, and then closes them

param (
    [Parameter(Mandatory = $true)]
    [string]$FolderPath
)

# Function to log messages with timestamp
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

# Validate input path exists
if (-not (Test-Path -Path $FolderPath -PathType Container)) {
    Write-Log "Error: The specified folder path does not exist: $FolderPath"
    exit 1
}

Write-Log "Starting Word document processing in folder: $FolderPath"

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
    
    # Set Word to be invisible - comment out this line if you want to see Word working
    $word.Visible = $false
    
    # Process each Word document
    foreach ($file in $wordFiles) {
        try {
            $docPath = $file.FullName
            Write-Log "Processing: $docPath"
            
            # Open the document
            $doc = $word.Documents.Open($docPath)
            Write-Log "Current sensitivity label: $($doc.SensitivityLabel.GetLabel().LabelId)"

            #Update the sensitivity label if needed
            $doc.SensitivityLabel.SetLabel("87867195-f2b8-4ac2-b0b6-6bb73cb33afc") # Change "General" to the desired label

            # Save the document
            $doc.Save()
            Write-Log "Document saved: $docPath"
            
            # Close the document
            $doc.Close()
            
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
