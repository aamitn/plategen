<#
.SYNOPSIS
Deletes specified directories (build, dist, node_modules) and files (*.spec)
recursively from the script's location.

.DESCRIPTION
This script first sets its own directory as the working directory.
It then locates and deletes all specified directories (build, dist, node_modules)
and files (*.spec) within that directory and its subdirectories.

.PARAMETER WhatIf
Runs the script in "WhatIf" mode, showing only what WOULD be deleted
without performing the actual deletion. It is highly recommended to
run the script with -WhatIf first.

.EXAMPLE
# Recommended first run (safe mode)
.\cleanup-project-folders.ps1 -WhatIf

.EXAMPLE
# Run the script to perform the deletion
.\cleanup-project-folders.ps1
#>
param(
    [switch]$WhatIf
)

# Define the names of the directories to be purged
$FoldersToPurge = @("build", "dist", "node_modules")

# Define the file patterns to be purged (e.g., test files)
$FilesToPurge = @("*.spec")

Write-Host "--- Starting project cleanup ---"

try {
    # Store the original location
    $OriginalLocation = Get-Location
    Write-Host "Initial Directory: $($OriginalLocation.Path)"

    # Set the working directory to the script's root path ($PSScriptRoot) for predictable cleanup.
    # $PSScriptRoot is only available when running a script file.
    if ($PSScriptRoot) {
        Set-Location -Path $PSScriptRoot
        Write-Host "Working directory successfully set to script root: $(Get-Location)" -ForegroundColor Cyan
    } else {
        Write-Host "Warning: \$PSScriptRoot not found. Continuing with initial directory: $(Get-Location)" -ForegroundColor DarkYellow
    }

    Write-Host "Folders to Delete: $($FoldersToPurge -join ', ')"
    Write-Host "Files to Delete: $($FilesToPurge -join ', ')"
    Write-Host ""
    
    # Initialize counters
    $DirCount = 0
    $FileCount = 0

    # --- 1. CLEAN UP DIRECTORIES ---
    Write-Host "--- Cleaning up directories ---" -ForegroundColor Green

    # Get-ChildItem finds all matching directories recursively.
    $DirsToDelete = Get-ChildItem -Path (Get-Location) -Recurse -Directory -Include $FoldersToPurge -ErrorAction SilentlyContinue

    if ($DirsToDelete) {
        $DirCount = $DirsToDelete.Count
        Write-Host "Found $($DirCount) directories to purge." -ForegroundColor Green
        
        # Use Remove-Item to delete the found directories.
        $DirsToDelete | ForEach-Object {
            Write-Host "Deleting Directory: $($_.FullName)" -ForegroundColor Yellow
            Remove-Item -Path $_.FullName -Recurse -Force -WhatIf:$WhatIf
        }
    } else {
        Write-Host "No directories matching target names found."
    }

    # --- 2. CLEAN UP FILES ---
    Write-Host ""
    Write-Host "--- Cleaning up files ---" -ForegroundColor Green

    # Get-ChildItem finds all matching files recursively. Use -File flag.
    $FilesFound = Get-ChildItem -Path (Get-Location) -Recurse -File -Include $FilesToPurge -ErrorAction SilentlyContinue

    if ($FilesFound) {
        $FileCount = $FilesFound.Count
        Write-Host "Found $($FileCount) files to purge." -ForegroundColor Green

        # Use Remove-Item to delete the found files.
        $FilesFound | ForEach-Object {
            Write-Host "Deleting File: $($_.FullName)" -ForegroundColor Yellow
            # Note: No -Recurse needed for files.
            Remove-Item -Path $_.FullName -Force -WhatIf:$WhatIf
        }
    } else {
        Write-Host "No files matching '*.spec' found."
    }

    # --- 3. FINAL REPORT ---
    $TotalCount = $DirCount + $FileCount
    Write-Host ""
    
    if ($TotalCount -eq 0) {
        Write-Host "Nothing to clean up. Script finished." -ForegroundColor Green
    } elseif ($WhatIf) {
        Write-Host "Cleanup simulation complete ($TotalCount total items). Re-run without '-WhatIf' to perform the actual deletion." -ForegroundColor Green
    } else {
        Write-Host "Cleanup complete! ($TotalCount total items purged)." -ForegroundColor Green
    }

} catch {
    Write-Error "An error occurred during cleanup: $($_.Exception.Message)"
    Write-Host "Operation failed. Check permissions and try running PowerShell as Administrator if the issue persists." -ForegroundColor Red
}

Write-Host "--------------------------------"