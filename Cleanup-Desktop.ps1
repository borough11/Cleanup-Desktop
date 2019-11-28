<#
.SYNOPSIS
  Move all files/folders to their respective monthly folders in -ArchiveFolder based on the LastWriteTime of each
  file/folder - except files/folders with a LastWriteTime newer than the number of days passed in as a Parameter -DaysToKeep.

.DESCRIPTION
  Loop through each file/folder in the current logged on user's desktop where the file isn't a link (.lnk) and where
  the LastWriteTime for the file/folder is older than -DaysToKeep ago.
  Move each valid file/folder to the provided -ArchiveFolder parameter, into a subfolder named to match the year
  and month based on the LastWriteTime of the file/folder.

.PARAMETER ArchiveFolder
  Folder location where the yearly and monthly folders will be created, and then desktop files/folders copied into.
  * This folder MUST exist, otherwise the script will throw an error at parameter validation.

.PARAMETER DaysToKeep
  The number of days of files/folders to keep on the desktop based on LastWriteTime

.PARAMETER Exclude
  A list of folder names, file names or file types to exclude from being moved. Exact names or wildcards can be used.
  If there are spaces, place quotations around the string, i.e. -Exclude "test images",O365*,*.jpg will exclude the
  file or folder named 'test images', any file or folder beginning 'O365' and all .JPG files

.EXAMPLE
  .\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30

.EXAMPLE
  .\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude "test images",O365*,*.jpg

.EXAMPLE
  .\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude *.jpg,"Office 365",test*

.OUTPUT
  Host-written output to PowerShell session
  Transcript file created in the archive folder location. These files are automatically purged (with
  confirmation), deleting those older than 90 days.

.SCHEDULE
  Run manually or run as a scheduled task
  Run as a scheduled task to suit your archiving, example...
  User Account: SYSTEM (Run With Highest Privileges)
  Action: powershell.exe
  Arguments: -executionpolicy remotesigned -File "C:\Scripts\Cleanup-Desktop\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30
  Arguments: -executionpolicy remotesigned -File "C:\Scripts\Cleanup-Desktop\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude "test images",O365*,*.jpg
  Arguments: -executionpolicy remotesigned -File "C:\Scripts\Cleanup-Desktop\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude *.jpg,"Office 365",test*


.HISTORY
  v0.1 - Created October 2018
  v0.2 - Change path to write transcript file to the archive folder rather than the script location "$transcriptRootFolder\$($MyInvocation.MyCommand.Name)-Transcript_$(Get-Random).log"

.NOTES
   Author: Steve Geall
     Date: October 2018
  Version: 0.2
  Updated: 18 Nov 2019
#>

[CmdletBinding()]
param (
    # only continue if archive folder path exists
    [Parameter(Mandatory=$True)]
    [ValidateScript({
        If ((Test-Path $_)) {
            $True
        } Else {
            Throw "The value provided for the -ArchiveFolder is invalid, ensure the folder exists!"
        }
    })]
    [String]$ArchiveFolder,

    [Parameter(Mandatory=$True)]
    [Int]$DaysToKeep,

    [String[]]$Exclude
)

Function Move-DesktopItem {
    [CmdletBinding()]
    param (
        $ItemToMove,
        $FileOrFolder
    )

    $destinationPath = "$ArchiveFolder\$($ItemToMove.LastWriteTime.Year)\$((Get-Culture).DateTimeFormat.GetMonthName($ItemToMove.LastWriteTime.Month))"
    If (!(Test-Path -path $destinationPath)) {
        Write-Host "[+] Creating destination subfolder ($destinationPath)" -ForegroundColor Magenta
        New-Item -Path $destinationPath -ItemType Directory | Out-Null
    }

    $script:movedOK = $null
    Try {
        # Write-Host "DEBUG :: Move-Item -LiteralPath $($ItemToMove.FullName) -Destination $destinationPath -ErrorAction Stop -Force"
        Move-Item -LiteralPath "$($ItemToMove.FullName)" -Destination $destinationPath -ErrorAction Stop -Force
        $script:movedOK = $True
    } Catch {
        $ErrorMessage = $_.Exception.Message
        $script:movedOK = $False
    }

    If ($script:movedOK -eq $True) {
        Write-Host "[*]" -ForegroundColor Green -NoNewline
    } ElseIf ($script:movedOK -eq $False) {
        Write-Host "[!]" -ForegroundColor Red -NoNewline
    }
    Write-Host " $FileOrFolder`:" -ForegroundColor White -NoNewline
    If ($script:movedOK -eq $True) {
        Write-Host " $($ItemToMove.Name)" -ForegroundColor Green -NoNewline
    } ElseIf ($script:movedOK -eq $False) {
        Write-Host " $($ItemToMove.Name)" -ForegroundColor Red -NoNewline
    }
    Write-Host " ->" -ForegroundColor Gray -NoNewline
    Write-Host " $destinationPath" -ForegroundColor Magenta
    If ($script:movedOK -eq $False) {
        Write-Host "    (error message: $ErrorMessage)" -ForegroundColor Red
    }

}


Try {
    Stop-Transcript
} Catch {
    # continue
}

$transcriptRootFolder = $ArchiveFolder
Start-Transcript "$transcriptRootFolder\$($MyInvocation.MyCommand.Name)-Transcript_$(Get-Random).log" -Force

Write-Host "`r`nCLEANUP DESKTOP" -ForegroundColor Cyan
Write-Host "---------------" -ForegroundColor White

Write-Host "Will move all old files/folders from the desktop to their respective monthly folders." -ForegroundColor White
Write-Host "Destination archive folder: " -ForegroundColor Gray -NoNewLine
Write-Host "$ArchiveFolder" -ForegroundColor Yellow
Write-Host "Move if older than (days): " -ForegroundColor Gray -NoNewLine
Write-Host "$DaysToKeep" -ForegroundColor Yellow
$Exclude | foreach-object {
    Write-Host "Excluding: " -ForegroundColor Gray -NoNewLine
    Write-Host "$_" -ForegroundColor Yellow
}

$validFiles = Get-ChildItem -Path ([Environment]::GetFolderPath("Desktop")) -Exclude $Exclude |
   Where-Object { ($_ -notlike "*lnk") -And ($_.LastWriteTime -lt (Get-Date).AddDays(-$DaysToKeep)) -And (!$_.PSIsContainer) }
$validFilesCount = $validFiles.Count
$validFolders = Get-ChildItem -Path ([Environment]::GetFolderPath("Desktop")) -Exclude $Exclude |
   Where-Object { ($_ -notlike "*lnk") -And ($_.LastWriteTime -lt (Get-Date).AddDays(-$DaysToKeep)) -And ($_.PSIsContainer) }
$validFoldersCount = $validFolders.Count
$totalItemsToMove = $validFilesCount + $validFoldersCount
Write-Host "Valid files to move: " -ForegroundColor Gray -NoNewLine
Write-Host "$validFilesCount" -ForegroundColor Yellow
Write-Host "Valid folders to move: " -ForegroundColor Gray -NoNewLine
Write-Host "$validFoldersCount" -ForegroundColor Yellow
Write-Host "Total items to move: " -ForegroundColor Gray -NoNewLine
Write-Host "$totalItemsToMove" -ForegroundColor Yellow

Write-Host "`r`nAre you sure you want to continue? [Y or y]"
$confirmation = Read-Host " "
If ($confirmation -eq "y" -Or $confirmation -eq "Y") {
    # continue
} Else {
    Write-Host "`r`nY or y not selected, do not continue with script.`r`n" -ForegroundColor Red
    Try {
        Stop-Transcript
    } Catch {
        # continue
    }
    Exit
}

Write-Host "`r`nBeginning Cleanup-Desktop..." -ForegroundColor Yellow
Write-Host "(only moving files/folders with LastWriteTime older than: $((Get-Date).AddDays(-$DaysToKeep)))`r`n" -ForegroundColor Yellow

$movedCountOK = 0
$movedCountFAIL = 0
$filesMovedOK = 0
$filesMovedFAIL = 0
$foldersMovedOK = 0
$foldersMovedFAIL = 0
$script:movedOK = $False

$validFiles | ForEach-Object {
    Move-DesktopItem -ItemToMove $_ -FileOrFolder "File"
    If ($script:movedOK -eq $True) {
        $filesMovedOK++
        $movedCountOK++
    } Else {
        $filesMovedFAIL++
        $movedCountFAIL++
    }
}

$validFolders | ForEach-Object {
    Move-DesktopItem -ItemToMove $_ -FileOrFolder "Folder"
    If ($script:movedOK -eq $True) {
        $foldersMovedOK++
        $movedCountOK++
    } Else {
        $foldersMovedFAIL++
        $movedCountFAIL++
    }
}

Write-Host "`r`nComplete ($movedCountOK/$totalItemsToMove items moved: $filesMovedOK/$validFilesCount files, $foldersMovedOK/$validFoldersCount folders)." -ForegroundColor Yellow
If ($movedCountFAIL -gt 0) {
    Write-Host "($filesMovedFAIL files and/or $foldersMovedFAIL folders failed!)" -ForegroundColor Red
}
Write-Host "`n"

# purge transcript files older than 90 days
If ([string]::IsNullOrEmpty($transcriptRootFolder)) {
    # if $transcriptRootFolder Null or Blank, then Get-ChildItem -Path "$transcriptRootFolder\*Transcript*.log" will resolve to C:\*Transcript*.log
    Write-Host "Sanity checking `$transcriptRootFolder variable and it's Null or Empty so skip purging of transcript files"
} Else {
    $purgableTranscriptFiles = Get-ChildItem -Path "$transcriptRootFolder\*Transcript*.log" | Where-Object { !$_.PSIsContainer -And $_.CreationTime -lt (Get-Date).AddDays(-90) }
    If ($purgableTranscriptFiles ) {
        Write-Host "Purging old transcript files..."
        $purgableTranscriptFiles | ForEach-Object {
            Write-Host "[-] $_"
        }
        $purgableTranscriptFiles | Remove-Item -Force -Confirm
    }
}
Write-Host "`n"

Try {
    Stop-Transcript
} Catch {
    # continue
}

Start-Sleep -Seconds 10