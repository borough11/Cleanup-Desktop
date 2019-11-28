# Cleanup-Desktop
PowerShell script: archives old items from Windows desktop

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
