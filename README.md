# Cleanup-Desktop
PowerShell script: archives old items from Windows desktop

## SYNOPSIS
  Move all files/folders to their respective monthly folders in -ArchiveFolder based on the LastWriteTime of each
  file/folder - except files/folders with a LastWriteTime newer than the number of days passed in as a Parameter -DaysToKeep.

## DESCRIPTION
  Loop through each file/folder in the current logged on user's desktop where the file isn't a link (.lnk) and where
  the LastWriteTime for the file/folder is older than -DaysToKeep ago.
  Move each valid file/folder to the provided -ArchiveFolder parameter, into a subfolder named to match the year
  and month based on the LastWriteTime of the file/folder.

## EXAMPLES
  `.\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30`

  `.\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude "test images",O365*,*.jpg`

 `.\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude *.jpg,"Office 365",test*`

## OUTPUT
  Host-written output to PowerShell session
  Transcript file created in the archive folder location. These files are automatically purged (with
  confirmation), deleting those older than 90 days.

## SCHEDULE
  Run manually or run as a scheduled task
  
  Run as a scheduled task to suit your archiving, example...
  
  User Account: SYSTEM (Run With Highest Privileges)
  
  Action: `powershell.exe`
  
  Arguments: `-executionpolicy remotesigned -File "C:\Scripts\Cleanup-Desktop\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30`
  
  Arguments: `-executionpolicy remotesigned -File "C:\Scripts\Cleanup-Desktop\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude "test images",O365*,*.jpg`
  
  Arguments: `-executionpolicy remotesigned -File "C:\Scripts\Cleanup-Desktop\Cleanup-Desktop.ps1 -ArchiveFolder "C:\Steve\DesktopArchive" -DaysToKeep 30 -Exclude *.jpg,"Office 365",test*`


