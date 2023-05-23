#****************************************************************************************************
#
# Version: 17.12.09.2022
#
#********* README *********************************************************************************
#
# The Microsoft Support and Recovery Assistant is governed by the terms applicable to "software" as
# detailed in the Microsoft Services Agreement (https://www.microsoft.com/servicesagreement). If you
# do not agree to these terms, do not use the software.
#
# Sample script that demonstrates the following capabilities:
#
# a. Copy SaRA Enterprise (command-line) files to the client from the specified location
# b. Launch SaraCmd.exe to run the specified scenario with the specified switches and parameters
# c. Gather log files generated during the scenario
# d. Create a consolidated zip file with all of the logs
#
#****************************************************************************************************
#
# Available Scenarios that can be run and their corresponding minimum <required> switches and arguments.
#
# NOTE: There are additional <optional> switches and parameters.
# Please see https://learn.microsoft.com/microsoft-365/troubleshoot/administration/sara-command-line-version
# for complete details. Or, run SaraCmd.exe /?
#
# Outlook Advanced Diagnostic Scan:
# -S ExpertExperienceAdminTask -Script -AcceptEula
#
# Office Uninstall:
# -S OfficeScrubScenario -Script -AcceptEula
#
# Teams Meeting Add-in for Outlook troubleshooter:
# -S TeamsAddinScenario -Script -AcceptEula -CloseOutlook
#
# Office Shared Computer Activation:
# -S OfficeSharedComputerScenario -Script -AcceptEula -CloseOffice
#
# Outlook Calendar Scan using CalCheck:
# -S OutlookCalendarCheckTask -Script -AcceptEula
#
# Office Activation troubleshooter:
# -S OfficeActivationScenario -Script -AcceptEula -CloseOffice
#
# Reset Office Subscription Activation:
# -S ResetOfficeActivation -Script -AcceptEula -CloseOffice
#
#****************************************************************************************************
#
# Configurable Variables
# ======================
#
# (1) Variable Name: SaraCmdSourcePath
# ------------------------------------
#
# You have two options:
# a. Leave the default variable value. By default, the script uses https://aka.ms/SaRA_EnterpriseVersionFiles
# to download and use the latest files from the web
# -or-
# b. If desired, modify $SaraCmdSourcePath to use a local (full) path to reference the SaRA Enterprise files
# folder (files extracted) or .zip file
#
$SaraCmdSourcePath = "https://aka.ms/SaRA_EnterpriseVersionFiles"
#
# (2) Variable Name: SaraScenarioArgument
# ---------------------------------------
#
# - Replace with arguments for your scenario, refer to the list above in 'Available Scenarios...'
# - Check latest SaRA Commandline documentation for updates -- please see https://aka.ms/SaRA_CommandLineVersion
#
# Example: $SaraScenarioArgument = "-S TeamsAddinScenario -Script -AcceptEula -CloseOutlook"
#
$SaraScenarioArgument = "-S OfficeScrubScenario -Script -AcceptEula"
#
# (3) Variable Name: currentTimeStamp
# ------------------------------------
#
# - Default timestamp format used for the file name of the .zip file created by this script
# - Changing the format is optional
#
$currentTimeStamp = Get-Date -Format "yyyy-MMM-dd_HH.mm.ss"
#
# (4) Variable Name: resultsFileName
# -----------------------------------
#
# - Do Not remove <scenario> from this variable (you will break the script)
# - Remove $env:USERNAME if you do not wish to have the username included
# - Changing this variable is optional
#
$resultsFileName = $env:USERNAME + "_<scenario>_$currentTimeStamp.zip"
#
#****************************************************************************************************
#
# ==================
# Begin Main Section ** Nothing below this line <requires> any edits **
# ==================

$currentLocation = $PSScriptRoot
$resultsFilePath = "$currentLocation\$resultsFileName"
$SaraCmdExecutableFolder = "$currentLocation\SaraCMDExecutable"
$SaraCmdExecutablePath = "$SaraCmdExecutableFolder\SaraCMD.exe"
$LocalLogFolder = "$currentLocation\LogFiles"
$scriptLogFile = "$LocalLogFolder\SaraCmd-$currentTimeStamp.txt"
$scriptStartTime = Get-Date

# ------------------------
# Starting Local Functions
# ------------------------

# Create local folders that contain SaRA files and log files
Function Create-LocalFolders
{
New-Item -Path $SaraCmdExecutableFolder -ItemType "directory" -Force | Out-Null
New-Item -Path $LocalLogFolder -ItemType "directory" -Force | Out-Null
}

# Cleanup the local folders that were created by the script
Function Clean-LocalFiles
{
Remove-Item -Path $SaraCmdExecutableFolder -Force -Recurse
Remove-Item -Path $LocalLogFolder -Force -Recurse
}

# Cleanup files created by the script that may remain from a previous run
Function Clean-InitialFiles
{
$targetZipFileLocation = "$currentLocation\SaraCmd.zip"
# Delete local zip file if it exists
if (Test-Path -Path $targetZipFileLocation -PathType Leaf)
{
Remove-Item -Path $targetZipFileLocation -Force | Out-File -FilePath $scriptLogFile -Append
}

if (Test-Path -Path $SaraCmdExecutableFolder)
{
Remove-Item -Path $SaraCmdExecutableFolder -Force -Recurse | Out-File -FilePath $scriptLogFile -Append
}
}

# Copy the SaraCmd Execution folders locally
Function Copy-SaraLocally($SaraCmdSourcePath)
{
$targetZipFileLocation = "$currentLocation\SaraCmd.zip"

Write-Output "Copying Files from $SaraCmdSourcePath" | Out-File -FilePath $scriptLogFile -Append

# if the source starts with https, just download it
if ($SaraCmdSourcePath.StartsWith("http", 'CurrentCultureIgnoreCase'))
{
Write-Output "Getting zip file from web location $SaraCmdSourcePath" | Out-File -FilePath $scriptLogFile -Append

Invoke-WebRequest -URI $SaraCmdSourcePath -OutFile $targetZipFileLocation
$SaraCmdSourcePath = $targetZipFileLocation
}
else
{
if ($SaraCmdSourcePath.EndsWith(".zip", 'CurrentCultureIgnoreCase'))
{
Copy-Item -Path $SaraCmdSourcePath -Destination $targetZipFileLocation -Force | Out-File -FilePath $scriptLogFile -Append
$SaraCmdSourcePath = $targetZipFileLocation
}
}

# if source ends with zip, copy it locally, overwrite if it was downloaded in the previous step
# if the source ends with zip, extract it
if ($SaraCmdSourcePath.EndsWith(".zip", 'CurrentCultureIgnoreCase'))
{
Write-Output "Expanding zip file from $SaraCmdSourcePath" | Out-File -FilePath $scriptLogFile -Append
Expand-Archive -Path $targetZipFileLocation -DestinationPath $SaraCmdExecutableFolder -Force
$SaraCmdSourcePath = $SaraCmdExecutableFolder
}

if($SaraCmdSourcePath -ne $SaraCmdExecutableFolder)
{
# copy files to expected folder
Write-Output "Copying files from $SaraCmdSourcePath" | Out-File -FilePath $scriptLogFile -Append
Copy-Item -Path $SaraCmdSourcePath\* -Destination $SaraCmdExecutableFolder -Recurse -Force
Write-Output "Copied Files To $SaraCmdExecutableFolder" | Out-File -FilePath $scriptLogFile -Append
}

# Delete local zip file if it exists
if (Test-Path -Path $targetZipFileLocation -PathType Leaf)
{
Remove-Item -Path $targetZipFileLocation -Force
}

# Check if the source contains SaraCmd.Exe
if (Test-Path -Path "$SaraCmdExecutableFolder\SaraCmd.exe" -PathType Leaf)
{
return $true
}
else
{
return $false
}

}

# Copies log files from the Sara CommandLine folders into current working folder tree
Function Copy-LogFiles()
{
$SaraLogsRootFolder = $env:LOCALAPPDATA
$saraLogsFolder = "$SaraLogsRootFolder\SaraLogs\Log\"
$SaraUploadLogsFolder = "$SaraLogsRootFolder\SaraLogs\UploadLogs\"
$SaraResultsFolder = "$SaraLogsRootFolder\SaraResults\"

# Copy contents from Sara Logs
Get-ChildItem -Path $saraLogsFolder |
Where-Object {
$_.LastWriteTime `
-gt $scriptStartTime } |
ForEach-Object { $_ | Copy-Item -Destination $LocalLogFolder -Recurse }

# Copy contents from Sara Upload Logs
Get-ChildItem -Path $SaraUploadLogsFolder |
Where-Object {
$_.LastWriteTime `
-gt $scriptStartTime } |
ForEach-Object { $_ | Copy-Item -Destination $LocalLogFolder -Recurse }

# Copy contents from Sara results
Get-ChildItem -Path $SaraResultsFolder |
Where-Object {
$_.LastWriteTime `
-gt $scriptStartTime } |
ForEach-Object { $_ | Copy-Item -Destination $LocalLogFolder -Recurse }
}

# Create Zip file with all logs attached
Function Create-LogArchive()
{
Compress-Archive -Path "$localLogFolder\*" -DestinationPath $resultsFilePath
}

# Checks if elevated execution is needed for this scenario
Function Check-AdminAccess($scenario)
{
$elevationRequired = $false

if ($scenario -in "OfficeActivationScenario", "OfficeScrubScenario", "OfficeSharedComputerScenario", "ResetOfficeActivation")
{
$elevationRequired = $true
}

return $elevationRequired
}

# Checks if the current window is elevated
Function Test-IsAdmin # Function credit to: https://devblogs.microsoft.com/scripting/use-function-to-determine-elevation-of-powershell-console/
{
# Returns True if the script is run from an elevated PowerShell console (Run as administrator)
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal $identity
return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# Download and execute the SaRA scenario
Function Execute-SaraCMD($saraCmdSourcePath, $arguments)
{
$success = $false
$filesCopied = Copy-SaraLocally($saraCmdSourcePath)
if ([bool]::Parse($filesCopied) -ne $true)
{
Write-Host "Could not get Sara CMD File locally, exiting..."
exit
}

Write-Output "Executing sara cmd from $SaraCmdExecutablePath" | Out-File -FilePath $scriptLogFile -Append
Write-Output "With arguments : $arguments" | Out-File -FilePath $scriptLogFile -Append

$scenario = Get-Scenario($arguments)

Write-Host ""
Write-Host ">>> Starting the scenario with the following arguments:"
Write-Host ""
Write-Host " $SaraScenarioArgument"
Write-Host ""
Write-Host ">>> Please wait ..."
Write-Host ""

$processInfo = new-Object System.Diagnostics.ProcessStartInfo($SaraCmdExecutablePath);
$processInfo.Arguments = $arguments # Do NOT modify - These are required parameters for this scenario

if(Check-AdminAccess($scenario) -eq $true)
{
$processInfo.Verb = "RunAs"
}

$processInfo.CreateNoWindow = $true;
$processInfo.UseShellExecute = $false;
$processInfo.RedirectStandardOutput = $true;
$process = [System.Diagnostics.Process]::Start($processInfo);
$process.StandardOutput.ReadToEnd();
$process.WaitForExit();

# https://learn.microsoft.com/microsoft-365/troubleshoot/administration/sara-command-line-version
# See the above article for possible ExitCode values

if($process.HasExited -and ($process.ExitCode -eq 0 -or ($process.ExitCode -eq 80) -or ($scenario="TeamsAddinScenario" -and $process.ExitCode -eq 23) -or ($scenario="OfficeSharedActivationScenario" -and $process.ExitCode -eq 63) -or ($scenario="OfficeActvationScenario" -and $process.ExitCode -eq 36) -or ($scenario="OutlookCalendarCheckTask" -and $process.ExitCode -eq 43) -or ($scenario="ExpertExperienceAdminTask" -and ($process.ExitCode -eq 01 -or $process.ExitCode -eq 02 -or $process.ExitCode -eq 3 -or $process.ExitCode -eq 66 -or $process.ExitCode -eq 67))))
{
$success = $true
}

$process.Dispose();

# Returns True if the scenario's execution PASSED, otherwise False
return $success;
}

# Extract scenario name from arguments
Function Get-Scenario($arguments)
{
$scenario = ""

$args = $arguments.Split("-")

foreach ($arg in $args)
{
if ($arg.StartsWith("S ") -or $arg.StartsWith("s "))
{
$scenario = $arg.Split(" ")[1]
break;
}
}
return $scenario
}
#
# -------------------
# End Local Functions
# -------------------
#
# ------------
# Begin Script
# ------------
#
#Check for an empty $SaraCmdSourcePath variable
if(($SaraCmdSourcePath -eq "") -or ($SaraCmdSourcePath -eq $null))
{

Write-Host ">>>"
Write-Host ">>> A value for `$SaraCmdSourcPath has not be specified in the script."
exit
}
# Check to see if the path is exists
if (-not ($saracmdsourcepath -like "https*") -and -not (Test-Path $SaraCmdSourcePath))
{
Write-Host ">>>"
Write-Host ">>> The path specified for `$SaraCmdSourcePath: '$SaraCmdSourcePath' does not exist."
Write-Host ">>>"
Write-Host ">>> Please check the specified path and update `$SaraCmdSourcePath to point to a valid path."
Write-Host ">>>"
exit
}
# Check to see if https path is correct
if ($saracmdsourcepath -like "https*" -and -not($saracmdsourcepath -eq "https://aka.ms/SaRA_EnterpriseVersionFiles"))
{
Write-Host ">>>"
Write-Host ">>> https URL used, but the path ($SaraCmdSourcePath) specified for `$SaraCmdSourcePath is not correct."
Write-Host ">>>"
Write-Host ">>> Please update `$SaraCmdSourcePath to 'https://aka.ms/SaRA_EnterpriseVersionFiles'."
Write-Host ">>>"
exit
}
# Check for an empty $SaraScenarioArgument variable
if (($SaraScenarioArgument -eq "") -or ($SaraScenarioArgument -eq $null))
{
Write-Host ">>>"
Write-Host ">>> `$SaraScenarioArgument is blank"
Write-Host ">>>"
Write-Host ">>> `$SaraScenarioArgument = $SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> Please refer to the Configurable Variables section of the script for the `$SaraScenarioArgument variable"
exit
}

# Check for an empty $currentTimeStamp variable
if (($currentTimeStamp -eq "") -or ($currentTimeStamp -eq $null))
{
Write-Host ">>>"
Write-Host ">>> `$currentTimeStamp is blank"
Write-Host ">>>"
Write-Host ">>> `$currentTimeStamp = $currentTimeStamp"
Write-Host ">>>"
Write-Host ">>> Please refer to the Configurable Variables section of the script for the `$currentTimeStamp variable"
exit
}
# Check for an empty $resultsFileName variable
if (($resultsFileName -eq "") -or ($resultsFileName -eq $null))
{
Write-Host ">>>"
Write-Host ">>> `$resultsFileName is blank"
Write-Host ">>>"
Write-Host ">>> `$resultsFileName = $resultsFileName"
Write-Host ">>>"
Write-Host ">>> Please refer to the Configurable Variables section of the script for the `$resultsFileName variable"
exit
}
# Check for the existence and spelling of the required -AcceptEula switch
if ($SaraScenarioArgument -notlike "*-accepteula*")
{
Write-Host ">>>"
Write-Host ">>> Required switch -AcceptEula missing or misspelled in `$SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> `$SaraScenarioArgument = $SaraScenarioArgument"
exit
}
# Check for the existence and spelling of the required -Script switch
if ($SaraScenarioArgument -notlike "*-script*")
{
Write-Host ">>>"
Write-Host ">>> Required switch -Script missing or misspelled in `$SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> `$SaraScenarioArgument = $SaraScenarioArgument"
exit
}
# Check for the required -S switch
if ($SaraScenarioArgument -notlike "*-s *")
{
Write-Host ">>>"
Write-Host ">>> Required switch -S missing in `$SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> `$SaraScenarioArgument = $SaraScenarioArgument"
exit
}

$scenario = Get-Scenario($SaraScenarioArgument)

# Check to ensure specified scenario name exists and is spelled correctly
if ($scenario -notin "ExpertExperienceAdminTask", "OfficeActivationScenario", "OfficeScrubScenario", "TeamsAddinScenario", "OutlookCalendarCheckTask", "OfficeSharedComputerScenario", "ResetOfficeActivation")
{
Write-Host ">>>"
Write-Host ">>> The scenario name used for the -S switch in `$SaraScenarioArgument is not valid."
Write-Host ">>>"
Write-Host ">>> $SaraScenarioArgument = $SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> Valid scenario names are: "
Write-Host ">>>"
Write-Host ">>> ExpertExperienceAdminTask, OfficeActivationScenario, OfficeScrubScenario, TeamsAddinScenario, "
Write-Host ">>> OutlookCalendarCheckTask, OfficeSharedComputerScenario, ResetOfficeActivation"
Write-Host ">>>"
write-Host ">>> See https://learn.microsoft.com/microsoft-365/troubleshoot/administration/sara-command-line-version for details."
exit
}

#
# Ensure the minimum required switches and parameters were used for the specified scenario
#
switch ($scenario)
{
ExpertExperienceAdminTask
{
# The required switches for this scenario are -S, -Script and -AcceptEula and there's a check for them elsewhere
}
OfficeActivationScenario
{
# Check for required -CloseOffice switch
if ($SaraScenarioArgument -notlike "*closeoffice*")
{
Write-Host ">>>"
Write-Host ">>> You specified the following switches and parameters:"
Write-Host ">>>"
Write-Host ">>> $SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> The $scenario scenario requires the -CloseOffice switch"
Write-Host ">>>"
Write-Host ">>> Please see https://learn.microsoft.com/microsoft-365/troubleshoot/administration/assistant-office-activation for complete details"
exit
}
}
OfficeScrubScenario
{
# The required switches for this scenario are -S, -Script and -AcceptEula and there's a check for them elsewhere
}
TeamsAddinScenario
{
# Check for required -CloseOutlook switch
if ($SaraScenarioArgument -notlike "*closeoutlook*")
{
Write-Host ">>>"
Write-Host ">>> You specified the following switches and parameters:"
Write-Host ">>>"
Write-Host ">>> $SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> The $scenario scenario requires the -CloseOutlook switch"
Write-Host ">>>"
Write-Host ">>> Please see https://learn.microsoft.com/microsoft-365/troubleshoot/administration/assistant-teams-meeting-add-in-outlook for complete details"
exit
}
}
OutlookCalendarCheckTask
{
# The required switches for this scenario are -S, -Script and -AcceptEula and there's a check for them elsewhere
}
OfficeSharedComputerScenario
{
# Check for required -CloseOffice switch
if ($SaraScenarioArgument -notlike "*closeoffice*")
{
Write-Host ">>>"
Write-Host ">>> You specified the following switches and parameters:"
Write-Host ">>>"
Write-Host ">>> $SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> The $scenario scenario requires the -CloseOffice switch"
Write-Host ">>>"
Write-Host ">>> Please see https://learn.microsoft.com/microsoft-365/troubleshoot/administration/assistant-office-shared-computer-activation for complete details"
exit
}
}
ResetOfficeActivation
{
# Check for required -CloseOffice switch
if ($SaraScenarioArgument -notlike "*closeoffice*")
{
Write-Host ">>>"
Write-Host ">>> You specified the following switches and parameters:"
Write-Host ">>>"
Write-Host ">>> $SaraScenarioArgument"
Write-Host ">>>"
Write-Host ">>> The $scenario scenario requires the -CloseOffice switch"
Write-Host ">>>"
Write-Host ">>> Please see https://learn.microsoft.com/microsoft-365/troubleshoot/administration/assistant-reset-office-activation for complete details"
exit
}
}

}

try
{
Clean-InitialFiles
Create-LocalFolders
Write-Output "--------------------------------------------" | Out-File -FilePath $scriptLogFile # First log statement to create the file
}
catch
{
Write-Host ">>> Unable to create the local log file folders. You may not have permissions to write into this folder."
Write-Host ">>>"
Write-Host ">>> Execute this script in a different folder."
exit
}

$elevationNeeded = Check-AdminAccess($scenario)

if (($elevationNeeded -ne $true) -or (($elevationNeeded -eq $true) -and (Test-IsAdmin -eq $true)))
{
$resultsFileName = $resultsFileName.Replace("<scenario>", $scenario)
$resultsFilePath = $resultsFilePath.Replace("<scenario>", $scenario)

$executionSuccess = Execute-SaraCMD $SaraCmdSourcePath $SaraScenarioArgument

Write-Host ">>> SaraCmd.exe output"
Write-Host ""
Write-Host "SaRA Command Line script execution status: $executionSuccess"
Write-Host ""

Write-Output "SaRA Command Line script execution status: $executionSuccess" | Out-File -FilePath $scriptLogFile -Append
Write-Output "" | Out-File -FilePath $scriptLogFile -Append

if($executionSuccess -eq $true)
{
Write-Output ">>> Scenario execution completed successfully" | Out-File -FilePath $scriptLogFile -Append
Write-Host ">>> Scenario execution completed successfully"
if (Test-Path -Path 'C:\Support') {
} else {
    New-Item -Path 'C:\Support\' -ItemType Directory
}

#Copy Installer
Invoke-WebRequest -uri "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365ProPlusRetail&platform=Def&language=en-us&TaxRegion=pr&correlationId=3dda7fc1-09ce-43fe-a9e3-8e6b7ea16401&token=0642b10e-77b5-44d4-a9df-982c54186524&version=O16GA&source=O15OLSO365&Br=2" -OutFile C:\Support\OfficeSetup.exe

New-Item C:\Support\OfficeConfig.xml -ItemType File
Set-Content C:\support\OfficeConfig.xml @"
<Configuration ID="6de936b1-0302-4bba-8810-b647b8ed37c7" DeploymentConfigurationID="00000000-0000-0000-0000-000000000000">
  <Add OfficeClientEdition="64" Channel="Monthly" ForceUpgrade="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
Set-Location C:\Support\
Start-Process .\OfficeSetup.exe -Wait -ArgumentList '/configure OfficeConfig.xml'
}
else
{
Write-Output ">>> SaRA Commandline ran into a problem or had an error. Please check the SaraLog-<date>.log file for details." | Out-File -FilePath $scriptLogFile -Append
Write-Host ">>> SaRA Commandline ran into a problem or had an error. Please check the SaraLog-<date>.log files for details."
Write-Host ""
}

Copy-LogFiles
Create-LogArchive
Write-Output ">>> All Generated Logs are found at: $resultsFilePath"
}
else
{
Write-Host ""
Write-Host ">>> $scenario needs to be run with elevated privileges (Run As Administrator)"
Write-Host ">>> Execute this script from a new PowerShell window using 'Run As Administrator'"
Write-Host ""
}

Clean-LocalFiles
#
# ----------
# End script
# ----------
#
# ================
# End Main section
# ================
#