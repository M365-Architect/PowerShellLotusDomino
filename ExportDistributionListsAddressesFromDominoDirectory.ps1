#========================================================================
# Created on:		06.05.2016 10:51
# Created by:		Andreas Hähnel
# Organization:		infoWAN Datenkommunikation GmbH
# Filename:			ExportDistributionListsAddressesFromDominoDirectory.ps1
# Script Version: 	0.1
#========================================================================
#
# Changelog:
# Version 0.1 06.05.2016
# - initial creation
#========================================================================


#========================================================================
# Script parameters
#========================================================================
#region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory=$true)][String]$DominoServer,
	[Parameter(Mandatory=$true)][String]$CSVOutFilePath,
	[Parameter(Mandatory=$true)][String]$CSVMultiValueDelimiter
	#[Parameter(Mandatory=$true)][String]$NotesIDFilePath
)

#endregion


#========================================================================
# Global Variables
#========================================================================
#region global variables

#endregion


#========================================================================
# Functions
#========================================================================
#region functions
#========================================================================
<#
	.SYNOPSIS
		Gets the currents timestamp in Format "yyyyMMddhhmm"

	.DESCRIPTION
		Gets the currents timestamp in Format "yyyyMMddhhmm"
	
	.OUTPUTS
		System.String

#>
function Get-TimestampyyyyMMdd
{
	#returns a padded timestamp string like 10.02.2014 17:02 
	$now = Get-Date
	$year = $now.Year.ToString()
	$month = $now.Month.ToString()
	$day = $now.Day.ToString()
	$hour = $now.Hour.ToString()
	$minute = $now.Minute.ToString()
	$second = $now.Second.ToString()

#region make sure all numbers have 2 digits	
	if ($month.length -lt 2) { $month = "0" + $month }
	if ($day.length -lt 2) { $day = "0" + $day }
	if ($hour.length -lt 2) { $hour = "0" + $hour }
	if ($minute.length -lt 2) {	$minute = "0" + $minute }
	if ($second.lenth -lt 2) { $second = "0" + $second }
#endregion
	
	#write-output $yr$mo$dy$hr$mi
	Write-Output ($year + $month + $day) #+ $hour + $minute)
}

#========================================================================
<#
	.SYNOPSIS
		Returns the UNC path in which the script is running

	.DESCRIPTION
		Returns the UNC path in which the script is running

	.EXAMPLE
		$CurrentDirectory = Get-ScriptDirectory

#>
function Get-ScriptDirectory
{ 
	if($hostinvocation -ne $null)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#========================================================================
<#
	.SYNOPSIS
		Writes the specified text to the specified Logfile

	.DESCRIPTION
		The text is appended to the logfile. If the file does not exist, it will be created

	.PARAMETER  filename
		The file the text should be written to.

	.PARAMETER  text
		The text to append to the file.

	.PARAMETER  timestamp
		Writes the current timestamp at the beginning of the line.

	.EXAMPLE
		writeLog -filename "C:\log.txt" -text "Hallo Welt"

	.INPUTS
		System.String
#>
function Write-LogFile
{
	param(
	[Parameter(Mandatory=$true)][string]$filename,
	[Parameter(Mandatory=$true)][string]$text,
	[Parameter(Mandatory=$false)][boolean]$timestamp
	)
	if(-not $timestamp) 
	{
		$text = ";" + $text
		Out-File $filename -Append -NoClobber -InputObject $text 
	}
	else
	{
		$stamp = Get-Timestamp
		$text = $stamp + ";" + $text
		Out-File $filename -Append -NoClobber -InputObject $text
	}
}

#========================================================================
<#
	.SYNOPSIS
		Gets the currents timestamp in Format "dd:MM:yyyy hh:mm"

	.DESCRIPTION
		Gets the currents timestamp in Format "dd:MM:yyyy hh:mm"
	
	.OUTPUTS
		System.String

#>
function Get-Timestamp 
{
	#returns a padded timestamp string like 10.02.2014 17:02 
	$now = Get-Date
	$year = $now.Year.ToString()
	$month = $now.Month.ToString()
	$day = $now.Day.ToString()
	$hour = $now.Hour.ToString()
	$minute = $now.Minute.ToString()
	$second = $now.Second.ToString()

#region make sure all numbers have 2 digits	
	if ($month.length -lt 2) { $month = "0" + $month }
	if ($day.length -lt 2) { $day = "0" + $day }
	if ($hour.length -lt 2) { $hour = "0" + $hour }
	if ($minute.length -lt 2) {	$minute = "0" + $minute }
	if ($second.lenth -lt 2) { $second = "0" + $second }
#endregion
	
	#write-output $yr$mo$dy$hr$mi
	Write-Output ($day + "." + $month + "." + $year + " " + $hour + ":" + $minute + ":" + $second)
}

#========================================================================
<#
	.SYNOPSIS
		Writes a Text to Console and Logfile
	
	.PARAMETER text
		The text to write

	.PARAMETER $textcolor
		The color the text should be written to console

#>
function Tee-ToLogAndConsole
{
	param
	(
		[Parameter(Mandatory=$true)][string]$text,
		[Parameter(Mandatory=$false)][System.ConsoleColor]$textcolour
	)
	
	if($textcolour) { Write-Host $text -ForegroundColor $textcolour }
	else {Write-Host $text}
	
	
	Write-LogFile -text $text -filename $logfile -timestamp $true
}


#endregion



#========================================================================
# Scriptstart
#========================================================================


clear

$logfile = ((Get-ScriptDirectory) + "\" + "scriptlog_" + (Get-TimestampyyyyMMdd) + ".log")
$count = 0

Tee-ToLogAndConsole -text "###################"
Tee-ToLogAndConsole -text "#   Scriptstart   #"
Tee-ToLogAndConsole -text "###################"

if($Error){$Error.Clear()}

$PowerShellRunSpace = [System.Runtime.InterOpServices.Marshal]::SizeOf([System.IntPtr])
	
if (($PowerShellRunSpace -ne 4) -and ($PowerShellRunSpace -ne 2))
{
	Tee-ToLogAndConsole -text "ERROR;PowerShell is running in 64-bit or unknown runspace. Please execute the script in a 32-bit runspace for Lotus Notes related functionality." -textcolour 'Red'
	Write-Host "PowerShell is running in 64-bit or unknown runspace. Please execute the script in a 32-bit runspace for Lotus Notes related functionality."  -ForegroundColor 'Red'
}
else
{
	if($Error){$Error.Clear()}
	$NotesSession = New-Object -ComObject Lotus.NotesSession
	if($Error){$Error.Clear()}
	$NotesSession.Initialize()
	if($Error)
	{Tee-ToLogAndConsole -text ("ERROR;Cannot initialize Notes Session; Exiting Script") -textcolour 'Red'; exit }
	
	#DominoDir öffnen
	$DominoDirectory = $NotesSession.GetDatabase($DominoServer, "names.nsf", 0)
	if($Error){$Error.Clear()}
	$PeopleView = $DominoDirectory.GetView("People")
	$GroupsView = $DominoDirectory.GetView("(`$VIMGroups)")
	
	if($Error)
	{Tee-ToLogAndConsole -text "Cannot access Notes View People! Exiting Script" -textcolour 'Red';	exit }
	
	"GroupName;Members" | Out-File -FilePath $CSVOutFilePath -Encoding "UTF8"	

	$NotesDocument = $GroupsView.GetFirstDocument()
	while($NotesDocument -ne $null)
	{
		$count++
		$GroupMembers = $null
		$GroupName=""
		Tee-ToLogAndConsole -text ("Found Group " + $NotesDocument.GetItemValue("ListName")[0] + "! Extracting members...") 
		$GroupMembers = $NotesDocument.GetItemValue("Members")
		$GroupName = $NotesDocument.GetItemValue("ListName")[0]
		$textToWrite = $NotesDocument.GetItemValue("ListName")[0] + ";"
		$GroupMembersString=""
		
		foreach($member in $GroupMembers)
		{
			$found=$false
			#case 1/3: external contact
			if($member -like "*@*"){$GroupMembersString += ($member + $CSVMultiValueDelimiter);$found=$true}
			
			#case 2/3: person or group
			$PersonDocument = $PeopleView.GetfirstDocument()
			if($found){$PersonDocument = $PeopleView.GetLastDocument()}
			while($PersonDocument -ne $null)
			{	
				if($found){break}
				$FullName = $PersonDocument.GetItemValue("FullName")
				foreach($fn in $FullName)
				{
					if($member -ieq $fn)
					{
						$GroupMembersString += ($PersonDocument.GetItemValue("InternetAddress")[0] + $CSVMultiValueDelimiter)
						$PersonDocument = $PeopleView.GetLastDocument()
						$found=$true
						break
					}
				}
				if($found = $false) # search for corresponding group
				{
					$GroupSearchDoc = $GroupsView.GetFirstDocument()
					while ($GroupSearchDoc -ne $null)
					{
						if($found){break}
						if($member -ieq $GroupSearchDoc.GetItemValue("ListName")[0])
						{
							$found=$true
							if($GroupSearchDoc.GetItemValue("InternetAddress")[0] -ne "")
							{$GroupMembersString +=($GroupSearchDoc.GetItemValue("InternetAddress")[0] + $CSVMultiValueDelimiter) }
							else{$GroupMembersString +=($GroupSearchDoc.GetItemValue("ListName")[0] + "_NOMAIL" + $CSVMultiValueDelimiter)	}
							$GroupSearchDoc = $GroupsView.GetLastDocument()
						}
						$GroupSearchDoc = $GroupsView.GetNextDocument($GroupSearchDoc)
					}
				}
				$PersonDocument = $PeopleView.GetNextDocument($PersonDocument)
			}
			
			#case 3/3: Server - doesn't matter - will be ignored
			if($found -eq $false)
			{
				Tee-ToLogAndConsole -text ("'" + $member + "' cannot be resolved and will be skipped!") -textcolour 'Yellow' 
				$GroupMembersString += ($member + "_UNRESOLVABLE" + $CSVMultiValueDelimiter)
			}
						
		}
		

		$textToWrite = $GroupName + ";" + $GroupMembersString
		$textToWrite | Out-File -FilePath $CSVOutFilePath -Encoding "UTF8" -Append
		
		$NotesDocument = $PeopleView.GetNextDocument($NotesDocument)
		
	}
	Tee-ToLogAndConsole -text ("Script execution finished - " + $count.ToString() + " groups exported") -textcolour 'Green'
}
