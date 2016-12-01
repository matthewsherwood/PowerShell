#DISCLAIMER
#This script is not supported under any CDW standard support program or service. 
#The sample script is provided AS-IS without warranty of any kind. 
#CDW further disclaims all implied warranties including, without limitation, 
#any implied warranties of merchantability or of fitness for a particular purpose. 
#The entire risk arising out of the use or performance of the sample script and documentation remains with you. 
#In no event shall CDW, its authors, or anyone else involved in the creation, production, 
#or delivery of the script be liable for any damages whatsoever (including, without limitation, 
#damages for loss of business profits, business interruption, loss of business information, 
#or other pecuniary loss) arising out of the use of or inability to use the sample script or documentation, 
#even if CDW has been advised of the possibility of such damages.

#    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.


<#
.SYNOPSIS
	Completes AutoSuspended mailbox moves.

.DESCRIPTION
	Completes mailbox moves that are AutoSuspended. Filters based on BatchName.

.EXAMPLE
	C:\PS> .\CompleteMoves.ps1 
	
	This example simply runs the script using default functionality.

.INPUTS
	BatchName
		This specifies the move request BatchName to use as a filter. If you want to complete all AutoSuspended move requests, enter "*" as the BatchName.
		This string accepts wildcards, so to get all BatchNames that start with Nov29, use "Nov29*". Or to get all BatchNames that contain MoveASAP, use "*MoveASAP*".

.OUTPUTS
	Default log file for all output.
		".\CompleteMovesLog $DateTime.txt"
	
.NOTES
	Company			: CDW
	File Name		: CompleteMoves.ps1
	Author			: Matthew Sherwood
	Author Email	: matt.sherwood@cdw.com
	Date Created	: 11/30/2016
	
	Requirements:
		Exchange PowerShell Module


	Change log:
		11/30/2016
			- Initial script creation.
		12/01/2016
			- Added safeguard to verify correct move requests to be completed.
			- Added parameters

#>




param (
# Batch name. Used to specify which batch name to use. Enter a string such as Nov29* or April20_Office1Office2.
	[Parameter(Mandatory=$true,HelpMessage="Enter a string such as Nov29* or April20_Office1Office2")]
	[String[]]
	$BatchName,

# Reporting. Should detailed reporting be displayed on screen
    [Parameter()]
    [Switch[]]
    $Reporting = $true,

# Logging. Should logs be saved for this task
    [Parameter()]
    [Switch[]]
    $Logging = $true 
    
)



### Adjust these variables to fit your needs ###

# DateTimeStart
	$DateTimeStart             = Get-Date 
	$DateTimeStartHuman        = "$($DateTimeStart.ToString('yyyy-MM-ddTHH-mm-ss'))"
# OutputFile(s) 
	$LogFile = ".\Logs\CompleteMovesLog $DateTimeStartHuman.txt"        # Use ".\ResourceCreationLogFile $DateTime.txt" for default setting
	#$TranscriptFile        = ".\Logs\CompleteMovesTranscriptFile $DateTimeStartHuman.txt"                # Use ".\ResourceCreationLogFile $DateTime.txt" for default setting


<#
!                                                                               !
	!                                                                       !
		!                                                               !
			!                                                       !
================================================================================
				######### Don't change below this line #########                
================================================================================
			!                                                       !
		!                                                               !
	!                                                                       !
!                                                                               !
#>

# ****************************** #
# Setup                          #
# ****************************** #

# Script's Name for summary report
$ScriptName             = "CompleteMoves.ps1"


# ****************************** #
# Main script                    #
# ****************************** #

# Gather data
	# All moves in batch
	$MoveRequests              = Get-MoveRequest -BatchName $BatchName
	$MoveRequestStats          = $MoveRequests | Get-MoveRequestStatistics
	# Staged moves in batch
	$StagedMoveRequests        = $MoveRequests | Where-Object {$_.Status -eq "AutoSuspended"}
	$StagedMoveRequestStats    = $StagedMoveRequests | Get-MoveRequestStatistics
	# Gather Mailbox Offices
	$StagedMailboxes           = $StagedMoveRequests | Get-Mailbox | Select DisplayName, Office, Alias, UserPrincipalName, ExchangeGuid, Database, ArchiveDatabase, ServerName


if ($Logging) {
	"" >> $LogFile
	"All Moves" >> $LogFile
	$MoveRequests | ft -AutoSize >> $LogFile

	"" >> $LogFile
	"All Move Stats" >> $LogFile
	$MoveRequestStats | ft -AutoSize >> $LogFile

	"" >> $LogFile
	"Staged Moves" >> $LogFile
	$StagedMoveRequests | ft -AutoSize >> $LogFile

	"" >> $LogFile
	"Staged Move Stats" >> $LogFile
	$StagedMoveRequestStats | ft -AutoSize >> $LogFile

	"" >> $LogFile
	"Staged Mailboxes" >> $LogFile
	$StagedMailboxes | ft -AutoSize >> $LogFile
}


	Write-Host "
	# ****************************** #
	# Staged Moves                   #
	# ****************************** #" -ForegroundColor Magenta
	$StagedMoveRequests | ft -AutoSize

	Write-Host "
	# ****************************** #
	# Staged Move Statistics         #
	# ****************************** #" -ForegroundColor Magenta
	$StagedMoveRequestStats | Select DisplayName, StatusDetail, TotalMailboxSize, PercentComplete, TargetDatabase | Sort-Object TargetDatabase | ft -AutoSize
	#$StagedMoveRequestStats | Select DisplayName, StatusDetail, TotalMailboxSize, PercentComplete, TargetDatabase | Out-GridView 

	Write-Host "
	# ****************************** #
	# Staged Mailboxes               #
	# ****************************** #" -ForegroundColor Magenta
	$StagedMailboxes | Select DisplayName, Office ft -AutoSize


$Check  = Read-Host -Prompt '
!!!!!!!!!!!!
!! VERIFY !!
!!!!!!!!!!!!

Please verify that the above users are ready to complete!
	Press "Enter" to complete the mailboxes listed above, or "Ctrl+C" to stop the script'



$StagedMoveRequests | Resume-MoveRequest








# Build Summary Report
# Get end DateTime
$DateTimeEnd                = Get-Date
$DateTimeEndHuman           = "$($DateTimeEnd.ToString('yyyy-MM-ddTHH-mm-ss'))"
$ElapsedTime                = (Get-Date) - $DateTimeStart
$AverageSecPerMailbox       = [math]::Round(($ElapsedTime.TotalSeconds/$StagedMoveRequests.Count))


Write-Host "
# ****************************** #
# Summary                        #
# ****************************** #" -ForegroundColor Magenta

$OutputSummaryReportString = "
=================================================

	Script name:        $ScriptName
	Script start time:  $DateTimeStartHuman
	Script end time:    $DateTimeEndHuman

	Reports saved to...
		Log file:       $LogFile

	Summary Statistics
		Batch name:             $BatchName

		Mailbox Move Status 
			Total mailboxes in batch:   $($MoveRequests.count)
			Staged mailboxes processed for completion:           $($StagedMoveRequests.count)
 
	Time statistics
		Time elapsed to run report:     		$($ElapsedTime.ToString("%d")) days $($ElapsedTime.ToString("%h")) hours $($ElapsedTime.ToString("%m")) minutes $($ElapsedTime.ToString("%s")) seconds
		Average seconds per mailbox processed:  $AverageSecPerMailbox

=================================================
	"

Write-Host $OutputSummaryReportString -ForegroundColor Cyan
if ($Logging) {
	$OutputSummaryReportString >> $LogFile
}



