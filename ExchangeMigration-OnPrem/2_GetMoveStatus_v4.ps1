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
	Gathers status of mailbox move jobs in Exchange.

.DESCRIPTION
    Gathers status of mailbox move jobs in Exchange. Script can be used to gather status of all move requests, but can also filter based on BatchName.

.EXAMPLE
    C:\PS> .\GetMoveStatus.ps1 
    
    This example simply runs the script using default functionality. Default functionality includes status of all move tasks, as well as provides logging and detailed reporting to PowerShell session.

.EXAMPLE
    C:\PS> .\GetMoveStatus.ps1 -BatchName Dec01*
    
    This example will provide a the status of move requests that have a BatchName that starts with "Dec01". It will also provide a log file and detailed reporting to the PowerShell session.

.EXAMPLE
    C:\PS> .\GetMoveStatus.ps1 -BatchName Dec01-Office1 -Logging $false $Reporting $false
    
    This example will provide a the status of move requests that have "Dec01-Office1" specified as a BatchName.
    The Logging and Reporting parameters are also configured to disable the creation of a log file, and will supress detailed reporting to the PowerShell session.

.INPUTS
    BatchName
		This specifies the move request BatchName to use as a filter. If you want to complete all AutoSuspended move requests, enter "*" as the BatchName.
		This string accepts wildcards, so to get all BatchNames that start with Nov29, use "Nov29*". Or to get all BatchNames that contain MoveASAP, use "*MoveASAP*".

.OUTPUTS
    Default log file for all output.
        ".\GetMoveStatusLog $DateTime.txt"
    
.NOTES
	Company			: CDW
    File Name		: GetMoveStatus.ps1
	Author			: Matthew Sherwood
	Author Email	: matt.sherwood@cdw.com
	Date Created	: 11/29/2016
    
    Requirements:
        Exchange PowerShell Module


    Change log:
        11/29/2016
            - Initial script creation.
        11/30/2016
            - Added parameters.
        12/02/2016
            - Adjusted filters to more accuratly summarize MoveStatus.

#>


# ****************************** #
# Parameters                     #
# ****************************** #


param (
# Batch name. Used to specify which batch name to use. Enter a string such as Nov29* or April20_Office1Office2.
	[Parameter(HelpMessage="Enter a string such as Nov29* or April20_Office1Office2")]
	[String[]]
	$BatchName = "*",

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
    $LogFile = ".\Logs\GetMoveStatusLog $DateTimeStartHuman.txt"        # Use ".\ResourceCreationLogFile $DateTime.txt" for default setting
	#$TranscriptFile        = ".\Logs\GetMoveStatusTranscriptFile $DateTimeStartHuman.txt"                # Use ".\ResourceCreationLogFile $DateTime.txt" for default setting

# Specify Batch Name (Nov30_Office1Office2)
    #$BatchName = "Nov29_LexingtonHuntington"




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
$ScriptName             = "GetMoveStatus.ps1"

# Set indexes to 0
$i                      = 0
$CountErrors            = 0

# Null all arrays
$Mailboxes              = @()
$MoveRequests           = @()


$MoveRequestStats               = @()
$MoveRequestStatsAutoSuspended  = @()
$MoveRequestStatsCompleted      = @()
$MoveRequestStatsCopying        = @()
$MoveRequestStatsStalled        = @()
$MoveRequestStatsOther          = @()

# Gather other data used in script
$CountOffices           = $Offices.count





# ****************************** #
# Main script                    #
# ****************************** #

# Gather move status data
$MoveRequests           = Get-MoveRequest -BatchName $BatchName
$MoveRequestStats       = $MoveRequests | Get-MoveRequestStatistics
$Mailboxes              = $MoveRequests | Get-Mailbox



if ($Logging) {
    "
    # ****************************** #
    # Mailboxes Being Moved          #
    # ****************************** #" >> $LogFile
    $Mailboxes | ft Name, ServerName, Database, Office, UserPrincipalName -AutoSize >> $LogFile

    "
    # ****************************** #
    # Move Requests                  #
    # ****************************** #" >> $LogFile
    $MoveRequests | ft -AutoSize  >> $LogFile

    "
    # ****************************** #
    # Move Request Statistics        #
    # ****************************** #" >> $LogFile
    $MoveRequestStats | Select DisplayName, StatusDetail, TotalMailboxSize, PercentComplete, TargetDatabase | Sort-Object TargetDatabase | ft -AutoSize  >> $LogFile 
}


if ($Reporting) {

    Write-Host "
    # ****************************** #
    # Mailboxes Being Moved          #
    # ****************************** #" -ForegroundColor Magenta
    $Mailboxes | ft Name, ServerName, Database, Office, UserPrincipalName -AutoSize

    Write-Host "
    # ****************************** #
    # Move Requests                  #
    # ****************************** #" -ForegroundColor Magenta
    $MoveRequests | ft -AutoSize

    Write-Host "
    # ****************************** #
    # Move Request Statistics        #
    # ****************************** #" -ForegroundColor Magenta
    $MoveRequestStats | Select DisplayName, StatusDetail, TotalMailboxSize, PercentComplete, TargetDatabase | Sort-Object TargetDatabase | ft -AutoSize
    #$MoveRequestStats | Select DisplayName, StatusDetail, TotalMailboxSize, PercentComplete, TargetDatabase | Out-GridView 

}



# Count various move request statistics statuses for summary report 

$MoveRequestsCompleted              = $MoveRequests | where {$_.Status -eq "Completed"}
$MoveRequestsStaged                 = $MoveRequests | where {$_.Status -eq "AutoSuspended"}
$MoveRequestsNonStaged              = $MoveRequests | where {$_.Status -ne "Completed" -and $_.Status -ne "AutoSuspended"}
$MoveRequestStatsNonStaged          = $MoveRequestsNonStaged | Get-MoveRequestStatistics
$MoveRequestStatsNonStagedCopying   = $MoveRequestStatsNonStaged | Where-Object {$_.StatusDetail -like "Copying*"}
$MoveRequestStatsNonStagedStalled   = $MoveRequestStatsNonStaged | Where-Object {$_.StatusDetail -like "Stalled*"}
$MoveRequestStatsNonStagedOther     = $MoveRequestStatsNonStaged | Where-Object {$_.StatusDetail -notlike "Copying*" -and $_.StatusDetail -notlike "Stalled*"}






# Gather Summary Stats
$AllBytesTransferred = (($MoveRequestStats.BytesTransferred | Measure-Object -Sum).Sum)/1024/1024/1024
$AllTotalMailboxSize = (($MoveRequestStats.TotalMailboxSize | Measure-Object -Sum).Sum)/1024/1024/1024
$AllPercentComplete  = $AllBytesTransferred/$AllTotalMailboxSize*100
$MaxOverallDuration  = (($MoveRequestStats.OverallDuration | Measure-Object -Maximum).Maximum).ToString()
#$MinOverallDuration  = (($MoveRequestStats.OverallDuration | Measure-Object -Minimum).Minimum).ToString()

# Get end DateTime
$DateTimeEnd                = Get-Date
$DateTimeEndHuman           = "$($DateTimeEnd.ToString('yyyy-MM-ddTHH-mm-ss'))"
$ElapsedTime                = (Get-Date) - $DateTimeStart
$AverageSecPerMailbox       = [math]::Round(($ElapsedTime.TotalSeconds/$Mailboxes.Count))


# Build summary report 
$OutputSummaryReportStringHeader = "
# ****************************** #
# Summary                        #
# ****************************** #"

$OutputSummaryReportString1 = "
=================================================

    Script Name:        $ScriptName
    Script Start Time:  $DateTimeStartHuman
    Script End Time:    $DateTimeEndHuman

    Reports saved to...
        Log File:       $LogFile

    Summary Statistics
        Batch Name:             $BatchName
        Batch Overall Duration: $MaxOverallDuration

        Mailbox Move Status
            Completed:  $($MoveRequestsCompleted.count)"
$OutputSummaryReportStringStaged    = "            Staged:     $($MoveRequestsStaged.count)"
$OutputSummaryReportString2         = "            Copying:    $($MoveRequestStatsNonStagedCopying.count)
            Stalled:    $($MoveRequestStatsNonStagedStalled.count)"
$OutputSummaryReportStringOther     = "            Other:      $($MoveRequestStatsNonStagedOther.count)"
$OutputSummaryReportString3         = "            -----------------------
            Total:      $($MoveRequests.count)

        Data Transfer Status
            Batch Bytes Transferred:    $($AllBytesTransferred.ToString("#.00")) GB
            Batch Total Mailbox Size:   $($AllTotalMailboxSize.ToString("#.00")) GB
            Batch Percent Complete:     $($AllPercentComplete.ToString("00.0"))


    Time statistics
        Time elapsed to run report:     $($ElapsedTime.ToString("%d")) days $($ElapsedTime.ToString("%h")) hours $($ElapsedTime.ToString("%m")) minutes $($ElapsedTime.ToString("%s")) seconds
        Average seconds per Mailbox:    $AverageSecPerMailbox

=================================================
    "


# Output summary report to screen
Write-Host $OutputSummaryReportStringHeader -ForegroundColor Magenta 
Write-Host $OutputSummaryReportString1 -ForegroundColor Cyan

if ($MoveRequestsStaged) {
    Write-Host $OutputSummaryReportStringStaged -ForegroundColor Green
} else {
    Write-Host $OutputSummaryReportStringStaged -ForegroundColor Cyan
}

Write-Host $OutputSummaryReportString2 -ForegroundColor Cyan

if ($MoveRequestStatsNonStagedOther) {
    Write-Host $OutputSummaryReportStringOther -ForegroundColor Yellow
} else {
    Write-Host $OutputSummaryReportStringOther -ForegroundColor Cyan
}

Write-Host $OutputSummaryReportString3 -ForegroundColor Cyan

# Output summary report to log file
if ($Logging) {
    $OutputSummaryReportString1 >> $LogFile
    $OutputSummaryReportStringStaged >> $LogFile
    $OutputSummaryReportString2 >> $LogFile
    $OutputSummaryReportStringOther >> $LogFile
    $OutputSummaryReportString3 >> $LogFile
}

