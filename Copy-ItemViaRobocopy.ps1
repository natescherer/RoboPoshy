<#
.SYNOPSIS
This script wraps Robocopy in a way works better with PowerShell. It also can provide email alerts at the beginning
and/or end of a job.

.DESCRIPTION
TBD

.INPUTS
None

.OUTPUTS
TBD

.EXAMPLE
./robocopy_wrapper.ps1 -source "\\server1\source" -dest "\\server2\dest" -log "c:\robologs\log1.log" -email "email1@contoso.com" -threads 64 -sec -e
FILL OUT MORE HERE

.EXAMPLE


.LINK


.NOTES
TBD
#>
[CmdletBinding(DefaultParameterSetName="Default")]
param (
    [Parameter(ParameterSetName="Default",Mandatory=$true)]
    [Parameter(ParameterSetName="Email",Mandatory=$true)]
    # Source directory
    [string]$Source,

    [Parameter(ParameterSetName="Default",Mandatory=$true)]
    [Parameter(ParameterSetName="Email",Mandatory=$true)]
    # Destination directory
    [string]$Dest,

    [Parameter(ParameterSetName="Default",Mandatory=$false)]
    [Parameter(ParameterSetName="Email",Mandatory=$false)]
    # Specifies log file path. Defaults to a file in the user temp directory.
    [string]$Log = $($env:temp) + "\Copy-ItemViaRobocopy_" + $(Get-Date -format "yyMMdd_Hmmss") + ".log",

    [Parameter(ParameterSetName="Default",Mandatory=$false)]
    [Parameter(ParameterSetName="Email",Mandatory=$false)]
    # Specifies a number of threads 1-128 to use for copy. Defaults to 128.
    [int]$Threads = 128,

    [Parameter(ParameterSetName="Default",Mandatory=$false)]
    [Parameter(ParameterSetName="Email",Mandatory=$false)]
    # Sets robocopy to copy security ACLs along with files
    [switch]$Sec,

    [Parameter(ParameterSetName="Default",Mandatory=$false)]
    [Parameter(ParameterSetName="Email",Mandatory=$false)]
    # Sets robocopy to leave files in the dest that aren't in the source (/e) rather than deleting them (/mir)
    [switch]$E,

    [Parameter(ParameterSetName="Email",Mandatory=$true)]
    # Specifies SMTP server used to send email
    [string]$SmtpServer,

    [Parameter(ParameterSetName="Email",Mandatory=$false)]
    # Specifies to send mail either Before, After, or BeforeAndAfter robocopy
    [ValidateSet("Before","After","BeforeAndAfter")] 
    [string]$EmailMode,

    [parameter(ParameterSetName="Email",Mandatory=$false)]
    # TCP Port to connect to SMTP server on, if it is different than 25.
    [int]$SmtpPort = 25,

    [Parameter(ParameterSetName="Email",Mandatory=$false)]
    # Specifies a source address for messages. Defaults to computername@domain
    [string]$EmailFrom = "$($env:computername)@$($env:userdnsdomain)",

    [Parameter(ParameterSetName="Email",Mandatory=$true)]
    # Specifies a comma-separated (i.e. "a@b.com","b@b.com") list of email addresses to email upon job completion
    [string[]]$EmailTo
)
process {
    $StartTime = Get-Date

    if ($EmailMode -like "Before*") {
        $SmtpParamsBefore = @{
            From = $EmailFrom
            To = $EmailTo
            Subject = "Robocopy Start: '$Source' to '$Dest' on $env:computername"
            Body = "Robocopy Starting at $($StartTime):  '$Source' to '$Dest' on $env:computername"
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            Port = $SmtpPort
            UseSsl = $true
        }
        Send-MailMessage @SMTPParamsBefore
    }

    if ($e) {
        $CopyMode = "/e"
    } else {
        $CopyMode = "/mir"
    }

    if ((Get-CimInstance Win32_OperatingSystem).Version -lt 6.2) {
        $DCopy = "/dcopy:T"
    } else {
        $DCopy = "/dcopy:DAT"
    }

    $SecText = ""
    if ($Sec) {
        $SecText = "/sec"
    }

    robocopy $Source $Dest $CopyMode $SecText $DCopy /r:5 /w:5 /mt:$Threads /log:$Log /NP

    if ($EmailMode -like "*After") {
        $Elapsed = $(Get-Date) - $StartTime
        $ElapsedString = "$($Elapsed.days) Days $($Elapsed.Hours) Hours $($Elapsed.Minutes) Minutes $($Elapsed.Seconds) Seconds"

        $SmtpParamsAfter = @{
            From = $EmailFrom
            To = $EmailTo
            Subject = "Robocopy End: '$Source' to '$Dest' on $env:computername"
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            Port = $SmtpPort
            UseSsl = $true
        }
        if ((get-item $Log).length -lt 45MB) {
            $SmtpParamsAfter += @{Body = "Robocopy job has completed with details as follows. Log is attached.`n`nSource: $Source`nDestination: $Dest`nElapsed Time: $ElapsedString`nClient: $env:computername`nLog location on client: $Log"}
            $SmtpParamsAfter += @{Attachments = $Log}
        } else {
            $SmtpParamsAfter += @{Body = "Robocopy job has completed with details as follows. Log is too large to email, but is located at address below.`n`nSource: $Source`nDestination: $Dest`nElapsed Time: $Elapsedsrting`nClient: $env:computername`nLog location on client: $Log"}
        }
        Send-MailMessage @SMTPParamsAfter
    }
}