<#
.SYNOPSIS
    Quick description of this script

.DESCRIPTION
    Longer description of what this script does

.PARAMETER FirstNumber
    This parameter does blablabla

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.INPUTS
    None. You cannot pipe objects to that script.

.OUTPUTS
    None for now

.EXAMPLE
.\Do-Something.ps1
This will launch the script and do someting

.EXAMPLE
.\Do-Something.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : Do-Something.ps1
VERSION : v1.0

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string]$pathToLogParserExe= "C:\Program Files (x86)\Log Parser 2.2\logparser.exe",
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "0.1"
<# Version changes
v0.1 : first script version
v0.1 -> v0.5 : 
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
$UserDocumentsFolder = "$($env:Userprofile)\Documents"
$OutputReport = "$UserDocumentsFolder\$($ScriptName)_Output_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$UserDocumentsFolder\$($ScriptName)_Logging_$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function Write-Log
{
	<#
	.SYNOPSIS
		This function creates or appends a line to a log file.
	.PARAMETER  Message
		The message parameter is the log message you'd like to record to the log file.
	.EXAMPLE
		PS C:\> Write-Log -Message 'Value1'
		This example shows how to call the Write-Log function with named parameters.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true,position = 0)]
		[string]$Message,
		[Parameter(Mandatory=$false,position = 1)]
        [string]$LogFileName=$ScriptLog,
        [Parameter(Mandatory=$false, position = 2)][switch]$Silent,
        [Parameter(Mandatory=$false)][switch]$Error
	)
	
	try
	{
		$DateTime = Get-Date -Format 'MM-dd-yy HH:mm:ss'
		$Invocation = "$($MyInvocation.MyCommand.Source | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
		Add-Content -Value "$DateTime - $Invocation - $Message" -Path $LogFileName
		if (!($Silent)){
            if ($Error){
                Write-Host $Message -ForegroundColor Red}
            } Else {
                Write-Host $Message -ForegroundColor Green}
            }   
	}
	catch
	{
		Write-Error $_.Exception.Message
	}
}

function Show-OpenFileDialog {
    <#
    .DESCRIPTION
        This function is a PowerShell call to the OpenFileDialog box using WPF.
        We'll see many examples of OpenFileDialog using Windows Forms, but we'll try to
        get off Windows Forms as it's legacy tech.
    
    .EXAMPLE
        PS_>Show-OpenFileDialog
        This will open a dialog box to enable the user to select a file. The output of
        the function is the full path of that file. It's useful for example to Import-CSV
        from a CSV file, or to select an Office document to be opened with PowerShell
        application automation...
    
    .EXAMPLE
        PS_>$FileName = Show-OpenFileDialog
        This will open a dialog box to enable the user to select a file, and the file name
        will be stored in the $FileName variable to be reused as described on the first example.
    
    .EXAMPLE
        PS_>$FileName = Show-OpenFileDialog -Title "Open an .XLSX file to be parsed" -Filter "Excel file|*.xlsx"  -InitialDirectory c:\MyExcelFiles
        This will open a dialog box to select a file, with a customized title, and with a default filter on *.xlsx Excel files. This
        dialog box will open the C:\MyExcelFiles directory to look for files. User can select later any other folder to look for.
    
    .LINK
        https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32?view=net-5.0
    
    #>
    
        param
        ($Title = 'Select a file to use', $Filter = 'Comma Separated|*.csv|Text|*.txt',$InitialDirectory = "c:\temp")
        
        Add-Type -AssemblyName PresentationFramework
    
        $dialog = New-Object -TypeName 'Microsoft.Win32.OpenFileDialog'
        $dialog.Title = $Title
        $dialog.Filter = $Filter
        If (!(Test-Path $InitialDirectory)){$InitialDirectory = "$($env:Userprofile)\Documents"} #If the default C:\temp doesn't exist, defaults to user's Document folder
        $dialog.InitialDirectory = $InitialDirectory
      
        if ($dialog.ShowDialog() -eq $true)
        {
            Return $dialog.FileName
        }
        else
        {
            Write-Warning 'Cancelled'
        }
    }
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
Write-Log "************************** Script Start **************************"
################################# Input ###########################################

$inputFile = Show-OpenFileDialog

########################### Check if Logparser is on the default installation folder ################################

If (!(Test-File $pathToLogParserExe)){
    Write-Log -Error "ERROR: Logparser not found on $pathToLogParserExe. Make sure you installed Logparser on this machine"
    exit
}

################################ DO NOT CHANGE ##################################################
$outExt = ".csv"
Get-ChildItem $inputFile -Filter *.log | Foreach-Object {
    $inFile = $_.FullName
    "Starting " + $inFile
    $outFile =  $_.FullName -replace "\.log",".csv"
    $cmd = "SELECT date, time, s-ip, cs-method, cs-uri-stem, cs-uri-query, s-port, cs-username, c-ip, cs(User-Agent) as cs-user-agent, sc-status, sc-substatus, sc-win32-status,sc-bytes,cs-bytes, time-taken INTO '"+$outFile+"' FROM '"+$inFile+"'"
    # on Windows 2016 with Exchange 2016 installed, I have the below fields::
    # Fields: date time s-ip cs-method cs-uri-stem cs-uri-query s-port cs-username c-ip cs(User-Agent) cs(Referer) sc-status sc-substatus sc-win32-status time-taken
    # on CSV it's:
    # date, time, s-ip, cs-method, cs-uri-stem, cs-uri-query, s-port, cs-username, c-ip, cs(User-Agent), sc-status, sc-substatus, sc-win32-status, time-taken, cs(Referer)
    
    $output =  & $pathToLogParserExe -i:W3C -o:csv $cmd | Out-String
    
    "Output of first conversion " +$output.Length + ". If it is 0, it will rerun to include missing fields with default values. Make sure sc-bytes or cs-bytes columns available in IISLogs."
 
    # Hack in case the log file dont have all fields it will rerun to include those with default values. Mainly it happens for sc-bytes, cs-bytes as those are not enabled by default.
    # TODO move the default logic to PowerBI if possible.
 
    if ($output.Length -eq 0) {
        #There may be missing columns such as sc-bytes, cs-bytes. Rerun without those fields & Add default value 0
        Write-Information "Converting again"
        $cmd = "SELECT date, time, s-ip, cs-method, cs-uri-stem, cs-uri-query, s-port, cs-username, c-ip, cs(User-Agent) as cs-user-agent, sc-status, sc-substatus, sc-win32-status,0 as sc-bytes,0 as cs-bytes, time-taken INTO '"+$outFile+"' FROM '"+$inFile+"'"
        
        $output =  & $pathToLogParserExe -i:W3C -o:csv $cmd | Out-String
        
        "Output of rerun is " + $output.Length
        if($output.Length -eq 0){
            Write-Error "Not able to convert. Please make sure the IIS log file is W3C format and minimum columns date, time, s-ip, cs-method, cs-uri-stem, cs-uri-query, s-port, cs-username, c-ip, cs(User-Agent), sc-status, sc-substatus, sc-win32-status, time-taken are included"
        }
    } else {
        "Completed file " + $inFile
    }
    "-"*100
}
"Completed all files"




<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
Write-Log "************************** Script End **************************"
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>











